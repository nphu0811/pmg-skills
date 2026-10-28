"""Source-first media research with explicit licensing and provenance.

Providers supply actual results; generation is a separate optional adapter. No
network result, license, URL, or generative capability is invented by fallback.
This module prepares media and never compiles PowerPoint animations.
"""
import hashlib
import html
import json
import mimetypes
import re
import subprocess
import urllib.parse
import urllib.request
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Protocol

import numpy as np
from PIL import Image


USER_AGENT = 'PMG-Skills/1.6 (+https://github.com/nphu0811/pmg-skills)'
NASA_TERMS = 'https://www.nasa.gov/nasa-brand-center/images-and-media/'


def _json_get(url):
    request = urllib.request.Request(url, headers={'User-Agent': USER_AGENT})
    with urllib.request.urlopen(request, timeout=25) as response:
        return json.load(response)


def _plain(value):
    return html.unescape(re.sub(r'<[^>]+>', '', str(value or ''))).strip()


@dataclass
class AssetRequirement:
    query: str
    kind: str = 'image'
    min_width: int = 640
    min_height: int = 360
    aspect_ratio: Optional[float] = None
    transparency: bool = False
    purpose: str = 'educational presentation'
    visual_style: str = ''
    needs_decomposition: bool = False


@dataclass
class AssetRecord:
    asset_id: str
    title: str
    kind: str
    source_url: str
    download_url: str
    provider: str
    provenance: str = 'retrieved'
    license_name: str = 'unknown'
    license_url: str = ''
    attribution: str = ''
    restrictions: List[str] = field(default_factory=list)
    rights_status: str = 'review_required'
    width: int = 0
    height: int = 0
    local_path: Optional[str] = None
    sha256: Optional[str] = None
    quality: Dict[str, Any] = field(default_factory=dict)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self):
        return asdict(self)


class AssetProvider(Protocol):
    def search(self, requirement: AssetRequirement, limit: int = 10) -> List[AssetRecord]:
        ...


class AssetGenerator(Protocol):
    def generate(self, requirement: AssetRequirement, output_dir: str) -> AssetRecord:
        ...


class WikimediaProvider:
    """Commons API metadata is retained, including trademark/usage restrictions."""
    endpoint = 'https://commons.wikimedia.org/w/api.php'

    def _parse(self, data):
        records = []
        for page in data.get('query', {}).get('pages', {}).values():
            info = (page.get('imageinfo') or [{}])[0]
            if not info.get('url'):
                continue
            metadata = info.get('extmetadata', {})
            def value(key):
                return _plain(metadata.get(key, {}).get('value', ''))
            mime = info.get('mime', mimetypes.guess_type(info['url'].split('?')[0])[0] or '')
            kind = 'video' if mime.startswith('video/') else ('icon' if 'svg' in mime else 'image')
            license_name = value('LicenseShortName') or value('UsageTerms') or 'unknown'
            restrictions = [item for item in value('Restrictions').split('|') if item]
            # Copyright and trademark are separate questions. Keep restrictions;
            # engineering/educational use is recorded by the caller's assessment.
            permissive = any(word in license_name.lower() for word in ('public domain', 'cc0', 'cc by', 'cc-by'))
            records.append(AssetRecord(
                f'commons:{page["pageid"]}', page['title'].removeprefix('File:') if hasattr(str, 'removeprefix') else page['title'][5:],
                kind, info.get('descriptionurl', ''), info['url'], 'Wikimedia Commons',
                license_name=license_name, license_url=value('LicenseUrl') or info.get('descriptionurl', ''),
                attribution=value('Artist') or value('Credit'), restrictions=restrictions,
                rights_status='compatible' if permissive and not restrictions else 'review_required',
                width=info.get('width', 0), height=info.get('height', 0),
                metadata={'description': value('ImageDescription'), 'credit': value('Credit'),
                          'license_metadata': metadata, 'mime': mime}))
        return records

    def search(self, requirement, limit=10):
        query = requirement.query + (' filetype:svg' if requirement.kind == 'icon' else '')
        parameters = dict(action='query', format='json', generator='search', gsrsearch=query,
                          gsrnamespace=6, gsrlimit=limit, prop='imageinfo', iiprop='url|size|mime|extmetadata')
        return self._parse(_json_get(self.endpoint+'?'+urllib.parse.urlencode(parameters)))

    def lookup(self, title: str):
        parameters = dict(action='query', format='json', titles=title if title.startswith('File:') else 'File:'+title,
                          prop='imageinfo', iiprop='url|size|mime|extmetadata')
        return self._parse(_json_get(self.endpoint+'?'+urllib.parse.urlencode(parameters)))


class NASAProvider:
    """NASA media is conditional; never label every image/logo unrestricted."""
    def search(self, requirement, limit=10):
        media_type = 'video' if requirement.kind == 'video' else 'image'
        parameters = urllib.parse.urlencode(dict(q=requirement.query, media_type=media_type, page_size=limit))
        results = _json_get('https://images-api.nasa.gov/search?'+parameters)
        records = []
        for item in results.get('collection', {}).get('items', []):
            data = (item.get('data') or [{}])[0]
            identity = data.get('nasa_id')
            if not identity:
                continue
            links = item.get('links') or []
            preview = next((link['href'] for link in links if link.get('rel') == 'preview'), '')
            records.append(AssetRecord(
                'nasa:'+identity, data.get('title', identity), media_type,
                'https://images.nasa.gov/details/'+urllib.parse.quote(identity), preview,
                'NASA Image and Video Library', license_name='NASA Media Usage Guidelines (conditional)',
                license_url=NASA_TERMS, attribution=data.get('secondary_creator') or data.get('photographer') or 'NASA/'+data.get('center', ''),
                restrictions=['third-party rights', 'protected insignia/logotypes', 'privacy/publicity', 'no endorsement'],
                metadata={'nasa_id': identity, 'description': data.get('description', ''), 'search_metadata': data,
                          'manifest_url': item.get('href', '')}))
        return records

    def resolve_media(self, asset: AssetRecord):
        manifest = _json_get('https://images-api.nasa.gov/asset/'+urllib.parse.quote(asset.metadata['nasa_id']))
        urls = [item.get('href', '') for item in manifest.get('collection', {}).get('items', [])]
        if asset.kind == 'video':
            candidates = [url for url in urls if urllib.parse.urlsplit(url).path.lower().endswith('.mp4')]
            candidates.sort(key=lambda url: ('~orig' not in url, '~medium' not in url))
        else:
            candidates = [url for url in urls if urllib.parse.urlsplit(url).path.lower().endswith(('.jpg', '.jpeg', '.png'))]
            candidates.sort(key=lambda url: ('~orig' not in url, '~large' not in url))
        if not candidates:
            raise ValueError('NASA manifest contains no supported media file')
        asset.download_url = candidates[0]
        asset.metadata['asset_manifest'] = manifest
        return asset


def approve_usage(asset: AssetRecord, purpose: str, evidence: str, reviewer: str = 'agent'):
    if not purpose.strip() or not evidence.strip():
        raise ValueError('A source-grounded rights assessment is required')
    if asset.license_name == 'unknown':
        raise ValueError('Cannot approve an invented or unknown license')
    asset.metadata['rights_assessment'] = {'purpose': purpose, 'evidence': evidence, 'reviewer': reviewer,
                                         'timestamp_utc': datetime.now(timezone.utc).isoformat()}
    asset.rights_status = 'compatible'
    return asset


def score_asset(asset: AssetRecord, requirement: AssetRequirement):
    tokens = set(re.findall(r'\w+', requirement.query.lower()))
    description = (asset.title+' '+asset.metadata.get('description', '')).lower()
    relevance = sum(token in description for token in tokens)/max(1, len(tokens))
    resolution = min(1, min(asset.width/max(1, requirement.min_width), asset.height/max(1, requirement.min_height)))
    aspect = 1.0
    if requirement.aspect_ratio and asset.height:
        aspect = min(asset.width/asset.height, requirement.aspect_ratio)/max(asset.width/asset.height, requirement.aspect_ratio)
    quality = float(asset.quality.get('technical_score', .5))
    transparency = 1.0 if not requirement.transparency or asset.quality.get('has_transparency') else .0
    rights = 1.0 if asset.rights_status == 'compatible' else 0.0
    score = .3*relevance + .15*resolution + .15*aspect + .15*quality + .1*transparency + .15*rights
    return {'score': score, 'semantic_relevance': relevance, 'resolution': resolution,
            'composition_aspect_fit': aspect, 'technical_quality': quality,
            'transparency_fit': transparency, 'license_compatibility': rights,
            'method': 'deterministic multi-criterion ranking; no model semantic claim'}


class VisualAssetEngine:
    def __init__(self, providers=None, generator=None):
        self.providers = providers if providers is not None else [WikimediaProvider(), NASAProvider()]
        self.generator = generator

    def research(self, requirement: AssetRequirement, limit: int = 10):
        records, errors = [], []
        for provider in self.providers:
            try:
                records.extend(asset for asset in provider.search(requirement, limit)
                               if asset.kind == requirement.kind)
            except Exception as exc:
                errors.append({'provider': type(provider).__name__, 'error': str(exc)})
        unique = {record.asset_id: record for record in records}
        ranked = sorted(unique.values(), key=lambda asset: score_asset(asset, requirement)['score'], reverse=True)
        return {'status': 'FOUND' if ranked else 'UNAVAILABLE', 'assets': ranked, 'provider_errors': errors}

    def download(self, asset: AssetRecord, output_dir: str, max_bytes: int = 256*1024*1024,
                 research_only: bool = False):
        if asset.rights_status != 'compatible' and not research_only:
            raise ValueError('Asset usage rights require source review before production')
        if urllib.parse.urlsplit(asset.download_url).scheme not in ('https', 'http'):
            raise ValueError('Asset must have an actual HTTP(S) source URL')
        directory = Path(output_dir).resolve()/('research' if research_only else asset.provenance)
        directory.mkdir(parents=True, exist_ok=True)
        extension = Path(urllib.parse.urlsplit(asset.download_url).path).suffix.lower() or '.bin'
        identity = hashlib.sha256(asset.asset_id.encode('utf-8')).hexdigest()[:20]
        output = directory/(identity+extension)
        request = urllib.request.Request(asset.download_url, headers={'User-Agent': USER_AGENT})
        total = 0
        temporary = output.with_suffix(extension+'.partial')
        try:
            with urllib.request.urlopen(request, timeout=30) as response, temporary.open('wb') as stream:
                while True:
                    chunk = response.read(1024*1024)
                    if not chunk:
                        break
                    total += len(chunk)
                    if total > max_bytes:
                        raise ValueError('Media exceeds the configured download resource budget')
                    stream.write(chunk)
            temporary.replace(output)
        finally:
            if temporary.exists():
                temporary.unlink()
        asset.local_path = str(output)
        asset.sha256 = hashlib.sha256(output.read_bytes()).hexdigest()
        asset.metadata['retrieved_at_utc'] = datetime.now(timezone.utc).isoformat()
        asset.metadata['research_only'] = research_only
        if asset.kind in ('image', 'icon') and extension != '.svg':
            with Image.open(output) as image:
                asset.width, asset.height = image.size
                rgba = np.asarray(image.convert('RGBA'))
            contrast = float(np.std(rgba[:, :, :3]))
            asset.quality.update(has_transparency=bool((rgba[:, :, 3] < 255).any()),
                                 contrast_std=contrast, technical_score=min(1.0, contrast/60))
        self.save_record(asset, directory/(identity+'.json'))
        return asset

    def acquire(self, requirement: AssetRequirement, output_dir: str):
        result = self.research(requirement)
        failures = list(result['provider_errors'])
        for asset in result['assets']:
            if asset.rights_status != 'compatible':
                continue
            try:
                self.download(asset, output_dir)
                if asset.kind not in ('image', 'icon') or (
                    asset.width >= requirement.min_width and asset.height >= requirement.min_height
                ):
                    return {'status': 'READY', 'asset': asset, 'provenance': 'retrieved', 'diagnostics': failures}
            except Exception as exc:
                failures.append({'asset_id': asset.asset_id, 'error': str(exc)})
        pending = [asset for asset in result['assets'] if asset.rights_status == 'review_required']
        if pending:
            return {'status': 'REVIEW_REQUIRED', 'asset': None, 'candidates': pending,
                    'diagnostics': failures, 'reason': 'Existing candidates need source/rights assessment before generation fallback.'}
        if self.generator is not None:
            generated = self.generator.generate(requirement, str(Path(output_dir)/'generated'))
            if generated.provenance != 'generated' or not generated.local_path or not Path(generated.local_path).is_file():
                raise ValueError('Generation adapter must return a real generated file and provenance')
            if generated.rights_status != 'compatible':
                raise ValueError('Generated asset usage rights are unverified')
            return {'status': 'READY', 'asset': generated, 'provenance': 'generated', 'diagnostics': failures}
        return {'status': 'UNAVAILABLE', 'asset': None,
                'diagnostics': failures, 'reason': 'No suitable permitted asset retrieved; no generation adapter configured.'}

    @staticmethod
    def save_record(asset, path):
        Path(path).write_text(json.dumps(asset.to_dict(), indent=2, ensure_ascii=False), encoding='utf-8')


def prepare_video(asset: AssetRecord, output_path: str, start_sec: float = 0, duration_sec: Optional[float] = None):
    if not asset.local_path or asset.kind != 'video' or asset.rights_status != 'compatible':
        raise ValueError('A retrieved, permitted video is required')
    if start_sec < 0 or (duration_sec is not None and duration_sec <= 0):
        raise ValueError('Invalid video trim interval')
    destination = Path(output_path).resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    command = ['ffmpeg', '-y', '-ss', str(start_sec), '-i', asset.local_path]
    if duration_sec is not None:
        command.extend(['-t', str(duration_sec)])
    command.extend(['-c:v', 'libx264', '-an', str(destination)])
    subprocess.run(command, check=True, capture_output=True, timeout=180)
    asset.metadata.setdefault('preparation', []).append({'operation': 'trim', 'start_sec': start_sec,
                                                       'duration_sec': duration_sec, 'output_path': str(destination)})
    return str(destination)


def prepare_image(asset: AssetRecord, output_path: str, crop=None, size=None):
    """Explicit derived crop/normalization; preserve original and provenance."""
    if not asset.local_path or asset.rights_status != 'compatible':
        raise ValueError('A real permitted asset is required')
    source = Path(asset.local_path)
    destination = Path(output_path).resolve()
    if destination == source.resolve():
        raise ValueError('Prepared assets must not overwrite the original')
    destination.parent.mkdir(parents=True, exist_ok=True)
    with Image.open(source) as image:
        prepared = image.convert('RGBA')
        if crop is not None:
            left, top, right, bottom = crop
            if not 0 <= left < right <= image.width or not 0 <= top < bottom <= image.height:
                raise ValueError('Crop is outside source bounds')
            prepared = prepared.crop(crop)
        if size is not None:
            prepared = prepared.resize(size, Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS)
        prepared.save(destination)
    asset.metadata.setdefault('preparation', []).append({'operation': 'image_crop_resize', 'input': str(source),
                                                        'source_sha256': asset.sha256, 'crop': crop,
                                                        'size': size, 'output': str(destination)})
    asset.local_path = str(destination)
    asset.width, asset.height = prepared.size
    asset.sha256 = hashlib.sha256(destination.read_bytes()).hexdigest()
    return asset


def prepare_svg_icon(asset: AssetRecord, output_path: str, color: Optional[str] = None,
                     stroke_width: Optional[float] = None):
    """Normalize real SVG source while retaining vector paths and attribution."""
    from lxml import etree
    if not asset.local_path or asset.rights_status != 'compatible':
        raise ValueError('A permitted vector icon is required')
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    root = etree.parse(asset.local_path, parser).getroot()
    if root.tag != '{http://www.w3.org/2000/svg}svg':
        raise ValueError('Source is not an SVG icon')
    if color is not None and not re.fullmatch(r'[0-9A-Fa-f]{6}', color.lstrip('#')):
        raise ValueError('Icon color must be a six-digit RGB hex value')
    for node in root.iter():
        if color:
            for attribute in ('fill', 'stroke'):
                if node.get(attribute) not in (None, 'none'):
                    node.set(attribute, '#'+color.lstrip('#'))
        if stroke_width is not None and node.get('stroke') not in (None, 'none'):
            node.set('stroke-width', str(stroke_width))
    destination = Path(output_path).resolve()
    if destination == Path(asset.local_path).resolve():
        raise ValueError('Normalize into a separate derived asset')
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(etree.tostring(root, xml_declaration=True, encoding='UTF-8'))
    asset.metadata.setdefault('preparation', []).append({'operation': 'svg_style_normalization', 'color': color,
                                                       'stroke_width': stroke_width, 'output': str(destination)})
    asset.local_path = str(destination)
    asset.sha256 = hashlib.sha256(destination.read_bytes()).hexdigest()
    return asset
