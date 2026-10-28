"""One atomic OPC part writer shared by engine metadata and asset references."""
import os
import tempfile
import zipfile
from typing import Dict
from lxml import etree

CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'


def write_package_parts(pptx_path: str, parts: Dict[str, bytes], mime_types: Dict[str, str]):
    for part in parts:
        if part.startswith('/') or '..' in part.split('/') or part == '[Content_Types].xml':
            raise ValueError(f'Invalid OPC part name: {part}')
        if part not in mime_types or '/' not in mime_types[part]:
            raise ValueError(f'Explicit ContentType required for {part}')
    descriptor, temporary = tempfile.mkstemp(suffix='.pptx', dir=os.path.dirname(os.path.abspath(pptx_path)))
    os.close(descriptor)
    try:
        with zipfile.ZipFile(pptx_path) as source, zipfile.ZipFile(temporary, 'w', zipfile.ZIP_DEFLATED) as output:
            root = etree.fromstring(source.read('[Content_Types].xml'))
            for part, mime in mime_types.items():
                extension = part.rsplit('.', 1)[-1].lower()
                # JSON must keep the canonical Default registration invariant.
                if extension == 'json':
                    for node in list(root):
                        if node.get('Extension', '').lower() == 'json' or node.get('PartName') == '/'+part:
                            root.remove(node)
                    if not any(node.get('Extension') == 'json' for node in root):
                        etree.SubElement(root, f'{{{CONTENT_TYPES_NS}}}Default', Extension='json', ContentType='application/json')
                else:
                    for node in list(root):
                        if node.get('PartName') == '/'+part:
                            root.remove(node)
                    etree.SubElement(root, f'{{{CONTENT_TYPES_NS}}}Override', PartName='/'+part, ContentType=mime)
            for item in source.infolist():
                if item.filename in parts or item.filename == '[Content_Types].xml':
                    continue
                output.writestr(item, source.read(item.filename))
            output.writestr('[Content_Types].xml', etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True))
            for part, data in parts.items():
                output.writestr(part, data)
        os.replace(temporary, pptx_path)
    finally:
        if os.path.exists(temporary):
            os.remove(temporary)
