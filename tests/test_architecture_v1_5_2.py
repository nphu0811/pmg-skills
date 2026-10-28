"""v1.5.2 ownership, compatibility and certification regressions.

Graph/decomposition/camera acceptance belongs to the gated v1.6/v1.7 work.
These tests never represent simulated playback as real PowerPoint execution.
"""
import ast
import json
import os
import subprocess
import zipfile
from pathlib import Path
from unittest.mock import Mock

import pytest
from lxml import etree

from pmg import MotionDeck, PMGOrchestrator, SlideAnimationTimeline, StaticValidator
from pmg.animations import AnimationStep
from pmg.models import ExpectedTimeline, ReviewDefect, ReviewReport, ValidationStatus, VideoMetadata
from pmg.orchestrator import _embed_manifest_in_pptx
from pmg.self_healing import execute_repair_loop
from pmg.worker.powerpoint_probe import PP_ALERTS_ALL, file_sha256


def build_deck(path, transition="morph"):
    deck = MotionDeck()
    for index in range(2):
        slide = deck.add_slide(transition="none" if index == 0 else transition)
        deck.add_hero_title(slide, f"Scene {index + 1}", subtitle="Editable native content")
        deck.add_glass_card(slide, 1, 3, 4, 2, name="!!Content_Panel")
    deck.save(str(path))
    return str(path)


def rewrite(path, replacements=None, additions=None, removed=()):
    target = path.with_name('rewritten.pptx')
    with zipfile.ZipFile(path) as source, zipfile.ZipFile(target, 'w') as output:
        for item in source.infolist():
            if item.filename not in removed:
                output.writestr(item, (replacements or {}).get(item.filename, source.read(item.filename)))
        for name, data in (additions or {}).items():
            output.writestr(name, data)
    os.replace(target, path)


def test_qa_is_read_only_and_shared_legacy_reports(tmp_path, monkeypatch, capsys):
    from scripts.motion_validator import validate_deck
    from scripts.slide_inspector import audit_presentation
    import pmg.validator
    path = Path(build_deck(tmp_path / 'audit.pptx'))
    monkeypatch.setattr(pmg.validator, 'verify_clean_powerpoint_open', Mock(side_effect=AssertionError('COM in static audit')))
    before = path.stat().st_mtime_ns, file_sha256(str(path))
    audit = StaticValidator().validate(str(path))
    legacy = validate_deck(str(path))
    assert audit['is_valid'] and audit['status'] == 'NOT_VERIFIED'
    assert legacy.audit == audit and not legacy.certified
    assert audit_presentation(str(path)) is True
    assert (path.stat().st_mtime_ns, file_sha256(str(path))) == before
    legacy.print_summary()
    assert 'NOT_VERIFIED' in capsys.readouterr().out


@pytest.mark.parametrize('effect,expected', [('zoom_in', 'zoom'), ('fade_in', 'fade'), ('spin', 'rotate'),
                                            ('grow_shrink', 'scale'), ('bezier_motion_path', 'motion_path')])
def test_animation_aliases_use_existing_compiler(effect, expected):
    deck = MotionDeck()
    slide = deck.add_slide(transition='none')
    shape = deck.add_glass_card(slide, 1, 1, 2, 2)
    timeline = SlideAnimationTimeline()
    step = timeline.add_effect(shape, effect, motion_path='M 0 0 L 0.1 0.1 E')
    assert step.effect_type == expected
    timeline.compile_to_openxml(slide)
    assert slide._element.find('.//{http://schemas.openxmlformats.org/presentationml/2006/main}spTgt') is not None


@pytest.mark.parametrize('effect', ['unknown', '', None, 'zoom_out', 'morph'])
def test_unknown_effects_are_rejected_before_compilation(effect):
    with pytest.raises(ValueError, match='Unsupported native'):
        AnimationStep('hero', 2, effect)


def test_mutated_effect_or_invalid_target_does_not_destroy_timing():
    deck = MotionDeck()
    slide = deck.add_slide(transition='none')
    shape = deck.add_glass_card(slide, 1, 1, 2, 2)
    timeline = SlideAnimationTimeline()
    step = timeline.add_effect(shape, 'fade')
    timeline.compile_to_openxml(slide)
    before = etree.tostring(slide._element)
    step.effect_type = 'broken'
    with pytest.raises(ValueError):
        timeline.compile_to_openxml(slide)
    assert etree.tostring(slide._element) == before
    step.effect_type = 'fade'
    step.target_spid = 123456
    with pytest.raises(ValueError, match='does not exist'):
        timeline.compile_to_openxml(slide)
    assert etree.tostring(slide._element) == before


def test_unregistered_manifest_negative_fixture_and_single_step_repair(tmp_path):
    path = Path(build_deck(tmp_path / 'unmapped.pptx'))
    rewrite(path, additions={'ppt/pmg_execution_manifest.json': b'{}'})
    audit = StaticValidator().validate(str(path))
    assert not audit['is_valid']
    assert any(d['object_id'] == 'ppt/pmg_execution_manifest.json' and
               d['remediation_hint'] == 'declare_part_content_type' for d in audit['defects'])
    execute_repair_loop(str(path), 1, [ReviewDefect(**d) for d in audit['defects']], audit)
    assert StaticValidator().validate(str(path))['is_valid']


@pytest.mark.parametrize('corruption', ['duplicate', 'invalid_xml', 'missing_types', 'missing_media', 'missing_rels'])
def test_opc_defects_are_structured_and_critical(tmp_path, corruption):
    path = Path(build_deck(tmp_path / 'package.pptx'))
    if corruption == 'duplicate':
        with zipfile.ZipFile(path, 'a') as output:
            output.writestr('ppt/presentation.xml', output.read('ppt/presentation.xml'))
    elif corruption == 'invalid_xml':
        rewrite(path, replacements={'docProps/core.xml': b'<bad>'})
    elif corruption == 'missing_types':
        rewrite(path, removed=('[Content_Types].xml',))
    elif corruption == 'missing_media':
        with zipfile.ZipFile(path) as source:
            root = etree.fromstring(source.read('ppt/slides/_rels/slide1.xml.rels'))
        etree.SubElement(root, '{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
                         Id='rId404', Type='image', Target='../media/missing.png')
        rewrite(path, replacements={'ppt/slides/_rels/slide1.xml.rels': etree.tostring(root)})
    else:
        rewrite(path, removed=('ppt/_rels/presentation.xml.rels',))
    audit = StaticValidator().validate(str(path))
    assert not audit['is_valid']
    assert any(d['category'] == 'openxml_structure' and d['severity'] == 'critical' for d in audit['defects'])


def test_legacy_repair_calls_canonical_orchestrator(tmp_path, monkeypatch):
    import pmg.orchestrator
    from scripts.motion_validator import review_and_heal_deck
    path = build_deck(tmp_path / 'legacy.pptx')
    manifest = Mock(static_validation=StaticValidator().validate(path), status=ValidationStatus.NOT_VERIFIED,
                    delivered_pptx=None, repair_history=[])
    run = Mock(return_value=manifest)
    monkeypatch.setattr(pmg.orchestrator.PMGOrchestrator, 'run_pipeline', run)
    accepted, report, actions = review_and_heal_deck(path, allow_substitution=False)
    assert run.call_count == 1 and run.call_args.kwargs['allow_substitution'] is False
    assert not accepted and not report.certified and report.status == 'NOT_VERIFIED'
    assert actions == []


def test_only_orchestrator_owns_retry_loop():
    import inspect
    import scripts.motion_validator
    for function in (scripts.motion_validator.review_and_heal_deck, execute_repair_loop):
        tree = ast.parse(inspect.getsource(function))
        assert not any(isinstance(node, ast.While) for node in ast.walk(tree))
    with pytest.raises(ValueError):
        PMGOrchestrator().run_pipeline('unused.pptx', max_repair_loops=4)


def fake_playback(engine, mutate=False, omit_proof=False):
    def run(candidate_pptx, qa_dir, **kwargs):
        with zipfile.ZipFile(candidate_pptx) as source:
            assert 'ppt/pmg_execution_manifest.json' in source.namelist()
            assert json.loads(source.read('ppt/pmg_execution_manifest.json'))['final_status'] == 'NOT_VERIFIED'
        audit = engine.validator.validate(candidate_pptx)
        from pmg.validator import extract_expected_timeline
        timeline = extract_expected_timeline(candidate_pptx)
        proof = dict(is_clean=True, skipped=False, source_unchanged=True, sha256=file_sha256(candidate_pptx))
        telemetry = {} if omit_proof else dict(clean_open=proof, validated_pptx_sha256=proof['sha256'])
        if mutate:
            with open(candidate_pptx, 'ab') as stream:
                stream.write(b'changed after validation')
        report = ReviewReport(ValidationStatus.PASS, visual_quality_passed=True, animation_behavior_passed=True,
                              temporal_quality_passed=True, capture_quality_passed=True,
                              slide_steady_states_verified=timeline.total_slides)
        video = VideoMetadata('mock.mp4', exists=True, decodable=True, actual_fps=60, width=1920, height=1080)
        return audit, timeline, video, video, telemetry, report, 'remote_windows_worker', 'test worker'
    return run


def test_packaged_bytes_are_frozen_and_hash_is_in_companion(tmp_path, monkeypatch):
    engine = PMGOrchestrator(str(tmp_path / 'qa'))
    monkeypatch.setattr(engine, '_run_validation_and_playback_pass', fake_playback(engine))
    output = tmp_path / 'final.pptx'
    manifest = engine.run_pipeline(str(output), builder_fn=build_deck)
    assert manifest.status == ValidationStatus.PASS
    assert file_sha256(str(output)) == manifest.provenance['final_pptx_sha256'] == file_sha256(manifest.source_pptx)
    companion = json.loads(Path(str(output) + '.manifest.json').read_text(encoding='utf-8'))
    assert companion['final_verification']['sha256'] == file_sha256(str(output))
    assert manifest.final_verification['playback_verified'] is True


@pytest.mark.parametrize('mutate,omit_proof,expected', [(True, False, ValidationStatus.FAIL),
                                                      (False, True, ValidationStatus.BLOCKED)])
def test_changed_bytes_or_unbound_remote_evidence_never_delivers(tmp_path, monkeypatch, mutate, omit_proof, expected):
    engine = PMGOrchestrator(str(tmp_path / 'qa'))
    monkeypatch.setattr(engine, '_run_validation_and_playback_pass', fake_playback(engine, mutate, omit_proof))
    output = tmp_path / 'existing.pptx'
    output.write_bytes(b'preserve the existing user file')
    manifest = engine.run_pipeline(str(output), builder_fn=build_deck, max_repair_loops=0)
    assert manifest.status == expected and manifest.delivered_pptx is None
    assert output.read_bytes() == b'preserve the existing user file'


def test_static_audit_never_certifies_and_repairs_use_one_budget(tmp_path):
    engine = PMGOrchestrator(str(tmp_path / 'qa'))
    output = tmp_path / 'not_certified.pptx'
    manifest = engine.run_pipeline(str(output), builder_fn=lambda p: build_deck(p, 'none'), worker_mode='static')
    assert manifest.status == ValidationStatus.NOT_VERIFIED
    assert manifest.repair_loops_executed == 1
    assert not output.exists() and manifest.delivered_pptx is None


def test_failed_final_destination_reopen_restores_previous_output(tmp_path, monkeypatch):
    import pmg.orchestrator as orchestration
    engine = PMGOrchestrator(str(tmp_path / 'qa'))
    playback = fake_playback(engine)
    def local_playback(**kwargs):
        result = list(playback(**kwargs))
        result[-2] = 'local_powerpoint_com'
        return tuple(result)
    monkeypatch.setattr(engine, '_run_validation_and_playback_pass', local_playback)
    probe = Mock(return_value={'is_clean': False, 'error': 'repair dialog observed'})
    monkeypatch.setattr(orchestration, 'verify_clean_powerpoint_open', probe)
    output = tmp_path / 'existing.pptx'
    output.write_bytes(b'existing user file')
    manifest = engine.run_pipeline(str(output), builder_fn=build_deck)
    assert probe.call_args.args[0] == str(output)
    assert output.read_bytes() == b'existing user file'
    assert manifest.status == ValidationStatus.FAIL and manifest.delivered_pptx is None


def test_missing_video_evidence_cannot_pass_remote_gate(tmp_path, monkeypatch):
    engine = PMGOrchestrator(str(tmp_path / 'qa'))
    playback = fake_playback(engine)
    def incomplete_playback(**kwargs):
        result = list(playback(**kwargs))
        result[2] = VideoMetadata('missing.mp4')
        return tuple(result)
    monkeypatch.setattr(engine, '_run_validation_and_playback_pass', incomplete_playback)
    manifest = engine.run_pipeline(str(tmp_path / 'output.pptx'), builder_fn=build_deck)
    assert manifest.status == ValidationStatus.BLOCKED and manifest.delivered_pptx is None


def test_missing_powerpoint_is_never_clean(tmp_path, monkeypatch):
    import pmg.worker.powerpoint_probe as probe
    path = tmp_path / 'source.pptx'
    path.write_bytes(b'fixture')
    monkeypatch.setattr(probe.sys, 'platform', 'linux')
    result = probe.verify_clean_powerpoint_open(str(path))
    assert result['is_clean'] is False and result['skipped'] is True
    assert PP_ALERTS_ALL == 2


def test_probe_timeout_is_bounded_and_fails_closed(tmp_path, monkeypatch):
    import pmg.worker.powerpoint_probe as probe
    path = tmp_path / 'source.pptx'
    path.write_bytes(b'fixture')
    monkeypatch.setattr(probe.sys, 'platform', 'win32')
    monkeypatch.setattr(probe.subprocess, 'run', Mock(side_effect=subprocess.TimeoutExpired('probe', .01)))
    result = probe.verify_clean_powerpoint_open(str(path), timeout_sec=.01)
    assert not result['is_clean'] and result['timed_out'] and result['source_unchanged']


def test_failed_manifest_packaging_never_delivers(tmp_path, monkeypatch):
    import pmg.orchestrator as orchestration
    engine = PMGOrchestrator(str(tmp_path / 'qa'))
    monkeypatch.setattr(orchestration, '_embed_manifest_in_pptx', Mock(side_effect=OSError('disk write failed')))
    manifest = engine.run_pipeline(str(tmp_path / 'output.pptx'), builder_fn=build_deck,
                                   worker_mode='static', max_repair_loops=0)
    assert manifest.status == ValidationStatus.FAIL and manifest.delivered_pptx is None
    assert any('packaging failed' in d['description'] for d in manifest.review_summary['defects'])


def test_installed_provenance_uses_published_vcs_commit(tmp_path, monkeypatch):
    import pmg.version as version
    distribution = Mock()
    distribution.locate_file.return_value = version.__file__
    distribution.read_text.return_value = json.dumps({'vcs_info': {
        'commit_id': 'published-exact-commit', 'requested_revision': 'published-exact-commit'}})
    monkeypatch.setattr(version, 'get_repo_root', lambda: str(tmp_path))
    monkeypatch.setattr(version.metadata, 'distribution', Mock(return_value=distribution))
    assert version.get_git_commit() == 'published-exact-commit'
    assert version.get_git_branch() == 'published-exact-commit'
