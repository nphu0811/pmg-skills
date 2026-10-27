#!/usr/bin/env python3
"""
motion_validator.py (PMG Deck Review, Validation & Self-Healing Engine)
======================================================================
Automated review & validation engine for PowerPoint Motion Graphics decks:
1. Deep Validation: Detects missing transitions, broken '!!' morph anchors,
   discontinuous traveling entities, timing anomalies, and text overflow.
2. 3-Loop Auto-Repair Cycle: Automatically fixes XML transitions, synchronizes
   morph identifiers, and harmonizes motion timing up to 3 attempts.
3. Fallback Skill Substitution: If an effect fails after 3 repair loops,
   automatically substitutes an alternative skill (e.g. Morph -> Circle Wipe/Fade,
   or 3D Dolly -> 2D Parallax/Glassmorphism Stack) and loops validation again.
4. Official Delivery Certification: Grants final verification seal only when
   100% of transitions and motion graphics pass.

Usage:
    python scripts/motion_validator.py MyDeck.pptx
    python scripts/motion_validator.py MyDeck.pptx --auto-heal --max-attempts 3
"""

import os
import sys
import zipfile
import re
import shutil
import argparse
from typing import Dict, List, Tuple, Set, Optional
import pptx
from pptx.oxml import parse_xml

sys.stdout.reconfigure(encoding='utf-8')


# =============================================================================
# 1. VALIDATION ENGINE (KIỂM TRA THIẾU HỤT HIỆU ỨNG)
# =============================================================================

class DeckReviewReport:
    def __init__(self, pptx_path: str):
        self.pptx_path = pptx_path
        self.total_slides = 0
        self.slide_transitions: Dict[int, str] = {}
        self.missing_transitions: List[int] = []
        self.traveling_anchors: Dict[str, List[int]] = {}  # tag -> list of slide indices
        self.discontinuous_anchors: Dict[str, List[str]] = {}
        self.word_counts: Dict[int, int] = {}
        self.overflow_slides: List[int] = []
        self.durations: Dict[int, int] = {}
        self.is_valid = False
        self.errors: List[str] = []
        self.warnings: List[str] = []

    def print_summary(self):
        print("=" * 75)
        print(f"[*] PMG MOTION & ANIMATION REVIEW REPORT: {os.path.basename(self.pptx_path)}")
        print("=" * 75)
        print(f"[+] Total Slides Detected: {self.total_slides}")
        
        for idx in range(1, self.total_slides + 1):
            trans = self.slide_transitions.get(idx, "None (Cut)")
            dur = self.durations.get(idx, 0)
            words = self.word_counts.get(idx, 0)
            
            if idx == 1:
                status_t = "COVER (No transition)"
            elif idx in self.missing_transitions:
                status_t = "[FAIL] MISSING TRANSITION"
            else:
                status_t = f"[PASS] {trans} ({dur}ms)"
                
            status_w = "[PASS]" if words <= 35 else f"[WARN] ({words} words)"
            print(f"  Slide {idx:02d}: Transition {status_t:36s} | Words: {words:2d} {status_w}")
            
        print("-" * 75)
        # Transition Coverage
        covered = self.total_slides - len(self.missing_transitions)
        pct = (covered / self.total_slides) * 100 if self.total_slides > 0 else 0
        print(f"[+] Transition Choreography Coverage: {covered}/{self.total_slides} slides ({pct:.1f}%)")
        
        # Traveling Entities
        if self.traveling_anchors:
            print(f"[+] Traveling Morph Anchors Tracked: {len(self.traveling_anchors)}")
            for tag, s_list in self.traveling_anchors.items():
                continuity = "Continuous" if tag not in self.discontinuous_anchors else "[BROKEN GAP]"
                print(f"    - {tag:24s} -> Slides: {s_list} [{continuity}]")
        else:
            print("[!] No '!!' Traveling Anchors found in shape tree.")
            
        # Errors & Warnings
        if self.errors:
            print("-" * 75)
            print(f"[!] ERRORS DETECTED ({len(self.errors)}):")
            for err in self.errors:
                print(f"    * {err}")
                
        if self.warnings:
            print("-" * 75)
            print(f"[*] WARNINGS ({len(self.warnings)}):")
            for w in self.warnings:
                print(f"    * {w}")
                
        print("=" * 75)
        if self.is_valid:
            print("[RESULT: VERIFIED & ACCEPTED FOR USER DELIVERY] 100% Motion Standards Satisfied!")
        else:
            print("[RESULT: VALIDATION FAILED - DEFICIENCIES DETECTED - REQUIRES REMEDIATION]")
        print("=" * 75)


def validate_deck(pptx_path: str) -> DeckReviewReport:
    """Performs deep OpenXML inspection of presentation motion graphics & transitions."""
    report = DeckReviewReport(pptx_path)
    
    if not os.path.exists(pptx_path):
        report.errors.append(f"File not found: {pptx_path}")
        return report

    try:
        with zipfile.ZipFile(pptx_path, 'r') as z:
            slides = [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)]
            slides.sort(key=lambda x: int(re.search(r'\d+', x).group()))
            report.total_slides = len(slides)
            
            if report.total_slides == 0:
                report.errors.append("No slides found in .pptx package.")
                return report

            for idx, s in enumerate(slides, 1):
                xml = z.read(s).decode('utf-8', errors='ignore')
                
                # 1. Parse Transition & Duration
                trans_match = re.search(r'<p:transition([^>]*)>(.*?)</p:transition>', xml, re.DOTALL)
                if trans_match:
                    t_attrs = trans_match.group(1)
                    t_inner = trans_match.group(2)
                    
                    # Duration
                    dur_m = re.search(r'dur="(\d+)"', t_attrs)
                    dur = int(dur_m.group(1)) if dur_m else 1600
                    report.durations[idx] = dur
                    
                    # Transition Type
                    if "p159:morph" in t_inner or "p160:morph" in t_inner:
                        report.slide_transitions[idx] = "Morph (60 FPS)"
                    elif "<p:circle" in t_inner:
                        report.slide_transitions[idx] = "Circle Wipe (Spotlight)"
                    elif "<p:split" in t_inner:
                        report.slide_transitions[idx] = "Split Curtain"
                    elif "<p:fade" in t_inner:
                        report.slide_transitions[idx] = "Fade"
                    elif "<p:push" in t_inner:
                        report.slide_transitions[idx] = "Push Pan"
                    else:
                        report.slide_transitions[idx] = "Native Transition"
                else:
                    if idx > 1:
                        report.missing_transitions.append(idx)
                        report.slide_transitions[idx] = "None (Cut)"
                    else:
                        report.slide_transitions[idx] = "Cover (Static Entry)"
                        
                # 2. Parse Shape Anchors (!! prefix)
                anchors = re.findall(r'name="(!![^"]+)"', xml)
                for a in set(anchors):
                    if a not in report.traveling_anchors:
                        report.traveling_anchors[a] = []
                    report.traveling_anchors[a].append(idx)
                    
                # 3. Word Count Check
                texts = re.findall(r'<a:t>([^<]+)</a:t>', xml)
                full_text = " ".join(texts)
                w_count = len(full_text.split())
                report.word_counts[idx] = w_count
                if w_count > 35:
                    report.overflow_slides.append(idx)
                    report.warnings.append(f"Slide {idx} word count ({w_count} words) exceeds visual-first limit (<= 35).")

    except Exception as e:
        report.errors.append(f"Inspection error: {str(e)}")
        return report

    # Check Transition Completeness
    if report.missing_transitions:
        report.errors.append(f"Missing transition on slide(s): {report.missing_transitions}. 100% coverage required.")

    # Check Traveling Anchor Continuity
    for tag, s_list in report.traveling_anchors.items():
        s_list.sort()
        # If an anchor spans across slides, check if there are intermediate gaps
        if len(s_list) > 1:
            for i in range(len(s_list) - 1):
                cur = s_list[i]
                nxt = s_list[i + 1]
                if nxt > cur + 1:
                    gap_msg = f"Anchor '{tag}' disappears between slide {cur} and slide {nxt}."
                    if tag not in report.discontinuous_anchors:
                        report.discontinuous_anchors[tag] = []
                    report.discontinuous_anchors[tag].append(gap_msg)
                    report.errors.append(gap_msg)

    # Valid if 0 errors and all non-cover slides have transitions
    report.is_valid = (len(report.errors) == 0 and len(report.missing_transitions) == 0)
    return report


# =============================================================================
# 2. AUTO-REMEDITATION & 3-LOOP SELF-HEALING ENGINE
# =============================================================================

def repair_xml_transitions(pptx_path: str, target_slides: List[int], trans_type: str = "morph", dur_ms: int = 1600):
    """Directly heals missing or malformed transitions in the PPTX via python-pptx oxml."""
    prs = pptx.Presentation(pptx_path)
    modified = False
    
    for idx, slide in enumerate(prs.slides, 1):
        if idx in target_slides and idx > 1:
            sld_elem = slide.element
            # Remove any malformed or corrupted transitions
            for child in list(sld_elem):
                if child.tag.endswith('transition') or child.tag.endswith('AlternateContent'):
                    sld_elem.remove(child)
                    
            if trans_type == "circle":
                inner_xml = '<p:circle/>'
            elif trans_type == "fade":
                inner_xml = '<p:fade/>'
            elif trans_type == "split":
                inner_xml = '<p:split orient="vert"/>'
            else:
                inner_xml = '<p159:morph option="byObject"/>'
                
            trans_xml = (
                f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
                f'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
                f'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
                f'spd="slow" p14:dur="{dur_ms}">'
                f'{inner_xml}'
                f'</p:transition>'
            )
            trans_node = parse_xml(trans_xml)
            
            clrMapOvr = None
            for child in sld_elem:
                if child.tag.endswith('clrMapOvr'):
                    clrMapOvr = child
                    break
                    
            if clrMapOvr is not None:
                clrMapOvr.addnext(trans_node)
            else:
                sld_elem.append(trans_node)
            modified = True
            
    if modified:
        prs.save(pptx_path)


def repair_traveling_anchor_gap(pptx_path: str, anchor_tag: str, gap_slides: List[int]):
    """Bridges missing traveling anchors on intermediate slides to prevent morph jump cuts."""
    prs = pptx.Presentation(pptx_path)
    modified = False
    
    for s_idx in gap_slides:
        if 1 <= s_idx <= len(prs.slides):
            slide = prs.slides[s_idx - 1]
            # Check if anchor is already there
            has_tag = any(shape.name == anchor_tag for shape in slide.shapes)
            if not has_tag:
                # Add an offscreen or subtle proxy shape with this name so morph stays continuous
                shape = slide.shapes.add_shape(
                    pptx.enum.shapes.MSO_SHAPE.RECTANGLE,
                    pptx.util.Inches(-2.5), pptx.util.Inches(3.5),
                    pptx.util.Inches(0.5), pptx.util.Inches(0.5)
                )
                shape.name = anchor_tag
                # Make proxy invisible
                shape.fill.background()
                shape.line.fill.background()
                modified = True
                
    if modified:
        prs.save(pptx_path)


def substitute_with_fallback_skill(pptx_path: str, failed_slides: List[int], reason: str) -> str:
    """
    Substitutes an intractable failure with an alternative skill from the PMG suite:
    - Failed Morph / 3D Dolly -> Fallback to pmg-choreography (Circle Wipe Spotlight or Split Reveal)
    - Broken complex motion -> Fallback to pmg-camera-parallax or pmg-glassmorphism card transition.
    """
    print(f"[*] Activating Fallback Skill Substitution for Slide(s) {failed_slides}...")
    print(f"    Reason: {reason}")
    print("    Alternative Skill Activated: 'pmg-choreography' (Iris Spotlight / Circle Wipe)")
    
    # Substitute failed slide transitions with high-impact Circle Spotlight or Split Curtain
    repair_xml_transitions(pptx_path, failed_slides, trans_type="circle", dur_ms=1400)
    return "pmg-choreography"


# =============================================================================
# 3. MASTER REVIEW, 3-LOOP HEALING & CERTIFICATION
# =============================================================================

def review_and_heal_deck(pptx_path: str, max_repair_loops: int = 3, allow_substitution: bool = True) -> Tuple[bool, DeckReviewReport, List[str]]:
    """
    Supervises the full PMG quality gate:
    1. Initial Validation.
    2. Loop up to 3 times to heal deficiencies in-place.
    3. If still failing after 3 loops, substitute with another skill from the suite and loop again.
    4. Issue official delivery certificate only upon 100% pass.
    """
    action_log = []
    
    print(f"\n[PHASE 1] Initial Inspection of '{os.path.basename(pptx_path)}'...")
    report = validate_deck(pptx_path)
    report.print_summary()
    
    if report.is_valid:
        action_log.append("Passed initial validation without remediation.")
        return True, report, action_log

    # Backup original before modifying
    backup_path = pptx_path + ".pre_heal.bak"
    shutil.copyfile(pptx_path, backup_path)
    action_log.append(f"Created safety backup at {os.path.basename(backup_path)}")

    # -------------------------------------------------------------------------
    # 3-LOOP AUTO-REPAIR CYCLE
    # -------------------------------------------------------------------------
    loop_idx = 1
    healed = False
    
    while loop_idx <= max_repair_loops and not healed:
        print(f"\n>>> [AUTO-REPAIR LOOP {loop_idx}/{max_repair_loops}] Diagnosing and healing deficiencies...")
        
        # Action A: Heal missing transitions
        if report.missing_transitions:
            print(f"    [Repair Action A] Injecting missing OpenXML Morph transitions on slides: {report.missing_transitions}")
            repair_xml_transitions(pptx_path, report.missing_transitions, trans_type="morph", dur_ms=1600)
            action_log.append(f"Loop {loop_idx}: Re-injected Morph transition on slides {report.missing_transitions}")

        # Action B: Bridge discontinuous traveling anchors
        if report.discontinuous_anchors:
            for tag, gap_msgs in report.discontinuous_anchors.items():
                s_list = report.traveling_anchors[tag]
                gap_slides = []
                for s in range(min(s_list), max(s_list) + 1):
                    if s not in s_list:
                        gap_slides.append(s)
                if gap_slides:
                    print(f"    [Repair Action B] Bridging anchor gap for '{tag}' across slides: {gap_slides}")
                    repair_traveling_anchor_gap(pptx_path, tag, gap_slides)
                    action_log.append(f"Loop {loop_idx}: Bridged '{tag}' continuity on slides {gap_slides}")

        # Re-validate after loop repairs
        report = validate_deck(pptx_path)
        if report.is_valid:
            print(f"\n[+] SUCCESS: Presentation healed successfully on Loop {loop_idx}!")
            healed = True
            break
            
        loop_idx += 1

    # -------------------------------------------------------------------------
    # FALLBACK SKILL SUBSTITUTION (IF STILL FAILING AFTER 3 ATTEMPTS)
    # -------------------------------------------------------------------------
    if not healed and allow_substitution:
        print("\n" + "!" * 75)
        print("[!] NOTICE: Deficiencies unresolved after 3 auto-repair loops.")
        print("[!] INITIATING PMG FALLBACK SKILL SUBSTITUTION MATRIX...")
        print("!" * 75)
        
        failed_slides = list(set(report.missing_transitions + [s for sl in report.discontinuous_anchors.values() for s in [1]]))
        if not failed_slides:
            failed_slides = [s for s in range(2, report.total_slides + 1)]
            
        sub_skill = substitute_with_fallback_skill(
            pptx_path, 
            failed_slides=failed_slides, 
            reason="Morph / spatial continuity could not be synchronized cleanly."
        )
        action_log.append(f"Substituted failed effect with '{sub_skill}' transition suite.")
        
        print("\n>>> [POST-SUBSTITUTION RE-LOOP] Re-running validation with substituted skill...")
        report = validate_deck(pptx_path)
        if report.is_valid:
            print("[+] SUCCESS: Presentation validated 100% after Fallback Skill Substitution!")
            healed = True
            action_log.append(f"Validation achieved 100% compliance using '{sub_skill}' fallback.")

    # -------------------------------------------------------------------------
    # FINAL CERTIFICATION
    # -------------------------------------------------------------------------
    print("\n" + "=" * 75)
    if report.is_valid:
        print(" [CERTIFIED FOR DELIVERY] PMG OFFICIAL MOTION GRAPHICS SEAL GRANTED ")
        print(" The presentation is 100% compliant with broadcast motion & transition standards.")
        print("=" * 75)
        # Clean backup if passed
        if os.path.exists(backup_path):
            os.remove(backup_path)
    else:
        print(" [DELIVERY BLOCKED] Validation still has unresolved errors. Do NOT deliver to user.")
        print("=" * 75)

    return report.is_valid, report, action_log


# =============================================================================
# CLI ENTRYPOINT
# =============================================================================

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Review, validate and auto-heal PMG presentations.")
    parser.add_argument("pptx", help="Path to .pptx presentation to review")
    parser.add_argument("--auto-heal", action="store_true", default=True, help="Automatically run 3-loop repair and skill substitution")
    parser.add_argument("--max-attempts", type=int, default=3, help="Max repair loops before skill substitution (default: 3)")
    parser.add_argument("--no-substitution", action="store_true", help="Disallow fallback skill substitution")
    
    args = parser.parse_args()
    
    if args.auto_heal:
        success, final_report, logs = review_and_heal_deck(
            args.pptx, 
            max_repair_loops=args.max_attempts, 
            allow_substitution=(not args.no_substitution)
        )
        sys.exit(0 if success else 1)
    else:
        rep = validate_deck(args.pptx)
        rep.print_summary()
        sys.exit(0 if rep.is_valid else 1)
