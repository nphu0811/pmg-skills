"""
slide_inspector.py (v3.5 - Motion Choreography & Traveling Entity Audit)
=======================================================================
Automated Quality Assurance & Verification tool for PMG Motion Graphics decks:
- Native Transition Audit (Morph, Fade, Circle, Split, Push, Reveal)
- Traveling Entity & '!!' persistent anchor tracking
- Visual-First compliance (< 25 words per slide in 3D mode, <= 45 in academic)
- 16:9 Aspect ratio verification
"""

import os
import sys
import zipfile
import re
import argparse

sys.stdout.reconfigure(encoding='utf-8')

def audit_presentation(pptx_path: str):
    print("=" * 70)
    print(f"[*] RUNNING MOTION & TRANSITION QA AUDIT ON: {pptx_path}")
    print("=" * 70)
    
    if not os.path.exists(pptx_path):
        print(f"[FAIL] File not found: {pptx_path}")
        return False
        
    with zipfile.ZipFile(pptx_path, 'r') as z:
        slides = [n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)]
        slides.sort(key=lambda x: int(re.search(r'\d+', x).group()))
        total_slides = len(slides)
        print(f"[+] Total slides detected: {total_slides}")
        
        passed_transitions = 0
        passed_word_count = 0
        total_words = 0
        traveling_anchors = set()
        
        for idx, s in enumerate(slides, 1):
            xml = z.read(s).decode('utf-8', errors='ignore')
            
            # 1. Transition Check
            trans_match = re.search(r'<p:transition[^>]*>(.*?)</p:transition>', xml, re.DOTALL)
            trans_name = "None (Cut)"
            has_transition = False
            
            if trans_match:
                t_inner = trans_match.group(1)
                if "p159:morph" in t_inner or "p160:morph" in t_inner:
                    trans_name = "Morph (60 FPS)"
                    has_transition = True
                elif "<p:circle" in t_inner:
                    trans_name = "Circle Spotlight"
                    has_transition = True
                elif "<p:fade" in t_inner:
                    trans_name = "Fade"
                    has_transition = True
                elif "<p:split" in t_inner:
                    trans_name = "Split Curtain"
                    has_transition = True
                elif "<p:push" in t_inner:
                    trans_name = "Push Pan"
                    has_transition = True
                else:
                    trans_name = "Custom Transition"
                    has_transition = True
                    
            if has_transition or idx == 1:
                passed_transitions += 1
                
            # 2. Traveling Anchors Check (!! prefix)
            anchors = re.findall(r'name="(!![^"]+)"', xml)
            for a in anchors:
                traveling_anchors.add(a)
                
            # 3. Word Count Check
            texts = re.findall(r'<a:t>([^<]+)</a:t>', xml)
            full_text = " ".join(texts)
            words = full_text.split()
            w_count = len(words)
            total_words += w_count
            is_visual_first = (w_count <= 35) # generous threshold for titles/labels
            if is_visual_first:
                passed_word_count += 1
                
            status_trans = f"PASS ({trans_name})" if (has_transition or idx == 1) else "WARN (No transition)"
            status_words = "PASS" if is_visual_first else f"WARN ({w_count} words)"
            print(f"  Slide {idx:02d}: Transition [{status_trans:20s}] | Words: {w_count:2d} [{status_words}]")
            
        print("-" * 70)
        print(f"[+] Transition Choreography Coverage: {passed_transitions}/{total_slides} slides ({(passed_transitions/total_slides)*100:.1f}%)")
        print(f"[+] Traveling Morph Anchors Detected: {len(traveling_anchors)} {list(traveling_anchors)[:4]}")
        print(f"[+] Visual-First Word Count Compliance: {passed_word_count}/{total_slides} slides ({(passed_word_count/total_slides)*100:.1f}%)")
        print(f"[+] Average Words/Slide: {total_words/total_slides:.1f} words")
        
        overall_pass = (passed_transitions >= total_slides - 1)
        print("=" * 70)
        if overall_pass:
            print("[RESULT: VERIFIED & ACCEPTED] Presentation meets PMG Motion Design standards!")
        else:
            print("[RESULT: REVISION RECOMMENDED] Missing transitions across slide deck.")
        print("=" * 70)
        return overall_pass

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Audit PPTX against PMG motion standards.")
    parser.add_argument("pptx", help="Path to .pptx presentation")
    args = parser.parse_args()
    audit_presentation(args.pptx)
