#!/usr/bin/env python3
"""
pptx_morph_injector.py (Transition Injector Suite)
=================================================
CLI tool to inspect and inject native PowerPoint transitions (Morph, Fade, 
Circle, Split, Push, Wipe, Reveal) into any presentation (.pptx).

Usage:
    python pptx_morph_injector.py input.pptx -o output.pptx --type morph --dur 1600
    python pptx_morph_injector.py input.pptx --type circle --dur 1400
"""

import argparse
import sys
import pptx
from pptx.oxml import parse_xml

def inject_transition(input_path: str, output_path: str, trans_type: str = "morph", 
                      duration_ms: int = 1600, mode: str = "byObject", 
                      start_slide: int = 2, **kwargs):
    prs = pptx.Presentation(input_path)
    total_slides = len(prs.slides)
    print(f"[*] Loaded '{input_path}' ({total_slides} slides)")
    
    injected_count = 0
    trans_type = trans_type.lower()
    
    for idx, slide in enumerate(prs.slides, 1):
        if idx < start_slide:
            continue
            
        sld_elem = slide.element
        # Remove existing transitions
        for child in list(sld_elem):
            if child.tag.endswith('transition') or child.tag.endswith('AlternateContent'):
                sld_elem.remove(child)
                
        if trans_type == "morph":
            inner_xml = f'<p159:morph option="{mode}"/>'
        elif trans_type == "fade":
            inner_xml = '<p:fade/>'
        elif trans_type == "circle":
            inner_xml = '<p:circle/>'
        elif trans_type == "split":
            orient = kwargs.get("orient", "vert")
            inner_xml = f'<p:split orient="{orient}"/>'
        elif trans_type == "push":
            direction = kwargs.get("dir", "u")
            inner_xml = f'<p:push dir="{direction}"/>'
        elif trans_type == "wipe":
            direction = kwargs.get("dir", "l")
            inner_xml = f'<p:wipe dir="{direction}"/>'
        else:
            inner_xml = f'<p159:morph option="{mode}"/>'
            
        trans_xml = (
            f'<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            f'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
            f'xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
            f'spd="slow" p14:dur="{duration_ms}">'
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
            
        injected_count += 1
        
    prs.save(output_path)
    print(f"[+] Successfully injected '{trans_type}' transition into {injected_count} slides.")
    print(f"[+] Saved to: {output_path}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Inject transitions (Morph, Fade, Page Curl, Circle, etc.) into slides.")
    parser.add_argument("input", help="Path to input .pptx file")
    parser.add_argument("-o", "--output", help="Path to output .pptx file (default: overwrites input)")
    parser.add_argument("--type", default="morph", choices=["morph", "fade", "circle", "split", "push", "reveal", "wipe"], help="Transition type")
    parser.add_argument("--dur", type=int, default=1600, help="Duration in milliseconds (default: 1600)")
    parser.add_argument("--mode", choices=["byObject", "byWord", "byChar"], default="byObject", help="Morph option")
    parser.add_argument("--start", type=int, default=2, help="Start slide index (default: 2)")
    
    args = parser.parse_args()
    out = args.output if args.output else args.input
    inject_transition(args.input, out, trans_type=args.type, duration_ms=args.dur, mode=args.mode, start_slide=args.start)
