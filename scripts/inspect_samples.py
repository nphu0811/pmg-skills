import sys, os, zipfile, re
sys.stdout.reconfigure(encoding='utf-8')

def analyze_pptx(filename):
    path = os.path.join('DataPowerPoint', filename)
    with zipfile.ZipFile(path) as z:
        slides = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)],
                        key=lambda x: int(re.search(r'\d+', x).group()))
        print(f"================ {filename} ({len(slides)} slides) ================")
        for i, s in enumerate(slides[:10]):
            xml = z.read(s).decode('utf-8', errors='ignore')
            trans = re.findall(r'<p:transition[^>]*>(.*?)</p:transition>', xml, re.DOTALL)
            trans_str = trans[0].strip() if trans else 'None'
            morph_shapes = re.findall(r'name="(!![^"]+)"', xml)
            all_shapes = re.findall(r'name="([^"]+)"', xml)
            srgb = set(re.findall(r'val="([0-9A-Fa-f]{6})"', xml))
            anims = re.findall(r'<p:(animMotion|animEffect|animScale|animRot|cMediaNode)', xml)
            print(f"  Slide {i+1}: Trans={trans_str[:40]} | !!Shapes={morph_shapes[:4]} | ShapesCount={len(all_shapes)} | Anims={len(anims)}")

def inspect_group_coords(filename, group_name, slide_indices):
    path = os.path.join('DataPowerPoint', filename)
    with zipfile.ZipFile(path) as z:
        print(f"\n--- Tracking {group_name} in {filename} ---")
        for idx in slide_indices:
            s_name = f'ppt/slides/slide{idx}.xml'
            if s_name not in z.namelist():
                continue
            xml = z.read(s_name).decode('utf-8', errors='ignore')
            # Look for shape or group with name
            pattern = rf'name="{re.escape(group_name)}".*?<a:off x="(-?\d+)" y="(-?\d+)".*?<a:ext cx="(\d+)" cy="(\d+)"'
            m = re.search(pattern, xml, re.DOTALL)
            if m:
                print(f"Slide {idx}: x={m.group(1)} EMU ({int(m.group(1))/914400:.2f} in), y={m.group(2)} EMU ({int(m.group(2))/914400:.2f} in), w={m.group(3)} EMU, h={m.group(4)} EMU")
            else:
                print(f"Slide {idx}: {group_name} not found or different structure")

inspect_group_coords('Vợ-Chồng-A-Phủ.pptx', '!!Group 60', [3, 4, 5])
inspect_group_coords('Vợ-Chồng-A-Phủ.pptx', '!!Group 10', [5, 6, 7])

def inspect_transitions_all(filename):
    path = os.path.join('DataPowerPoint', filename)
    with zipfile.ZipFile(path) as z:
        slides = sorted([n for n in z.namelist() if re.match(r'ppt/slides/slide\d+\.xml', n)],
                        key=lambda x: int(re.search(r'\d+', x).group()))
        print(f"\n--- All transitions in {filename} ---")
        for i, s in enumerate(slides):
            xml = z.read(s).decode('utf-8', errors='ignore')
            m_trans = re.search(r'<p:transition[^>]*>(.*?)</p:transition>', xml, re.DOTALL)
            m_spd = re.search(r'<p:transition[^>]*spd="([^"]+)"', xml)
            m_dur = re.search(r'<p14:dur[^>]*val="([^"]+)"', xml)
            t_content = m_trans.group(1).strip() if m_trans else 'Cut/None'
            spd = m_spd.group(1) if m_spd else ''
            dur = m_dur.group(1) if m_dur else ''
            # check for timing animations
            anims = re.findall(r'<p:(animMotion|animEffect|animScale|animRot|cMediaNode)', xml)
            if t_content != 'Cut/None' or anims:
                print(f"  Slide {i+1}: Transition=[{t_content}] spd={spd} dur={dur} | InSlideAnims={len(anims)}")

inspect_transitions_all('Vợ-Chồng-A-Phủ.pptx')
inspect_transitions_all('MAC-LENIN.pptx')


