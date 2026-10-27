"""
build_showcase_deck.py
======================
Builds the official PMG demonstration presentation:
- Bright Modern / Academic Sky Cyan theme
- Continuous Traveling Entity (Origami Crane swooping across slides)
- PowerPoint Native Transitions Suite (Morph, Circle Wipe, Fade, Split)
- Renders 60 FPS MP4 video to examples/pmg_motion_showcase.mp4
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from scripts.motion_graph_kit import MotionDeck, Themes, export_deck_to_video

sys.stdout.reconfigure(encoding='utf-8')

def build_showcase():
    print("[*] Initializing MotionDeck with LIGHT_ACADEMIC_CYAN theme...")
    deck = MotionDeck(theme=Themes.LIGHT_ACADEMIC_CYAN)
    crane_asset = os.path.abspath("assets/origami_crane.png")
    
    # =========================================================================
    # SLIDE 1: TRANG BÌA — THE APERTURE UNVEIL
    # =========================================================================
    s1 = deck.add_slide(transition="none")
    
    # Header Badge
    deck.add_badge(s1, "NGHIÊN CỨU HỌC THUẬT // DỰ ÁN MÔI TRƯỜNG 2026", 
                   left=3.66, top=1.4, width=6.0, height=0.42, name="!!badge_top",
                   bg_color=RGBColor(2, 132, 199), text_color=RGBColor(255, 255, 255))
    
    # Title Box
    deck.add_hero_title(s1, "PHÂN LOẠI RÁC THẢI SINH HOẠT", 
                        subtitle="Khảo sát nhận thức sinh viên và giải pháp nâng cao thói quen bền vững",
                        left=2.5, top=2.2, width=8.33, height=2.2,
                        title_size=38, subtitle_size=14,
                        title_color=RGBColor(15, 23, 42), name="!!title_main")
    
    # Center Frosted Card with Presenter Information
    c1 = deck.add_glass_card(s1, left=4.16, top=4.5, width=5.0, height=1.6, name="!!card_center",
                             border_width_pt=1.5, corner_radius=6000)
    tx_meta = s1.shapes.add_textbox(Inches(4.3), Inches(4.7), Inches(4.7), Inches(1.2))
    tf_meta = tx_meta.text_frame
    tf_meta.word_wrap = True
    p1 = tf_meta.paragraphs[0]
    p1.text = "NHÓM THỰC HIỆN: F6 HUTECH  |  LỚP: Q05"
    p1.font.name = "Segoe UI"
    p1.font.size = Pt(11)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(2, 132, 199)
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf_meta.add_paragraph()
    p2.space_before = Pt(4)
    p2.text = "Giảng viên hướng dẫn & Hội đồng khoa học"
    p2.font.name = "Segoe UI"
    p2.font.size = Pt(10)
    p2.font.color.rgb = RGBColor(71, 85, 105)
    p2.alignment = PP_ALIGN.CENTER
    
    # Traveling Crane staged on the left margin (poised for takeoff)
    deck.add_traveling_entity(s1, crane_asset, left=0.6, top=4.0, width=2.0, height=2.0, rot_deg=20, name="!!Traveling_Crane")
    
    # Bottom Progress Bar
    deck.add_progress_bar(s1, current_step=1, total_steps=5)

    # =========================================================================
    # SLIDE 2: THE CRANE TAKES FLIGHT — MORPH TRANSITION (1600ms)
    # =========================================================================
    s2 = deck.add_slide(transition="morph", duration_ms=1600)
    
    deck.add_badge(s2, "01 // THỰC TRẠNG KHẢO SÁT", left=1.0, top=0.9, width=3.2, height=0.38, name="!!badge_top",
                   bg_color=RGBColor(2, 132, 199), text_color=RGBColor(255, 255, 255))
    
    deck.add_hero_title(s2, "VẤN ĐỀ XUẤT HIỆN XUNG QUANH", 
                        subtitle="97.7% người học xác nhận tình trạng phân loại rác chưa đúng quy chuẩn",
                        left=1.0, top=1.45, width=9.0, height=1.3,
                        title_size=28, subtitle_size=12.5,
                        title_color=RGBColor(15, 23, 42), name="!!title_main")
    
    # Metric Cards Left
    deck.add_metric_card(s2, "TỶ LỆ NHẬN DIỆN THỰC TẾ", "97.7%", 
                         "Sinh viên nhận thấy khó khăn ngay tại khuôn viên học tập",
                         left=1.0, top=2.8, width=5.2, height=1.7, 
                         accent_color=RGBColor(2, 132, 199), name="!!card_left")
                         
    deck.add_metric_card(s2, "MỨC ĐỘ NGHIÊM TRỌNG", "81.8%", 
                         "Đánh giá vấn đề ở mức Nghiêm trọng hoặc Rất nghiêm trọng",
                         left=1.0, top=4.7, width=5.2, height=1.7, 
                         accent_color=RGBColor(239, 68, 68), name="!!card_sub")
    
    # Right Showcase Card
    c_right = deck.add_glass_card(s2, left=6.6, top=2.8, width=5.7, height=3.6, name="!!card_center",
                                  border_width_pt=1.5, corner_radius=7000)
    tx_r = s2.shapes.add_textbox(Inches(7.0), Inches(3.1), Inches(4.9), Inches(3.0))
    tf_r = tx_r.text_frame
    tf_r.word_wrap = True
    pr1 = tf_r.paragraphs[0]
    pr1.text = "NGUYÊN NHÂN CỐT LÕI"
    pr1.font.name = "Segoe UI"
    pr1.font.size = Pt(13)
    pr1.font.bold = True
    pr1.font.color.rgb = RGBColor(2, 132, 199)
    
    bullet_items = [
        ("38.6%", "Thiếu ý thức và thói quen sinh hoạt hằng ngày"),
        ("29.5%", "Chưa biết phương pháp phân loại rác đúng"),
        ("15.9%", "Thiếu hệ thống thùng rác chuyên dụng tiện lợi"),
        ("15.9%", "Ít được tuyên truyền hướng dẫn thực tế")
    ]
    for pct, desc in bullet_items:
        bp = tf_r.add_paragraph()
        bp.space_before = Pt(6)
        bp.text = f"• {pct} — {desc}"
        bp.font.name = "Segoe UI"
        bp.font.size = Pt(10.5)
        bp.font.color.rgb = RGBColor(51, 65, 85)
        
    # Traveling Crane soaring smoothly across center canvas into the right quadrant!
    deck.add_traveling_entity(s2, crane_asset, left=6.0, top=1.3, width=1.9, height=1.9, rot_deg=-12, name="!!Traveling_Crane")
    
    deck.add_progress_bar(s2, current_step=2, total_steps=5)

    # =========================================================================
    # SLIDE 3: CHUYỂN PHẦN ĐIỀU TRA — CIRCLE WIPE (Iris Spotlight Focus)
    # =========================================================================
    s3 = deck.add_slide(transition="circle", duration_ms=1400)
    
    deck.add_badge(s3, "02 // ĐIỀU TRA THỰC ĐỊA", left=1.0, top=0.9, width=3.2, height=0.38, name="!!badge_top",
                   bg_color=RGBColor(16, 185, 129), text_color=RGBColor(255, 255, 255))
                   
    deck.add_hero_title(s3, "GÓC NHÌN TỪ CHUYÊN GIA & NGƯỜI TRONG CUỘC", 
                        subtitle="Ghi nhận trực tiếp từ giảng viên và cộng tác viên trường học",
                        left=1.0, top=1.45, width=9.5, height=1.3,
                        title_size=28, subtitle_size=12.5,
                        title_color=RGBColor(15, 23, 42), name="!!title_main")
                        
    # Person 1 Card
    card_p1 = deck.add_glass_card(s3, left=1.0, top=2.8, width=5.3, height=3.6, name="!!card_left",
                                  border_hex="10B981", border_width_pt=1.5, corner_radius=6000)
    tx_p1 = s3.shapes.add_textbox(Inches(1.3), Inches(3.1), Inches(4.7), Inches(3.0))
    tf_p1 = tx_p1.text_frame
    tf_p1.word_wrap = True
    pp1 = tf_p1.paragraphs[0]
    pp1.text = "NGUYỄN TRUNG HIẾU"
    pp1.font.name = "Segoe UI"
    pp1.font.size = Pt(14)
    pp1.font.bold = True
    pp1.font.color.rgb = RGBColor(6, 78, 59)
    
    pp2 = tf_p1.add_paragraph()
    pp2.text = "Giảng viên Tiếng Anh — Saigon Campus"
    pp2.font.name = "Segoe UI"
    pp2.font.size = Pt(10)
    pp2.font.bold = True
    pp2.font.color.rgb = RGBColor(16, 185, 129)
    
    pp3 = tf_p1.add_paragraph()
    pp3.space_before = Pt(8)
    pp3.text = "“Thói quen phân loại cần được rèn luyện từ môi trường giáo dục. Cần lồng ghép hướng dẫn vào các chuyên đề học thuật thường xuyên.”"
    pp3.font.name = "Segoe UI"
    pp3.font.size = Pt(11)
    pp3.font.color.rgb = RGBColor(51, 65, 85)
    
    # Person 2 Card
    card_p2 = deck.add_glass_card(s3, left=6.8, top=2.8, width=5.5, height=3.6, name="!!card_center",
                                  border_hex="0284C7", border_width_pt=1.5, corner_radius=6000)
    tx_p2 = s3.shapes.add_textbox(Inches(7.1), Inches(3.1), Inches(4.9), Inches(3.0))
    tf_p2 = tx_p2.text_frame
    tf_p2.word_wrap = True
    pq1 = tf_p2.paragraphs[0]
    pq1.text = "PHẠM ĐĂNG HẢI ANH"
    pq1.font.name = "Segoe UI"
    pq1.font.size = Pt(14)
    pq1.font.bold = True
    pq1.font.color.rgb = RGBColor(2, 132, 199)
    
    pq2 = tf_p2.add_paragraph()
    pq2.text = "Cộng tác viên — Sinh viên thực tế"
    pq2.font.name = "Segoe UI"
    pq2.font.size = Pt(10)
    pq2.font.bold = True
    pq2.font.color.rgb = RGBColor(37, 99, 235)
    
    pq3 = tf_p2.add_paragraph()
    pq3.space_before = Pt(8)
    pq3.text = "“Rất nhiều bạn muốn phân loại nhưng xung quanh chỉ có 1 thùng rác chung. Khi có đủ thùng tái chế riêng biệt, mọi người sẽ tự giác làm đúng.”"
    pq3.font.name = "Segoe UI"
    pq3.font.size = Pt(11)
    pq3.font.color.rgb = RGBColor(51, 65, 85)
    
    # Crane hovers gently over the investigation
    deck.add_traveling_entity(s3, crane_asset, left=10.5, top=1.4, width=1.8, height=1.8, rot_deg=10, name="!!Traveling_Crane")
    
    deck.add_progress_bar(s3, current_step=3, total_steps=5)

    # =========================================================================
    # SLIDE 4: GIẢI PHÁP ĐỀ XUẤT — MORPH CONTINUOUS DOCKING (1600ms)
    # =========================================================================
    s4 = deck.add_slide(transition="morph", duration_ms=1600)
    
    deck.add_badge(s4, "03 // HỆ THỐNG GIẢI PHÁP", left=1.0, top=0.9, width=3.2, height=0.38, name="!!badge_top",
                   bg_color=RGBColor(2, 132, 199), text_color=RGBColor(255, 255, 255))
                   
    deck.add_hero_title(s4, "4 TRỤ CỘT HÀNH ĐỘNG THIẾT THỰC", 
                        subtitle="100% sinh viên tham gia khảo sát bày tỏ mong muốn giải quyết triệt để",
                        left=1.0, top=1.45, width=9.5, height=1.3,
                        title_size=28, subtitle_size=12.5,
                        title_color=RGBColor(15, 23, 42), name="!!title_main")
                        
    # 4 Solution Pillar Cards
    pillars = [
        ("01", "TUYÊN TRUYỀN TRỰC QUAN", "Infographic & video ngắn hướng dẫn phân loại dễ nhớ", RGBColor(2, 132, 199)),
        ("02", "HẠ TẦNG THÙNG RÁC", "Bố trí đồng bộ 3 ngăn: Hữu cơ - Tái chế - Vô cơ", RGBColor(16, 185, 129)),
        ("03", "QUY CHẾ & KHÍCH LỆ", "Tích điểm đổi quà xanh cho sinh viên phân loại đúng", RGBColor(245, 158, 11)),
        ("04", "KẾT NỐI TÁI CHẾ", "Phối hợp các đơn vị thu mua ve chai, tái chế tuần hoàn", RGBColor(124, 58, 237))
    ]
    card_w = 2.65
    gap = 0.28
    for i, (num, p_title, p_desc, col) in enumerate(pillars):
        c_x = 1.0 + i * (card_w + gap)
        c_pil = deck.add_glass_card(s4, left=c_x, top=2.8, width=card_w, height=3.6, name=f"!!card_pillar_{i}",
                                    border_hex=f"{col[0]:02X}{col[1]:02X}{col[2]:02X}", border_width_pt=1.5, corner_radius=6000)
        tx_pil = s4.shapes.add_textbox(Inches(c_x + 0.18), Inches(3.0), Inches(card_w - 0.36), Inches(3.2))
        tf_pil = tx_pil.text_frame
        tf_pil.word_wrap = True
        
        p_num = tf_pil.paragraphs[0]
        p_num.text = num
        p_num.font.name = "Segoe UI"
        p_num.font.size = Pt(24)
        p_num.font.bold = True
        p_num.font.color.rgb = col
        
        p_t = tf_pil.add_paragraph()
        p_t.space_before = Pt(4)
        p_t.text = p_title
        p_t.font.name = "Segoe UI"
        p_t.font.size = Pt(11)
        p_t.font.bold = True
        p_t.font.color.rgb = RGBColor(15, 23, 42)
        
        p_d = tf_pil.add_paragraph()
        p_d.space_before = Pt(6)
        p_d.text = p_desc
        p_d.font.name = "Segoe UI"
        p_d.font.size = Pt(10)
        p_d.font.color.rgb = RGBColor(71, 85, 105)
        
    # Crane swoops down to highlight the 4th pillar
    deck.add_traveling_entity(s4, crane_asset, left=10.4, top=4.4, width=1.8, height=1.8, rot_deg=-20, name="!!Traveling_Crane")
    
    deck.add_progress_bar(s4, current_step=4, total_steps=5)

    # =========================================================================
    # SLIDE 5: KẾT LUẬN & THÔNG ĐIỆP — CIRCLE SPOTLIGHT WIPE (1600ms)
    # =========================================================================
    s5 = deck.add_slide(transition="circle", duration_ms=1600)
    
    deck.add_badge(s5, "KẾT LUẬN & CAM KẾT HÀNH ĐỘNG", left=4.46, top=1.2, width=4.4, height=0.42, name="!!badge_top",
                   bg_color=RGBColor(2, 132, 199), text_color=RGBColor(255, 255, 255))
                   
    # Central Grand Showcase Card
    c_final = deck.add_glass_card(s5, left=2.4, top=1.8, width=8.53, height=4.6, name="!!card_center",
                                  border_hex="0284C7", border_width_pt=2.0, corner_radius=7000)
                                  
    tx_fin = s5.shapes.add_textbox(Inches(2.8), Inches(2.2), Inches(7.73), Inches(3.6))
    tf_fin = tx_fin.text_frame
    tf_fin.word_wrap = True
    
    pf1 = tf_fin.paragraphs[0]
    pf1.alignment = PP_ALIGN.CENTER
    pf1.text = "PHÂN LOẠI RÁC THẢI KHÔNG CHỈ Ở THÙNG RÁC,"
    pf1.font.name = "Segoe UI"
    pf1.font.size = Pt(22)
    pf1.font.bold = True
    pf1.font.color.rgb = RGBColor(15, 23, 42)
    
    pf2 = tf_fin.add_paragraph()
    pf2.alignment = PP_ALIGN.CENTER
    pf2.text = "MÀ BẮT ĐẦU TỪ NHẬN THỨC VÀ THÓI QUEN."
    pf2.font.name = "Segoe UI"
    pf2.font.size = Pt(22)
    pf2.font.bold = True
    pf2.font.color.rgb = RGBColor(2, 132, 199)
    
    pf3 = tf_fin.add_paragraph()
    pf3.alignment = PP_ALIGN.CENTER
    pf3.space_before = Pt(14)
    pf3.text = "“Một hành động nhỏ mỗi ngày — Kiến tạo môi trường học đường văn minh, bền vững.”"
    pf3.font.name = "Segoe UI"
    pf3.font.size = Pt(13)
    pf3.font.color.rgb = RGBColor(71, 85, 105)
    
    pf4 = tf_fin.add_paragraph()
    pf4.alignment = PP_ALIGN.CENTER
    pf4.space_before = Pt(18)
    pf4.text = "TRÂN TRỌNG CẢM ƠN THẦY CÔ & CÁC BẠN ĐÃ LẮNG NGHE"
    pf4.font.name = "Segoe UI"
    pf4.font.size = Pt(12)
    pf4.font.bold = True
    pf4.font.color.rgb = RGBColor(16, 185, 129)
    
    # Crane hovers high in victory flight at top-right
    deck.add_traveling_entity(s5, crane_asset, left=10.6, top=1.0, width=2.1, height=2.1, rot_deg=15, name="!!Traveling_Crane")
    
    deck.add_progress_bar(s5, current_step=5, total_steps=5)

    # Save presentation to temporary/render file
    output_pptx = os.path.abspath("examples/pmg_motion_showcase.pptx")
    deck.save(output_pptx)
    print(f"[+] Showcase presentation created: {output_pptx}")
    
    # Export to 60 FPS MP4 Video
    output_mp4 = os.path.abspath("examples/pmg_motion_showcase.mp4")
    print(f"[*] Rendering high-definition 60 FPS video to: {output_mp4}")
    export_deck_to_video(output_pptx, output_mp4, slide_duration=3.5, fps=60, resolution=1080)
    
    # Clean up the intermediate pptx per user requirement: "xoá luôn cái pptx hiện tại. cái example thì quay video đi, k dùng ảnh nữa"
    if os.path.exists(output_pptx):
        os.remove(output_pptx)
        print(f"[+] Cleaned up temporary pptx; video artifact preserved in: {output_mp4}")

if __name__ == "__main__":
    build_showcase()
