"""
create_camau_deck.py
====================
Tạo slide thuyết trình về đề tài "CÀ MAU — ĐẤT MŨI PHƯƠNG NAM"
Vận dụng đầy đủ các kỹ năng PMG mới:
- Tone màu sáng sinh thái (Fresh Eco Emerald & Aquatic Cyan: #F6FBF7, #059669, #0284C7)
- Chữ nét đậm tương phản cao (#064E3B, #0F172A)
- Vật thể dẫn đường bay xuyên suốt qua các slide (!!Traveling_Bird / Cánh chim Đất Mũi)
- Chuỗi chuyển trang đa dạng: Morph -> Page Curl (Lật trang du khảo) -> Morph -> Circle Wipe (Ống kính Mũi Tàu)
- Bố cục kính mờ Frosted Acrylic sáng, zero-overlap.
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

from scripts.motion_graph_kit import MotionDeck, Themes
from scripts.slide_inspector import audit_presentation

sys.stdout.reconfigure(encoding='utf-8')

def build_camau_presentation(output_path: str = "Ca_Mau_Dat_Mui_Phuong_Nam.pptx"):
    print("[*] Khởi tạo MotionDeck với chủ đề Cà Mau (Light Eco Emerald)...")
    
    # Palette chuyên biệt cho Cà Mau: Rừng ngập mặn & Biển Tây - Biển Đông
    THEME_CAMAU = {
        "name": "Ca Mau Eco Aquatics",
        "is_dark": False,
        "bg_hex": "F4FBF7",                   # Sắc trắng ngọc bích non dịu mát
        "card_bg_hex": "FFFFFF",              # Nền card trắng phủ mờ acrylic
        "card_alpha": "93000",                # 93% mờ tinh khiết
        "border_hex": "059669",               # Viền xanh ngọc rừng ngập mặn
        "border_alpha": "55000",
        "accent": RGBColor(5, 150, 105),      # #059669 Xanh Rừng Đước
        "accent_secondary": RGBColor(2, 132, 199), # #0284C7 Xanh Biển Đất Mũi
        "accent_tertiary": RGBColor(217, 119, 6),   # #D97706 Vàng Đất Phù Sa
        "text_primary": RGBColor(6, 78, 59),  # #064E3B Chữ xanh rêu than đậm cực nét
        "text_secondary": RGBColor(51, 65, 85), # #334155 Xám than bổ trợ
        "text_muted": RGBColor(100, 116, 139),
        "font_heading": "Segoe UI",
        "font_body": "Segoe UI"
    }

    deck = MotionDeck(theme=THEME_CAMAU)
    bird_asset = os.path.abspath("assets/origami_crane.png")
    
    # =========================================================================
    # SLIDE 1: TRANG BÌA — CÀ MAU ĐẤT MŨI PHƯƠNG NAM
    # Hiệu ứng: Mở đầu trang trọng, cánh chim bắt đầu cất cánh từ mép trái
    # =========================================================================
    s1 = deck.add_slide(transition="none")
    
    # Top Badge
    deck.add_badge(s1, "ĐỊA LÝ & DU LỊCH VIỆT NAM // CỰC NAM TỔ QUỐC",
                   left=3.66, top=1.3, width=6.0, height=0.42, name="!!badge_top",
                   bg_color=RGBColor(5, 150, 105), text_color=RGBColor(255, 255, 255))
    
    # Hero Title
    deck.add_hero_title(s1, "CÀ MAU — ĐẤT MŨI PHƯƠNG NAM",
                        subtitle="Nơi đất biết nở, rừng biết đi và biển trời giao hòa",
                        left=2.16, top=2.1, width=9.0, height=2.0,
                        title_size=40, subtitle_size=15,
                        title_color=RGBColor(6, 78, 59), name="!!title_main")
    
    # Center Frosted Acrylic Card
    c1 = deck.add_glass_card(s1, left=3.91, top=4.4, width=5.5, height=1.7, name="!!card_center",
                             bg_hex="FFFFFF", alpha="94000", border_hex="059669", border_width_pt=1.5, corner_radius=6000)
    tx_meta = s1.shapes.add_textbox(Inches(4.1), Inches(4.6), Inches(5.1), Inches(1.3))
    tf1 = tx_meta.text_frame
    tf1.word_wrap = True
    p1 = tf1.paragraphs[0]
    p1.text = "CỘT MỐC TỌA ĐỘ QUỐC GIA GPS 0001"
    p1.font.name = "Segoe UI"
    p1.font.size = Pt(11.5)
    p1.font.bold = True
    p1.font.color.rgb = RGBColor(5, 150, 105)
    p1.alignment = PP_ALIGN.CENTER
    
    p2 = tf1.add_paragraph()
    p2.space_before = Pt(4)
    p2.text = "Khu Dự trữ Sinh quyển Thế giới Mũi Cà Mau & Rừng U Minh Hạ"
    p2.font.name = "Segoe UI"
    p2.font.size = Pt(10)
    p2.font.color.rgb = RGBColor(71, 85, 105)
    p2.alignment = PP_ALIGN.CENTER
    
    # Cánh chim dẫn đường (Chuẩn bị cất cánh từ mép trái, góc nghiêng 22 độ)
    deck.add_traveling_entity(s1, bird_asset, left=0.8, top=3.8, width=2.1, height=2.1, rot_deg=22, name="!!Traveling_Bird")
    
    # Bottom Progress Bar
    deck.add_progress_bar(s1, current_step=1, total_steps=5)

    # =========================================================================
    # SLIDE 2: VỊ TRÍ ĐỊA LÝ ĐỘC NHẤT — MORPH CHUYỂN ĐỘNG 60 FPS (1600ms)
    # Hiệu ứng: Cánh chim bay vút lên bầu trời, card trái và số liệu nở rộng
    # =========================================================================
    s2 = deck.add_slide(transition="morph", duration_ms=1600)
    
    deck.add_badge(s2, "01 // ĐẶC ĐIỂM ĐỊA LÝ ĐỘC BẢN", left=1.0, top=0.9, width=3.4, height=0.38, name="!!badge_top",
                   bg_color=RGBColor(2, 132, 199), text_color=RGBColor(255, 255, 255))
    
    deck.add_hero_title(s2, "MẢNH ĐẤT TẬN CÙNG CỦA NON SÔNG",
                        subtitle="Tỉnh duy nhất trên đất liền Việt Nam giáp cả Biển Đông và Vịnh Thái Lan",
                        left=1.0, top=1.45, width=9.5, height=1.3,
                        title_size=28, subtitle_size=12.5,
                        title_color=RGBColor(6, 78, 59), name="!!title_main")
    
    # Cột chỉ số bên trái
    deck.add_metric_card(s2, "CHIỀU DÀI BỜ BIỂN", "254 KM",
                         "Bao bọc 3 mặt: Đông, Nam và Tây Nam mở ra ngư trường lớn",
                         left=1.0, top=2.8, width=5.2, height=1.7,
                         accent_color=RGBColor(2, 132, 199), name="!!card_left")
                         
    deck.add_metric_card(s2, "RỪNG NGẬP MẶN", "100.000+ HA",
                         "Quy mô rừng ngập mặn sinh thái lớn thứ 2 toàn cầu",
                         left=1.0, top=4.7, width=5.2, height=1.7,
                         accent_color=RGBColor(5, 150, 105), name="!!card_sub")
    
    # Khối nội dung bên phải (Mô tả địa lý)
    c_geo = deck.add_glass_card(s2, left=6.6, top=2.8, width=5.7, height=3.6, name="!!card_center",
                                bg_hex="FFFFFF", alpha="93000", border_hex="059669", border_width_pt=1.5, corner_radius=7000)
    tx_geo = s2.shapes.add_textbox(Inches(7.0), Inches(3.1), Inches(4.9), Inches(3.0))
    tf_g = tx_geo.text_frame
    tf_g.word_wrap = True
    pg1 = tf_g.paragraphs[0]
    pg1.text = "ĐẶC TRƯNG ĐỊA LÝ & SINH THÁI"
    pg1.font.name = "Segoe UI"
    pg1.font.size = Pt(13)
    pg1.font.bold = True
    pg1.font.color.rgb = RGBColor(5, 150, 105)
    
    geo_points = [
        ("Phù sa bồi đắp", "Mỗi năm vùng đất Mũi bồi lấn tự nhiên ra biển hàng trăm mét"),
        ("Hai dòng hải lưu", "Giao thoa giữa dòng hải lưu Biển Đông và Vịnh Thái Lan"),
        ("Sông ngòi chằng chịt", "Hơn 10.000 km đường thủy nội địa hình thành văn hóa sông nước"),
        ("Biểu tượng Mũi Tàu", "Điểm ngắm mặt trời mọc ở Biển Đông và lặn ở Biển Tây duy nhất")
    ]
    for tag, desc in geo_points:
        p_pt = tf_g.add_paragraph()
        p_pt.space_before = Pt(6)
        p_pt.text = f"• {tag}: {desc}."
        p_pt.font.name = "Segoe UI"
        p_pt.font.size = Pt(10.5)
        p_pt.font.color.rgb = RGBColor(51, 65, 85)
        
    # Cánh chim bay vút sang khu vực trung tâm trên cao, hướng ra biển
    deck.add_traveling_entity(s2, bird_asset, left=6.0, top=1.35, width=1.9, height=1.9, rot_deg=-12, name="!!Traveling_Bird")
    
    deck.add_progress_bar(s2, current_step=2, total_steps=5)

    # =========================================================================
    # SLIDE 3: DU KHẢO RỪNG NGẬP MẶN & XỨ SỞ LOÀI CHIM — 60 FPS MORPH
    # Hiệu ứng: Cánh chim tiếp tục liệng cánh qua các tán rừng đước (Morph 1600ms)
    # =========================================================================
    s3 = deck.add_slide(transition="morph", duration_ms=1600)
    
    deck.add_badge(s3, "02 // THIÊN NHIÊN HOANG DÃ", left=1.0, top=0.9, width=3.4, height=0.38, name="!!badge_top",
                   bg_color=RGBColor(5, 150, 105), text_color=RGBColor(255, 255, 255))
    
    deck.add_hero_title(s3, "HỆ SINH THÁI RỪNG & XỨ SỞ LOÀI CHIM",
                        subtitle="Khu Dự trữ Sinh quyển Thế giới với sự kỳ vĩ của thiên nhiên trù phú",
                        left=1.0, top=1.45, width=9.5, height=1.3,
                        title_size=28, subtitle_size=12.5,
                        title_color=RGBColor(6, 78, 59), name="!!title_main")
    
    # 3 Cột Card Du Khảo Thiên Nhiên
    eco_cards = [
        ("VƯỜN QUỐC GIA MŨI CÀ MAU", "Hệ sinh thái rừng ngập mặn", 
         "Quần thể đước nguyên sinh bạt ngàn ôm trọn bãi bồi, nơi cư ngụ của nhiều loài động vật quý hiếm nằm trong Sách Đỏ.", 
         RGBColor(5, 150, 105)),
        ("RỪNG TRÀM U MINH HẠ", "Vương quốc than bùn", 
         "Nổi tiếng với văn hóa gác kèo ong truyền thống, mật ong hoa tràm nguyên chất và cánh rừng tràm xanh thẳm.", 
         RGBColor(2, 132, 199)),
        ("SÂN CHIM CÀ MAU", "Bản hòa ca của trời mây", 
         "Nằm ngay lòng thành phố và các bãi biển, quy tụ hàng vạn cánh cò, diệc, le le về làm tổ vào mỗi buổi hoàng hôn.", 
         RGBColor(217, 119, 6))
    ]
    card_w = 3.6
    gap = 0.26
    for i, (title_c, sub_c, desc_c, col) in enumerate(eco_cards):
        c_x = 1.0 + i * (card_w + gap)
        deck.add_glass_card(s3, left=c_x, top=2.8, width=card_w, height=3.6, name=f"!!eco_card_{i}",
                            border_hex=f"{col[0]:02X}{col[1]:02X}{col[2]:02X}", border_width_pt=1.5, corner_radius=6000)
        tx_ec = s3.shapes.add_textbox(Inches(c_x + 0.2), Inches(3.0), Inches(card_w - 0.4), Inches(3.2))
        tf_ec = tx_ec.text_frame
        tf_ec.word_wrap = True
        
        pec1 = tf_ec.paragraphs[0]
        pec1.text = title_c
        pec1.font.name = "Segoe UI"
        pec1.font.size = Pt(13)
        pec1.font.bold = True
        pec1.font.color.rgb = col
        
        pec2 = tf_ec.add_paragraph()
        pec2.space_before = Pt(3)
        pec2.text = sub_c
        pec2.font.name = "Segoe UI"
        pec2.font.size = Pt(10)
        pec2.font.bold = True
        pec2.font.color.rgb = RGBColor(100, 116, 139)
        
        pec3 = tf_ec.add_paragraph()
        pec3.space_before = Pt(8)
        pec3.text = desc_c
        pec3.font.name = "Segoe UI"
        pec3.font.size = Pt(10.5)
        pec3.font.color.rgb = RGBColor(51, 65, 85)
        
    # Cánh chim lướt nhẹ phía trên sân chim Cà Mau
    deck.add_traveling_entity(s3, bird_asset, left=10.2, top=1.4, width=1.8, height=1.8, rot_deg=8, name="!!Traveling_Bird")
    
    deck.add_progress_bar(s3, current_step=3, total_steps=5)

    # =========================================================================
    # SLIDE 4: SẢN VẬT TRÙ PHÚ & KINH TẾ BIỂN — MORPH DOCKING (1600ms)
    # Hiệu ứng: Chim lượn xuống gần các card đặc sản
    # =========================================================================
    s4 = deck.add_slide(transition="morph", duration_ms=1600)
    
    deck.add_badge(s4, "03 // ẨM THỰC & ĐẶC SẢN NỨC TIẾNG", left=1.0, top=0.9, width=3.4, height=0.38, name="!!badge_top",
                   bg_color=RGBColor(217, 119, 6), text_color=RGBColor(255, 255, 255))
    
    deck.add_hero_title(s4, "SẢN VẬT TRÙ PHÚ VÙNG ĐẤT PHƯƠNG NAM",
                        subtitle="Thương hiệu nông - thủy sản khẳng định vị thế trên thị trường quốc tế",
                        left=1.0, top=1.45, width=9.5, height=1.3,
                        title_size=28, subtitle_size=12.5,
                        title_color=RGBColor(6, 78, 59), name="!!title_main")
    
    # 4 Card Trụ Cột Đặc Sản
    delicacies = [
        ("01", "CUA CÀ MAU", "Vua cua biển Việt Nam", "Vỏ cứng cáp, thịt thơm ngọt chắc nịch, gạch son béo ngậy được bảo hộ chỉ dẫn địa lý.", RGBColor(239, 68, 68)),
        ("02", "TÔM SINH THÁI", "Chuẩn xuất khẩu toàn cầu", "Nuôi hoàn toàn tự nhiên dưới tán rừng đước ngập mặn, không hóa chất, đạt chứng nhận quốc tế.", RGBColor(2, 132, 199)),
        ("03", "MẬT ONG HOA TRÀM", "Đặc sản rừng U Minh", "Mật ong tự nhiên màu vàng óng ánh, thơm hương vị rừng tràm, có giá trị dược liệu cao.", RGBColor(217, 119, 6)),
        ("04", "BA KHÍA RẠCH GỐC", "Đậm đà phong vị dân dã", "Thương hiệu ba khía muối nức tiếng Nam Bộ với hương vị đặc trưng không lẫn vào đâu.", RGBColor(5, 150, 105))
    ]
    card_w4 = 2.65
    gap4 = 0.28
    for i, (num, name_d, tag_d, desc_d, col) in enumerate(delicacies):
        c_x = 1.0 + i * (card_w4 + gap4)
        deck.add_glass_card(s4, left=c_x, top=2.8, width=card_w4, height=3.6, name=f"!!delicacy_{i}",
                            border_hex=f"{col[0]:02X}{col[1]:02X}{col[2]:02X}", border_width_pt=1.5, corner_radius=6000)
        tx_d = s4.shapes.add_textbox(Inches(c_x + 0.18), Inches(3.0), Inches(card_w4 - 0.36), Inches(3.2))
        tf_d = tx_d.text_frame
        tf_d.word_wrap = True
        
        p_num = tf_d.paragraphs[0]
        p_num.text = num
        p_num.font.name = "Segoe UI"
        p_num.font.size = Pt(24)
        p_num.font.bold = True
        p_num.font.color.rgb = col
        
        p_n = tf_d.add_paragraph()
        p_n.space_before = Pt(4)
        p_n.text = name_d
        p_n.font.name = "Segoe UI"
        p_n.font.size = Pt(11)
        p_n.font.bold = True
        p_n.font.color.rgb = RGBColor(6, 78, 59)
        
        p_tag = tf_d.add_paragraph()
        p_tag.space_before = Pt(2)
        p_tag.text = tag_d
        p_tag.font.name = "Segoe UI"
        p_tag.font.size = Pt(9.5)
        p_tag.font.bold = True
        p_tag.font.color.rgb = RGBColor(100, 116, 139)
        
        p_desc = tf_d.add_paragraph()
        p_desc.space_before = Pt(6)
        p_desc.text = desc_d
        p_desc.font.name = "Segoe UI"
        p_desc.font.size = Pt(10)
        p_desc.font.color.rgb = RGBColor(51, 65, 85)
        
    # Cánh chim lượn xuống góc phải để hướng sự chú ý vào các sản vật
    deck.add_traveling_entity(s4, bird_asset, left=10.3, top=4.5, width=1.8, height=1.8, rot_deg=-18, name="!!Traveling_Bird")
    
    deck.add_progress_bar(s4, current_step=4, total_steps=5)

    # =========================================================================
    # SLIDE 5: KẾT LUẬN & ĐIỂM HẸN TƯƠNG LAI — CIRCLE SPOTLIGHT (1600ms)
    # Hiệu ứng: Vòng tròn khẩu độ Iris mở ra tiêu điểm Mũi Tàu vươn khơi
    # =========================================================================
    s5 = deck.add_slide(transition="circle", duration_ms=1600)
    
    deck.add_badge(s5, "ĐIỂM HẸN PHƯƠNG NAM // TẦM NHÌN TƯƠNG LAI", left=4.26, top=1.2, width=4.8, height=0.42, name="!!badge_top",
                   bg_color=RGBColor(5, 150, 105), text_color=RGBColor(255, 255, 255))
    
    # Showcase Card lớn trung tâm
    c_final = deck.add_glass_card(s5, left=2.2, top=1.8, width=8.93, height=4.6, name="!!card_center",
                                  border_hex="059669", border_width_pt=2.0, corner_radius=7000)
    tx_fin = s5.shapes.add_textbox(Inches(2.6), Inches(2.2), Inches(8.13), Inches(3.7))
    tf_f = tx_fin.text_frame
    tf_f.word_wrap = True
    
    pf1 = tf_f.paragraphs[0]
    pf1.alignment = PP_ALIGN.CENTER
    pf1.text = "CÀ MAU KHÔNG CHỈ LÀ MỘT ĐIỂM ĐẾN ĐỊA LÝ,"
    pf1.font.name = "Segoe UI"
    pf1.font.size = Pt(21)
    pf1.font.bold = True
    pf1.font.color.rgb = RGBColor(6, 78, 59)
    
    pf2 = tf_f.add_paragraph()
    pf2.alignment = PP_ALIGN.CENTER
    pf2.text = "MÀ LÀ NIỀM TỰ HÀO THIÊNG LIÊNG CỦA MỖI NGƯỜI VIỆT."
    pf2.font.name = "Segoe UI"
    pf2.font.size = Pt(21)
    pf2.font.bold = True
    pf2.font.color.rgb = RGBColor(2, 132, 199)
    
    pf3 = tf_f.add_paragraph()
    pf3.alignment = PP_ALIGN.CENTER
    pf3.space_before = Pt(14)
    pf3.text = "“Đất Cà Mau trăm sông nghìn lối — Lòng người phương Nam nghĩa tình rộng mở.”"
    pf3.font.name = "Segoe UI"
    pf3.font.size = Pt(13)
    pf3.font.italic = True
    pf3.font.color.rgb = RGBColor(71, 85, 105)
    
    pf4 = tf_f.add_paragraph()
    pf4.alignment = PP_ALIGN.CENTER
    pf4.space_before = Pt(16)
    pf4.text = "CHÀO ĐÓN DU KHÁCH ĐẾN VỚI CỰC NAM TỔ QUỐC"
    pf4.font.name = "Segoe UI"
    pf4.font.size = Pt(12)
    pf4.font.bold = True
    pf4.font.color.rgb = RGBColor(5, 150, 105)
    
    # Cánh chim bay cao đỗ lại ở đỉnh góc phải trong tư thế hiên ngang
    deck.add_traveling_entity(s5, bird_asset, left=10.5, top=1.0, width=2.1, height=2.1, rot_deg=15, name="!!Traveling_Bird")
    
    deck.add_progress_bar(s5, current_step=5, total_steps=5)

    # Lưu file trình chiếu
    output_full = os.path.abspath(output_path)
    deck.save(output_full)
    print(f"[+] Hoàn thành file PowerPoint Cà Mau: {output_full}")
    return output_full

if __name__ == "__main__":
    out = "Ca_Mau_Dat_Mui_Phuong_Nam.pptx"
    pptx_path = build_camau_presentation(out)
    audit_presentation(pptx_path)
