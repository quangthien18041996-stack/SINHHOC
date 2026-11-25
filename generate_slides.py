import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Colors
    C_DARK_BG = RGBColor(19, 44, 32)      # Deep Forest
    C_FOREST = RGBColor(27, 58, 43)       # Brand Dark Green #1b3a2b
    C_EMERALD = RGBColor(16, 185, 129)    # Emerald Green #10b981
    C_MINT = RGBColor(236, 253, 245)      # Mint Light #ecfdf5
    C_AMBER = RGBColor(217, 119, 6)       # Warm Amber #d97706
    C_AMBER_LT = RGBColor(254, 243, 199)  # Light Amber #fef3c7
    C_WHITE = RGBColor(255, 255, 255)
    C_SLATE = RGBColor(30, 41, 59)        # Body text #1e293b
    C_MUTED = RGBColor(100, 116, 139)     # Subtitle #64748b
    C_LIGHT_BG = RGBColor(248, 249, 245)  # Soft Ivory #f8f9f5
    C_CARD_BORDER = RGBColor(226, 232, 220)
    C_ACCENT_BLUE = RGBColor(2, 132, 199)
    C_ACCENT_ROSE = RGBColor(225, 29, 72)

    def set_slide_bg(slide, color):
        background = slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_header(slide, title_text, category_text="KHTN 7 • BÀI 23", light_theme=True):
        # Category pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(4.5), Inches(0.36))
        pill.fill.solid()
        pill.fill.fore_color.rgb = C_MINT if light_theme else RGBColor(27, 58, 43)
        pill.line.color.rgb = C_EMERALD
        pill.line.width = Pt(1)
        tf = pill.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = f"🌿 {category_text}"
        run.font.bold = True
        run.font.size = Pt(10.5)
        run.font.color.rgb = C_FOREST if light_theme else C_EMERALD
        run.font.name = "Arial"

        # Main Title
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.8), Inches(11.7), Inches(0.65))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = title_text
        run.font.bold = True
        run.font.size = Pt(21)
        run.font.color.rgb = C_FOREST if light_theme else C_WHITE
        run.font.name = "Georgia"

        # Accent line
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.733), Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(226, 232, 220) if light_theme else RGBColor(46, 93, 66)
        line.line.fill.background()

    def add_footer(slide, slide_num, total_slides=17, light_theme=True):
        # Footer text
        tx_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.35))
        tf = tx_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        
        # Left part
        run1 = p.add_run()
        run1.text = "Trường THCS Tân Tạo A  •  Giáo viên: Mai Thuý Đức  •  Học liệu số: https://sinhhoc-amber.vercel.app/"
        run1.font.size = Pt(9.5)
        run1.font.color.rgb = C_MUTED if light_theme else RGBColor(167, 243, 208)
        run1.font.name = "Arial"

        # Right part
        p2 = tf.add_paragraph() if False else None # or just add spaces
        run2 = p.add_run()
        run2.text = f"                 Trang {slide_num}/{total_slides}"
        run2.font.size = Pt(9.5)
        run2.font.bold = True
        run2.font.color.rgb = C_FOREST if light_theme else C_WHITE
        run2.font.name = "Arial"

    def create_card(slide, left, top, width, height, bg_color=C_WHITE, border_color=C_CARD_BORDER, border_width=1):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(border_width)
        else:
            card.line.fill.background()
        return card

    # ==========================================================
    # SLIDE 1: TRANG BÌA (TITLE SLIDE)
    # ==========================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, C_DARK_BG)

    # Accent decorative background banner
    card_banner = create_card(s1, 0.8, 0.7, 11.733, 6.1, bg_color=C_FOREST, border_color=RGBColor(46, 93, 66), border_width=1.5)

    # Add Logo
    if os.path.exists("images/logo.png"):
        s1.shapes.add_picture("images/logo.png", Inches(1.3), Inches(1.1), width=Inches(0.9), height=Inches(0.9))

    # Badge Top
    tb_badge = s1.shapes.add_textbox(Inches(2.35), Inches(1.15), Inches(9.8), Inches(0.4))
    tf = tb_badge.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH • HỘI THI DẠY HỌC SỐ & STEM NĂM HỌC 2026 - 2027"
    r.font.size = Pt(10)
    r.font.bold = True
    r.font.color.rgb = RGBColor(167, 243, 208)
    r.font.name = "Arial"

    # Main Titles
    tb_title = s1.shapes.add_textbox(Inches(1.3), Inches(2.2), Inches(10.7), Inches(2.0))
    tf = tb_title.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "BÀI 23: MỘT SỐ YẾU TỐ ẢNH HƯỞNG\nĐẾN QUANG HỢP"
    r.font.bold = True
    r.font.size = Pt(32)
    r.font.color.rgb = C_WHITE
    r.font.name = "Georgia"

    p2 = tf.add_paragraph()
    p2.space_before = Pt(12)
    r2 = p2.add_run()
    r2.text = "CHỦ ĐỀ DẠY HỌC TÍCH HỢP STEM & CHUYỂN ĐỔI SỐ (90 PHÚT) • BỘ KẾT NỐI TRI THỨC VỚI CUỘC SỐNG"
    r2.font.bold = True
    r2.font.size = Pt(13)
    r2.font.color.rgb = RGBColor(217, 119, 6) # Amber
    r2.font.name = "Arial"

    # Teacher Profile Card at Bottom
    card_info = create_card(s1, 1.3, 4.6, 10.7, Inches(1.8), bg_color=RGBColor(19, 44, 32), border_color=RGBColor(52, 110, 78), border_width=1)
    
    tb_info = s1.shapes.add_textbox(Inches(1.5), Inches(4.7), Inches(10.3), Inches(1.6))
    tf = tb_info.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "THÔNG TIN GIÁO VIÊN & ĐƠN VỊ DỰ THI:"
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(167, 243, 208)
    r.font.name = "Arial"

    p_gv = tf.add_paragraph()
    p_gv.space_before = Pt(6)
    r_gv = p_gv.add_run()
    r_gv.text = "• Giáo viên thực hiện:  "
    r_gv.font.size = Pt(13)
    r_gv.font.color.rgb = C_WHITE
    r_gv = p_gv.add_run()
    r_gv.text = "MAI THUÝ ĐỨC"
    r_gv.font.bold = True
    r_gv.font.size = Pt(14)
    r_gv.font.color.rgb = RGBColor(254, 243, 199)

    r_sch = p_gv.add_run()
    r_sch.text = "          • Đơn vị công tác:  "
    r_sch.font.size = Pt(13)
    r_sch.font.color.rgb = C_WHITE
    r_sch = p_gv.add_run()
    r_sch.text = "TRƯỜNG THCS TÂN TẠO A (BÌNH TÂN)"
    r_sch.font.bold = True
    r_sch.font.size = Pt(13)
    r_sch.font.color.rgb = RGBColor(254, 243, 199)

    p_web = tf.add_paragraph()
    p_web.space_before = Pt(4)
    r_web = p_web.add_run()
    r_web.text = "• Nền tảng học liệu số tương tác chính thức: "
    r_web.font.size = Pt(11)
    r_web.font.color.rgb = C_WHITE
    r_web = p_web.add_run()
    r_web.text = "https://sinhhoc-amber.vercel.app/"
    r_web.font.bold = True
    r_web.font.size = Pt(11)
    r_web.font.color.rgb = C_EMERALD

    # ==========================================================
    # SLIDE 2: YÊU CẦU CẦN ĐẠT (CV 5512)
    # ==========================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s2, C_LIGHT_BG)
    add_header(s2, "Mục Tiêu & Yêu Cầu Cần Đạt Của Bài Dạy (Theo CV 5512)", "MỤC TIÊU BÀI DẠY")
    add_footer(s2, 2)

    # 3 Cards for Objectives
    # Card 1: Năng lực KHTN
    c1 = create_card(s2, 0.8, 1.7, 3.7, 5.1, bg_color=C_WHITE, border_color=C_EMERALD, border_width=1.5)
    tb = s2.shapes.add_textbox(Inches(0.95), Inches(1.85), Inches(3.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌱 NĂNG LỰC KHTN"
    r.font.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = C_FOREST

    items1 = [
        ("Nhận thức KHTN:", "Nêu và phân tích được ảnh hưởng định lượng của 4 yếu tố môi trường (Ánh sáng, Nước, CO₂, Nhiệt độ) đến tốc độ quang hợp."),
        ("Tìm hiểu tự nhiên:", "Thu thập, quan sát và phân tích số liệu thí nghiệm số (giải phóng bọt khí O₂ cành rong đuôi chó; đồ thị động học quang hợp)."),
        ("Vận dụng kiến thức:", "Đóng vai Kỹ sư sinh thái, đề xuất giải pháp bố trí cây xanh thích ứng vi khí hậu 4 vùng đặc thù TP.HCM.")
    ]
    for title, desc in items1:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_FOREST
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_SLATE

    # Card 2: Năng lực Số & Chung
    c2 = create_card(s2, 4.8, 1.7, 3.7, 5.1, bg_color=C_WHITE, border_color=C_ACCENT_BLUE, border_width=1.5)
    tb = s2.shapes.add_textbox(Inches(4.95), Inches(1.85), Inches(3.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "💻 NĂNG LỰC SỐ & CHUNG"
    r.font.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = C_ACCENT_BLUE

    items2 = [
        ("Năng lực số:", "Thao tác thuần thục phòng thí nghiệm ảo, phân tích dữ liệu cảm biến IoT vi khí hậu thời gian thực, nộp bài và trình chiếu E-Poster trực tuyến."),
        ("Tự chủ & tự học:", "Chủ động điều tra các biến số, tự rút ra quy luật sinh học qua trải nghiệm mô phỏng tương tác."),
        ("Giao tiếp & hợp tác:", "Phân công nhịp nhàng trong nhóm 4 học sinh, phản biện khoa học và thống nhất phương án sinh thái tối ưu.")
    ]
    for title, desc in items2:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_ACCENT_BLUE
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_SLATE

    # Card 3: Phẩm chất
    c3 = create_card(s2, 8.8, 1.7, 3.7, 5.1, bg_color=C_WHITE, border_color=C_AMBER, border_width=1.5)
    tb = s2.shapes.add_textbox(Inches(8.95), Inches(1.85), Inches(3.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "⭐ PHẨM CHẤT CHỦ YẾU"
    r.font.bold = True
    r.font.size = Pt(14)
    r.font.color.rgb = C_AMBER

    items3 = [
        ("Chăm chỉ:", "Tích cực tìm tòi tài liệu, kiên trì thử nghiệm các biến số trên mô hình số hóa để khám phá chân lý khoa học."),
        ("Trung thực:", "Ghi nhận khách quan các số liệu thực nghiệm số, không ngụy tạo kết quả thí nghiệm."),
        ("Trách nhiệm:", "Ý thức sâu sắc về vai trò của cây xanh trong việc hấp thụ CO₂, lọc bụi mịn PM2.5, giảm hiệu ứng nhà kính và chủ động lan tỏa lối sống xanh đô thị.")
    ]
    for title, desc in items3:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_AMBER
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 3: TIẾN TRÌNH DẠY HỌC 90 PHÚT
    # ==========================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s3, C_LIGHT_BG)
    add_header(s3, "Tiến Trình Dạy Học 90 Phút (2 Tiết Học Tích Hợp STEM)", "KẾ HOẠCH BÀI DẠY (CV 5512)")
    add_footer(s3, 3)

    # Left Column: Tiết 1
    c_t1 = create_card(s3, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_FOREST, border_width=1.5)
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "⏱️ TIẾT 1 (45 PHÚT) — KHÁM PHÁ QUY LUẬT"
    r.font.bold = True
    r.font.size = Pt(13.5)
    r.font.color.rgb = C_FOREST

    t1_steps = [
        ("Hoạt động 1: Khởi động (5 phút)", "Nêu tình huống nghịch lý sinh học: Cây cháy lá ngoài trời nắng gắt vs cây héo úa trong nhà tối. Học sinh thảo luận cặp đôi xác định vấn đề cần nghiên cứu."),
        ("Hoạt động 2: Khám phá kiến thức (35 phút)", "Học sinh sử dụng thiết bị số truy cập Module 2 (Phòng Thí Nghiệm Số), luân phiên điều tra 4 yếu tố:\n• Yếu tố Ánh sáng: Bão hòa ánh sáng, bẫy cháy lá.\n• Yếu tố Nước: Quang phân ly, cơ chế mở lỗ khí khổng.\n• Yếu tố Khí CO₂: Nồng độ tối ưu, ngưỡng ngộ độc CO₂.\n• Yếu tố Nhiệt độ: Hoạt tính & biến tính enzyme quang hợp."),
        ("Đúc kết & Chốt Tiết 1 (5 phút)", "Các nhóm hoàn thành Phiếu học tập số 1; chuẩn hóa kiến thức trọng tâm về 4 yếu tố ngoại cảnh.")
    ]
    for st_title, st_desc in t1_steps:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"📍 {st_title}\n"
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_FOREST
        r2 = p.add_run()
        r2.text = st_desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_SLATE

    # Right Column: Tiết 2
    c_t2 = create_card(s3, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_AMBER, border_width=1.5)
    tb = s3.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🚀 TIẾT 2 (45 PHÚT) — VẬN DỤNG STEM & XUẤT BẢN"
    r.font.bold = True
    r.font.size = Pt(13.5)
    r.font.color.rgb = C_AMBER

    t2_steps = [
        ("Hoạt động 3: Luyện tập nhanh (10 phút)", "Học sinh tham gia mini quiz trắc nghiệm tương tác trên nền tảng số để củng cố các mốc giới hạn sinh thái quang hợp."),
        ("Hoạt động 4: Vận dụng STEM (25 phút)", "Đóng vai Kỹ sư sinh thái kiến tạo 4 Trạm Không Gian Xanh:\n• Trạm 1: Tháp rừng thẳng đứng (Bosco Verticale).\n• Trạm 2: Đảo nhiệt trung tâm thương mại & phố đi bộ.\n• Trạm 3: Nông nghiệp công nghệ cao nhà màng Củ Chi.\n• Trạm 4: Không gian xanh Trường THCS Tân Tạo A.\nKéo thả cây giống, điều tiết vi khí hậu, theo dõi IoT."),
        ("Hoạt động 5: Báo cáo & Đánh giá (10 phút)", "Xuất bản E-Poster khoa học, Cô Mai Thuý Đức thẩm định chuẩn y, nhận chứng nhận số và đánh giá theo Rubric.")
    ]
    for st_title, st_desc in t2_steps:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"📍 {st_title}\n"
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_AMBER
        r2 = p.add_run()
        r2.text = st_desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 4: KHỞI ĐỘNG - NGHỊCH LÝ CÂY XANH ĐÔ THỊ
    # ==========================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s4, C_LIGHT_BG)
    add_header(s4, "Khởi Động: Tình Huống Nghịch Lý Sinh Học Đô Thị", "TIẾT 1 • HOẠT ĐỘNG 1 (5 PHÚT)")
    add_footer(s4, 4)

    # 2 Comparison Cards
    c_left = create_card(s4, 0.8, 1.7, 5.7, 3.8, bg_color=C_WHITE, border_color=C_ACCENT_ROSE, border_width=1.5)
    tb = s4.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "☀️ TÌNH HUỐNG 1: NGOÀI TRỜI NẮNG GẮT"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_ACCENT_ROSE

    p = tf.add_paragraph()
    p.space_before = Pt(8)
    r = p.add_run()
    r.text = "• Hiện tượng: "
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = C_SLATE
    r2 = p.add_run()
    r2.text = "Cây cảnh đặt ngoài sân thượng giữa trưa hè (bức xạ >45.000 Lux, nhiệt độ >38°C) dù sáng nào cũng được tưới nước đẫm nhưng mép lá vẫn bị cháy xém, cây héo rũ."
    r2.font.size = Pt(10.5)

    p2 = tf.add_paragraph()
    p2.space_before = Pt(8)
    r = p2.add_run()
    r.text = "• Nghịch lý: "
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = C_ACCENT_ROSE
    r2 = p2.add_run()
    r2.text = "Quang hợp cần ánh sáng để tổng hợp chất hữu cơ, tại sao khi thừa thãi ánh sáng cây lại bị cháy khô và ngừng phát triển?"
    r2.font.size = Pt(10.5)

    c_right = create_card(s4, 6.8, 1.7, 5.7, 3.8, bg_color=C_WHITE, border_color=C_ACCENT_BLUE, border_width=1.5)
    tb = s4.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌑 TÌNH HUỐNG 2: TRONG NHÀ TỐI THIẾU SÁNG"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_ACCENT_BLUE

    p = tf.add_paragraph()
    p.space_before = Pt(8)
    r = p.add_run()
    r.text = "• Hiện tượng: "
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = C_SLATE
    r2 = p.add_run()
    r2.text = "Cây thân gỗ lớn được bứng trồng trang trí ở góc hành lang tối kín gió trong nhà, dù được tưới nước đều đặn nhưng lá cứ vàng vọt dần rồi rụng trơ cành."
    r2.font.size = Pt(10.5)

    p2 = tf.add_paragraph()
    p2.space_before = Pt(8)
    r = p2.add_run()
    r.text = "• Nghịch lý: "
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = C_ACCENT_BLUE
    r2 = p2.add_run()
    r2.text = "Có nước và không khí đầy đủ, vì sao cây vẫn không thể quang hợp để tự nuôi sống mình?"
    r2.font.size = Pt(10.5)

    # Bottom Question Plate
    c_bot = create_card(s4, 0.8, 5.7, 11.7, 1.2, bg_color=C_MINT, border_color=C_EMERALD, border_width=1.5)
    tb = s4.shapes.add_textbox(Inches(1.0), Inches(5.8), Inches(11.3), Inches(1.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "❓ CÂU HỎI TRỌNG TÂM CỦA BÀI HỌC:"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = C_FOREST
    p2 = tf.add_paragraph()
    p2.space_before = Pt(3)
    r = p2.add_run()
    r.text = "Quang hợp của thực vật chịu ảnh hưởng bởi những yếu tố môi trường nào? Ngưỡng giới hạn thích hợp của từng yếu tố là bao nhiêu để cây sinh trưởng tối ưu nhất?"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 5: TƯ LIỆU THỰC NGHIỆM SGK TRANG 107 (BOSCO VERTICALE)
    # ==========================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s5, C_LIGHT_BG)
    add_header(s5, "Tư Liệu Thực Nghiệm SGK Trang 107: 'Khu Rừng Thẳng Đứng' Lọc Không Khí", "TƯ LIỆU SGK KHTN 7")
    add_footer(s5, 5)

    # Left: Text Excerpt
    c_txt = create_card(s5, 0.8, 1.7, 6.2, 5.1, bg_color=C_WHITE, border_color=C_FOREST, border_width=1.5)
    tb = s5.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.8), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🏢 TÒA THÁP RỪNG THẲNG ĐỨNG (BOSCO VERTICALE)"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_FOREST

    p = tf.add_paragraph()
    p.space_before = Pt(6)
    r = p.add_run()
    r.text = "Trích SGK KHTN 7 (Bộ Kết nối tri thức - Trang 107):"
    r.font.italic = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = C_MUTED

    sgk_text = (
        "Kiến trúc sư người Italia là Stê-pha-nô Bo-e-ri (Stefano Boeri) đã tạo ra toà tháp rừng thẳng đứng "
        "đầu tiên mang tên Bosco Verticale ở thành phố Milan, Italia.\n\n"
        "Toà tháp đôi này được phủ bằng một rừng cây gồm khoảng 900 cây xanh, 5.000 bụi cây với trên 11.000 loài thực vật khác nhau. "
        "Khu rừng này có thể hấp thụ CO₂, ngăn ngừa các khí gây hiệu ứng nhà kính phát tán lên bầu khí quyển.\n\n"
        "Cây xanh trên nóc toà nhà có thể lọc bụi mịn PM2.5. Kiến trúc sư Boeri đã chọn những loại cây có khả năng hấp thụ khí gây ô nhiễm "
        "và bụi mịn PM2.5 tốt nhất."
    )
    p2 = tf.add_paragraph()
    p2.space_before = Pt(6)
    r2 = p2.add_run()
    r2.text = sgk_text
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = C_SLATE

    # Right: Image & Challenge
    if os.path.exists("screenshots/bosco_updated.png"):
        s5.shapes.add_picture("screenshots/bosco_updated.png", Inches(7.2), Inches(1.7), width=Inches(5.33), height=Inches(3.3))
    
    c_chal = create_card(s5, 7.2, 5.15, 5.33, 1.65, bg_color=C_AMBER_LT, border_color=C_AMBER, border_width=1)
    tb = s5.shapes.add_textbox(Inches(7.35), Inches(5.25), Inches(5.0), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🎯 BÀI TOÁN KHOA HỌC DÀNH CHO HỌC SINH:"
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = C_AMBER

    p2 = tf.add_paragraph()
    p2.space_before = Pt(4)
    r2 = p2.add_run()
    r2.text = "Ở độ cao tầng thượng (nắng gắt, gió to) và chân tòa nhà (bóng râm, thiếu sáng), ta phải chọn và bố trí các loại cây như thế nào để cây đạt hiệu suất quang hợp tối đa?"
    r2.font.size = Pt(10)
    r2.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 6: YẾU TỐ 1 - ÁNH SÁNG
    # ==========================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s6, C_LIGHT_BG)
    add_header(s6, "Khám Phá Yếu Tố 1: Ánh Sáng Đối Với Quang Hợp", "TIẾT 1 • HOẠT ĐỘNG 2 (PHÒNG THÍ NGHIỆM SỐ)")
    add_footer(s6, 6)

    # Left: Experiment & Science
    c_left = create_card(s6, 0.8, 1.7, 6.2, 5.1, bg_color=C_WHITE, border_color=C_EMERALD, border_width=1.5)
    tb = s6.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.8), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "☀️ THÍ NGHIỆM CÀNH RONG ĐUÔI CHÓ (ELODEA)"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_FOREST

    light_points = [
        ("Cơ chế giải phóng bọt khí O₂:", "Khi chiếu sáng, diệp lục hấp thụ photon kích hoạt quang phân ly nước, giải phóng bọt khí O₂ từ vết cắt cành rong. Số bọt khí/phút phản ánh trực tiếp tốc độ quang hợp."),
        ("Điểm bù ánh sáng:", "Cường độ ánh sáng mà tại đó tốc độ quang hợp vừa vặn bằng tốc độ hô hấp (lượng CO₂ hấp thụ = lượng CO₂ giải phóng)."),
        ("Điểm bão hòa ánh sáng (~25.000 Lux):", "Khi tăng cường độ ánh sáng, tốc độ quang hợp tăng dần đến cực đại; vượt qua điểm này, quang hợp không tăng thêm."),
        ("Bẫy cháy lá (>45.000 Lux):", "Ánh sáng quá mạnh đốt nóng phiến lá, phân hủy diệp lục, khí khổng đóng để hạn chế mất nước làm quang hợp tụt giảm dốc đứng."),
        ("Phân nhóm thực vật:", "Cây ưa sáng (phiến lá dày, tầng cutin bóng, cần ánh sáng mạnh) vs Cây ưa bóng (phiến lá mỏng, chứa nhiều diệp lục b, quang hợp tốt ở nơi râm mát).")
    ]
    for title, desc in light_points:
        p = tf.add_paragraph()
        p.space_before = Pt(5)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_FOREST
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_SLATE

    # Right: Screenshot Lab
    if os.path.exists("screenshots/lab_elodea_realistic.png"):
        s6.shapes.add_picture("screenshots/lab_elodea_realistic.png", Inches(7.2), Inches(1.7), width=Inches(5.33), height=Inches(5.1))

    # ==========================================================
    # SLIDE 7: YẾU TỐ 2 - NƯỚC
    # ==========================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s7, C_LIGHT_BG)
    add_header(s7, "Khám Phá Yếu Tố 2: Nước Đối Với Quang Hợp", "TIẾT 1 • HOẠT ĐỘNG 2 (PHÒNG THÍ NGHIỆM SỐ)")
    add_footer(s7, 7)

    # 2 Big Cards
    c1 = create_card(s7, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_ACCENT_BLUE, border_width=1.5)
    tb = s7.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "💧 1. NGUYÊN LIỆU TRỰC TIẾP CỦA QUANG HỢP"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_ACCENT_BLUE

    points_w1 = [
        ("Nguyên liệu phản ứng:", "Nước chiếm 70% – 90% khối lượng tươi của lá, là nguyên liệu trực tiếp tham gia phản ứng quang phân ly nước trong pha sáng của quang hợp:"),
        ("Phương trình quang phân ly:", "2H₂O  ➔  4H⁺  +  4e⁻  +  O₂ ↑\n(Giải phóng dưỡng khí Oxygen cho sinh quyển)."),
        ("Môi trường hòa tan & vận chuyển:", "Nước là dung môi hòa tan muối khoáng, vận chuyển sản phẩm quang hợp (chất hữu cơ) từ lá đến các cơ quan dự trữ (thân, rễ, củ, quả)."),
        ("Điều hòa nhiệt độ phiến lá:", "Sự thoát hơi nước qua bề mặt lá giúp hạ nhiệt độ phiến lá trong những ngày nắng nóng, bảo vệ cấu trúc màng thylakoid không bị phá hủy.")
    ]
    for title, desc in points_w1:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_ACCENT_BLUE
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_SLATE

    c2 = create_card(s7, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_FOREST, border_width=1.5)
    tb = s7.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🔬 2. ĐIỀU TIẾT CƠ CHẾ ĐÓNG MỞ KHÍ KHỔNG"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_FOREST

    points_w2 = [
        ("Khi cây đủ nước (Trương nước):", "Tế bào hạt đậu no nước căng tròn, thành mỏng phía ngoài dãn ra làm thành dày phía trong uốn cong ➔ Lỗ khí khổng mở rộng tối đa ➔ Khí CO₂ ồ ạt khuếch tán vào mô xốp lá để quang hợp."),
        ("Khi cây thiếu nước (Mất nước):", "Tế bào hạt đậu mất sức trương trở nên xẹp, thành mỏng co lại ➔ Lỗ khí khổng đóng chặt ➔ Khí CO₂ không thể vào lá ➔ Tốc độ quang hợp giảm về 0."),
        ("Nguy cơ ngập úng rễ:", "Tưới thừa nước ngập úng đáy chậu làm đất nghèo O₂, rễ bị ngạt không thể hô hấp để hấp thụ khoáng ➔ Cây bị 'héo sinh lý' dù đất ướt sũng."),
        ("Ứng dụng thực tiễn:", "Bố trí công nghệ tưới nhỏ giọt hẹn giờ tự động, giữ ẩm tối ưu 60% – 75%.")
    ]
    for title, desc in points_w2:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_FOREST
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 8: YẾU TỐ 3 & 4 - KHÍ CO2 VÀ NHIỆT ĐỘ
    # ==========================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s8, C_LIGHT_BG)
    add_header(s8, "Khám Phá Yếu Tố 3 & 4: Nồng Độ Khí CO₂ và Nhiệt Độ", "TIẾT 1 • HOẠT ĐỘNG 2 (PHÒNG THÍ NGHIỆM SỐ)")
    add_footer(s8, 8)

    # Card Left: CO2
    c1 = create_card(s8, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_AMBER, border_width=1.5)
    tb = s8.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "☁️ YẾU TỐ 3: HÀM LƯỢNG KHÍ CO₂"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_AMBER

    co2_points = [
        ("Vai trò sinh học:", "CO₂ là nguyên liệu trực tiếp được đồng hóa trong pha tối (chu trình Calvin) để tổng hợp nên glucose và tinh bột."),
        ("Nồng độ bình thường trong khí quyển:", "~0,04% (tương đương 400 ppm)."),
        ("Khoảng nồng độ tối ưu:", "Từ 0,03% đến 0,15% (300 ppm – 1.500 ppm). Khi tăng nồng độ CO₂ trong khoảng này, cường độ quang hợp tăng tỷ lệ thuận."),
        ("Điểm bão hòa CO₂ (~0,15%):", "Tốc độ quang hợp đạt cực đại; nếu tiếp tục tăng CO₂, quang hợp không tăng thêm."),
        ("Ngưỡng ngộ độc CO₂ (>0,20%):", "Nồng độ CO₂ quá cao làm tê liệt tế bào khí khổng, ức chế enzyme và gây ngộ độc tế bào làm cây suy kiệt.")
    ]
    for title, desc in co2_points:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_AMBER
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_SLATE

    # Card Right: Nhiệt độ
    c2 = create_card(s8, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_ACCENT_ROSE, border_width=1.5)
    tb = s8.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌡️ YẾU TỐ 4: NHIỆT ĐỘ MÔI TRƯỜNG"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_ACCENT_ROSE

    temp_points = [
        ("Bản chất sinh học:", "Các phản ứng quang hợp đều được xúc tác bởi hệ enzyme đặc hiệu (đặc biệt là enzyme RuBisCO). Hoạt tính của enzyme phụ thuộc chặt chẽ vào nhiệt độ."),
        ("Khoảng nhiệt độ tối ưu:", "Từ 25°C đến 35°C ở đa số thực vật nhiệt đới. Tại dải nhiệt độ này, tốc độ xúc tác enzyme đạt đỉnh cao nhất."),
        ("Dưới 10°C (Nhiệt độ thấp):", "Enzyme quang hợp bị ức chế hoạt tính; màng sinh chất giảm tính linh động, quang hợp diễn ra rất chậm hoặc đình trệ."),
        ("Trên 40°C (Nhiệt độ cao):", "Protein enzyme bị biến tính bất thuận nghịch, mất hoạt tính xúc tác; khí khổng đóng để hạn chế mất nước làm quang hợp ngưng hoàn toàn.")
    ]
    for title, desc in temp_points:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_ACCENT_ROSE
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 9: CHỐT KIẾN THỨC TIẾT 1 (PHIẾU HỌC TẬP SỐ 1)
    # ==========================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s9, C_LIGHT_BG)
    add_header(s9, "Đúc Kết Tiết 1: Bảng Chuẩn Hóa Kiến Thức (Phiếu Học Tập Số 1)", "TIẾT 1 • TỔNG KẾT & ĐÁNH GIÁ (5 PHÚT)")
    add_footer(s9, 9)

    # Big Summary Table
    rows = [
        ("Yếu tố ngoại cảnh", "Khoảng tối ưu cho quang hợp", "Hiện tượng khi vượt ngưỡng / cực đoan", "Giải pháp kỹ thuật trồng trọt"),
        ("1. Ánh sáng", "15.000 – 30.000 Lux\n(Điểm bão hòa ~25.000 Lux)", "> 45.000 Lux: Bão hòa quang, cháy diệp lục lá\n< 3.000 Lux: Dưới điểm bù, không quang hợp", "Trồng xen canh phân tầng tán (cây ưa sáng tầng trên, cây ưa bóng tầng dưới)"),
        ("2. Nước cấp", "Độ ẩm đất 60% – 75%\n(Tế bào hạt đậu no nước)", "Hạn hán: Tế bào xẹp, đóng khí khổng\nNgập úng: Ngạt rễ, héo sinh lý", "Lắp đặt hệ thống tưới nhỏ giọt hẹn giờ, đất tơi xốp thoát nước tốt"),
        ("3. Khí CO₂", "0,03% – 0,15%\n(300 – 1.500 ppm)", "< 0,008%: Ngừng quang hợp\n> 0,20%: Khí khổng tê liệt, ngộ độc tế bào", "Bón phân hữu cơ sinh học giải phóng CO₂, tăng cường thông gió nhà màng"),
        ("4. Nhiệt độ", "25°C – 35°C\n(Enzyme RuBisCO tối ưu)", "< 10°C: Enzyme bị ức chế hoạt tính\n> 40°C: Enzyme biến tính, ngừng quang hợp", "Lắp lưới lan giảm bức xạ nhiệt, phun sương làm mát vi khí hậu")
    ]

    table_shape = s9.shapes.add_table(5, 4, Inches(0.8), Inches(1.7), Inches(11.733), Inches(5.1))
    table = table_shape.table
    table.columns[0].width = Inches(2.3)
    table.columns[1].width = Inches(2.9)
    table.columns[2].width = Inches(3.4)
    table.columns[3].width = Inches(3.133)

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_FOREST
                p.font.bold = True
                p.font.color.rgb = C_WHITE
                p.font.size = Pt(11.5)
                p.alignment = PP_ALIGN.CENTER
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_WHITE if r_idx % 2 == 1 else C_MINT
                p.font.size = Pt(10)
                p.font.color.rgb = C_SLATE
                if c_idx == 0:
                    p.font.bold = True
                    p.font.color.rgb = C_FOREST

    # ==========================================================
    # SLIDE 10: TIẾT 2 - HOẠT ĐỘNG LUYỆN TẬP (MINI QUIZ)
    # ==========================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s10, C_LIGHT_BG)
    add_header(s10, "Luyện Tập: Thử Thách Trắc Nghiệm Tương Tác Số", "TIẾT 2 • HOẠT ĐỘNG 3 (10 PHÚT)")
    add_footer(s10, 10)

    # 4 Quiz Question Cards
    quiz_items = [
        ("Câu 1: Yếu tố nào vừa là nguyên liệu trực tiếp giải phóng O₂, vừa điều khiển cơ chế đóng mở khí khổng?",
         "A. Khí Carbon dioxide (CO₂)\nB. Nước (H₂O)  ➔  [ĐÁP ÁN ĐÚNG]\nC. Nhiệt độ môi trường\nD. Bức xạ ánh sáng"),
        ("Câu 2: Khi cường độ ánh sáng vượt ngưỡng 45.000 Lux vào giữa trưa nắng gắt, điều gì sẽ xảy ra?",
         "A. Quang hợp tăng vọt không ngừng\nB. Cây hấp thụ thêm nhiều khí CO₂\nC. Diệp lục bị phá hủy, cháy mô lá, quang hợp giảm mạnh  ➔  [ĐÁP ÁN ĐÚNG]\nD. Lỗ khí khổng mở rộng tối đa"),
        ("Câu 3: Nồng độ khí CO₂ trong không khí ở ngưỡng nào sẽ làm cây ngộ độc sinh học?",
         "A. Khoảng 0,04%\nB. Khoảng 0,10%\nC. Vượt quá 0,20%  ➔  [ĐÁP ÁN ĐÚNG]\nD. Dưới 0,01%"),
        ("Câu 4: Nhiệt độ môi trường vượt quá 40°C làm quang hợp ngừng trệ là do nguyên nhân nào?",
         "A. Nước bị đóng băng\nB. Hệ enzyme quang hợp bị biến tính mất hoạt tính  ➔  [ĐÁP ÁN ĐÚNG]\nC. Nồng độ O₂ trong không khí cạn kiệt\nD. Diệp lục biến thành màu xanh đậm hơn")
    ]

    for idx, (q, opts) in enumerate(quiz_items):
        col = idx % 2
        row = idx // 2
        left = 0.8 + col * 6.0
        top = 1.7 + row * 2.55
        c = create_card(s10, left, top, 5.7, 2.35, bg_color=C_WHITE, border_color=C_EMERALD, border_width=1.2)
        tb = s10.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.1), Inches(5.4), Inches(2.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = q
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_FOREST

        p2 = tf.add_paragraph()
        p2.space_before = Pt(4)
        r2 = p2.add_run()
        r2.text = opts
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 11: VẬN DỤNG STEM - DỰ ÁN KỸ SƯ SINH THÁI
    # ==========================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s11, C_LIGHT_BG)
    add_header(s11, "Vận Dụng STEM: Dự Án Kỹ Sư Sinh Thái 4 Trạm Vi Khí Hậu", "TIẾT 2 • HOẠT ĐỘNG 4 (25 PHÚT)")
    add_footer(s11, 11)

    # 4 Station Cards for 4 Groups
    stations = [
        ("TRẠM 1: ĐÔ THỊ NÉN CAO TẦNG", "Nhóm 1 phụ trách", "Khu tháp căn hộ bê tông hóa: Nắng gắt tầng trên (>45.000 Lux), ngạt khí thải khói xe (CO₂ 850 ppm), tầng trệt thiếu sáng (3.500 Lux).", "Thiết kế tháp rừng thẳng đứng Bosco Verticale đa tầng tán.", C_FOREST),
        ("TRẠM 2: TRUNG TÂM THƯƠNG MẠI", "Nhóm 2 phụ trách", "Đảo nhiệt đô thị bê tông & kính: Nhiệt độ bức xạ >38°C, gió quẩn khô nóng, bão hòa nhiệt nhân tạo.", "Thiết kế ốc đảo sinh thái phố đi bộ, vách tường rêu và phun sương.", C_ACCENT_BLUE),
        ("TRẠM 3: NÔNG NGHIỆP CÔNG NGHỆ CAO", "Nhóm 3 phụ trách", "Khu nhà kính / nhà màng Củ Chi: Hầm nhiệt bí gió, CO₂ cạn kiệt cục bộ vào giữa trưa, độ ẩm biến động mạnh.", "Điều khiển mái che tự động, cấp CO₂ sinh học, tưới nhỏ giọt hẹn giờ.", C_AMBER),
        ("TRẠM 4: KHÔNG GIAN HỌC ĐƯỜNG", "Nhóm 4 phụ trách", "Trường THCS Tân Tạo A: Bụi mịn PM2.5 đường lộ, hành lang hướng Tây đón nắng gay gắt mùa khô TP.HCM.", "Thiết kế hành lang xanh lọc bụi, giàn cây leo cản nhiệt hướng Tây.", C_ACCENT_ROSE)
    ]

    for idx, (st_name, grp, problem, goal, col_border) in enumerate(stations):
        c_left = 0.8 + (idx % 2) * 6.0
        c_top = 1.7 + (idx // 2) * 2.55
        c = create_card(s11, c_left, c_top, 5.7, 2.35, bg_color=C_WHITE, border_color=col_border, border_width=1.5)
        tb = s11.shapes.add_textbox(Inches(c_left + 0.15), Inches(c_top + 0.1), Inches(5.4), Inches(2.15))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = f"📍 {st_name} "
        r.font.bold = True
        r.font.size = Pt(11.5)
        r.font.color.rgb = col_border

        r_g = p.add_run()
        r_g.text = f"({grp})"
        r_g.font.bold = True
        r_g.font.size = Pt(10)
        r_g.font.color.rgb = C_MUTED

        p_prob = tf.add_paragraph()
        p_prob.space_before = Pt(4)
        r = p_prob.add_run()
        r.text = "• Thách thức: "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = C_SLATE
        r2 = p_prob.add_run()
        r2.text = problem
        r2.font.size = Pt(9.5)

        p_goal = tf.add_paragraph()
        p_goal.space_before = Pt(3)
        r = p_goal.add_run()
        r.text = "• Mục tiêu kỹ thuật: "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = col_border
        r2 = p_goal.add_run()
        r2.text = goal
        r2.font.size = Pt(9.5)

    # ==========================================================
    # SLIDE 12: THAO TÁC TRẠM 1 - KHU VƯỜN THẲNG ĐỨNG
    # ==========================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s12, C_LIGHT_BG)
    add_header(s12, "Trải Nghiệm Trạm 1: Kiến Tạo Rừng Thẳng Đứng Bosco Verticale", "THAO TÁC THỰC NGHIỆM SỐ TRÊN WEB")
    add_footer(s12, 12)

    # Left: Steps
    c_left = create_card(s12, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_FOREST, border_width=1.5)
    tb = s12.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌿 BỐ TRÍ CÂY GIỐNG 4 TẦNG SINH THÁI"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_FOREST

    st1_steps = [
        ("Tầng Thượng: Cây Ưa Sáng Cản Bụi", "Trồng Cây Bàng Đài Loan, Cây Trúc Cần Câu. Lá dày, tầng cutin bóng chịu nắng gắt 45.000 Lux, đón gió cản bụi mịn PM2.5."),
        ("Vách Tường Đứng: Giàn Cây Leo Hút Nhiệt", "Trồng Dây Thường Xuân, Hoa Giấy bám tường. Tạo tấm rèm xanh cách nhiệt, hạ nhiệt độ bề mặt tường bê tông từ 36°C xuống 28°C."),
        ("Tầng Trệt: Cây Ưa Bóng Lọc Độc", "Trồng Cây Lan Ý, Dương Xỉ, Trầu Bà. Lá mỏng giàu diệp lục b quang hợp hiệu quả ở nơi ánh sáng yếu 3.500 Lux, lọc khí độc VOCs."),
        ("Hạ tầng cấp nước:", "Lắp đặt Hệ Thống Tưới Nhỏ Giọt Hẹn Giờ Tiết Kiệm Nước, đảm bảo độ ẩm rễ không bị ngập úng."),
        ("Thao tác chăm sóc sinh học:", "Sử dụng 4 nút: [Tưới Nước], [Kéo Lưới Che Nắng], [Thổi Gió Mát], [Tái lập]. Đạt 4/4 giải pháp tối ưu!")
    ]
    for title, desc in st1_steps:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r = p.add_run()
        r.text = f"• {title}:\n  "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_FOREST
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_SLATE

    # Right: Screenshot Garden
    if os.path.exists("screenshots/garden_planted.png"):
        s12.shapes.add_picture("screenshots/garden_planted.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 13: BÀN LÀM VIỆC CỦA GIÁO VIÊN (CÔ MAI THUÝ ĐỨC)
    # ==========================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s13, C_LIGHT_BG)
    add_header(s13, "Cổng Phê Duyệt Của Giáo Viên: Cô Mai Thuý Đức", "QUY TRÌNH KIỂM SOÁT SƯ PHẠM (TEACHER HUB)")
    add_footer(s13, 13)

    # Left: Explanation
    c_left = create_card(s13, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_AMBER, border_width=1.5)
    tb = s13.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "📋 QUY TRÌNH PHÊ DUYỆT 3 BƯỚC CHẶT CHẼ"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_AMBER

    rev_steps = [
        ("Bước 1: Tiếp nhận hồ sơ số", "Sau khi các nhóm bấm [Nộp Hồ Sơ Thiết Kế], dữ liệu được truyền thời gian thực về Bàn Làm Việc của Giáo viên. Trạng thái chuyển thành: 'Đang Chờ Cô Mai Thuý Đức Duyệt'."),
        ("Bước 2: Giáo viên thẩm định chuyên môn", "Cô Mai Thuý Đức mở hồ sơ từng nhóm để kiểm tra:\n• Tính chính xác của các phương án cây trồng.\n• Mức độ cải thiện các chỉ số vi khí hậu (CO₂, Lux, °C).\n• Lập luận cơ chế quang hợp sinh học theo Bài 23 SGK."),
        ("Bước 3: Cho điểm, nhận xét & Xuất bản E-Poster", "Giáo viên chấm điểm (thang 100đ), nhập lời nhận xét sư phạm động viên và bấm [Phê Duyệt & Chuẩn Y]."),
        ("Chế độ Demo dành cho Giám khảo:", "Nút [Duyệt Nhanh Tất Cả 4 Nhóm] hỗ trợ Giám khảo hội thi kiểm tra nhanh toàn bộ sản phẩm hoàn chỉnh mà không mất thời gian chờ đợi.")
    ]
    for title, desc in rev_steps:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r = p.add_run()
        r.text = f"• {title}:\n  "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_AMBER
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_SLATE

    # Right: Screenshot Poster Gallery
    if os.path.exists("screenshots/exp_4_poster_gallery_approved.png"):
        s13.shapes.add_picture("screenshots/exp_4_poster_gallery_approved.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 14: THUYẾT TRÌNH BÁO CÁO E-POSTER KHOA HỌC
    # ==========================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s14, C_LIGHT_BG)
    add_header(s14, "Triển Lãm & Thuyết Trình Báo Cáo E-Poster Khoa Học", "BÁO CÁO ĐỒ ÁN STEM TOÀN MÀN HÌNH")
    add_footer(s14, 14)

    # Left: Explanation
    c_left = create_card(s14, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_FOREST, border_width=1.5)
    tb = s14.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "📊 CẤU TRÚC E-POSTER 3 CỘT CHUẨN KHOA HỌC"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_FOREST

    poster_cols = [
        ("Cột 1: Đặt vấn đề & Thách thức vi khí hậu", "Mô tả đặc điểm khí hậu trạm được giao (hiệu ứng đảo nhiệt, thiếu sáng tầng trệt, khói bụi) và nguy cơ sinh học với lá cây."),
        ("Cột 2: Phương án kỹ thuật & Cơ chế Bài 23", "Trình bày danh mục cây trồng đã chọn và luận giải cơ chế sinh học: Vì sao cây ưa sáng đặt ở tầng thượng? Cơ chế thoát hơi nước và đóng mở khí khổng diễn ra như thế nào?"),
        ("Cột 3: Kết quả thực nghiệm vi khí hậu & Lời phê GV", "Bảng dữ liệu đo đạc cảm biến IoT: Nồng độ CO₂ giảm về 410 ppm, Lux đạt 25.000, Nhiệt độ hạ còn 28°C, Điểm quang hợp đạt 96/100đ.\nKèm Lời phê chính thức của Cô Mai Thuý Đức."),
        ("Đánh giá đồng đẳng (Peer Review):", "Học sinh các nhóm khác tham gia phản biện câu hỏi và bình chọn 1 - 5 sao trực tiếp trên màn hình E-Poster.")
    ]
    for title, desc in poster_cols:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r = p.add_run()
        r.text = f"• {title}:\n  "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_FOREST
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = C_SLATE

    # Right: Screenshot Poster
    if os.path.exists("screenshots/exp_5_presentation_poster.png"):
        s14.shapes.add_picture("screenshots/exp_5_presentation_poster.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 15: TRAO GIẤY CHỨNG NHẬN KỸ SƯ SINH THÁI SỐ
    # ==========================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s15, C_LIGHT_BG)
    add_header(s15, "Vinh Danh: Trao Giấy Chứng Nhận Kỹ Sư Sinh Học Xanh", "HOẠN TẤT NHIỆM VỤ DỰ ÁN SỐ")
    add_footer(s15, 15)

    # Left: Explanation
    c_left = create_card(s15, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_AMBER, border_width=1.5)
    tb = s15.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🎖️ HỆ THỐNG VINH DANH ECO-POINTS"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_AMBER

    cert_points = [
        ("Cấp bậc vinh danh dựa trên điểm số:", "• 👑 ĐẠI SỨ XANH GDPT 2018 (Từ 480 – 500đ)\n• 🥇 KỸ SƯ SINH THÁI ĐÔ THỊ TP.HCM (Từ 400 – 479đ)\n• 🥈 CHIẾN BINH SINH THÁI (Từ 250 – 399đ)\n• 🥉 MẦM NON KHÁM PHÁ (Dưới 250đ)"),
        ("Các yếu tố pháp lý trên chứng nhận số:", "• Mã định danh xác thực duy nhất: STEM-HCM2026-8892\n• Đơn vị: Trường THCS Tân Tạo A\n• Con dấu đỏ chính thức: ĐÃ NGHIỆM THU DỰ ÁN NHÓM\n• Chữ ký xác nhận: Giáo viên hướng dẫn Cô Mai Thuý Đức."),
        ("Tính năng xuất file ảnh độ nét cao:", "Bấm nút [Tải Giấy Chứng Nhận (PNG)], hệ thống tự động render Canvas 1200 x 850 pixel để in màu hoặc lưu vào hồ sơ học tập số của học sinh.")
    ]
    for title, desc in cert_points:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title}\n"
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = C_AMBER
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_SLATE

    # Right: Screenshot Cert
    if os.path.exists("screenshots/screen_certificate.png"):
        s15.shapes.add_picture("screenshots/screen_certificate.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 16: ĐÁNH GIÁ RUBRIC & HƯỚNG DẪN TỰ HỌC
    # ==========================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s16, C_LIGHT_BG)
    add_header(s16, "Đánh Giá Rubric & Hướng Dẫn Tự Học Mở Rộng", "TỔNG KẾT BÀI DẠY (CV 5512)")
    add_footer(s16, 16)

    # Left: Rubric summary
    c_rubric = create_card(s16, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_FOREST, border_width=1.5)
    tb = s16.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "📝 TIÊU CHÍ ĐÁNH GIÁ RUBRIC (4 MỨC ĐỘ)"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_FOREST

    rub_criteria = [
        ("Tiêu chí 1: Nhận thức kiến thức (25%):", "Nắm vững ảnh hưởng của 4 yếu tố môi trường đến quang hợp; phân biệt rõ cây ưa sáng vs cây ưa bóng."),
        ("Tiêu chí 2: Kỹ năng thực hành số (25%):", "Thao tác thành thạo phòng thí nghiệm ảo, phân tích chính xác đồ thị động học và đọc bảng đo cảm biến vi khí hậu."),
        ("Tiêu chí 3: Sản phẩm E-Poster STEM (30%):", "Đồ án không gian xanh tối ưu, đạt hiệu quả quang hợp >85%, cấu trúc poster 3 cột thẩm mỹ, chuẩn mực khoa học."),
        ("Tiêu chí 4: Hợp tác & Báo cáo (20%):", "Làm việc nhóm trách nhiệm, tự tin thuyết trình, lập luận logic và trả lời xuất sắc câu hỏi phản biện của bạn.")
    ]
    for title, desc in rub_criteria:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_FOREST
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_SLATE

    # Right: Homework
    c_hw = create_card(s16, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=C_ACCENT_BLUE, border_width=1.5)
    tb = s16.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🏠 NHIỆM VỤ TỰ HỌC & MỞ RỘNG VỀ NHÀ"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = C_ACCENT_BLUE

    hw_tasks = [
        ("Nhiệm vụ 1: Hoàn thiện vở ghi chép KHTN 7", "Hệ thống hóa sơ đồ tư duy về mối quan hệ giữa 4 yếu tố: Ánh sáng, Nước, CO₂ và Nhiệt độ đối với quang hợp."),
        ("Nhiệm vụ 2: Dự án gia đình 'Góc Ban Công Xanh'", "Khảo sát một góc ban công, cửa sổ hoặc sân vườn nhà em:\n• Xác định hướng nắng (Đông, Tây, Nam, Bắc) và ước tính cường độ ánh sáng.\n• Đề xuất bố trí ít nhất 2 loại cây ưa sáng và 2 loại cây ưa bóng phù hợp."),
        ("Nhiệm vụ 3: Tự học trực tuyến", "Truy cập https://sinhhoc-amber.vercel.app/ để:\n• Ôn luyện bộ câu hỏi trắc nghiệm.\n• Thử nghiệm thiết kế thêm 3 Trạm vi khí hậu còn lại."),
        ("Chuẩn bị bài mới:", "Đọc trước Bài 24: Thực hành: Chứng minh quang hợp ở cây xanh.")
    ]
    for title, desc in hw_tasks:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = C_ACCENT_BLUE
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = C_SLATE

    # ==========================================================
    # SLIDE 17: TRANG KẾT THÚC (THANK YOU)
    # ==========================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s17, C_DARK_BG)

    # Card Banner
    card_end = create_card(s17, 0.8, 0.7, 11.733, 6.1, bg_color=C_FOREST, border_color=RGBColor(46, 93, 66), border_width=1.5)

    if os.path.exists("images/logo.png"):
        s17.shapes.add_picture("images/logo.png", Inches(5.9), Inches(1.2), width=Inches(1.5), height=Inches(1.5))

    tb_end = s17.shapes.add_textbox(Inches(1.3), Inches(3.0), Inches(10.7), Inches(3.4))
    tf = tb_end.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "XIN TRÂN TRỌNG CẢM ƠN!"
    r.font.bold = True
    r.font.size = Pt(32)
    r.font.color.rgb = C_WHITE
    r.font.name = "Georgia"

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(10)
    r2 = p2.add_run()
    r2.text = "Kính chúc quý Thầy Cô Ban Giám Khảo dồi dào sức khỏe & thành công!\nChúc các em học sinh luôn giữ trọn niềm đam mê khám phá Khoa học tự nhiên!"
    r2.font.size = Pt(14)
    r2.font.color.rgb = RGBColor(167, 243, 208)
    r2.font.name = "Arial"

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(20)
    r3 = p3.add_run()
    r3.text = "GIÁO VIÊN: MAI THUÝ ĐỨC  •  TRƯỜNG THCS TÂN TẠO A (BÌNH TÂN, TP.HCM)\n"
    r3.font.bold = True
    r3.font.size = Pt(13)
    r3.font.color.rgb = RGBColor(254, 243, 199)

    r4 = p3.add_run()
    r4.text = "Nền tảng học liệu số mở: https://sinhhoc-amber.vercel.app/"
    r4.font.bold = True
    r4.font.size = Pt(12)
    r4.font.color.rgb = C_EMERALD

    # Save presentation
    output_filename = "KHTN7_Bai23_QuangHop_MaiThuyDuc.pptx"
    prs.save(output_filename)
    print(f"Presentation successfully created: {output_filename}")

if __name__ == "__main__":
    create_deck()
