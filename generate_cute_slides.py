import os
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE


def create_cute_deck():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # --- SOFT BOTANICAL CLASSROOM PALETTE ---
    BG_SOFT_MINT = RGBColor(243, 249, 245)    # #f3f9f5
    BG_CREAM     = RGBColor(254, 252, 248)    # #fefcf8
    BG_TITLE     = RGBColor(234, 247, 239)    # #eaf7ef

    C_WHITE      = RGBColor(255, 255, 255)
    C_PILL_MINT  = RGBColor(209, 250, 229)    # #d1fae5
    C_PILL_SKY   = RGBColor(224, 242, 254)    # #e0f2fe
    C_PILL_PEACH = RGBColor(254, 243, 199)    # #fef3c7
    C_PILL_ROSE  = RGBColor(255, 228, 230)    # #ffe4e6

    BORDER_MINT  = RGBColor(110, 231, 183)
    BORDER_SKY   = RGBColor(125, 211, 252)
    BORDER_AMBER = RGBColor(252, 211, 77)
    BORDER_ROSE  = RGBColor(253, 164, 175)

    T_PRIMARY    = RGBColor(20, 56, 41)       # Deep forest green
    T_GREEN      = RGBColor(13, 110, 72)
    T_BLUE       = RGBColor(3, 105, 161)
    T_AMBER      = RGBColor(180, 83, 9)
    T_ROSE       = RGBColor(190, 24, 93)
    T_BODY       = RGBColor(30, 41, 59)
    T_MUTED      = RGBColor(71, 85, 105)

    FONT_MAIN = "Arial"
    TOTAL_SLIDES = 20

    def set_slide_bg(slide, color):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def create_card(slide, left, top, width, height, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=2.0):
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        card.adjustments[0] = 0.04
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        if border_color:
            card.line.color.rgb = border_color
            card.line.width = Pt(border_width)
        else:
            card.line.fill.background()
        return card

    def add_header(slide, title_text, badge_text, slide_num, title_size=30):
        # Top row: section badge on left (24pt) + slide counter pill on right (24pt)
        tb_top = slide.shapes.add_textbox(Inches(0.6), Inches(0.22), Inches(9.5), Inches(0.48))
        tf_top = tb_top.text_frame
        tf_top.word_wrap = False
        tf_top.margin_left = tf_top.margin_top = tf_top.margin_right = tf_top.margin_bottom = Inches(0)
        p_top = tf_top.paragraphs[0]
        r_top = p_top.add_run()
        r_top.text = badge_text
        r_top.font.name = FONT_MAIN
        r_top.font.bold = True
        r_top.font.size = Pt(24)
        r_top.font.color.rgb = T_GREEN

        tb_num = slide.shapes.add_textbox(Inches(10.3), Inches(0.22), Inches(2.433), Inches(0.48))
        tf_num = tb_num.text_frame
        tf_num.word_wrap = False
        tf_num.margin_left = tf_num.margin_top = tf_num.margin_right = tf_num.margin_bottom = Inches(0)
        p_num = tf_num.paragraphs[0]
        p_num.alignment = PP_ALIGN.RIGHT
        r_num = p_num.add_run()
        r_num.text = f"{slide_num}/{TOTAL_SLIDES}"
        r_num.font.name = FONT_MAIN
        r_num.font.bold = True
        r_num.font.size = Pt(24)
        r_num.font.color.rgb = T_MUTED

        # Main slide title (30pt - 34pt)
        tb_title = slide.shapes.add_textbox(Inches(0.6), Inches(0.72), Inches(12.133), Inches(0.68))
        tf_t = tb_title.text_frame
        tf_t.word_wrap = True
        tf_t.margin_left = tf_t.margin_top = tf_t.margin_right = tf_t.margin_bottom = Inches(0)
        p_t = tf_t.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title_text
        r_t.font.name = FONT_MAIN
        r_t.font.bold = True
        r_t.font.size = Pt(title_size)
        r_t.font.color.rgb = T_PRIMARY

        # Accent divider bar at y=1.42"
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.6), Inches(1.42), Inches(12.133), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(167, 243, 208)
        line.line.fill.background()

    def make_textbox(slide, left, top, width, height):
        tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.06)
        tf.margin_right = Inches(0.06)
        tf.margin_top = Inches(0.04)
        tf.margin_bottom = Inches(0.04)
        return tf

    def add_para(tf, runs_spec, space_before=10, align=PP_ALIGN.LEFT, first=False):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        p.alignment = align
        if not first and space_before:
            p.space_before = Pt(space_before)
        for txt, sz, bold, col in runs_spec:
            r = p.add_run()
            r.text = txt
            r.font.name = FONT_MAIN
            r.font.size = Pt(sz)
            r.font.bold = bold
            r.font.color.rgb = col
        return p

    def add_picture_framed(slide, path, left, top, width, height, border_col=BORDER_MINT):
        create_card(slide, left, top, width, height, bg_color=C_WHITE, border_color=border_col, border_width=2.0)
        if os.path.exists(path):
            slide.shapes.add_picture(
                path,
                Inches(left + 0.12), Inches(top + 0.12),
                width=Inches(width - 0.24), height=Inches(height - 0.24)
            )

    # ==========================================================
    # SLIDE 1: TRANG BÌA (SPARSE -> 25pt - 40pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_TITLE)
    create_card(s, 0.6, 0.5, 12.133, 6.5, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=2.5)

    if os.path.exists("images/logo.png"):
        s.shapes.add_picture("images/logo.png", Inches(1.05), Inches(0.85), width=Inches(1.1), height=Inches(1.1))

    tf = make_textbox(s, 2.35, 0.95, 9.9, 0.95)
    add_para(tf, [("KHOA HỌC TỰ NHIÊN 7 • KẾT NỐI TRI THỨC VỚI CUỘC SỐNG", 25, True, T_GREEN)], first=True)

    tf_main = make_textbox(s, 1.0, 2.05, 11.333, 2.55)
    add_para(tf_main, [("BÀI 23: MỘT SỐ YẾU TỐ\nẢNH HƯỞNG ĐẾN QUANG HỢP", 40, True, T_PRIMARY)], align=PP_ALIGN.CENTER, first=True)
    add_para(tf_main, [("Chủ đề STEM & Học liệu số tương tác (Thời lượng: 02 tiết)", 26, True, T_GREEN)], space_before=14, align=PP_ALIGN.CENTER)

    create_card(s, 1.0, 4.8, 11.333, 1.85, bg_color=C_PILL_MINT, border_color=BORDER_MINT, border_width=1.8)
    tf_info = make_textbox(s, 1.25, 4.95, 10.833, 1.55)
    add_para(tf_info, [
        ("Giáo viên: ", 25, False, T_BODY),
        ("MAI THUÝ ĐỨC", 26, True, T_PRIMARY),
        ("   •   Trường: ", 25, False, T_BODY),
        ("THCS TÂN TẠO A", 26, True, T_PRIMARY)
    ], align=PP_ALIGN.CENTER, first=True)
    add_para(tf_info, [
        ("Học liệu số: ", 24, False, T_BODY),
        ("https://sinhhoc-amber.vercel.app/", 25, True, T_GREEN)
    ], space_before=10, align=PP_ALIGN.CENTER)

    # ==========================================================
    # SLIDE 2: MỤC TIÊU BÀI HỌC (3 COLUMNS -> 24pt - 32pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Mục Tiêu Cần Đạt Của Chủ Đề 🎯", "KHTN 7 • BÀI 23", 2, title_size=32)

    cols_s2 = [
        ("🌱 1. KIẾN THỨC", BORDER_MINT, T_GREEN, [
            "- 4 yếu tố: Ánh sáng, Nước, CO₂, Nhiệt độ.",
            "- Phân biệt cây ưa sáng và cây ưa bóng.",
            "- Giải thích cách trồng, chăm sóc cây."
        ]),
        ("💻 2. NĂNG LỰC", BORDER_SKY, T_BLUE, [
            "- Khảo sát thí nghiệm số rong đuôi chó.",
            "- Thiết kế 4 Trạm không gian xanh.",
            "- Hợp tác nhóm và báo cáo E-Poster."
        ]),
        ("💖 3. PHẨM CHẤT", BORDER_AMBER, T_AMBER, [
            "- Chăm chỉ đọc bài và thử nghiệm.",
            "- Trung thực khi ghi nhận số liệu.",
            "- Có trách nhiệm bảo vệ cây xanh."
        ])
    ]
    for idx, (head, bcol, tcol, bullets) in enumerate(cols_s2):
        x = 0.6 + idx * 4.11
        create_card(s, x, 1.62, 3.91, 5.55, bg_color=C_WHITE, border_color=bcol)
        tf = make_textbox(s, x + 0.16, 1.80, 3.59, 5.20)
        add_para(tf, [(head, 26, True, tcol)], first=True)
        for b in bullets:
            add_para(tf, [(b, 24, False, T_BODY)], space_before=14)

    # ==========================================================
    # SLIDE 3: TIẾN TRÌNH 2 TIẾT HỌC (2 SPACIOUS CARDS -> 25pt - 32pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Tiến Trình Tổ Chức 02 Tiết Học (90 Phút) 🗺️", " LỘ TRÌNH HỌC TẬP", 3, title_size=32)

    create_card(s, 0.6, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf1 = make_textbox(s, 0.82, 1.76, 5.48, 5.28)
    add_para(tf1, [("🌟 TIẾT 1: KHÁM PHÁ KIẾN THỨC", 27, True, T_GREEN)], first=True)
    add_para(tf1, [("• Hoạt động 1 (07 phút):", 26, True, T_PRIMARY)], space_before=14)
    add_para(tf1, [("  Tình huống hai chậu cây & Tháp xanh.", 25, False, T_BODY)], space_before=4)
    add_para(tf1, [("• Hoạt động 2 (38 phút):", 26, True, T_PRIMARY)], space_before=14)
    add_para(tf1, [("  Thí nghiệm số 4 yếu tố trên cành rong;\n  hoàn thành Phiếu học tập số 1.", 25, False, T_BODY)], space_before=4)

    create_card(s, 6.81, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_AMBER)
    tf2 = make_textbox(s, 7.03, 1.76, 5.48, 5.28)
    add_para(tf2, [("🚀 TIẾT 2: LUYỆN TẬP & DỰ ÁN STEM", 27, True, T_AMBER)], first=True)
    add_para(tf2, [("• Hoạt động 3 (15 phút):", 26, True, T_PRIMARY)], space_before=14)
    add_para(tf2, [("  Trắc nghiệm tình huống chăm sóc cây.", 25, False, T_BODY)], space_before=4)
    add_para(tf2, [("• Hoạt động 4 (30 phút):", 26, True, T_PRIMARY)], space_before=14)
    add_para(tf2, [("  Thiết kế xanh tại 4 Trạm thực nghiệm;\n  ghi Phiếu số 2 & báo cáo E-Poster.", 25, False, T_BODY)], space_before=4)

    # ==========================================================
    # SLIDE 4: HOẠT ĐỘNG 1 - TÌNH HUỐNG HAI BẠN CÂY (24pt - 30pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Tình Huống: Vì Sao Hai Chậu Cây Đều Suy Yếu? 🧐", "HOẠT ĐỘNG 1 • MỞ ĐẦU", 4, title_size=30)

    # Left card: Lan Ý
    create_card(s, 0.6, 1.60, 5.92, 4.45, bg_color=C_WHITE, border_color=BORDER_AMBER)
    if os.path.exists("images/cay_san_thuong.png"):
        s.shapes.add_picture("images/cay_san_thuong.png", Inches(0.80), Inches(1.85), width=Inches(2.15), height=Inches(2.35))
    tf_l = make_textbox(s, 3.05, 1.78, 3.32, 4.10)
    add_para(tf_l, [("☀️ SÂN THƯỢNG", 25, True, T_AMBER)], first=True)
    add_para(tf_l, [("Cây Lan Ý (Ưa bóng)", 25, True, T_PRIMARY)], space_before=8)
    add_para(tf_l, [("- Sáng: 48.000 Lux\n- Nhiệt độ: 39°C\n- Đủ nước tưới", 24, False, T_BODY)], space_before=10)
    add_para(tf_l, [("→ Cháy xém mép lá!", 25, True, T_ROSE)], space_before=10)

    # Right card: Bàng Cảnh
    create_card(s, 6.81, 1.60, 5.92, 4.45, bg_color=C_WHITE, border_color=BORDER_SKY)
    if os.path.exists("images/cay_phong_khach.png"):
        s.shapes.add_picture("images/cay_phong_khach.png", Inches(7.01), Inches(1.85), width=Inches(2.15), height=Inches(2.35))
    tf_r = make_textbox(s, 9.26, 1.78, 3.32, 4.10)
    add_para(tf_r, [("🏡 GÓC PHÒNG TỐI", 25, True, T_BLUE)], first=True)
    add_para(tf_r, [("Bàng Cảnh (Ưa sáng)", 25, True, T_PRIMARY)], space_before=8)
    add_para(tf_r, [("- Sáng: 650 Lux\n- Nhiệt độ: 25°C\n- Đủ nước tưới", 24, False, T_BODY)], space_before=10)
    add_para(tf_r, [("→ Vàng và rụng lá!", 25, True, T_BLUE)], space_before=10)

    # Bottom prompt bar
    create_card(s, 0.6, 6.20, 12.133, 1.02, bg_color=C_PILL_MINT, border_color=BORDER_MINT)
    tf_b = make_textbox(s, 0.75, 6.32, 11.833, 0.80)
    add_para(tf_b, [("Thảo luận: Nếu đổi chỗ hai chậu cây cho nhau thì kết quả ra sao?", 25, True, T_PRIMARY)], align=PP_ALIGN.CENTER, first=True)

    # ==========================================================
    # SLIDE 5: KẾT QUẢ HOÁN ĐỔI VỊ TRÍ & CÂU HỎI BÀI HỌC (26pt - 32pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Thử Nghiệm Đổi Vị Trí Hai Chậu Cây Trên Web 🎉", "HOẠT ĐỘNG 1 • XÁC ĐỊNH VẤN ĐỀ", 5, title_size=31)

    create_card(s, 0.6, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf1 = make_textbox(s, 0.85, 1.82, 5.42, 5.15)
    add_para(tf1, [("✅ KẾT QUẢ SAU KHI ĐỔI CHỖ", 27, True, T_GREEN)], first=True)
    add_para(tf1, [("- Cây Lan Ý vào phòng mát (650 Lux):\n  Lá xanh tươi, phục hồi tốt.", 26, False, T_BODY)], space_before=18)
    add_para(tf1, [("- Cây Bàng Cảnh ra ban công nắng:\n  Quang hợp mạnh, lá vươn khỏe.", 26, False, T_BODY)], space_before=18)

    create_card(s, 6.81, 1.62, 5.92, 5.55, bg_color=C_PILL_PEACH, border_color=BORDER_AMBER)
    tf2 = make_textbox(s, 7.06, 1.82, 5.42, 5.15)
    add_para(tf2, [("❓ VẤN ĐỀ CẦN NGHIÊN CỨU", 27, True, T_AMBER)], first=True)
    add_para(tf2, [("Ánh sáng, Nước, Khí CO₂ và Nhiệt độ ảnh hưởng như thế nào đến tốc độ quang hợp của thực vật?", 27, True, T_PRIMARY)], space_before=20)
    add_para(tf2, [("→ Cùng khảo sát tại Phòng thí nghiệm số!", 25, True, T_GREEN)], space_before=20)

    # ==========================================================
    # SLIDE 6: TƯ LIỆU SGK TRANG 107 - BOSCO VERTICALE (25pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Tư Liệu SGK Trang 107: Tháp Rừng Bosco Verticale 🏢", "LIÊN HỆ THỰC TIỄN", 6, title_size=30)

    create_card(s, 0.6, 1.62, 6.2, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf = make_textbox(s, 0.84, 1.80, 5.72, 5.20)
    add_para(tf, [("🌿 TÒA THÁP XANH TẠI MILAN (Ý)", 26, True, T_GREEN)], first=True)
    add_para(tf, [("- Trồng ~900 cây thân gỗ, 5.000 cây bụi.", 25, False, T_BODY)], space_before=14)
    add_para(tf, [("- Hấp thụ CO₂, giảm nhiệt bức xạ bê tông.", 25, False, T_BODY)], space_before=12)
    add_para(tf, [("- Lọc sạch bụi mịn PM2.5 cho đô thị.", 25, False, T_BODY)], space_before=12)
    add_para(tf, [("💡 Nóc tháp nắng gắt nên chọn cây ưa sáng hay cây ưa bóng?", 25, True, T_AMBER)], space_before=16)

    add_picture_framed(s, "screenshots/bosco_updated.png", 7.0, 1.62, 5.733, 5.55, BORDER_MINT)

    # ==========================================================
    # SLIDE 7: YẾU TỐ 1 - ÁNH SÁNG (24pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "1. Ảnh Hưởng Của Ánh Sáng Đến Quang Hợp ☀️", "HOẠT ĐỘNG 2 • THÍ NGHIỆM SỐ", 7, title_size=31)

    create_card(s, 0.6, 1.62, 6.3, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf = make_textbox(s, 0.82, 1.78, 5.86, 5.20)
    add_para(tf, [("🌿 THÍ NGHIỆM RONG ĐUÔI CHÓ", 26, True, T_GREEN)], first=True)
    add_para(tf, [("- Tăng sáng → bọt khí O₂ tăng nhanh\n  (đạt tối ưu quanh ~25.000 Lux).", 24, False, T_BODY)], space_before=12)
    add_para(tf, [("- Quá gắt (>45.000 Lux): giảm quang hợp.", 24, False, T_BODY)], space_before=10)
    add_para(tf, [("- Cây ưa sáng: lá dày, cần nơi nắng mạnh.", 24, True, T_PRIMARY)], space_before=12)
    add_para(tf, [("- Cây ưa bóng: lá mỏng, xanh sẫm, ưa râm.", 24, True, T_GREEN)], space_before=10)

    add_picture_framed(s, "screenshots/lab_elodea_realistic.png", 7.1, 1.62, 5.633, 5.55, BORDER_MINT)

    # ==========================================================
    # SLIDE 8: YẾU TỐ 2 - NƯỚC (25pt - 32pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "2. Ảnh Hưởng Của Nước Đến Quang Hợp 💧", "HOẠT ĐỘNG 2 • THÍ NGHIỆM SỐ", 8, title_size=32)

    create_card(s, 0.6, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_SKY)
    tf1 = make_textbox(s, 0.84, 1.82, 5.44, 5.15)
    add_para(tf1, [("💧 VAI TRÒ CỦA NƯỚC", 28, True, T_BLUE)], first=True)
    add_para(tf1, [("- Là nguyên liệu trực tiếp của quang hợp (tạo chất hữu cơ và khí O₂).", 26, False, T_BODY)], space_before=20)
    add_para(tf1, [("- Thoát hơi nước qua lá giúp điều hòa, làm giảm nhiệt độ bề mặt lá.", 26, False, T_BODY)], space_before=20)

    create_card(s, 6.81, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf2 = make_textbox(s, 7.05, 1.82, 5.44, 5.15)
    add_para(tf2, [("🔬 ĐIỀU TIẾT ĐỘ MỞ KHÍ KHỔNG", 28, True, T_GREEN)], first=True)
    add_para(tf2, [("- Đủ nước: Khí khổng mở rộng → khí CO₂ khuếch tán vào lá thuận lợi.", 26, False, T_BODY)], space_before=18)
    add_para(tf2, [("- Thiếu nước / ngập úng: Khí khổng đóng lại → quang hợp giảm hoặc ngừng.", 26, False, T_BODY)], space_before=18)

    # ==========================================================
    # SLIDE 9: YẾU TỐ 3 & 4 - KHÍ CO2 VÀ NHIỆT ĐỘ (25pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "3 & 4. Hàm Lượng Khí CO₂ Và Nhiệt Độ ☁️🌡️", "HOẠT ĐỘNG 2 • THÍ NGHIỆM SỐ", 9, title_size=31)

    create_card(s, 0.6, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_AMBER)
    tf1 = make_textbox(s, 0.84, 1.82, 5.44, 5.15)
    add_para(tf1, [("☁️ HÀM LƯỢNG KHÍ CO₂", 27, True, T_AMBER)], first=True)
    add_para(tf1, [("- Không khí tự nhiên: ~0,03% – 0,04%.", 25, False, T_BODY)], space_before=16)
    add_para(tf1, [("- Tăng CO₂ (0,03% – 0,15%): tốc độ quang hợp tăng.", 25, False, T_BODY)], space_before=14)
    add_para(tf1, [("- Quá thấp (<0,008%) hoặc quá cao (>0,2%): ngừng quang hợp / ngộ độc.", 25, False, T_BODY)], space_before=14)

    create_card(s, 6.81, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_ROSE)
    tf2 = make_textbox(s, 7.05, 1.82, 5.44, 5.15)
    add_para(tf2, [("🌡️ NHIỆT ĐỘ MÔI TRƯỜNG", 27, True, T_ROSE)], first=True)
    add_para(tf2, [("- Thuận lợi nhất: 25°C – 35°C.", 25, False, T_BODY)], space_before=16)
    add_para(tf2, [("- Quá thấp (<10°C): enzyme giảm hoạt tính, quang hợp chậm.", 25, False, T_BODY)], space_before=14)
    add_para(tf2, [("- Quá cao (>40°C): biến tính enzyme, cây ngừng quang hợp.", 25, False, T_BODY)], space_before=14)

    # ==========================================================
    # SLIDE 10: BẢNG TỔNG KẾT TIẾT 1 (TABLE 24pt MIN)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Bảng Tổng Kết Kiến Thức Trọng Tâm Bài 23 📋", "CHỐT KIẾN THỨC TIẾT 1", 10, title_size=31)

    t_rows = [
        ("Yếu tố", "Ngưỡng thuận lợi", "Khi vượt ngưỡng / bất lợi"),
        ("☀️ Ánh sáng", "15.000 – 30.000 Lux", "Tối (0 Lux) hoặc gắt >45.000 Lux"),
        ("💧 Nước", "Độ ẩm đất 60% – 75%", "Khô hạn hoặc ngập úng thối rễ"),
        ("☁️ Khí CO₂", "0,03% – 0,15%", "<0,008% hoặc >0,20% (ngộ độc)"),
        ("🌡️ Nhiệt độ", "25°C – 35°C", "<10°C hoặc >40°C (hỏng enzyme)")
    ]
    tbl_sh = s.shapes.add_table(5, 3, Inches(0.6), Inches(1.62), Inches(12.133), Inches(5.55))
    tbl = tbl_sh.table
    tbl.columns[0].width = Inches(2.85)
    tbl.columns[1].width = Inches(4.15)
    tbl.columns[2].width = Inches(5.133)

    for r_i, row_vals in enumerate(t_rows):
        tbl.rows[r_i].height = Inches(1.11)
        for c_i, val in enumerate(row_vals):
            cell = tbl.cell(r_i, c_i)
            cell.text = ""
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r_i == 0:
                cell.fill.fore_color.rgb = C_PILL_MINT
            else:
                cell.fill.fore_color.rgb = C_WHITE if r_i % 2 == 1 else BG_CREAM
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if (r_i == 0 or c_i == 1) else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.name = FONT_MAIN
            r.font.size = Pt(25 if r_i == 0 else 24)
            r.font.bold = (r_i == 0 or c_i == 0)
            r.font.color.rgb = T_PRIMARY if r_i == 0 else (T_GREEN if c_i == 0 else T_BODY)

    # ==========================================================
    # SLIDE 11: LUYỆN TẬP NHANH (CÂU 1 & CÂU 2 -> 25pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Luyện Tập Tình Huống Thực Tế (Phần 1) 🎮", "HOẠT ĐỘNG 3 • ĐẦU TIẾT 2", 11, title_size=31)

    create_card(s, 0.6, 1.62, 12.133, 2.65, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf1 = make_textbox(s, 0.85, 1.76, 11.633, 2.38)
    add_para(tf1, [("Câu 1: Yếu tố nào vừa là nguyên liệu quang hợp, vừa điều tiết đóng mở khí khổng?", 26, True, T_PRIMARY)], first=True)
    add_para(tf1, [("A. Khí CO₂      B. Nước (H₂O) ✅      C. Nhiệt độ      D. Gió", 25, True, T_GREEN)], space_before=16)

    create_card(s, 0.6, 4.50, 12.133, 2.68, bg_color=C_WHITE, border_color=BORDER_AMBER)
    tf2 = make_textbox(s, 0.85, 4.64, 11.633, 2.40)
    add_para(tf2, [("Câu 2: Giữa trưa nắng gắt (>45.000 Lux, >40°C), vì sao quang hợp giảm mạnh?", 26, True, T_PRIMARY)], first=True)
    add_para(tf2, [("A. Cây tăng hút nước      B. Khí khổng đóng và biến tính enzyme ✅", 25, True, T_AMBER)], space_before=16)

    # ==========================================================
    # SLIDE 12: LUYỆN TẬP NHANH (CÂU 3 & CÂU 4 -> 25pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Luyện Tập Tình Huống Thực Tế (Phần 2) ⭐", "HOẠT ĐỘNG 3 • CỦNG CỐ", 12, title_size=31)

    create_card(s, 0.6, 1.62, 12.133, 2.65, bg_color=C_WHITE, border_color=BORDER_SKY)
    tf1 = make_textbox(s, 0.85, 1.76, 11.633, 2.38)
    add_para(tf1, [("Câu 3: Cây trồng bị ngộ độc khi nồng độ khí CO₂ vượt quá ngưỡng nào?", 26, True, T_PRIMARY)], first=True)
    add_para(tf1, [("A. Khoảng 0,04%      B. Khoảng 0,10%      C. Vượt quá 0,20% ✅", 25, True, T_BLUE)], space_before=16)

    create_card(s, 0.6, 4.50, 12.133, 2.68, bg_color=C_WHITE, border_color=BORDER_ROSE)
    tf2 = make_textbox(s, 0.85, 4.64, 11.633, 2.40)
    add_para(tf2, [("Câu 4: Cây lưỡi hổ, trầu bà (nhóm ưa bóng) phù hợp đặt ở vị trí nào?", 26, True, T_PRIMARY)], first=True)
    add_para(tf2, [("A. Mái sân thượng nắng gắt      B. Hành lang, phòng học râm mát ✅", 25, True, T_ROSE)], space_before=16)

    # ==========================================================
    # SLIDE 13: DỰ ÁN STEM - NHÓM 1 & NHÓM 2 (24pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Dự Án STEM: Nhiệm Vụ Nhóm 1 & Nhóm 2 🚀", "HOẠT ĐỘNG 4 • THIẾT KẾ 4 TRẠM", 13, title_size=31)

    create_card(s, 0.6, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf1 = make_textbox(s, 0.84, 1.80, 5.44, 5.20)
    add_para(tf1, [("🌱 NHÓM 1 — TRẠM 1", 26, True, T_GREEN)], first=True)
    add_para(tf1, [("Đô Thị Nén & Tháp Rừng Đứng", 25, True, T_PRIMARY)], space_before=6)
    add_para(tf1, [("- Hiện trạng: CO₂ 850 ppm, nóng 36°C,\n  tầng mái 45.000 Lux, tầng trệt râm.", 24, False, T_BODY)], space_before=14)
    add_para(tf1, [("- Nhiệm vụ: Phân tầng cây ưa sáng, giàn dây leo, cây ưa bóng & tưới nhỏ giọt.", 24, False, T_GREEN)], space_before=14)

    create_card(s, 6.81, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_SKY)
    tf2 = make_textbox(s, 7.05, 1.80, 5.44, 5.20)
    add_para(tf2, [("🌱 NHÓM 2 — TRẠM 2", 26, True, T_BLUE)], first=True)
    add_para(tf2, [("Nhà Màng Nông Nghiệp Số (Củ Chi)", 25, True, T_PRIMARY)], space_before=6)
    add_para(tf2, [("- Hiện trạng: Sáng 12.000 Lux (thiếu),\n  CO₂ 400 ppm, nhiệt độ 34°C.", 24, False, T_BODY)], space_before=14)
    add_para(tf2, [("- Nhiệm vụ: Bổ sung đèn LED, tăng CO₂ (1.200 ppm), quạt mát 26°C & tưới tự động.", 24, False, T_BLUE)], space_before=14)

    # ==========================================================
    # SLIDE 14: DỰ ÁN STEM - NHÓM 3 & NHÓM 4 (24pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Dự Án STEM: Nhiệm Vụ Nhóm 3 & Nhóm 4 🌿", "HOẠT ĐỘNG 4 • THIẾT KẾ 4 TRẠM", 14, title_size=31)

    create_card(s, 0.6, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_AMBER)
    tf1 = make_textbox(s, 0.84, 1.80, 5.44, 5.20)
    add_para(tf1, [("☀️ NHÓM 3 — TRẠM 3", 26, True, T_AMBER)], first=True)
    add_para(tf1, [("Sân Thượng Nắng Gắt Hướng Tây", 25, True, T_PRIMARY)], space_before=6)
    add_para(tf1, [("- Hiện trạng: Nắng 58.000 Lux, sàn bê tông 43°C, khí khổng đóng chặt.", 24, False, T_BODY)], space_before=14)
    add_para(tf1, [("- Nhiệm vụ: Căng lưới lan, chọn cây chịu hạn cutin dày, phun sương & phủ gốc.", 24, False, T_AMBER)], space_before=14)

    create_card(s, 6.81, 1.62, 5.92, 5.55, bg_color=C_WHITE, border_color=BORDER_ROSE)
    tf2 = make_textbox(s, 7.05, 1.80, 5.44, 5.20)
    add_para(tf2, [("🏡 NHÓM 4 — TRẠM 4", 26, True, T_ROSE)], first=True)
    add_para(tf2, [("Trong Nhà & Hành Lang Râm Mát", 25, True, T_PRIMARY)], space_before=6)
    add_para(tf2, [("- Hiện trạng: Sáng yếu 800 Lux, thiếu thông gió tự nhiên, CO₂ 650 ppm.", 24, False, T_BODY)], space_before=14)
    add_para(tf2, [("- Nhiệm vụ: Chọn cây ưa bóng giàu diệp lục b, đo sáng, đèn LED 4000K & lau lá.", 24, False, T_ROSE)], space_before=14)

    # ==========================================================
    # SLIDE 15: THAO TÁC THỰC HÀNH TRÊN WEB & GHI PHIẾU SỐ 2 (24pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Các Bước Thực Hành Tại Trạm (12 Phút) 💻", "HƯỚNG DẪN THAO TÁC", 15, title_size=31)

    create_card(s, 0.6, 1.62, 6.1, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf = make_textbox(s, 0.84, 1.80, 5.62, 5.20)
    add_para(tf, [("📋 QUY TRÌNH 4 BƯỚC CỦA NHÓM", 26, True, T_GREEN)], first=True)
    add_para(tf, [("1. Ghi chỉ số ban đầu vào Phiếu số 2.", 25, False, T_BODY)], space_before=16)
    add_para(tf, [("2. Đưa lần lượt 4 giải pháp vào trạm.", 25, False, T_BODY)], space_before=14)
    add_para(tf, [("3. Đo và ghi CO₂, Lux, °C sau mỗi bước.", 25, False, T_BODY)], space_before=14)
    add_para(tf, [("4. Tính độ chênh lệch & bấm Gửi Báo Cáo.", 25, True, T_PRIMARY)], space_before=14)

    add_picture_framed(s, "screenshots/garden_planted.png", 6.9, 1.62, 5.833, 5.55, BORDER_MINT)

    # ==========================================================
    # SLIDE 16: CỔNG DUYỆT BÀI CỦA GIÁO VIÊN (25pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Giáo Viên Kiểm Tra & Phê Duyệt Hồ Sơ 👩‍🏫", "QUẢN LÝ BÀI NỘP TRỰC TUYẾN", 16, title_size=31)

    create_card(s, 0.6, 1.62, 6.1, 5.55, bg_color=C_WHITE, border_color=BORDER_AMBER)
    tf = make_textbox(s, 0.84, 1.85, 5.62, 5.10)
    add_para(tf, [("💌 TỔNG HỢP KẾT QUẢ 4 NHÓM", 26, True, T_AMBER)], first=True)
    add_para(tf, [("- Kiểm tra đủ 4 giải pháp tối ưu (100%).", 25, False, T_BODY)], space_before=18)
    add_para(tf, [("- Đối chiếu các chỉ số CO₂, Lux, °C.", 25, False, T_BODY)], space_before=16)
    add_para(tf, [("- Giáo viên nhận xét và mở khóa E-Poster lên màn hình lớn của lớp.", 25, False, T_BODY)], space_before=16)

    add_picture_framed(s, "screenshots/exp_4_poster_gallery_approved.png", 6.9, 1.62, 5.833, 5.55, BORDER_AMBER)

    # ==========================================================
    # SLIDE 17: BÁO CÁO & THẢO LUẬN E-POSTER (25pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Báo Cáo Sản Phẩm E-Poster (2,5 Phút/Nhóm) 🎙️", "THUYẾT TRÌNH & PHẢN BIỆN", 17, title_size=30)

    create_card(s, 0.6, 1.62, 6.1, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf = make_textbox(s, 0.84, 1.82, 5.62, 5.15)
    add_para(tf, [("📊 NỘI DUNG BÁO CÁO NHÓM", 26, True, T_GREEN)], first=True)
    add_para(tf, [("1. Thách thức vi khí hậu ban đầu.", 25, False, T_BODY)], space_before=16)
    add_para(tf, [("2. Giải thích cơ sở chọn cây (Bài 23).", 25, False, T_BODY)], space_before=14)
    add_para(tf, [("3. Kết quả cải thiện CO₂, Lux, °C.", 25, False, T_BODY)], space_before=14)
    add_para(tf, [("4. Trao đổi câu hỏi cùng nhóm bạn.", 25, True, T_PRIMARY)], space_before=14)

    add_picture_framed(s, "screenshots/exp_5_presentation_poster.png", 6.9, 1.62, 5.833, 5.55, BORDER_MINT)

    # ==========================================================
    # SLIDE 18: TIÊU CHÍ ĐÁNH GIÁ & CHỨNG NHẬN (25pt - 31pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Tiêu Chí Đánh Giá Dự Án STEM (Rubric) 🏆", "ĐÁNH GIÁ TỔNG KẾT", 18, title_size=31)

    create_card(s, 0.6, 1.62, 6.1, 5.55, bg_color=C_WHITE, border_color=BORDER_AMBER)
    tf = make_textbox(s, 0.84, 1.82, 5.62, 5.15)
    add_para(tf, [("🎯 THANG ĐIỂM ĐÁNH GIÁ (10Đ)", 26, True, T_AMBER)], first=True)
    add_para(tf, [("- Tiêu chí 1 (40%): Chọn đúng 4/4 giải pháp và ghi bảng số liệu chính xác.", 25, False, T_BODY)], space_before=16)
    add_para(tf, [("- Tiêu chí 2 (30%): Vận dụng kiến thức Bài 23 giải thích rõ cơ chế.", 25, False, T_BODY)], space_before=14)
    add_para(tf, [("- Tiêu chí 3 (30%): Phối hợp nhóm tốt, báo cáo rõ ràng, tự tin.", 25, False, T_BODY)], space_before=14)

    add_picture_framed(s, "screenshots/screen_certificate.png", 6.9, 1.62, 5.833, 5.55, BORDER_AMBER)

    # ==========================================================
    # SLIDE 19: HƯỚNG DẪN TỰ HỌC Ở NHÀ (SPARSE -> 26pt - 34pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_SOFT_MINT)
    add_header(s, "Hướng Dẫn Tự Học & Mở Rộng Tại Nhà 🏡🌱", "HOẠT ĐỘNG TIẾP NỐI", 19, title_size=33)

    create_card(s, 0.6, 1.62, 12.133, 5.55, bg_color=C_WHITE, border_color=BORDER_MINT)
    tf = make_textbox(s, 0.95, 1.90, 11.433, 5.0)
    add_para(tf, [("1. Hệ thống kiến thức: ", 28, True, T_GREEN), ("Vẽ sơ đồ tư duy 4 yếu tố ảnh hưởng đến quang hợp vào vở ghi.", 27, False, T_BODY)], first=True)
    add_para(tf, [("2. Quan sát góc xanh gia đình: ", 28, True, T_BLUE), ("Khảo sát ánh sáng tại ban công hoặc bàn học ở nhà, đề xuất 01 loài cây phù hợp để trồng và chăm sóc.", 27, False, T_BODY)], space_before=24)
    add_para(tf, [("3. Chuẩn bị bài mới: ", 28, True, T_AMBER), ("Xem trước Bài 24 – Thực hành: Chứng minh quang hợp ở cây xanh.", 27, False, T_BODY)], space_before=24)

    # ==========================================================
    # SLIDE 20: TRANG KẾT THÚC (SPARSE -> 26pt - 38pt)
    # ==========================================================
    s = prs.slides.add_slide(blank_layout)
    set_slide_bg(s, BG_TITLE)
    create_card(s, 0.6, 0.5, 12.133, 6.5, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=2.5)

    if os.path.exists("images/logo.png"):
        s.shapes.add_picture("images/logo.png", Inches(6.016), Inches(0.85), width=Inches(1.3), height=Inches(1.3))

    tf_end = make_textbox(s, 1.0, 2.35, 11.333, 4.3)
    add_para(tf_end, [("🌿 CẢM ƠN QUÝ THẦY CÔ\nVÀ CÁC EM HỌC SINH! 🌿", 38, True, T_PRIMARY)], align=PP_ALIGN.CENTER, first=True)
    add_para(tf_end, [("Chúc các em học tốt và tích cực chăm sóc mảng xanh quanh mình!", 27, False, T_GREEN)], space_before=18, align=PP_ALIGN.CENTER)
    add_para(tf_end, [("GV: Mai Thuý Đức  •  Trường THCS Tân Tạo A\nHọc liệu số: https://sinhhoc-amber.vercel.app/", 25, True, T_BODY)], space_before=22, align=PP_ALIGN.CENTER)

    prs.core_properties.author = "Mai Thuý Đức"
    prs.core_properties.last_modified_by = "Mai Thuý Đức"
    prs.core_properties.title = "Bài giảng KHTN 7 - Bài 23: Một số yếu tố ảnh hưởng đến quang hợp"
    prs.core_properties.comments = ""

    prs.save("KHTN7_Bai23_QuangHop_MaiThuyDuc.pptx")
    prs.save("03_Slide_Trinh_Chieu_Giang_Day_Mai_Thuy_Duc.pptx")
    print(f"Saved {len(prs.slides)}-slide deck (>=24pt) to both pptx files!")


if __name__ == "__main__":
    create_cute_deck()
