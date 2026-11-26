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

    # --- CUTE & GENTLE PASTEL PALETTE ---
    # Backgrounds
    BG_SOFT_MINT = RGBColor(243, 249, 245)    # #f3f9f5 Milk mint
    BG_CREAM     = RGBColor(254, 252, 248)    # #fefcf8 Warm milk cream
    BG_TITLE     = RGBColor(234, 247, 239)    # #eaf7ef Soft pastel green

    # Cards & Fills
    C_WHITE      = RGBColor(255, 255, 255)
    C_PILL_MINT  = RGBColor(209, 250, 229)    # #d1fae5
    C_PILL_SKY   = RGBColor(224, 242, 254)    # #e0f2fe
    C_PILL_PEACH = RGBColor(254, 243, 199)    # #fef3c7
    C_PILL_ROSE  = RGBColor(255, 228, 230)    # #ffe4e6
    C_PILL_PURPLE= RGBColor(243, 232, 255)    # #f3e8ff

    # Borders
    BORDER_MINT  = RGBColor(167, 243, 208)    # #a7f3d0
    BORDER_SKY   = RGBColor(186, 230, 253)    # #bae6fd
    BORDER_AMBER = RGBColor(253, 230, 138)    # #fde68a
    BORDER_ROSE  = RGBColor(254, 205, 211)    # #fecdd3
    BORDER_SOFT  = RGBColor(229, 231, 235)

    # Texts
    T_PRIMARY    = RGBColor(27, 67, 50)       # #1b4332 Deep matcha/forest
    T_GREEN      = RGBColor(16, 122, 80)      # #107a50 Fresh botanical
    T_BLUE       = RGBColor(2, 132, 199)      # #0284c7 Sky blue
    T_AMBER      = RGBColor(180, 83, 9)       # #b45309 Honey amber
    T_ROSE       = RGBColor(190, 24, 93)      # #be185d Berry rose
    T_BODY       = RGBColor(51, 65, 85)       # #334155 Slate body
    T_MUTED      = RGBColor(100, 116, 139)    # #64748b Muted

    def set_slide_bg(slide, color):
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = color

    def add_header(slide, title_text, pill_text="KHTN 7 • BÀI 23", pill_bg=C_PILL_MINT, pill_border=BORDER_MINT, pill_color=T_PRIMARY):
        # Category pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.4), Inches(4.8), Inches(0.4))
        pill.fill.solid()
        pill.fill.fore_color.rgb = pill_bg
        pill.line.color.rgb = pill_border
        pill.line.width = Pt(1.2)
        tf = pill.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = pill_text
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = pill_color
        r.font.name = "Arial"

        # Main Title (Big, friendly, legible)
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.85), Inches(11.733), Inches(0.65))
        tf = tb_title.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = title_text
        r.font.bold = True
        r.font.size = Pt(21)
        r.font.color.rgb = T_PRIMARY
        r.font.name = "Arial"

        # Decorative soft dotted line
        line = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.52), Inches(11.733), Inches(0.03))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(209, 235, 220)
        line.line.fill.background()

    def add_footer(slide, slide_num, total_slides=18):
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.35))
        tf = tb.text_frame
        p = tf.paragraphs[0]
        
        r1 = p.add_run()
        r1.text = "🏫 Trường THCS Tân Tạo A  •  👩‍🏫 Cô giáo: Mai Thuý Đức  •  🌱 Học liệu số: sinhhoc-amber.vercel.app"
        r1.font.size = Pt(9.5)
        r1.font.color.rgb = T_MUTED
        r1.font.name = "Arial"

        r2 = p.add_run()
        r2.text = f"                 Trang {slide_num}/{total_slides} ✨"
        r2.font.bold = True
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_GREEN
        r2.font.name = "Arial"

    def create_pill(slide, left, top, width, height, text, bg_color, border_color=None, text_color=T_PRIMARY, font_size=10, bold=True):
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        pill.fill.solid()
        pill.fill.fore_color.rgb = bg_color
        if border_color:
            pill.line.color.rgb = border_color
            pill.line.width = Pt(1.2)
        else:
            pill.line.fill.background()
        tf = pill.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = text
        r.font.bold = bold
        r.font.size = Pt(font_size)
        r.font.color.rgb = text_color
        r.font.name = "Arial"
        return pill

    def create_card(slide, left, top, width, height, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5):
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
    # SLIDE 1: TRANG BÌA SIÊU DỄ THƯƠNG (TITLE SLIDE)
    # ==========================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s1, BG_TITLE)

    # Large Soft Card in center
    c_main = create_card(s1, 0.8, 0.7, 11.733, 6.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=2)

    # Cute Logo Badge
    if os.path.exists("images/logo.png"):
        s1.shapes.add_picture("images/logo.png", Inches(1.3), Inches(1.15), width=Inches(1.0), height=Inches(1.0))

    # Top Welcome Badge
    tb_badge = s1.shapes.add_textbox(Inches(2.5), Inches(1.2), Inches(9.5), Inches(0.45))
    tf = tb_badge.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "✨ CHÀO MỪNG CÁC EM ĐẾN VỚI LỚP HỌC KHOA HỌC TỰ NHIÊN 7 ✨"
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = T_GREEN
    r.font.name = "Arial"
    p2 = tf.add_paragraph()
    r2 = p2.add_run()
    r2.text = "Bộ sách Kết nối tri thức với cuộc sống  •  Chủ đề STEM & Chuyển đổi số sáng tạo"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = T_MUTED

    # Title
    tb_title = s1.shapes.add_textbox(Inches(1.3), Inches(2.4), Inches(10.7), Inches(1.9))
    tf = tb_title.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌿 BÀI 23: MỘT SỐ YẾU TỐ ẢNH HƯỞNG\nĐẾN QUANG HỢP 🌿"
    r.font.bold = True
    r.font.size = Pt(30)
    r.font.color.rgb = T_PRIMARY
    r.font.name = "Arial"

    p_sub = tf.add_paragraph()
    p_sub.space_before = Pt(10)
    r_sub = p_sub.add_run()
    r_sub.text = "🌱 Cùng Đi Tìm Bí Mật Giúp Cây Xanh Luôn Tươi Vui & Kiến Tạo Không Gian Xanh Đô Thị (90 Phút)"
    r_sub.font.bold = True
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = T_GREEN

    # Teacher Pill Card
    c_info = create_card(s1, 1.3, 4.65, 10.7, 1.7, bg_color=C_PILL_MINT, border_color=BORDER_MINT, border_width=1.5)
    tb_info = s1.shapes.add_textbox(Inches(1.5), Inches(4.75), Inches(10.3), Inches(1.5))
    tf = tb_info.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌸 THÔNG TIN LỚP HỌC:"
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = T_PRIMARY

    p_gv = tf.add_paragraph()
    p_gv.space_before = Pt(5)
    r = p_gv.add_run()
    r.text = "👩‍🏫 Cô giáo hướng dẫn:  "
    r.font.size = Pt(12)
    r.font.color.rgb = T_BODY
    r_name = p_gv.add_run()
    r_name.text = "MAI THUÝ ĐỨC"
    r_name.font.bold = True
    r_name.font.size = Pt(13)
    r_name.font.color.rgb = T_PRIMARY

    r_sch = p_gv.add_run()
    r_sch.text = "            🏫 Trường:  "
    r_sch.font.size = Pt(12)
    r_sch.font.color.rgb = T_BODY
    r_sname = p_gv.add_run()
    r_sname.text = "THCS TÂN TẠO A (QUẬN BÌNH TÂN)"
    r_sname.font.bold = True
    r_sname.font.size = Pt(12.5)
    r_sname.font.color.rgb = T_PRIMARY

    p_web = tf.add_paragraph()
    p_web.space_before = Pt(4)
    r = p_web.add_run()
    r.text = "📞 SĐT: 0975515659  •  ✉️ Email: maithuyduc36@gmail.com  •  🌐 Học liệu số: "
    r.font.size = Pt(10.5)
    r.font.color.rgb = T_BODY
    r_link = p_web.add_run()
    r_link.text = "https://sinhhoc-amber.vercel.app/"
    r_link.font.bold = True
    r_link.font.size = Pt(10.5)
    r_link.font.color.rgb = T_GREEN

    # ==========================================================
    # SLIDE 2: MỤC TIÊU BÀI HỌC (NHÀ THÁM HIỂM NHÍ)
    # ==========================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s2, BG_SOFT_MINT)
    add_header(s2, "Hôm Nay Chúng Mình Cùng Khám Phá Điều Gì Nhé? 🎯", "MỤC TIÊU BÀI HỌC")
    add_footer(s2, 2, 18)

    # 3 Pastel Cards
    # Card 1: KHTN
    c1 = create_card(s2, 0.8, 1.7, 3.7, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s2.shapes.add_textbox(Inches(0.95), Inches(1.85), Inches(3.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌱 1. BÍ MẬT KHOA HỌC"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = T_GREEN

    p = tf.add_paragraph()
    p.space_before = Pt(8)
    r = p.add_run()
    r.text = "Gặp gỡ 4 'người bạn' thân thiết của cây xanh:\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = T_BODY
    r2 = p.add_run()
    r2.text = "• ☀️ Ánh Nắng Ấm Áp\n• 💧 Giọt Nước Mát Lành\n• ☁️ Khí CO₂ Trong Lành\n• 🌡️ Nhiệt Độ Dịu Êm\n\n"
    r2.font.bold = True
    r2.font.size = Pt(10.5)
    r2.font.color.rgb = T_PRIMARY

    r3 = p.add_run()
    r3.text = "🔍 Phân biệt được cây thích tắm nắng (ưa sáng) và cây thích núp bóng mát (ưa bóng) để chăm sóc cây thật chuẩn!"
    r3.font.size = Pt(10)
    r3.font.color.rgb = T_BODY

    # Card 2: Năng lực số
    c2 = create_card(s2, 4.8, 1.7, 3.7, 5.1, bg_color=C_WHITE, border_color=BORDER_SKY, border_width=1.5)
    tb = s2.shapes.add_textbox(Inches(4.95), Inches(1.85), Inches(3.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "💻 2. TRẢI NGHIỆM CÔNG NGHỆ"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = T_BLUE

    p = tf.add_paragraph()
    p.space_before = Pt(8)
    r = p.add_run()
    r.text = "Khám phá phòng thí nghiệm ảo như chơi game sinh thái:\n\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = T_BODY
    r2 = p.add_run()
    r2.text = "• 🧪 Đếm từng bọt khí O₂ cành rong đuôi chó thở ra.\n• 🖱️ Kéo thả chậu cây giống vào 4 tầng tháp xanh.\n• 📊 Đọc bảng đồng hồ đo nhiệt độ, ánh sáng thời gian thực.\n• 🎙️ Trình diễn E-Poster khoa học rực rỡ trước lớp!"
    r2.font.size = Pt(10)
    r2.font.color.rgb = T_BODY

    # Card 3: Phẩm chất
    c3 = create_card(s2, 8.8, 1.7, 3.7, 5.1, bg_color=C_WHITE, border_color=BORDER_AMBER, border_width=1.5)
    tb = s2.shapes.add_textbox(Inches(8.95), Inches(1.85), Inches(3.4), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "💖 3. YÊU QUÝ THIÊN NHIÊN"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = T_AMBER

    p = tf.add_paragraph()
    p.space_before = Pt(8)
    r = p.add_run()
    r.text = "Trở thành Đại Sứ Xanh nhí của TP.HCM:\n\n"
    r.font.size = Pt(10.5)
    r.font.color.rgb = T_BODY
    r2 = p.add_run()
    r2.text = "• 🤝 Làm việc nhóm nhịp nhàng, biết lắng nghe và cùng bạn giải quyết thử thách.\n• 🌿 Yêu quý cây xanh, hiểu rằng lá cây đang ngày đêm lọc sạch không khí cho chúng ta thở.\n• 🏡 Tự tin tự tay trồng một góc cây xanh mát ở nhà và lớp học!"
    r2.font.size = Pt(10)
    r2.font.color.rgb = T_BODY

    # ==========================================================
    # SLIDE 3: BẢN ĐỒ PHIÊU LƯU 90 PHÚT
    # ==========================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s3, BG_SOFT_MINT)
    add_header(s3, "Bản Đồ Cuộc Phiêu Lưu Xanh 90 Phút Cùng Cô Thúy Đức 🗺️", "TIẾN TRÌNH 2 TIẾT HỌC")
    add_footer(s3, 3, 18)

    # 2 Big Columns
    c_t1 = create_card(s3, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s3.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌟 TIẾT 1 (45 PHÚT): ĐI TÌM CÔNG THỨC QUANG HỢP"
    r.font.bold = True
    r.font.size = Pt(12.5)
    r.font.color.rgb = T_GREEN

    t1_cute = [
        ("Trạm 1: Khởi động tò mò (5p)", "Nghe câu chuyện của hai bạn cây: Bạn Cây Cháy Nắng và Bạn Cây Trong Tối. Tìm ra câu hỏi bí ẩn cần giải đáp!"),
        ("Trạm 2: Vào phòng thí nghiệm ảo (35p)", "Hóa thân thành nhà khoa học nhí:\n• Thử ánh nắng: Xem bọt khí O₂ cành rong nổi lên rộn ràng.\n• Thử giọt nước: Soi 'chiếc miệng' khí khổng đóng mở.\n• Thử nồng độ CO₂ & nhiệt độ: Khám phá ngưỡng chịu đựng của lá."),
        ("Trạm 3: Hoàn thành Phiếu Thần Kỳ (5p)", "Cùng chốt lại các mốc bí kíp quan trọng vào Phiếu học tập số 1 để bước sang Tiết 2!")
    ]
    for st_title, st_desc in t1_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"📍 {st_title}\n"
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = T_GREEN
        r2 = p.add_run()
        r2.text = st_desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    c_t2 = create_card(s3, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_AMBER, border_width=1.5)
    tb = s3.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🚀 TIẾT 2 (45 PHÚT): TRỞ THÀNH KỸ SƯ SINH THÁI NHÍ"
    r.font.bold = True
    r.font.size = Pt(12.5)
    r.font.color.rgb = T_AMBER

    t2_cute = [
        ("Trạm 4: Thử tài đố vui nhanh (10p)", "Cả lớp cùng thi đấu trắc nghiệm 4 câu hỏi vui nhộn trên máy tính để 'nạp năng lượng' kiến thức."),
        ("Trạm 5: Trồng rừng xanh 4 Trạm TP.HCM (25p)", "Mỗi nhóm nhận 1 vùng vi khí hậu:\n• Kéo cây ưa sáng lên nóc tháp, đưa cây ưa bóng xuống trệt.\n• Bấm nút tưới nước, che lưới, thổi gió mát lành!"),
        ("Trạm 6: Khoe Poster & Nhận Giấy Khen (10p)", "Chiếu E-Poster lung linh, Cô Mai Thuý Đức duyệt bài và trao Giấy chứng nhận Kỹ sư sinh học có mã hash xịn xò!")
    ]
    for st_title, st_desc in t2_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"📍 {st_title}\n"
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = T_AMBER
        r2 = p.add_run()
        r2.text = st_desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # ==========================================================
    # SLIDE 4: KHỞI ĐỘNG - CÂU CHUYỆN CỦA HAI BẠN CÂY (3-CỘT CHUẨN GIAO DIỆN WEB)
    # ==========================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s4, BG_SOFT_MINT)
    add_footer(s4, 4, 18)

    # Header: Pill, Title & Subtitle y hệt Web
    create_pill(s4, 0.8, 0.35, 4.3, 0.35, "🌱 KHỞI ĐỘNG BÀI HỌC • TÌNH HUỐNG THỰC TIỄN", C_PILL_MINT, BORDER_MINT, RGBColor(6, 95, 70), font_size=10)
    create_pill(s4, 9.2, 0.35, 3.333, 0.35, "🔊 Giọng Đọc AI & Tương Tác Số 🌿", C_PILL_PEACH, BORDER_AMBER, RGBColor(180, 83, 9), font_size=9.5)

    tb_h4 = s4.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.733), Inches(0.78))
    tf = tb_h4.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = 'Câu Chuyện Của Hai Bạn Cây: "Sao Bạn Lại Buồn?" 🧐'
    r.font.bold = True
    r.font.size = Pt(20)
    r.font.color.rgb = T_PRIMARY
    r.font.name = "Arial"

    p_sub = tf.add_paragraph()
    p_sub.space_before = Pt(3)
    r_sub = p_sub.add_run()
    r_sub.text = "Lắng nghe cuộc đối thoại giữa Bạn Cây Sân Thượng (thừa nắng cháy lá) và Bạn Cây Phòng Khách (thiếu sáng đói lả) để tìm ra câu hỏi nghiên cứu trọng tâm!"
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = T_MUTED
    r_sub.font.name = "Arial"

    # Line divider
    line = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.52), Inches(11.733), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(209, 235, 220)
    line.line.fill.background()

    # --- CỘT TRÁI: BÉ LAN Ý (SÂN THƯỢNG NẮNG GẮT) ---
    c_left = create_card(s4, 0.8, 1.62, 3.25, 4.25, bg_color=C_WHITE, border_color=RGBColor(251, 191, 36), border_width=1.8)
    create_pill(s4, 1.15, 1.76, 2.55, 0.32, "☀️ SÂN THƯỢNG NẮNG GẮT", RGBColor(254, 243, 199), RGBColor(253, 230, 138), RGBColor(180, 83, 9), font_size=9.5)
    
    if os.path.exists("images/cay_san_thuong.png"):
        s4.shapes.add_picture("images/cay_san_thuong.png", Inches(1.47), Inches(2.18), width=Inches(1.9), height=Inches(1.8))

    tb = s4.shapes.add_textbox(Inches(0.9), Inches(4.02), Inches(3.05), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Bé Lan Ý (Ưa Bóng)"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_PRIMARY
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "\"Nóng rát phiến lá quá!\""
    r2.font.italic = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(220, 38, 38)

    # Box thông số vi khí hậu
    c_param_l = create_card(s4, 0.95, 4.68, 2.95, 1.05, bg_color=RGBColor(255, 251, 235), border_color=RGBColor(254, 215, 170), border_width=1.0)
    tb_pl = s4.shapes.add_textbox(Inches(1.05), Inches(4.72), Inches(2.75), Inches(0.95))
    tf = tb_pl.text_frame
    tf.word_wrap = True
    params_l = [
        ("Ánh sáng: ", "48.000 Lux ⚠️", RGBColor(220, 38, 38)),
        ("Nhiệt độ: ", "39.0°C ⚠️", RGBColor(220, 38, 38)),
        ("Tưới nước: ", "Đầy đủ (100%)", RGBColor(16, 122, 80))
    ]
    for idx, (lbl, val, col) in enumerate(params_l):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        if idx > 0: p.space_before = Pt(2)
        r = p.add_run()
        r.text = lbl
        r.font.size = Pt(9)
        r.font.color.rgb = T_MUTED
        r2 = p.add_run()
        r2.text = val
        r2.font.bold = True
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = col

    # --- CỘT GIỮA: KHUNG ĐỐI THOẠI CỦA HAI BẠN CÂY ---
    c_mid = create_card(s4, 4.25, 1.62, 4.833, 4.25, bg_color=C_WHITE, border_color=RGBColor(56, 189, 248), border_width=1.8)
    create_pill(s4, 4.5, 1.76, 2.5, 0.32, "🟣 Bé Bàng Cảnh  •  4 / 6", RGBColor(243, 232, 255), RGBColor(216, 180, 254), RGBColor(126, 34, 206), font_size=9.5)

    tb_dialogue = s4.shapes.add_textbox(Inches(4.45), Inches(2.18), Inches(4.433), Inches(1.65))
    tf = tb_dialogue.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "“Đúng đấy! Hình như tớ là cây ưa sáng phiến lá dày, cần nhiều nắng; còn cậu là cây ưa bóng phiến lá mỏng, chỉ thích ánh sáng dịu mát thôi!"
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(31, 41, 55)
    r.font.name = "Arial"
    p2 = tf.add_paragraph()
    p2.space_before = Pt(6)
    r2 = p2.add_run()
    r2.text = "Lại còn nhiệt độ nữa, trên 40°C thì các chú thợ enzyme quang hợp bị đình công hết!”"
    r2.font.size = Pt(11)
    r2.font.color.rgb = RGBColor(31, 41, 55)
    r2.font.name = "Arial"
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(31, 41, 55)
    r.font.name = "Arial"

    # Pagination indicator
    tb_dots = s4.shapes.add_textbox(Inches(4.45), Inches(3.92), Inches(4.433), Inches(0.3))
    tf = tb_dots.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "‹ Quay lại          ● ● ● ● ○ ○          Tiếp theo ›"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(156, 163, 175)

    # Mini Swap Prompt Card
    c_swap = create_card(s4, 4.4, 4.30, 4.533, 1.43, bg_color=RGBColor(240, 253, 244), border_color=RGBColor(187, 247, 208), border_width=1.2)
    tb_sw = s4.shapes.add_textbox(Inches(4.5), Inches(4.35), Inches(4.333), Inches(1.3))
    tf = tb_sw.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🔄 Thử nghiệm nhanh: Hoán đổi vị trí của 2 chậu cây!"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(22, 101, 52)
    
    p2 = tf.add_paragraph()
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "Nếu đưa Bé Lan Ý vào bóng râm và đưa Bé Bàng Cảnh ra đón nắng thì điều gì sẽ xảy ra? Liệu hai bạn có hồi phục không?"
    r2.font.italic = True
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(55, 65, 81)

    p3 = tf.add_paragraph()
    p3.space_before = Pt(4)
    r3 = p3.add_run()
    r3.text = "👉 [ 🔄 Bấm Thử Hoán Đổi Vị Trí Ngay Trên Web ]"
    r3.font.bold = True
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(5, 150, 105)

    # --- CỘT PHẢI: BÉ BÀNG CẢNH (PHÒNG KHÁCH TỐI MỜ) ---
    c_right = create_card(s4, 9.283, 1.62, 3.25, 4.25, bg_color=C_WHITE, border_color=RGBColor(6, 182, 212), border_width=1.8)
    create_pill(s4, 9.633, 1.76, 2.55, 0.32, "🔵 PHÒNG KHÁCH TỐI MỜ", RGBColor(224, 242, 254), RGBColor(186, 230, 253), RGBColor(2, 132, 199), font_size=9.5)

    if os.path.exists("images/cay_phong_khach.png"):
        s4.shapes.add_picture("images/cay_phong_khach.png", Inches(9.953), Inches(2.18), width=Inches(1.9), height=Inches(1.8))

    tb = s4.shapes.add_textbox(Inches(9.383), Inches(4.02), Inches(3.05), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Bé Bàng Cảnh (Ưa Sáng)"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_PRIMARY
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "\"Tối mờ, đói lả cả người!\""
    r2.font.italic = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(2, 132, 199)

    # Box thông số vi khí hậu
    c_param_r = create_card(s4, 9.433, 4.68, 2.95, 1.05, bg_color=RGBColor(240, 249, 255), border_color=RGBColor(186, 230, 253), border_width=1.0)
    tb_pr = s4.shapes.add_textbox(Inches(9.533), Inches(4.72), Inches(2.75), Inches(0.95))
    tf = tb_pr.text_frame
    tf.word_wrap = True
    params_r = [
        ("Ánh sáng: ", "650 Lux ⚠️", RGBColor(180, 83, 9)),
        ("Nhiệt độ: ", "25.0°C (Mát)", RGBColor(2, 132, 199)),
        ("Tưới nước: ", "Đầy đủ (100%)", RGBColor(16, 122, 80))
    ]
    for idx, (lbl, val, col) in enumerate(params_r):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        if idx > 0: p.space_before = Pt(2)
        r = p.add_run()
        r.text = lbl
        r.font.size = Pt(9)
        r.font.color.rgb = T_MUTED
        r2 = p.add_run()
        r2.text = val
        r2.font.bold = True
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = col

    # --- BANNER CHÂN TRANG: NHIỆM VỤ HỌC TẬP TRỌNG TÂM ---
    c_bot = create_card(s4, 0.8, 6.00, 11.733, 0.95, bg_color=RGBColor(6, 78, 59), border_color=RGBColor(16, 185, 129), border_width=1.5)
    tb_bot = s4.shapes.add_textbox(Inches(1.05), Inches(6.08), Inches(7.6), Inches(0.8))
    tf = tb_bot.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🚀 Nhiệm Vụ Học Tập Trọng Tâm: Giải Cứu Cây Xanh!"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(255, 255, 255)
    
    p2 = tf.add_paragraph()
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "Cùng vào Phòng Thí Nghiệm Số (Module 1) để điều tra ngưỡng tối ưu của Ánh Sáng, Nước, CO₂ và Nhiệt Độ."
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(167, 243, 208)

    # Nút bấm bên phải
    btn_bot = create_card(s4, 8.85, 6.12, 3.5, 0.7, bg_color=RGBColor(255, 255, 255), border_color=RGBColor(52, 211, 153), border_width=1.2)
    tb_btn = s4.shapes.add_textbox(Inches(8.9), Inches(6.12), Inches(3.4), Inches(0.7))
    tf = tb_btn.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "🧪 Vào Phòng Thí Nghiệm Số Ngay ➔"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(6, 78, 59)
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(2)
    r2 = p2.add_run()
    r2.text = "sinhhoc-amber.vercel.app"
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(5, 150, 105)

    # ==========================================================
    # SLIDE 6: KHỞI ĐỘNG - KẾT QUẢ HOÁN ĐỔI VỊ TRÍ (GIẢI CỨU THÀNH CÔNG!)
    # ==========================================================
    s5_res = prs.slides.add_slide(blank_layout)
    set_slide_bg(s5_res, BG_SOFT_MINT)
    add_footer(s5_res, 5, 18)

    create_pill(s5_res, 0.8, 0.35, 4.3, 0.35, "🌱 KHỞI ĐỘNG BÀI HỌC • KẾT QUẢ THỬ NGHIỆM SỐ", C_PILL_MINT, BORDER_MINT, RGBColor(6, 95, 70), font_size=10)
    create_pill(s5_res, 9.2, 0.35, 3.333, 0.35, "✨ Giải Cứu Cây Trồng Thành Công!", C_PILL_MINT, BORDER_MINT, RGBColor(6, 95, 70), font_size=9.5)

    tb_h5 = s5_res.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.733), Inches(0.78))
    tf = tb_h5.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Khởi Động: Thử Nghiệm Hoán Đổi Vị Trí — Hai Bạn Cây Cùng Tươi Vui! 🎉"
    r.font.bold = True
    r.font.size = Pt(19)
    r.font.color.rgb = T_PRIMARY
    r.font.name = "Arial"

    p_sub = tf.add_paragraph()
    p_sub.space_before = Pt(3)
    r_sub = p_sub.add_run()
    r_sub.text = "Khi đưa Bé Lan Ý vào phòng khách mát mẻ và đưa Bé Bàng Cảnh ra ban công đón nắng: Cả hai bạn cây đều hồi sinh rực rỡ!"
    r_sub.font.size = Pt(10)
    r_sub.font.color.rgb = T_MUTED
    r_sub.font.name = "Arial"

    line = s5_res.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.52), Inches(11.733), Inches(0.02))
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(209, 235, 220)
    line.line.fill.background()

    # --- CỘT TRÁI: LAN Ý HỒI SINH TRONG PHÒNG KHÁCH ---
    c_l5 = create_card(s5_res, 0.8, 1.62, 3.25, 4.25, bg_color=C_WHITE, border_color=RGBColor(52, 211, 153), border_width=1.8)
    create_pill(s5_res, 1.05, 1.76, 2.75, 0.32, "🏡 PHÒNG KHÁCH MÁT (VỊ TRÍ MỚI)", RGBColor(209, 250, 229), RGBColor(167, 243, 208), RGBColor(6, 95, 70), font_size=9)

    if os.path.exists("images/cay_san_thuong.png"):
        s5_res.shapes.add_picture("images/cay_san_thuong.png", Inches(1.47), Inches(2.18), width=Inches(1.9), height=Inches(1.8))

    tb = s5_res.shapes.add_textbox(Inches(0.9), Inches(4.02), Inches(3.05), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Bé Lan Ý (Hồi Sinh Tươi Tốt! 🌸)"
    r.font.bold = True
    r.font.size = Pt(11.5)
    r.font.color.rgb = RGBColor(6, 95, 70)
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "\"Mát dịu quá, mép lá tớ đã lành rồi!\""
    r2.font.italic = True
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(16, 185, 129)

    c_param_l5 = create_card(s5_res, 0.95, 4.68, 2.95, 1.05, bg_color=RGBColor(240, 253, 244), border_color=RGBColor(187, 247, 208), border_width=1.0)
    tb_pl5 = s5_res.shapes.add_textbox(Inches(1.05), Inches(4.72), Inches(2.75), Inches(0.95))
    tf = tb_pl5.text_frame
    tf.word_wrap = True
    params_l5 = [
        ("Ánh sáng: ", "650 Lux (Tán xạ dịu) ✅", RGBColor(16, 122, 80)),
        ("Nhiệt độ: ", "25.0°C (Mát mẻ) ✅", RGBColor(16, 122, 80)),
        ("Quang hợp: ", "Ổn định, tạo tinh bột ✅", RGBColor(6, 95, 70))
    ]
    for idx, (lbl, val, col) in enumerate(params_l5):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        if idx > 0: p.space_before = Pt(2)
        r = p.add_run()
        r.text = lbl
        r.font.size = Pt(9)
        r.font.color.rgb = T_MUTED
        r2 = p.add_run()
        r2.text = val
        r2.font.bold = True
        r2.font.size = Pt(9)
        r2.font.color.rgb = col

    # --- CỘT GIỮA: ĐÚC KẾT KHOA HỌC & CÂU HỎI MỞ BÀI ---
    c_m5 = create_card(s5_res, 4.25, 1.62, 4.833, 4.25, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.8)
    create_pill(s5_res, 4.5, 1.76, 3.2, 0.32, "💡 BẬT MÍ KHOA HỌC TỪ THỬ NGHIỆM", RGBColor(254, 243, 199), RGBColor(253, 230, 138), RGBColor(180, 83, 9), font_size=9.5)

    tb_disc = s5_res.shapes.add_textbox(Inches(4.45), Inches(2.18), Inches(4.433), Inches(2.15))
    tf = tb_disc.text_frame
    tf.word_wrap = True
    points = [
        ("1. Nhu cầu Ánh sáng khác biệt: ", "Cây ưa sáng cần quang năng mạnh, phiến lá dày. Cây ưa bóng chỉ cần ánh sáng tán xạ dịu nhẹ, nếu phơi nắng gắt sẽ bị cháy lá."),
        ("2. Giới hạn Nhiệt độ: ", "Khi trời quá nóng (>40°C), enzyme quang hợp bị đình công; khí khổng đóng chặt để giữ nước khiến cây ngừng quang hợp."),
        ("3. Bí mật cần giải mã: ", "Còn NƯỚC 💧 và KHÍ CO₂ ☁️ ảnh hưởng như thế nào? Cây cần bao nhiêu thì đủ để năng suất quang hợp đạt đỉnh cao nhất?")
    ]
    for idx, (head, body) in enumerate(points):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        if idx > 0: p.space_before = Pt(6)
        r = p.add_run()
        r.text = head
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_PRIMARY
        r2 = p.add_run()
        r2.text = body
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # Mini Call to action
    c_cta5 = create_card(s5_res, 4.4, 4.48, 4.533, 1.25, bg_color=RGBColor(234, 247, 239), border_color=BORDER_MINT, border_width=1.2)
    tb_cta5 = s5_res.shapes.add_textbox(Inches(4.5), Inches(4.52), Inches(4.333), Inches(1.15))
    tf = tb_cta5.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🎯 CÙNG BƯỚC VÀO PHÒNG THÍ NGHIỆM SỐ (MODULE 1)!"
    r.font.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = T_GREEN
    p2 = tf.add_paragraph()
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "Các em hãy tự tay điều chỉnh 4 yếu tố trên mô hình số để quan sát bọt khí O₂ cành rong và giải mã toàn bộ ngưỡng tối ưu!"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = T_BODY

    # --- CỘT PHẢI: BÀNG CẢNH TẮM NẮNG QUANG HỢP ---
    c_r5 = create_card(s5_res, 9.283, 1.62, 3.25, 4.25, bg_color=C_WHITE, border_color=RGBColor(251, 191, 36), border_width=1.8)
    create_pill(s5_res, 9.533, 1.76, 2.75, 0.32, "☀️ BAN CÔNG ĐÓN NẮNG (VỊ TRÍ MỚI)", RGBColor(254, 243, 199), RGBColor(253, 230, 138), RGBColor(180, 83, 9), font_size=9)

    if os.path.exists("images/cay_phong_khach.png"):
        s5_res.shapes.add_picture("images/cay_phong_khach.png", Inches(9.953), Inches(2.18), width=Inches(1.9), height=Inches(1.8))

    tb = s5_res.shapes.add_textbox(Inches(9.383), Inches(4.02), Inches(3.05), Inches(0.65))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "Bé Bàng Cảnh (Quang Hợp Hăng Say! ⚡)"
    r.font.bold = True
    r.font.size = Pt(11.5)
    r.font.color.rgb = T_PRIMARY
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "\"Nắng ngập tràn rồi, tớ no nê tinh bột!\""
    r2.font.italic = True
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(180, 83, 9)

    c_param_r5 = create_card(s5_res, 9.433, 4.68, 2.95, 1.05, bg_color=RGBColor(255, 251, 235), border_color=RGBColor(254, 215, 170), border_width=1.0)
    tb_pr5 = s5_res.shapes.add_textbox(Inches(9.533), Inches(4.72), Inches(2.75), Inches(0.95))
    tf = tb_pr5.text_frame
    tf.word_wrap = True
    params_r5 = [
        ("Ánh sáng: ", "45.000 Lux (Nắng tốt) ✅", RGBColor(180, 83, 9)),
        ("Nhiệt độ: ", "30.0°C (Gió thoáng) ✅", RGBColor(16, 122, 80)),
        ("Quang hợp: ", "Bọt O₂ thoát mạnh mẽ! ✅", RGBColor(6, 95, 70))
    ]
    for idx, (lbl, val, col) in enumerate(params_r5):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        if idx > 0: p.space_before = Pt(2)
        r = p.add_run()
        r.text = lbl
        r.font.size = Pt(9)
        r.font.color.rgb = T_MUTED
        r2 = p.add_run()
        r2.text = val
        r2.font.bold = True
        r2.font.size = Pt(9)
        r2.font.color.rgb = col

    # --- BANNER CHÂN TRANG: CHUYỂN TIẾP HOẠT ĐỘNG 2 ---
    c_bot5 = create_card(s5_res, 0.8, 6.00, 11.733, 0.95, bg_color=RGBColor(6, 78, 59), border_color=RGBColor(16, 185, 129), border_width=1.5)
    tb_bot5 = s5_res.shapes.add_textbox(Inches(1.05), Inches(6.08), Inches(7.6), Inches(0.8))
    tf = tb_bot5.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌿 BẮT ĐẦU HOẠT ĐỘNG 2: HÌNH THÀNH KIẾN THỨC MỚI (35 PHÚT)"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(255, 255, 255)
    
    p2 = tf.add_paragraph()
    p2.space_before = Pt(3)
    r2 = p2.add_run()
    r2.text = "Khám phá 4 yếu tố quang hợp qua Thí nghiệm cành rong ảo & Điền Phiếu học tập số 1."
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(167, 243, 208)

    # Nút bấm bên phải
    btn_bot5 = create_card(s5_res, 8.85, 6.12, 3.5, 0.7, bg_color=RGBColor(255, 255, 255), border_color=RGBColor(52, 211, 153), border_width=1.2)
    tb_btn5 = s5_res.shapes.add_textbox(Inches(8.9), Inches(6.15), Inches(3.4), Inches(0.65))
    tf = tb_btn5.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "📝 Mở Phiếu Học Tập Số 1 ➔"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(6, 78, 59)
    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(2)
    r2 = p2.add_run()
    r2.text = "Khám phá bọt khí O₂ cành rong"
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(5, 150, 105)

    # ==========================================================
    # SLIDE 7: BOSCO VERTICALE - KHU RỪNG THẲNG ĐỨNG
    # ==========================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s5, BG_SOFT_MINT)
    add_header(s5, "Tư Liệu SGK Trang 107: 'Khu Rừng Thẳng Đứng' Lọc Không Khí 🏢🌿", "TƯ LIỆU THỰC TẾ SGK")
    add_footer(s5, 6, 18)

    # Left: Text Excerpt in soft pastel card
    c_txt = create_card(s5, 0.8, 1.7, 6.2, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s5.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.8), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🏢 TÒA THÁP KỲ DIỆU BOSCO VERTICALE (MILAN, Ý)"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_GREEN

    p = tf.add_paragraph()
    p.space_before = Pt(5)
    r = p.add_run()
    r.text = "📖 Hãy cùng đọc đoạn tư liệu thú vị trong SGK Trang 107 nhé:"
    r.font.italic = True
    r.font.size = Pt(10)
    r.font.color.rgb = T_MUTED

    sgk_text = (
        "Kiến trúc sư người Italia là Stê-pha-nô Bo-e-ri (Stefano Boeri) đã tạo ra toà tháp rừng thẳng đứng "
        "đầu tiên mang tên Bosco Verticale ở thành phố Milan, Italia.\n\n"
        "Toà tháp đôi này được phủ bằng một rừng cây gồm khoảng 900 cây xanh, 5.000 bụi cây với trên 11.000 loài thực vật khác nhau. "
        "Khu rừng này có thể hấp thụ CO₂, ngăn ngừa các khí gây hiệu ứng nhà kính phát tán lên bầu khí quyển.\n\n"
        "Cây xanh trên nóc toà nhà có thể lọc bụi mịn PM2.5. Bác Boeri đã khéo léo chọn những loại cây có khả năng hấp thụ khí ô nhiễm "
        "và bụi mịn PM2.5 tốt nhất."
    )
    p2 = tf.add_paragraph()
    p2.space_before = Pt(6)
    r2 = p2.add_run()
    r2.text = sgk_text
    r2.font.size = Pt(10)
    r2.font.color.rgb = T_BODY

    # Right: Image & Challenge
    if os.path.exists("screenshots/bosco_updated.png"):
        s5.shapes.add_picture("screenshots/bosco_updated.png", Inches(7.2), Inches(1.7), width=Inches(5.33), height=Inches(3.3))
    
    c_chal = create_card(s5, 7.2, 5.15, 5.33, 1.65, bg_color=C_PILL_PEACH, border_color=BORDER_AMBER, border_width=1.2)
    tb = s5.shapes.add_textbox(Inches(7.35), Inches(5.25), Inches(5.0), Inches(1.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌱 THỬ THÁCH CHO CÁC KỸ SƯ NHÍ:"
    r.font.bold = True
    r.font.size = Pt(11)
    r.font.color.rgb = T_AMBER

    p2 = tf.add_paragraph()
    p2.space_before = Pt(4)
    r2 = p2.add_run()
    r2.text = "Nếu em là trợ lý của bác kiến trúc sư, em sẽ chọn cây chịu nắng giỏi hay cây thích bóng mát để trồng lên nóc tháp gió to nắng gắt? Cùng tìm hiểu nào!"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = T_BODY

    # ==========================================================
    # SLIDE 7: YẾU TỐ 1 - ÁNH SÁNG
    # ==========================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s6, BG_SOFT_MINT)
    add_header(s6, "Yếu Tố 1: Ánh Nắng Ấm Áp — Bí Mật Của Cành Rong Nhỏ ☀️", "TIẾT 1 • THÍ NGHIỆM SỐ")
    add_footer(s6, 7, 18)

    # Left: Science & Friendly explanation
    c_left = create_card(s6, 0.8, 1.7, 6.2, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s6.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.8), Inches(4.8))
    tf = tf_box = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌿 BỌT KHÍ O₂ CỦA BẠN RONG ĐUÔI CHÓ (ELODEA)"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_GREEN

    light_cute = [
        ("Càng sáng thì thở càng hăng hái:", "Khi chiếu đèn vào bình nước, cành rong bắt đầu quang hợp và 'thở' ra từng bọt khí O₂ tròn xoe lấp lánh nổi lên bề mặt. Đếm số bọt khí là biết cây quang hợp nhanh hay chậm!"),
        ("Điểm bão hòa ánh sáng (~25.000 Lux):", "Khi tăng độ sáng vừa phải, cây thở bọt khí O₂ nhanh nhất. Đây là điểm cây 'vui sướng' nhất!"),
        ("Cẩn thận bẫy cháy lá (>45.000 Lux):", "Nếu bật đèn quá chói hoặc phơi nắng gắt mùa hè, lá cây bị nóng rát, hạt diệp lục bị bỏng rách ➔ cây giảm quang hợp ngay lập tức!"),
        ("Hai hội bạn cây thân thiết:", "• Cây ưa sáng (Bàng, Trúc, Bạch đàn): Phiến lá dày, thích đón nắng.\n• Cây ưa bóng (Lan ý, Dương xỉ, Trầu bà): Phiến lá mỏng manh, thích ở nơi râm mát.")
    ]
    for title, desc in light_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(5)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_GREEN
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # Right: Screenshot Lab
    if os.path.exists("screenshots/lab_elodea_realistic.png"):
        s6.shapes.add_picture("screenshots/lab_elodea_realistic.png", Inches(7.2), Inches(1.7), width=Inches(5.33), height=Inches(5.1))

    # ==========================================================
    # SLIDE 8: YẾU TỐ 2 - NƯỚC
    # ==========================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s7, BG_SOFT_MINT)
    add_header(s7, "Yếu Tố 2: Giọt Nước Mát Lành — 'Chiếc Miệng' Của Lá Cây 💧", "TIẾT 1 • THÍ NGHIỆM SỐ")
    add_footer(s7, 8, 18)

    # 2 Big Cards
    c1 = create_card(s7, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_SKY, border_width=1.5)
    tb = s7.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "💧 1. NGUYÊN LIỆU NẤU THỨC ĂN CỦA CÂY"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_BLUE

    water_cute_1 = [
        ("Nước chiếm 70% – 90% cơ thể cây:", "Cũng như chúng mình cần uống nước mỗi ngày, cây xanh cũng cần nước để giữ cho thân cành lá luôn căng tràn sức sống."),
        ("Nguyên liệu trực tiếp tạo ra Oxy:", "Dưới ánh nắng, nước được tách ra giải phóng khí Oxygen (O₂) tinh khiết cho con người và động vật hít thở:"),
        ("Công thức kỳ diệu:", "2H₂O  ➔  4H⁺  +  4e⁻  +  O₂ ↑ (Bay lên bầu trời)"),
        ("Làm mát cơ thể cây:", "Khi trời nắng, lá thoát hơi nước để hạ nhiệt độ, giống như khi chúng mình toát mồ hôi vậy!")
    ]
    for title, desc in water_cute_1:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = T_BLUE
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = T_BODY

    c2 = create_card(s7, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s7.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🔬 2. BÍ MẬT 'CHIẾC MIỆNG' KHÍ KHỔNG"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_GREEN

    water_cute_2 = [
        ("Khi cây uống đủ nước (Cười mở miệng):", "Hai tế bào hạt đậu no nước căng mọng lên, làm lỗ khí khổng mở toang để đón khí CO₂ ùa vào làm nguyên liệu quang hợp!"),
        ("Khi cây bị khát (Ngậm chặt miệng):", "Thiếu nước làm tế bào hạt đậu xẹp lép, lỗ khí khổng đóng chặt lại để giữ nước. Cửa đóng nên CO₂ không vào được ➔ quang hợp ngừng lại!"),
        ("Đừng tưới ngập úng nhé!:", "Tưới nước sũng ngập đáy chậu làm rễ cây bị ngạt thở không hít thở được O₂ ➔ rễ thối và cây bị 'chết khát trên đầm nước'."),
        ("Bí kíp tưới cây thông minh:", "Lắp hệ thống tưới nhỏ giọt hẹn giờ, đất ẩm vừa vặn 60% – 75% là tuyệt nhất!")
    ]
    for title, desc in water_cute_2:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = T_GREEN
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = T_BODY

    # ==========================================================
    # SLIDE 9: YẾU TỐ 3 & 4 - KHÍ CO2 VÀ NHIỆT ĐỘ
    # ==========================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s8, BG_SOFT_MINT)
    add_header(s8, "Yếu Tố 3 & 4: Khí CO₂ Và Nhiệt Độ — Cần Vừa Đủ Thôi Nhé! ☁️🌡️", "TIẾT 1 • THÍ NGHIỆM SỐ")
    add_footer(s8, 9, 18)

    # Card Left: CO2
    c1 = create_card(s8, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_AMBER, border_width=1.5)
    tb = s8.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "☁️ YẾU TỐ 3: HÀM LƯỢNG KHÍ CO₂"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_AMBER

    co2_cute = [
        ("Thức ăn dạng khí của cây:", "Lá cây hít khí CO₂ vào để tạo thành đường ngọt và tinh bột nuôi dưỡng toàn bộ cây."),
        ("Nồng độ tự nhiên trong không khí:", "Khoảng 0,04% (khoảng 400 phần triệu - ppm)."),
        ("Khoảng cây thích nhất:", "Từ 0,03% đến 0,15%. Trong khoảng này, có thêm chút CO₂ là cây mừng rỡ quang hợp nhanh hơn!"),
        ("Nguy cơ ngộ độc (>0,20%):", "Nếu khói xe xả ngạt thở vượt quá 0,20%, cây sẽ bị 'say ngạt' CO₂, khí khổng bị tê liệt và cây suy yếu dần.")
    ]
    for title, desc in co2_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = T_AMBER
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = T_BODY

    # Card Right: Nhiệt độ
    c2 = create_card(s8, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_ROSE, border_width=1.5)
    tb = s8.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌡️ YẾU TỐ 4: NHIỆT ĐỘ MÔI TRƯỜNG"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_ROSE

    temp_cute = [
        ("Nhiệt độ ấm áp lý tưởng (25°C – 35°C):", "Ở mức nhiệt độ dễ chịu này, các 'chú thợ tí hon' enzyme trong lá làm việc năng suất nhất, quang hợp đạt điểm 10!"),
        ("Trời lạnh dưới 10°C:", "Cây bị rét cóng, các chú thợ enzyme lười biếng ngủ đông, quang hợp chậm chạp."),
        ("Trời sốt nóng trên 40°C:", "Nhiệt độ quá nóng làm các enzyme bị hỏng (biến tính), lá cây bị kiệt sức và dừng quang hợp hoàn toàn."),
        ("Giải pháp làm mát:", "Kéo lưới lan che mát và phun sương để giảm nhiệt cho vườn cây đô thị.")
    ]
    for title, desc in temp_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title} "
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = T_ROSE
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(10)
        r2.font.color.rgb = T_BODY

    # ==========================================================
    # SLIDE 10: BẢNG BÍ KÍP QUANG HỢP (PHIẾU HỌC TẬP SỐ 1)
    # ==========================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s9, BG_SOFT_MINT)
    add_header(s9, "Bảng Bí Kíp Quang Hợp Của Nhà Sinh Học Nhí 📋✨", "TỔNG KẾT TIẾT 1")
    add_footer(s9, 10, 18)

    # Soft Pastel Table
    rows = [
        ("Yếu tố môi trường", "Mức cây thích nhất (Tối ưu)", "Khi quá mức hoặc thiếu thốn", "Bí kíp chăm sóc của em"),
        ("☀️ 1. Ánh sáng", "15.000 – 30.000 Lux\n(Điểm bão hòa ~25.000 Lux)", "Nắng gắt > 45.000 Lux: Bỏng lá\nTối < 3.000 Lux: Không quang hợp", "Cây ưa sáng để nơi nhiều nắng, cây ưa bóng để chỗ râm mát"),
        ("💧 2. Nước tưới", "Độ ẩm đất 60% – 75%\n(Khí khổng mở to đón CO₂)", "Thiếu nước: Khí khổng đóng lại\nNgập úng: Rễ ngạt thở, thối rễ", "Tưới nhỏ giọt đều đặn, chậu có lỗ thoát nước tốt"),
        ("☁️ 3. Khí CO₂", "0,03% – 0,15%\n(Khoảng 300 – 1.500 ppm)", "Quá ít (<0,008%): Đói CO₂\nQuá nhiều (>0,20%): Cây bị ngạt độc", "Để nơi thoáng khí, bón phân hữu cơ giải phóng CO₂ tự nhiên"),
        ("🌡️ 4. Nhiệt độ", "25°C – 35°C\n(Enzyme làm việc hăng say)", "Lạnh < 10°C: Cây lười quang hợp\nNóng > 40°C: Enzyme bị hỏng", "Kéo lưới che nắng, phun sương làm mát vào buổi trưa")
    ]

    table_shape = s9.shapes.add_table(5, 4, Inches(0.8), Inches(1.7), Inches(11.733), Inches(5.1))
    table = table_shape.table
    table.columns[0].width = Inches(2.4)
    table.columns[1].width = Inches(2.9)
    table.columns[2].width = Inches(3.3)
    table.columns[3].width = Inches(3.133)

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_PILL_MINT
                p.font.bold = True
                p.font.color.rgb = T_PRIMARY
                p.font.size = Pt(11)
                p.alignment = PP_ALIGN.CENTER
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_WHITE if r_idx % 2 == 1 else BG_CREAM
                p.font.size = Pt(9.8)
                p.font.color.rgb = T_BODY
                if c_idx == 0:
                    p.font.bold = True
                    p.font.color.rgb = T_GREEN

    # ==========================================================
    # SLIDE 11: TIẾT 2 - ĐỐ VUI NHANH
    # ==========================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s10, BG_SOFT_MINT)
    add_header(s10, "Thử Thách Vui: Bạn Nào Trả Lời Nhanh Nhất? 🎮⭐", "TIẾT 2 • KHỞI ĐỘNG VUI (10 PHÚT)")
    add_footer(s10, 11, 18)

    # 4 Quiz Question Cards (Pastel Style)
    quiz_items = [
        ("Câu 1: Yếu tố nào vừa là nguyên liệu tạo ra Oxy, vừa giúp 'chiếc miệng' khí khổng mở rộng?",
         "A. Khí CO₂\nB. Nước (H₂O)  ➔  [CHÍNH XÁC! 🎉]\nC. Nhiệt độ môi trường\nD. Bụi mịn PM2.5"),
        ("Câu 2: Khi nắng gắt chói chang (>45.000 Lux), điều gì sẽ xảy ra với lá cây?",
         "A. Cây quang hợp càng nhanh\nB. Cây tạo ra nhiều quả ngọt hơn\nC. Diệp lục bị bỏng rát, cháy lá, quang hợp giảm mạnh  ➔  [ĐÚNG RỒI! ☀️]\nD. Lá cây biến thành màu đỏ"),
        ("Câu 3: Khí khổng sẽ bị tê liệt và cây bị ngộ độc khi nồng độ khí CO₂ vượt quá mức nào?",
         "A. Khoảng 0,04%\nB. Khoảng 0,10%\nC. Vượt quá 0,20%  ➔  [CHUẨN LUÔN! ☁️]\nD. Dưới 0,02%"),
        ("Câu 4: Vì sao nhiệt độ nóng trên 40°C lại làm cây ngừng quang hợp?",
         "A. Nước bị bốc hơi thành mây\nB. Các chú thợ enzyme quang hợp bị hỏng (biến tính)  ➔  [XUẤT SẮC! 🌡️]\nC. Hết khí oxy\nD. Cây biến thành thân gỗ")
    ]

    for idx, (q, opts) in enumerate(quiz_items):
        col = idx % 2
        row = idx // 2
        left = 0.8 + col * 6.0
        top = 1.7 + row * 2.55
        c = create_card(s10, left, top, 5.7, 2.35, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
        tb = s10.shapes.add_textbox(Inches(left + 0.15), Inches(top + 0.1), Inches(5.4), Inches(2.15))
        tf = tb.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = q
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = T_PRIMARY

        p2 = tf.add_paragraph()
        p2.space_before = Pt(4)
        r2 = p2.add_run()
        r2.text = opts
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # ==========================================================
    # SLIDE 12: VẬN DỤNG STEM - 4 ĐỘI KỸ SƯ NHÍ
    # ==========================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s11, BG_SOFT_MINT)
    add_header(s11, "Dự Án STEM: 4 Đội Kỹ Sư Sinh Thái Nhí Xuất Kích! 🚀🌿", "TIẾT 2 • VẬN DỤNG DỰ ÁN")
    add_footer(s11, 12, 18)

    # 4 Team Cards
    teams = [
        ("NHÓM 1: ĐÔ THỊ NÉN & THÁP RỪNG ĐỨNG", "Trạm 1: Chung cư & Nhà phố TP.HCM", "Khói xe CO₂ cao (850 ppm), bê tông nóng 36°C; tầng trên nắng gắt 45.000 Lux, tầng trệt râm.", "Phân tầng cây ưa sáng trên mái, giàn leo vách đứng, cây ưa bóng tầng trệt & tưới nhỏ giọt.", BORDER_MINT, T_GREEN),
        ("NHÓM 2: NHÀ MÀNG NÔNG NGHIỆP SỐ", "Trạm 2: Nông nghiệp công nghệ cao Củ Chi", "Cần tối ưu cả 4 yếu tố quanh năm (hiện trạng thiếu sáng 12.000 Lux, CO₂ 400 ppm, nóng 34°C).", "Lắp đèn LED quang phổ, bổ sung CO₂ (1.200 ppm), quạt hạ nhiệt 26°C & tưới nhỏ giọt tự động.", BORDER_SKY, T_BLUE),
        ("NHÓM 3: SÂN THƯỢNG NẮNG GẮT", "Trạm 3: Mái bê tông & Ban công hướng Tây", "Nắng trưa gay gắt 58.000 Lux, sàn bê tông 43°C làm khí khổng đóng chặt, lá dễ cháy khô.", "Căng lưới lan giảm nắng gắt, trồng cây chịu hạn cutin dày, phun sương làm mát & phủ gốc giữ ẩm.", BORDER_AMBER, T_AMBER),
        ("NHÓM 4: TRONG NHÀ & HÀNH LANG RÂM", "Trạm 4: Lớp học & Khu vực thiếu sáng", "Ánh sáng rất yếu (800 Lux, dưới điểm bù cây ưa sáng), thiếu thông gió tự nhiên, CO₂ 650 ppm.", "Chọn cây ưa bóng giàu diệp lục b, đo sáng bằng smartphone, bù đèn LED 4000K & lau sạch bụi lá.", BORDER_ROSE, T_ROSE)
    ]

    for idx, (t_name, sub, problem, goal, col_border, col_title) in enumerate(teams):
        c_left = 0.8 + (idx % 2) * 6.0
        c_top = 1.7 + (idx // 2) * 2.55
        c = create_card(s11, c_left, c_top, 5.7, 2.35, bg_color=C_WHITE, border_color=col_border, border_width=1.5)
        tb = s11.shapes.add_textbox(Inches(c_left + 0.15), Inches(c_top + 0.1), Inches(5.4), Inches(2.15))
        tf = tb.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        r = p.add_run()
        r.text = f"🌱 {t_name} "
        r.font.bold = True
        r.font.size = Pt(11)
        r.font.color.rgb = col_title

        r_sub = p.add_run()
        r_sub.text = f"({sub})"
        r_sub.font.bold = True
        r_sub.font.size = Pt(9.5)
        r_sub.font.color.rgb = T_MUTED

        p_prob = tf.add_paragraph()
        p_prob.space_before = Pt(4)
        r = p_prob.add_run()
        r.text = "• Thách thức: "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_BODY
        r2 = p_prob.add_run()
        r2.text = problem
        r2.font.size = Pt(9.5)

        p_goal = tf.add_paragraph()
        p_goal.space_before = Pt(3)
        r = p_goal.add_run()
        r.text = "• Nhiệm vụ của đội: "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = col_title
        r2 = p_goal.add_run()
        r2.text = goal
        r2.font.size = Pt(9.5)

    # ==========================================================
    # SLIDE 13: THAO TÁC TRẠM 1 - KHU VƯỜN THẲNG ĐỨNG
    # ==========================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s12, BG_SOFT_MINT)
    add_header(s12, "Trải Nghiệm Trạm 1: Kéo Cây Vào Trồng Cực Vui Nhé! 🌿🏡", "THAO TÁC THỰC HÀNH WEB")
    add_footer(s12, 13, 18)

    # Left: Steps
    c_left = create_card(s12, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s12.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌱 CÁCH TRỒNG CÂY TRÊN MÀN HÌNH MÁY TÍNH:"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_GREEN

    st1_cute = [
        ("Bước 1: Chọn cây tầng thượng", "Kéo Cây Bàng Đài Loan hoặc Trúc Cần Câu vào tầng trên cùng. Cây ưa sáng lá dày, chịu nắng gió và cản bụi rất cừ!"),
        ("Bước 2: Phủ xanh vách tường đứng", "Kéo Dây Thường Xuân hoặc Hoa Giấy leo bám tường, che mát bề mặt bê tông nóng bức."),
        ("Bước 3: Trồng cây tầng trệt", "Kéo Cây Lan Ý hoặc Dương Xỉ vào tầng trệt. Cây ưa bóng lá mỏng lọc khí độc cực tốt!"),
        ("Bước 4: Cấp nước thông minh", "Kéo Hệ Thống Tưới Nhỏ Giọt Hẹn Giờ để rễ luôn đủ ẩm mà không lo ngập úng."),
        ("Các nút chăm sóc rộn ràng", "Bấm [Tưới Nước 🌧️], [Kéo Lưới Che Nắng ⛱️], [Thổi Gió Mát 🍃]. Theo dõi điểm hiệu suất quang hợp tăng vọt đạt 100/100đ!")
    ]
    for title, desc in st1_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(5)
        r = p.add_run()
        r.text = f"• {title}:\n  "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_GREEN
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # Right: Screenshot Garden
    if os.path.exists("screenshots/garden_planted.png"):
        s12.shapes.add_picture("screenshots/garden_planted.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 14: BÀN LÀM VIỆC CÔ MAI THUÝ ĐỨC
    # ==========================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s13, BG_SOFT_MINT)
    add_header(s13, "Bàn Làm Việc Cô Mai Thuý Đức: Chờ Cô Chấm Điểm & Chuẩn Y 👩‍🏫✨", "CỔNG PHÊ DUYỆT BÀI")
    add_footer(s13, 14, 18)

    # Left: Explanation
    c_left = create_card(s13, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_AMBER, border_width=1.5)
    tb = s13.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "💌 GỬI BÀI ĐẾN BÀN LÀM VIỆC CỦA CÔ:"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_AMBER

    hub_cute = [
        ("Bấm nút nộp bài thật nhanh:", "Khi đã chọn đủ 4 giải pháp xanh, nhóm bấm [Nộp Hồ Sơ Thiết Kế]. Bài thi sẽ bay vèo đến Bàn Làm Việc của Cô Mai Thuý Đức!"),
        ("Trạng thái hồi hộp:", "Màn hình nhóm sẽ hiện dòng chữ: 'Đang Chờ Cô Mai Thuý Đức Duyệt'."),
        ("Cô giáo nhận xét & cho điểm:", "Cô Đức mở bài của nhóm, kiểm tra xem cây ưa sáng và ưa bóng đã đúng vị trí chưa, sau đó tặng điểm thưởng và lời khen động viên!"),
        ("Mở khóa vinh quang:", "Ngay khi cô bấm [Chuẩn Y & Phê Duyệt], chiếc E-Poster của nhóm sẽ chính thức bừng sáng trên Sảnh Triển Lãm để cả lớp cùng ngắm nhìn!")
    ]
    for title, desc in hub_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_AMBER
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # Right: Screenshot Poster Gallery
    if os.path.exists("screenshots/exp_4_poster_gallery_approved.png"):
        s13.shapes.add_picture("screenshots/exp_4_poster_gallery_approved.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 15: THUYẾT TRÌNH BÁO CÁO E-POSTER KHOA HỌC
    # ==========================================================
    s14 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s14, BG_SOFT_MINT)
    add_header(s14, "Triển Lãm E-Poster: Cả Lớp Cùng Khoe Thành Quả Rực Rỡ! 🎨🎙️", "THUYẾT TRÌNH DỰ ÁN")
    add_footer(s14, 15, 18)

    # Left: Explanation
    c_left = create_card(s14, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s14.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "📊 CHIẾC E-POSTER CHUẨN XỊN 3 CỘT:"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_GREEN

    poster_cute = [
        ("Cột 1: Câu chuyện trạm của nhóm", "Kể cho các bạn nghe trạm nhóm mình nóng bức thế nào, bụi bặm ra sao và cây xanh ở đây gặp khó khăn gì."),
        ("Cột 2: Giải pháp thông minh của nhóm", "Khoe những loại cây xanh đặc biệt nhóm đã chọn và giải thích cơ chế quang hợp Bài 23 thật tự tin!"),
        ("Cột 3: Kết quả vi khí hậu & Lời phê của cô", "Bảng đồng hồ đo CO₂ hạ xuống 410 ppm mát lành, điểm quang hợp 96/100 kèm lời khen của Cô Mai Thuý Đức."),
        ("Bình chọn bạn yêu thích ⭐:", "Mỗi bạn học sinh trong lớp đều có thể bấm tặng 1 đến 5 sao yêu thích cho poster của nhóm bạn!")
    ]
    for title, desc in poster_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_GREEN
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # Right: Screenshot Poster
    if os.path.exists("screenshots/exp_5_presentation_poster.png"):
        s14.shapes.add_picture("screenshots/exp_5_presentation_poster.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 16: TRAO GIẤY CHỨNG NHẬN KỸ SƯ SINH THÁI SỐ
    # ==========================================================
    s15 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s15, BG_SOFT_MINT)
    add_header(s15, "Vinh Danh: Nhận Giấy Chứng Nhận Kỹ Sư Sinh Học Xanh! 🎖️🏆", "TRAO THƯỞNG DỰ ÁN")
    add_footer(s15, 16, 18)

    # Left: Explanation
    c_left = create_card(s15, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_AMBER, border_width=1.5)
    tb = s15.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "⭐ CÁC DANH HIỆU VINH DANH ECO-POINTS:"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_AMBER

    cert_cute = [
        ("👑 ĐẠI SỨ XANH GDPT 2018 (Từ 480 – 500đ):", "Dành cho các đội kỹ sư xuất sắc toàn diện, tối ưu trọn vẹn cả 4 yếu tố quang hợp."),
        ("🥇 KỸ SƯ SINH THÁI ĐÔ THỊ (Từ 400 – 479đ):", "Huy chương vàng sáng giá, làm chủ giải pháp phủ xanh thích ứng biến đổi khí hậu."),
        ("🥈 CHIẾN BINH SINH THÁI (Từ 250 – 399đ):", "Huy chương bạc nhiệt huyết, tích cực chăm sóc cây xanh."),
        ("Tải Giấy Chứng Nhận về khoe bố mẹ:", "Bấm nút [Tải Giấy Chứng Nhận (PNG)] để lưu tấm bằng khen có con dấu đỏ ĐÃ NGHIỆM THU và chữ ký của Cô Mai Thuý Đức (Trường THCS Tân Tạo A)!")
    ]
    for title, desc in cert_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_AMBER
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # Right: Screenshot Cert
    if os.path.exists("screenshots/screen_certificate.png"):
        s15.shapes.add_picture("screenshots/screen_certificate.png", Inches(6.8), Inches(1.7), width=Inches(5.733), height=Inches(5.1))

    # ==========================================================
    # SLIDE 17: DỰ ÁN NHỎ VỀ NHÀ
    # ==========================================================
    s16 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s16, BG_SOFT_MINT)
    add_header(s16, "Dự Án Nhỏ Về Nhà: 'Góc Xanh Ban Công Của Em' 🏡🌱", "TỰ HỌC & MỞ RỘNG")
    add_footer(s16, 17, 18)

    # Left: Rubric friendly
    c_rubric = create_card(s16, 0.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=1.5)
    tb = s16.shapes.add_textbox(Inches(1.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🎯 TIÊU CHÍ ĐỂ ĐẠT ĐIỂM 10 CỦA CÔ:"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_GREEN

    rub_cute = [
        ("1. Hiểu bài thật sâu (25%):", "Nhớ được 4 người bạn của cây (Nắng, Nước, CO₂, Nhiệt độ) và phân biệt cây ưa sáng / ưa bóng."),
        ("2. Thao tác siêu mượt (25%):", "Bấm chuột làm thí nghiệm ảo thành thạo, đọc biểu đồ và đếm bọt khí O₂ chuẩn xác."),
        ("3. Sản phẩm E-Poster đẹp mắt (30%):", "Trồng cây phủ xanh đủ 4 tầng, poster trình bày khoa học, thẩm mỹ và đạt điểm quang hợp cao."),
        ("4. Tinh thần đồng đội tuyệt vời (20%):", "Cùng bạn thảo luận vui vẻ, đoàn kết và tự tin thuyết trình trước lớp.")
    ]
    for title, desc in rub_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_GREEN
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # Right: Homework friendly
    c_hw = create_card(s16, 6.8, 1.7, 5.7, 5.1, bg_color=C_WHITE, border_color=BORDER_SKY, border_width=1.5)
    tb = s16.shapes.add_textbox(Inches(7.0), Inches(1.85), Inches(5.3), Inches(4.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "🌸 NHIỆM VỤ DỄ THƯƠNG TẠI NHÀ:"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = T_BLUE

    hw_cute = [
        ("Nhiệm vụ 1: Hoàn thành vở ghi KHTN 7", "Vẽ một sơ đồ tư duy hình cây xanh có 4 cành tượng trưng cho 4 yếu tố quang hợp."),
        ("Nhiệm vụ 2: Dự án 'Góc Xanh Ban Công'", "Tìm một góc cửa sổ hoặc ban công nhà em:\n• Xem hướng nắng buổi sáng hay buổi trưa.\n• Chọn 1 chậu cây ưa sáng và 1 chậu cây ưa bóng rồi tưới nước chăm sóc nhé!"),
        ("Nhiệm vụ 3: Vào web chơi game sinh thái", "Truy cập https://sinhhoc-amber.vercel.app/ để thử thách bản thân với 3 trạm sinh thái còn lại."),
        ("Chuẩn bị bài mới:", "Đọc trước Bài 24: Thực hành Chứng minh quang hợp ở cây xanh.")
    ]
    for title, desc in hw_cute:
        p = tf.add_paragraph()
        p.space_before = Pt(8)
        r = p.add_run()
        r.text = f"• {title}\n  "
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = T_BLUE
        r2 = p.add_run()
        r2.text = desc
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = T_BODY

    # ==========================================================
    # SLIDE 18: TRANG KẾT THÚC SIÊU NGỌT NGÀO
    # ==========================================================
    s17 = prs.slides.add_slide(blank_layout)
    set_slide_bg(s17, BG_TITLE)

    # Card Banner Center
    card_end = create_card(s17, 0.8, 0.7, 11.733, 6.1, bg_color=C_WHITE, border_color=BORDER_MINT, border_width=2)

    if os.path.exists("images/logo.png"):
        s17.shapes.add_picture("images/logo.png", Inches(5.9), Inches(1.15), width=Inches(1.5), height=Inches(1.5))

    tb_end = s17.shapes.add_textbox(Inches(1.3), Inches(2.9), Inches(10.7), Inches(3.6))
    tf = tb_end.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "💖 CẢM ƠN CÁC BẠN NHỎ ĐÃ HỌC THẬT CHĂM CHỈ! 💖"
    r.font.bold = True
    r.font.size = Pt(26)
    r.font.color.rgb = T_PRIMARY
    r.font.name = "Arial"

    p2 = tf.add_paragraph()
    p2.alignment = PP_ALIGN.CENTER
    p2.space_before = Pt(12)
    r2 = p2.add_run()
    r2.text = "Chúc các em luôn yêu quý thiên nhiên, say mê khám phá Khoa học tự nhiên\nvà cùng nhau xây dựng môi trường sống xanh tươi, an lành!"
    r2.font.size = Pt(13.5)
    r2.font.color.rgb = T_GREEN
    r2.font.name = "Arial"

    p3 = tf.add_paragraph()
    p3.alignment = PP_ALIGN.CENTER
    p3.space_before = Pt(22)
    r3 = p3.add_run()
    r3.text = "CÔ GIÁO: MAI THUÝ ĐỨC (KHTN SINH HỌC - 9 NĂM KINH NGHIỆM)  🌸  TRƯỜNG THCS TÂN TẠO A\n"
    r3.font.bold = True
    r3.font.size = Pt(12)
    r3.font.color.rgb = T_PRIMARY

    r4 = p3.add_run()
    r4.text = "📞 SĐT: 0975515659  •  ✉️ Email: maithuyduc36@gmail.com  •  🌐 Web: https://sinhhoc-amber.vercel.app/"
    r4.font.bold = True
    r4.font.size = Pt(11)
    r4.font.color.rgb = T_GREEN

    # Save presentation to both names
    prs.save("KHTN7_Bai23_QuangHop_MaiThuyDuc.pptx")
    prs.save("03_Slide_Trinh_Chieu_Giang_Day_Mai_Thuy_Duc.pptx")
    print("Cute presentation saved to both files.")

if __name__ == "__main__":
    create_cute_deck()
