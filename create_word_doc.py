# -*- coding: utf-8 -*-
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_demo_doc(filename):
    doc = docx.Document()
    
    # Page setup: A4, Margins 2cm
    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.79)
        section.bottom_margin = Inches(0.79)
        section.left_margin = Inches(0.79)
        section.right_margin = Inches(0.79)

    # Styles
    COLOR_PRIMARY = RGBColor(27, 54, 93)     # Navy #1B365D
    COLOR_SECONDARY = RGBColor(46, 117, 182) # Blue #2E75B6
    COLOR_TEXT = RGBColor(38, 38, 38)        # Charcoal #262626
    COLOR_MUTED = RGBColor(100, 100, 100)
    
    # Base Normal Style
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11.5)
    normal_style.font.color.rgb = COLOR_TEXT
    normal_style.paragraph_format.line_spacing = 1.25
    normal_style.paragraph_format.space_after = Pt(4)

    def set_cell_background(cell, hex_color):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
            node = OxmlElement(m)
            node.set(qn('w:w'), str(val))
            node.set(qn('w:type'), 'dxa')
            tcMar.append(node)
        tcPr.append(tcMar)

    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(16)
        run.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(14)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.italic = True
        run.font.color.rgb = COLOR_MUTED
        return p

    def add_heading1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_heading2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11.5)
        run.bold = True
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_bullet(text, bold_prefix="", level=0):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.left_indent = Inches(0.25 * (level + 1))
        if bold_prefix:
            r_b = p.add_run(bold_prefix)
            r_b.bold = True
            r_b.font.color.rgb = COLOR_PRIMARY
        r_t = p.add_run(text)
        return p

    def add_callout(text, bold_title=""):
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        cell.width = Inches(6.69)
        set_cell_background(cell, "F0F4F8")
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        
        # Left border only
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}>'
                            f'<w:top w:val="none"/>'
                            f'<w:left w:val="single" w:sz="24" w:space="0" w:color="1B365D"/>'
                            f'<w:bottom w:val="none"/>'
                            f'<w:right w:val="none"/>'
                            f'</w:tcBorders>')
        tcPr.append(borders)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.2
        if bold_title:
            rb = p.add_run(bold_title + " ")
            rb.bold = True
            rb.font.color.rgb = COLOR_PRIMARY
        rt = p.add_run(text)
        rt.font.color.rgb = COLOR_TEXT
        
        doc.add_paragraph().paragraph_format.space_after = Pt(4)

    # ------------------ DOCUMENT CONTENT ------------------

    # Header block
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_top1 = p_top.add_run("SỞ GIÁO DỤC VÀ ĐÀO TẠO THÀNH PHỐ HỒ CHÍ MINH\n")
    r_top1.bold = True
    r_top1.font.size = Pt(10)
    r_top2 = p_top.add_run("HỘI THI THIẾT KẾ CHỦ ĐỀ DẠY HỌC STEM/STEAM, CHỦ ĐỀ DẠY HỌC SỐ NĂM HỌC 2026 - 2027\n")
    r_top2.bold = True
    r_top2.font.size = Pt(11)
    r_top2.font.color.rgb = COLOR_PRIMARY
    r_top3 = p_top.add_run("(Căn cứ Kế hoạch số 12070/KH-SGDĐT ngày 09/11/2026)")
    r_top3.italic = True
    r_top3.font.size = Pt(9.5)
    r_top3.font.color.rgb = COLOR_MUTED
    p_top.paragraph_format.space_after = Pt(12)

    add_title("BẢN DEMO ĐỀ XUẤT Ý TƯỞNG DỰ THI")
    add_title("THIẾT KẾ MÔ HÌNH \"NGÔI NHÀ CÁCH ÂM THÔNG MINH\"\nCHỐNG Ô NHIỄM TIẾNG ỒN ĐÔ THỊ")
    add_subtitle("Môn học: Khoa học tự nhiên 7 | Bài 14: Phản xạ âm, chống ô nhiễm tiếng ồn\nThời lượng: 02 tiết (90 phút) | Quy mô: 03 nhóm học sinh")

    add_callout(
        "Ý tưởng chủ đạo: Kết hợp giữa Khoa học Vật lý (Âm học), Kỹ thuật chế tạo mô hình kiến trúc với 3 vật liệu chủ lực "
        "(Tấm Fomex, Tấm xốp, Giấy khay trứng gà) và Công nghệ Dạy học số (đo lường cường độ âm thanh thời gian thực bằng "
        "ứng dụng số trên Smartphone). Dự án vừa vặn trong 2 tiết, giải quyết vấn đề thực tế nhức nhối của đô thị TP.HCM.",
        "★ ĐIỂM SÁNG DỰ THI:"
    )

    # PHẦN 1
    add_heading1("I. PHẦN TỔNG QUAN CHỦ ĐỀ (QUY CHUẨN FOLDER 1 - STEM_P1)")
    
    add_bullet(" Khoa học tự nhiên 7 (Bộ sách Kết nối tri thức với cuộc sống).", "1. Môn học dự thi:")
    add_bullet(" Bài 14: Phản xạ âm, chống ô nhiễm tiếng ồn (Chương IV – Âm thanh).", "2. Bài học trọng tâm:")
    add_bullet(" Không quá 02 tiết (90 phút).", "3. Thời lượng thực hiện:")
    add_bullet(" Phòng học bộ môn Khoa học tự nhiên / Phòng học đa phương tiện có máy chiếu và kết nối Internet.", "4. Không gian tổ chức:")
    
    add_bullet("", "5. Bối cảnh thực tiễn & Tình huống xuất phát (Hook):")
    p_ctx = doc.add_paragraph()
    p_ctx.paragraph_format.left_indent = Inches(0.25)
    p_ctx.add_run(
        "Thành phố Hồ Chí Minh là đô thị đặc biệt với mật độ giao thông và xây dựng dày đặc. Tiếng ồn từ xe cộ, còi xe, "
        "máy móc công trình và máy bay (đặc biệt khu vực Gò Vấp, Tân Bình gần sân bay Tân Sơn Nhất) thường xuyên vượt ngưỡng "
        "an toàn 70 - 85 dB, gây ảnh hưởng nghiêm trọng đến sức khỏe, giấc ngủ và hiệu quả học tập. Đặt học sinh vào vai "
        "\"Kỹ sư âm học đô thị\", các em nhận nhiệm vụ thiết kế một mô hình ngôi nhà có khả năng cách âm, bảo vệ không gian sống."
    )

    add_bullet("", "6. Vị trí kiến thức trong chương trình:")
    add_bullet("Khái niệm sóng âm, nguồn âm, độ to của âm (đo bằng decibel - dB), sự truyền âm qua chất rắn, lỏng, khí (Bài 12, 13 – KHTN 7). Kỹ năng hình học, đo đạc kích thước và thao tác ứng dụng trên điện thoại thông minh.", "• Kiến thức, kỹ năng nền: ", level=1)
    add_bullet("Nhận biết hiện tượng phản xạ âm; phân loại vật liệu phản xạ âm tốt (cứng, nhẵn) và vật liệu hấp thụ âm tốt (mềm, xốp, gồ ghề); vận dụng 3 nguyên tắc chống ô nhiễm tiếng ồn (ngăn chặn, phân tán, hấp thụ âm); quy trình thiết kế kỹ thuật STEM 5 bước.", "• Kiến thức, kỹ năng mới hình thành: ", level=1)

    add_bullet("", "7. Mục tiêu cần đạt (Chuẩn CT GDPT 2018):")
    add_bullet("Lấy được ví dụ về vật liệu phản xạ âm tốt, phản xạ âm kém; giải thích được cơ chế chống ồn và đề xuất các biện pháp khả thi để giảm thiểu tiếng ồn trong đời sống.", "• Năng lực Khoa học tự nhiên: ", level=1)
    add_bullet("Thực hiện quy trình thiết kế kỹ thuật; biết lập bản vẽ mặt cắt kết cấu vách tường; tính toán diện tích, tối ưu hóa vật liệu.", "• Năng lực Kỹ thuật & Công nghệ: ", level=1)
    add_bullet("Sử dụng ứng dụng số (Decibel X / Phyphox) để thu thập và đo lường cường độ âm thanh khách quan; số hóa kết quả đo và báo cáo trên nền tảng Padlet.", "• Năng lực Dạy học số: ", level=1)
    add_bullet("Có ý thức giữ gìn sự yên tĩnh nơi công cộng, tinh thần làm việc nhóm kỷ luật, cẩn thận và an toàn lao động.", "• Phẩm chất chủ yếu: ", level=1)

    # PHẦN 2
    add_heading1("II. GIẢI MÃ BỘ 3 VẬT LIỆU CHỦ LỰC VÀ GIẢI PHÁP KỸ THUẬT")
    
    add_heading2("1. Cơ chế khoa học của 3 loại vật liệu")
    p_mat = doc.add_paragraph()
    p_mat.add_run(
        "Điểm đột phá của dự án là việc phối hợp nhịp nhàng giữa 3 loại vật liệu rất dễ tìm, giá thành rẻ nhưng mang đầy đủ "
        "các đặc tính vật lý âm học đối lập và bổ trợ cho nhau:"
    )

    # Table of materials
    table_mat = doc.add_table(rows=4, cols=3)
    table_mat.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = [Inches(1.5), Inches(2.2), Inches(3.0)]
    
    headers = ["Tên vật liệu", "Đặc tính vật lý", "Vai trò trong cách âm & tiêu âm"]
    for i, h in enumerate(headers):
        cell = table_mat.cell(0, i)
        cell.width = col_widths[i]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(10.5)

    mat_data = [
        ("Tấm Fomex\n(Formex 3mm/5mm)", "Cứng cáp, bề mặt phẳng mịn, đàn hồi nhẹ, cắt gọt dễ dàng, màu trắng thẩm mỹ.", 
         "Đóng vai trò VỎ CỨNG phản xạ âm: Ngăn chặn sóng âm truyền thẳng, làm dội ngược phần lớn năng lượng âm thanh ban đầu và tạo khung nhà vững chãi."),
        ("Tấm xốp\n(PE-Foam / Xốp bọt biển)", "Mềm, xốp, nhẹ, có vô số túi khí li ti đàn hồi bên trong cấu trúc.", 
         "Đóng vai trò LỚP ĐỆM HẤP THỤ ÂM: Triệt tiêu các dao động âm thanh lọt qua vỏ Fomex, biến đổi cơ năng dao động của sóng âm thành nhiệt năng li ti."),
        ("Giấy khay trứng gà\n(Bìa giấy ép định hình)", "Chất liệu bột giấy xốp, bề mặt lồi lõm zíc zắc hình nón cụt đa chiều.", 
         "Đóng vai trò BỀ MẶT TÁN ÂM & TIÊU ÂM NỘI THẤT: Phá vỡ các sóng âm phản xạ song song, triệt tiêu hiện tượng dội âm (tiếng vang), giúp bên trong nhà tĩnh lặng tối đa.")
    ]

    for row_idx, data in enumerate(mat_data, start=1):
        bg = "FFFFFF" if row_idx % 2 != 0 else "F7FAFC"
        for col_idx, text in enumerate(data):
            cell = table_mat.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.size = Pt(10)
            if col_idx == 0:
                r.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(4)

    add_heading2("2. Định hướng giải pháp kỹ thuật cho 3 nhóm học sinh")
    p_grp = doc.add_paragraph()
    p_grp.add_run(
        "Để tạo không khí học tập hào hứng, lớp học được chia thành 3 Đội Kỹ sư. Cả 3 nhóm đều được cấp cùng một khung "
        "nhà bằng tấm Fomex tiêu chuẩn (kích thước 20 x 20 x 25 cm), nhưng tiếp cận xử lý cách âm theo 3 trường phái kỹ thuật khác nhau:"
    )

    add_bullet("Tường Fomex cứng + Lót toàn bộ mặt trong bằng Giấy khay trứng gà.", "• NHÓM 1: Đội Kiến trúc Tiêu âm & Tán âm (Acoustic Studio)")
    p_g1 = doc.add_paragraph()
    p_g1.paragraph_format.left_indent = Inches(0.4)
    p_g1.add_run(
        "→ Giải pháp: Bề mặt lồi lõm của khay trứng hướng vào lòng nhà. Khi tiếng ồn đi vào, sóng âm bị phân tán nhiều hướng "
        "và chất bột giấy tiêu âm triệt để hiện tượng cộng hưởng âm thanh. Mô phỏng phòng thu âm hoặc phòng karaoke hiện đại."
    )

    add_bullet("Tường 3 lớp Fomex ngoài – Tấm xốp đệm giữa – Fomex trong (Vách Sandwich kép).", "• NHÓM 2: Đội Kiến trúc Vách kép Chống ồn Đô thị (Double-Wall Tech)")
    p_g2 = doc.add_paragraph()
    p_g2.paragraph_format.left_indent = Inches(0.4)
    p_g2.add_run(
        "→ Giải pháp: Tạo cấu trúc tường đa tầng cản âm tối đa: Fomex chặn âm → Tấm xốp triệt tiêu dao động cơ học → Lớp Fomex "
        "bên trong ngăn nốt lượng âm còn lại. Viền cửa sổ được chèn thêm dải xốp làm gioăng làm kín."
    )

    add_bullet("Kết hợp cả 3 vật liệu: Fomex ngoài + Tấm xốp ở giữa + Giấy khay trứng gà trong cùng.", "• NHÓM 3: Đội Kiến trúc Hỗn hợp Đa tầng (Hybrid Master System)")
    p_g3 = doc.add_paragraph()
    p_g3.paragraph_format.left_indent = Inches(0.4)
    p_g3.add_run(
        "→ Giải pháp: Giải pháp toàn diện nhất: Tận dụng ưu điểm phản xạ của Fomex, khả năng nuốt âm của Tấm xốp và khả năng "
        "tán xạ triệt tiêu tiếng vang của Khay trứng. Nhóm này đại diện cho mô hình công trình cách âm cao cấp."
    )

    # PHẦN 3
    add_heading1("III. KẾ HOẠCH BÀI DẠY CHI TIẾT 02 TIẾT (QUY CHUẨN FOLDER 2 - STEM_P2)")

    add_heading2("TIẾT 1 (45 PHÚT): KHÁM PHÁ VẬT LIỆU & THIẾT KẾ GIẢI PHÁP KỸ THUẬT")
    
    add_bullet("Khởi động & Chuyển giao nhiệm vụ (8 phút)", "1. Hoạt động 1: ")
    p_a1 = doc.add_paragraph()
    p_a1.paragraph_format.left_indent = Inches(0.3)
    p_a1.add_run(
        "• Giáo viên chiếu video clip 1,5 phút: Mô phỏng tiếng ồn đường phố và tiếng máy bay gầm rú tại TP.HCM.\n"
        "• GV đặt vấn đề: Làm sao để bảo vệ giấc ngủ và sự tập trung của học sinh trong ngôi nhà đô thị?\n"
        "• GV giao đơn đặt hàng: Mỗi nhóm thiết kế mô hình nhà cách âm đạt tiêu chuẩn kỹ thuật: Giảm tối thiểu 15 - 20 dB "
        "khi kiểm định tại Trạm đo âm thanh số của lớp học.\n"
        "• GV giới thiệu Bảng tiêu chí đánh giá (Rubric) và mã QR dẫn vào bảng Padlet của lớp."
    )

    add_bullet("Nghiên cứu kiến thức nền & Thí nghiệm số khám phá vật liệu (15 phút)", "2. Hoạt động 2: ")
    p_a2 = doc.add_paragraph()
    p_a2.paragraph_format.left_indent = Inches(0.3)
    p_a2.add_run(
        "• Thí nghiệm Mini-Lab số: GV phát cho mỗi nhóm 3 tấm vật liệu mẫu (10 x 10 cm): (1) Tấm Fomex nhẵn, (2) Tấm xốp, (3) Miếng khay trứng gà.\n"
        "• Học sinh dùng điện thoại phát 1 đoạn âm thanh chuẩn (75 dB), lần lượt đặt từng tấm vật liệu chắn trước nguồn âm và mở app Decibel X "
        "trên điện thoại thứ 2 để đo cường độ âm lọt qua.\n"
        "• Rút ra định luật khoa học: Fomex chặn âm ban đầu nhưng phản xạ lại; Tấm xốp làm giảm biên độ sóng âm; Khay trứng làm tiêu biến tiếng dội.\n"
        "• Học sinh điền kết quả vào Phiếu học tập số số 1."
    )

    add_bullet("Đề xuất giải pháp và Phác thảo bản vẽ kỹ thuật (15 phút)", "3. Hoạt động 3: ")
    p_a2 = doc.add_paragraph()
    p_a2.paragraph_format.left_indent = Inches(0.3)
    p_a2.add_run(
        "• Các nhóm thảo luận, phân công nhiệm vụ (Nhóm trưởng, Kỹ sư thiết kế, Kỹ sư chế tạo, Chuyên viên dữ liệu số).\n"
        "• Vẽ Bản thiết kế mặt cắt ngang tường và mái nhà (ghi rõ thứ tự sắp xếp các lớp Fomex, xốp và khay trứng).\n"
        "• Tính toán kích thước, diện tích các mặt tường để tối ưu hóa vật liệu, không gây lãng phí.\n"
        "• Chuyên viên dữ liệu số chụp ảnh bản vẽ và upload lên trang Padlet của nhóm."
    )

    add_bullet("Báo cáo, chia sẻ & Hoàn thiện bản vẽ (7 phút)", "4. Hoạt động 4: ")
    p_a4 = doc.add_paragraph()
    p_a4.paragraph_format.left_indent = Inches(0.3)
    p_a4.add_run(
        "• GV chiếu màn hình Padlet hiển thị bản vẽ của 3 nhóm.\n"
        "• Đại diện mỗi nhóm thuyết minh 60 giây về điểm độc đáo trong kết cấu của nhóm mình.\n"
        "• GV nhận xét, lưu ý kỹ thuật bịt kín các khe hở mép tường và gioăng cửa sổ trước khi bước sang tiết 2."
    )

    add_heading2("TIẾT 2 (45 PHÚT): CHẾ TẠO, ĐO LƯỜNG SỐ ĐỊNH LƯỢNG & BÁO CÁO TỔNG KẾT")

    add_bullet("Thi công, chế tạo mô hình (15 phút)", "5. Hoạt động 5: ")
    p_a5 = doc.add_paragraph()
    p_a5.paragraph_format.left_indent = Inches(0.3)
    p_a5.add_run(
        "• Các nhóm nhận khay vật tư: Tấm Fomex khung, Tấm xốp, Giấy khay trứng gà, keo dính, dao rọc giấy an toàn, thước đo.\n"
        "• Tiến hành ốp vật liệu vào khung Fomex theo đúng bản vẽ thiết kế đã duyệt.\n"
        "• Gia công gioăng cửa ra vào và cửa sổ bằng dải mút xốp mỏng để chống hiện tượng \"rò rỉ âm học\"."
    )

    add_bullet("Thử nghiệm tại \"Trạm kiểm định âm học số\" (15 phút)", "6. Hoạt động 6: ")
    p_a6 = doc.add_paragraph()
    p_a6.paragraph_format.left_indent = Inches(0.3)
    p_a6.add_run(
        "• Thiết lập Trạm kiểm định chuẩn tại bục giảng: 01 Loa phát tiếng ồn đô thị liên tục ở mức cố định 80 dB (khoảng cách 30 cm).\n"
        "• Quy trình đo kiểm định lần lượt 3 nhóm:\n"
        "  + Bước 1: Mở app Decibel X (chế độ giữ đỉnh Max và Avg), đặt điện thoại vào bên trong nhà, hướng mic về phía cửa.\n"
        "  + Bước 2: Đóng kín cửa nhà, bắt đầu bật loa tiếng ồn 80 dB trong 15 giây.\n"
        "  + Bước 3: Mở cửa, chụp lại màn hình chỉ số âm trong nhà (L_trong).\n"
        "  + Bước 4: Nhập ngay số liệu vào bảng tổng hợp Google Sheets của lớp chiếu trên màn hình.\n"
        "• Công thức tính mức giảm ồn định lượng: ΔL = 80 dB - L_trong (dB)."
    )

    add_bullet("Báo cáo số liệu, Đối chiếu 3 mô hình & Đánh giá Rubric (10 phút)", "7. Hoạt động 7: ")
    p_a7 = doc.add_paragraph()
    p_a7.paragraph_format.left_indent = Inches(0.3)
    p_a7.add_run(
        "• Chiếu bảng so sánh dữ liệu thực tế của 3 nhóm (Kỳ vọng: Nhóm 1 giảm ~18 dB; Nhóm 2 giảm ~22 dB; Nhóm 3 giảm ~24 dB).\n"
        "• Thảo luận khoa học: Vì sao giải pháp kết hợp đa tầng hoặc vách kép lại vượt trội hơn hẳn so với vách đơn?\n"
        "• GV và học sinh đánh giá chéo theo Bảng tiêu chí Rubric."
    )

    add_bullet("Kết luận, nhận định & Mở rộng thực tiễn (5 phút)", "8. Hoạt động 8: ")
    p_a8 = doc.add_paragraph()
    p_a8.paragraph_format.left_indent = Inches(0.3)
    p_a8.add_run(
        "• GV hệ thống hóa 3 nguyên tắc chống ô nhiễm tiếng ồn trong thực tiễn xây dựng:\n"
        "  (1) Ngăn chặn đường truyền âm: Dùng tường Fomex/bê tông dày, kính hộp cách âm 2 lớp.\n"
        "  (2) Hấp thụ và tiêu âm: Ốp mút xốp, rèm cửa dày, thảm len.\n"
        "  (3) Phân tán âm: Bề mặt zíc zắc tán âm (khay trứng/gỗ tiêu âm), trồng dải cây xanh quanh nhà.\n"
        "• Giao nhiệm vụ số về nhà: Dùng điện thoại đo tiếng ồn tại phòng ngủ của em và đề xuất 1 cách cải tạo đơn giản."
    )

    # PHẦN 4
    add_heading1("IV. CÔNG CỤ DẠY HỌC SỐ & BẢNG RUBRIC ĐÁNH GIÁ ĐẦU RA")

    add_heading2("1. Hệ sinh thái Dạy học số tích hợp trong bài dạy")
    add_bullet("Decibel X / Phyphox (Cài trên Smartphone): Biến điện thoại thành thiết bị đo cường độ âm thanh chuẩn xác, hiển thị biểu đồ sóng thời gian thực.", "• Ứng dụng đo lường số: ")
    add_bullet("Padlet lớp học: Không gian nộp bản vẽ thiết kế, hình ảnh thi công và nhật ký thí nghiệm của 3 nhóm.", "• Không gian số tương tác: ")
    add_bullet("Google Sheets: Bảng tính tự động tính toán ΔL và vẽ biểu đồ so sánh hiệu quả cách âm trực quan giữa các nhóm.", "• Xử lý dữ liệu số: ")
    add_bullet("Video clip tình huống mở đầu, slide bài giảng điện tử tương tác Canva.", "• Học liệu số dùng chung: ")

    add_heading2("2. Bảng Rubric đánh giá sản phẩm và năng lực học sinh")

    table_rub = doc.add_table(rows=5, cols=5)
    table_rub.alignment = WD_TABLE_ALIGNMENT.CENTER
    rub_widths = [Inches(1.2), Inches(1.3), Inches(1.4), Inches(1.4), Inches(1.4)]
    
    rub_headers = ["Tiêu chí", "Trọng số", "Mức Đạt (5 - 6đ)", "Mức Khá (7 - 8đ)", "Mức Tốt (9 - 10đ)"]
    for i, h in enumerate(rub_headers):
        cell = table_rub.cell(0, i)
        cell.width = rub_widths[i]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, top=100, bottom=100, left=100, right=100)
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(10)

    rub_data = [
        ("1. Hiệu quả cách âm định lượng (ΔL)", "40%", 
         "Giảm được từ 10 - 14 dB so với nguồn âm ngoài.", 
         "Giảm được từ 15 - 19 dB so với nguồn âm ngoài.", 
         "Giảm được ≥ 20 dB (hiệu quả cách âm vượt trội)."),
        ("2. Bản vẽ kỹ thuật & Tính toán", "20%", 
         "Bản vẽ sơ lược, thiếu thông số kích thước và tính toán diện tích.", 
         "Bản vẽ rõ ràng, có phân lớp vật liệu, tính toán tương đối đủ.", 
         "Bản vẽ mặt cắt tỉ mỉ, tính toán diện tích tối ưu, chú thích vật liệu chi tiết."),
        ("3. Kết cấu mô hình & Thẩm mỹ", "20%", 
         "Khung Fomex lỏng lẻo, dán keo lem nhem, mép cửa bị hở nhiều.", 
         "Khung nhà vững chắc, các mép Fomex và mút khít tương đối, có cửa đóng mở.", 
         "Khung nhà vuông vắn, thẩm mỹ cao, cửa kín khít có gioăng xốp, phối hợp vật liệu sáng tạo."),
        ("4. Kỹ năng Dạy học số & Báo cáo", "20%", 
         "Thao tác app đo đạc chậm, cần GV hỗ trợ nhiều, nộp Padlet trễ.", 
         "Sử dụng thành thạo app đo decibel, nộp bản vẽ lên Padlet đúng hạn.", 
         "Thao tác số rất nhanh nhẹn, số liệu chụp rõ nét, thuyết trình mạch lạc, tự tin.")
    ]

    for row_idx, data in enumerate(rub_data, start=1):
        bg = "FFFFFF" if row_idx % 2 != 0 else "F7FAFC"
        for col_idx, text in enumerate(data):
            cell = table_rub.cell(row_idx, col_idx)
            cell.width = rub_widths[col_idx]
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.size = Pt(9.5)
            if col_idx in [0, 1]:
                r.bold = True

    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # Footer note
    p_foot = doc.add_paragraph()
    p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_f = p_foot.add_run("Kế hoạch bài dạy được biên soạn theo định hướng Chương trình GDPT 2018 và Kế hoạch 12070/KH-SGDĐT TP.HCM")
    r_f.italic = True
    r_f.font.size = Pt(9.5)
    r_f.font.color.rgb = COLOR_MUTED

    doc.save(filename)
    print(f"File created successfully: {filename}")

if __name__ == '__main__':
    create_demo_doc("E:\\4. THIET KE WEB\\Cô Thúy\\Y_Tuong_STEM_Ngoi_Nha_Cach_Am_KHTN7.docx")
