import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from generate_measured_worksheets import stations

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="000000", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def build_student_bw_worksheet_single_page(st, filename):
    doc = docx.Document()
    
    # Page setup: A4 with compact margins (0.4 in top/bottom, 0.5 in left/right) to guarantee 1-page fit
    for section in doc.sections:
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    # 1. School Header (Pure Black)
    p_hdr = doc.add_paragraph()
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(2)
    p_hdr.paragraph_format.line_spacing = 1.05
    
    r = p_hdr.add_run("TRƯỜNG THCS TÂN TẠO A — TỔ KHOA HỌC TỰ NHIÊN\n")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_hdr.add_run("DỰ ÁN STEM SINH HỌC 7: XƯỞNG NGHIÊN CỨU VI KHÍ HẬU & KHÔNG GIAN XANH TP.HCM")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)

    # 2. Main Title Box (Pure White Background with Crisp Black Border - No Color)
    t_box = doc.add_table(rows=1, cols=1)
    t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t_box.rows[0].cells[0]
    cell.width = Inches(7.27)
    set_cell_background(cell, "FFFFFF")
    set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
    set_table_borders(t_box, color="000000", sz="10", val="single")
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.05
    
    r = p.add_run(f"PHIẾU HỌC TẬP THỰC NGHIỆM SỐ {st['code']} — DÀNH CHO {st['group']}\n")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p.add_run(f"TIẾN TRÌNH ĐO LƯỜNG VI KHÍ HẬU & HIỆU QUẢ QUANG HỢP (BÀI 23 KHTN 7)\n")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p.add_run(f"{st['title']} ({st['subtitle']})")
    r3.font.name = "Times New Roman"
    r3.font.bold = True
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(0, 0, 0)

    # 3. Student Identification Line (Pure Black)
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(4)
    p_info.paragraph_format.space_after = Pt(3)
    p_info.paragraph_format.line_spacing = 1.05
    
    r = p_info.add_run("Tên nhóm: ..................................................... Lớp: 7/......  •  GV hướng dẫn: ")
    r.font.name = "Times New Roman"
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_info.add_run("Cô Mai Thuý Đức")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_info.add_run("  •  Năm học: 2026 - 2027\n")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(9)
    r3.font.color.rgb = RGBColor(0, 0, 0)
    
    r4 = p_info.add_run("Thành viên nhóm: 1............................................... 2............................................... 3............................................... 4...............................................")
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(8.5)
    r4.font.color.rgb = RGBColor(0, 0, 0)

    # 4. Section I: Context & Mission (Pure Black)
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(3)
    p_sec1.paragraph_format.space_after = Pt(1)
    r = p_sec1.add_run("I. KHẢO SÁT HIỆN TRẠNG & THÁCH THỨC VI KHÍ HẬU TẠI TRẠM:")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(0)
    p_desc.paragraph_format.space_after = Pt(3)
    p_desc.paragraph_format.line_spacing = 1.05
    
    r = p_desc.add_run("• Thách thức vi khí hậu: ")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_desc.add_run(st['challenge'] + "\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_desc.add_run("• Nhiệm vụ thực nghiệm số: ")
    r3.font.name = "Times New Roman"
    r3.font.bold = True
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = RGBColor(0, 0, 0)
    
    r4 = p_desc.add_run(st['mission'])
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(8.5)
    r4.font.color.rgb = RGBColor(0, 0, 0)

    # 5. Section II: Telemetry Table (Clean Blank for Student Handwriting)
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(3)
    p_sec2.paragraph_format.space_after = Pt(2)
    r = p_sec2.add_run("II. BẢNG GHI LẠI THÔNG SỐ ĐO LƯỜNG KHI CHO TỪNG GIẢI PHÁP VÀO:")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r_sub = p_sec2.add_run(" (Học sinh đo đạc trực tiếp trên mô hình mô phỏng Module 3 và ghi vào bảng)")
    r_sub.font.name = "Times New Roman"
    r_sub.font.italic = True
    r_sub.font.size = Pt(8)
    r_sub.font.color.rgb = RGBColor(0, 0, 0)

    col_widths = [Inches(0.95), Inches(1.85), Inches(0.75), Inches(0.80), Inches(0.65), Inches(0.80), Inches(1.47)]
    table = doc.add_table(rows=6, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="000000", sz="6", val="single")

    headers = [
        "Bước thực nghiệm",
        "Giải pháp kỹ thuật đưa vào\n(Kéo thả trên web)",
        "Nồng độ CO₂\n(ppm)",
        "Ánh sáng\n(Lux)",
        "Nhiệt độ\n(°C)",
        "Hiệu quả quang hợp",
        "Hiện tượng quan sát\n(Trạng thái lá & Khí khổng)"
    ]

    hdr_row = table.rows[0]
    for c_idx, title in enumerate(headers):
        c = hdr_row.cells[c_idx]
        c.width = col_widths[c_idx]
        set_cell_background(c, "F2F2F2")
        set_cell_margins(c, top=60, bottom=60, left=40, right=40)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(0, 0, 0)

    step_labels = [
        "Ban đầu\n(Hiện trạng)",
        "Bước 1\n(+ 1 Giải pháp)",
        "Bước 2\n(+ 2 Giải pháp)",
        "Bước 3\n(+ 3 Giải pháp)",
        "Bước 4\n(+ 4 Giải pháp)"
    ]

    for step_idx in range(5):
        row = table.rows[step_idx + 1]
        for c_idx in range(7):
            c = row.cells[c_idx]
            c.width = col_widths[c_idx]
            set_cell_background(c, "FFFFFF")
            set_cell_margins(c, top=100, bottom=100, left=50, right=50) # Generous height for handwriting
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(step_labels[step_idx])
                r.font.name = "Times New Roman"
                r.font.bold = True
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(0, 0, 0)
            elif c_idx == 1 and step_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run("Chưa kích hoạt giải pháp nào")
                r.font.name = "Times New Roman"
                r.font.italic = True
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(0, 0, 0)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 6. Section III: Delta Improvement (Pure Black, Compact)
    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.space_before = Pt(5)
    p_sec3.paragraph_format.space_after = Pt(2)
    r = p_sec3.add_run("III. ĐÁNH GIÁ MỨC ĐỘ CẢI THIỆN VI KHÍ HẬU (Δ SAU KHI HOÀN TẤT - BAN ĐẦU):")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_delta = doc.add_paragraph()
    p_delta.paragraph_format.space_before = Pt(0)
    p_delta.paragraph_format.space_after = Pt(3)
    p_delta.paragraph_format.line_spacing = 1.15
    
    r = p_delta.add_run("• Nồng độ CO₂: Thay đổi từ ........................... ppm  →  ........................... ppm  (Tăng/Giảm: ....................................................)\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_delta.add_run("• Cường độ ánh sáng: Thay đổi từ ........................... Lux  →  ........................... Lux  (Tăng/Giảm: ............................................)\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_delta.add_run("• Nhiệt độ vi khí hậu: Thay đổi từ ........................... °C  →  ........................... °C  (Tăng/Giảm: ...................................................)\n")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = RGBColor(0, 0, 0)
    
    r4 = p_delta.add_run("• Hiệu quả quang hợp: Tăng vọt từ ........................... %  →  ........................... %  (Trạng thái lá: ......................................................)")
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(8.5)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(0, 0, 0)

    # (SECTION IV REMOVED AS REQUESTED!)

    # 7. Evaluation & Signatures (Pure Black)
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(8)
    p_sign.paragraph_format.space_after = Pt(0)
    r = p_sign.add_run("ĐÁNH GIÁ CỦA NHÓM: ...................................................          NHẬN XÉT & CHỮ KÝ GIÁO VIÊN: ")
    r.font.name = "Times New Roman"
    r.font.size = Pt(8.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_sign.add_run("Cô Mai Thuý Đức")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)

    doc.save(filename)
    print(f"Created 1-page BW student worksheet without Section IV: {filename}")

def main():
    file_configs = [
        (stations[0], "Phieu_Hoc_Tap_2A_Nhom_1_In_Phat_Hoc_Sinh.docx"),
        (stations[1], "Phieu_Hoc_Tap_2B_Nhom_2_In_Phat_Hoc_Sinh.docx"),
        (stations[2], "Phieu_Hoc_Tap_2C_Nhom_3_In_Phat_Hoc_Sinh.docx"),
        (stations[3], "Phieu_Hoc_Tap_2D_Nhom_4_In_Phat_Hoc_Sinh.docx"),
    ]
    for st, fn in file_configs:
        build_student_bw_worksheet_single_page(st, fn)

if __name__ == "__main__":
    main()
