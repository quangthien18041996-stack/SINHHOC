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

def set_cell_margins(cell, top=100, bottom=100, left=100, right=100):
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

def build_student_bw_worksheet(st, filename):
    doc = docx.Document()
    
    # Page setup: A4 with standard 0.5 inch margins for clean 2-page double-sided printout
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)

    # 1. School Header (Pure Black)
    p_hdr = doc.add_paragraph()
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(2)
    
    r = p_hdr.add_run("TRƯỜNG THCS TÂN TẠO A — TỔ KHOA HỌC TỰ NHIÊN\n")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_hdr.add_run("DỰ ÁN STEM SINH HỌC 7: XƯỞNG NGHIÊN CỨU VI KHÍ HẬU & KHÔNG GIAN XANH TP.HCM")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(0, 0, 0)

    # 2. Main Title Box (Pure White Background with Double/Thick Black Border - No Dark Shading)
    t_box = doc.add_table(rows=1, cols=1)
    t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t_box.rows[0].cells[0]
    cell.width = Inches(7.17)
    set_cell_background(cell, "FFFFFF") # Pure white, no ink waste
    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
    set_table_borders(t_box, color="000000", sz="12", val="single") # Solid 1.5pt crisp black box
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    
    r = p.add_run(f"PHIẾU HỌC TẬP THỰC NGHIỆM SỐ {st['code']} — DÀNH CHO {st['group']}\n")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p.add_run(f"TIẾN TRÌNH ĐO LƯỜNG VI KHÍ HẬU & HIỆU QUẢ QUANG HỢP (BÀI 23 KHTN 7)\n")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p.add_run(f"{st['title']} ({st['subtitle']})")
    r3.font.name = "Times New Roman"
    r3.font.bold = True
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(0, 0, 0)

    # 3. Student Identification Line (Pure Black)
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(6)
    p_info.paragraph_format.space_after = Pt(4)
    
    r = p_info.add_run("Tên nhóm: ..................................................... Lớp: 7/......  •  GV hướng dẫn: ")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_info.add_run("Cô Mai Thuý Đức")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_info.add_run("  •  Năm học: 2026 - 2027\n")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(0, 0, 0)
    
    r4 = p_info.add_run("Thành viên nhóm: 1............................................... 2............................................... 3............................................... 4...............................................")
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(9)
    r4.font.color.rgb = RGBColor(0, 0, 0)

    # 4. Section I: Context & Mission (Pure Black)
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(4)
    p_sec1.paragraph_format.space_after = Pt(2)
    r = p_sec1.add_run("I. KHẢO SÁT HIỆN TRẠNG & THÁCH THỨC VI KHÍ HẬU TẠI TRẠM:")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(0)
    p_desc.paragraph_format.space_after = Pt(4)
    
    r = p_desc.add_run("• Thách thức vi khí hậu: ")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_desc.add_run(st['challenge'] + "\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_desc.add_run("• Nhiệm vụ thực nghiệm số: ")
    r3.font.name = "Times New Roman"
    r3.font.bold = True
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(0, 0, 0)
    
    r4 = p_desc.add_run(st['mission'])
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(9.5)
    r4.font.color.rgb = RGBColor(0, 0, 0)

    # 5. Section II: Telemetry & Measurement Table (Pure Black & White for Handwriting)
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(5)
    p_sec2.paragraph_format.space_after = Pt(3)
    r = p_sec2.add_run("II. BẢNG GHI LẠI THÔNG SỐ ĐO LƯỜNG KHI CHO TỪNG GIẢI PHÁP VÀO:")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r_sub = p_sec2.add_run(" (Học sinh đo đạc trực tiếp trên mô hình mô phỏng Module 3 và dùng bút ghi vào bảng)")
    r_sub.font.name = "Times New Roman"
    r_sub.font.italic = True
    r_sub.font.size = Pt(9)
    r_sub.font.color.rgb = RGBColor(0, 0, 0)

    col_widths = [Inches(1.0), Inches(1.85), Inches(0.75), Inches(0.80), Inches(0.65), Inches(0.80), Inches(1.32)]
    table = doc.add_table(rows=6, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="000000", sz="6", val="single") # Crisp black borders

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
        set_cell_background(c, "F2F2F2") # Very light 5% gray, photocopies super cleanly without ink blotch
        set_cell_margins(c, top=80, bottom=80, left=50, right=50)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(0, 0, 0) # Black text

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
            set_cell_background(c, "FFFFFF") # 100% Pure White
            set_cell_margins(c, top=140, bottom=140, left=60, right=60) # Ample handwriting height!
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

            if c_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(step_labels[step_idx])
                r.font.name = "Times New Roman"
                r.font.bold = True
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(0, 0, 0)
            elif c_idx == 1 and step_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run("Chưa kích hoạt giải pháp nào")
                r.font.name = "Times New Roman"
                r.font.italic = True
                r.font.size = Pt(8.5)
                r.font.color.rgb = RGBColor(0, 0, 0)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Left completely blank for student to write with pen

    # 6. Section III: Delta Improvement (Pure Black)
    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.space_before = Pt(8)
    p_sec3.paragraph_format.space_after = Pt(3)
    r = p_sec3.add_run("III. ĐÁNH GIÁ MỨC ĐỘ CẢI THIỆN VI KHÍ HẬU (Δ SAU KHI HOÀN TẤT - BAN ĐẦU):")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_delta = doc.add_paragraph()
    p_delta.paragraph_format.space_before = Pt(0)
    p_delta.paragraph_format.space_after = Pt(4)
    r = p_delta.add_run("• Nồng độ CO₂: Thay đổi từ ........................... ppm  →  ........................... ppm  (Tăng/Giảm: ...........................)\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_delta.add_run("• Cường độ ánh sáng: Thay đổi từ ........................... Lux  →  ........................... Lux  (Tăng/Giảm: ...........................)\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_delta.add_run("• Nhiệt độ vi khí hậu: Thay đổi từ ........................... °C  →  ........................... °C  (Tăng/Giảm: ...........................)\n")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(9.5)
    r3.font.color.rgb = RGBColor(0, 0, 0)
    
    r4 = p_delta.add_run("• Hiệu quả quang hợp: Tăng vọt từ ........................... %  →  ........................... %  (Trạng thái lá: ......................................................)")
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(9.5)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor(0, 0, 0)

    # 7. Section IV: Scientific Questions (Pure Black)
    p_sec4 = doc.add_paragraph()
    p_sec4.paragraph_format.space_before = Pt(6)
    p_sec4.paragraph_format.space_after = Pt(2)
    r = p_sec4.add_run("IV. CÂU HỎI KHOA HỌC & RÚT RA QUY LUẬT SINH HỌC (BÁO CÁO E-POSTER):")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(0, 0, 0)

    p_q1 = doc.add_paragraph()
    p_q1.paragraph_format.space_before = Pt(2)
    p_q1.paragraph_format.space_after = Pt(2)
    r = p_q1.add_run("1. Câu hỏi 1: ")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_q1.add_run(st['q1'] + "\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_q1.add_run("Trả lời: ...........................................................................................................................................................................................................................................\n..........................................................................................................................................................................................................................................................")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(9)
    r3.font.color.rgb = RGBColor(0, 0, 0)

    p_q2 = doc.add_paragraph()
    p_q2.paragraph_format.space_before = Pt(5)
    p_q2.paragraph_format.space_after = Pt(3)
    r = p_q2.add_run("2. Câu hỏi 2: ")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_q2.add_run(st['q2'] + "\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(0, 0, 0)
    
    r3 = p_q2.add_run("Trả lời: ...........................................................................................................................................................................................................................................\n..........................................................................................................................................................................................................................................................")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(9)
    r3.font.color.rgb = RGBColor(0, 0, 0)

    # 8. Evaluation & Signatures (Pure Black)
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(8)
    p_sign.paragraph_format.space_after = Pt(0)
    r = p_sign.add_run("ĐÁNH GIÁ CỦA NHÓM: ...................................................          NHẬN XÉT & CHỮ KÝ GIÁO VIÊN: ")
    r.font.name = "Times New Roman"
    r.font.size = Pt(9)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0, 0, 0)
    
    r2 = p_sign.add_run("Cô Mai Thuý Đức")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(9)
    r2.font.color.rgb = RGBColor(0, 0, 0)

    doc.save(filename)
    print(f"Created black-and-white student worksheet: {filename}")

def main():
    file_configs = [
        (stations[0], "Phieu_Hoc_Tap_2A_Nhom_1_In_Phat_Hoc_Sinh.docx"),
        (stations[1], "Phieu_Hoc_Tap_2B_Nhom_2_In_Phat_Hoc_Sinh.docx"),
        (stations[2], "Phieu_Hoc_Tap_2C_Nhom_3_In_Phat_Hoc_Sinh.docx"),
        (stations[3], "Phieu_Hoc_Tap_2D_Nhom_4_In_Phat_Hoc_Sinh.docx"),
    ]
    for st, fn in file_configs:
        build_student_bw_worksheet(st, fn)

if __name__ == "__main__":
    main()
