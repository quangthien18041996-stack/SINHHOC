import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from generate_measured_worksheets import stations

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="9CA3AF", sz="4", val="single"):
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

def create_single_worksheet(st, filename):
    doc = docx.Document()
    
    # Page setup: A4 with compact 0.5 inch margins for clean 2-page printout
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)
        section.page_width = Inches(8.27)  # A4 width
        section.page_height = Inches(11.69) # A4 height

    # 1. School Header
    p_hdr = doc.add_paragraph()
    p_hdr.paragraph_format.space_before = Pt(0)
    p_hdr.paragraph_format.space_after = Pt(2)
    
    r = p_hdr.add_run("TRƯỜNG THCS TÂN TẠO A — TỔ KHOA HỌC TỰ NHIÊN\n")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(75, 85, 99)
    
    r2 = p_hdr.add_run("DỰ ÁN STEM SINH HỌC 7: XƯỞNG NGHIÊN CỨU VI KHÍ HẬU & KHÔNG GIAN XANH TP.HCM")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(10)
    r2.font.color.rgb = RGBColor(16, 122, 80)

    # 2. Main Title Banner Box
    t_box = doc.add_table(rows=1, cols=1)
    t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t_box.rows[0].cells[0]
    cell.width = Inches(7.17)
    set_cell_background(cell, "064E3B") # Emerald 900
    set_cell_margins(cell, top=120, bottom=120, left=160, right=160)
    
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    
    r = p.add_run(f"PHIẾU HỌC TẬP THỰC NGHIỆM SỐ {st['code']} — DÀNH CHO {st['group']}\n")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(12.5)
    r.font.color.rgb = RGBColor(255, 255, 255)
    
    r2 = p.add_run(f"TIẾN TRÌNH ĐO LƯỜNG VI KHÍ HẬU & HIỆU QUẢ QUANG HỢP (BÀI 23 KHTN 7)\n")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(167, 243, 208)
    
    r3 = p.add_run(f"{st['title']} ({st['subtitle']})")
    r3.font.name = "Times New Roman"
    r3.font.bold = True
    r3.font.size = Pt(10)
    r3.font.color.rgb = RGBColor(253, 230, 138)

    # 3. Student Identification Line
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(6)
    p_info.paragraph_format.space_after = Pt(5)
    
    r = p_info.add_run("Tên nhóm: ..................................................... Lớp: 7/......  •  GV hướng dẫn: ")
    r.font.name = "Times New Roman"
    r.font.size = Pt(9.5)
    
    r2 = p_info.add_run("Cô Mai Thuý Đức")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(9.5)
    r2.font.color.rgb = RGBColor(16, 122, 80)
    
    r3 = p_info.add_run("  •  Năm học: 2026 - 2027\n")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(9.5)
    
    r4 = p_info.add_run("Thành viên nhóm: 1............................................... 2............................................... 3............................................... 4...............................................")
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(8.5)
    r4.font.color.rgb = RGBColor(75, 85, 99)

    # 4. Section I: Context & Mission
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_before = Pt(4)
    p_sec1.paragraph_format.space_after = Pt(2)
    r = p_sec1.add_run("I. KHẢO SÁT HIỆN TRẠNG & THÁCH THỨC VI KHÍ HẬU TẠI TRẠM:")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(27, 67, 50)

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.space_before = Pt(0)
    p_desc.paragraph_format.space_after = Pt(4)
    
    r = p_desc.add_run("• Thách thức vi khí hậu: ")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9)
    
    r2 = p_desc.add_run(st['challenge'] + "\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(9)
    
    r3 = p_desc.add_run("• Nhiệm vụ thực nghiệm số: ")
    r3.font.name = "Times New Roman"
    r3.font.bold = True
    r3.font.size = Pt(9)
    r3.font.color.rgb = RGBColor(16, 122, 80)
    
    r4 = p_desc.add_run(st['mission'])
    r4.font.name = "Times New Roman"
    r4.font.size = Pt(9)

    # 5. Section II: Telemetry & Measurement Table (7 Columns, 5 Steps)
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_before = Pt(5)
    p_sec2.paragraph_format.space_after = Pt(3)
    r = p_sec2.add_run("II. BẢNG GHI LẠI THÔNG SỐ ĐO LƯỜNG KHI CHO TỪNG GIẢI PHÁP VÀO:")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(27, 67, 50)
    
    r_sub = p_sec2.add_run(" (Học sinh đo đạc và quan sát trực tiếp trên mô hình mô phỏng Module 3)")
    r_sub.font.name = "Times New Roman"
    r_sub.font.italic = True
    r_sub.font.size = Pt(8.5)
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    col_widths = [Inches(0.9), Inches(1.85), Inches(0.78), Inches(0.85), Inches(0.64), Inches(0.80), Inches(1.35)]
    table = doc.add_table(rows=6, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(table, color="9CA3AF", sz="4")

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
        set_cell_background(c, "1E3A8A") # Navy Blue 900
        set_cell_margins(c, top=70, bottom=70, left=50, right=50)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(title)
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(8)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for step_idx, step_data in enumerate(st['steps']):
        row = table.rows[step_idx + 1]
        bg_color = "F9FAFB" if step_idx % 2 == 0 else "FFFFFF"
        if step_idx == 0:
            bg_color = "FEF2F2" # Red tint for initial baseline
        elif step_idx == 4:
            bg_color = "ECFDF5" # Green tint for optimal 100%

        for c_idx in range(7):
            c = row.cells[c_idx]
            c.width = col_widths[c_idx]
            set_cell_background(c, bg_color)
            set_cell_margins(c, top=60, bottom=60, left=60, right=60)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            
            if c_idx in [0, 2, 3, 4, 5]:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

            val = step_data[c_idx]
            r = p.add_run(val)
            r.font.name = "Times New Roman"
            r.font.size = Pt(8)
            if c_idx == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(31, 41, 55)
            elif c_idx == 1:
                r.font.bold = (step_idx == 4)
                r.font.color.rgb = RGBColor(17, 24, 39)
            elif c_idx == 5:
                r.font.bold = True
                r.font.color.rgb = RGBColor(220, 38, 38) if step_idx == 0 else RGBColor(16, 122, 80)
            elif c_idx == 6:
                r.font.italic = True
                r.font.size = Pt(7.5)
                r.font.color.rgb = RGBColor(55, 65, 81)

    # 6. Section III: Delta Improvement
    p_sec3 = doc.add_paragraph()
    p_sec3.paragraph_format.space_before = Pt(6)
    p_sec3.paragraph_format.space_after = Pt(2)
    r = p_sec3.add_run("III. ĐÁNH GIÁ MỨC ĐỘ CẢI THIỆN VI KHÍ HẬU (Δ SAU KHI HOÀN TẤT - BAN ĐẦU):")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(27, 67, 50)

    p_delta = doc.add_paragraph()
    p_delta.paragraph_format.space_before = Pt(0)
    p_delta.paragraph_format.space_after = Pt(4)
    r = p_delta.add_run(st['delta'])
    r.font.name = "Times New Roman"
    r.font.size = Pt(8.5)
    r.font.color.rgb = RGBColor(4, 120, 87)

    # 7. Section IV: Scientific Questions
    p_sec4 = doc.add_paragraph()
    p_sec4.paragraph_format.space_before = Pt(5)
    p_sec4.paragraph_format.space_after = Pt(2)
    r = p_sec4.add_run("IV. CÂU HỎI KHOA HỌC & RÚT RA QUY LUẬT SINH HỌC (BÁO CÁO E-POSTER):")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = RGBColor(27, 67, 50)

    p_q1 = doc.add_paragraph()
    p_q1.paragraph_format.space_before = Pt(2)
    p_q1.paragraph_format.space_after = Pt(2)
    r = p_q1.add_run("1. Câu hỏi 1: ")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(180, 83, 9)
    
    r2 = p_q1.add_run(st['q1'] + "\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(9)
    
    r3 = p_q1.add_run("Trả lời: ...........................................................................................................................................................................................................................................\n..........................................................................................................................................................................................................................................................")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = RGBColor(107, 114, 128)

    p_q2 = doc.add_paragraph()
    p_q2.paragraph_format.space_before = Pt(4)
    p_q2.paragraph_format.space_after = Pt(3)
    r = p_q2.add_run("2. Câu hỏi 2: ")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(180, 83, 9)
    
    r2 = p_q2.add_run(st['q2'] + "\n")
    r2.font.name = "Times New Roman"
    r2.font.size = Pt(9)
    
    r3 = p_q2.add_run("Trả lời: ...........................................................................................................................................................................................................................................\n..........................................................................................................................................................................................................................................................")
    r3.font.name = "Times New Roman"
    r3.font.size = Pt(8.5)
    r3.font.color.rgb = RGBColor(107, 114, 128)

    # 8. Evaluation & Signatures
    p_sign = doc.add_paragraph()
    p_sign.paragraph_format.space_before = Pt(8)
    p_sign.paragraph_format.space_after = Pt(0)
    
    r = p_sign.add_run("ĐÁNH GIÁ CỦA NHÓM: ...................................................          NHẬN XÉT & CHỮ KÝ GIÁO VIÊN: ")
    r.font.name = "Times New Roman"
    r.font.size = Pt(8.5)
    r.font.bold = True
    
    r2 = p_sign.add_run("Cô Mai Thuý Đức")
    r2.font.name = "Times New Roman"
    r2.font.bold = True
    r2.font.size = Pt(8.5)
    r2.font.color.rgb = RGBColor(16, 122, 80)

    doc.save(filename)
    print(f"Created: {filename}")

def main():
    file_configs = [
        (stations[0], "Phieu_Hoc_Tap_So_2A_Nhom_1_Tram_1.docx"),
        (stations[1], "Phieu_Hoc_Tap_So_2B_Nhom_2_Tram_2.docx"),
        (stations[2], "Phieu_Hoc_Tap_So_2C_Nhom_3_Tram_3.docx"),
        (stations[3], "Phieu_Hoc_Tap_So_2D_Nhom_4_Tram_4.docx"),
    ]
    for st, fn in file_configs:
        create_single_worksheet(st, fn)

if __name__ == "__main__":
    main()
