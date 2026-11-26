import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=60, bottom=60, left=50, right=50):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_cell_valign(cell, valign="center"):
    tcPr = cell._tc.get_or_add_tcPr()
    vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{valign}"/>')
    tcPr.append(vAlign)

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

def remove_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>
            <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

from stations_data import stations

def build_all_worksheets(filename="Phieu_Hoc_Tap_KHTN7_Bai23_Quang_Hop_4_Tram.docx"):
    doc = docx.Document()
    for s_idx, st in enumerate(stations):
        if s_idx > 0:
            doc.add_page_break()

        section = doc.sections[0] if s_idx == 0 else doc.sections[-1]
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.55)

        # 1. Header (2 cols borderless)
        hdr_table = doc.add_table(rows=1, cols=2)
        hdr_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(hdr_table)
        
        c1 = hdr_table.rows[0].cells[0]
        c1.width = Inches(3.2)
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        p1.paragraph_format.line_spacing = 1.05
        r = p1.add_run("TRƯỜNG THCS TÂN TẠO A\nTỔ: KHOA HỌC TỰ NHIÊN\n-----------***-----------")
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(10)

        c2 = hdr_table.rows[0].cells[1]
        c2.width = Inches(3.77)
        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        p2.paragraph_format.line_spacing = 1.05
        r = p2.add_run("DỰ ÁN STEM SINH HỌC 7\nBÀI 23: QUANG HỢP Ở THỰC VẬT\nNăm học: 2026 – 2027")
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(10)

        # 2. Title
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(6)
        p_title.paragraph_format.space_after = Pt(1)
        p_title.paragraph_format.line_spacing = 1.1
        r = p_title.add_run(f"PHIẾU HỌC TẬP THỰC NGHIỆM SỐ {st['code']}\n")
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(13)
        r2 = p_title.add_run(f"DÀNH CHO {st['group'].upper()} — {st['title'].upper()}\n")
        r2.font.name = "Times New Roman"
        r2.font.bold = True
        r2.font.size = Pt(10.5)
        r3 = p_title.add_run(f"({st['subtitle']})")
        r3.font.name = "Times New Roman"
        r3.font.italic = True
        r3.font.size = Pt(9.5)

        # 3. Student info
        p_info = doc.add_paragraph()
        p_info.paragraph_format.space_before = Pt(4)
        p_info.paragraph_format.space_after = Pt(3)
        p_info.paragraph_format.line_spacing = 1.15
        r = p_info.add_run("Tên nhóm: .............................................................. Lớp: 7/......  •  GV hướng dẫn: ")
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r2 = p_info.add_run("Cô Mai Thuý Đức\n")
        r2.font.name = "Times New Roman"
        r2.font.bold = True
        r2.font.size = Pt(9.5)
        r3 = p_info.add_run("Thành viên: ....................................................................................................................................................................\n..................................................................................................................................................................................................")
        r3.font.name = "Times New Roman"
        r3.font.size = Pt(9)

        # 4. Section I
        p_sec1 = doc.add_paragraph()
        p_sec1.paragraph_format.space_before = Pt(3)
        p_sec1.paragraph_format.space_after = Pt(1)
        r = p_sec1.add_run("I. KHẢO SÁT HIỆN TRẠNG & THÁCH THỨC VI KHÍ HẬU TẠI TRẠM:")
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(10)

        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(0)
        p_desc.paragraph_format.space_after = Pt(3)
        p_desc.paragraph_format.line_spacing = 1.1
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
        r4 = p_desc.add_run(st['mission'])
        r4.font.name = "Times New Roman"
        r4.font.size = Pt(9)

        # 5. Section II
        p_sec2 = doc.add_paragraph()
        p_sec2.paragraph_format.space_before = Pt(3)
        p_sec2.paragraph_format.space_after = Pt(2)
        r = p_sec2.add_run("II. BẢNG GHI LẠI THÔNG SỐ ĐO LƯỜNG KHI CHO TỪNG GIẢI PHÁP VÀO:")
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(10)
        r_sub = p_sec2.add_run(" (Học sinh đo đạc trên mô hình mô phỏng web và ghi vào bảng)")
        r_sub.font.name = "Times New Roman"
        r_sub.font.italic = True
        r_sub.font.size = Pt(8.5)

        col_widths = [Inches(1.0), Inches(1.85), Inches(0.72), Inches(0.75), Inches(0.65), Inches(0.75), Inches(1.25)]
        table = doc.add_table(rows=6, cols=7)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table, color="000000", sz="4", val="single")

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
            set_cell_valign(c, "center")
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
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
                set_cell_margins(c, top=90, bottom=90, left=50, right=50)
                set_cell_valign(c, "center")
                p = c.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.05

                if c_idx == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r = p.add_run(step_labels[step_idx])
                    r.font.name = "Times New Roman"
                    r.font.bold = True
                    r.font.size = Pt(8)
                    r.font.color.rgb = RGBColor(0, 0, 0)
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # 6. Section III
        p_sec3 = doc.add_paragraph()
        p_sec3.paragraph_format.space_before = Pt(4)
        p_sec3.paragraph_format.space_after = Pt(1)
        r = p_sec3.add_run("III. ĐÁNH GIÁ MỨC ĐỘ CẢI THIỆN VI KHÍ HẬU (Δ SAU KHI HOÀN TẤT - BAN ĐẦU):")
        r.font.name = "Times New Roman"
        r.font.bold = True
        r.font.size = Pt(9.5)

        p_delta = doc.add_paragraph()
        p_delta.paragraph_format.space_before = Pt(0)
        p_delta.paragraph_format.space_after = Pt(4)
        p_delta.paragraph_format.line_spacing = 1.15
        r = p_delta.add_run("• Nồng độ CO₂: Thay đổi từ ........................... ppm  →  ........................... ppm  (Tăng/Giảm: ....................................................)\n")
        r.font.name = "Times New Roman"
        r.font.size = Pt(8.5)
        r2 = p_delta.add_run("• Cường độ ánh sáng: Thay đổi từ ........................... Lux  →  ........................... Lux  (Tăng/Giảm: ............................................)\n")
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(8.5)
        r3 = p_delta.add_run("• Nhiệt độ vi khí hậu: Thay đổi từ ........................... °C  →  ........................... °C  (Tăng/Giảm: ...................................................)\n")
        r3.font.name = "Times New Roman"
        r3.font.size = Pt(8.5)
        r4 = p_delta.add_run("• Hiệu quả quang hợp: Tăng vọt từ ........................... %  →  ........................... %  (Trạng thái lá: ......................................................)")
        r4.font.name = "Times New Roman"
        r4.font.size = Pt(8.5)
        r4.font.bold = True

        # 7. Signatures
        sign_table = doc.add_table(rows=1, cols=2)
        sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        remove_table_borders(sign_table)
        
        sc1 = sign_table.rows[0].cells[0]
        sc1.width = Inches(3.4)
        sp1 = sc1.paragraphs[0]
        sp1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp1.paragraph_format.space_before = Pt(2)
        sp1.paragraph_format.space_after = Pt(0)
        sp1.paragraph_format.line_spacing = 1.1
        r = sp1.add_run("ĐÁNH GIÁ CỦA NHÓM\n(Mức độ hoàn thành nhiệm vụ)\n........................................................................")
        r.font.name = "Times New Roman"
        r.font.size = Pt(8.5)

        sc2 = sign_table.rows[0].cells[1]
        sc2.width = Inches(3.57)
        sp2 = sc2.paragraphs[0]
        sp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sp2.paragraph_format.space_before = Pt(2)
        sp2.paragraph_format.space_after = Pt(0)
        sp2.paragraph_format.line_spacing = 1.1
        r = sp2.add_run("NHẬN XÉT CỦA GIÁO VIÊN\n(Ký và ghi rõ họ tên)\n\n\nCô Mai Thuý Đức")
        r.font.name = "Times New Roman"
        r.font.size = Pt(9)

    doc.save(filename)
    print(f"Created all-in-one standard worksheet: {filename}")

if __name__ == "__main__":
    build_all_worksheets()
