import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from generate_measured_worksheets import stations

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=70, bottom=70, left=70, right=70):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def update_lesson_plan(filepath):
    print(f"Updating lesson plan: {filepath}")
    doc = docx.Document(filepath)

    # Tables 2, 3, 4, 5 correspond to Station 1, 2, 3, 4
    for st_idx, st in enumerate(stations):
        t_idx = 2 + st_idx
        table = doc.tables[t_idx]

        # Row 0: header
        headers = [
            "Bước thực nghiệm",
            "Giải pháp kỹ thuật đưa vào\n(Kéo thả trên web)",
            "Nồng độ CO₂",
            "Ánh sáng\n(Lux)",
            "Nhiệt độ\n(°C)",
            "Hiệu quả quang hợp",
            "Hiện tượng quan sát\n(Trạng thái lá & Khí khổng)"
        ]
        hdr_row = table.rows[0]
        for c_idx, h_text in enumerate(headers):
            cell = hdr_row.cells[c_idx]
            set_cell_background(cell, "1E3A8A")
            set_cell_margins(cell, top=80, bottom=80, left=60, right=60)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.text = ""
            r = p.add_run(h_text)
            r.font.bold = True
            r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(255, 255, 255)

        # Rows 1 to 5: 5 steps
        for step_idx, step_data in enumerate(st['steps']):
            row = table.rows[step_idx + 1]
            bg_color = "F9FAFB" if step_idx % 2 == 0 else "FFFFFF"
            if step_idx == 0:
                bg_color = "FEF2F2"
            elif step_idx == 4:
                bg_color = "ECFDF5"

            for c_idx in range(7):
                cell = row.cells[c_idx]
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=65, bottom=65, left=65, right=65)
                p = cell.paragraphs[0]
                p.text = ""
                if c_idx in [0, 2, 3, 4, 5]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                val = step_data[c_idx]
                r = p.add_run(val)
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

    # Now let's update the text descriptions before and after the tables in Phụ lục 2
    # P23 to P34
    p_map = {
        23: f"=== PHIẾU HỌC TẬP 2A (DÀNH CHO NHÓM 1) ===\nĐỒ ÁN STEM: {stations[0]['title']} ({stations[0]['subtitle'].upper()})\n(Dành cho giáo viên tổ chức và học sinh ghi nhận khi thực hành kéo thả)",
        24: f"• Thách thức vi khí hậu: {stations[0]['challenge']}\n• Nhiệm vụ thực nghiệm số: {stations[0]['mission']}",
        25: f"• Đánh giá mức độ cải thiện vi khí hậu (Δ Sau khi hoàn tất cả 4 giải pháp - Ban đầu):\n{stations[0]['delta']}",

        26: f"=== PHIẾU HỌC TẬP 2B (DÀNH CHO NHÓM 2) ===\nĐỒ ÁN STEM: {stations[1]['title']} ({stations[1]['subtitle'].upper()})\n(Dành cho giáo viên tổ chức và học sinh ghi nhận khi thực hành kéo thả)",
        27: f"• Thách thức vi khí hậu: {stations[1]['challenge']}\n• Nhiệm vụ thực nghiệm số: {stations[1]['mission']}",
        28: f"• Đánh giá mức độ cải thiện vi khí hậu (Δ Sau khi hoàn tất cả 4 giải pháp - Ban đầu):\n{stations[1]['delta']}",

        29: f"=== PHIẾU HỌC TẬP 2C (DÀNH CHO NHÓM 3) ===\nĐỒ ÁN STEM: {stations[2]['title']} ({stations[2]['subtitle'].upper()})\n(Dành cho giáo viên tổ chức và học sinh ghi nhận khi thực hành kéo thả)",
        30: f"• Thách thức vi khí hậu: {stations[2]['challenge']}\n• Nhiệm vụ thực nghiệm số: {stations[2]['mission']}",
        31: f"• Đánh giá mức độ cải thiện vi khí hậu (Δ Sau khi hoàn tất cả 4 giải pháp - Ban đầu):\n{stations[2]['delta']}",

        32: f"=== PHIẾU HỌC TẬP 2D (DÀNH CHO NHÓM 4) ===\nĐỒ ÁN STEM: {stations[3]['title']} ({stations[3]['subtitle'].upper()})\n(Dành cho giáo viên tổ chức và học sinh ghi nhận khi thực hành kéo thả)",
        33: f"• Thách thức vi khí hậu: {stations[3]['challenge']}\n• Nhiệm vụ thực nghiệm số: {stations[3]['mission']}",
        34: f"• Đánh giá mức độ cải thiện vi khí hậu (Δ Sau khi hoàn tất cả 4 giải pháp - Ban đầu):\n{stations[3]['delta']}"
    }

    for p_idx, text in p_map.items():
        if p_idx < len(doc.paragraphs):
            p = doc.paragraphs[p_idx]
            p.text = text
            # Format nicely
            for r in p.runs:
                r.font.size = Pt(9.5)
                r.font.name = "Times New Roman"

    doc.save(filepath)
    print(f"Successfully updated {filepath} with 100% exact web telemetry data!")

if __name__ == "__main__":
    update_lesson_plan("02_Phan_2_Ke_Hoach_Bai_Day_CV5512_Mai_Thuy_Duc.docx")
    update_lesson_plan("KHTN_THCS_Mai-Thuy-Duc_STEM_P2.docx")
