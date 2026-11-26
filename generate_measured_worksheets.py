import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def set_table_borders(table, color="D1D5DB", sz="4", val="single"):
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

stations = [
    {
        "code": "2A",
        "group": "NHÓM 1",
        "title": "TRẠM 1: ĐÔ THỊ NÉN & THÁP RỪNG THẲNG ĐỨNG",
        "subtitle": "Chung cư cao tầng & Nhà phố liền kề TP.HCM",
        "challenge": "Không gian chật hẹp, nồng độ CO₂ cao do khí thải xe cộ, nhiệt độ hấp thụ từ bê tông cao (hiệu ứng đảo nhiệt). Tầng trên nắng gắt, tầng trệt thiếu sáng.",
        "mission": "Ứng dụng mô hình phân tầng sinh thái Bosco Verticale: Tuyển chọn và kích hoạt từng giải pháp cây xanh, hệ thống tưới tự động để hạ nhiệt, lọc không khí và tối ưu hóa quang hợp.",
        "steps": [
            ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp nào", "850 ppm", "45,000 Lux", "36.0°C", "Chưa tối ưu (0%)", "Hiện trạng ban đầu: Bê tông giữ nhiệt, khói xe dày đặc, quang hợp bị ức chế nghiêm trọng."),
            ("Bước 1", "Trồng cây ưa sáng có tầng cutin dày (Hoa giấy, Trúc quân tử)", "730 ppm", "45,000 Lux", "33.5°C", "Kém (25%)", "Cây tầng thượng có tán lá dày giúp cản 15.000 Lux nắng gắt, giảm bức xạ nhiệt xuống tầng dưới."),
            ("Bước 2", "Bố trí vườn tường đứng nhiều tầng (Vertical Wall) bằng giàn cây leo bản rộng", "580 ppm", "30,000 Lux", "31.5°C", "Trung bình (50%)", "Giàn cây leo che phủ bê tông làm mát tường thêm 2.5°C và hấp thụ 120 ppm khí CO₂."),
            ("Bước 3", "Trồng cây ưa bóng (Lưỡi hổ, Lan ý) có khả năng lọc formaldehyde", "460 ppm", "30,000 Lux", "30.5°C", "Khá tốt (75%)", "Cây ưa bóng tầng trệt hấp thụ CO₂ khói xe và lọc formaldehyde mà không bị cháy lá."),
            ("Bước 4", "Cài đặt van tưới nhỏ giọt hẹn giờ tự động kết hợp khay hứng tràn", "410 ppm", "25,000 Lux", "28.0°C", "Tối ưu 100% 🌿", "Cung cấp nước đều đặn làm khí khổng mở tối đa, CO₂ khuếch tán vào mô lá đạt đỉnh quang hợp 100%.")
        ],
        "delta": "• Nồng độ CO₂: Giảm từ 850 ppm → 410 ppm (-440 ppm, giảm ô nhiễm khí thải, đạt nồng độ tối ưu)\n• Cường độ ánh sáng: Điều tiết từ 45,000 Lux → 25,000 Lux (cản bớt nắng gắt tầng thượng, tán xạ dịu xuống trệt)\n• Nhiệt độ vi khí hậu: Hạ từ 36.0°C → 28.0°C (hạ 8.0°C, triệt tiêu hiệu ứng đảo nhiệt bê tông)\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (cả 4 tầng cây xanh tốt)",
        "q1": "Sự thay đổi của các thông số đo lường (CO₂, Ánh sáng, Nhiệt độ, Nước) sau khi cho từng giải pháp vào đã tác động như thế nào đến sự đóng/mở khí khổng và hiệu suất quang hợp của cây xanh tại trạm?",
        "q2": "Từ kết quả thực nghiệm đo lường trên, nhóm hãy rút ra quy luật sinh học về sự phối hợp đồng bộ giữa 4 yếu tố môi trường (Ánh sáng, Nước, CO₂, Nhiệt độ) đối với quá trình quang hợp."
    },
    {
        "code": "2B",
        "group": "NHÓM 2",
        "title": "TRẠM 2: NHÀ MÀNG NÔNG NGHIỆP THÔNG MINH",
        "subtitle": "Nhà màng công nghệ cao Củ Chi - TP.HCM",
        "challenge": "Yêu cầu năng suất và chất lượng quả tối đa, cần chủ động kiểm soát cả 4 yếu tố quang hợp quanh năm bất chấp thời tiết bất lợi bên ngoài.",
        "mission": "Thiết lập hệ thống cảm biến vi khí hậu và kích hoạt từng giải pháp nông nghiệp chính xác (đèn LED, làm giàu CO₂, quạt làm mát, tưới nhỏ giọt) để thúc đẩy quang hợp đạt cực đại.",
        "steps": [
            ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp nào", "400 ppm (Tự nhiên)", "12.000 Lux (Thiếu)", "34.0°C", "Chưa tối ưu (0%)", "Hiện trạng ban đầu: Cây phụ thuộc thời tiết bên ngoài, thiếu sáng và nhiệt độ bí nóng."),
            ("Bước 1", "Lắp đặt đèn LED chuyên dụng phát dải sáng xanh (450nm) và đỏ (660nm)", "400 ppm (Tự nhiên)", "35.000 Lux (LED)", "34.0°C", "Kém (25%)", "Đèn LED đa phổ cung cấp đúng 2 bước sóng 450nm và 660nm, kích thích mạnh pha sáng."),
            ("Bước 2", "Bơm bổ sung khí CO₂ duy trì ở ngưỡng tối ưu 0.10% – 0.15% (1.000 - 1.500 ppm)", "1.200 ppm (Tối ưu)", "35.000 Lux (LED)", "34.0°C", "Trung bình (50%)", "Bơm CO₂ bổ sung lên 1.200 ppm làm tăng tốc độ tổng hợp đường bột trước điểm bão hòa."),
            ("Bước 3", "Lưới cắt nắng tự động kích hoạt kết hợp quạt hút và màng nước làm mát", "1.200 ppm (Tối ưu)", "35.000 Lux (LED)", "28.0°C", "Khá tốt (75%)", "Quạt hút gió và màng nước hạ nhiệt về 27°C, bảo vệ các enzyme quang hợp hoạt động cực đại."),
            ("Bước 4", "Tưới dinh dưỡng nhỏ giọt tự động kiểm soát theo độ ẩm đất đo bằng cảm biến IoT", "1.200 ppm (Tối ưu)", "35.000 Lux (LED)", "26.0°C", "Tối ưu 100% 🌿", "Tưới nhỏ giọt IoT cung cấp nước liên tục, duy trì tế bào khí khổng mở suốt ngày đêm.")
        ],
        "delta": "• Nồng độ CO₂: Tăng từ 400 ppm → 1.200 ppm (làm giàu nguyên liệu cho pha tối quang hợp tổng hợp đường)\n• Cường độ ánh sáng: Tăng từ 12.000 Lux → 35.000 Lux (đèn LED chuyên dụng tập trung dải xanh 450nm & đỏ 660nm)\n• Nhiệt độ vi khí hậu: Hạ từ 34.0°C → 26.0°C (hạ 8.0°C, đưa về khoảng nhiệt độ tối ưu của enzyme xúc tác quang hợp)\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (năng suất sinh khối và độ ngọt đạt tối đa)",
        "q1": "Tại sao người trồng dưa lưới trong nhà màng công nghệ cao lại chủ động bơm thêm khí CO₂ lên 1.200 ppm vào buổi sáng? Mức CO₂ tăng lên có tác dụng gì đối với quang hợp?",
        "q2": "Nếu chỉ làm giàu CO₂ và bổ sung đèn LED mà không bật quạt làm mát (để nhiệt độ >38°C) thì cây dưa lưới có quang hợp tốt không? Vì sao?"
    },
    {
        "code": "2C",
        "group": "NHÓM 3",
        "title": "TRẠM 3: SÂN THƯỢNG NẮNG GẮT & NHIỆT ĐỘ CAO",
        "subtitle": "Vùng mái nhà bê tông & Ban công hướng Tây TP.HCM",
        "challenge": "Nắng trưa gay gắt > 55.000 Lux, nhiệt độ sàn bê tông lên tới 43°C. Khí khổng đóng chặt tránh mất nước làm ngừng quang hợp, lá dễ bị 'đốt nóng' và cháy khô.",
        "mission": "Thiết kế giải pháp giải cứu vườn sân thượng: Căng lưới lan, tuyển chọn cây chịu hạn cutin dày, phun sương sàn và phủ gốc giữ ẩm để tái mở khí khổng và phục hồi quang hợp.",
        "steps": [
            ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp nào", "Khí khổng đóng", "58,000 Lux", "43.0°C", "Chưa tối ưu (0%)", "Hiện trạng ban đầu: Nắng trưa >55.000 Lux làm đốt nóng lá, nhiệt độ 43°C khiến khí khổng đóng chặt."),
            ("Bước 1", "Căng lưới lan đen che chắn giảm 40 - 50% ánh sáng trong khung giờ 11h - 14h", "Khí khổng đóng", "28,000 Lux", "39.0°C", "Kém (25%)", "Lưới lan đen giảm 50% cường độ bức xạ, ngăn chặn phá hủy diệp lục tố của lá."),
            ("Bước 2", "Trồng các giống cây chịu hạn, ưa sáng mạnh (Hoa sứ Thái, Hoa giấy, Xương rồng)", "Khí khổng đóng", "28,000 Lux", "38.0°C", "Trung bình (50%)", "Cây hoa sứ và xương rồng có tầng cutin dày, thích nghi và bắt đầu quang hợp dưới nhiệt độ cao."),
            ("Bước 3", "Phun sương làm mát bề mặt sàn bê tông vào giữa trưa để hạ nhiệt độ về dưới 33°C", "Khí khổng mở", "28,000 Lux", "33.0°C", "Khá tốt (75%)", "Phun sương sàn hạ nhiệt độ mặt bê tông xuống 34°C, kích thích tế bào khí khổng hé mở lại."),
            ("Bước 4", "Phủ lớp vỏ trấu / sỏi nhẹ giữ ẩm đất chậu và chỉ tưới vào sáng sớm hoặc chiều mát", "Khí khổng mở", "28,000 Lux", "31.0°C", "Tối ưu 100% 🌿", "Lớp phủ rơm rạ và sỏi nhẹ ngăn bốc hơi nước rễ, giúp cây quang hợp bền bỉ cả ngày.")
        ],
        "delta": "• Cường độ ánh sáng: Điều tiết từ 58,000 Lux (nắng gắt cháy lá) → 28,000 Lux (vùng bão hòa ánh sáng an toàn)\n• Trạng thái khí khổng: Từ đóng chặt chống mất nước → Khí khổng hé mở và mở hoàn toàn đón nhận khí CO₂\n• Nhiệt độ vi khí hậu: Hạ từ 43.0°C (sàn bê tông nung nóng) → 31.0°C (hạ 12.0°C bảo vệ tế bào và hệ rễ cây)\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (vườn sân thượng xanh tươi, dịu mát)",
        "q1": "Vì sao sáng nào bạn nhỏ cũng tưới no nước vào chậu nhưng Bé Lan Ý đặt trên sân thượng giữa trưa hè vẫn bị cháy mép lá khô giòn và ngừng quang hợp?",
        "q2": "Hãy giải thích vai trò của việc phun sương làm mát sàn bê tông đối với việc kích thích khí khổng mở ra trong quá trình quang hợp."
    },
    {
        "code": "2D",
        "group": "NHÓM 4",
        "title": "TRẠM 4: TRONG NHÀ & HÀNH LANG RÂM MÁT",
        "subtitle": "Lớp học, văn phòng & Khu vực gầm cầu thang",
        "challenge": "Ánh sáng rất yếu (< 1.500 Lux, dưới điểm bù ánh sáng của cây ưa sáng), thiếu thông gió tự nhiên, nồng độ CO₂ tăng cao do hơi thở con người.",
        "mission": "Thiết kế giải pháp vi khí hậu trong nhà: Tuyển chọn cây ưa bóng giàu diệp lục b, đo độ rọi bằng smartphone, bổ sung đèn LED 4000K và vệ sinh bụi lá để khôi phục quang hợp và thanh lọc không khí.",
        "steps": [
            ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp nào", "650 ppm", "800 Lux", "26.0°C", "Chưa tối ưu (0%)", "Hiện trạng ban đầu: Ánh sáng <800 Lux dưới điểm bù quang hợp, cây tiêu hao nhiều hơn tạo ra."),
            ("Bước 1", "Trang bị đèn rọi LED ánh sáng trắng ấm 4000K (15W) chiếu bổ sung 3 - 4 giờ/ngày", "650 ppm", "1,500 Lux", "26.0°C", "Kém (25%)", "App Lux Meter định vị được vị trí đón ánh sáng tán xạ 1.500 Lux gần cửa sổ."),
            ("Bước 2", "Tuyển chọn 100% cây ưa bóng có phiến lá to mỏng, giàu diệp lục b (Lá lốt, Trầu không, Kim tiền)", "590 ppm", "1,500 Lux", "26.0°C", "Trung bình (50%)", "Cây lá lốt ưa bóng với phiến lá to, mỏng, chứa nhiều diệp lục b bắt trọn ánh sáng yếu."),
            ("Bước 3", "Dùng app 'Lux Meter' quét chọn các vị trí gần cửa sổ / giếng trời có độ rọi > 1.500 Lux", "590 ppm", "2,500 Lux", "26.0°C", "Khá tốt (75%)", "Đèn rọi LED 15W bổ sung thêm 1.000 Lux, kích thích pha sáng diễn ra bình thường."),
            ("Bước 4", "Định kỳ lau sạch bụi bám trên bề mặt lá bằng khăn ẩm mềm", "450 ppm", "2,500 Lux", "26.0°C", "Tối ưu 100% 🌿", "Lau sạch bụi bám trên bề mặt lá giúp lỗ khí khổng thông thoáng, CO₂ khuếch tán dễ dàng.")
        ],
        "delta": "• Cường độ ánh sáng: Tăng từ 800 Lux (dưới điểm bù ánh sáng) → 2,500 Lux (quang hợp hiệu quả, vượt điểm bù)\n• Nồng độ CO₂: Giảm từ 650 ppm (khí thở tù đọng) → 450 ppm (cây quang hợp tích cực, thanh lọc khí thở)\n• Nhiệt độ vi khí hậu: Ổn định ở mức 26.0°C (mát mẻ, phù hợp tối ưu cho enzyme quang hợp trong nhà)\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (cây phát triển xanh tốt, lá bóng mọng)",
        "q1": "Tại sao cây đặt sâu trong phòng tối dù được tưới nước đầy đủ nhưng lá cứ nhạt màu dần rồi rụng lả tả? (Liên hệ khái niệm điểm bù ánh sáng).",
        "q2": "Vì sao việc dùng khăn mềm ẩm lau sạch bụi bám trên phiến lá lại giúp cây trong nhà tăng cường hiệu quả quang hợp?"
    }
]

def build_station_worksheets(filename="Phieu_Hoc_Tap_KHTN7_Bai23_Quang_Hop_4_Tram.docx"):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.55)
        section.bottom_margin = Inches(0.55)
        section.left_margin = Inches(0.55)
        section.right_margin = Inches(0.55)

    for s_idx, st in enumerate(stations):
        if s_idx > 0:
            doc.add_page_break()

        # Header Block
        p_hdr = doc.add_paragraph()
        p_hdr.paragraph_format.space_before = Pt(0)
        p_hdr.paragraph_format.space_after = Pt(2)
        r = p_hdr.add_run("TRƯỜNG THCS TÂN TẠO A - TỔ KHOA HỌC TỰ NHIÊN\n")
        r.font.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(75, 85, 99)
        r2 = p_hdr.add_run("BỘ PHIẾU HỌC TẬP DỰ ÁN STEM (PHIẾU HỌC TẬP SỐ 2) • BÀI 23 KHTN 7")
        r2.font.bold = True
        r2.font.size = Pt(10)
        r2.font.color.rgb = RGBColor(16, 122, 80)

        # Title Box
        t_box = doc.add_table(rows=1, cols=1)
        t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = t_box.rows[0].cells[0]
        cell.width = Inches(7.2)
        set_cell_background(cell, "064E3B") # Emerald 900
        set_cell_margins(cell, top=130, bottom=130, left=180, right=180)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"PHIẾU HỌC TẬP SỐ {st['code']} — DÀNH CHO {st['group']}\n")
        r.font.bold = True
        r.font.size = Pt(12.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        r2 = p.add_run(f"NHIỆM VỤ THỰC NGHIỆM ĐO LƯỜNG TIẾN TRÌNH VI KHÍ HẬU & QUANG HỢP\n")
        r2.font.bold = True
        r2.font.size = Pt(10)
        r2.font.color.rgb = RGBColor(167, 243, 208)
        r3 = p.add_run(f"{st['title']} ({st['subtitle']})")
        r3.font.bold = True
        r3.font.size = Pt(10)
        r3.font.color.rgb = RGBColor(253, 230, 138)

        # Student info bar
        p_info = doc.add_paragraph()
        p_info.paragraph_format.space_before = Pt(5)
        p_info.paragraph_format.space_after = Pt(5)
        r = p_info.add_run("Tên nhóm: ..................................................... Lớp: 7/......  •  GV hướng dẫn: ")
        r.font.size = Pt(9.5)
        r2 = p_info.add_run("Cô Mai Thuý Đức")
        r2.font.bold = True
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = RGBColor(16, 122, 80)
        r3 = p_info.add_run("  •  Năm học: 2026 - 2027\n")
        r3.font.size = Pt(9.5)
        r4 = p_info.add_run("Thành viên nhóm: 1............................................... 2............................................... 3............................................... 4...............................................")
        r4.font.size = Pt(9)
        r4.font.color.rgb = RGBColor(75, 85, 99)

        # Section I: Context & Mission
        p_sec1 = doc.add_paragraph()
        p_sec1.paragraph_format.space_before = Pt(3)
        p_sec1.paragraph_format.space_after = Pt(2)
        r = p_sec1.add_run("I. KHẢO SÁT HIỆN TRẠNG & THÁCH THỨC VI KHÍ HẬU TẠI TRẠM:")
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(27, 67, 50)

        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(0)
        p_desc.paragraph_format.space_after = Pt(4)
        r = p_desc.add_run("• Thách thức vi khí hậu: ")
        r.font.bold = True
        r.font.size = Pt(9.5)
        r2 = p_desc.add_run(st['challenge'] + "\n")
        r2.font.size = Pt(9.5)
        r3 = p_desc.add_run("• Nhiệm vụ thực nghiệm số: ")
        r3.font.bold = True
        r3.font.size = Pt(9.5)
        r3.font.color.rgb = RGBColor(16, 122, 80)
        r4 = p_desc.add_run(st['mission'])
        r4.font.size = Pt(9.5)

        # Section II: Observation & Measurement Table (5 Steps)
        p_sec2 = doc.add_paragraph()
        p_sec2.paragraph_format.space_before = Pt(5)
        p_sec2.paragraph_format.space_after = Pt(3)
        r = p_sec2.add_run("II. BẢNG GHI LẠI THÔNG SỐ ĐO LƯỜNG KHI CHO TỪNG GIẢI PHÁP VÀO:")
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(27, 67, 50)
        r_sub = p_sec2.add_run(" (Học sinh đo đạc trực tiếp trên mô hình mô phỏng Module 3)")
        r_sub.font.italic = True
        r_sub.font.size = Pt(9)
        r_sub.font.color.rgb = RGBColor(100, 116, 139)

        # Table: 6 rows (1 header + 5 steps), 7 columns
        col_widths = [Inches(0.9), Inches(1.85), Inches(0.80), Inches(0.85), Inches(0.65), Inches(0.80), Inches(1.35)]
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
            set_cell_margins(c, top=80, bottom=80, left=60, right=60)
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(title)
            r.font.bold = True
            r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(255, 255, 255)

        for step_idx, step_data in enumerate(st['steps']):
            row = table.rows[step_idx + 1]
            bg_color = "F9FAFB" if step_idx % 2 == 0 else "FFFFFF"
            if step_idx == 0:
                bg_color = "FEF2F2" # Red tint for initial condition
            elif step_idx == 4:
                bg_color = "ECFDF5" # Green tint for optimal 100%

            for c_idx in range(7):
                c = row.cells[c_idx]
                c.width = col_widths[c_idx]
                set_cell_background(c, bg_color)
                set_cell_margins(c, top=65, bottom=65, left=65, right=65)
                p = c.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                
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

        # Section III: Delta Improvement
        p_sec3 = doc.add_paragraph()
        p_sec3.paragraph_format.space_before = Pt(5)
        p_sec3.paragraph_format.space_after = Pt(2)
        r = p_sec3.add_run("III. ĐÁNH GIÁ MỨC ĐỘ CẢI THIỆN VI KHÍ HẬU (Δ SAU KHI HOÀN TẤT - BAN ĐẦU):")
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(27, 67, 50)

        p_delta = doc.add_paragraph()
        p_delta.paragraph_format.space_before = Pt(0)
        p_delta.paragraph_format.space_after = Pt(4)
        r = p_delta.add_run(st['delta'])
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(4, 120, 87)

        # Section IV: Scientific Questions
        p_sec4 = doc.add_paragraph()
        p_sec4.paragraph_format.space_before = Pt(4)
        p_sec4.paragraph_format.space_after = Pt(2)
        r = p_sec4.add_run("IV. CÂU HỎI KHOA HỌC & RÚT RA QUY LUẬT SINH HỌC (BÁO CÁO E-POSTER):")
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(27, 67, 50)

        p_q1 = doc.add_paragraph()
        p_q1.paragraph_format.space_before = Pt(1)
        p_q1.paragraph_format.space_after = Pt(1)
        r = p_q1.add_run("1. Câu hỏi 1: ")
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(180, 83, 9)
        r2 = p_q1.add_run(st['q1'] + "\n")
        r2.font.size = Pt(9)
        r3 = p_q1.add_run("Trả lời: ...........................................................................................................................................................................................................................................\n..........................................................................................................................................................................................................................................................")
        r3.font.size = Pt(8.5)
        r3.font.color.rgb = RGBColor(107, 114, 128)

        p_q2 = doc.add_paragraph()
        p_q2.paragraph_format.space_before = Pt(3)
        p_q2.paragraph_format.space_after = Pt(3)
        r = p_q2.add_run("2. Câu hỏi 2: ")
        r.font.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(180, 83, 9)
        r2 = p_q2.add_run(st['q2'] + "\n")
        r2.font.size = Pt(9)
        r3 = p_q2.add_run("Trả lời: ...........................................................................................................................................................................................................................................\n..........................................................................................................................................................................................................................................................")
        r3.font.size = Pt(8.5)
        r3.font.color.rgb = RGBColor(107, 114, 128)

        # Signature Footer
        p_sign = doc.add_paragraph()
        p_sign.paragraph_format.space_before = Pt(5)
        p_sign.paragraph_format.space_after = Pt(0)
        r = p_sign.add_run("ĐÁNH GIÁ CỦA NHÓM: ....................................................          CHỮ KÝ DUYỆT CỦA GIÁO VIÊN: ")
        r.font.size = Pt(8.5)
        r.font.bold = True
        r2 = p_sign.add_run("Cô Mai Thuý Đức")
        r2.font.bold = True
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = RGBColor(16, 122, 80)

    doc.save(filename)
    print(f"File {filename} created successfully with 100% exact web telemetry data!")

if __name__ == "__main__":
    build_station_worksheets()
