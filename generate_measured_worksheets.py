import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

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
            "title": "TRẠM 1: ĐÔ THỊ NÉN & THÁP RỪNG THẲNG ĐỨNG (BOSCO VERTICALE SÀI GÒN)",
            "context": "Chung cư cao tầng tại Quận 1, TP.HCM. Hiệu ứng đảo nhiệt đô thị gay gắt, nhiệt độ bề mặt bê tông 36°C - 39°C. Khí thải xe cộ khiến nồng độ CO₂ tăng cao cục bộ (850 ppm). Tầng mái nắng gắt (>45.000 Lux), sảnh kín và tầng trệt thiếu sáng (<650 Lux). Bụi mịn PM2.5 cao.",
            "mission": "Ứng dụng mô hình phân tầng sinh thái Bosco Verticale: Tuyển chọn và đưa từng giải pháp cây xanh, hệ thống tưới vào để hạ nhiệt, lọc không khí và tối ưu hóa quang hợp.",
            "steps": [
                ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp (Bê tông khô nóng)", "850 ppm (Cao ⚠️)", "45.000 Lux (Nắng gắt) / 650 Lux (Trệt)", "36.0°C (Nóng ⚠️)", "15% (Kém)", "Khí khổng đóng chặt để giữ nước; mép lá Lan Ý cháy vàng khô rụm."),
                ("Bước 1", "Đưa Giải pháp 1: Vách tường đứng nhiều tầng (Trầu bà, Dây cúc tần Ấn Độ, Lưỡi hổ)", "720 ppm", "42.000 Lux", "34.0°C", "45% (Bắt đầu phục hồi)", "Tán lá cản bức xạ trực tiếp vào tường, hấp thụ bước đầu CO₂ và giữ bụi mịn PM2.5."),
                ("Bước 2", "Đưa Giải pháp 2: Tầng thượng trồng cây ưa sáng phiến lá dày cản bụi (Hoa giấy, Trúc quân tử)", "620 ppm", "38.000 Lux", "31.5°C", "70% (Khá tốt)", "Phiến lá dày, tầng cutin bóng cản bức xạ, quang hợp tạo tinh bột mạnh mẽ trên cao."),
                ("Bước 3", "Đưa Giải pháp 3: Tầng trệt & sảnh kín bố trí cây ưa bóng phiến lá mỏng lọc khí độc (Bé Lan Ý)", "510 ppm", "1.200 Lux (Tán xạ dịu)", "29.5°C", "85% (Tốt)", "Lan Ý hấp thu hiệu quả ánh sáng yếu trong nhà, phiến lá hồi phục xanh mọng nước."),
                ("Bước 4", "Đưa Giải pháp 4: Hệ thống tưới nhỏ giọt hẹn giờ tự động & Thu gom nước mưa tuần hoàn", "410 ppm (Lý tưởng)", "Phân tầng tối ưu", "28.0°C (Mát mẻ)", "100% (Cực đại ⚡)", "Nước cung cấp đầy đủ khiến tế bào hạt đậu trương nước mở to khí khổng; quang hợp đạt đỉnh cao.")
            ],
            "delta": "• Nồng độ CO₂: Giảm từ 850 ppm → 410 ppm (giảm 440 ppm, không khí sạch mát)\n• Nhiệt độ vi khí hậu: Hạ từ 36.0°C → 28.0°C (hạ 8.0°C, triệt tiêu hiệu ứng đảo nhiệt)\n• Hiệu suất quang hợp: Tăng vọt từ 15% → 100% (cả 4 tầng cây đều xanh tốt)",
            "q1": "Sự thay đổi của các thông số đo lường (CO₂, Ánh sáng, Nhiệt độ, Nước) sau khi cho từng giải pháp vào đã tác động như thế nào đến sự đóng/mở khí khổng và hiệu suất quang hợp của cây xanh tại trạm?",
            "q2": "Từ kết quả thực nghiệm đo lường trên, nhóm hãy rút ra quy luật sinh học về sự phối hợp đồng bộ giữa 4 yếu tố môi trường (Ánh sáng, Nước, CO₂, Nhiệt độ) đối với quá trình quang hợp."
        },
        {
            "code": "2B",
            "group": "NHÓM 2",
            "title": "TRẠM 2: NHÀ MÀNG NÔNG NGHIỆP CÔNG NGHỆ CAO (CỦ CHI - TP.HCM)",
            "context": "Nhà màng công nghệ cao trồng dưa lưới F1 và rau thủy canh tại Củ Chi. Vào giữa trưa hè, hiệu ứng nhà kính làm nhiệt độ không khí vọt lên 38.5°C - 43°C gây sốc nhiệt cho cây. Mùa khô thiếu ẩm, nồng độ CO₂ tự nhiên bị cây tiêu thụ cạn kiệt.",
            "mission": "Thiết lập hệ thống cảm biến và lập trình đưa từng giải pháp vi khí hậu thông minh vào nhà màng để kích thích quang hợp tích lũy đường Brix và sinh khối.",
            "steps": [
                ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp (Nhà màng hầm nóng, kín khí)", "380 ppm (Cạn kiệt ⚠️)", "12.000 Lux (Kém sáng)", "38.5°C (Quá nóng ⚠️)", "20% (Kém)", "Enzyme quang hợp bị ức chế do sốc nhiệt; cây dưa lưới héo rũ, tích lũy đường ngưng trệ."),
                ("Bước 1", "Đưa Giải pháp 1: Lưới cắt nắng tự động 50% & Quạt đối lưu không khí lưu thông", "400 ppm", "22.000 Lux", "31.0°C", "50% (Phục hồi)", "Hạ cường độ bức xạ thiêu đốt, luồng gió đối lưu giúp giảm nhiệt độ mặt lá."),
                ("Bước 2", "Đưa Giải pháp 2: Hệ thống phun sương làm mát tự động theo cảm biến ẩm độ", "420 ppm", "25.000 Lux", "27.5°C", "75% (Tốt)", "Ẩm độ đạt mức tối ưu 65-70%, hạn chế bốc thoát hơi nước quá mức, giữ tế bào lá căng mọng."),
                ("Bước 3", "Đưa Giải pháp 3: Làm giàu nồng độ khí CO₂ sinh học (800 - 1.000 ppm) vào buổi sáng", "900 ppm (Bão hòa CO₂)", "28.000 Lux", "26.5°C", "90% (Rất cao)", "Cung cấp dồi dào nguyên liệu cho pha tối quang hợp, tốc độ cố định Carbon tăng vọt, trái tích đường nhanh."),
                ("Bước 4", "Đưa Giải pháp 4: Tưới nhỏ giọt bù áp tự động điều hòa nồng độ dinh dưỡng khoáng EC chuẩn", "1.000 ppm", "32.000 Lux", "26.0°C (Lý tưởng)", "100% (Cực đại ⚡)", "Dinh dưỡng khoáng và nước đầy đủ giúp tổng hợp chất hữu cơ tối đa; độ ngọt Brix vượt 14.")
            ],
            "delta": "• Nồng độ CO₂: Tăng từ 380 ppm → 1.000 ppm (làm giàu nguyên liệu cho chu trình Calvin)\n• Nhiệt độ nhà màng: Hạ từ 38.5°C → 26.0°C (hạ 12.5°C, đưa về vùng tối ưu của enzyme)\n• Hiệu suất quang hợp: Tăng từ 20% → 100% (trái dưa lưới lớn nhanh, năng suất kỷ lục)",
            "q1": "Tại sao người trồng dưa lưới nhà kính công nghệ cao lại chủ động bơm thêm khí CO₂ vào buổi sáng sớm (6h - 9h)? Mức CO₂ tăng lên có tác dụng gì đối với quang hợp?",
            "q2": "Nếu chỉ làm giàu CO₂ mà không bật quạt đối lưu và phun sương hạ nhiệt độ (để nhiệt độ >40°C) thì năng suất dưa lưới có đạt đỉnh được không? Vì sao?"
        },
        {
            "code": "2C",
            "group": "NHÓM 3",
            "title": "TRẠM 3: SÂN THƯỢNG NẮNG GẮT & VƯỜN MÁI NHÀ PHỐ TP.HCM",
            "context": "Sân thượng nhà phố hướng Tây tại TP.HCM. Trưa hè nắng chang chang (>55.000 Lux), sàn bê tông hấp thụ nhiệt nóng rực lên đến 43°C. Gió khô nóng thổi mạnh làm mất nước cực nhanh; mép lá cây bị cháy xém vàng giòn dù sáng nào cũng được tưới đẫm nước.",
            "mission": "Thiết kế giải pháp giải cứu cây trồng sân thượng, chống cháy lá và biến vườn mái thành lá phổi xanh điều hòa nhiệt độ cho toàn bộ ngôi nhà.",
            "steps": [
                ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp (Sàn bê tông nóng rực, nắng gắt)", "420 ppm", "55.000 Lux (Quá ngưỡng ⚠️)", "43.0°C (Cực đoan ⚠️)", "10% (Nguy kịch)", "Hiện tượng quang ức chế (quá tải quang năng); diệp lục bị phá hủy, mép lá khô xém giòn rụm."),
                ("Bước 1", "Đưa Giải pháp 1: Giàn cây leo Sử quân tử & Bàng cảnh che bóng tầng cao", "430 ppm", "35.000 Lux", "38.0°C", "40% (Giảm sốc nhiệt)", "Tán lá giàn leo cản bớt ánh nắng chói chang, tạo lớp che chắn tự nhiên cho tầng dưới."),
                ("Bước 2", "Đưa Giải pháp 2: Tấm đệm thoát nước và lớp thảm thực vật cách nhiệt sàn mái bê tông", "440 ppm", "30.000 Lux", "34.0°C", "65% (Hồi tỉnh)", "Nhiệt độ sàn bê tông giảm 8°C, chấm dứt hiện tượng rễ cây bị nung nóng trong chậu đất."),
                ("Bước 3", "Đưa Giải pháp 3: Bố trí nhóm cây ưa bóng (Lan Ý, Trầu bà, Cau tiểu trâm) dưới bóng giàn", "450 ppm", "15.000 Lux (Tán xạ dịu)", "32.0°C", "85% (Tươi tỉnh)", "Lan Ý được bảo vệ khỏi nắng gắt, phiến lá mỏng xanh mướt trở lại, không còn cháy mép lá."),
                ("Bước 4", "Đưa Giải pháp 4: Hệ thống tưới phun mưa hẹn giờ bán tự động buổi sáng sớm và chiều mát", "450 ppm", "25.000 Lux", "31.0°C (Dịu mát)", "100% (Cực đại ⚡)", "Đảm bảo bù đắp đủ lượng nước thoát hơi; khí khổng mở đều, vi khí hậu sân thượng trở thành ốc đảo xanh.")
            ],
            "delta": "• Cường độ ánh sáng: Điều tiết từ 55.000 Lux (cháy nắng) → 25.000 Lux (vùng bão hòa ánh sáng)\n• Nhiệt độ sàn và không khí: Hạ từ 43.0°C → 31.0°C (hạ 12.0°C, cứu sống hoàn toàn bộ rễ)\n• Hiệu suất quang hợp: Tăng từ 10% → 100% (vườn mái xanh mướt, cách nhiệt phòng ngủ bên dưới)",
            "q1": "Vì sao sáng nào bạn nhỏ cũng tưới no nước vào chậu nhưng Bé Lan Ý đặt trên sân thượng giữa trưa hè vẫn bị cháy mép lá khô giòn?",
            "q2": "Hãy phân tích vai trò của việc phối hợp giữa 'giàn cây che nắng' và 'hệ thống tưới nước giữ ẩm' trong việc cứu sống cây trồng sân thượng."
        },
        {
            "code": "2D",
            "group": "NHÓM 4",
            "title": "TRẠM 4: TRONG NHÀ, GÓC HỌC TẬP & HÀNH LANG RÂM MÁT",
            "context": "Không gian phòng khách, góc cầu thang và góc học tập của học sinh. Ánh sáng tự nhiên rất yếu (<250 Lux, nằm dưới điểm bù ánh sáng). Không khí tù đọng, bụi mịn bám đầy phiến lá. Cây bàng cảnh bị vàng lá, đói lả và rụng lả tả.",
            "mission": "Thiết kế giải pháp bổ sung ánh sáng quang học nhân tạo, thông gió và chăm sóc cây ưa bóng trong nhà để cây quang hợp khỏe mạnh, thanh lọc khí thở.",
            "steps": [
                ("Ban đầu (Hiện trạng)", "Chưa kích hoạt giải pháp (Tối mờ, thiếu sáng trầm trọng)", "750 ppm (Tù đọng ⚠️)", "250 Lux (Dưới điểm bù ⚠️)", "29.0°C", "5% (Đình trệ)", "Cường độ quang hợp nhỏ hơn hô hấp (tiêu hao chất hữu cơ); cây không tự nuôi sống được, lá úa rụng."),
                ("Bước 1", "Đưa Giải pháp 1: Bổ sung đèn quang phổ chuyên dụng cho cây trồng (LED 4000K bước sóng xanh/đỏ)", "700 ppm", "1.200 Lux (Vượt điểm bù)", "28.0°C", "45% (Bắt đầu quang hợp)", "Cung cấp photon năng lượng kích thích diệp lục tố tổng hợp phân tử glucose đầu tiên."),
                ("Bước 2", "Đưa Giải pháp 2: Vệ sinh lau sạch bụi bẩn bám trên bề mặt phiến lá định kỳ", "650 ppm", "1.500 Lux", "27.5°C", "70% (Khá tốt)", "Lỗ khí khổng được giải phóng khỏi bụi bẩn, tăng khả năng hấp thụ ánh sáng và trao đổi khí CO₂."),
                ("Bước 3", "Đưa Giải pháp 3: Quạt đối lưu thông gió nhẹ nhàng lưu chuyển không khí tươi", "550 ppm", "1.500 Lux", "26.0°C (Mát dịu)", "85% (Tốt)", "Tránh tích tụ nhiệt từ đèn điện, giúp lá cây thoát hơi nước tự nhiên và hấp thụ CO₂ lưu thông."),
                ("Bước 4", "Đưa Giải pháp 4: Chế độ tưới nước giữ ẩm 60%, giá thể xơ dừa thoát nước tốt chống úng rễ", "500 ppm", "1.800 Lux (Lý tưởng)", "25.5°C", "100% (Cực đại ⚡)", "Đất ẩm thoáng khí, rễ hút nước và khoáng dễ dàng; lá cây căng bóng mọng nước, thanh lọc formaldehyde.")
            ],
            "delta": "• Cường độ ánh sáng: Tăng từ 250 Lux (dưới điểm bù) → 1.800 Lux (quang hợp hiệu quả)\n• Nồng độ CO₂ và thông gió: Giảm từ 750 ppm tù đọng → 500 ppm tươi mát\n• Hiệu suất quang hợp: Tăng từ 5% → 100% (cây trong nhà chấm dứt rụng lá, phát triển xanh tốt)",
            "q1": "Tại sao Bé Bàng Cảnh để trong góc tối dù được tưới nước đầy đủ nhưng lá cứ nhạt màu dần rồi rụng lả tả?",
            "q2": "Ánh sáng từ bóng đèn huỳnh quang thông thường trong phòng học có giúp cây quang hợp tốt như ánh sáng mặt trời hoặc đèn LED chuyên dụng không? Vì sao?"
        }
    ]

def build_station_worksheets(filename="Phieu_Hoc_Tap_KHTN7_Bai23_Quang_Hop_4_Tram.docx"):
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)

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
        cell.width = Inches(7.1)
        set_cell_background(cell, "064E3B") # Emerald 900
        set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(f"PHIẾU HỌC TẬP SỐ {st['code']} — DÀNH CHO {st['group']}\n")
        r.font.bold = True
        r.font.size = Pt(13)
        r.font.color.rgb = RGBColor(255, 255, 255)
        r2 = p.add_run(f"NHIỆM VỤ THỰC NGHIỆM ĐO LƯỜNG TIẾN TRÌNH VI KHÍ HẬU & QUANG HỢP\n")
        r2.font.bold = True
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = RGBColor(167, 243, 208)
        r3 = p.add_run(st['title'])
        r3.font.bold = True
        r3.font.size = Pt(10)
        r3.font.color.rgb = RGBColor(253, 230, 138)

        # Student info bar
        p_info = doc.add_paragraph()
        p_info.paragraph_format.space_before = Pt(6)
        p_info.paragraph_format.space_after = Pt(6)
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
        p_sec1.paragraph_format.space_before = Pt(4)
        p_sec1.paragraph_format.space_after = Pt(3)
        r = p_sec1.add_run("I. KHẢO SÁT HIỆN TRẠNG & THÁCH THỨC VI KHÍ HẬU TẠI TRẠM:")
        r.font.bold = True
        r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor(27, 67, 50)

        p_desc = doc.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(0)
        p_desc.paragraph_format.space_after = Pt(4)
        r = p_desc.add_run("• Hiện trạng vi khí hậu: ")
        r.font.bold = True
        r.font.size = Pt(9.5)
        r2 = p_desc.add_run(st['context'] + "\n")
        r2.font.size = Pt(9.5)
        r3 = p_desc.add_run("• Nhiệm vụ thực nghiệm số: ")
        r3.font.bold = True
        r3.font.size = Pt(9.5)
        r3.font.color.rgb = RGBColor(16, 122, 80)
        r4 = p_desc.add_run(st['mission'])
        r4.font.size = Pt(9.5)

        # Section II: Observation & Measurement Table (5 Steps)
        p_sec2 = doc.add_paragraph()
        p_sec2.paragraph_format.space_before = Pt(6)
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
        # Col widths: 1.0, 1.8, 0.75, 0.85, 0.75, 0.75, 1.2 => sum = 7.1 inches
        col_widths = [Inches(0.9), Inches(1.85), Inches(0.75), Inches(0.85), Inches(0.70), Inches(0.75), Inches(1.30)]
        table = doc.add_table(rows=6, cols=7)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_borders(table, color="9CA3AF", sz="4")

        headers = [
            "Bước thực nghiệm",
            "Giải pháp kỹ thuật đưa vào\n(Kéo thả trên web)",
            "Nồng độ CO₂\n(ppm)",
            "Ánh sáng\n(Lux)",
            "Nhiệt độ\n(°C)",
            "Hiệu suất quang hợp",
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
                bg_color = "FEF2F2" # Red tint for initial poor condition
            elif step_idx == 4:
                bg_color = "ECFDF5" # Green tint for optimal 100%

            for c_idx in range(7):
                c = row.cells[c_idx]
                c.width = col_widths[c_idx]
                set_cell_background(c, bg_color)
                set_cell_margins(c, top=70, bottom=70, left=70, right=70)
                p = c.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                
                # Center align numerical/status columns, left align solution & observation
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
        p_sec3.paragraph_format.space_before = Pt(6)
        p_sec3.paragraph_format.space_after = Pt(2)
        r = p_sec3.add_run("III. ĐÁNH GIÁ MỨC ĐỘ CẢI THIỆN VI KHÍ HẬU (Δ SAU KHI HOÀN TẤT - BAN ĐẦU):")
        r.font.bold = True
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(27, 67, 50)

        p_delta = doc.add_paragraph()
        p_delta.paragraph_format.space_before = Pt(0)
        p_delta.paragraph_format.space_after = Pt(4)
        r = p_delta.add_run(st['delta'])
        r.font.size = Pt(9)
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
        p_sign.paragraph_format.space_before = Pt(6)
        p_sign.paragraph_format.space_after = Pt(0)
        r = p_sign.add_run("ĐÁNH GIÁ CỦA NHÓM: ....................................................          CHỮ KÝ DUYỆT CỦA GIÁO VIÊN: ")
        r.font.size = Pt(8.5)
        r.font.bold = True
        r2 = p_sign.add_run("Cô Mai Thuý Đức")
        r2.font.bold = True
        r2.font.size = Pt(8.5)
        r2.font.color.rgb = RGBColor(16, 122, 80)

    doc.save(filename)
    print(f"File {filename} created successfully!")

if __name__ == "__main__":
    build_station_worksheets()
