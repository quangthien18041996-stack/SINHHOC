const fs = require('fs');

let html = fs.readFileSync('index.html', 'utf8');

const newStationWorksheetData = `const stationWorksheetData = {
            'station-1': {
                code: '2A',
                team: 'Nhóm 1',
                stationName: 'TRẠM 1: ĐÔ THỊ NÉN & THÁP RỪNG THẲNG ĐỨNG',
                context: 'Không gian chật hẹp, nồng độ CO₂ cao do khí thải xe cộ, nhiệt độ hấp thụ từ bê tông cao (hiệu ứng đảo nhiệt). Tầng trên nắng gắt, tầng trệt thiếu sáng.',
                mission: 'Ứng dụng mô hình phân tầng sinh thái Bosco Verticale: Tuyển chọn và kích hoạt từng giải pháp cây xanh, hệ thống tưới tự động để hạ nhiệt, lọc không khí và tối ưu hóa quang hợp.',
                steps: [
                    { step: 'Ban đầu (Hiện trạng)', sol: 'Chưa kích hoạt giải pháp nào', co2: '850 ppm', light: '45,000 Lux', temp: '36.0°C', eff: 'Chưa tối ưu (0%)', obs: 'Hiện trạng ban đầu: Bê tông giữ nhiệt, khói xe dày đặc, quang hợp bị ức chế nghiêm trọng.' },
                    { step: 'Bước 1', sol: 'Trồng cây ưa sáng có tầng cutin dày (Hoa giấy, Trúc quân tử)', co2: '730 ppm', light: '45,000 Lux', temp: '33.5°C', eff: 'Kém (25%)', obs: 'Cây tầng thượng có tán lá dày giúp cản 15.000 Lux nắng gắt, giảm bức xạ nhiệt xuống tầng dưới.' },
                    { step: 'Bước 2', sol: 'Bố trí vườn tường đứng nhiều tầng (Vertical Wall) bằng giàn cây leo bản rộng', co2: '580 ppm', light: '30,000 Lux', temp: '31.5°C', eff: 'Trung bình (50%)', obs: 'Giàn cây leo che phủ bê tông làm mát tường thêm 2.5°C và hấp thụ 120 ppm khí CO₂.' },
                    { step: 'Bước 3', sol: 'Trồng cây ưa bóng (Lưỡi hổ, Lan ý) có khả năng lọc formaldehyde', co2: '460 ppm', light: '30,000 Lux', temp: '30.5°C', eff: 'Khá tốt (75%)', obs: 'Cây ưa bóng tầng trệt hấp thụ CO₂ khói xe và lọc formaldehyde mà không bị cháy lá.' },
                    { step: 'Bước 4', sol: 'Cài đặt van tưới nhỏ giọt hẹn giờ tự động kết hợp khay hứng tràn', co2: '410 ppm', light: '25,000 Lux', temp: '28.0°C', eff: 'Tối ưu 100% 🌿', obs: 'Cung cấp nước đều đặn làm khí khổng mở tối đa, CO₂ khuếch tán vào mô lá đạt đỉnh quang hợp 100%.' }
                ],
                delta: '• Nồng độ CO₂: Giảm từ 850 ppm → 410 ppm (-440 ppm, giảm ô nhiễm khí thải)\\n• Cường độ ánh sáng: Điều tiết từ 45,000 Lux → 25,000 Lux (cản bớt nắng gắt tầng trên, tán xạ xuống trệt)\\n• Nhiệt độ vi khí hậu: Hạ từ 36.0°C → 28.0°C (hạ 8.0°C, triệt tiêu hiệu ứng đảo nhiệt)\\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (cả 4 tầng cây xanh tốt)',
                questions: [
                    'Sự thay đổi của các thông số đo lường (CO₂, Ánh sáng, Nhiệt độ, Nước) sau khi cho từng giải pháp vào đã tác động như thế nào đến sự đóng/mở khí khổng và hiệu suất quang hợp của cây xanh tại trạm?',
                    'Từ kết quả thực nghiệm đo lường trên, nhóm hãy rút ra quy luật sinh học về sự phối hợp đồng bộ giữa 4 yếu tố môi trường (Ánh sáng, Nước, CO₂, Nhiệt độ) đối với quá trình quang hợp.'
                ]
            },
            'station-2': {
                code: '2B',
                team: 'Nhóm 2',
                stationName: 'TRẠM 2: NHÀ MÀNG NÔNG NGHIỆP THÔNG MINH',
                context: 'Yêu cầu năng suất và chất lượng quả tối đa, cần chủ động kiểm soát cả 4 yếu tố quang hợp quanh năm bất chấp thời tiết bất lợi bên ngoài.',
                mission: 'Thiết lập hệ thống cảm biến vi khí hậu và kích hoạt từng giải pháp nông nghiệp chính xác (đèn LED, làm giàu CO₂, quạt làm mát, tưới nhỏ giọt) để thúc đẩy quang hợp đạt cực đại.',
                steps: [
                    { step: 'Ban đầu (Hiện trạng)', sol: 'Chưa kích hoạt giải pháp nào', co2: '400 ppm (Tự nhiên)', light: '12.000 Lux (Thiếu)', temp: '34.0°C', eff: 'Chưa tối ưu (0%)', obs: 'Hiện trạng ban đầu: Cây phụ thuộc thời tiết bên ngoài, thiếu sáng và nhiệt độ bí nóng.' },
                    { step: 'Bước 1', sol: 'Lắp đặt đèn LED chuyên dụng phát dải sáng xanh (450nm) và đỏ (660nm)', co2: '400 ppm (Tự nhiên)', light: '35.000 Lux (LED)', temp: '34.0°C', eff: 'Kém (25%)', obs: 'Đèn LED đa phổ cung cấp đúng 2 bước sóng 450nm và 660nm, kích thích mạnh pha sáng.' },
                    { step: 'Bước 2', sol: 'Bơm bổ sung khí CO₂ duy trì ở ngưỡng tối ưu 0.10% – 0.15% (1.000 - 1.500 ppm)', co2: '1.200 ppm (Tối ưu)', light: '35.000 Lux (LED)', temp: '34.0°C', eff: 'Trung bình (50%)', obs: 'Bơm CO₂ bổ sung lên 1.200 ppm làm tăng tốc độ tổng hợp đường bột trước điểm bão hòa.' },
                    { step: 'Bước 3', sol: 'Lưới cắt nắng tự động kích hoạt kết hợp quạt hút và màng nước làm mát', co2: '1.200 ppm (Tối ưu)', light: '35.000 Lux (LED)', temp: '28.0°C', eff: 'Khá tốt (75%)', obs: 'Quạt hút gió và màng nước hạ nhiệt về 27°C, bảo vệ các enzyme quang hợp hoạt động cực đại.' },
                    { step: 'Bước 4', sol: 'Tưới dinh dưỡng nhỏ giọt tự động kiểm soát theo độ ẩm đất đo bằng cảm biến IoT', co2: '1.200 ppm (Tối ưu)', light: '35.000 Lux (LED)', temp: '26.0°C', eff: 'Tối ưu 100% 🌿', obs: 'Tưới nhỏ giọt IoT cung cấp nước liên tục, duy trì tế bào khí khổng mở suốt ngày đêm.' }
                ],
                delta: '• Nồng độ CO₂: Tăng từ 400 ppm → 1.200 ppm (làm giàu nguyên liệu cho pha tối quang hợp)\\n• Cường độ ánh sáng: Tăng từ 12.000 Lux → 35.000 Lux (đèn LED chuyên dụng dải xanh 450nm & đỏ 660nm)\\n• Nhiệt độ vi khí hậu: Hạ từ 34.0°C → 26.0°C (hạ 8.0°C, đưa về vùng tối ưu của enzyme)\\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (năng suất và độ ngọt đạt tối đa)',
                questions: [
                    'Tại sao người trồng dưa lưới trong nhà màng công nghệ cao lại chủ động bơm thêm khí CO₂ lên 1.200 ppm vào buổi sáng? Mức CO₂ tăng lên có tác dụng gì đối với quang hợp?',
                    'Nếu chỉ làm giàu CO₂ và bổ sung đèn LED mà không bật quạt làm mát (để nhiệt độ >38°C) thì cây dưa lưới có quang hợp tốt không? Vì sao?'
                ]
            },
            'station-3': {
                code: '2C',
                team: 'Nhóm 3',
                stationName: 'TRẠM 3: SÂN THƯỢNG NẮNG GẮT & NHIỆT ĐỘ CAO',
                context: 'Nắng trưa gay gắt > 55.000 Lux, nhiệt độ sàn bê tông lên tới 43°C. Khí khổng đóng chặt tránh mất nước làm ngừng quang hợp, lá dễ bị \\'đốt nóng\\' và cháy khô.',
                mission: 'Thiết kế giải pháp giải cứu vườn sân thượng: Căng lưới lan, tuyển chọn cây chịu hạn cutin dày, phun sương sàn và phủ gốc giữ ẩm để tái mở khí khổng và phục hồi quang hợp.',
                steps: [
                    { step: 'Ban đầu (Hiện trạng)', sol: 'Chưa kích hoạt giải pháp nào', co2: 'Khí khổng đóng', light: '58,000 Lux', temp: '43.0°C', eff: 'Chưa tối ưu (0%)', obs: 'Hiện trạng ban đầu: Nắng trưa >55.000 Lux làm đốt nóng lá, nhiệt độ 43°C khiến khí khổng đóng chặt.' },
                    { step: 'Bước 1', sol: 'Căng lưới lan đen che chắn giảm 40 - 50% ánh sáng trong khung giờ 11h - 14h', co2: 'Khí khổng đóng', light: '28,000 Lux', temp: '39.0°C', eff: 'Kém (25%)', obs: 'Lưới lan đen giảm 50% cường độ bức xạ, ngăn chặn phá hủy diệp lục tố của lá.' },
                    { step: 'Bước 2', sol: 'Trồng các giống cây chịu hạn, ưa sáng mạnh (Hoa sứ Thái, Hoa giấy, Xương rồng)', co2: 'Khí khổng đóng', light: '28,000 Lux', temp: '38.0°C', eff: 'Trung bình (50%)', obs: 'Cây hoa sứ và xương rồng có tầng cutin dày, thích nghi và bắt đầu quang hợp dưới nhiệt độ cao.' },
                    { step: 'Bước 3', sol: 'Phun sương làm mát bề mặt sàn bê tông vào giữa trưa để hạ nhiệt độ về dưới 33°C', co2: 'Khí khổng mở', light: '28,000 Lux', temp: '33.0°C', eff: 'Khá tốt (75%)', obs: 'Phun sương sàn hạ nhiệt độ mặt bê tông xuống 34°C, kích thích tế bào khí khổng hé mở lại.' },
                    { step: 'Bước 4', sol: 'Phủ lớp vỏ trấu / sỏi nhẹ giữ ẩm đất chậu và chỉ tưới vào sáng sớm hoặc chiều mát', co2: 'Khí khổng mở', light: '28,000 Lux', temp: '31.0°C', eff: 'Tối ưu 100% 🌿', obs: 'Lớp phủ rơm rạ và sỏi nhẹ ngăn bốc hơi nước rễ, giúp cây quang hợp bền bỉ cả ngày.' }
                ],
                delta: '• Cường độ ánh sáng: Điều tiết từ 58,000 Lux (nắng gắt cháy lá) → 28,000 Lux (vùng an toàn)\\n• Trạng thái khí khổng: Từ đóng chặt chống mất nước → Khí khổng hé mở và mở hoàn toàn đón CO₂\\n• Nhiệt độ vi khí hậu: Hạ từ 43.0°C sàn nung nóng → 31.0°C (hạ 12.0°C bảo vệ tế bào và enzyme)\\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (vườn sân thượng xanh tươi, dịu mát)',
                questions: [
                    'Vì sao sáng nào bạn nhỏ cũng tưới no nước vào chậu nhưng Bé Lan Ý đặt trên sân thượng giữa trưa hè vẫn bị cháy mép lá khô giòn và ngừng quang hợp?',
                    'Hãy giải thích vai trò của việc phun sương làm mát sàn bê tông đối với việc kích thích khí khổng mở ra trong quá trình quang hợp.'
                ]
            },
            'station-4': {
                code: '2D',
                team: 'Nhóm 4',
                stationName: 'TRẠM 4: TRONG NHÀ & HÀNH LANG RÂM MÁT',
                context: 'Ánh sáng rất yếu (< 1.500 Lux, dưới điểm bù ánh sáng của cây ưa sáng), thiếu thông gió tự nhiên, nồng độ CO₂ tăng cao do hơi thở con người.',
                mission: 'Thiết kế giải pháp vi khí hậu trong nhà: Tuyển chọn cây ưa bóng giàu diệp lục b, đo độ rọi bằng smartphone, bổ sung đèn LED 4000K và vệ sinh bụi lá để khôi phục quang hợp và thanh lọc không khí.',
                steps: [
                    { step: 'Ban đầu (Hiện trạng)', sol: 'Chưa kích hoạt giải pháp nào', co2: '650 ppm', light: '800 Lux', temp: '26.0°C', eff: 'Chưa tối ưu (0%)', obs: 'Hiện trạng ban đầu: Ánh sáng <800 Lux dưới điểm bù quang hợp, cây tiêu hao nhiều hơn tạo ra.' },
                    { step: 'Bước 1', sol: 'Trang bị đèn rọi LED ánh sáng trắng ấm 4000K (15W) chiếu bổ sung 3 - 4 giờ/ngày', co2: '650 ppm', light: '1,500 Lux', temp: '26.0°C', eff: 'Kém (25%)', obs: 'App Lux Meter định vị được vị trí đón ánh sáng tán xạ 1.500 Lux gần cửa sổ.' },
                    { step: 'Bước 2', sol: 'Tuyển chọn 100% cây ưa bóng có phiến lá to mỏng, giàu diệp lục b (Lá lốt, Trầu không, Kim tiền)', co2: '590 ppm', light: '1,500 Lux', temp: '26.0°C', eff: 'Trung bình (50%)', obs: 'Cây lá lốt ưa bóng với phiến lá to, mỏng, chứa nhiều diệp lục b bắt trọn ánh sáng yếu.' },
                    { step: 'Bước 3', sol: 'Dùng app \\'Lux Meter\\' quét chọn các vị trí gần cửa sổ / giếng trời có độ rọi > 1.500 Lux', co2: '590 ppm', light: '2,500 Lux', temp: '26.0°C', eff: 'Khá tốt (75%)', obs: 'Đèn rọi LED 15W bổ sung thêm 1.000 Lux, kích thích pha sáng diễn ra bình thường.' },
                    { step: 'Bước 4', sol: 'Định kỳ lau sạch bụi bám trên bề mặt lá bằng khăn ẩm mềm', co2: '450 ppm', light: '2,500 Lux', temp: '26.0°C', eff: 'Tối ưu 100% 🌿', obs: 'Lau sạch bụi bám trên bề mặt lá giúp lỗ khí khổng thông thoáng, CO₂ khuếch tán dễ dàng.' }
                ],
                delta: '• Cường độ ánh sáng: Tăng từ 800 Lux (dưới điểm bù) → 2,500 Lux (vượt điểm bù quang hợp)\\n• Nồng độ CO₂: Giảm từ 650 ppm (khí thở tù đọng) → 450 ppm (cây quang hợp tích cực, thanh lọc khí thở)\\n• Nhiệt độ vi khí hậu: Ổn định ở mức 26.0°C (mát mẻ, phù hợp tối ưu cho enzyme quang hợp trong nhà)\\n• Hiệu quả quang hợp: Tăng vọt từ Chưa tối ưu (0%) → Tối ưu 100% 🌿 (cây phát triển xanh tốt, lá bóng mọng)',
                questions: [
                    'Tại sao cây đặt sâu trong phòng tối dù được tưới nước đầy đủ nhưng lá cứ nhạt màu dần rồi rụng lả tả? (Liên hệ khái niệm điểm bù ánh sáng).',
                    'Vì sao việc dùng khăn mềm ẩm lau sạch bụi bám trên phiến lá lại giúp cây trong nhà tăng cường hiệu quả quang hợp?'
                ]
            }
        };`;

const startIdx = html.indexOf('const stationWorksheetData =');
const endIdx = html.indexOf('let currentActiveWorksheetTab =');

if (startIdx === -1 || endIdx === -1) {
    console.error('Could not find stationWorksheetData markers in index.html');
    process.exit(1);
}

const updatedHtml = html.slice(0, startIdx) + newStationWorksheetData + '\n\n        ' + html.slice(endIdx);
fs.writeFileSync('index.html', updatedHtml, 'utf8');
console.log('Successfully updated stationWorksheetData in index.html!');
