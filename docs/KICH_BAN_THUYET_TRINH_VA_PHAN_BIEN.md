# KỊCH BẢN THUYẾT TRÌNH & BỘ CÂU HỎI PHẢN BIỆN BẢO VỆ ĐỒ ÁN (AI301)
## ĐỀ DỰ ÁN 09: PHÂN TÍCH TIẾN ĐỘ CƠ HỘI VÀ DỰ KIẾN DOANH SỐ — CÔNG TY BÌNH MINH

---

## PHẦN 1: KỊCH BẢN THUYẾT TRÌNH LIVE DEMO (5 - 7 PHÚT)

### ⏱️ Phút 1: Đặt vấn đề & Bối cảnh Công ty Bình Minh
* **Lời thoại**:
  > *"Kính thưa Thầy/Cô và Hội đồng, nhóm em xin trình bày giải pháp AI Agent cho Đề dự án 09. Công ty Bình Minh là doanh nghiệp B2B với 10 Nhân viên Kinh doanh và khoảng 300 cơ hội đang theo dõi. Trước đây, báo cáo hằng tuần gặp 4 vấn đề nghiêm trọng:
  > 1. Trùng lặp dữ liệu do NVKD cập nhật ở nhiều tệp khác nhau.
  > 2. Cơ hội đã mất khách hoặc quá hạn nhưng NVKD quên không đóng (Zombie Deals).
  > 3. Báo cáo cộng cả doanh số đã ký hợp đồng vào cơ hội mở và dùng tỷ lệ dự kiến tùy ý.
  > 4. Chỉ ưu tiên các thương vụ tiền lớn mà bỏ qua việc xác minh thông tin.
  > Để giải quyết triệt để, nhóm em đã xây dựng hệ thống AI Agent phối hợp 3 kỹ năng nghiệp vụ cùng 2 điểm kiểm soát của Con người theo quy trình 5 bước chuẩn hóa."*

---

### ⏱️ Phút 2 - 3: Giới thiệu Sơ đồ Quy trình 5 Bước & Kiến trúc Phối hợp
*(Chiếu sơ đồ workflow hoặc mở ảnh trên giao diện: nút **"Xem Sơ Đồ Workflow Gốc"**)*
* **Lời thoại**:
  > *"Quy trình của nhóm gồm 5 bước chặt chẽ qua 3 làn trách nhiệm:
  > - **Bước 1 (Sales Ops)**: Tiếp nhận dữ liệu từ file Excel, CSV hoặc Google Sheets và khóa Ngày đối chiếu thống nhất (Cut-off Date).
  > - **Bước 2 (AI Agent - Skill 1)**: Tự động gộp cơ hội trùng, đối chiếu bảng Orders để tách doanh số đã bán khỏi cơ hội mở, và áp trọng số giai đoạn chính thức (Lead 10%, Qualification 25%, Proposal 50%, Negotiation 75%, Closing 90%).
  > - **Bước 3 (AI Agent - Skill 2)**: Tự động quét 7 loại rủi ro: quá hạn chốt, đóng băng tương tác trên 14 ngày, thiếu thông tin xác nhận và mâu thuẫn trạng thái.
  > - **Bước 4 (Con người - HITL #1)**: Trưởng phòng KD xem xét bằng chứng và duyệt các ca nghi vấn trực tiếp trên giao diện.
  > - **Bước 5 (AI Agent - Skill 3 & Quản lý duyệt cuối)**: Tính mức hoàn thành KPI theo 10 NVKD, đề xuất Top 5 cơ hội ưu tiên tuần tới kèm căn cứ và kế hoạch hành động cụ thể."*

---

### ⏱️ Phút 4: Thao tác Live Demo trên Giao diện Localhost
*(Thực hiện trực tiếp trên màn hình `http://localhost:8000`)*
1. **Nạp dữ liệu**:
   - Chọn kịch bản **"1. Trường hợp thông thường (Normal Case)"** và bấm **"Kích Hoạt AI Agent"**.
   - Chỉ cho Hội đồng thấy: 4 thẻ KPI tổng quan hiển thị tách bạch rõ:
     - **Doanh số đã ghi nhận**: Doanh số chắc chắn từ đơn hàng đã duyệt.
     - **Giá trị cơ hội còn mở**: Đã loại trừ toàn bộ đơn hàng đã chốt (không bị cộng trùng).
     - **Giá trị dự kiến (Trọng số)**: Có nhãn cảnh báo bắt buộc *"Ước tính theo quy tắc, không phải doanh số chắc chắn"*.
     - **Doanh số ảo đã loại bỏ**: Số tiền bị thổi phồng đã được Agent làm sạch.
2. **Chứng minh xử lý ngoại lệ**:
   - Chỉ vào cơ hội `OPP-007` (1.25 tỷ): *"Dù trong file cơ hội ghi Closing, nhưng vì đã có đơn hàng ORD-1001 trong Orders nên Agent tự động chuyển thành Closed Won và loại khỏi cơ hội mở"*.
   - Chỉ vào cơ hội `OPP-009`: *"Trong file ghi Negotiation, nhưng lịch sử cuộc gọi ghi khách đã chọn đối thủ nên Agent tự động phát hiện mâu thuẫn và đề xuất Closed Lost"*.

---

### ⏱️ Phút 5: Thao tác Con người kiểm duyệt (Bước 4) & Kiểm chứng Ranh giới (< 5 cơ hội)
1. **Human-in-the-Loop (Bước 4)**:
   - Vào tab **"Bước 3 & 4: Kiểm Duyệt Rủi Ro"**, chọn ca `OPP-009`, chọn thao tác: **"❌ Xác nhận Đóng (Closed Lost)"** $\rightarrow$ Bấm **"Áp Dụng Xác Minh Của Quản Lý"**.
   - Chỉ cho Thầy/Cô thấy: Báo cáo lập tức cập nhật lại và lưu vết vào **Audit Log**.
2. **Kiểm chứng Ranh giới quy tắc (< 5 cơ hội)**:
   - Chuyển sang **"Kịch bản 4: Ranh giới quy tắc (Boundary)"** $\rightarrow$ Bấm chạy.
   - Chỉ vào thông báo: *"Hệ thống xuất đúng **3/5 cơ hội hợp lệ**, hiển thị rõ lý do tại sao thiếu 2 cơ hội và liệt kê bảng các cơ hội lớn bị loại trừ (như OPP-B05 trị giá 3 tỷ bị loại vì chưa xác nhận ngân sách và người quyết định)"*.

---

### ⏱️ Phút 6: Bằng chứng Agent đọc dữ liệu thực tế (Audit Trail)
- Mở tab **"Nhật Ký Truy Vết (Audit Log)"**:
  > *"Toàn bộ các bước từ đọc dữ liệu, gọi Tool gộp trùng, lọc Cut-off date, áp trọng số và quyết định của Con người đều được ghi nhận theo chuẩn JSONL với mã Run ID và timestamp chi tiết, chứng minh kết quả là suy luận và tính toán động 100%, không hề nhập tay kết luận."*

---

## PHẦN 2: BỘ CÂU HỎI & CÂU TRẢ LỜI PHẢN BIỆN (Q&A CHEATSHEET)

### ❓ Câu 1: Làm sao chứng minh AI Agent thực sự đọc dữ liệu và tính toán theo quy tắc chứ không phải nhóm nhập tay kết luận có sẵn?
* **Trả lời**:
  > *"Dạ thưa Thầy/Cô, nhóm chứng minh qua 3 bằng chứng kỹ thuật:
  > 1. **Khả năng nạp file động**: Hệ thống có cổng nạp file `.xlsx` / `.csv` và Google Sheets trực tiếp tại Bước 1. Thầy/Cô có thể đưa bất kỳ bộ dữ liệu mới nào, Agent sẽ đọc các sheet bằng `pandas` và tính toán ngay tại chỗ.
  > 2. **Kiến trúc Tool-Calling rõ ràng**: Mỗi bước chuẩn hóa đều được module hóa thành các hàm Python độc lập (`_skill_1_consolidate_and_valuate`, `_skill_2_detect_risks_and_conflicts`, `_skill_3_evaluate_kpi_and_prioritize`).
  > 3. **Nhật ký truy vết (Audit Trail)**: Mọi thao tác đều xuất ra file `logs/agent_audit_log.jsonl` ghi rõ thời gian thực thi, quy tắc áp dụng và dữ liệu trước/sau biến đổi."*

---

### ❓ Câu 2: Tại sao cơ hội OPP-011 trị giá 2.4 tỷ rất lớn mà không được Agent chọn vào Top 5 ưu tiên tuần tới?
* **Trả lời**:
  > *"Dạ thưa Thầy/Cô, đây chính là điểm mấu chốt để giải quyết vấn đề số 4 trong đề bài: **'Xem mọi cơ hội giá trị lớn là ưu tiên dù thiếu thông tin xác nhận'**.
  > Trong hồ sơ cơ hội `OPP-011`, NVKD tuy điền giá trị 2.4 tỷ nhưng lại bỏ trống **Người quyết định**, ngân sách để **'Chưa xác nhận'** và không có **Bước tiếp theo**.
  > AI Agent áp dụng Bộ lọc hợp lệ (Eligibility Filter): chỉ những cơ hội đã xác thực đầy đủ thông tin mới được xét vào Top ưu tiên. Nếu chọn cơ hội 2.4 tỷ này thì Trưởng phòng KD sẽ dồn nguồn lực vào một 'thương vụ bánh vẽ', có nguy cơ thất bại phút chót."*

---

### ❓ Câu 3: Nếu giảng viên đưa bộ dữ liệu mới vào ngày thi thì nhóm thao tác thế nào?
* **Trả lời**:
  > *"Dạ, nhóm đã thiết kế sẵn chức năng nạp dữ liệu thi tại Bước 1:
  > - Nếu Thầy/Cô đưa 1 file Excel tổng: Nhóm chọn tab **'Tải Tệp (.XLSX / .CSV)'**, kéo file vào ô 'File Excel tổng' và bấm nút chạy.
  > - Nếu Thầy/Cô đưa các file CSV rời: Nhóm tải riêng từng file vào các ô Opportunities, Activities, Orders, KPI.
  > - Nếu Thầy/Cô đưa link Google Sheets: Nhóm dán link vào ô Google Sheets.
  > Toàn bộ quá trình chạy hoàn toàn tự động trong chưa đầy 2 giây mà không cần sửa bất kỳ dòng code nào."*

---

### ❓ Câu 4: Vì sao cần 2 điểm Con người kiểm duyệt (Human-in-the-loop) mà không để AI tự động 100%?
* **Trả lời**:
  > *"Dạ, trong nghiệp vụ B2B, mối quan hệ khách hàng rất phức tạp và có những yếu tố định tính mà dữ liệu chưa phản ánh hết:
  > - **Tại Điểm 1 (Bước 4 - Xử lý ca nghi vấn)**: Ví dụ ghi chú cuộc gọi ghi 'Khách tạm hoãn 1 tuần để xin chữ ký', AI có thể cảnh báo nguy cơ mất khách, nhưng Trưởng phòng KD biết rõ khách hàng này quen thuộc nên có quyền chọn 'Xác nhận giữ mở'. Con người giữ vai trò phán quyết cuối cùng để tránh AI hiểu sai ngữ cảnh.
  > - **Tại Điểm 2 (Cuối Bước 5 - Phê duyệt giao việc)**: AI đề xuất Top 5 dựa trên điểm số toán học, nhưng việc phân công ai hỗ trợ (ví dụ cử Trưởng nhóm Presales nào đi cùng NVKD nào) phụ thuộc vào năng lực thực tế của nhân sự, cần Trưởng phòng quyết định."*

---

### ❓ Câu 5: Hệ thống xử lý thế nào khi ngày dự kiến chốt hoặc ngày hoạt động rơi đúng vào Ngày đối chiếu?
* **Trả lời**:
  > *"Dạ, đây là trường hợp ranh giới (Boundary Case) nhóm đã kiểm thử rất kỹ:
  > - Nếu `expected_close_date == report_date`: Đúng ngày báo cáo, cơ hội **không bị tính là quá hạn** (còn 0 ngày).
  > - Nếu `expected_close_date < report_date`: Dù chỉ trước 1 ngày thì vẫn bị tính là **quá hạn chốt**.
  > - Nếu khoảng cách liên hệ đúng bằng 14 ngày: Nằm trong ngưỡng cho phép; nếu từ 15 ngày trở lên sẽ bị gắn cờ **đóng băng tương tác**.
  > - Mọi đơn hàng hoặc hoạt động phát sinh sau ngày đối chiếu (`> report_date`) đều bị bộ lọc Cut-off Date tách ra ngoài để đảm bảo tính toàn vẹn của kỳ báo cáo."*

---

### ❓ Câu 6: Làm thế nào để nhóm tự tin hệ thống không gặp lỗi khi giảng viên kiểm thử?
* **Trả lời**:
  > *"Dạ, nhóm đã xây dựng module kiểm thử tự động `test_runner.py` tích hợp sẵn trên giao diện (nút **'Chạy Kiểm Thử 3 Kịch Bản'**).
  > Module này tự động chạy và kiểm chứng qua 11 điều kiện logic khắt khe trên cả 3 bộ dữ liệu: Normal Case, Missing/Conflict Case, và Boundary Case. Hiện tại cả 3 kịch bản đều đạt **100% PASS**."*
