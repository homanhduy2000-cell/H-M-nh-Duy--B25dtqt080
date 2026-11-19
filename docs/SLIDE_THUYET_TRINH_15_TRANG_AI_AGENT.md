# 🎓 BỘ 15 SLIDE THUYẾT TRÌNH: CÁCH XÂY DỰNG DUYDUDI AI AGENT
**Dự án:** Hệ thống AI Agent Phân tích Tiến độ Cơ hội & Dự báo Doanh số B2B  
**Đơn vị:** Công ty của Duydudi • Nền tảng Duydudi AI Sales Ops (`http://localhost:8000`)

---

## 📑 SLIDE 1: TRANG BÌA (TITLE SLIDE)
### **Xây Dựng Duydudi AI Agent: Phân Tích Tiến Độ Cơ Hội & Dự Báo Doanh Số B2B**
- **Phụ đề:** Từ Dữ liệu B2B Phân mảnh đến Quyết định Điều hành Chính xác với Kiến trúc *Deterministic Python + Google Gemini AI*
- **Thương hiệu:** **Công ty của Duydudi** (`Duydudi • AI Sales Ops`)
- **Công nghệ cốt lõi:** Python FastAPI • Pandas Engine • Google Gemini 3.5 Flash • Chrome/Smallpdf Minimalist UI
- **Demo trực tiếp:** `http://localhost:8000`

---

## 📑 SLIDE 2: BÀI TOÁN THỰC TẾ & NỖI ĐAU NGHIỆP VỤ (PROBLEM STATEMENT)
### **Tại Sao Báo Cáo Doanh Số B2B Truyền Thống Thường Sai Lệch?**
- **1. Dữ liệu phân mảnh đa nguồn:** Nằm rời rạc ở 4 bảng (`Opportunities`, `Activities`, `Orders`, `KPI Targets`) trên nhiều file Excel, CSV và Google Sheets.
- **2. Trùng lặp & Mâu thuẫn (Data Conflicts):** Cùng một dự án nhưng 2 nhân viên báo cáo khác nhau, hoặc khách đã ký đơn hàng (`Orders`) nhưng cơ hội vẫn để mở -> **Tính trùng doanh số**.
- **3. "Cơ hội xác sống" (Zombie Deals):** Lịch sử liên hệ (`Activities`) ghi khách đã từ chối/chọn đối thủ, hoặc quá hạn chốt nhiều tuần không tương tác nhưng vẫn nằm trong dự báo.
- **4. Xác suất cảm tính (Arbitrary Probability):** Sales tự điền tỷ lệ thắng 80%–90% theo cảm tính làm **phồng ảo dự báo doanh số** toàn công ty.

---

## 📑 SLIDE 3: TRIẾT LÝ THIẾT KẾ "DETERMINISTIC + LLM" (DESIGN PHILOSOPHY)
### **Chống Ảo Giác Số Liệu (Zero-Hallucination Architecture)**
- **Nguyên tắc vàng:** Tuyệt đối **không** dùng LLM để cộng trừ toàn bộ bảng dữ liệu thô (tránh tràn token và sai số học).
- **Sự kết hợp 2 tầng tư duy:**
  1. **Deterministic Python Engine (`agent_engine.py`):** Đảm nhận 100% việc lọc ngày chốt sổ, khử trùng lặp, đối chiếu chéo 4 bảng, áp công thức trọng số và chấm điểm ưu tiên.
  2. **Generative AI Planner (`Google Gemini 3.5 Flash`):** Nhận bảng chỉ số sạch (JSON Metadata) để tổng hợp bản **Nhận xét Điều hành (Executive Brief)** và khuyến nghị hành động.
- **Minh bạch hóa dự báo:** Mọi con số dự báo đều gắn nhãn *"Ước tính theo quy tắc trọng số, không phải doanh số chắc chắn"*.

---

## 📑 SLIDE 4: KIẾN TRÚC TỔNG THỂ 4 TẦNG END-TO-END (SYSTEM ARCHITECTURE)
### **Sơ Đồ Luồng Xử Lý Toàn Diện Của Duydudi AI Agent**
- **Tầng 1 (Ingestion Gateway):** Tiếp nhận file `.xlsx`, `.csv`, `.zip` từ giao diện Smallpdf, link Google Sheets trực tuyến và Webhook API bên ngoài.
- **Tầng 2 (Validation & Quarantine):** Tự động nhận diện 4 bảng dữ liệu, lọc mốc chốt sổ (`report_date`), khử trùng lặp và cách ly bản ghi lỗi ra `data_quarantine.csv`.
- **Tầng 3 (Agent Reasoning & HITL):** Thực thi 3 Kỹ năng (Skills) kết hợp điểm dừng kiểm duyệt của Quản lý tại Bước 4 và tổng hợp nhận xét bằng Gemini 3.5 Flash.
- **Tầng 4 (Output & Delivery):** Trực quan hóa trên Studio Dashboard, xuất file CSV/Excel sạch và lưu Nhật ký truy vết (`agent_audit_log.jsonl`).

---

## 📑 SLIDE 5: QUY TRÌNH PHỐI HỢP 5 BƯỚC GIỮA CON NGƯỜI & AI AGENT
### **Thiết Kế Vòng Lặp Cộng Tác (Human-AI Collaboration Loop)**
1. **Bước 1 (Con người / API):** Tải tệp Excel/CSV/ZIP hoặc dán link Google Sheets và thiết lập Ngày đối chiếu (`report_date`).
2. **Bước 2 (AI Agent - Skill 1):** Chuẩn hóa thời gian, gộp cơ hội trùng lặp giữa các file và tính giá trị dự kiến theo trọng số giai đoạn.
3. **Bước 3 (AI Agent - Skill 2):** Đối chiếu chéo 3 bảng (`Opportunities` - `Activities` - `Orders`) để phát hiện 8 nhóm rủi ro, quá hạn và mâu thuẫn.
4. **Bước 4 (Con người - HITL):** Trưởng phòng Kinh doanh xem xét danh sách nghi vấn và ra quyết định (`Giữ mở`, `Chốt Won`, hoặc `Đóng Lost`).
5. **Bước 5 (AI Agent - Skill 3 & Gemini):** Tính mức hoàn thành KPI của 10 Sales, chấm điểm xếp hạng Top 5 cơ hội ưu tiên tuần tới và sinh Báo cáo điều hành.

---

## 📑 SLIDE 6: TẦNG 1 — XÂY DỰNG CỔNG TIẾP NHẬN ĐA NGUỒN (DATA INGESTION)
### **Làm Sao Để Agent Tự Động Đọc Mọi Định Dạng File & Google Sheets?**
- **1. Bộ phân tích Excel & ZIP đa bảng (`parse_excel_bytes`):**
  - Quét tên Sheet tự động (`Opportunities`, `Activities`, `Orders`, `KPI_Targets`).
  - **Fallback thông minh:** Nếu người dùng đặt tên Sheet tùy ý (`Sheet1`, `Data`), Agent tự động quét **tên cột đặc trưng** (`deal_value`, `activity_date`, `order_amount`, `kpi_target`) để ghép đúng bảng.
- **2. Trình kết nối Google Sheets trực tiếp (`load_from_google_sheets_url`):**
  - Dùng Regex bóc tách `sheet_id` từ link chia sẻ `/spreadsheets/d/<ID>/edit` và chuyển đổi sang luồng `/export?format=xlsx` để tải trọn bộ nhiều Sheet cùng lúc.
- **3. Cổng Webhook tích hợp hệ thống ngoài (`POST /api/webhook/ingest`):**
  - Cho phép CRM/POS/n8n đẩy file tự động qua HTTP Multipart mà không cần mở trình duyệt.

---

## 📑 SLIDE 7: SKILL 1 — LỌC NGÀY ĐỐI CHIẾU & KHỬ TRÙNG LẶP ĐA TỆP
### **Chuẩn Hóa Thời Gian & Hợp Nhất Bản Ghi Trùng (`_skill_1_consolidate_and_valuate`)**
- **1. Bộ lọc mốc chốt sổ (`RULE-CUTOFF-FILTER`):**
  - Mọi đơn hàng (`order_date > report_date`) hoặc hoạt động (`activity_date > report_date`) phát sinh sau ngày báo cáo đều bị tách riêng vào `cutoff_filtered`, đảm bảo tính nhất quán theo đúng thời điểm đối chiếu.
- **2. Thuật toán Hợp nhất Trùng lặp Ngữ nghĩa (`RULE-DEDUP`):**
  - Tạo khóa kép: kết hợp mã `opp_id` và khóa ngữ nghĩa `lower(customer_name) || lower(project_name)`.
  - Khi phát hiện nhiều bản ghi của cùng 1 dự án từ các file khác nhau:
    - Tự động chọn bản ghi có `updated_date` mới nhất (`<= report_date`).
    - Ghi nhận cảnh báo nếu lệch nhân viên phụ trách (`rep_id`) hoặc chênh lệch giá trị >= 20%.

---

## 📑 SLIDE 8: SKILL 1 — TÁCH BẠCH DOANH SỐ & CHUẨN HÓA TRỌNG SỐ
### **Chấm Dứt Tình Trạng Tính Trùng & Báo Cáo Xác Suất Cảm Tính**
- **Nguyên tắc tách bạch tuyệt đối 3 đại lượng:**
  - **Doanh số thực tế đã ghi nhận (`actual_recorded_revenue`):** Chỉ lấy từ đơn hàng hợp lệ (`Orders` / `Closed Won`). Khi đã chốt Won -> Giá trị mở = `0`, Dự báo = `0`.
  - **Giá trị cơ hội đang mở (`open_pipeline_value`):** Tổng giá trị gốc của các cơ hội chưa đóng.
  - **Doanh số dự kiến ước tính (`weighted_forecast_value`):** `Giá trị cơ hội mở * Trọng số giai đoạn chuẩn`.
- **Bảng trọng số quy tắc chuẩn hóa (`STAGE_RULES`):**
  - `Lead`: **10%** | `Qualification`: **25%** | `Proposal`: **50%** | `Negotiation`: **75%** | `Closing`: **90%**
  - `Closed Won`: Tách sang doanh số thực tế | `Closed Lost`: **0%**

---

## 📑 SLIDE 9: SKILL 2 — ĐỘNG CƠ QUÉT 8 NHÓM RỦI RO & MÂU THUẪN DỮ LIỆU
### **Đối Chiếu Chéo 3 Bảng Để Bắt Lỗi Nghiệp Vụ (`_skill_2_detect_risks_and_conflicts`)**
1. **`CONFLICT_ORDER_STATUS` (Critical):** Đã có đơn hàng bên bảng `Orders` nhưng cột `stage` ở bảng Cơ hội vẫn để `Negotiation`/`Closing` (nguy cơ tính trùng).
2. **`CONFLICT_ZOMBIE_LOST` (Critical):** Nhật ký `Activities` có tín hiệu từ chối (*"mất khách", "đối thủ", "hủy dự án"*) nhưng trạng thái vẫn để Mở.
3. **`CONFLICT_DUPLICATE_MERGE` (High):** Gộp từ nhiều file nhưng lệch tên Sales hoặc lệch giá trị >= 20%.
4. **`OVERDUE_CLOSE_DATE` (High):** Ngày dự kiến chốt `expected_close_date < report_date` mà chưa đóng.
5. **`INACTIVE_STALE` (Medium/High):** Đóng băng tương tác quá ngưỡng quy định (`days_since_activity > 14 ngày`).
6. **`MISSING_VERIFICATION_INFO` (Critical/Medium):** Cơ hội mở (đặc biệt >= 500 triệu) nhưng thiếu *Người quyết định*, *Ngân sách* hoặc *Bước tiếp theo*.
7. **`MISSING_STAGE` / `MISSING_CLOSE_DATE` (High):** Bỏ trống giai đoạn hoặc ngày dự kiến chốt.
8. **`ARBITRARY_PROBABILITY` (Low):** Tỷ lệ tự điền của Sales lệch >= 20% so với quy tắc chuẩn.

---

## 📑 SLIDE 10: VÙNG CÁCH LY (QUARANTINE ZONE) & NHẬT KÝ TRUY VẾT (AUDIT LOG)
### **Đảm Bảo Tính Minh Bạch & Khả Năng Giải Trình (Explainable AI)**
- **1. Vùng cách ly dữ liệu lỗi (`Quarantine Zone` - `/api/export/quarantine.csv`):**
  - Toàn bộ bản ghi mâu thuẫn trạng thái, quá hạn, thiếu thông tin hoặc sau ngày chốt sổ được trích xuất thành file `data_quarantine.csv` kèm cột lý do chi tiết và hướng xử lý đề xuất.
  - Dữ liệu sạch sau chuẩn hóa được xuất riêng tại `/api/export/cleaned.csv`.
- **2. Hệ thống Nhật ký Truy vết (`AuditLogger` - `agent_audit_log.jsonl`):**
  - Mọi bước thực thi của Agent đều ghi lại: `run_id`, `timestamp`, `step`, `skill_or_tool`, `rule_applied` và `details`.
  - Chứng minh Agent xử lý dựa trên dữ liệu thực và quy tắc rõ ràng, không phải "hộp đen".

---

## 📑 SLIDE 11: BƯỚC 4 — CƠ CHẾ CON NGƯỜI KIỂM DUYỆT (HUMAN-IN-THE-LOOP)
### **Trao Quyền Quyết Định Cuối Cùng Cho Trưởng Phòng Kinh Doanh**
- **Tại sao AI không tự ý xóa dữ liệu?**
  - Trong B2B, một cơ hội quá hạn ngày chốt có thể do khách hàng đang chờ ký phụ lục, hoặc đã mất thật. Con người cần xác nhận thực tế.
- **Cách lập trình cơ chế `human_overrides`:**
  - Giao diện cung cấp bảng Kiểm duyệt Rủi ro với 3 quyền quyết định cho từng mã cơ hội (`opp_id`):
    - `SET_CLOSED_LOST`: Xác nhận mất khách -> Trọng số về `0%`.
    - `SET_CLOSED_WON`: Xác nhận đã ký kết -> Chuyển sang Doanh số thực tế.
    - `VERIFY_AND_KEEP_OPEN`: Đã xác minh bổ sung thông tin & gia hạn -> Giữ mở và cho phép xét ưu tiên.
  - Khi bấm **"Áp Dụng Quyết Định & Chạy Lại"**, Agent nạp `human_overrides` và tính toán lại toàn bộ KPI + Top 5 trong `< 1 giây`.

---

## 📑 SLIDE 12: SKILL 3 — ĐÁNH GIÁ KPI 10 SALES & THUẬT TOÁN CHẤM ĐIỂM TOP 5
### **Công Thức Xếp Hạng Đa Tiêu Chí & Kiểm Soát Ranh Giới Quy Tắc**
- **1. Điều kiện tiên quyết (Hard Gatekeeper):**
  - Chỉ chấm điểm các cơ hội: **Đang mở** + **Đủ 3 thông tin xác nhận** (Người quyết định, Ngân sách, Bước tiếp theo) + **Không quá hạn ngày chốt** + **Có tương tác trong 14 ngày qua**.
- **2. Công thức chấm điểm ưu tiên (Thang điểm 100):**
  - **Giá trị dự kiến trọng số (35đ):** Tỷ lệ thuận với `weighted_forecast_value`.
  - **Độ chín giai đoạn (25đ):** `Closing` (25đ), `Negotiation` (20đ), `Proposal` (14đ)...
  - **Độ khẩn cấp ngày chốt (20đ):** Chốt trong 0–7 ngày tới đạt tối đa 20đ.
  - **Hỗ trợ thu hẹp thiếu hụt KPI của NVKD (20đ):** Ưu tiên hỗ trợ nhân viên đang thiếu KPI nhiều.
- **3. Xử lý ranh giới `< 5` cơ hội (`RULE-PRIORITY-TOP5`):**
  - Nếu cả công ty chỉ có `< 5` cơ hội đủ chuẩn, Agent chỉ xuất đúng số lượng thực tế (ví dụ 3/5) và giải thích rõ danh sách các cơ hội lớn bị loại trừ do thiếu xác nhận.

---

## 📑 SLIDE 13: TÍCH HỢP GOOGLE GEMINI AI & CƠ CHẾ TỰ PHỤC HỒI (FALLBACK)
### **Tổng Hợp Nhận Xét Điều Hành Thông Minh (`_generate_executive_synthesis`)**
- **1. Thiết kế Context tinh gọn (Lightweight Context Prompting):**
  - Chỉ truyền vào Gemini bản tóm tắt JSON đã tính chuẩn (`recorded_revenue`, `weighted_forecast_estimate`, `anomalies_count`, `top_priority_count`, `boundary_note`).
  - Tiết kiệm 95% lượng Token, thời gian phản hồi nhanh và không bao giờ bịa số liệu.
- **2. Cơ chế Multi-Model Rotation & Offline Fallback:**
  - Tự động xoay vòng qua danh sách model (`gemini-3.5-flash`, `gemini-3-flash-preview`, `gemini-flash-latest`).
  - Nếu mạng bị ngắt hoặc hết hạn mức API, hệ thống tự động kích hoạt **Built-in Deterministic Synthesis** để xuất báo cáo điều hành 3 phần chuẩn xác mà không bao giờ báo lỗi sập trang.

---

## 📑 SLIDE 14: THIẾT KẾ GIAO DIỆN TỐI GIẢN & TRẢI NGHIỆM NGƯỜI DÙNG (UI/UX)
### **Kết Hợp Phong Cách Google Chrome, Smallpdf & AY Studio Dashboard**
- **Luồng trải nghiệm 3 màn hình mượt mà (`static/index.html`):**
  1. **Màn hình 1 — Trang Bìa Duydudi:** Tối giản, nổi bật thương hiệu **Công ty của Duydudi**, huy hiệu đánh giá **4.9/5.0 ⭐** và nút **"BẮT ĐẦU ->"**.
  2. **Màn hình 2 — Trang Tải Tệp Chuẩn Smallpdf:** Khung kéo thả màu xanh dương `#0070e0` viền đứt nét, hoàn toàn trống và sạch sẽ, hỗ trợ chọn file `.xlsx/.csv/.zip`, dán link Google Sheets và đánh giá 5 sao tương tác.
  3. **Màn hình 3 — Bàn Làm Việc Studio Dashboard:** 4 thẻ KPI màu sắc trực quan, biểu đồ tròn 3D, biểu đồ sóng Earning Summary, nhận xét Gemini Live và các hộp thoại kiểm duyệt chi tiết.
- **Tiện ích giao diện:** Chuyển đổi chế độ **Ngày / Đêm (Light / Dark Mode)** và thanh Tab điều hướng kiểu trình duyệt Google Chrome.

---

## 📑 SLIDE 15: KIỂM THỬ TỰ ĐỘNG, TRIỂN KHAI CLOUD & KẾT LUẬN
### **Đảm Bảo Chất Lượng Thực Chiến (`test_runner.py` & Serverless Ready)**
- **1. Bộ Kiểm thử Tự động 3 Kịch bản (`GET /api/self-test`):**
  - **Kịch bản 1 (Normal Case):** Kiểm tra gộp trùng, tách đơn hàng đã chốt và trích xuất đủ 5/5 cơ hội ưu tiên -> **PASS**.
  - **Kịch bản 2 (Missing & Conflict):** Kiểm tra phát hiện cơ hội thiếu ngân sách, mâu thuẫn trạng thái và quá hạn -> **PASS**.
  - **Kịch bản 3 (Boundary Case):** Kiểm tra xử lý ngày chốt sổ và trường hợp chỉ có `< 5` cơ hội hợp lệ -> **PASS**.
- **2. Sẵn sàng cho môi trường Cloud / Vercel Serverless:**
  - Tự động nhận diện biến môi trường `VERCEL` để chuyển hướng ghi log sang `/tmp/logs`, khắc phục triệt để lỗi `Read-only file system`.
- **3. Tổng kết giá trị:**
  - Rút ngắn thời gian tổng hợp báo cáo tuần của Trưởng phòng KD từ **4 giờ xuống còn 5 giây**, loại bỏ 100% doanh số ảo và chuẩn hóa kỷ luật dữ liệu B2B cho **Công ty của Duydudi**.
