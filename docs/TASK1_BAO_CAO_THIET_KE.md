# BÁO CÁO ĐỀ DỰ ÁN 09 (AI301): PHÂN TÍCH TIẾN ĐỘ CƠ HỘI VÀ DỰ KIẾN DOANH SỐ — CÔNG TY BÌNH MINH

---

## PHẦN 1 (TASK 1): PHÂN TÍCH QUY TRÌNH, NHẬN DIỆN VẤN ĐỀ VÀ THIẾT KẾ HỆ THỐNG AI AGENT

### 1. Phân tích quy trình hiện tại (As-Is Process)
Tại **Công ty Bình Minh** (doanh nghiệp B2B dịch vụ công nghệ với 10 Nhân viên Kinh doanh - NVKD và ~300 cơ hội đang theo dõi), quy trình tổng hợp báo cáo hằng tuần hiện đang diễn ra thủ công qua các bước:
1. **Tiếp nhận dữ liệu phân tán**: 10 NVKD tự cập nhật cơ hội bán hàng vào các file Excel/Google Sheets cá nhân ở các thời điểm khác nhau trong tuần; bộ phận Kế toán/Admin xuất danh sách đơn hàng đã ký (`Orders`); lịch sử liên hệ khách hàng (`Activities`) nằm rời rạc ở nhật ký cuộc gọi/email.
2. **Tổng hợp thủ công (Sales Ops / Trưởng phòng)**: Vào cuối tuần, Trưởng phòng Kinh doanh gom các tệp từ NVKD, copy-paste vào một bảng tổng hợp chung mà không có bước khử trùng lặp chuẩn hóa hoặc khóa mốc thời gian đối chiếu (`Cut-off Date`).
3. **Tính toán cảm tính**: Cộng gộp cả các cơ hội đã thành đơn hàng vào danh sách cơ hội đang mở; sử dụng cột "Tỷ lệ thành công dự kiến (%)" do từng NVKD tự điền cảm tính để nhân với giá trị hợp đồng.
4. **Chọn cơ hội ưu tiên và giao việc**: Trưởng phòng lọc các cơ hội có giá trị tiền lớn nhất (Big Deals) để ưu tiên hỗ trợ trong tuần tới mà không kiểm tra xem cơ hội đó đã có đủ thông tin xác nhận (ngân sách, người quyết định, hoạt động tương tác gần nhất) hay chưa.

---

### 2. Xác định các vấn đề trong bối cảnh và ảnh hưởng cụ thể (Ít nhất 3 vấn đề cốt lõi)

| STT | Vấn đề trong quy trình hiện tại | Nguyên nhân gốc rễ | Ảnh hưởng cụ thể đến doanh nghiệp |
| :--- | :--- | :--- | :--- |
| **1** | **Trùng lặp dữ liệu và lệch mốc thời gian đối chiếu (Data Duplication & Inconsistent Cut-off)** | NVKD cập nhật ở nhiều tệp và nhiều thời điểm khác nhau; cùng 1 cơ hội khách hàng xuất hiện ở nhiều bản ghi hoặc phát sinh dữ liệu sau ngày chốt sổ. | **Ảo tưởng quy mô Pipeline (Pipeline Inflation)**: Tổng giá trị cơ hội mở bị thổi phồng; số liệu giữa các kỳ báo cáo không đồng nhất, khiến Ban Giám đốc đánh giá sai năng lực tạo nguồn cơ hội. |
| **2** | **Sai lệch trạng thái thực tế ("Zombie Opportunities" - Cơ hội chết/Quá hạn chưa đóng)** | Cơ hội đã mất khách (khách từ chối trong lịch sử liên hệ) nhưng NVKD quên hoặc ngại đổi sang `Closed Lost`; cơ hội đã qua ngày dự kiến chốt (`Expected Close Date`) nhiều tuần nhưng không có hoạt động tương tác mới. | **Lãng phí nguồn lực hỗ trợ & Dự báo ảo**: Trưởng phòng phân bổ thời gian hỗ trợ vào các thương vụ đã thất bại hoặc bị "đóng băng", bỏ lỡ các cơ hội thực sự đang cần thúc đẩy. |
| **3** | **Tính trùng doanh số đã bán vào cơ hội mở & Dùng tỷ lệ dự kiến tùy ý (Double Counting & Arbitrary Weighting)** | Không đối chiếu chéo danh sách Cơ hội (`Opportunities`) với Đơn hàng đã ghi nhận (`Orders`); cho phép NVKD tự gán xác suất chốt (ví dụ 80%-90% dù mới ở bước Gửi báo giá). | **Sai lệch nghiêm trọng dự báo doanh số & KPI**: Doanh số đã ghi nhận bị cộng 2 lần (vừa ở Doanh số thực tế, vừa ở Pipeline mở); doanh số dự kiến bị lạc quan quá mức, dẫn đến hụt thu thực tế cuối kỳ. |
| **4** | **Ưu tiên sai mục tiêu — Chạy theo giá trị lớn nhưng thiếu thông tin xác nhận (Blind Big-Deal Prioritization)** | Chỉ sắp xếp ưu tiên theo `Deal Value` mà bỏ qua tiêu chí xác nhận (Người quyết định, Ngân sách, Bước tiếp theo, Tương tác gần nhất). | **Rủi ro thất bại phút chót**: Tập trung nguồn lực cho các "cơ hội bánh vẽ" (giá trị lớn nhưng chưa xác minh nhu cầu/ngân sách), làm giảm tỷ lệ chuyển đổi (Win-rate) toàn phòng. |

---

### 3. Quy trình đề xuất (To-Be Process) — 5 Bước chuẩn hóa theo Sơ đồ Kiến trúc

Quy trình mới kết hợp giữa **Bộ phận Sales & Kế toán (Sales Ops)**, **AI Agent (Tự động)** và **Quản lý / Con người (Human-in-the-loop)** qua 5 bước chặt chẽ:

| Bước | Tên bước | Đầu vào (Input) | Công việc thực hiện (Activities) | Đầu ra (Output) | Chịu trách nhiệm |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Bước 1** | **Thu thập & Hợp nhất đầu vào** | - Hồ sơ cơ hội (`Opportunities`)<br>- Lịch sử trạng thái & liên hệ (`Activities`)<br>- Đơn hàng đã ghi nhận (`Orders`)<br>- Nhân viên & Mục tiêu (`KPI Targets`)<br>- Ngày đối chiếu (`Report Cut-off Date`) & Bộ quy tắc (`Rules`) | Chuẩn bị dữ liệu từ các nguồn (CSV, XLSX, hoặc Google Sheets Link), thiết lập ngày đối chiếu thống nhất và nạp vào hệ thống. | Bộ dữ liệu thô đã nạp + Cấu hình ngày đối chiếu & Quy tắc trọng số. | **Bộ phận Sales & Kế toán (Sales Ops)** |
| **Bước 2** | **Chuẩn hóa, Hợp nhất & Tính toán (Skill 1)** | Dữ liệu thô từ Bước 1 + Ngày đối chiếu + Bảng trọng số giai đoạn chuẩn. | - Lọc dữ liệu đúng mốc ngày đối chiếu (`<= report_date`).<br>- Tự động phát hiện và hợp nhất cơ hội trùng lặp (`opp_id` hoặc cùng `customer_name + product/deal`).<br>- Đối chiếu với `Orders` để tách bạch **Doanh số đã ghi nhận** khỏi **Cơ hội còn mở**.<br>- Xác định trạng thái tại ngày báo cáo và áp **trọng số giai đoạn chuẩn** (bỏ tỷ lệ tùy ý của Sales) để tính **Giá trị dự kiến (Ước tính)**. | - Danh sách cơ hội duy nhất đã chuẩn hóa.<br>- Bảng tách bạch Doanh số thực tế vs Cơ hội mở vs Doanh số dự kiến theo trọng số. | **AI Agent (Tự động)** |
| **Bước 3** | **Phát hiện rủi ro, Quá hạn & Mâu thuẫn (Skill 2)** | Danh sách cơ hội đã hợp nhất + Lịch sử `Activities` + `Orders`. | - Quét các cơ hội **quá hạn chốt** (`expected_close_date < report_date`) hoặc **đóng băng** (không có hoạt động liên hệ mới quá $N$ ngày).<br>- Phát hiện **thiếu thông tin xác nhận** (thiếu người quyết định, ngân sách, bước tiếp theo, giá trị $\le 0$).<br>- Phát hiện **mâu thuẫn dữ liệu** (VD: đã có đơn hàng nhưng trạng thái chưa `Closed Won`; lịch sử liên hệ ghi khách từ chối/hủy nhưng trạng thái vẫn mở; tỷ lệ Sales tự nhập lệch lớn so với quy tắc). | Danh sách cảnh báo chi tiết phân loại theo: `CONFLICT` (Mâu thuẫn), `OVERDUE` (Quá hạn/Không tương tác), `MISSING_INFO` (Thiếu thông tin). | **AI Agent (Tự động)** |
| **Bước 4** | **Kiểm duyệt & Xác minh (Human-in-the-loop #1)** | Danh sách các ca nghi vấn, mâu thuẫn, quá hạn và thiếu thông tin từ Bước 3. | Trưởng phòng KD / Sales Ops xem xét từng ca nghi vấn trên giao diện: quyết định xác nhận chuyển trạng thái thực tế (`Closed Won`, `Closed Lost`, giữ mở có điều kiện, hoặc yêu cầu Sales bổ sung thông tin). Hệ thống cập nhật ngược lại cho AI Agent tính toán lại. | Dữ liệu cơ hội đã được con người xác thực (Verified Clean Pipeline) + Nhật ký phê duyệt (Audit Log). | **Quản lý / Con người (Trưởng phòng KD & Sales Ops)** |
| **Bước 5** | **Phân tích KPI, Đề xuất 5 Cơ hội Ưu tiên & Phê duyệt giao việc (Skill 3 + HITL #2)** | Dữ liệu sạch sau Bước 4 + Bảng `KPI Targets` theo kỳ. | - **AI Agent**: Tính mức hoàn thành KPI từng NVKD (Doanh số đã ghi nhận, Cơ hội mở, Giá trị dự kiến có trọng số, Khoảng cách mục tiêu - Gap).<br>- **AI Agent**: Chấm điểm ưu tiên (Priority Score) cho các cơ hội mở *hợp lệ* và trích xuất **Top 5 cơ hội cần ưu tiên trong tuần tới** (hoặc nêu số thực tế & lý do nếu $< 5$), kèm căn cứ và hành động cụ thể.<br>- **Con người**: Trưởng phòng duyệt báo cáo cuối cùng và giao việc cho NVKD. | - Báo cáo tổng hợp theo NVKD & toàn phòng.<br>- Danh sách Top 5 cơ hội ưu tiên kèm Action Plan.<br>- Biên bản giao việc tuần tiếp theo. | **AI Agent (Tự động)** phối hợp **Trưởng phòng KD (Duyệt cuối)** |

---

### 4. Lý do lựa chọn các bước cho AI Agent tự thực hiện (Ít nhất 3 bước)

Hệ thống giao cho **AI Agent tự động thực hiện 3 bước cốt lõi (Bước 2, Bước 3, và phần Phân tích/Đề xuất của Bước 5)** vì các lý do sau:
1. **Bước 2 — Hợp nhất cơ hội trùng, đối chiếu đơn hàng và tính giá trị dự kiến theo quy tắc**:
   - *Lý do*: Với ~300 cơ hội từ 10 NVKD và hàng nghìn dòng lịch sử hoạt động/đơn hàng, con người đối chiếu thủ công rất dễ bỏ sót bản ghi trùng hoặc cộng nhầm đơn hàng đã chốt vào cơ hội mở. AI Agent kết hợp công cụ xử lý dữ liệu chính xác 100% theo quy tắc trọng số, loại bỏ hoàn toàn sự thiên vị/cảm tính của từng NVKD.
2. **Bước 3 — Quét rủi ro, quá hạn, thiếu thông tin và phát hiện mâu thuẫn ngữ nghĩa**:
   - *Lý do*: Mâu thuẫn thường ẩn trong ghi chú hoạt động tự nhiên (ví dụ: ghi chú cuộc gọi viết *"Khách báo tạm dừng dự án chuyển sang đối thủ X"* nhưng trạng thái cơ hội vẫn để *Đàm phán*). AI Agent có khả năng đọc hiểu ngữ cảnh ghi chú (`Activities`) kết hợp kiểm tra chéo ngày tháng (`expected_close_date`, `last_activity_date`) để phát hiện tức thì các "cơ hội chết" mà mắt thường dễ lướt qua.
3. **Bước 5 (Phần AI) — Tính toán KPI đa chiều và xếp hạng Top 5 cơ hội ưu tiên kèm khuyến nghị hành động**:
   - *Lý do*: Việc xếp hạng ưu tiên đòi hỏi kết hợp nhiều tiêu chí đồng thời (Giá trị dự kiến, Giai đoạn, Độ cấp thiết theo ngày chốt, Mức độ đầy đủ thông tin xác nhận, và Khoảng cách KPI của NVKD phụ trách). AI Agent tính điểm nhất quán, giải thích rõ căn cứ cho từng cơ hội, tự động kiểm tra điều kiện ranh giới (nếu $<5$ cơ hội hợp lệ thì báo cáo số thực tế kèm lý do) và sinh khuyến nghị hành động sát với tình trạng thực tế.

---

### 5. Điểm kiểm soát của Con người (Human-in-the-loop Checkpoints)

Hệ thống thiết kế **2 điểm kiểm tra và quyết định bắt buộc của con người**:

#### Điểm kiểm soát 1: Kiểm duyệt & Xác minh dữ liệu bất thường (Bước 4)
- **Nội dung con người quyết định**: Xác nhận trạng thái thực tế cho các cơ hội có mâu thuẫn (giữa `Opportunities`, `Orders`, `Activities`), quyết định đóng (`Closed Lost`) hay gia hạn cho các cơ hội quá hạn lâu ngày, và xử lý các bản ghi trùng có thông tin lệch nhau.
- **Điều kiện chuyển xử lý (Escalation Trigger)**:
  - Cơ hội đã xuất hiện trong bảng `Orders` (đã duyệt) nhưng trạng thái cơ hội khác `Closed Won`.
  - Ghi chú hoạt động mới nhất có tín hiệu từ chối/mất khách nhưng trạng thái vẫn đang mở.
  - Cơ hội có giá trị lớn ($\ge 200$ triệu VNĐ) nhưng thiếu thông tin người quyết định/ngân sách hoặc quá hạn chốt $> 14$ ngày không có tương tác.
  - Hai bản ghi trùng khách hàng + dự án nhưng lệch nhau về NVKD phụ trách hoặc lệch giá trị $> 20\%$.
- **Cách phát hiện kết quả sai của AI Agent**:
  - Giao diện hiển thị song song **"Trạng thái gốc"**, **"Bằng chứng truy vết (Mã đơn hàng / Nội dung liên hệ / Số ngày quá hạn)"** và **"Đề xuất của AI"**.
  - Nếu AI hiểu nhầm ghi chú (ví dụ khách nói *"Tạm hoãn 1 tuần để xin chữ ký Sếp"* nhưng AI gắn cờ mất khách), Trưởng phòng dễ dàng phát hiện qua cột bằng chứng gốc và bấm **"Giữ mở (Override AI)"** ngay trên giao diện.

#### Điểm kiểm soát 2: Phê duyệt Top 5 Cơ hội Ưu tiên & Phân công tuần tới (Cuối Bước 5)
- **Nội dung con người quyết định**: Duyệt báo cáo mức hoàn thành KPI, kiểm tra danh sách 5 cơ hội ưu tiên do AI đề xuất, điều chỉnh người hỗ trợ (Trưởng phòng/Presales/Giám đốc) và phê duyệt kế hoạch hành động tuần tới.
- **Điều kiện chuyển xử lý (Escalation Trigger)**:
  - Mọi danh sách ưu tiên tuần mới trước khi phát hành chính thức cho đội ngũ Sales.
  - Trường hợp số cơ hội hợp lệ đủ điều kiện ưu tiên $< 5$ (cần Trưởng phòng quyết định chỉ đạo bổ sung thông tin cho các cơ hội đang bị treo).
  - Nhân viên có tỷ lệ hoàn thành mục tiêu dự kiến (Expected Attainment) $< 60\%$ cần Trưởng phòng can thiệp hỗ trợ.
- **Cách phát hiện kết quả sai của AI Agent**:
  - Kiểm tra đối chiếu tự động (Sanity Check): Tổng Doanh số đã ghi nhận + Cơ hội mở phải khớp 100% với sổ chi tiết; không có cơ hội nào chưa xác nhận thông tin hoặc đã chốt/mất lọt vào danh sách Top 5 ưu tiên; mọi con số dự kiến đều đi kèm nhãn cảnh báo *"Ước tính theo quy tắc trọng số"*.

---

### 6. Bản thiết kế hệ thống AI Agent (Mục tiêu – Đầu vào – Quy trình – Đầu ra & Kiến trúc)

#### A. Mục tiêu (Objectives)
- Chuẩn hóa dữ liệu bán hàng theo **một ngày đối chiếu thống nhất (`Report Date`)**.
- Tách bạch tuyệt đối **Doanh số đã ghi nhận (Actual Revenue)** và **Giá trị cơ hội mở (Open Pipeline)**; tính **Giá trị dự kiến (Weighted Forecast)** bằng trọng số quy tắc.
- Tự động phát hiện 100% cơ hội trùng, thiếu thông tin, quá hạn hoặc mâu thuẫn dữ liệu để con người xác minh.
- Xuất báo cáo KPI theo từng NVKD và đề xuất tối đa **5 cơ hội ưu tiên hợp lệ** trong tuần tới kèm căn cứ, hành động cụ thể và nhật ký truy vết (Audit Trail) đầy đủ.

#### B. Đầu vào (Inputs)
1. **Dữ liệu nghiệp vụ (hỗ trợ tải file `.csv`, `.xlsx` hoặc dán link Google Sheets)**:
   - `opportunities`: Mã cơ hội, Tên khách hàng, Tên dự án/dịch vụ, Mã NVKD, Giai đoạn, Giá trị, Tỷ lệ Sales tự nhập, Ngày tạo, Ngày cập nhật, Ngày dự kiến chốt, Người quyết định, Xác nhận ngân sách, Bước tiếp theo.
   - `activities`: Mã hoạt động, Mã cơ hội, Ngày liên hệ, Loại hình (Call/Meeting/Email/Demo), Tóm tắt kết quả, Tín hiệu trạng thái.
   - `orders`: Mã đơn hàng, Mã cơ hội, Khách hàng, Mã NVKD, Giá trị đơn hàng, Ngày ghi nhận, Trạng thái đơn.
   - `kpi_targets`: Mã NVKD, Họ tên, Nhóm, Mục tiêu doanh số kỳ (Target), Kỳ đánh giá.
2. **Tham số điều khiển & Quy tắc (`rules`)**:
   - Ngày đối chiếu báo cáo (`report_date`).
   - Bảng trọng số giai đoạn chuẩn: `Lead/Khởi tạo (10%)`, `Qualification/Khảo sát (25%)`, `Proposal/Báo giá (50%)`, `Negotiation/Đàm phán (75%)`, `Closing/Chốt HĐ (90%)`, `Closed Won (100% - chuyển sang Doanh số thực tế)`, `Closed Lost (0%)`.
   - Ngưỡng cảnh báo không tương tác (`inactivity_days_threshold = 14` ngày) và Tiêu chí chấm điểm ưu tiên.

#### C. Quy trình & Kiến trúc Phối hợp Đa Kỹ năng (Multi-Skill Agent Architecture)
Hệ thống áp dụng kiến trúc **Orchestrated Tool-Calling AI Agent với Human-in-the-Loop và Audit Trail**, gồm 3 kỹ năng nghiệp vụ (Business Skills) phối hợp tuần tự:
- **Skill 1 (`DataConsolidationAndValuationSkill`)**: Đọc dữ liệu đa nguồn $\rightarrow$ Lọc theo `report_date` $\rightarrow$ Khử trùng lặp cơ hội $\rightarrow$ Đối chiếu `Orders` & `Activities` để xác định trạng thái chuẩn $\rightarrow$ Tách doanh số đã bán và tính giá trị dự kiến có trọng số.
- **Skill 2 (`AnomalyAndConflictDetectionSkill`)**: Phân tích chéo Hồ sơ cơ hội + Lịch sử hoạt động + Đơn hàng $\rightarrow$ Phát hiện cơ hội thiếu thông tin, quá hạn chốt, quá hạn tương tác, hoặc mâu thuẫn trạng thái $\rightarrow$ Đẩy sang trạm Kiểm duyệt Con người (Bước 4).
- **Skill 3 (`KPIAndPriorityActionSkill`)**: Sau khi nhận quyết định xác minh từ Bước 4 $\rightarrow$ Tổng hợp bảng kết quả theo từng NVKD $\rightarrow$ Xếp hạng các cơ hội mở hợp lệ đã xác nhận thông tin $\rightarrow$ Chọn Top 5 cơ hội ưu tiên (hoặc giải trình nếu $<5$) $\rightarrow$ Gọi LLM (Gemini / OpenAI) kết hợp Rule Engine để sinh phân tích chiến lược và hành động cụ thể cho từng cơ hội.
- **Audit Logger (`TraceabilityEngine`)**: Ghi nhận toàn bộ lịch sử Agent đọc file, gọi từng tool/skill, quy tắc áp dụng, các bản ghi bị gộp/đổi trạng thái, và quyết định can thiệp của con người theo chuẩn JSONL/Table truy vết.

#### D. Đầu ra (Outputs)
1. **Dashboard Tổng quan & Bảng kết quả theo Nhân viên**: Doanh số đã ghi nhận, Số cơ hội mở, Tổng giá trị cơ hội mở, Giá trị dự kiến có trọng số (kèm nhãn ước tính), Mục tiêu KPI, Tỷ lệ hoàn thành thực tế (%) và Tỷ lệ hoàn thành dự kiến (%).
2. **Danh sách Cơ hội bất thường cần xác minh**: Phân loại rõ Thiếu thông tin, Quá hạn, Mâu thuẫn dữ liệu kèm nút thao tác duyệt trực tiếp trên giao diện.
3. **Danh sách Top 5 Cơ hội Ưu tiên tuần tới**: Nêu rõ Điểm ưu tiên, Căn cứ lựa chọn, Rủi ro cần lưu ý, Việc cần làm tiếp theo và Người hỗ trợ. (Nếu $<5$ cơ hội hợp lệ, hiển thị rõ số lượng thực tế và lý do loại trừ các cơ hội khác).
4. **Nhật ký Truy vết (Execution & Audit Log)** và chức năng xuất báo cáo hoàn chỉnh.

