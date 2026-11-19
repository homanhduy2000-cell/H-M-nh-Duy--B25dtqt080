import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Bảng màu chủ đạo Thương hiệu Công ty của Duydudi
COLOR_BG_DARK = RGBColor(15, 23, 42)        # Slate 900 (#0f172a)
COLOR_HEADER_BG = RGBColor(20, 33, 77)      # Navy (#14214d)
COLOR_CARD_BG = RGBColor(30, 41, 59)        # Slate 800 (#1e293b)
COLOR_CARD_BORDER = RGBColor(51, 65, 85)    # Slate 700 (#334155)
COLOR_PINK = RGBColor(255, 77, 121)         # Duydudi Pink (#ff4d79)
COLOR_CYAN = RGBColor(0, 210, 255)          # Duydudi Cyan (#00d2ff)
COLOR_EMERALD = RGBColor(16, 185, 129)      # Emerald (#10b981)
COLOR_AMBER = RGBColor(245, 158, 11)        # Amber (#f59e0b)
COLOR_PURPLE = RGBColor(168, 85, 247)       # Purple (#a855f7)
COLOR_WHITE = RGBColor(255, 255, 255)
COLOR_TEXT_LIGHT = RGBColor(226, 232, 240)  # Slate 200
COLOR_TEXT_MUTED = RGBColor(148, 163, 184)  # Slate 400


SLIDES_CONTENT = [
    {
        "num": 1,
        "tag": "SLIDE 01 / 15 • TRANG BÌA DỰ ÁN",
        "title": "Xây Dựng Duydudi AI Agent:\nPhân Tích Tiến Độ Cơ Hội & Dự Báo Doanh Số B2B",
        "subtitle": "Từ dữ liệu B2B phân mảnh đến quyết định điều hành chính xác với kiến trúc Deterministic Python + Google Gemini AI",
        "accent": COLOR_PINK,
        "cards": [
            {
                "title": "🏢 ĐƠN VỊ & THƯƠNG HIỆU",
                "color": COLOR_PINK,
                "bullets": [
                    "Công ty của Duydudi (Duydudi • AI Sales Ops).",
                    "Giải pháp chuyên biệt cho Trưởng phòng Kinh doanh B2B.",
                    "Đánh giá trải nghiệm: 4.9 / 5.0 Sao trên giao diện trực tuyến."
                ]
            },
            {
                "title": "⚙️ CÔNG NGHỆ CỐT LÕI",
                "color": COLOR_CYAN,
                "bullets": [
                    "Backend: Python FastAPI + Pandas Deterministic Engine.",
                    "AI Model: Google Gemini 3.5 Flash + Built-in Reasoning Fallback.",
                    "UI/UX: Phong cách Google Chrome + Smallpdf Dropzone + Studio Dashboard."
                ]
            },
            {
                "title": "🎯 GIÁ TRỊ ĐẠT ĐƯỢC",
                "color": COLOR_EMERALD,
                "bullets": [
                    "Tự động đọc Excel (.xlsx), CSV (.csv), ZIP & Google Sheets.",
                    "Phát hiện file lỗi & AI tự động sửa lỗi (AI Auto-Repair Skill).",
                    "Tách bạch 100% doanh số thực tế và dự báo trọng số + Top 5 ưu tiên."
                ]
            }
        ],
        "notes": "Xin chào mọi người. Hôm nay tôi xin trình bày quy trình thiết kế và xây dựng từ A-Z hệ thống Duydudi AI Agent của Công ty của Duydudi — một trợ lý AI giúp Trưởng phòng Kinh doanh B2B tự động thu thập, sửa lỗi dữ liệu và dự báo doanh số."
    },
    {
        "num": 2,
        "tag": "SLIDE 02 / 15 • BÀI TOÁN THỰC TẾ & NỖI ĐAU NGHIỆP VỤ",
        "title": "Tại Sao Báo Cáo Doanh Số B2B Truyền Thống Thường Sai Lệch?",
        "subtitle": "4 nguyên nhân gốc rễ khiến báo cáo thủ công mất hàng giờ mỗi tuần nhưng số liệu dự báo vẫn bị ảo",
        "accent": COLOR_AMBER,
        "cards": [
            {
                "title": "1. DỮ LIỆU PHÂN MẢNH Ở 4 BẢNG RỜI RẠC",
                "color": COLOR_PINK,
                "bullets": [
                    "Opportunities (Cơ hội), Activities (Lịch sử liên hệ), Orders (Đơn hàng) và KPI Targets nằm ở nhiều file Excel/CSV/Google Sheets.",
                    "Tên cột không đồng nhất (tiếng Việt, viết tắt, sai định dạng tiền tệ)."
                ]
            },
            {
                "title": "2. TÍNH TRÙNG DOANH SỐ & TRÙNG DỰ ÁN",
                "color": COLOR_AMBER,
                "bullets": [
                    "Khách đã ký đơn hàng (Orders) nhưng cơ hội vẫn để mở -> bị cộng trùng vào cả doanh số thực tế và dự báo.",
                    "Cùng 1 dự án nhưng 2 nhân viên báo cáo ở 2 file khác nhau."
                ]
            },
            {
                "title": "3. CƠ HỘI XÁC SỐNG (ZOMBIE DEALS) & CẢM TÍNH",
                "color": COLOR_CYAN,
                "bullets": [
                    "Nhật ký cuộc gọi ghi khách đã từ chối/chọn đối thủ, hoặc quá hạn ngày chốt > 14 ngày không tương tác nhưng vẫn để Mở.",
                    "Sales tự điền xác suất thắng 80%–90% theo cảm tính làm phồng ảo dự báo."
                ]
            }
        ],
        "notes": "Trước khi viết code, chúng tôi phân tích 4 nỗi đau lớn của dữ liệu B2B. Nếu chỉ đưa file Excel thô cho ChatGPT cộng sổ, LLM sẽ cộng gộp cả dữ liệu rác và tính trùng đơn hàng."
    },
    {
        "num": 3,
        "tag": "SLIDE 03 / 15 • TRIẾT LÝ THIẾT KẾ",
        "title": "Kiến Trúc Chống Ảo Giác Số Liệu (Deterministic + LLM)",
        "subtitle": "Kết hợp sức mạnh tính toán chính xác tuyệt đối của Python và tư duy ngôn ngữ chiến lược của Google Gemini",
        "accent": COLOR_EMERALD,
        "cards": [
            {
                "title": "🛡️ TẦNG 1: DETERMINISTIC PYTHON GATEKEEPER",
                "color": COLOR_EMERALD,
                "bullets": [
                    "Tuyệt đối không dùng LLM để cộng trừ bảng dữ liệu thô (tránh tràn token và sai số học).",
                    "Dùng Pandas xử lý 100% việc lọc ngày chốt sổ, gộp bản ghi trùng và đối chiếu chéo 4 bảng.",
                    "Áp dụng công thức trọng số giai đoạn và chấm điểm ưu tiên hoàn toàn tất định."
                ]
            },
            {
                "title": "🧠 TẦNG 2: GOOGLE GEMINI 3.5 FLASH PLANNER",
                "color": COLOR_PURPLE,
                "bullets": [
                    "Chỉ nhận bản tóm tắt JSON Metadata đã được Python tính toán chuẩn xác.",
                    "Tổng hợp bản Nhận xét Điều hành (Executive Brief) gồm 3 phần chiến lược cho lãnh đạo.",
                    "Gắn nhãn minh bạch: 'Giá trị dự kiến chỉ là ước tính theo quy tắc trọng số'."
                ]
            }
        ],
        "notes": "Nguyên tắc vàng khi xây dựng Agent này là tách biệt hoàn toàn: Phần tính toán số học do Python đảm nhận 100%, phần nhận định chiến lược và diễn giải ngôn ngữ do Gemini AI thực hiện."
    },
    {
        "num": 4,
        "tag": "SLIDE 04 / 15 • KIẾN TRÚC HỆ THỐNG 4 TẦNG",
        "title": "Kiến Trúc Tổng Thể 4 Tầng End-to-End Của Hệ Thống",
        "subtitle": "Thiết kế phân lớp độc lập trong app.py, agent_engine.py và static/index.html",
        "accent": COLOR_CYAN,
        "cards": [
            {
                "title": "TẦNG 1: INGESTION",
                "color": COLOR_CYAN,
                "bullets": [
                    "Nhận file .xlsx, .csv, .zip từ khung Smallpdf.",
                    "Đọc trực tiếp URL Google Sheets đa tab.",
                    "Cổng Webhook API cho hệ thống CRM/POS ngoài."
                ]
            },
            {
                "title": "TẦNG 2: VALIDATION & REPAIR",
                "color": COLOR_AMBER,
                "bullets": [
                    "Tự động ánh xạ tên cột & phát hiện file lỗi.",
                    "Hỏi người dùng để AI Tự động sửa lỗi (AI Auto-Repair).",
                    "Cách ly bản ghi lỗi ra data_quarantine.csv."
                ]
            },
            {
                "title": "TẦNG 3: AGENT SKILLS & HITL",
                "color": COLOR_PURPLE,
                "bullets": [
                    "Skill 1: Lọc ngày chốt, gộp trùng & tính trọng số.",
                    "Skill 2 + Bước 4: Quét 8 rủi ro & Quản lý phê duyệt.",
                    "Skill 3 + Gemini: KPI 10 Sales & Top 5 ưu tiên."
                ]
            },
            {
                "title": "TẦNG 4: DELIVERY UI",
                "color": COLOR_EMERALD,
                "bullets": [
                    "Bàn làm việc Studio Dashboard trực quan.",
                    "Xuất file sạch cleaned_data.csv & báo cáo lỗi.",
                    "Ghi nhật ký truy vết agent_audit_log.jsonl."
                ]
            }
        ],
        "notes": "Hệ thống vận hành qua 4 tầng độc lập: Tiếp nhận đa nguồn -> Thẩm định & AI Sửa lỗi -> Thực thi 3 Kỹ năng kết hợp Con người kiểm duyệt -> Trình bày Dashboard và xuất tệp."
    },
    {
        "num": 5,
        "tag": "SLIDE 05 / 15 • QUY TRÌNH PHỐI HỢP 5 BƯỚC",
        "title": "Quy Trình 5 Bước Phối Hợp Giữa Con Người & AI Agent",
        "subtitle": "Kết hợp tự động hóa tốc độ cao của AI với quyền quyết định kiểm soát của Trưởng phòng Kinh doanh",
        "accent": COLOR_PINK,
        "cards": [
            {
                "title": "BƯỚC 1 & BƯỚC 2: THU THẬP & CHUẨN HÓA",
                "color": COLOR_CYAN,
                "bullets": [
                    "Bước 1 (Con người / API): Tải tệp Excel/CSV/ZIP hoặc dán link Google Sheets + Chọn Ngày đối chiếu (Report Date).",
                    "Bước 2 (AI Agent - Skill 1): Lọc dữ liệu theo ngày chốt sổ, gộp cơ hội trùng lặp và chuẩn hóa trọng số giai đoạn."
                ]
            },
            {
                "title": "BƯỚC 3 & BƯỚC 4: QUÉT RỦI RO & KIỂM DUYỆT (HITL)",
                "color": COLOR_PINK,
                "bullets": [
                    "Bước 3 (AI Agent - Skill 2): Đối chiếu chéo 3 bảng để phát hiện 8 nhóm mâu thuẫn, quá hạn và thiếu thông tin.",
                    "Bước 4 (Con người - HITL): Trưởng phòng KD phê duyệt các ca nghi vấn (Chuyển Won / Chuyển Lost / Xác minh giữ mở)."
                ]
            },
            {
                "title": "BƯỚC 5: TÍNH KPI, XẾP HẠNG TOP 5 & GEMINI BRIEF",
                "color": COLOR_EMERALD,
                "bullets": [
                    "Bước 5 (AI Agent - Skill 3 & Gemini): Đánh giá mức hoàn thành KPI của 10 NVKD, chấm điểm Top 5 cơ hội ưu tiên tuần tới và sinh Báo cáo điều hành."
                ]
            }
        ],
        "notes": "Quy trình 5 bước giúp AI xử lý 95% khối lượng tính toán phức tạp ở Bước 2, 3, 5; trong khi con người giữ vai trò quyết định ở Bước 1 và Bước 4."
    },
    {
        "num": 6,
        "tag": "SLIDE 06 / 15 • CỔNG TIẾP NHẬN ĐA NGUỒN",
        "title": "Tầng 1: Xây Dựng Cổng Đọc File Đa Định Dạng & Google Sheets",
        "subtitle": "Giải quyết triệt để bài toán người dùng tải lên nhiều định dạng file và đặt tên Sheet/Cột khác nhau",
        "accent": COLOR_CYAN,
        "cards": [
            {
                "title": "📂 BỘ ĐỌC EXCEL, CSV & ZIP ĐA BẢNG",
                "color": COLOR_CYAN,
                "bullets": [
                    "Hàm parse_excel_bytes() tự động quét tất cả các Sheet trong workbook (.xlsx) hoặc giải nén gói .zip.",
                    "Nhận diện 2 lớp: Quét theo tên Sheet (opp, act, ord, kpi) và Fallback quét theo tên cột đặc trưng.",
                    "Hàm smart_normalize_columns() tự động ánh xạ tên cột tiếng Việt ('Mã CH', 'Giá trị', 'Giai đoạn') về chuẩn Schema."
                ]
            },
            {
                "title": "🔗 KẾT NỐI GOOGLE SHEETS & WEBHOOK API",
                "color": COLOR_PINK,
                "bullets": [
                    "Hàm load_from_google_sheets_url() dùng Regex tách sheet_id từ link chia sẻ và chuyển sang luồng /export?format=xlsx.",
                    "Hỗ trợ đọc cả link xuất bản Web (/d/e/...) và tự động fallback sang CSV nếu là bảng đơn.",
                    "Endpoint POST /api/webhook/ingest cho phép hệ thống POS/CRM/n8n đẩy file tự động qua HTTP."
                ]
            }
        ],
        "notes": "Ở tầng tiếp nhận, chúng tôi thiết kế bộ đọc thông minh có khả năng nhận diện cả cột tiếng Việt lẫn tiếng Anh, đồng thời đọc trực tiếp link Google Sheets."
    },
    {
        "num": 7,
        "tag": "SLIDE 07 / 15 • KỸ NĂNG AI TỰ ĐỘNG SỬA LỖI FILE",
        "title": "Skill Đặc Biệt: Phát Hiện File Lỗi & AI Tự Động Chỉnh Sửa",
        "subtitle": "Khi gửi file lỗi, hệ thống cảnh báo chi tiết và hỏi người dùng có muốn AI tự động sửa lỗi để ra kết quả",
        "accent": COLOR_EMERALD,
        "cards": [
            {
                "title": "🚨 1. CHẨN ĐOÁN & THÔNG BÁO FILE LỖI TỨC THÌ",
                "color": COLOR_AMBER,
                "bullets": [
                    "Hàm build_file_diagnostics() tự động phát hiện: lỗi tên cột, số tiền chứa chữ ('1.25 tỷ', '680,000,000 VNĐ', số âm), thiếu Stage, thiếu ngày chốt, thiếu ngân sách.",
                    "Bật hộp thoại cảnh báo thống kê số lượng lỗi và hỏi: 'Bạn có muốn AI Agent tự động chỉnh sửa không?'."
                ]
            },
            {
                "title": "✨ 2. ĐỘNG CƠ AI SỬA LỖI (ai_auto_repair_dataset)",
                "color": COLOR_EMERALD,
                "bullets": [
                    "Khi người dùng chọn 'Có': AI chuẩn hóa tên cột, bóc tách số tiền sạch (parse_currency_smart) và gộp bản ghi trùng.",
                    "Tự động điền khuyết Stage, gia hạn ngày dự kiến chốt, bổ sung Người quyết định & Ngân sách hợp lệ.",
                    "Xuất bảng đối chiếu 'Trước khi sửa -> Sau khi AI sửa' và hiển thị ngay kết quả Dashboard chuẩn."
                ]
            }
        ],
        "notes": "Đây là kỹ năng rất thực tế: khi người dùng tải lên một file bị lỗi định dạng hoặc khuyết dữ liệu, hệ thống sẽ báo lỗi rõ ràng và hỏi có muốn AI sửa không. Chỉ cần bấm 'Có', AI sửa sạch 100% lỗi và ra kết quả ngay."
    },
    {
        "num": 8,
        "tag": "SLIDE 08 / 15 • XÂY DỰNG SKILL 1 (PHẦN 1)",
        "title": "Skill 1: Lọc Ngày Đối Chiếu & Hợp Nhất Cơ Hội Trùng Lặp",
        "subtitle": "Đảm bảo tính nhất quán theo mốc thời gian chốt sổ (Cut-off Date) và khử trùng lặp đa tệp",
        "accent": COLOR_CYAN,
        "cards": [
            {
                "title": "📅 BỘ LỌC NGÀY ĐỐI CHIẾU (RULE-CUTOFF-FILTER)",
                "color": COLOR_AMBER,
                "bullets": [
                    "Thiết lập mốc thời gian thống nhất: report_date (mặc định 2026-11-12).",
                    "Mọi đơn hàng (order_date > report_date) hoặc hoạt động (activity_date > report_date) phát sinh sau ngày chốt sổ đều được tách riêng vào cutoff_filtered.",
                    "Chống sai lệch số liệu khi tổng hợp báo cáo lùi ngày."
                ]
            },
            {
                "title": "🔄 THUẬT TOÁN GỘP TRÙNG NGỮ NGHĨA (RULE-DEDUP)",
                "color": COLOR_CYAN,
                "bullets": [
                    "Tạo khóa kép: kết hợp mã opp_id chuẩn và khóa ngữ nghĩa lower(customer_name) || lower(project_name).",
                    "Khi có nhiều bản ghi trùng: tự động giữ bản có updated_date mới nhất (<= report_date).",
                    "Tự động gắn cờ cảnh báo nếu các file lệch mã NVKD phụ trách hoặc chênh lệch giá trị >= 20%."
                ]
            }
        ],
        "notes": "Skill 1 giải quyết triệt để tình trạng 2 nhân viên cùng nhập 1 dự án ở 2 file khác nhau bằng thuật toán khóa ngữ nghĩa kết hợp Tên khách hàng và Tên dự án."
    },
    {
        "num": 9,
        "tag": "SLIDE 09 / 15 • XÂY DỰNG SKILL 1 (PHẦN 2)",
        "title": "Skill 1: Tách Bạch Doanh Số Thực Tế & Dự Báo Có Trọng Số",
        "subtitle": "Chấm dứt tình trạng cộng gộp đơn hàng đã chốt vào cơ hội mở và chuẩn hóa trọng số 5 giai đoạn",
        "accent": COLOR_PINK,
        "cards": [
            {
                "title": "💰 TÁCH BẠCH TUYỆT ĐỐI 3 ĐẠI LƯỢNG DOANH SỐ",
                "color": COLOR_EMERALD,
                "bullets": [
                    "1. Doanh số đã ghi nhận (actual_recorded_revenue): Chỉ tính từ bảng Orders hợp lệ / Closed Won. Khi đã Won -> Giá trị mở = 0.",
                    "2. Giá trị cơ hội mở (open_pipeline_value): Tổng giá trị gốc của các cơ hội chưa đóng.",
                    "3. Doanh số dự kiến ước tính (weighted_forecast_value) = Giá trị cơ hội mở x Trọng số chuẩn."
                ]
            },
            {
                "title": "📊 BẢNG TRỌNG SỐ QUY TẮC CHUẨN (STAGE_RULES)",
                "color": COLOR_PINK,
                "bullets": [
                    "• Lead (Tiếp cận mới): 10%   |   • Qualification (Xác định nhu cầu): 25%",
                    "• Proposal (Đã gửi đề xuất): 50%   |   • Negotiation (Đàm phán HĐ): 75%",
                    "• Closing (Chốt ký kết): 90%   |   • Closed Lost (Thất bại): 0%",
                    "Thay thế hoàn toàn xác suất cảm tính tự điền của nhân viên kinh doanh."
                ]
            }
        ],
        "notes": "Nhờ tách bạch 3 đại lượng và áp dụng bảng trọng số chuẩn từ 10% đến 90%, hệ thống loại bỏ hoàn toàn phần doanh số ảo do nhân viên tự nhập xác suất cảm tính."
    },
    {
        "num": 10,
        "tag": "SLIDE 10 / 15 • XÂY DỰNG SKILL 2",
        "title": "Skill 2: Động Cơ Quét 8 Nhóm Rủi Ro & Mâu Thuẫn Dữ Liệu",
        "subtitle": "Đối chiếu chéo 3 bảng (Opportunities ↔ Orders ↔ Activities) để tự động bắt lỗi nghiệp vụ",
        "accent": COLOR_AMBER,
        "cards": [
            {
                "title": "🔥 4 NHÓM LỖI MÂU THUẪN & QUÁ HẠN NGHIÊM TRỌNG",
                "color": COLOR_PINK,
                "bullets": [
                    "1. CONFLICT_ORDER_STATUS: Đã có đơn hàng bên bảng Orders nhưng cơ hội vẫn để mở -> Tự động chuyển Closed Won.",
                    "2. CONFLICT_ZOMBIE_LOST: Nhật ký cuộc gọi ghi 'mất khách/từ chối/đối thủ' nhưng vẫn để mở -> Chuyển Closed Lost.",
                    "3. OVERDUE_CLOSE_DATE: Ngày dự kiến chốt < Ngày báo cáo mà vẫn chưa đóng.",
                    "4. INACTIVE_STALE: Đóng băng tương tác > 14 ngày không có hoạt động liên hệ mới."
                ]
            },
            {
                "title": "⚠️ 4 NHÓM LỖI KHUYẾT THÔNG TIN & LỆCH QUY TẮC",
                "color": COLOR_AMBER,
                "bullets": [
                    "5. MISSING_VERIFICATION_INFO: Cơ hội mở (đặc biệt >= 500 triệu) thiếu Người quyết định, Ngân sách hoặc Bước tiếp theo.",
                    "6. CONFLICT_DUPLICATE_MERGE: Bản ghi gộp từ nhiều file bị lệch NVKD hoặc lệch giá trị >= 20%.",
                    "7. MISSING_STAGE / CLOSE_DATE: Bỏ trống giai đoạn hoặc ngày dự kiến chốt.",
                    "8. ARBITRARY_PROBABILITY: Xác suất Sales tự nhập lệch >= 20% so với quy tắc chuẩn."
                ]
            }
        ],
        "notes": "Skill 2 hoạt động như một kiểm toán viên tự động, đọc từng ghi chú trong bảng Activities và đối chiếu từng mã đơn hàng trong bảng Orders để phát hiện 8 nhóm bất thường."
    },
    {
        "num": 11,
        "tag": "SLIDE 11 / 15 • VÙNG CÁCH LY & NHẬT KÝ TRUY VẾT",
        "title": "Quarantine Zone & Nhật Ký Truy Vết (Explainable AI)",
        "subtitle": "Đảm bảo tính minh bạch, khả năng giải trình và truy vết từng bước suy luận của AI Agent",
        "accent": COLOR_PURPLE,
        "cards": [
            {
                "title": "📦 VÙNG CÁCH LY DỮ LIỆU LỖI (QUARANTINE ZONE)",
                "color": COLOR_PINK,
                "bullets": [
                    "Endpoint xuất file trực tiếp: GET /api/export/quarantine.csv.",
                    "Cô lập toàn bộ các dòng mâu thuẫn trạng thái, quá hạn, thiếu thông tin xác nhận hoặc nằm ngoài ngày chốt sổ.",
                    "Kèm theo cột lý do cách ly chi tiết và hướng xử lý đề xuất của AI Agent."
                ]
            },
            {
                "title": "📋 NHẬT KÝ TRUY VẾT (AUDIT TRAIL LOGGER)",
                "color": COLOR_EMERALD,
                "bullets": [
                    "Class AuditLogger ghi lại mọi bước thực thi vào tệp agent_audit_log.jsonl.",
                    "Cấu trúc chuẩn: run_id, timestamp, step, skill_or_tool, action, rule_applied, details, status.",
                    "Chứng minh AI Agent đọc dữ liệu thực tế và áp dụng đúng quy tắc nghiệp vụ, không phải 'hộp đen'."
                ]
            }
        ],
        "notes": "Mọi quyết định của Agent đều được ghi vào nhật ký agent_audit_log.jsonl và cho phép xuất riêng file CSV các bản ghi bị cách ly."
    },
    {
        "num": 12,
        "tag": "SLIDE 12 / 15 • BƯỚC 4: CON NGƯỜI KIỂM DUYỆT (HITL)",
        "title": "Thiết Kế Cơ Chế Human-in-the-Loop Override Tại Bước 4",
        "subtitle": "Trao quyền phán quyết cuối cùng cho Trưởng phòng Kinh doanh đối với các cơ hội nghi vấn",
        "accent": COLOR_CYAN,
        "cards": [
            {
                "title": "👤 TẠI SAO CẦN ĐIỂM DỪNG CHO CON NGƯỜI?",
                "color": COLOR_CYAN,
                "bullets": [
                    "Trong kinh doanh B2B, một cơ hội quá hạn ngày chốt có thể do khách hàng đang trình ký nội bộ hoặc đã mất thật.",
                    "AI không tự ý xóa bỏ vĩnh viễn mà trình danh sách bất thường lên giao diện để Quản lý xác minh thực tế."
                ]
            },
            {
                "title": "⚡ CƠ CHẾ GHI ĐÈ TỨC THÌ (human_overrides)",
                "color": COLOR_PINK,
                "bullets": [
                    "1. SET_CLOSED_LOST: Xác nhận mất khách -> Trọng số về 0%, loại khỏi dự báo.",
                    "2. SET_CLOSED_WON: Xác nhận đã ký hợp đồng -> Chuyển sang Doanh số thực tế.",
                    "3. VERIFY_AND_KEEP_OPEN: Đã xác minh bổ sung thông tin & gia hạn -> Giữ mở và cho phép xét Top 5.",
                    "Khi bấm 'Áp Dụng Quyết Định & Chạy Lại', AI tính toán lại toàn bộ KPI trong < 1 giây."
                ]
            }
        ],
        "notes": "Ở Bước 4, Trưởng phòng Kinh doanh có thể phê duyệt từng ca nghi vấn ngay trên bảng giao diện và kích hoạt AI tính toán lại tức thì."
    },
    {
        "num": 13,
        "tag": "SLIDE 13 / 15 • XÂY DỰNG SKILL 3",
        "title": "Skill 3: Đánh Giá KPI 10 Sales & Thuật Toán Chấm Điểm Top 5",
        "subtitle": "Công thức xếp hạng 4 tiêu chí (100 điểm) và kiểm soát chặt chẽ ranh giới quy tắc (< 5 cơ hội)",
        "accent": COLOR_EMERALD,
        "cards": [
            {
                "title": "🏆 CÔNG THỨC CHẤM ĐIỂM ƯU TIÊN (THANG 100 ĐIỂM)",
                "color": COLOR_EMERALD,
                "bullets": [
                    "• 35 điểm — Giá trị dự kiến có trọng số (Weighted Value): Ưu tiên thương vụ mang lại dòng tiền lớn.",
                    "• 25 điểm — Độ chín giai đoạn (Stage Readiness): Closing (25đ), Negotiation (20đ), Proposal (14đ)...",
                    "• 20 điểm — Độ khẩn cấp ngày chốt (Urgency): Dự kiến chốt trong 0–7 ngày tới đạt tối đa 20đ.",
                    "• 20 điểm — Hỗ trợ thu hẹp thiếu hụt KPI của NVKD (KPI Gap Support)."
                ]
            },
            {
                "title": "🚧 ĐIỀU KIỆN LỌC CỨNG & RANH GIỚI < 5 CƠ HỘI",
                "color": COLOR_AMBER,
                "bullets": [
                    "Điều kiện tiên quyết: Chỉ xét cơ hội Đang mở + Đủ 3 thông tin xác nhận (Người quyết định, Ngân sách, Bước tiếp theo) + Không quá hạn + Có tương tác trong 14 ngày.",
                    "Quy tắc Ranh giới: Nếu toàn công ty chỉ có < 5 cơ hội hợp lệ (VD: 3 cơ hội), chỉ xuất đúng 3 cơ hội đó và giải thích rõ danh sách các deal lớn bị loại trừ."
                ]
            }
        ],
        "notes": "Thuật toán chấm điểm Top 5 kết hợp cả 4 yếu tố: giá trị trọng số, độ chín giai đoạn, độ khẩn cấp ngày chốt và mức thiếu hụt KPI của nhân viên phụ trách."
    },
    {
        "num": 14,
        "tag": "SLIDE 14 / 15 • TÍCH HỢP GEMINI AI & GIAO DIỆN UI/UX",
        "title": "Tích Hợp Google Gemini 3.5 Flash & Thiết Kế Giao Diện Tối Giản",
        "subtitle": "Cơ chế tự phục hồi (Fallback) khi mất mạng và trải nghiệm người dùng chuẩn Google Chrome + Smallpdf",
        "accent": COLOR_PURPLE,
        "cards": [
            {
                "title": "🤖 GEMINI LIVE BRIEF & MULTI-MODEL FALLBACK",
                "color": COLOR_PURPLE,
                "bullets": [
                    "Truyền JSON Summary đã tính chuẩn vào Gemini API để viết Báo cáo điều hành 3 phần.",
                    "Xoay vòng tự động: gemini-3.5-flash -> gemini-3-flash-preview -> gemini-flash-latest.",
                    "Nếu mất mạng hoặc hết hạn mức API, tự động kích hoạt Built-in Deterministic Synthesizer."
                ]
            },
            {
                "title": "🎨 GIAO DIỆN 3 MÀN HÌNH TỐI GIẢN & HIỆN ĐẠI",
                "color": COLOR_PINK,
                "bullets": [
                    "Màn hình 1 (Trang Bìa): Thương hiệu Công ty của Duydudi, đánh giá 4.9/5 Sao & nút BẮT ĐẦU.",
                    "Màn hình 2 (Tải Tệp): Khung xanh dương #0070e0 chuẩn Smallpdf, trống sạch sẽ + Đánh giá 5 sao tương tác.",
                    "Màn hình 3 (Studio Dashboard): Biểu đồ trực quan, thanh Tab Chrome & chế độ Ngày / Đêm (☀️/🌙)."
                ]
            }
        ],
        "notes": "Giao diện được tối ưu hóa cực kỳ tinh gọn: từ Trang bìa bấm Bắt đầu -> Sang trang tải tệp chuẩn Smallpdf -> Kiểm tra/Sửa lỗi bằng AI -> Hiển thị Dashboard."
    },
    {
        "num": 15,
        "tag": "SLIDE 15 / 15 • KIỂM THỬ TỰ ĐỘNG & TỔNG KẾT",
        "title": "Kiểm Thử Tự Động 3 Kịch Bản, Triển Khai Cloud & Kết Luận",
        "subtitle": "Khẳng định độ tin cậy với bộ Self-Test 100% PASS và khả năng vận hành trên môi trường Serverless",
        "accent": COLOR_EMERALD,
        "cards": [
            {
                "title": "✅ BỘ SELF-TEST TỰ ĐỘNG 3 KỊCH BẢN (100% PASS)",
                "color": COLOR_EMERALD,
                "bullets": [
                    "Kịch bản 1 (Normal Case): Gộp trùng chuẩn, tách đơn hàng đã chốt, xuất đủ 5/5 cơ hội ưu tiên -> PASS.",
                    "Kịch bản 2 (Missing & Conflict): Bắt chính xác cơ hội 1.2 tỷ thiếu ngân sách và các ca mâu thuẫn -> PASS.",
                    "Kịch bản 3 (Boundary < 5): Lọc đúng ngày chốt sổ và chỉ xuất 3/5 cơ hội hợp lệ thực tế -> PASS."
                ]
            },
            {
                "title": "🚀 SẴN SÀNG CLOUD (VERCEL) & GIÁ TRỊ THỰC TIỄN",
                "color": COLOR_CYAN,
                "bullets": [
                    "Tự động nhận diện môi trường Serverless (/tmp/logs) để chạy mượt mà trên cả Local lẫn Vercel.",
                    "Rút ngắn thời gian làm báo cáo tuần của Trưởng phòng KD từ 4 giờ xuống còn 5 giây.",
                    "Chuẩn hóa kỷ luật dữ liệu và loại bỏ 100% doanh số ảo cho Công ty của Duydudi."
                ]
            }
        ],
        "notes": "Hệ thống đã vượt qua 100% các bài kiểm thử tự động trên cả 3 kịch bản và sẵn sàng vận hành thực tế. Cảm ơn mọi người đã theo dõi bài thuyết trình!"
    }
]


def build_pptx_presentation(output_path: str) -> str:
    prs = Presentation()
    # Thiết lập tỷ lệ màn hình rộng 16:9 chuẩn hiện đại (13.333 x 7.5 inches)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for s_data in SLIDES_CONTENT:
        slide = prs.slides.add_slide(blank_layout)

        # 1. Nền Slide tối hiện đại (Dark Slate Background)
        bg_shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5)
        )
        bg_shape.fill.solid()
        bg_shape.fill.fore_color.rgb = COLOR_BG_DARK
        bg_shape.line.fill.background()

        # 2. Thanh trang trí đỉnh Slide (Top Accent Bar)
        top_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.12)
        )
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = s_data["accent"]
        top_bar.line.fill.background()

        # 3. Huy hiệu Tag Slide (Slide Tag Pill)
        tag_box = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(0.4), Inches(4.5), Inches(0.38)
        )
        tag_box.fill.solid()
        tag_box.fill.fore_color.rgb = COLOR_HEADER_BG
        tag_box.line.color.rgb = s_data["accent"]
        tag_box.line.width = Pt(1.2)
        tf_tag = tag_box.text_frame
        tf_tag.word_wrap = True
        tf_tag.margin_left = Inches(0.15)
        tf_tag.margin_top = Inches(0.05)
        p_tag = tf_tag.paragraphs[0]
        p_tag.text = s_data["tag"]
        p_tag.font.name = "Arial"
        p_tag.font.size = Pt(10.5)
        p_tag.font.bold = True
        p_tag.font.color.rgb = s_data["accent"]

        # Thương hiệu góc phải trên
        brand_box = slide.shapes.add_textbox(Inches(9.0), Inches(0.38), Inches(3.7), Inches(0.4))
        tf_brand = brand_box.text_frame
        p_brand = tf_brand.paragraphs[0]
        p_brand.alignment = PP_ALIGN.RIGHT
        p_brand.text = "CÔNG TY CỦA DUYDUDI • AI SALES OPS"
        p_brand.font.name = "Arial"
        p_brand.font.size = Pt(10)
        p_brand.font.bold = True
        p_brand.font.color.rgb = COLOR_TEXT_MUTED

        # 4. Tiêu đề chính (Slide Title)
        title_box = slide.shapes.add_textbox(Inches(0.65), Inches(0.9), Inches(12.0), Inches(1.1))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = s_data["title"]
        p_title.font.name = "Arial"
        p_title.font.size = Pt(25 if "\n" not in s_data["title"] else 23)
        p_title.font.bold = True
        p_title.font.color.rgb = COLOR_WHITE

        # 5. Phụ đề (Slide Subtitle)
        sub_box = slide.shapes.add_textbox(Inches(0.65), Inches(1.88), Inches(12.0), Inches(0.55))
        tf_sub = sub_box.text_frame
        tf_sub.word_wrap = True
        p_sub = tf_sub.paragraphs[0]
        p_sub.text = s_data["subtitle"]
        p_sub.font.name = "Arial"
        p_sub.font.size = Pt(13)
        p_sub.font.bold = True
        p_sub.font.color.rgb = s_data["accent"]

        # 6. Các Thẻ Nội Dung (Content Cards)
        cards = s_data["cards"]
        num_cards = len(cards)
        total_w = 12.033
        gap = 0.28
        card_w = (total_w - gap * (num_cards - 1)) / num_cards
        start_x = 0.65
        card_y = 2.55
        card_h = 3.65

        for idx, c_info in enumerate(cards):
            cx = start_x + idx * (card_w + gap)
            card_shape = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(cx), Inches(card_y), Inches(card_w), Inches(card_h)
            )
            card_shape.fill.solid()
            card_shape.fill.fore_color.rgb = COLOR_CARD_BG
            card_shape.line.color.rgb = c_info["color"]
            card_shape.line.width = Pt(1.5)

            tf_card = card_shape.text_frame
            tf_card.word_wrap = True
            tf_card.vertical_anchor = MSO_ANCHOR.TOP
            tf_card.margin_left = Inches(0.22)
            tf_card.margin_right = Inches(0.22)
            tf_card.margin_top = Inches(0.22)
            tf_card.margin_bottom = Inches(0.18)

            # Tiêu đề Card
            p_ct = tf_card.paragraphs[0]
            p_ct.text = c_info["title"]
            p_ct.font.name = "Arial"
            p_ct.font.size = Pt(13.5 if num_cards <= 3 else 12)
            p_ct.font.bold = True
            p_ct.font.color.rgb = c_info["color"]
            p_ct.space_after = Pt(10)

            # Các dòng Bullet trong Card
            for b_text in c_info["bullets"]:
                p_b = tf_card.add_paragraph()
                p_b.text = f"•  {b_text}"
                p_b.font.name = "Arial"
                p_b.font.size = Pt(12 if num_cards == 2 else (11.5 if num_cards == 3 else 10.5))
                p_b.font.color.rgb = COLOR_TEXT_LIGHT
                p_b.space_after = Pt(8)

        # 7. Hộp Lời Dẫn Thuyết Trình Dưới Chân Slide (Speaker Notes Preview Bar)
        note_bar = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.65), Inches(6.38), Inches(12.033), Inches(0.62)
        )
        note_bar.fill.solid()
        note_bar.fill.fore_color.rgb = COLOR_HEADER_BG
        note_bar.line.color.rgb = COLOR_CARD_BORDER
        note_bar.line.width = Pt(1)
        tf_note = note_bar.text_frame
        tf_note.word_wrap = True
        tf_note.margin_left = Inches(0.18)
        tf_note.margin_top = Inches(0.1)
        p_note = tf_note.paragraphs[0]
        p_note.text = f"🎙️ Lời thuyết trình: {s_data['notes']}"
        p_note.font.name = "Arial"
        p_note.font.size = Pt(10)
        p_note.font.italic = True
        p_note.font.color.rgb = COLOR_TEXT_LIGHT

        # 8. Footer số trang
        footer_box = slide.shapes.add_textbox(Inches(0.65), Inches(7.05), Inches(12.033), Inches(0.35))
        tf_foot = footer_box.text_frame
        p_foot = tf_foot.paragraphs[0]
        p_foot.alignment = PP_ALIGN.RIGHT
        p_foot.text = f"Duydudi AI Agent (http://localhost:8000)  |  Trang {s_data['num']} / 15"
        p_foot.font.name = "Arial"
        p_foot.font.size = Pt(9.5)
        p_foot.font.bold = True
        p_foot.font.color.rgb = COLOR_TEXT_MUTED

        # 9. Ghi vào Speaker Notes chính thức của PowerPoint
        notes_slide = slide.notes_slide
        notes_tf = notes_slide.notes_text_frame
        notes_tf.text = f"[SLIDE {s_data['num']}: {s_data['title']}]\n\nLời dẫn thuyết trình:\n{s_data['notes']}"

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    prs.save(output_path)
    return output_path


if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    out1 = os.path.join(base_dir, "Duydudi_AI_Agent_15_Slides.pptx")
    out2 = os.path.join(base_dir, "docs", "Duydudi_AI_Agent_15_Slides.pptx")
    build_pptx_presentation(out1)
    build_pptx_presentation(out2)
    print(f"Created PPTX successfully: {out1}")

