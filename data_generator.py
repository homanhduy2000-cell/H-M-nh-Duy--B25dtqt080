import os
import json
import random
from datetime import datetime, timedelta
import pandas as pd

DEFAULT_REPORT_DATE = "2026-11-08"

STAGE_RULES = {
    "Lead": {"weight": 0.10, "label": "Khởi tạo (Lead)", "order": 1},
    "Qualification": {"weight": 0.25, "label": "Khảo sát & Xác nhận nhu cầu", "order": 2},
    "Proposal": {"weight": 0.50, "label": "Gửi đề xuất & Báo giá", "order": 3},
    "Negotiation": {"weight": 0.75, "label": "Đàm phán hợp đồng", "order": 4},
    "Closing": {"weight": 0.90, "label": "Chờ ký kết (Closing)", "order": 5},
    "Closed Won": {"weight": 1.00, "label": "Thành công (Đã chốt)", "order": 6},
    "Closed Lost": {"weight": 0.00, "label": "Thất bại (Mất khách)", "order": 7},
}

SALES_REPS = [
    {"rep_id": "NV01", "rep_name": "Nguyễn Văn An", "team": "Enterprise B2B", "kpi_target": 2500000000, "period": "Q4/2026"},
    {"rep_id": "NV02", "rep_name": "Trần Thị Bích", "team": "Enterprise B2B", "kpi_target": 2200000000, "period": "Q4/2026"},
    {"rep_id": "NV03", "rep_name": "Lê Hoàng Cường", "team": "Enterprise B2B", "kpi_target": 2000000000, "period": "Q4/2026"},
    {"rep_id": "NV04", "rep_name": "Phạm Minh Dung", "team": "Cloud & AI Solutions", "kpi_target": 1800000000, "period": "Q4/2026"},
    {"rep_id": "NV05", "rep_name": "Hoàng Quốc Đạt", "team": "Cloud & AI Solutions", "kpi_target": 1800000000, "period": "Q4/2026"},
    {"rep_id": "NV06", "rep_name": "Vũ Thu Hà", "team": "Cloud & AI Solutions", "kpi_target": 1600000000, "period": "Q4/2026"},
    {"rep_id": "NV07", "rep_name": "Đặng Quang Huy", "team": "SMB Tech Services", "kpi_target": 1500000000, "period": "Q4/2026"},
    {"rep_id": "NV08", "rep_name": "Bùi Thanh Lan", "team": "SMB Tech Services", "kpi_target": 1400000000, "period": "Q4/2026"},
    {"rep_id": "NV09", "rep_name": "Đỗ Tuấn Minh", "team": "SMB Tech Services", "kpi_target": 1300000000, "period": "Q4/2026"},
    {"rep_id": "NV10", "rep_name": "Ngô Phương Nga", "team": "SMB Tech Services", "kpi_target": 1200000000, "period": "Q4/2026"},
]

CUSTOMERS = [
    "Tập đoàn Viễn thông Á Châu", "Ngân hàng TMCP Đông Nam", "Công ty CP Logistics Toàn Cầu",
    "Bệnh viện Quốc tế Hòa Bình", "Tập đoàn Bán lẻ Sao Việt", "Công ty CP Thép Hòa Phát Mới",
    "Tổng Công ty Điện lực Miền Bắc", "Công ty Bảo hiểm Nhân thọ An Tâm", "Tập đoàn Giáo dục Tri Thức",
    "Công ty CP Dược phẩm Trung Ương", "Công ty TNHH Sản xuất Linh kiện Điện tử", "Tập đoàn Bất động sản Hưng Thịnh",
    "Công ty CP Hàng không Việt Á", "Chuỗi Siêu thị Xanh Mart", "Công ty Chứng khoán Phú Hưng",
    "Công ty CP Thực phẩm Sạch", "Tập đoàn Xây dựng Thái Bình", "Công ty Vận tải Biển Đông",
    "Ngân hàng Số Thế Hệ Mới", "Công ty CP Cảng Quốc tế", "Tập đoàn Dệt may Thành Công",
    "Công ty CP Xi măng Hoàng Mai", "Hệ thống Phòng khám Đa khoa Tâm Đức", "Công ty Phần mềm Tài chính FinTech",
    "Tập đoàn Nông nghiệp Công nghệ cao", "Công ty CP Nhựa Bình Minh Mới", "Công ty Du lịch Lữ hành Quốc tế",
    "Công ty CP Đầu tư Hạ tầng Đô thị", "Tập đoàn Năng lượng Xanh", "Công ty Thương mại Điện tử Nhanh"
]

SOLUTIONS = [
    "Hệ thống ERP Cloud SAP/Odoo B2B",
    "Nền tảng CRM & AI Call Center",
    "Dịch vụ SOC & Bảo mật An toàn thông tin",
    "Hạ tầng Data Warehouse & BI Dashboard",
    "Giải pháp Tự động hóa Quy trình RPA",
    "Hệ thống Quản lý Chuỗi cung ứng SCM",
    "Dịch vụ Cloud Migration & Managed K8s",
    "Giải pháp AI OCR & Số hóa Tài liệu"
]

DECISION_MAKERS = [
    "Ông Trần Quốc Tuấn (CIO)", "Bà Nguyễn Mai Hương (CFO)", "Ông Lê Văn Kiên (TGĐ)",
    "Bà Phạm Thu Thủy (GĐ CNTT)", "Ông Vũ Đình Trọng (GĐ Vận hành)", "Ông Hoàng Minh Đức (Phó TGĐ)"
]


def _save_bundle(folder_path: str, name_prefix: str, opps_df: pd.DataFrame, acts_df: pd.DataFrame, orders_df: pd.DataFrame, kpi_df: pd.DataFrame):
    os.makedirs(folder_path, exist_ok=True)
    opps_df.to_csv(os.path.join(folder_path, "opportunities.csv"), index=False, encoding="utf-8-sig")
    acts_df.to_csv(os.path.join(folder_path, "activities.csv"), index=False, encoding="utf-8-sig")
    orders_df.to_csv(os.path.join(folder_path, "orders.csv"), index=False, encoding="utf-8-sig")
    kpi_df.to_csv(os.path.join(folder_path, "kpi_targets.csv"), index=False, encoding="utf-8-sig")

    xlsx_path = os.path.join(folder_path, f"{name_prefix}.xlsx")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        opps_df.to_excel(writer, sheet_name="Opportunities", index=False)
        acts_df.to_excel(writer, sheet_name="Activities", index=False)
        orders_df.to_excel(writer, sheet_name="Orders", index=False)
        kpi_df.to_excel(writer, sheet_name="KPI_Targets", index=False)


def build_core_featured_records():
    """
    Tạo tập các bản ghi điển hình chứa đầy đủ mọi tình huống của Đề 09:
    - Cơ hội hợp lệ đủ tiêu chuẩn ưu tiên (Top Priority Candidates)
    - Cơ hội trùng lặp ở nhiều tệp (Duplicates across files)
    - Cơ hội đã chốt đơn trong Orders nhưng trạng thái vẫn mở (Double counting)
    - Cơ hội khách đã từ chối trong Activities nhưng chưa đổi sang Closed Lost (Zombie Lost)
    - Cơ hội quá hạn ngày dự kiến chốt và không có hoạt động mới (Overdue & Inactive)
    - Cơ hội giá trị rất lớn nhưng thiếu thông tin xác nhận (Unverified Big Deals)
    - Cơ hội sử dụng tỷ lệ dự kiến tùy ý sai lệch quy tắc (Arbitrary Probability)
    """
    opps = [
        # --- NHÓM A: CÁC CƠ HỘI HỢP LỆ, ĐẦY ĐỦ THÔNG TIN XÁC NHẬN (ỨNG VIÊN TOP 5 ƯU TIÊN) ---
        {
            "opp_id": "OPP-001",
            "customer_name": "Ngân hàng TMCP Đông Nam",
            "project_name": "Dịch vụ SOC & Bảo mật An toàn thông tin",
            "rep_id": "NV01",
            "stage": "Closing",
            "deal_value": 950000000,
            "sales_custom_prob": 0.95,
            "created_date": "2026-09-10",
            "updated_date": "2026-11-06",
            "expected_close_date": "2026-11-13",
            "decision_maker": "Ông Trần Quốc Tuấn (CIO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Họp rà soát điều khoản SLA cuối cùng và trình ký HĐ ngày 12/11",
            "source_file": "CRM_Weekly_NV01.xlsx"
        },
        {
            "opp_id": "OPP-002",
            "customer_name": "Tập đoàn Viễn thông Á Châu",
            "project_name": "Hạ tầng Data Warehouse & BI Dashboard",
            "rep_id": "NV02",
            "stage": "Negotiation",
            "deal_value": 1100000000,
            "sales_custom_prob": 0.90,
            "created_date": "2026-09-15",
            "updated_date": "2026-11-05",
            "expected_close_date": "2026-11-15",
            "decision_maker": "Bà Nguyễn Mai Hương (CFO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Chốt phương án thanh toán 3 đợt với Phòng Mua sắm",
            "source_file": "CRM_Weekly_NV02.xlsx"
        },
        {
            "opp_id": "OPP-003",
            "customer_name": "Công ty CP Logistics Toàn Cầu",
            "project_name": "Hệ thống Quản lý Chuỗi cung ứng SCM",
            "rep_id": "NV04",
            "stage": "Closing",
            "deal_value": 720000000,
            "sales_custom_prob": 0.85,
            "created_date": "2026-09-20",
            "updated_date": "2026-11-07",
            "expected_close_date": "2026-11-12",
            "decision_maker": "Ông Vũ Đình Trọng (GĐ Vận hành)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Gửi bản cứng hợp đồng cho Pháp chế khách hàng đóng dấu",
            "source_file": "CRM_Weekly_NV04.xlsx"
        },
        {
            "opp_id": "OPP-004",
            "customer_name": "Bệnh viện Quốc tế Hòa Bình",
            "project_name": "Giải pháp AI OCR & Số hóa Tài liệu",
            "rep_id": "NV05",
            "stage": "Negotiation",
            "deal_value": 680000000,
            "sales_custom_prob": 0.80,
            "created_date": "2026-09-25",
            "updated_date": "2026-11-04",
            "expected_close_date": "2026-11-14",
            "decision_maker": "Bà Phạm Thu Thủy (GĐ CNTT)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Demo nghiệm thu PoC lần cuối với Ban Giám đốc Bệnh viện",
            "source_file": "CRM_Weekly_NV05.xlsx"
        },
        {
            "opp_id": "OPP-005",
            "customer_name": "Tập đoàn Bán lẻ Sao Việt",
            "project_name": "Nền tảng CRM & AI Call Center",
            "rep_id": "NV03",
            "stage": "Proposal",
            "deal_value": 890000000,
            "sales_custom_prob": 0.85,  # Sales tự điền 85% dù mới ở Proposal (규 tắc là 50%)
            "created_date": "2026-10-01",
            "updated_date": "2026-11-06",
            "expected_close_date": "2026-11-16",
            "decision_maker": "Ông Lê Văn Kiên (TGĐ)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Trình bày bảng ROI và phương án tích hợp tổng đài hiện hữu",
            "source_file": "CRM_Weekly_NV03.xlsx"
        },
        {
            "opp_id": "OPP-006",
            "customer_name": "Công ty Bảo hiểm Nhân thọ An Tâm",
            "project_name": "Giải pháp Tự động hóa Quy trình RPA",
            "rep_id": "NV06",
            "stage": "Negotiation",
            "deal_value": 540000000,
            "sales_custom_prob": 0.75,
            "created_date": "2026-10-05",
            "updated_date": "2026-11-05",
            "expected_close_date": "2026-11-18",
            "decision_maker": "Ông Hoàng Minh Đức (Phó TGĐ)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Thương lượng gói bảo trì 24/7 năm đầu tiên",
            "source_file": "CRM_Weekly_NV06.xlsx"
        },

        # --- NHÓM B: BẢN GHI TRÙNG LẶP Ở NHIỀU TỆP (DUPLICATES) ---
        {
            "opp_id": "OPP-001",  # Trùng mã OPP-001 từ file cũ của NV01
            "customer_name": "Ngân hàng TMCP Đông Nam",
            "project_name": "Dịch vụ SOC & Bảo mật An toàn thông tin",
            "rep_id": "NV01",
            "stage": "Negotiation",
            "deal_value": 900000000,
            "sales_custom_prob": 0.80,
            "created_date": "2026-09-10",
            "updated_date": "2026-10-25",
            "expected_close_date": "2026-11-10",
            "decision_maker": "Ông Trần Quốc Tuấn (CIO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Đang đàm phán giá",
            "source_file": "Backup_Sheet_Thang10.csv"
        },
        {
            "opp_id": "OPP-002-DUP",  # Khác mã nhưng trùng Khách hàng + Dự án với OPP-002
            "customer_name": "Tập đoàn Viễn thông Á Châu",
            "project_name": "Hạ tầng Data Warehouse & BI Dashboard",
            "rep_id": "NV02",
            "stage": "Proposal",
            "deal_value": 1100000000,
            "sales_custom_prob": 0.70,
            "created_date": "2026-09-15",
            "updated_date": "2026-10-20",
            "expected_close_date": "2026-11-15",
            "decision_maker": "Bà Nguyễn Mai Hương (CFO)",
            "budget_confirmed": "Chưa xác nhận",
            "next_step": "Gửi báo giá lần 1",
            "source_file": "Master_Tracking_Old.xlsx"
        },

        # --- NHÓM C: ĐÃ CHỐT ĐƠN HÀNG (CÓ TRONG ORDERS) NHƯNG VẪN ĐỂ MỞ TRONG PIPELINE (DOUBLE COUNTING) ---
        {
            "opp_id": "OPP-007",
            "customer_name": "Công ty CP Thép Hòa Phát Mới",
            "project_name": "Hệ thống ERP Cloud SAP/Odoo B2B",
            "rep_id": "NV01",
            "stage": "Closing",  # Sai! Đã có đơn hàng ORD-1001 ngày 02/11 nhưng chưa chuyển Closed Won
            "deal_value": 1250000000,
            "sales_custom_prob": 0.95,
            "created_date": "2026-08-12",
            "updated_date": "2026-11-01",
            "expected_close_date": "2026-11-05",
            "decision_maker": "Ông Lê Văn Kiên (TGĐ)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Chờ kế toán xuất hóa đơn đợt 1",
            "source_file": "CRM_Weekly_NV01.xlsx"
        },
        {
            "opp_id": "OPP-008",
            "customer_name": "Tổng Công ty Điện lực Miền Bắc",
            "project_name": "Dịch vụ Cloud Migration & Managed K8s",
            "rep_id": "NV03",
            "stage": "Negotiation",  # Đã có đơn hàng ORD-1002 ngày 04/11 nhưng Sales quên đóng cơ hội
            "deal_value": 820000000,
            "sales_custom_prob": 0.90,
            "created_date": "2026-08-20",
            "updated_date": "2026-10-29",
            "expected_close_date": "2026-11-04",
            "decision_maker": "Ông Trần Quốc Tuấn (CIO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Đã ký hợp đồng, bàn giao đội triển khai",
            "source_file": "CRM_Weekly_NV03.xlsx"
        },

        # --- NHÓM D: ĐÃ MẤT KHÁCH (TRONG ACTIVITIES GHI TỪ CHỐI/CHỌN ĐỐI THỦ) NHƯNG CHƯA ĐỔI TRẠNG THÁI ---
        {
            "opp_id": "OPP-009",
            "customer_name": "Tập đoàn Giáo dục Tri Thức",
            "project_name": "Nền tảng CRM & AI Call Center",
            "rep_id": "NV02",
            "stage": "Negotiation",  # Thực tế đã mất khách từ 28/10
            "deal_value": 640000000,
            "sales_custom_prob": 0.75,
            "created_date": "2026-09-01",
            "updated_date": "2026-10-22",
            "expected_close_date": "2026-11-10",
            "decision_maker": "Bà Phạm Thu Thủy (GĐ CNTT)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Chờ phản hồi báo giá cạnh tranh",
            "source_file": "CRM_Weekly_NV02.xlsx"
        },
        {
            "opp_id": "OPP-010",
            "customer_name": "Công ty CP Hàng không Việt Á",
            "project_name": "Dịch vụ SOC & Bảo mật An toàn thông tin",
            "rep_id": "NV07",
            "stage": "Proposal",  # Thực tế khách đã hủy ngân sách dự án năm 2026
            "deal_value": 780000000,
            "sales_custom_prob": 0.80,
            "created_date": "2026-09-10",
            "updated_date": "2026-10-18",
            "expected_close_date": "2026-11-09",
            "decision_maker": "Ông Hoàng Minh Đức (Phó TGĐ)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Theo dõi lại",
            "source_file": "CRM_Weekly_NV07.xlsx"
        },

        # --- NHÓM E: CƠ HỘI GIÁ TRỊ RẤT LỚN NHƯNG THIẾU THÔNG TIN XÁC NHẬN (BÁNH VẼ / UNVERIFIED BIG DEALS) ---
        {
            "opp_id": "OPP-011",
            "customer_name": "Tập đoàn Bất động sản Hưng Thịnh",
            "project_name": "Hệ thống ERP Cloud SAP/Odoo B2B",
            "rep_id": "NV08",
            "stage": "Proposal",
            "deal_value": 2400000000,  # Giá trị cực lớn 2.4 tỷ nhưng thiếu người quyết định, chưa xác nhận ngân sách, thiếu next step
            "sales_custom_prob": 0.90,  # NVKD tự điền 90%
            "created_date": "2026-10-12",
            "updated_date": "2026-11-02",
            "expected_close_date": "2026-11-15",
            "decision_maker": "",
            "budget_confirmed": "Chưa xác nhận",
            "next_step": "",
            "source_file": "CRM_Weekly_NV08.xlsx"
        },
        {
            "opp_id": "OPP-012",
            "customer_name": "Tập đoàn Xây dựng Thái Bình",
            "project_name": "Hạ tầng Data Warehouse & BI Dashboard",
            "rep_id": "NV09",
            "stage": "Qualification",
            "deal_value": 1950000000,  # Giá trị lớn 1.95 tỷ nhưng thiếu thông tin xác nhận
            "sales_custom_prob": 0.85,
            "created_date": "2026-10-18",
            "updated_date": "2026-11-01",
            "expected_close_date": "2026-11-20",
            "decision_maker": "",
            "budget_confirmed": "Chưa rõ",
            "next_step": "",
            "source_file": "CRM_Weekly_NV09.xlsx"
        },

        # --- NHÓM F: QUÁ HẠN NGÀY DỰ KIẾN CHỐT MÀ KHÔNG CÓ HOẠT ĐỘNG MỚI (OVERDUE & STALE) ---
        {
            "opp_id": "OPP-013",
            "customer_name": "Công ty CP Dược phẩm Trung Ương",
            "project_name": "Hệ thống Quản lý Chuỗi cung ứng SCM",
            "rep_id": "NV04",
            "stage": "Negotiation",
            "deal_value": 560000000,
            "sales_custom_prob": 0.75,
            "created_date": "2026-08-15",
            "updated_date": "2026-10-05",
            "expected_close_date": "2026-10-20",  # Đã quá hạn 19 ngày so với 08/11/2026, không có liên hệ mới từ 05/10
            "decision_maker": "Ông Vũ Đình Trọng (GĐ Vận hành)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Hẹn gặp lại sau",
            "source_file": "CRM_Weekly_NV04.xlsx"
        },
        {
            "opp_id": "OPP-014",
            "customer_name": "Chuỗi Siêu thị Xanh Mart",
            "project_name": "Giải pháp Tự động hóa Quy trình RPA",
            "rep_id": "NV10",
            "stage": "Proposal",
            "deal_value": 430000000,
            "sales_custom_prob": 0.70,
            "created_date": "2026-08-25",
            "updated_date": "2026-10-10",
            "expected_close_date": "2026-10-25",  # Quá hạn 14 ngày, không có tương tác mới
            "decision_maker": "Bà Nguyễn Mai Hương (CFO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Chờ khách xem báo giá",
            "source_file": "CRM_Weekly_NV10.xlsx"
        },

        # --- NHÓM G: ĐÃ CHỐT THÀNH CÔNG CHUẨN (CLOSED WON) ---
        {
            "opp_id": "OPP-015",
            "customer_name": "Công ty Chứng khoán Phú Hưng",
            "project_name": "Dịch vụ SOC & Bảo mật An toàn thông tin",
            "rep_id": "NV02",
            "stage": "Closed Won",
            "deal_value": 880000000,
            "sales_custom_prob": 1.0,
            "created_date": "2026-08-05",
            "updated_date": "2026-10-28",
            "expected_close_date": "2026-10-28",
            "decision_maker": "Ông Trần Quốc Tuấn (CIO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Đã nghiệm thu giai đoạn 1",
            "source_file": "CRM_Weekly_NV02.xlsx"
        },
        {
            "opp_id": "OPP-016",
            "customer_name": "Ngân hàng Số Thế Hệ Mới",
            "project_name": "Giải pháp AI OCR & Số hóa Tài liệu",
            "rep_id": "NV04",
            "stage": "Closed Won",
            "deal_value": 760000000,
            "sales_custom_prob": 1.0,
            "created_date": "2026-08-10",
            "updated_date": "2026-10-30",
            "expected_close_date": "2026-10-30",
            "decision_maker": "Bà Phạm Thu Thủy (GĐ CNTT)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Hoàn tất bàn giao license",
            "source_file": "CRM_Weekly_NV04.xlsx"
        },
    ]

    activities = [
        {"activity_id": "ACT-001", "opp_id": "OPP-001", "activity_date": "2026-11-06", "activity_type": "Meeting", "outcome_signal": "POSITIVE", "notes": "Họp với CIO Trần Quốc Tuấn, khách đồng ý toàn bộ phạm vi kỹ thuật, hẹn ký HĐ tuần tới."},
        {"activity_id": "ACT-002", "opp_id": "OPP-002", "activity_date": "2026-11-05", "activity_type": "Call", "outcome_signal": "POSITIVE", "notes": "CFO xác nhận ngân sách Q4 đã duyệt, đang rà soát điều khoản bảo hành."},
        {"activity_id": "ACT-003", "opp_id": "OPP-003", "activity_date": "2026-11-07", "activity_type": "Email", "outcome_signal": "POSITIVE", "notes": "Pháp chế khách hàng đã duyệt dự thảo hợp đồng SCM."},
        {"activity_id": "ACT-004", "opp_id": "OPP-004", "activity_date": "2026-11-04", "activity_type": "Demo", "outcome_signal": "POSITIVE", "notes": "Demo PoC AI OCR đạt độ chính xác 98.5%, khách hàng rất hài lòng."},
        {"activity_id": "ACT-005", "opp_id": "OPP-005", "activity_date": "2026-11-06", "activity_type": "Meeting", "outcome_signal": "POSITIVE", "notes": "TGĐ Lê Văn Kiên yêu cầu gửi phương án tích hợp chi tiết để chốt trong tuần tới."},
        {"activity_id": "ACT-006", "opp_id": "OPP-006", "activity_date": "2026-11-05", "activity_type": "Call", "outcome_signal": "NEUTRAL", "notes": "Đang thương lượng phí bảo trì năm đầu."},
        {"activity_id": "ACT-007", "opp_id": "OPP-007", "activity_date": "2026-11-02", "activity_type": "Contract", "outcome_signal": "WON", "notes": "Khách hàng đã ký hợp đồng chính thức và chuyển khoản tạm ứng đợt 1."},
        {"activity_id": "ACT-008", "opp_id": "OPP-008", "activity_date": "2026-11-04", "activity_type": "Contract", "outcome_signal": "WON", "notes": "Đã ký biên bản nghiệm thu đơn hàng ORD-1002."},
        {"activity_id": "ACT-009", "opp_id": "OPP-009", "activity_date": "2026-10-28", "activity_type": "Call", "outcome_signal": "LOST", "notes": "Khách hàng thông báo đã ký với đối thủ cạnh tranh do giá thấp hơn 20%. Đã mất khách."},
        {"activity_id": "ACT-010", "opp_id": "OPP-010", "activity_date": "2026-10-25", "activity_type": "Email", "outcome_signal": "LOST", "notes": "Khách hàng chính thức hủy dự án năm nay, từ chối tiếp tục triển khai."},
        {"activity_id": "ACT-011", "opp_id": "OPP-011", "activity_date": "2026-11-02", "activity_type": "Email", "outcome_signal": "NEUTRAL", "notes": "Mới gửi brochure giới thiệu qua email cho chuyên viên IT, chưa gặp được lãnh đạo."},
        {"activity_id": "ACT-012", "opp_id": "OPP-012", "activity_date": "2026-11-01", "activity_type": "Call", "outcome_signal": "NEUTRAL", "notes": "Gọi điện hỏi thăm sơ bộ, khách chưa có ngân sách cụ thể."},
        {"activity_id": "ACT-013", "opp_id": "OPP-013", "activity_date": "2026-10-05", "activity_type": "Call", "outcome_signal": "STALLED", "notes": "Gọi 3 lần khách báo bận, chưa phản hồi lại từ đầu tháng 10."},
        {"activity_id": "ACT-014", "opp_id": "OPP-014", "activity_date": "2026-10-10", "activity_type": "Email", "outcome_signal": "STALLED", "notes": "Đã gửi báo giá từ đầu tháng 10 nhưng khách không trả lời email."},
        {"activity_id": "ACT-015", "opp_id": "OPP-015", "activity_date": "2026-10-28", "activity_type": "Contract", "outcome_signal": "WON", "notes": "Hoàn tất ký kết hợp đồng."},
        {"activity_id": "ACT-016", "opp_id": "OPP-016", "activity_date": "2026-10-30", "activity_type": "Contract", "outcome_signal": "WON", "notes": "Hoàn tất đơn hàng và ghi nhận doanh số."},
    ]

    orders = [
        {"order_id": "ORD-1001", "opp_id": "OPP-007", "customer_name": "Công ty CP Thép Hòa Phát Mới", "rep_id": "NV01", "order_amount": 1250000000, "order_date": "2026-11-02", "order_status": "Approved"},
        {"order_id": "ORD-1002", "opp_id": "OPP-008", "customer_name": "Tổng Công ty Điện lực Miền Bắc", "rep_id": "NV03", "order_amount": 820000000, "order_date": "2026-11-04", "order_status": "Approved"},
        {"order_id": "ORD-1003", "opp_id": "OPP-015", "customer_name": "Công ty Chứng khoán Phú Hưng", "rep_id": "NV02", "order_amount": 880000000, "order_date": "2026-10-28", "order_status": "Approved"},
        {"order_id": "ORD-1004", "opp_id": "OPP-016", "customer_name": "Ngân hàng Số Thế Hệ Mới", "rep_id": "NV04", "order_amount": 760000000, "order_date": "2026-10-30", "order_status": "Approved"},
    ]

    return opps, activities, orders


def generate_all_datasets(base_dir: str):
    os.makedirs(base_dir, exist_ok=True)
    kpi_df = pd.DataFrame(SALES_REPS)

    # 1. NORMAL CASE (~45 opportunities across 10 reps)
    random.seed(42)
    base_opps, base_acts, base_orders = build_core_featured_records()
    normal_opps = list(base_opps)
    normal_acts = list(base_acts)
    normal_orders = list(base_orders)

    stages_pool = ["Lead", "Qualification", "Proposal", "Negotiation", "Closing", "Closed Won", "Closed Lost"]
    for idx in range(17, 46):
        opp_id = f"OPP-{idx:03d}"
        rep = SALES_REPS[idx % len(SALES_REPS)]["rep_id"]
        cust = f"{CUSTOMERS[idx % len(CUSTOMERS)]} - CN {(idx // 10) + 1}"
        proj = SOLUTIONS[idx % len(SOLUTIONS)]
        st = stages_pool[idx % len(stages_pool)]
        val = random.choice([250, 320, 400, 480, 520, 600]) * 1000000
        custom_prob = min(0.95, STAGE_RULES[st]["weight"] + random.choice([0.0, 0.15, 0.20]))
        exp_date = (datetime.strptime(DEFAULT_REPORT_DATE, "%Y-%m-%d") + timedelta(days=random.randint(4, 25))).strftime("%Y-%m-%d")
        act_date = (datetime.strptime(DEFAULT_REPORT_DATE, "%Y-%m-%d") - timedelta(days=random.randint(1, 9))).strftime("%Y-%m-%d")

        if st == "Closed Won":
            exp_date = "2026-10-25"
            custom_prob = 1.0
            normal_orders.append({
                "order_id": f"ORD-10{idx:02d}",
                "opp_id": opp_id,
                "customer_name": cust,
                "rep_id": rep,
                "order_amount": val,
                "order_date": "2026-10-25",
                "order_status": "Approved"
            })
        elif st == "Closed Lost":
            exp_date = "2026-10-22"
            custom_prob = 0.0

        normal_opps.append({
            "opp_id": opp_id,
            "customer_name": cust,
            "project_name": proj,
            "rep_id": rep,
            "stage": st,
            "deal_value": val,
            "sales_custom_prob": round(custom_prob, 2),
            "created_date": "2026-09-12",
            "updated_date": act_date,
            "expected_close_date": exp_date,
            "decision_maker": DECISION_MAKERS[idx % len(DECISION_MAKERS)],
            "budget_confirmed": "Đã xác nhận" if st in ["Proposal", "Negotiation", "Closing", "Closed Won"] else "Chưa xác nhận",
            "next_step": "Tiếp tục bám sát lịch trình triển khai với khách hàng",
            "source_file": f"CRM_Weekly_{rep}.xlsx"
        })
        normal_acts.append({
            "activity_id": f"ACT-{idx:03d}",
            "opp_id": opp_id,
            "activity_date": act_date,
            "activity_type": random.choice(["Call", "Meeting", "Email", "Demo"]),
            "outcome_signal": "WON" if st == "Closed Won" else ("LOST" if st == "Closed Lost" else "POSITIVE"),
            "notes": f"Trao đổi định kỳ tiến độ dự án {proj} với khách hàng."
        })

    _save_bundle(
        os.path.join(base_dir, "scenario_1_normal"),
        "BinhMinh_Scenario1_Normal",
        pd.DataFrame(normal_opps),
        pd.DataFrame(normal_acts),
        pd.DataFrame(normal_orders),
        kpi_df
    )

    # 2. FULL 300 CONTEXT (10 reps, ~300 opportunities as described in prompt)
    full_opps = list(normal_opps)
    full_acts = list(normal_acts)
    full_orders = list(normal_orders)
    for idx in range(46, 301):
        opp_id = f"OPP-{idx:03d}"
        rep = SALES_REPS[idx % len(SALES_REPS)]["rep_id"]
        cust = f"{CUSTOMERS[idx % len(CUSTOMERS)]} (Khối {idx % 7 + 1})"
        proj = SOLUTIONS[idx % len(SOLUTIONS)]
        st = stages_pool[idx % len(stages_pool)]
        val = random.choice([120, 180, 240, 300, 360, 450]) * 1000000
        custom_prob = min(0.95, STAGE_RULES[st]["weight"] + random.choice([0.0, 0.10, 0.25]))

        # Inject periodic anomalies
        is_overdue = (idx % 15 == 0) and st not in ["Closed Won", "Closed Lost"]
        is_missing = (idx % 18 == 0) and st not in ["Closed Won", "Closed Lost"]

        if is_overdue:
            exp_date = "2026-10-15"
            act_date = "2026-10-02"
        else:
            exp_date = (datetime.strptime(DEFAULT_REPORT_DATE, "%Y-%m-%d") + timedelta(days=random.randint(5, 35))).strftime("%Y-%m-%d")
            act_date = (datetime.strptime(DEFAULT_REPORT_DATE, "%Y-%m-%d") - timedelta(days=random.randint(1, 10))).strftime("%Y-%m-%d")

        if st == "Closed Won":
            exp_date = "2026-10-20"
            custom_prob = 1.0
            full_orders.append({
                "order_id": f"ORD-1{idx:03d}",
                "opp_id": opp_id,
                "customer_name": cust,
                "rep_id": rep,
                "order_amount": val,
                "order_date": "2026-10-20",
                "order_status": "Approved"
            })

        full_opps.append({
            "opp_id": opp_id,
            "customer_name": cust,
            "project_name": proj,
            "rep_id": rep,
            "stage": st,
            "deal_value": val,
            "sales_custom_prob": round(custom_prob, 2),
            "created_date": "2026-08-20",
            "updated_date": act_date,
            "expected_close_date": exp_date,
            "decision_maker": "" if is_missing else DECISION_MAKERS[idx % len(DECISION_MAKERS)],
            "budget_confirmed": "Chưa xác nhận" if is_missing else ("Đã xác nhận" if st in ["Negotiation", "Closing", "Closed Won"] else "Chưa xác nhận"),
            "next_step": "" if is_missing else "Bám sát tiến độ thương thảo hợp đồng",
            "source_file": f"CRM_Weekly_{rep}.xlsx"
        })
        full_acts.append({
            "activity_id": f"ACT-{idx:03d}",
            "opp_id": opp_id,
            "activity_date": act_date,
            "activity_type": random.choice(["Call", "Meeting", "Email"]),
            "outcome_signal": "WON" if st == "Closed Won" else ("LOST" if st == "Closed Lost" else ("STALLED" if is_overdue else "POSITIVE")),
            "notes": "Khách chưa phản hồi từ đầu tháng 10" if is_overdue else f"Cập nhật tiến độ thương vụ {proj}."
        })

    _save_bundle(
        os.path.join(base_dir, "scenario_0_full300"),
        "BinhMinh_Full300_Context",
        pd.DataFrame(full_opps),
        pd.DataFrame(full_acts),
        pd.DataFrame(full_orders),
        kpi_df
    )

    # 3. MISSING & CONFLICT DATA CASE (Kịch bản kiểm thử 2: Thiếu dữ liệu & Mâu thuẫn nặng)
    missing_opps = [
        # Nhiều bản ghi thiếu giá trị, thiếu giai đoạn, thiếu ngày chốt, thiếu người quyết định, mâu thuẫn đơn hàng
        {"opp_id": "OPP-M01", "customer_name": "Tập đoàn Năng lượng Xanh", "project_name": "Hệ thống ERP Cloud SAP/Odoo B2B", "rep_id": "NV01", "stage": "Closing", "deal_value": 1500000000, "sales_custom_prob": 0.95, "created_date": "2026-09-10", "updated_date": "2026-11-05", "expected_close_date": "", "decision_maker": "", "budget_confirmed": "Chưa xác nhận", "next_step": "", "source_file": "File_Loi_NV01.csv"},
        {"opp_id": "OPP-M02", "customer_name": "Công ty CP Cảng Quốc tế", "project_name": "Dịch vụ SOC & Bảo mật An toàn thông tin", "rep_id": "NV02", "stage": "Negotiation", "deal_value": 0, "sales_custom_prob": 0.85, "created_date": "2026-09-12", "updated_date": "2026-11-04", "expected_close_date": "2026-11-14", "decision_maker": "Ông Trần Quốc Tuấn (CIO)", "budget_confirmed": "Đã xác nhận", "next_step": "Bổ sung báo giá", "source_file": "File_Loi_NV02.csv"},
        {"opp_id": "OPP-M03", "customer_name": "Công ty Phần mềm Tài chính FinTech", "project_name": "Dịch vụ Cloud Migration & Managed K8s", "rep_id": "NV03", "stage": "", "deal_value": 850000000, "sales_custom_prob": 0.80, "created_date": "2026-09-15", "updated_date": "2026-11-03", "expected_close_date": "2026-11-16", "decision_maker": "", "budget_confirmed": "", "next_step": "", "source_file": "File_Loi_NV03.csv"},
        # Mâu thuẫn: Trong Orders đã chốt 920 triệu nhưng trong Opportunities ghi Closed Lost!
        {"opp_id": "OPP-M04", "customer_name": "Tập đoàn Dệt may Thành Công", "project_name": "Giải pháp Tự động hóa Quy trình RPA", "rep_id": "NV04", "stage": "Closed Lost", "deal_value": 920000000, "sales_custom_prob": 0.0, "created_date": "2026-09-01", "updated_date": "2026-11-01", "expected_close_date": "2026-11-05", "decision_maker": "Bà Nguyễn Mai Hương (CFO)", "budget_confirmed": "Đã xác nhận", "next_step": "Đã ký đơn hàng ORD-M01", "source_file": "File_Loi_NV04.csv"},
        # Trùng lặp lệch giá trị lớn giữa 2 tệp (700tr vs 1.2 tỷ) và lệch nhân viên phụ trách
        {"opp_id": "OPP-M05", "customer_name": "Công ty CP Nhựa Bình Minh Mới", "project_name": "Hạ tầng Data Warehouse & BI Dashboard", "rep_id": "NV05", "stage": "Proposal", "deal_value": 700000000, "sales_custom_prob": 0.60, "created_date": "2026-10-01", "updated_date": "2026-10-25", "expected_close_date": "2026-11-18", "decision_maker": "Ông Lê Văn Kiên (TGĐ)", "budget_confirmed": "Đã xác nhận", "next_step": "Thuyết trình giải pháp", "source_file": "Sheet_Cua_NV05.xlsx"},
        {"opp_id": "OPP-M05", "customer_name": "Công ty CP Nhựa Bình Minh Mới", "project_name": "Hạ tầng Data Warehouse & BI Dashboard", "rep_id": "NV06", "stage": "Negotiation", "deal_value": 1200000000, "sales_custom_prob": 0.90, "created_date": "2026-10-01", "updated_date": "2026-11-06", "expected_close_date": "2026-11-18", "decision_maker": "Ông Lê Văn Kiên (TGĐ)", "budget_confirmed": "Đã xác nhận", "next_step": "Đàm phán hợp đồng mở rộng", "source_file": "Sheet_Cua_NV06.xlsx"},
    ] + base_opps[:8]

    missing_acts = [
        {"activity_id": "ACT-M01", "opp_id": "OPP-M01", "activity_date": "2026-11-05", "activity_type": "Call", "outcome_signal": "NEUTRAL", "notes": "Chưa xác định được người phê duyệt ngân sách cuối cùng."},
        {"activity_id": "ACT-M02", "opp_id": "OPP-M04", "activity_date": "2026-11-03", "activity_type": "Contract", "outcome_signal": "WON", "notes": "Khách đã ký hợp đồng ORD-M01 nhưng NVKD bấm nhầm sang Closed Lost trên file cá nhân."},
    ] + base_acts[:8]

    missing_orders = [
        {"order_id": "ORD-M01", "opp_id": "OPP-M04", "customer_name": "Tập đoàn Dệt may Thành Công", "rep_id": "NV04", "order_amount": 920000000, "order_date": "2026-11-03", "order_status": "Approved"},
    ] + base_orders[:2]

    _save_bundle(
        os.path.join(base_dir, "scenario_2_missing"),
        "BinhMinh_Scenario2_MissingConflict",
        pd.DataFrame(missing_opps),
        pd.DataFrame(missing_acts),
        pd.DataFrame(missing_orders),
        kpi_df
    )

    # 4. BOUNDARY / EDGE CASE (Kịch bản kiểm thử 3: Ranh giới quy tắc & < 5 cơ hội hợp lệ)
    # Kiểm chứng:
    # - Chỉ có đúng 3 cơ hội mở hợp lệ đủ điều kiện ưu tiên (< 5) -> Hệ thống phải nêu số thực tế (3/5) và giải thích lý do!
    # - Cơ hội có expected_close_date == report_date (2026-11-08) -> Đúng tại ranh giới, KHÔNG bị tính là quá hạn.
    # - Cơ hội có expected_close_date == 2026-11-07 (trước 1 ngày) -> Bị tính là quá hạn 1 ngày.
    # - Cơ hội có khoảng cách tương tác đúng bằng 14 ngày (2026-10-25 đến 2026-11-08) vs 15 ngày (2026-10-24).
    # - Đơn hàng và hoạt động phát sinh SAU ngày báo cáo (2026-11-10 > 2026-11-08) -> Phải bị lọc bỏ khi chốt sổ tại ngày 2026-11-08!
    boundary_opps = [
        # Cơ hội hợp lệ #1: Ngay tại ranh giới ngày dự kiến chốt == report_date (2026-11-08) và tương tác cách đúng 14 ngày (2026-10-25)
        {
            "opp_id": "OPP-B01",
            "customer_name": "Ngân hàng TMCP Đông Nam",
            "project_name": "Dịch vụ SOC & Bảo mật An toàn thông tin",
            "rep_id": "NV01",
            "stage": "Closing",
            "deal_value": 850000000,
            "sales_custom_prob": 0.90,
            "created_date": "2026-09-15",
            "updated_date": "2026-10-25",
            "expected_close_date": "2026-11-08",  # Đúng bằng report_date (Boundary!)
            "decision_maker": "Ông Trần Quốc Tuấn (CIO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Ký hợp đồng ngay trong đầu tuần mới",
            "source_file": "Boundary_Test.xlsx"
        },
        # Cơ hội hợp lệ #2
        {
            "opp_id": "OPP-B02",
            "customer_name": "Tập đoàn Viễn thông Á Châu",
            "project_name": "Hạ tầng Data Warehouse & BI Dashboard",
            "rep_id": "NV02",
            "stage": "Negotiation",
            "deal_value": 760000000,
            "sales_custom_prob": 0.75,
            "created_date": "2026-09-20",
            "updated_date": "2026-11-06",
            "expected_close_date": "2026-11-12",
            "decision_maker": "Bà Nguyễn Mai Hương (CFO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Chốt phụ lục kỹ thuật và tiến độ thanh toán",
            "source_file": "Boundary_Test.xlsx"
        },
        # Cơ hội hợp lệ #3: Có đơn hàng phát sinh SAU ngày báo cáo (2026-11-10 > 2026-11-08), nên tại ngày báo cáo 08/11 vẫn phải tính là Cơ hội đang mở!
        {
            "opp_id": "OPP-B03",
            "customer_name": "Công ty CP Logistics Toàn Cầu",
            "project_name": "Hệ thống Quản lý Chuỗi cung ứng SCM",
            "rep_id": "NV03",
            "stage": "Closing",
            "deal_value": 690000000,
            "sales_custom_prob": 0.90,
            "created_date": "2026-09-25",
            "updated_date": "2026-11-07",
            "expected_close_date": "2026-11-10",
            "decision_maker": "Ông Vũ Đình Trọng (GĐ Vận hành)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Hoàn tất thủ tục ký hợp đồng ngày 10/11",
            "source_file": "Boundary_Test.xlsx"
        },
        # Cơ hội #4: Vượt ranh giới quá hạn đúng 1 ngày (expected_close_date = 2026-11-07 < 2026-11-08) và 15 ngày không tương tác (2026-10-24) -> Bị gắn cờ Quá hạn & Đóng băng!
        {
            "opp_id": "OPP-B04",
            "customer_name": "Bệnh viện Quốc tế Hòa Bình",
            "project_name": "Giải pháp AI OCR & Số hóa Tài liệu",
            "rep_id": "NV04",
            "stage": "Proposal",
            "deal_value": 920000000,
            "sales_custom_prob": 0.85,
            "created_date": "2026-09-10",
            "updated_date": "2026-10-24",
            "expected_close_date": "2026-11-07",  # Quá hạn đúng 1 ngày so với 2026-11-08
            "decision_maker": "Bà Phạm Thu Thủy (GĐ CNTT)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Liên hệ lại xác nhận lịch họp",
            "source_file": "Boundary_Test.xlsx"
        },
        # Cơ hội #5: Giá trị rất lớn (3.0 tỷ) nhưng thiếu thông tin xác nhận ngân sách & người quyết định -> Không đủ điều kiện vào Top 5 hợp lệ!
        {
            "opp_id": "OPP-B05",
            "customer_name": "Tập đoàn Bất động sản Hưng Thịnh",
            "project_name": "Hệ thống ERP Cloud SAP/Odoo B2B",
            "rep_id": "NV05",
            "stage": "Proposal",
            "deal_value": 3000000000,
            "sales_custom_prob": 0.95,
            "created_date": "2026-10-15",
            "updated_date": "2026-11-05",
            "expected_close_date": "2026-11-15",
            "decision_maker": "",
            "budget_confirmed": "Chưa xác nhận",
            "next_step": "",
            "source_file": "Boundary_Test.xlsx"
        },
        # Cơ hội #6: Đã có đơn hàng tại đúng ngày báo cáo (2026-11-08) -> Phải chuyển sang Doanh số đã ghi nhận!
        {
            "opp_id": "OPP-B06",
            "customer_name": "Công ty Chứng khoán Phú Hưng",
            "project_name": "Dịch vụ SOC & Bảo mật An toàn thông tin",
            "rep_id": "NV01",
            "stage": "Closing",
            "deal_value": 1100000000,
            "sales_custom_prob": 0.90,
            "created_date": "2026-09-01",
            "updated_date": "2026-11-08",
            "expected_close_date": "2026-11-08",
            "decision_maker": "Ông Trần Quốc Tuấn (CIO)",
            "budget_confirmed": "Đã xác nhận",
            "next_step": "Đã chốt đơn ngày 08/11",
            "source_file": "Boundary_Test.xlsx"
        }
    ]

    boundary_acts = [
        {"activity_id": "ACT-B01", "opp_id": "OPP-B01", "activity_date": "2026-10-25", "activity_type": "Meeting", "outcome_signal": "POSITIVE", "notes": "Tương tác cách đúng 14 ngày so với ngày báo cáo 08/11 (ngay tại ngưỡng biên 14 ngày)."},
        {"activity_id": "ACT-B02", "opp_id": "OPP-B02", "activity_date": "2026-11-06", "activity_type": "Call", "outcome_signal": "POSITIVE", "notes": "Khách đồng ý dự thảo hợp đồng."},
        {"activity_id": "ACT-B03", "opp_id": "OPP-B03", "activity_date": "2026-11-07", "activity_type": "Email", "outcome_signal": "POSITIVE", "notes": "Xác nhận lịch ký hợp đồng vào ngày 10/11."},
        # Hoạt động phát sinh SAU ngày báo cáo (2026-11-09 > 2026-11-08) -> Hệ thống phải lọc bỏ khi đối chiếu tại mốc 08/11
        {"activity_id": "ACT-B03-FUTURE", "opp_id": "OPP-B03", "activity_date": "2026-11-09", "activity_type": "Call", "outcome_signal": "LOST", "notes": "Bản ghi tương lai sau ngày đối chiếu 08/11 (Dùng để kiểm thử bộ lọc Cut-off Date)."},
        {"activity_id": "ACT-B04", "opp_id": "OPP-B04", "activity_date": "2026-10-24", "activity_type": "Call", "outcome_signal": "STALLED", "notes": "Tương tác cách 15 ngày (> ngưỡng 14 ngày) và ngày dự kiến chốt 07/11 (< 08/11)."},
        {"activity_id": "ACT-B05", "opp_id": "OPP-B05", "activity_date": "2026-11-05", "activity_type": "Email", "outcome_signal": "NEUTRAL", "notes": "Chưa có thông tin người quyết định và ngân sách."},
        {"activity_id": "ACT-B06", "opp_id": "OPP-B06", "activity_date": "2026-11-08", "activity_type": "Contract", "outcome_signal": "WON", "notes": "Ký hợp đồng đúng vào ngày đối chiếu 08/11/2026."},
    ]

    boundary_orders = [
        # Đơn hàng đúng ngày đối chiếu 2026-11-08 -> Được ghi nhận vào Doanh số thực tế kỳ này
        {"order_id": "ORD-B01", "opp_id": "OPP-B06", "customer_name": "Công ty Chứng khoán Phú Hưng", "rep_id": "NV01", "order_amount": 1100000000, "order_date": "2026-11-08", "order_status": "Approved"},
        # Đơn hàng phát sinh SAU ngày đối chiếu (2026-11-10 > 2026-11-08) -> Chưa được tính vào Doanh số thực tế tại ngày 08/11!
        {"order_id": "ORD-B02-FUTURE", "opp_id": "OPP-B03", "customer_name": "Công ty CP Logistics Toàn Cầu", "rep_id": "NV03", "order_amount": 690000000, "order_date": "2026-11-10", "order_status": "Approved"},
    ]

    _save_bundle(
        os.path.join(base_dir, "scenario_3_boundary"),
        "BinhMinh_Scenario3_Boundary",
        pd.DataFrame(boundary_opps),
        pd.DataFrame(boundary_acts),
        pd.DataFrame(boundary_orders),
        kpi_df
    )


if __name__ == "__main__":
    sample_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data")
    generate_all_datasets(sample_dir)
    print(f"Sample datasets generated successfully in: {sample_dir}")

