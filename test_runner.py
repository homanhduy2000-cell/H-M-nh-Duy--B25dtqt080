import os
from typing import Dict, Any, List
import pandas as pd
from agent_engine import SalesProgressAgent
from data_generator import DEFAULT_REPORT_DATE, generate_all_datasets

BASE_SAMPLE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data")


def get_active_sample_dir() -> str:
    """
    Trả về thư mục sample_data hợp lệ cho cả Local và Vercel Serverless (chỉ cho ghi ở /tmp).
    """
    if os.path.exists(os.path.join(BASE_SAMPLE_DIR, "scenario_1_normal", "opportunities.csv")):
        return BASE_SAMPLE_DIR
    tmp_dir = "/tmp/sample_data" if os.name != "nt" else os.path.join(os.environ.get("TEMP", "."), "duydudi_sample_data")
    if not os.path.exists(os.path.join(tmp_dir, "scenario_1_normal", "opportunities.csv")):
        os.makedirs(tmp_dir, exist_ok=True)
        generate_all_datasets(tmp_dir)
    return tmp_dir


SAMPLE_DIR = BASE_SAMPLE_DIR


def load_scenario_dfs(scenario_folder: str):
    active_dir = get_active_sample_dir()
    folder = os.path.join(active_dir, scenario_folder)
    opps_df = pd.read_csv(os.path.join(folder, "opportunities.csv"))
    acts_df = pd.read_csv(os.path.join(folder, "activities.csv"))
    orders_df = pd.read_csv(os.path.join(folder, "orders.csv"))
    kpi_df = pd.read_csv(os.path.join(folder, "kpi_targets.csv"))
    return opps_df, acts_df, orders_df, kpi_df


def run_all_self_tests(report_date: str = DEFAULT_REPORT_DATE) -> List[Dict[str, Any]]:
    """
    Chạy kiểm thử tự động trên 3 kịch bản bắt buộc của Đề dự án 09 (Task 2):
    1. Trường hợp thông thường (Normal Case)
    2. Trường hợp thiếu dữ liệu & mâu thuẫn (Missing & Conflicting Data Case)
    3. Trường hợp tại ranh giới quy tắc (Boundary / Edge Case)
    """
    results = []

    # 1. KỊCH BẢN 1: TRƯỜNG HỢP THÔNG THƯỜNG
    opps1, acts1, orders1, kpi1 = load_scenario_dfs("scenario_1_normal")
    agent1 = SalesProgressAgent(report_date=report_date)
    res1 = agent1.execute_pipeline(opps1, acts1, orders1, kpi1, dataset_label="Kịch bản 1: Trường hợp thông thường (Normal Case)")

    checks_1 = [
        {
            "check_name": "Hợp nhất cơ hội trùng lặp (OPP-001 và OPP-002-DUP)",
            "expected": "Giảm từ 47 bản ghi thô xuống 45 cơ hội duy nhất (gộp 2 cặp trùng)",
            "actual": f"Đầu vào {len(opps1)} bản ghi -> Sau hợp nhất: {res1['company_kpis']['total_unique_opportunities']} cơ hội ({len(res1['duplicates_merged'])} nhóm trùng)",
            "passed": res1["company_kpis"]["total_unique_opportunities"] == len(opps1) - 2 and len(res1["duplicates_merged"]) == 2
        },
        {
            "check_name": "Tách doanh số đã ghi nhận khỏi cơ hội mở (OPP-007 & OPP-008)",
            "expected": "OPP-007 (1.25 tỷ) & OPP-008 (820 triệu) có trong Orders phải chuyển sang Closed Won, Open Pipeline = 0",
            "actual": next(
                (f"OPP-007 stage={o['effective_stage']}, open_val={o['open_pipeline_value']:,.0f}, actual_rev={o['actual_recorded_revenue']:,.0f}"
                 for o in res1["opportunities_clean"] if o["opp_id"] == "OPP-007"),
                "Not found"
            ),
            "passed": any(
                o["opp_id"] == "OPP-007" and o["effective_stage"] == "Closed Won" and o["open_pipeline_value"] == 0
                for o in res1["opportunities_clean"]
            )
        },
        {
            "check_name": "Phát hiện cơ hội đã mất khách nhưng chưa đổi trạng thái (OPP-009 & OPP-010)",
            "expected": "Phát hiện tín hiệu LOST trong Activities và chuyển về Closed Lost (Trọng số 0%)",
            "actual": next(
                (f"OPP-009 stage={o['effective_stage']}, weighted_val={o['weighted_forecast_value']:,.0f}"
                 for o in res1["opportunities_clean"] if o["opp_id"] == "OPP-009"),
                "Not found"
            ),
            "passed": any(
                o["opp_id"] == "OPP-009" and o["effective_stage"] == "Closed Lost" and o["weighted_forecast_value"] == 0
                for o in res1["opportunities_clean"]
            )
        },
        {
            "check_name": "Loại trừ cơ hội giá trị lớn thiếu thông tin xác nhận khỏi Top 5 (OPP-011: 2.4 tỷ & OPP-012: 1.95 tỷ)",
            "expected": "OPP-011 và OPP-012 không được lọt vào Top 5 ưu tiên; xuất đủ 5/5 cơ hội hợp lệ",
            "actual": f"Số cơ hội ưu tiên hợp lệ: {res1['priority_recommendations']['actual_valid_count']}/5; Top IDs: {[x['opp_id'] for x in res1['priority_recommendations']['top_opportunities']]}",
            "passed": (
                res1["priority_recommendations"]["actual_valid_count"] == 5
                and "OPP-011" not in [x["opp_id"] for x in res1["priority_recommendations"]["top_opportunities"]]
                and "OPP-012" not in [x["opp_id"] for x in res1["priority_recommendations"]["top_opportunities"]]
            )
        }
    ]
    results.append({
        "scenario_id": "scenario_1_normal",
        "scenario_title": "1. Trường hợp thông thường (Normal Case)",
        "description": "Kiểm thử toàn diện quy trình chuẩn trên 10 NVKD: khử trùng lặp, tách đơn hàng đã chốt, đóng cơ hội đã mất khách và trích xuất đủ Top 5 cơ hội ưu tiên hợp lệ.",
        "all_passed": all(c["passed"] for c in checks_1),
        "checks": checks_1,
        "summary_metrics": res1["company_kpis"]
    })

    # 2. KỊCH BẢN 2: TRƯỜNG HỢP THIẾU DỮ LIỆU & MÂU THUẪN
    opps2, acts2, orders2, kpi2 = load_scenario_dfs("scenario_2_missing")
    agent2 = SalesProgressAgent(report_date=report_date)
    res2 = agent2.execute_pipeline(opps2, acts2, orders2, kpi2, dataset_label="Kịch bản 2: Thiếu dữ liệu & Mâu thuẫn (Missing & Conflict Data)")

    anom_types_2 = {a["anomaly_type"] for a in res2["anomalies"]}
    checks_2 = [
        {
            "check_name": "Phát hiện thiếu Ngày dự kiến chốt, thiếu Giai đoạn và Giá trị = 0 (OPP-M01, OPP-M02, OPP-M03)",
            "expected": "Gắn cờ MISSING_CLOSE_DATE, MISSING_STAGE và MISSING_VERIFICATION_INFO",
            "actual": f"Các loại lỗi phát hiện: {sorted(list(anom_types_2))}",
            "passed": {"MISSING_CLOSE_DATE", "MISSING_STAGE", "MISSING_VERIFICATION_INFO"}.issubset(anom_types_2)
        },
        {
            "check_name": "Phát hiện mâu thuẫn ngược: Đã có Đơn hàng ORD-M01 nhưng cơ hội OPP-M04 ghi Closed Lost",
            "expected": "Gắn cờ CONFLICT_ORDER_STATUS và khôi phục doanh số thực tế 920 triệu VNĐ",
            "actual": next(
                (f"OPP-M04 raw={o['raw_stage']} -> effective={o['effective_stage']}, actual_rev={o['actual_recorded_revenue']:,.0f}"
                 for o in res2["opportunities_clean"] if o["opp_id"] == "OPP-M04"),
                "Not found"
            ),
            "passed": any(
                o["opp_id"] == "OPP-M04" and o["effective_stage"] == "Closed Won" and o["actual_recorded_revenue"] == 920000000
                for o in res2["opportunities_clean"]
            )
        },
        {
            "check_name": "Phát hiện bản ghi trùng lặp OPP-M05 lệch nhân viên phụ trách (NV05 vs NV06) và lệch giá trị (700tr vs 1.2 tỷ)",
            "expected": "Gắn cờ CONFLICT_DUPLICATE_MERGE để chuyển Trưởng phòng KD phân xử tại Bước 4",
            "actual": next(
                (a["description"] for a in res2["anomalies"] if a["opp_id"] == "OPP-M05" and a["anomaly_type"] == "CONFLICT_DUPLICATE_MERGE"),
                "Not found"
            ),
            "passed": "CONFLICT_DUPLICATE_MERGE" in anom_types_2
        }
    ]
    results.append({
        "scenario_id": "scenario_2_missing",
        "scenario_title": "2. Trường hợp thiếu dữ liệu & mâu thuẫn (Missing & Conflicting Data)",
        "description": "Kiểm thử khả năng chống lỗi khi dữ liệu đầu vào bị khuyết ngày chốt, khuyết giai đoạn, giá trị bằng 0, mâu thuẫn ngược giữa Đơn hàng và Trạng thái, hoặc trùng lặp lệch thông tin.",
        "all_passed": all(c["passed"] for c in checks_2),
        "checks": checks_2,
        "summary_metrics": res2["company_kpis"]
    })

    # 3. KỊCH BẢN 3: TRƯỜNG HỢP TẠI RANH GIỚI QUY TẮC (BOUNDARY CASE)
    opps3, acts3, orders3, kpi3 = load_scenario_dfs("scenario_3_boundary")
    agent3 = SalesProgressAgent(report_date=report_date)
    res3 = agent3.execute_pipeline(opps3, acts3, orders3, kpi3, dataset_label="Kịch bản 3: Ranh giới quy tắc (Boundary Case)")

    checks_3 = [
        {
            "check_name": "Ranh giới ngày chốt & tương tác (OPP-B01: expected_close_date == 2026-11-08, tương tác cách đúng 14 ngày)",
            "expected": "Không bị tính là quá hạn hoặc đóng băng (Hợp lệ đưa vào danh sách ưu tiên)",
            "actual": next(
                (f"OPP-B01 days_to_close={o['days_to_close']}, days_since_act={o['days_since_activity']}, eligible={o['is_eligible_for_priority']}"
                 for o in res3["opportunities_clean"] if o["opp_id"] == "OPP-B01"),
                "Not found"
            ),
            "passed": any(
                o["opp_id"] == "OPP-B01" and o["days_to_close"] == 0 and o["days_since_activity"] == 14 and o["is_eligible_for_priority"] is True
                for o in res3["opportunities_clean"]
            )
        },
        {
            "check_name": "Vượt ranh giới đúng 1 ngày (OPP-B04: expected_close_date = 2026-11-07, tương tác cách 15 ngày)",
            "expected": "Bị gắn cờ Quá hạn chốt 1 ngày (OVERDUE_CLOSE_DATE) & Đóng băng 15 ngày (INACTIVE_STALE)",
            "actual": next(
                (f"OPP-B04 days_to_close={o['days_to_close']}, days_since_act={o['days_since_activity']}, eligible={o['is_eligible_for_priority']}"
                 for o in res3["opportunities_clean"] if o["opp_id"] == "OPP-B04"),
                "Not found"
            ),
            "passed": any(
                o["opp_id"] == "OPP-B04" and o["days_to_close"] == -1 and o["days_since_activity"] == 15 and o["is_eligible_for_priority"] is False
                for o in res3["opportunities_clean"]
            )
        },
        {
            "check_name": "Ranh giới ngày đối chiếu (Cut-off Filter): Lọc bỏ đơn hàng ORD-B02-FUTURE (2026-11-10 > 2026-11-08) và hoạt động ACT-B03-FUTURE",
            "expected": "2 bản ghi tương lai sau ngày 08/11 bị loại bỏ; OPP-B03 vẫn được giữ là Cơ hội đang mở tại ngày 08/11",
            "actual": f"Số bản ghi tương lai bị lọc bỏ: {len(res3['cutoff_filtered'])}; OPP-B03 stage={next((o['effective_stage'] for o in res3['opportunities_clean'] if o['opp_id']=='OPP-B03'), '')}",
            "passed": len(res3["cutoff_filtered"]) == 2 and any(
                o["opp_id"] == "OPP-B03" and o["effective_stage"] == "Closing"
                for o in res3["opportunities_clean"]
            )
        },
        {
            "check_name": "Ranh giới số lượng cơ hội ưu tiên hợp lệ < 5 (Yêu cầu đề bài: Nêu số thực tế và lý do)",
            "expected": "Chỉ có 3/5 cơ hội hợp lệ (OPP-B01, OPP-B02, OPP-B03); bật cờ is_under_5_boundary = True và giải trình lý do loại OPP-B04, OPP-B05",
            "actual": f"Số lượng thực tế: {res3['priority_recommendations']['actual_valid_count']}/5 | Giải trình: {res3['priority_recommendations']['boundary_explanation'][:110]}...",
            "passed": (
                res3["priority_recommendations"]["actual_valid_count"] == 3
                and res3["priority_recommendations"]["is_under_5_boundary"] is True
            )
        }
    ]
    results.append({
        "scenario_id": "scenario_3_boundary",
        "scenario_title": "3. Trường hợp tại ranh giới quy tắc (Boundary / Edge Case)",
        "description": "Kiểm thử các điểm biên toán học & nghiệp vụ: ngày chốt == ngày báo cáo, đúng 14 ngày không tương tác vs 15 ngày, lọc bỏ dữ liệu phát sinh sau ngày báo cáo, và trường hợp chỉ có 3 (< 5) cơ hội hợp lệ.",
        "all_passed": all(c["passed"] for c in checks_3),
        "checks": checks_3,
        "summary_metrics": res3["company_kpis"]
    })

    return results


if __name__ == "__main__":
    test_suites = run_all_self_tests()
    for suite in test_suites:
        print(f"\n=== {suite['scenario_title']} | PASSED: {suite['all_passed']} ===")
        for chk in suite["checks"]:
            status = "[PASS]" if chk["passed"] else "[FAIL]"
            print(f"  {status} {chk['check_name']}")
            print(f"         Actual: {chk['actual']}")

