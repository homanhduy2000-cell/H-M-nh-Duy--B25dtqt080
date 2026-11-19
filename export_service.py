# -*- coding: utf-8 -*-
"""
Dịch vụ xuất báo cáo đa định dạng cho Công ty của Duydudi (Đề 09 - AI301)
Hỗ trợ:
- Excel (.xlsx) đa sheet có định dạng số và màu sắc chuyên nghiệp
- Word (.docx) văn bản báo cáo điều hành hoàn chỉnh kèm bảng biểu và chữ ký
- PowerPoint (.pptx) slide trình chiếu 16:9 đẹp mắt, phong cách Shore theme
"""

import io
import datetime
from typing import Dict, Any, List

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import docx
    from docx.shared import Inches, Pt, RGBColor as DocxRGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
except Exception:
    docx = None

try:
    import pptx
    from pptx.util import Inches as PptxInches, Pt as PptxPt
    from pptx.dml.color import RGBColor as PptxRGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.enum.shapes import MSO_SHAPE
except Exception:
    pptx = None


def _fmt_vnd(num: Any) -> str:
    try:
        val = float(num)
        return f"{val:,.0f} đ"
    except Exception:
        return str(num or "0 đ")


def _fmt_pct(num: Any) -> str:
    try:
        val = float(num)
        return f"{val:.1f}%"
    except Exception:
        return str(num or "0.0%")


def _extract_top_opps(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    p_rec = data.get("priority_recommendations", {})
    if isinstance(p_rec, dict):
        return p_rec.get("top_opportunities", [])
    elif isinstance(p_rec, list):
        return p_rec
    return []


def _extract_brief_text(data: Dict[str, Any]) -> str:
    brief_data = data.get("ai_executive_brief", {})
    if isinstance(brief_data, dict):
        return brief_data.get("content", "")
    return str(brief_data or "")


# =============================================================================
# 1. EXCEL (.XLSX) EXPORT SERVICE
# =============================================================================

def generate_excel_report(data: Dict[str, Any]) -> bytes:
    output = io.BytesIO()
    report_date = data.get("report_date", datetime.date.today().strftime("%d/%m/%Y"))
    company_kpis = data.get("company_kpis", {})
    rep_summary = data.get("rep_summary", [])
    top_opps = _extract_top_opps(data)
    clean_opps = data.get("opportunities_clean", [])
    brief_text = _extract_brief_text(data)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default blank sheet

    # Fonts & Fills (Palette: #894b77 Plum, #1ed7b5 Mint, #f0f9a7 Citrus, #abbdff Periwinkle)
    font_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
    fill_title = PatternFill(start_color="894B77", end_color="894B77", fill_type="solid")

    font_header = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="129C82", end_color="129C82", fill_type="solid")

    font_bold = Font(name="Segoe UI", size=11, bold=True)
    font_regular = Font(name="Segoe UI", size=11)

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    fill_alt = PatternFill(start_color="F3F6FF", end_color="F3F6FF", fill_type="solid")

    # -------------------------------------------------------------
    # SHEET 1: TỔNG QUAN KPI
    # -------------------------------------------------------------
    ws1 = wb.create_sheet(title="1_Tong_Quan_KPI")
    ws1.views.sheetView[0].showGridLines = True

    ws1.merge_cells("A1:D1")
    t1 = ws1["A1"]
    t1.value = "CÔNG TY CỦA DUYDUDI - BÁO CÁO TỔNG QUAN TIẾN ĐỘ & DỰ KIẾN DOANH SỐ"
    t1.font = font_title
    t1.fill = fill_title
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 40

    ws1["A2"] = f"Ngày báo cáo: {report_date}"
    ws1["A2"].font = Font(name="Segoe UI", size=11, italic=True)
    ws1["A3"] = "Hệ thống: Sales Operations AI Agent (AI301)"
    ws1["A3"].font = Font(name="Segoe UI", size=11, italic=True)

    headers_kpi = ["Chỉ Số Doanh Nghiệp", "Giá Trị Số Liệu", "Đơn Vị", "Ghi Chú Đánh Giá"]
    for col_idx, h in enumerate(headers_kpi, start=1):
        c = ws1.cell(row=5, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
    ws1.row_dimensions[5].height = 28

    tgt = company_kpis.get("total_kpi_target") or company_kpis.get("target_revenue", 0)
    act = company_kpis.get("total_recorded_revenue") or company_kpis.get("actual_revenue", 0)
    wf = company_kpis.get("total_weighted_forecast_estimate") or company_kpis.get("weighted_forecast", 0)
    tot = company_kpis.get("total_projected_revenue_estimate") or company_kpis.get("expected_total", 0)
    rate = company_kpis.get("company_projected_attainment_pct") or company_kpis.get("expected_achievement_rate", 0)
    inflation = company_kpis.get("forecast_inflation_removed", 0)
    open_count = company_kpis.get("total_open_opportunities") or company_kpis.get("total_active_opps", 0)
    closed_won_count = company_kpis.get("total_closed_won", 0)

    gap = max(0, tgt - tot)

    kpi_rows = [
        ("Chỉ tiêu Doanh số Toàn công ty (Target)", tgt, "VNĐ", "Mục tiêu ban đầu"),
        ("Doanh số Thực tế Đã chốt (Closed Won)", act, "VNĐ", f"Đã ký kết ({closed_won_count} hợp đồng)"),
        ("Doanh số Dự kiến có trọng số (Forecast Weighted)", wf, "VNĐ", "Tính theo trọng số từng giai đoạn"),
        ("Tổng Doanh số Kỳ vọng (Expected Total)", tot, "VNĐ", "Đã chốt + Dự kiến trọng số"),
        ("Tỷ lệ Hoàn thành Kỳ vọng", (rate or 0) / 100.0, "%", "Đạt chỉ tiêu" if rate >= 100 else "Cần đôn đốc thêm"),
        ("Khoảng cách cần bù đắp (Gap)", gap, "VNĐ", "Số tiền còn thiếu để đạt 100% KPI"),
        ("Ảo số dự báo được loại bỏ (Inflation Removed)", inflation, "VNĐ", "Lọc deal trùng lặp & quá hạn"),
        ("Tổng số Cơ hội đang mở", open_count, "Cơ hội", "Các deal đang tiếp tục bám sát"),
    ]

    for idx, (metric, val, unit, note) in enumerate(kpi_rows, start=6):
        r_c1 = ws1.cell(row=idx, column=1, value=metric)
        r_c2 = ws1.cell(row=idx, column=2, value=val)
        r_c3 = ws1.cell(row=idx, column=3, value=unit)
        r_c4 = ws1.cell(row=idx, column=4, value=note)

        for col in (r_c1, r_c2, r_c3, r_c4):
            col.font = font_regular
            col.border = cell_border
            if idx % 2 == 1:
                col.fill = fill_alt

        if unit == "VNĐ":
            r_c2.number_format = '#,##0 "đ"'
            r_c2.alignment = Alignment(horizontal="right")
        elif unit == "%":
            r_c2.number_format = "0.0%"
            r_c2.alignment = Alignment(horizontal="right")
            r_c2.font = font_bold
        else:
            r_c2.number_format = '#,##0'
            r_c2.alignment = Alignment(horizontal="center")

        ws1.row_dimensions[idx].height = 24

    # Brief Section
    b_row = len(kpi_rows) + 8
    ws1.cell(row=b_row, column=1, value="NHẬN XÉT ĐIỀU HÀNH TỔNG QUAN TỪ AI:").font = font_bold
    ws1.merge_cells(start_row=b_row + 1, start_column=1, end_row=b_row + 6, end_column=4)
    b_cell = ws1.cell(row=b_row + 1, column=1, value=brief_text)
    b_cell.font = Font(name="Segoe UI", size=10, italic=True)
    b_cell.alignment = Alignment(vertical="top", wrap_text=True)

    # -------------------------------------------------------------
    # SHEET 2: TOP 5 ƯU TIÊN
    # -------------------------------------------------------------
    ws2 = wb.create_sheet(title="2_Top_5_Uu_Tien")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:H1")
    t2 = ws2["A1"]
    t2.value = "TOP 5 CƠ HỘI TRỌNG ĐIỂM CẦN ĐÔN ĐỐC CHỐT HỢP ĐỒNG TUẦN TỚI"
    t2.font = font_title
    t2.fill = fill_title
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 36

    headers_top5 = [
        "Hạng", "Mã Opp", "Tên Khách Hàng", "NVKD Phụ Trách",
        "Giá Trị Deal (VNĐ)", "Giai Đoạn", "Điểm Ưu Tiên", "Hành Động Đề Xuất & Lưu Ý Chốt Sale"
    ]
    for col_idx, h in enumerate(headers_top5, start=1):
        c = ws2.cell(row=3, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
    ws2.row_dimensions[3].height = 28

    for idx, opp in enumerate(top_opps[:5], start=4):
        rank = opp.get("rank", idx - 3)
        opp_id = opp.get("opp_id", "")
        cust_name = opp.get("customer_name", "")
        rep = opp.get("rep_name") or opp.get("sales_rep", "")
        deal_val = opp.get("deal_value", 0)
        stage = opp.get("stage_label") or opp.get("effective_stage", "")
        score = opp.get("priority_score", 0)
        action = opp.get("recommended_action_next_week") or opp.get("recommended_action", "")
        support = opp.get("support_assignment") or opp.get("talking_points", "")
        full_action = f"{action} [Phân công hỗ trợ: {support}]" if support else action

        row_vals = [rank, opp_id, cust_name, rep, deal_val, stage, score, full_action]
        for col_idx, cell_val in enumerate(row_vals, start=1):
            cell = ws2.cell(row=idx, column=col_idx, value=cell_val)
            cell.font = font_regular
            cell.border = cell_border
            if idx % 2 == 1:
                cell.fill = fill_alt
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center")
                cell.font = font_bold
            elif col_idx == 5:
                cell.number_format = '#,##0 "đ"'
                cell.alignment = Alignment(horizontal="right")
                cell.font = font_bold
            elif col_idx == 7:
                cell.number_format = "0.0"
                cell.alignment = Alignment(horizontal="center")
            elif col_idx == 8:
                cell.alignment = Alignment(wrap_text=True)

        ws2.row_dimensions[idx].height = 46

    # -------------------------------------------------------------
    # SHEET 3: TIẾN ĐỘ 10 SALES
    # -------------------------------------------------------------
    ws3 = wb.create_sheet(title="3_KPI_10_Sales")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:G1")
    t3 = ws3["A1"]
    t3.value = "BẢNG ĐÁNH GIÁ TIẾN ĐỘ & HIỆU SUẤT ĐỘI NGŨ 10 NHÂN SỰ KINH DOANH"
    t3.font = font_title
    t3.fill = fill_title
    t3.alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 36

    headers_sales = [
        "Nhân Viên Sales", "Mục Tiêu (VNĐ)", "Đã Chốt (VNĐ)",
        "Dự Kiến Thêm (VNĐ)", "Tổng Kỳ Vọng (VNĐ)", "Tỷ Lệ Đạt (%)", "Đánh Giá Trạng Thái"
    ]
    for col_idx, h in enumerate(headers_sales, start=1):
        c = ws3.cell(row=3, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
    ws3.row_dimensions[3].height = 28

    for idx, rep in enumerate(rep_summary, start=4):
        name = rep.get("rep_name") or rep.get("sales_rep", "")
        r_tgt = rep.get("kpi_target") or rep.get("target", 0)
        r_act = rep.get("recorded_revenue") or rep.get("actual", 0)
        r_wf = rep.get("weighted_forecast_value") or rep.get("weighted_forecast", 0)
        r_tot = rep.get("projected_total_estimate") or rep.get("expected_total", 0)
        r_rate = (rep.get("projected_attainment_pct") or rep.get("expected_rate", 0) or 0) / 100.0
        r_status = rep.get("kpi_status") or rep.get("status", "")

        row_vals = [name, r_tgt, r_act, r_wf, r_tot, r_rate, r_status]
        for col_idx, cell_val in enumerate(row_vals, start=1):
            cell = ws3.cell(row=idx, column=col_idx, value=cell_val)
            cell.font = font_regular
            cell.border = cell_border
            if idx % 2 == 1:
                cell.fill = fill_alt

            if col_idx in (2, 3, 4, 5):
                cell.number_format = '#,##0 "đ"'
                cell.alignment = Alignment(horizontal="right")
            elif col_idx == 6:
                cell.number_format = "0.0%"
                cell.alignment = Alignment(horizontal="center")
                cell.font = font_bold
                if r_rate >= 1.0:
                    cell.fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
                elif r_rate < 0.7:
                    cell.fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
            elif col_idx == 7:
                cell.alignment = Alignment(horizontal="center")

        ws3.row_dimensions[idx].height = 26

    # -------------------------------------------------------------
    # SHEET 4: TOÀN BỘ CƠ HỘI SẠCH
    # -------------------------------------------------------------
    ws4 = wb.create_sheet(title="4_Co_Hoi_Clean")
    ws4.views.sheetView[0].showGridLines = True

    headers_clean = [
        "Mã Opp", "Tên Khách Hàng", "NVKD", "Giai Đoạn",
        "Trọng Số (%)", "Giá Trị Deal (VNĐ)", "Ngày Dự Kiến Chốt", "Dự Kiến Trọng Số (VNĐ)", "Tình Trạng Rủi Ro"
    ]
    for col_idx, h in enumerate(headers_clean, start=1):
        c = ws4.cell(row=1, column=col_idx, value=h)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
    ws4.row_dimensions[1].height = 28

    for idx, opp in enumerate(clean_opps, start=2):
        weight = opp.get("rule_weight") or (opp.get("stage_weight", 0) / 100.0 if opp.get("stage_weight", 0) > 1 else opp.get("stage_weight", 0))
        wf_val = opp.get("weighted_forecast_value") or opp.get("weighted_value", 0)
        has_anom = opp.get("has_anomaly", False)

        row_vals = [
            opp.get("opp_id", ""),
            opp.get("customer_name", ""),
            opp.get("rep_name") or opp.get("rep_id") or opp.get("sales_rep", ""),
            opp.get("stage_label") or opp.get("effective_stage", ""),
            weight,
            opp.get("deal_value", 0),
            str(opp.get("expected_close_date", "")),
            wf_val,
            "Cần lưu ý" if has_anom else "Hợp lệ"
        ]
        for col_idx, cell_val in enumerate(row_vals, start=1):
            cell = ws4.cell(row=idx, column=col_idx, value=cell_val)
            cell.font = font_regular
            cell.border = cell_border
            if idx % 2 == 1:
                cell.fill = fill_alt

            if col_idx == 5:
                cell.number_format = "0%"
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in (6, 8):
                cell.number_format = '#,##0 "đ"'
                cell.alignment = Alignment(horizontal="right")
            elif col_idx in (1, 7):
                cell.alignment = Alignment(horizontal="center")

        ws4.row_dimensions[idx].height = 22

    # Auto-fit column widths
    for ws in wb.worksheets:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if cell.row == 1 and ws.title != "4_Co_Hoi_Clean":
                    continue
                max_len = max(max_len, len(val_str))
            ws.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 52)

    wb.save(output)
    output.seek(0)
    return output.getvalue()


# =============================================================================
# 2. WORD (.DOCX) EXPORT SERVICE
# =============================================================================

def generate_word_report(data: Dict[str, Any]) -> bytes:
    doc = docx.Document()

    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)

    report_date = data.get("report_date", datetime.date.today().strftime("%d/%m/%Y"))
    company_kpis = data.get("company_kpis", {})
    rep_summary = data.get("rep_summary", [])
    top_opps = _extract_top_opps(data)
    brief_text = _extract_brief_text(data)

    tgt = company_kpis.get("total_kpi_target") or company_kpis.get("target_revenue", 0)
    act = company_kpis.get("total_recorded_revenue") or company_kpis.get("actual_revenue", 0)
    wf = company_kpis.get("total_weighted_forecast_estimate") or company_kpis.get("weighted_forecast", 0)
    tot = company_kpis.get("total_projected_revenue_estimate") or company_kpis.get("expected_total", 0)
    rate = company_kpis.get("company_projected_attainment_pct") or company_kpis.get("expected_achievement_rate", 0)
    open_count = company_kpis.get("total_open_opportunities") or company_kpis.get("total_active_opps", 0)
    closed_won_count = company_kpis.get("total_closed_won", 0)
    inflation = company_kpis.get("forecast_inflation_removed", 0)
    gap = max(0, tgt - tot)

    # Document Header
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = title_p.add_run("BÁO CÁO PHÂN TÍCH TIẾN ĐỘ CƠ HỘI & DỰ KIẾN DOANH SỐ")
    r_title.bold = True
    r_title.font.size = Pt(17)
    r_title.font.name = "Segoe UI"
    r_title.font.color.rgb = DocxRGBColor(137, 75, 119)

    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_after = Pt(10)
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = sub_p.add_run("CÔNG TY CỦA DUYDUDI • SALES OPERATIONS AI AGENT (AI301)")
    r_sub.bold = True
    r_sub.font.size = Pt(11)
    r_sub.font.name = "Segoe UI"
    r_sub.font.color.rgb = DocxRGBColor(71, 85, 105)

    meta_p = doc.add_paragraph()
    meta_p.paragraph_format.space_after = Pt(14)
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = meta_p.add_run(f"Ngày lập báo cáo: {report_date}  |  Phạm vi: Toàn quốc  |  Trạng thái: Đã phê duyệt")
    r_meta.italic = True
    r_meta.font.size = Pt(10)
    r_meta.font.name = "Segoe UI"
    r_meta.font.color.rgb = DocxRGBColor(100, 116, 139)

    # -------------------------------------------------------------
    # PHẦN 1: TỔNG QUAN CHỈ SỐ DOANH NGHIỆP
    # -------------------------------------------------------------
    h1 = doc.add_heading("I. TỔNG QUAN CHỈ SỐ DOANH SỐ & KPI DOANH NGHIỆP", level=1)
    h1.paragraph_format.space_before = Pt(10)
    h1.paragraph_format.space_after = Pt(6)

    table_kpi = doc.add_table(rows=1, cols=4)
    table_kpi.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_kpi.autofit = False

    col_widths = [Inches(2.5), Inches(1.8), Inches(1.1), Inches(1.8)]
    hdr_cells = table_kpi.rows[0].cells
    hdr_titles = ["Chỉ Số Doanh Nghiệp", "Giá Trị Số Liệu", "Đơn Vị", "Đánh Giá"]
    for i, t in enumerate(hdr_titles):
        hdr_cells[i].text = t
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = DocxRGBColor(255, 255, 255)
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:color'), 'auto')
        shading.set(qn('w:fill'), '0B8199')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shading)

    kpi_items = [
        ("Mục tiêu Doanh số Toàn công ty", _fmt_vnd(tgt), "VNĐ", "Chỉ tiêu Tổng thể"),
        ("Doanh số Đã chốt (Closed Won)", _fmt_vnd(act), "VNĐ", f"Đã hoàn thành ({closed_won_count} hợp đồng)"),
        ("Dự kiến có trọng số (Forecast Weighted)", _fmt_vnd(wf), "VNĐ", "Tính theo độ chín giai đoạn"),
        ("Tổng Doanh số Kỳ vọng", _fmt_vnd(tot), "VNĐ", "Đã chốt + Dự kiến"),
        ("Tỷ lệ Hoàn thành Kỳ vọng", _fmt_pct(rate), "%", "Đạt chỉ tiêu" if rate >= 100 else "Cần đôn đốc thêm"),
        ("Khoảng cách cần bù đắp (Gap)", _fmt_vnd(gap), "VNĐ", "Cần nỗ lực bổ sung"),
        ("Ảo số dự báo được loại bỏ", _fmt_vnd(inflation), "VNĐ", "Lọc deal trùng & quá hạn"),
        ("Tổng số Cơ hội đang mở", str(open_count), "Cơ hội", "Pipeline đang hoạt động")
    ]

    for item in kpi_items:
        row_cells = table_kpi.add_row().cells
        for col_idx, text_val in enumerate(item):
            row_cells[col_idx].text = str(text_val)
            p = row_cells[col_idx].paragraphs[0]
            p.runs[0].font.size = Pt(9.5)
            p.runs[0].font.name = "Segoe UI"
            if col_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.runs[0].font.bold = True
            elif col_idx == 2:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for row in table_kpi.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    # Brief Callout
    doc.add_paragraph().paragraph_format.space_after = Pt(4)
    brief_head = doc.add_paragraph()
    r_bh = brief_head.add_run("Nhận Xét Điều Hành Chiến Lược Từ AI Agent:")
    r_bh.bold = True
    r_bh.font.color.rgb = DocxRGBColor(11, 129, 153)

    bp = doc.add_paragraph()
    bp.paragraph_format.left_indent = Inches(0.2)
    bp.paragraph_format.space_after = Pt(12)
    r_bp = bp.add_run(brief_text)
    r_bp.italic = True
    r_bp.font.size = Pt(9.5)

    # -------------------------------------------------------------
    # PHẦN 2: TOP 5 CƠ HỘI TRỌNG ĐIỂM
    # -------------------------------------------------------------
    h2 = doc.add_heading("II. TOP 5 CƠ HỘI TRỌNG ĐIỂM CẦN ĐÔN ĐỐC CHỐT HỢP ĐỒNG", level=1)
    h2.paragraph_format.space_before = Pt(10)
    h2.paragraph_format.space_after = Pt(6)

    table_top = doc.add_table(rows=1, cols=6)
    table_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_widths = [Inches(0.6), Inches(1.8), Inches(1.3), Inches(1.3), Inches(0.8), Inches(1.4)]
    hdr_cells2 = table_top.rows[0].cells
    hdr_titles2 = ["Hạng", "Khách Hàng", "Phụ Trách", "Giá Trị Deal", "Giai Đoạn", "Hành Động Đề Xuất"]
    for i, t in enumerate(hdr_titles2):
        hdr_cells2[i].text = t
        hdr_cells2[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells2[i].paragraphs[0].runs[0].font.color.rgb = DocxRGBColor(255, 255, 255)
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:color'), 'auto')
        shading.set(qn('w:fill'), '164E63')
        hdr_cells2[i]._tc.get_or_add_tcPr().append(shading)

    for idx, opp in enumerate(top_opps[:5], start=1):
        row_cells = table_top.add_row().cells
        rank = str(opp.get("rank", idx))
        cust = str(opp.get("customer_name", ""))
        rep = str(opp.get("rep_name") or opp.get("sales_rep", ""))
        val = _fmt_vnd(opp.get("deal_value", 0))
        stg = str(opp.get("stage_label") or opp.get("effective_stage", ""))
        act_text = str(opp.get("recommended_action_next_week") or opp.get("recommended_action", ""))

        vals = [rank, cust, rep, val, stg, act_text]
        for col_idx, text_val in enumerate(vals):
            row_cells[col_idx].text = text_val
            p = row_cells[col_idx].paragraphs[0]
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.name = "Segoe UI"
            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.bold = True
            elif col_idx == 3:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.runs[0].font.bold = True

    for row in table_top.rows:
        for i, w in enumerate(t_widths):
            row.cells[i].width = w

    # -------------------------------------------------------------
    # PHẦN 3: BẢNG TIẾN ĐỘ 10 SALES
    # -------------------------------------------------------------
    h3 = doc.add_heading("III. ĐÁNH GIÁ TIẾN ĐỘ ĐỘI NGŨ 10 NHÂN SỰ KINH DOANH", level=1)
    h3.paragraph_format.space_before = Pt(12)
    h3.paragraph_format.space_after = Pt(6)

    table_sales = doc.add_table(rows=1, cols=6)
    table_sales.alignment = WD_TABLE_ALIGNMENT.CENTER
    s_widths = [Inches(1.6), Inches(1.3), Inches(1.3), Inches(1.3), Inches(0.8), Inches(1.3)]
    hdr_cells3 = table_sales.rows[0].cells
    hdr_titles3 = ["Nhân Viên Sales", "Mục Tiêu", "Đã Chốt", "Tổng Kỳ Vọng", "Tỷ Lệ %", "Đánh Giá"]
    for i, t in enumerate(hdr_titles3):
        hdr_cells3[i].text = t
        hdr_cells3[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells3[i].paragraphs[0].runs[0].font.color.rgb = DocxRGBColor(255, 255, 255)
        shading = OxmlElement('w:shd')
        shading.set(qn('w:val'), 'clear')
        shading.set(qn('w:color'), 'auto')
        shading.set(qn('w:fill'), '0B8199')
        hdr_cells3[i]._tc.get_or_add_tcPr().append(shading)

    for rep in rep_summary:
        row_cells = table_sales.add_row().cells
        name = rep.get("rep_name") or rep.get("sales_rep", "")
        r_tgt = _fmt_vnd(rep.get("kpi_target") or rep.get("target", 0))
        r_act = _fmt_vnd(rep.get("recorded_revenue") or rep.get("actual", 0))
        r_tot = _fmt_vnd(rep.get("projected_total_estimate") or rep.get("expected_total", 0))
        r_rate = _fmt_pct(rep.get("projected_attainment_pct") or rep.get("expected_rate", 0))
        r_status = str(rep.get("kpi_status") or rep.get("status", ""))

        vals = [name, r_tgt, r_act, r_tot, r_rate, r_status]
        for col_idx, text_val in enumerate(vals):
            row_cells[col_idx].text = text_val
            p = row_cells[col_idx].paragraphs[0]
            p.runs[0].font.size = Pt(9)
            p.runs[0].font.name = "Segoe UI"
            if col_idx in (1, 2, 3):
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif col_idx == 4:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.bold = True

    for row in table_sales.rows:
        for i, w in enumerate(s_widths):
            row.cells[i].width = w

    # Signatures
    doc.add_paragraph().paragraph_format.space_before = Pt(16)
    sig_p = doc.add_paragraph()
    sig_p.paragraph_format.keep_with_next = True
    table_sig = doc.add_table(rows=2, cols=2)
    table_sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_sig.rows[0].cells[0].paragraphs[0].text = "TRƯỞNG PHÒNG KINH DOANH\n(Ký và ghi rõ họ tên)"
    table_sig.rows[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_sig.rows[0].cells[1].paragraphs[0].text = "HỆ THỐNG DUYDUDI AI AGENT\n(Xác thực dữ liệu tự động)"
    table_sig.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_sig.rows[1].cells[0].paragraphs[0].text = "\n\n_________________________\nHoàng Mạnh Duy"
    table_sig.rows[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    table_sig.rows[1].cells[1].paragraphs[0].text = f"\n\n_________________________\nNgày phê duyệt: {report_date}"
    table_sig.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out.getvalue()


# =============================================================================
# 3. POWERPOINT (.PPTX) EXPORT SERVICE
# =============================================================================

def generate_powerpoint_report(data: Dict[str, Any]) -> bytes:
    prs = pptx.Presentation()
    prs.slide_width = PptxInches(13.333)
    prs.slide_height = PptxInches(7.5)

    report_date = data.get("report_date", datetime.date.today().strftime("%d/%m/%Y"))
    company_kpis = data.get("company_kpis", {})
    rep_summary = data.get("rep_summary", [])
    top_opps = _extract_top_opps(data)
    brief_text = _extract_brief_text(data)

    tgt = company_kpis.get("total_kpi_target") or company_kpis.get("target_revenue", 0)
    act = company_kpis.get("total_recorded_revenue") or company_kpis.get("actual_revenue", 0)
    wf = company_kpis.get("total_weighted_forecast_estimate") or company_kpis.get("weighted_forecast", 0)
    tot = company_kpis.get("total_projected_revenue_estimate") or company_kpis.get("expected_total", 0)
    rate = company_kpis.get("company_projected_attainment_pct") or company_kpis.get("expected_achievement_rate", 0)

    c_dark_bg = PptxRGBColor(18, 14, 28)      # Deep Plum Slate
    c_card_bg = PptxRGBColor(32, 25, 48)      # Plum Card
    c_card_border = PptxRGBColor(171, 189, 255)  # Periwinkle #abbdff
    c_cyan = PptxRGBColor(30, 215, 181)       # Mint #1ed7b5
    c_peach = PptxRGBColor(240, 249, 167)     # Citrus #f0f9a7
    c_teal = PptxRGBColor(137, 75, 119)       # Deep Plum #894b77
    c_white = PptxRGBColor(255, 255, 255)
    c_muted = PptxRGBColor(171, 189, 255)     # Periwinkle #abbdff
    c_green = PptxRGBColor(30, 215, 181)      # Mint #1ed7b5

    blank_layout = prs.slide_layouts[6]

    def _set_slide_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = c_dark_bg
        bg.line.fill.background()
        return bg

    # -------------------------------------------------------------
    # SLIDE 1: TRANG BÌA
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    _set_slide_bg(s1)

    top_bar = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptxInches(1), PptxInches(0.8), PptxInches(11.333), PptxInches(0.08))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = c_cyan
    top_bar.line.fill.background()

    tb1 = s1.shapes.add_textbox(PptxInches(1), PptxInches(1.2), PptxInches(11.333), PptxInches(3.5))
    tf1 = tb1.text_frame
    tf1.word_wrap = True

    p_tag = tf1.paragraphs[0]
    p_tag.text = "BÁO CÁO ĐIỀU HÀNH DOANH SỐ • SALES OPERATIONS AI AGENT (AI301)"
    p_tag.font.size = PptxPt(13)
    p_tag.font.bold = True
    p_tag.font.color.rgb = c_cyan

    p_title = tf1.add_paragraph()
    p_title.text = "Phân Tích Tiến Độ Cơ Hội &\nDự Báo Doanh Số Doanh Nghiệp"
    p_title.font.size = PptxPt(38)
    p_title.font.bold = True
    p_title.font.color.rgb = c_white
    p_title.space_before = PptxPt(12)

    p_sub = tf1.add_paragraph()
    p_sub.text = f"Đơn vị: Công ty của Duydudi   •   Ngày lập: {report_date}   •   Trạng thái: Đã phê duyệt"
    p_sub.font.size = PptxPt(16)
    p_sub.font.color.rgb = c_peach
    p_sub.space_before = PptxPt(16)

    # 3 Summary Cards
    kpi_cards = [
        ("MỤC TIÊU TOÀN CÔNG TY", _fmt_vnd(tgt), c_peach),
        ("DOANH SỐ ĐÃ CHỐT", _fmt_vnd(act), c_green),
        ("TỔNG KỲ VỌNG HOÀN THÀNH", f"{_fmt_vnd(tot)} ({_fmt_pct(rate)})", c_cyan),
    ]

    card_w = PptxInches(3.55)
    card_h = PptxInches(1.6)
    start_x = PptxInches(1)
    gap = PptxInches(0.34)
    start_y = PptxInches(5.0)

    for i, (title, val, color) in enumerate(kpi_cards):
        x = start_x + i * (card_w + gap)
        card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, start_y, card_w, card_h)
        card.fill.solid()
        card.fill.fore_color.rgb = c_card_bg
        card.line.color.rgb = c_card_border

        tf = card.text_frame
        tf.word_wrap = True
        p_c1 = tf.paragraphs[0]
        p_c1.text = title
        p_c1.font.size = PptxPt(11)
        p_c1.font.bold = True
        p_c1.font.color.rgb = c_muted

        p_c2 = tf.add_paragraph()
        p_c2.text = val
        p_c2.font.size = PptxPt(18)
        p_c2.font.bold = True
        p_c2.font.color.rgb = color
        p_c2.space_before = PptxPt(8)

    # -------------------------------------------------------------
    # SLIDE 2: TỔNG QUAN CHỈ SỐ DOANH SỐ & KPI
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    _set_slide_bg(s2)

    tb_h = s2.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(0.8))
    p_h = tb_h.text_frame.paragraphs[0]
    p_h.text = "1. TỔNG QUAN HIỆU SUẤT DOANH SỐ & TIẾN ĐỘ KPI"
    p_h.font.size = PptxPt(24)
    p_h.font.bold = True
    p_h.font.color.rgb = c_cyan

    kpi_4 = [
        ("Mục tiêu Doanh số", _fmt_vnd(tgt), "Chỉ tiêu cần đạt", c_peach),
        ("Thực tế Đã chốt", _fmt_vnd(act), "Hợp đồng đã ký", c_green),
        ("Dự kiến thêm (Weighted)", _fmt_vnd(wf), "Pipeline có trọng số", c_cyan),
        ("Tỷ lệ Hoàn thành", _fmt_pct(rate), "Đạt chỉ tiêu" if rate >= 100 else "Cần đôn đốc", c_peach),
    ]

    card_w4 = PptxInches(2.7)
    card_h4 = PptxInches(1.5)
    gap4 = PptxInches(0.3)
    y4 = PptxInches(1.6)

    for i, (t, v, sub, col) in enumerate(kpi_4):
        x = PptxInches(0.8) + i * (card_w4 + gap4)
        c = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y4, card_w4, card_h4)
        c.fill.solid()
        c.fill.fore_color.rgb = c_card_bg
        c.line.color.rgb = c_card_border

        tf = c.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = t
        p1.font.size = PptxPt(11)
        p1.font.color.rgb = c_muted

        p2 = tf.add_paragraph()
        p2.text = v
        p2.font.size = PptxPt(19)
        p2.font.bold = True
        p2.font.color.rgb = col
        p2.space_before = PptxPt(6)

        p3 = tf.add_paragraph()
        p3.text = sub
        p3.font.size = PptxPt(10)
        p3.font.color.rgb = c_muted
        p3.space_before = PptxPt(4)

    # Executive Brief Box
    brief_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptxInches(0.8), PptxInches(3.4), PptxInches(11.733), PptxInches(3.3))
    brief_box.fill.solid()
    brief_box.fill.fore_color.rgb = c_card_bg
    brief_box.line.color.rgb = c_card_border

    tf_bb = brief_box.text_frame
    tf_bb.word_wrap = True
    p_bb_title = tf_bb.paragraphs[0]
    p_bb_title.text = "NHẬN XÉT ĐIỀU HÀNH TỔNG QUAN TỪ AI SALES OPERATIONS:"
    p_bb_title.font.size = PptxPt(13)
    p_bb_title.font.bold = True
    p_bb_title.font.color.rgb = c_cyan

    brief_display = brief_text[:750] + ("..." if len(brief_text) > 750 else "")
    p_bb_body = tf_bb.add_paragraph()
    p_bb_body.text = brief_display
    p_bb_body.font.size = PptxPt(12)
    p_bb_body.font.color.rgb = c_white
    p_bb_body.space_before = PptxPt(8)

    # -------------------------------------------------------------
    # SLIDE 3: TOP 5 CƠ HỘI TRỌNG ĐIỂM
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    _set_slide_bg(s3)

    tb_h3 = s3.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(0.8))
    p_h3 = tb_h3.text_frame.paragraphs[0]
    p_h3.text = "2. TOP 5 CƠ HỘI TRỌNG ĐIỂM CẦN ĐÔN ĐỐC CHỐT HỢP ĐỒNG"
    p_h3.font.size = PptxPt(24)
    p_h3.font.bold = True
    p_h3.font.color.rgb = c_cyan

    t_shape3 = s3.shapes.add_table(6, 6, PptxInches(0.8), PptxInches(1.5), PptxInches(11.733), PptxInches(5.2))
    table3 = t_shape3.table
    col_w3 = [PptxInches(0.8), PptxInches(2.6), PptxInches(1.8), PptxInches(1.8), PptxInches(1.4), PptxInches(3.333)]
    for i, w in enumerate(col_w3):
        table3.columns[i].width = w

    headers_s3 = ["Hạng", "Tên Khách Hàng", "Phụ Trách", "Giá Trị Deal", "Giai Đoạn", "Hành Động Đề Xuất"]
    for i, h in enumerate(headers_s3):
        c = table3.cell(0, i)
        c.fill.solid()
        c.fill.fore_color.rgb = c_teal
        p = c.text_frame.paragraphs[0]
        p.text = h
        p.font.size = PptxPt(12)
        p.font.bold = True
        p.font.color.rgb = c_white
        p.alignment = PP_ALIGN.CENTER

    for idx, opp in enumerate(top_opps[:5], start=1):
        action_text = opp.get("recommended_action_next_week") or opp.get("recommended_action", "")
        if len(action_text) > 130:
            action_text = action_text[:127] + "..."

        vals = [
            str(opp.get("rank", idx)),
            str(opp.get("customer_name", "")),
            str(opp.get("rep_name") or opp.get("sales_rep", "")),
            _fmt_vnd(opp.get("deal_value", 0)),
            str(opp.get("stage_label") or opp.get("effective_stage", "")),
            action_text
        ]
        for c_idx, v in enumerate(vals):
            c = table3.cell(idx, c_idx)
            c.fill.solid()
            c.fill.fore_color.rgb = c_card_bg if idx % 2 == 1 else PptxRGBColor(20, 29, 44)
            p = c.text_frame.paragraphs[0]
            p.text = v
            p.font.size = PptxPt(11)
            p.font.color.rgb = c_white
            if c_idx == 0:
                p.alignment = PP_ALIGN.CENTER
                p.font.bold = True
                p.font.color.rgb = c_cyan
            elif c_idx == 3:
                p.alignment = PP_ALIGN.RIGHT
                p.font.bold = True
                p.font.color.rgb = c_peach

    # -------------------------------------------------------------
    # SLIDE 4: TIẾN ĐỘ ĐỘI NGŨ 10 SALES
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    _set_slide_bg(s4)

    tb_h4 = s4.shapes.add_textbox(PptxInches(0.8), PptxInches(0.6), PptxInches(11.7), PptxInches(0.8))
    p_h4 = tb_h4.text_frame.paragraphs[0]
    p_h4.text = "3. ĐÁNH GIÁ TIẾN ĐỘ 10 NHÂN SỰ KINH DOANH"
    p_h4.font.size = PptxPt(24)
    p_h4.font.bold = True
    p_h4.font.color.rgb = c_cyan

    num_reps = min(len(rep_summary), 10)
    t_shape4 = s4.shapes.add_table(num_reps + 1, 6, PptxInches(0.8), PptxInches(1.5), PptxInches(11.733), PptxInches(5.2))
    table4 = t_shape4.table
    col_w4 = [PptxInches(2.3), PptxInches(1.9), PptxInches(1.9), PptxInches(1.9), PptxInches(1.5), PptxInches(2.233)]
    for i, w in enumerate(col_w4):
        table4.columns[i].width = w

    headers_s4 = ["Nhân Viên Sales", "Mục Tiêu", "Đã Chốt", "Tổng Kỳ Vọng", "Tỷ Lệ %", "Đánh Giá"]
    for i, h in enumerate(headers_s4):
        c = table4.cell(0, i)
        c.fill.solid()
        c.fill.fore_color.rgb = c_teal
        p = c.text_frame.paragraphs[0]
        p.text = h
        p.font.size = PptxPt(11)
        p.font.bold = True
        p.font.color.rgb = c_white
        p.alignment = PP_ALIGN.CENTER

    for idx, rep in enumerate(rep_summary[:10], start=1):
        r_tgt = rep.get("kpi_target") or rep.get("target", 0)
        r_act = rep.get("recorded_revenue") or rep.get("actual", 0)
        r_tot = rep.get("projected_total_estimate") or rep.get("expected_total", 0)
        r_rate = rep.get("projected_attainment_pct") or rep.get("expected_rate", 0)
        r_status = rep.get("kpi_status") or rep.get("status", "")

        vals = [
            str(rep.get("rep_name") or rep.get("sales_rep", "")),
            _fmt_vnd(r_tgt),
            _fmt_vnd(r_act),
            _fmt_vnd(r_tot),
            _fmt_pct(r_rate),
            str(r_status)
        ]
        for c_idx, v in enumerate(vals):
            c = table4.cell(idx, c_idx)
            c.fill.solid()
            c.fill.fore_color.rgb = c_card_bg if idx % 2 == 1 else PptxRGBColor(20, 29, 44)
            p = c.text_frame.paragraphs[0]
            p.text = v
            p.font.size = PptxPt(10)
            p.font.color.rgb = c_white
            if c_idx in (1, 2, 3):
                p.alignment = PP_ALIGN.RIGHT
            elif c_idx == 4:
                p.alignment = PP_ALIGN.CENTER
                p.font.bold = True
                p.font.color.rgb = c_green if r_rate >= 100 else (c_peach if r_rate >= 70 else PptxRGBColor(248, 113, 113))

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.getvalue()
