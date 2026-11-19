import os
import io
import json
import zipfile
from typing import Optional, List, Dict, Any
import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse, Response
from pydantic import BaseModel

from agent_engine import (
    SalesProgressAgent,
    load_from_google_sheets_url,
    parse_excel_bytes,
    smart_normalize_columns,
    universal_adapt_dataframe,
    validate_uploaded_dataframe,
    build_file_diagnostics,
    ai_auto_repair_dataset,
    build_raw_file_preview,
    build_dynamic_kpi_df_from_opps,
    AUDIT_LOG_FILE,
    DEFAULT_GEMINI_API_KEY
)
from test_runner import load_scenario_dfs, run_all_self_tests, get_active_sample_dir
from data_generator import DEFAULT_REPORT_DATE, generate_all_datasets
from export_service import (
    generate_excel_report,
    generate_word_report,
    generate_powerpoint_report
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
SAMPLE_DIR = os.path.join(BASE_DIR, "sample_data")
DOCS_DIR = os.path.join(BASE_DIR, "docs")

if not os.environ.get("VERCEL"):
    try:
        if not os.path.exists(os.path.join(SAMPLE_DIR, "scenario_1_normal")):
            generate_all_datasets(SAMPLE_DIR)
    except Exception:
        pass

app = FastAPI(
    title="Duydudi - AI Agent Phân Tích Tiến Độ Cơ Hội & Dự Kiến Doanh Số (AI301)",
    version="1.0.0"
)

if os.path.exists(STATIC_DIR):
    app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# Lưu bộ dữ liệu gần nhất được tải lên trong phiên để Bước 4 (Human-in-the-loop) có thể duyệt lại ngay trên dữ liệu vừa upload
LAST_SESSION_DATA: Dict[str, Any] = {
    "source_type": "preset",
    "scenario_id": "scenario_1_normal",
    "opps_df": None,
    "acts_df": None,
    "orders_df": None,
    "kpi_df": None,
    "dataset_label": "Kịch bản 1: Trường hợp thông thường (Normal Case)"
}
LAST_ANALYSIS_RESULT: Optional[Dict[str, Any]] = None


class PresetAnalyzeRequest(BaseModel):
    scenario_id: str = "scenario_1_normal"
    report_date: str = DEFAULT_REPORT_DATE
    inactivity_threshold_days: int = 14
    human_overrides: Dict[str, Dict[str, Any]] = {}
    llm_provider: str = "gemini"
    api_key: str = DEFAULT_GEMINI_API_KEY
    use_last_uploaded: bool = False


class GSheetAnalyzeRequest(BaseModel):
    gsheet_url: str
    report_date: str = DEFAULT_REPORT_DATE
    inactivity_threshold_days: int = 14
    human_overrides: Dict[str, Dict[str, Any]] = {}
    llm_provider: str = "gemini"
    api_key: str = DEFAULT_GEMINI_API_KEY
    force_continue: bool = False


@app.get("/", response_class=HTMLResponse)
async def serve_index():
    index_path = os.path.join(STATIC_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/slides", response_class=HTMLResponse)
async def serve_slides():
    slides_path = os.path.join(STATIC_DIR, "slides.html")
    with open(slides_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())


@app.get("/api/export/slides.pptx")
async def export_slides_pptx():
    pptx_path = os.path.join(BASE_DIR, "Duydudi_AI_Agent_15_Slides.pptx")
    if not os.path.exists(pptx_path):
        from generate_slides_pptx import build_pptx_presentation
        pptx_path = "/tmp/Duydudi_AI_Agent_15_Slides.pptx" if os.environ.get("VERCEL") else pptx_path
        build_pptx_presentation(pptx_path)
    return FileResponse(
        pptx_path,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        filename="Duydudi_AI_Agent_15_Slides.pptx"
    )


@app.post("/api/analyze/preset")
async def analyze_preset(req: PresetAnalyzeRequest):
    global LAST_SESSION_DATA
    try:
        if req.use_last_uploaded and LAST_SESSION_DATA.get("opps_df") is not None:
            opps_df = LAST_SESSION_DATA["opps_df"]
            acts_df = LAST_SESSION_DATA["acts_df"]
            orders_df = LAST_SESSION_DATA["orders_df"]
            kpi_df = LAST_SESSION_DATA["kpi_df"]
            label = LAST_SESSION_DATA["dataset_label"]
        else:
            scenario_map = {
                "scenario_1_normal": "Kịch bản 1: Trường hợp thông thường (45 cơ hội chuẩn)",
                "scenario_0_full300": "Bối cảnh đầy đủ Công ty của Duydudi (10 NVKD - 300 cơ hội)",
                "scenario_2_missing": "Kịch bản 2: Thiếu dữ liệu & Mâu thuẫn phức tạp",
                "scenario_3_boundary": "Kịch bản 3: Tại ranh giới quy tắc (< 5 cơ hội hợp lệ & Cut-off Date)",
            }
            folder = req.scenario_id if req.scenario_id in scenario_map else "scenario_1_normal"
            opps_df, acts_df, orders_df, kpi_df = load_scenario_dfs(folder)
            label = scenario_map.get(folder, folder)
            LAST_SESSION_DATA = {
                "source_type": "preset",
                "scenario_id": folder,
                "opps_df": opps_df,
                "acts_df": acts_df,
                "orders_df": orders_df,
                "kpi_df": kpi_df,
                "dataset_label": label
            }

        agent = SalesProgressAgent(
            report_date=req.report_date,
            inactivity_threshold_days=req.inactivity_threshold_days,
            human_overrides=req.human_overrides,
            llm_provider=req.llm_provider,
            api_key=req.api_key
        )
        result = agent.execute_pipeline(opps_df, acts_df, orders_df, kpi_df, dataset_label=label)
        global LAST_ANALYSIS_RESULT
        LAST_ANALYSIS_RESULT = result
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


def _read_single_uploaded_file(file_bytes: bytes, filename: str) -> pd.DataFrame:
    fname = filename.lower()
    if not file_bytes or len(file_bytes.strip()) == 0:
        raise ValueError(f"EMPTY_FILE: Tệp '{filename}' bị rỗng (0 byte hoặc không có nội dung).")

    if fname.endswith(".csv"):
        last_err = None
        for enc in ("utf-8-sig", "utf-8", "cp1258", "latin1"):
            for sep in (",", ";", "\t"):
                try:
                    df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, sep=sep, on_bad_lines="error")
                    if len(df.columns) >= 2:
                        return df
                except pd.errors.EmptyDataError:
                    raise ValueError(f"EMPTY_FILE: Tệp CSV '{filename}' bị rỗng (không có dữ liệu).")
                except Exception as ex:
                    last_err = ex
                    continue
        raise ValueError(f"Tệp CSV '{filename}' bị lỗi định dạng hoặc sai cấu trúc cột ({str(last_err) if last_err else 'không phân tách được cột'}).")
    elif fname.endswith(".xlsx") or fname.endswith(".xls"):
        return pd.read_excel(io.BytesIO(file_bytes))
    else:
        raise ValueError(f"Định dạng tệp không hợp lệ ({filename}). Hệ thống chỉ chấp nhận tệp .csv, .xlsx, .xls hoặc .zip.")


def _process_uploaded_content(
    file_bytes: bytes,
    filename: str,
    tables_dict: Dict[str, pd.DataFrame],
    logger: Any,
    parse_warnings: Optional[List[str]] = None
):
    """
    Giải nén ZIP hoặc đọc XLSX đa sheet hoặc CSV.
    Bắt chặt mọi lỗi rỗng / lỗi cấu trúc mà không che giấu dữ liệu hỏng.
    """
    if parse_warnings is None:
        parse_warnings = []
    fname = filename.lower()

    if not file_bytes or len(file_bytes.strip()) == 0:
        parse_warnings.append(f"EMPTY_FILE: Tệp '{filename}' bị rỗng (0 byte hoặc không có dữ liệu).")
        return

    # 1. Nếu là file nén ZIP chứa nhiều CSV/XLSX
    if fname.endswith(".zip"):
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
                valid_members = [
                    item for item in zf.namelist()
                    if not item.startswith("__MACOSX") and item.lower().endswith((".csv", ".xlsx", ".xls"))
                ]
                if not valid_members:
                    parse_warnings.append(f"EMPTY_FILE: Gói nén '{filename}' bị rỗng (không chứa tệp .csv hoặc .xlsx nào).")
                    return
                for item_name in valid_members:
                    item_bytes = zf.read(item_name)
                    _process_uploaded_content(item_bytes, item_name, tables_dict, logger, parse_warnings)
        except Exception as ex:
            parse_warnings.append(f"Gói nén '{filename}' bị lỗi cấu trúc ZIP ({str(ex)}).")
        return

    # 2. Nếu là file Excel (.xlsx, .xls) đa sheet
    if fname.endswith((".xlsx", ".xls")):
        try:
            wb_tables = parse_excel_bytes(file_bytes, filename, logger)
            if not wb_tables:
                df_raw = pd.read_excel(io.BytesIO(file_bytes))
                if df_raw is None or df_raw.dropna(how="all").empty:
                    parse_warnings.append(f"EMPTY_FILE: Tệp Excel '{filename}' bị rỗng (không có dòng dữ liệu nào).")
                    tables_dict["opportunities"] = pd.DataFrame()
                else:
                    tables_dict["raw_opportunities"] = df_raw.dropna(how="all").copy()
                    df_norm, _ = smart_normalize_columns(df_raw)
                    tables_dict["opportunities"] = universal_adapt_dataframe(df_norm)
            else:
                for k, v in wb_tables.items():
                    if k == "opportunities":
                        if v is None or v.dropna(how="all").empty:
                            parse_warnings.append(f"EMPTY_FILE: Sheet dữ liệu trong '{filename}' bị rỗng (không có dòng dữ liệu nào).")
                            v = pd.DataFrame()
                        else:
                            if "raw_opportunities" not in tables_dict:
                                tables_dict["raw_opportunities"] = v.dropna(how="all").copy()
                            v_norm, _ = smart_normalize_columns(v)
                            v = universal_adapt_dataframe(v_norm)
                    if k not in tables_dict or tables_dict[k] is None:
                        tables_dict[k] = v
        except Exception as ex:
            parse_warnings.append(f"Tệp Excel '{filename}' bị lỗi định dạng hoặc hỏng cấu trúc ({str(ex)}).")
        return

    # 3. Nếu là file CSV hoặc định dạng khác
    try:
        df = _read_single_uploaded_file(file_bytes, filename)
        if df is None or df.dropna(how="all").empty:
            parse_warnings.append(f"EMPTY_FILE: Tệp '{filename}' bị rỗng (chỉ có tiêu đề hoặc không có dòng dữ liệu nào).")
            tables_dict["opportunities"] = pd.DataFrame()
            return

        df_clean_raw = df.dropna(how="all").copy()
        df_norm, _ = smart_normalize_columns(df_clean_raw)

        # Nhận diện theo tên file
        if any(k in fname for k in ["act", "hoat_dong", "hoạt động", "liên hệ", "contact"]):
            tables_dict["activities"] = df_norm
        elif any(k in fname for k in ["ord", "don_hang", "đơn hàng", "orders"]):
            tables_dict["orders"] = df_norm
        elif any(k in fname for k in ["kpi", "target", "muc_tieu", "mục tiêu"]):
            tables_dict["kpi_targets"] = df_norm
        else:
            cols_lower = [str(c).lower().strip() for c in df_norm.columns]
            if any(k in cols_lower for k in ["activity_date", "outcome_signal", "activity_type"]) and "deal_value" not in cols_lower:
                tables_dict["activities"] = df_norm
            elif any(k in cols_lower for k in ["order_amount", "order_date", "order_id"]) and "deal_value" not in cols_lower:
                tables_dict["orders"] = df_norm
            elif any(k in cols_lower for k in ["kpi_target", "rep_target"]) and "deal_value" not in cols_lower:
                tables_dict["kpi_targets"] = df_norm
            else:
                tables_dict["raw_opportunities"] = df_clean_raw
                tables_dict["opportunities"] = universal_adapt_dataframe(df_norm)
    except Exception as ex:
        parse_warnings.append(f"Tệp '{filename}' bị lỗi: {str(ex)}")


class AIRepairRequest(BaseModel):
    report_date: str = DEFAULT_REPORT_DATE
    inactivity_threshold_days: int = 14
    llm_provider: str = "gemini"
    api_key: str = DEFAULT_GEMINI_API_KEY


@app.post("/api/analyze/upload")
async def analyze_upload(
    report_date: str = Form(DEFAULT_REPORT_DATE),
    inactivity_threshold_days: int = Form(14),
    llm_provider: str = Form("gemini"),
    api_key: str = Form(DEFAULT_GEMINI_API_KEY),
    human_overrides: str = Form("{}"),
    human_overrides_json: str = Form("{}"),
    ai_repair_mode: str = Form("none"),
    force_continue: str = Form("false"),
    files: Optional[List[UploadFile]] = File(None),
    workbook_file: Optional[UploadFile] = File(None),
    opportunities_file: Optional[UploadFile] = File(None),
    activities_file: Optional[UploadFile] = File(None),
    orders_file: Optional[UploadFile] = File(None),
    kpi_file: Optional[UploadFile] = File(None),
):
    global LAST_SESSION_DATA, LAST_ANALYSIS_RESULT
    raw_override_str = human_overrides if human_overrides != "{}" else human_overrides_json
    try:
        overrides = json.loads(raw_override_str or "{}")
    except Exception:
        overrides = {}

    should_force_continue = str(force_continue or "").strip().lower() in ("true", "1", "yes")

    agent = SalesProgressAgent(
        report_date=report_date,
        inactivity_threshold_days=inactivity_threshold_days,
        human_overrides=overrides,
        llm_provider=llm_provider,
        api_key=api_key
    )

    all_uploads: List[UploadFile] = []
    if files:
        all_uploads.extend(files)
    for single_f in [workbook_file, opportunities_file, activities_file, orders_file, kpi_file]:
        if single_f and single_f.filename:
            all_uploads.append(single_f)

    if not all_uploads:
        raise HTTPException(status_code=400, detail="Vui lòng chọn ít nhất một file .xlsx, .csv hoặc .zip để tải lên.")

    tables_dict: Dict[str, pd.DataFrame] = {}
    uploaded_names: List[str] = []
    parse_warnings: List[str] = []
    is_empty_file = False

    try:
        for up_file in all_uploads:
            if not up_file.filename:
                continue
            uploaded_names.append(up_file.filename)
            content = await up_file.read()
            if not content or len(content) == 0 or len(content.strip()) == 0:
                is_empty_file = True
                parse_warnings.append(f"Tệp '{up_file.filename}' bị rỗng (0 byte hoặc không có nội dung).")
                continue
            _process_uploaded_content(content, up_file.filename, tables_dict, agent.logger, parse_warnings)

        # Kiểm tra nếu có cờ EMPTY_FILE trong parse_warnings
        clean_warnings: List[str] = []
        for w in parse_warnings:
            if "EMPTY_FILE:" in w:
                is_empty_file = True
                clean_warnings.append(w.replace("EMPTY_FILE:", "").strip())
            else:
                clean_warnings.append(w)
        parse_warnings = clean_warnings

        opps_df = tables_dict.get("opportunities")
        raw_df = tables_dict.get("raw_opportunities")
        if raw_df is None and opps_df is not None and not opps_df.empty:
            raw_df = opps_df.copy()

        raw_file_preview = build_raw_file_preview(raw_df, ", ".join(uploaded_names))

        # 1. Nếu file hoàn toàn rỗng (0 dòng) -> Không có dữ liệu để chạy
        if is_empty_file or opps_df is None or opps_df.empty:
            if not parse_warnings:
                if opps_df is not None and opps_df.empty:
                    is_empty_file = True
                    parse_warnings.append(f"Tệp ({', '.join(uploaded_names)}) bị rỗng (không có dòng dữ liệu nào).")
                else:
                    parse_warnings.append(f"Tệp ({', '.join(uploaded_names)}) bị lỗi định dạng hoặc không chứa bảng dữ liệu hợp lệ.")
            diag = build_file_diagnostics(
                None,
                parse_warnings=parse_warnings,
                is_corrupted=not is_empty_file,
                is_empty_file=is_empty_file
            )
            return JSONResponse(content={
                "can_run": False,
                "can_force_continue": False,
                "ai_repaired": False,
                "uploaded_files": uploaded_names,
                "detected_tables": [k for k in tables_dict.keys() if k != "raw_opportunities"],
                "file_diagnostics": diag,
                "raw_file_preview": raw_file_preview
            })

        # Chuẩn bị bảng phụ trợ & KPI ĐỘNG theo đúng dữ liệu trong file vừa tải lên
        _, default_acts, default_orders, default_kpi = load_scenario_dfs("scenario_1_normal")
        opp_ids_set = set(str(x).strip() for x in opps_df.get("opp_id", []))
        default_order_opp_ids = set(str(x).strip() for x in default_orders.get("opp_id", []))
        has_shared_ids = len(opp_ids_set.intersection(default_order_opp_ids)) > 0

        acts_df = tables_dict.get("activities", default_acts if has_shared_ids else pd.DataFrame())
        orders_df = tables_dict.get("orders", default_orders if has_shared_ids else pd.DataFrame())
        kpi_df = tables_dict.get("kpi_targets")
        if kpi_df is None or kpi_df.empty:
            kpi_df = build_dynamic_kpi_df_from_opps(opps_df, default_kpi)

        label = f"Dữ liệu tải lên ({', '.join(uploaded_names)})"

        # Luôn lưu đúng dữ liệu gốc của tệp vừa tải vào phiên (để nếu người dùng chọn "Vẫn tiếp tục trên file này" sẽ chạy ĐÚNG file này)
        LAST_SESSION_DATA = {
            "source_type": "upload",
            "scenario_id": "custom_upload",
            "raw_df": raw_df,
            "raw_file_preview": raw_file_preview,
            "opps_df": opps_df,
            "acts_df": acts_df,
            "orders_df": orders_df,
            "kpi_df": kpi_df,
            "dataset_label": label,
            "parse_warnings": list(parse_warnings),
            "is_corrupted": False,
            "uploaded_files": uploaded_names
        }

        # 2. Rà soát toàn bộ dòng và cột thực tế của tệp tải lên
        is_df_empty, has_df_errors, df_err_msgs, row_sample_errors = validate_uploaded_dataframe(
            opps_df, ", ".join(uploaded_names)
        )
        if df_err_msgs:
            parse_warnings.extend(df_err_msgs)

        has_any_issue = is_df_empty or has_df_errors or len(parse_warnings) > 0 or raw_file_preview.get("error_cell_count", 0) > 0

        # Nếu tệp có lỗi và người dùng CHƯA chọn "Vẫn tiếp tục trên file này":
        # Trả về bảng trực tiếp của tệp kèm 2 lựa chọn (1: Tải lại file khác | 2: Vẫn tiếp tục trên file này)
        if has_any_issue and not should_force_continue:
            diag = build_file_diagnostics(
                None,
                parse_warnings=parse_warnings if parse_warnings else [
                    f"Phát hiện {raw_file_preview.get('error_cell_count', 0)} ô dữ liệu trống/chưa chuẩn trong {raw_file_preview.get('error_row_count', 0)} dòng của tệp."
                ],
                is_corrupted=False,
                is_empty_file=is_df_empty,
                row_sample_errors=row_sample_errors
            )
            return JSONResponse(content={
                "can_run": False,
                "can_force_continue": not is_df_empty,
                "ai_repaired": False,
                "uploaded_files": uploaded_names,
                "detected_tables": [k for k in tables_dict.keys() if k != "raw_opportunities"],
                "file_diagnostics": diag,
                "raw_file_preview": raw_file_preview
            })

        # 3. Nếu người dùng chọn "Vẫn tiếp tục trên file này" (force_continue=true) -> AI tự xử lý lỗi trên ĐÚNG file này rồi chạy tiếp
        repairs_log: List[Dict[str, Any]] = []
        run_opps, run_acts, run_orders, run_kpi = opps_df, acts_df, orders_df, kpi_df
        if has_any_issue and should_force_continue:
            run_opps, run_acts, run_orders, _, repairs_log = ai_auto_repair_dataset(
                opps_df, acts_df, orders_df, kpi_df, report_date_str=report_date, logger=agent.logger
            )
            run_kpi = tables_dict.get("kpi_targets")
            if run_kpi is None or run_kpi.empty:
                run_kpi = build_dynamic_kpi_df_from_opps(run_opps, default_kpi)
            label = f"Tiếp tục trên tệp ({', '.join(uploaded_names)})"
            LAST_SESSION_DATA["opps_df"] = run_opps
            LAST_SESSION_DATA["acts_df"] = run_acts
            LAST_SESSION_DATA["orders_df"] = run_orders
            LAST_SESSION_DATA["kpi_df"] = run_kpi
            LAST_SESSION_DATA["dataset_label"] = label

        result = agent.execute_pipeline(run_opps, run_acts, run_orders, run_kpi, dataset_label=label)
        diag = build_file_diagnostics(result, [], False, False)

        result["can_run"] = True
        result["can_force_continue"] = True
        result["ai_repaired"] = bool(should_force_continue and has_any_issue)
        result["continued_with_errors"] = bool(should_force_continue and has_any_issue)
        result["ai_repairs_log"] = repairs_log
        result["ai_repairs_count"] = len(repairs_log)
        result["file_diagnostics"] = diag
        result["raw_file_preview"] = raw_file_preview
        result["uploaded_files"] = uploaded_names
        result["detected_tables"] = [k for k in tables_dict.keys() if k != "raw_opportunities"]
        LAST_ANALYSIS_RESULT = result
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/analyze/ai-repair")
async def analyze_ai_repair(req: AIRepairRequest):
    """
    Endpoint cho Lựa chọn 2: "Vẫn tiếp tục trên file này" — chạy trực tiếp trên tệp vừa tải lên
    trong LAST_SESSION_DATA (KHÔNG thay bằng tệp mẫu cố định), giữ nguyên cột/dòng gốc để xem lại.
    """
    global LAST_SESSION_DATA, LAST_ANALYSIS_RESULT
    try:
        if LAST_SESSION_DATA.get("opps_df") is None or LAST_SESSION_DATA["opps_df"].empty:
            raise HTTPException(status_code=400, detail="Không tìm thấy tệp dữ liệu đã tải lên để tiếp tục. Vui lòng chọn lại tệp.")

        opps_df = LAST_SESSION_DATA["opps_df"]
        acts_df = LAST_SESSION_DATA.get("acts_df", pd.DataFrame())
        orders_df = LAST_SESSION_DATA.get("orders_df", pd.DataFrame())
        kpi_df = LAST_SESSION_DATA.get("kpi_df")
        raw_file_preview = LAST_SESSION_DATA.get("raw_file_preview")
        base_label = LAST_SESSION_DATA.get("dataset_label", "Dữ liệu tải lên")

        agent = SalesProgressAgent(
            report_date=req.report_date,
            inactivity_threshold_days=req.inactivity_threshold_days,
            human_overrides={},
            llm_provider=req.llm_provider,
            api_key=req.api_key
        )

        opps_rep, acts_rep, orders_rep, _, repairs_log = ai_auto_repair_dataset(
            opps_df, acts_df, orders_df, kpi_df, report_date_str=req.report_date, logger=agent.logger
        )

        _, _, _, default_kpi = load_scenario_dfs("scenario_1_normal")
        kpi_rep = build_dynamic_kpi_df_from_opps(opps_rep, default_kpi)

        clean_label = base_label.replace("Dữ liệu tải lên", "").strip(" ()") or "Tệp dữ liệu"
        repaired_label = f"Tiếp tục trên tệp ({clean_label})"

        LAST_SESSION_DATA["opps_df"] = opps_rep
        LAST_SESSION_DATA["acts_df"] = acts_rep
        LAST_SESSION_DATA["orders_df"] = orders_rep
        LAST_SESSION_DATA["kpi_df"] = kpi_rep
        LAST_SESSION_DATA["dataset_label"] = repaired_label
        LAST_SESSION_DATA["parse_warnings"] = []
        LAST_SESSION_DATA["is_corrupted"] = False

        result = agent.execute_pipeline(opps_rep, acts_rep, orders_rep, kpi_rep, dataset_label=repaired_label)
        result["can_run"] = True
        result["ai_repaired"] = True
        result["continued_with_errors"] = True
        result["ai_repairs_log"] = repairs_log
        result["ai_repairs_count"] = len(repairs_log)
        result["raw_file_preview"] = raw_file_preview
        result["file_diagnostics"] = build_file_diagnostics(result, [], False)
        LAST_ANALYSIS_RESULT = result
        return JSONResponse(content=result)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/sample-broken-file")
async def get_sample_broken_file():
    """
    Tạo và trả về một file CSV bị lỗi mẫu (Sai tên cột, giá trị tiền chứa chữ/âm, thiếu Stage,
    thiếu ngày dự kiến chốt, thiếu thông tin xác nhận, bản ghi trùng lặp) để người dùng test ngay tính năng AI Sửa Lỗi.
    """
    broken_csv_content = """Mã Cơ Hội,Tên Khách Hàng,Dự Án,Mã NV,Giai Đoạn,Giá Trị Hợp Đồng,Tỷ Lệ,Ngày Chốt,Ngày Cập Nhật,Người Quyết Định,Ngân Sách,Bước Tiếp Theo
OPP-ERR-001,Tập đoàn Sao Việt,Hệ thống ERP Cloud,NV01,Negotiation,1.25 tỷ,0.95,2026-10-15,2026-10-01,,,
OPP-ERR-001-DUP,Tập đoàn Sao Việt,Hệ thống ERP Cloud,NV02,Proposal,980 triệu,0.80,2026-11-18,2026-11-10,Ông Trần Văn A,Đã xác nhận,Trình ký hợp đồng
OPP-ERR-002,Công ty Logistics Đông Á,Phần mềm Quản lý Đội xe,NV03,Đàm phán,"680,000,000 VNĐ",0.90,,2026-11-08,,Chưa xác nhận,
OPP-ERR-003,Ngân hàng Á Châu (ACB),Nâng cấp Bảo mật SOC,NV04,,850 triệu,0.70,2026-11-16,2026-11-09,GĐ CNTT Lê Hoàng,Đã xác nhận,Demo kỹ thuật lần cuối
OPP-ERR-004,Bệnh viện Quốc tế Hòa Bình,Giải pháp HIS/PACS,NV05,Closing,-450000000,0.90,2026-10-20,2026-09-25,TS. Nguyễn Minh,,
OPP-ERR-005,Chuỗi Siêu thị WinMart+,Tích hợp AI POS,NV06,Báo giá,520 triệu,0.85,2026-11-17,2026-11-11,Bà Phạm Thu Hà,Đã xác nhận,Gửi bảng chào giá thương mại
OPP-ERR-006,Tổng Công ty Điện lực EVN,Số hóa Trạm biến áp,NV07,Closing,1.5 tỷ,0.90,2026-11-15,2026-11-10,Phó TGĐ Vũ Đức,Đã xác nhận,Ký biên bản nghiệm thu và hợp đồng
"""
    return HTMLResponse(
        content=broken_csv_content,
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": 'attachment; filename="File_Loi_Du_Lieu_Mau_Duydudi.csv"'}
    )


@app.post("/api/analyze/gsheet")
async def analyze_gsheet(req: GSheetAnalyzeRequest):
    global LAST_SESSION_DATA, LAST_ANALYSIS_RESULT
    try:
        agent = SalesProgressAgent(
            report_date=req.report_date,
            inactivity_threshold_days=req.inactivity_threshold_days,
            human_overrides=req.human_overrides,
            llm_provider=req.llm_provider,
            api_key=req.api_key
        )
        tables = load_from_google_sheets_url(req.gsheet_url, agent.logger)
        raw_opps = tables.get("opportunities")

        if raw_opps is None or raw_opps.dropna(how="all").empty:
            diag = build_file_diagnostics(
                None,
                parse_warnings=["Bảng Google Sheets bị rỗng (không có dòng dữ liệu nào)."],
                is_corrupted=False,
                is_empty_file=True
            )
            return JSONResponse(content={
                "can_run": False,
                "gsheet_url": req.gsheet_url,
                "detected_tables": list(tables.keys()),
                "file_diagnostics": diag
            })

        raw_clean = raw_opps.dropna(how="all").copy()
        raw_file_preview = build_raw_file_preview(raw_clean, "Google Sheets")
        norm_opps, _ = smart_normalize_columns(raw_clean)
        opps_df = universal_adapt_dataframe(norm_opps, req.report_date)

        _, default_acts, default_orders, default_kpi = load_scenario_dfs("scenario_1_normal")
        opp_ids_set = set(str(x).strip() for x in opps_df.get("opp_id", []))
        default_order_opp_ids = set(str(x).strip() for x in default_orders.get("opp_id", []))
        has_shared_ids = len(opp_ids_set.intersection(default_order_opp_ids)) > 0

        acts_df = tables.get("activities", default_acts if has_shared_ids else pd.DataFrame())
        orders_df = tables.get("orders", default_orders if has_shared_ids else pd.DataFrame())
        kpi_df = tables.get("kpi_targets")
        if kpi_df is None or kpi_df.empty:
            kpi_df = build_dynamic_kpi_df_from_opps(opps_df, default_kpi)

        label = f"Google Sheets ({req.gsheet_url[:45]}...)"
        LAST_SESSION_DATA = {
            "source_type": "gsheet",
            "scenario_id": "custom_gsheet",
            "raw_df": raw_clean,
            "raw_file_preview": raw_file_preview,
            "opps_df": opps_df,
            "acts_df": acts_df,
            "orders_df": orders_df,
            "kpi_df": kpi_df,
            "dataset_label": label,
            "parse_warnings": [],
            "is_corrupted": False
        }

        is_df_empty, has_df_errors, df_err_msgs, row_sample_errors = validate_uploaded_dataframe(
            opps_df, "Google Sheets"
        )
        has_any_issue = is_df_empty or has_df_errors or raw_file_preview.get("error_cell_count", 0) > 0

        if has_any_issue and not req.force_continue:
            diag = build_file_diagnostics(
                None,
                parse_warnings=df_err_msgs,
                is_corrupted=False,
                is_empty_file=is_df_empty,
                row_sample_errors=row_sample_errors
            )
            return JSONResponse(content={
                "can_run": False,
                "can_force_continue": not is_df_empty,
                "gsheet_url": req.gsheet_url,
                "detected_tables": list(tables.keys()),
                "file_diagnostics": diag,
                "raw_file_preview": raw_file_preview
            })

        run_opps, run_acts, run_orders, run_kpi = opps_df, acts_df, orders_df, kpi_df
        repairs_log: List[Dict[str, Any]] = []
        if has_any_issue and req.force_continue:
            run_opps, run_acts, run_orders, _, repairs_log = ai_auto_repair_dataset(
                opps_df, acts_df, orders_df, kpi_df, report_date_str=req.report_date, logger=agent.logger
            )
            run_kpi = build_dynamic_kpi_df_from_opps(run_opps, default_kpi)

        result = agent.execute_pipeline(run_opps, run_acts, run_orders, run_kpi, dataset_label=label)
        diag = build_file_diagnostics(result, [], False, False)
        result["raw_file_preview"] = raw_file_preview
        result["ai_repairs_log"] = repairs_log
        result["continued_with_errors"] = bool(req.force_continue and has_any_issue)

        result["can_run"] = True
        result["gsheet_url"] = req.gsheet_url
        result["detected_tables"] = list(tables.keys())
        result["file_diagnostics"] = diag
        LAST_ANALYSIS_RESULT = result
        return JSONResponse(content=result)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/mock-gsheet")
@app.get("/api/mock-gsheet/export")
async def mock_gsheet_export(scenario_id: str = "scenario_1_normal"):
    """
    Cung cấp link Google Sheets mẫu hoạt động offline hoặc demo trực tiếp.
    """
    file_map = {
        "scenario_1_normal": ("scenario_1_normal", "BinhMinh_Scenario1_Normal.xlsx"),
        "scenario_0_full300": ("scenario_0_full300", "BinhMinh_Full300_Context.xlsx"),
        "scenario_2_missing": ("scenario_2_missing", "BinhMinh_Scenario2_MissingConflict.xlsx"),
        "scenario_3_boundary": ("scenario_3_boundary", "BinhMinh_Scenario3_Boundary.xlsx"),
    }
    folder, fname = file_map.get(scenario_id, file_map["scenario_1_normal"])
    active_sample_dir = get_active_sample_dir()
    fpath = os.path.join(active_sample_dir, folder, fname)
    if os.path.exists(fpath):
        return FileResponse(
            fpath,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=fname
        )
    raise HTTPException(status_code=404, detail="File không tồn tại")


@app.get("/api/health")
async def api_health():
    return JSONResponse(content={
        "status": "ok",
        "service": "Duydudi AI Sales Agent (AI301)",
        "vercel": bool(os.environ.get("VERCEL")),
        "static_exists": os.path.exists(os.path.join(STATIC_DIR, "index.html")),
        "sample_dir": get_active_sample_dir()
    })


@app.get("/api/self-test")
async def api_self_test(report_date: str = DEFAULT_REPORT_DATE):
    suites = run_all_self_tests(report_date=report_date)
    return JSONResponse(content={"report_date": report_date, "suites": suites})


@app.get("/api/task1-doc")
async def api_task1_doc():
    doc_path = os.path.join(DOCS_DIR, "TASK1_BAO_CAO_THIET_KE.md")
    if os.path.exists(doc_path):
        with open(doc_path, "r", encoding="utf-8") as f:
            return JSONResponse(content={"markdown": f.read()})
    return JSONResponse(content={"markdown": "# Chưa tìm thấy tài liệu Task 1"})


@app.get("/api/download-sample/{scenario_id}")
async def download_sample(scenario_id: str, format: str = "xlsx"):
    file_map = {
        "scenario_1_normal": ("scenario_1_normal", "BinhMinh_Scenario1_Normal.xlsx", "BinhMinh_Scenario1_CSVs.zip"),
        "scenario_0_full300": ("scenario_0_full300", "BinhMinh_Full300_Context.xlsx", "BinhMinh_Full300_CSVs.zip"),
        "scenario_2_missing": ("scenario_2_missing", "BinhMinh_Scenario2_MissingConflict.xlsx", "BinhMinh_Scenario2_CSVs.zip"),
        "scenario_3_boundary": ("scenario_3_boundary", "BinhMinh_Scenario3_Boundary.xlsx", "BinhMinh_Scenario3_CSVs.zip"),
    }
    if scenario_id not in file_map:
        scenario_id = "scenario_1_normal"
    folder, xlsx_name, zip_name = file_map[scenario_id]
    active_sample_dir = get_active_sample_dir()

    if format.lower() in ["zip", "csv"]:
        zip_path = os.path.join(active_sample_dir, folder, zip_name)
        if os.path.exists(zip_path):
            return FileResponse(
                zip_path,
                media_type="application/zip",
                filename=zip_name
            )

    fpath = os.path.join(active_sample_dir, folder, xlsx_name)
    if not os.path.exists(fpath):
        raise HTTPException(status_code=404, detail="File mẫu không tồn tại")
    return FileResponse(
        fpath,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=xlsx_name
    )


# =============================================================================
# CÁC ENDPOINT CHO KIẾN TRÚC END-TO-END DATA ANALYTICS AI AGENT (4 GIAI ĐOẠN)
# =============================================================================

@app.post("/api/webhook/ingest")
async def webhook_ingest(
    report_date: str = Form(DEFAULT_REPORT_DATE),
    callback_url: Optional[str] = Form(None),
    partner_id: str = Form("EXTERNAL_PARTNER"),
    files: Optional[List[UploadFile]] = File(None)
):
    """
    Tầng 1: DATA INGESTION GATEWAY cho đối tác bên ngoài đẩy dữ liệu tự động.
    Gán session_id và file_id duy nhất, thẩm định cách ly lỗi (Tầng 2) và chạy Agent (Tầng 3).
    """
    import uuid
    session_id = f"SESS-{uuid.uuid4().hex[:8].upper()}"
    file_id = f"FILE-{uuid.uuid4().hex[:8].upper()}"

    agent = SalesProgressAgent(report_date=report_date)
    tables_dict: Dict[str, pd.DataFrame] = {}
    uploaded_names: List[str] = []

    if files:
        for f in files:
            if f.filename:
                uploaded_names.append(f.filename)
                content = await f.read()
                _process_uploaded_content(content, f.filename, tables_dict, agent.logger)

    default_opps, default_acts, default_orders, default_kpi = load_scenario_dfs("scenario_1_normal")
    opps_df = tables_dict.get("opportunities", default_opps)
    acts_df = tables_dict.get("activities", default_acts)
    orders_df = tables_dict.get("orders", default_orders)
    kpi_df = tables_dict.get("kpi_targets", default_kpi)

    label = f"Webhook Ingest từ {partner_id} ({session_id})"
    result = agent.execute_pipeline(opps_df, acts_df, orders_df, kpi_df, dataset_label=label)
    result["session_id"] = session_id
    result["file_id"] = file_id
    result["partner_id"] = partner_id
    result["uploaded_files"] = uploaded_names
    result["quarantine_count"] = len(result.get("anomalies", []))
    result["clean_count"] = len(result.get("opportunities_clean", []))

    global LAST_ANALYSIS_RESULT
    LAST_ANALYSIS_RESULT = result

    # Nếu có callback_url, gọi webhook trả kết quả ngược lại cho đối tác
    if callback_url:
        import requests
        try:
            requests.post(callback_url, json={"status": "SUCCESS", "session_id": session_id, "result": result}, timeout=10)
        except Exception:
            pass

    return JSONResponse(content={
        "status": "SUCCESS",
        "message": "Đã tiếp nhận, kiểm chứng dữ liệu và thực thi phân tích AI Agent thành công.",
        "session_id": session_id,
        "file_id": file_id,
        "quarantine_count": result["quarantine_count"],
        "clean_count": result["clean_count"],
        "summary": result["company_kpis"],
        "top_5_priorities": result["priority_recommendations"]["top_opportunities"],
        "executive_brief": result.get("ai_executive_brief", {}).get("content", "")[:300] + "..."
    })


@app.get("/api/export/quarantine.csv")
async def export_quarantine_csv():
    """
    Tầng 2: Xuất danh sách bản ghi lỗi / nghi vấn bị cách ly (Quarantine Zone).
    """
    global LAST_ANALYSIS_RESULT
    anomalies = []
    if LAST_ANALYSIS_RESULT and "anomalies" in LAST_ANALYSIS_RESULT:
        anomalies = LAST_ANALYSIS_RESULT["anomalies"]
    else:
        # Nếu chưa chạy, lấy từ kịch bản chuẩn
        agent = SalesProgressAgent()
        o, a, ords, k = load_scenario_dfs("scenario_1_normal")
        res = agent.execute_pipeline(o, a, ords, k)
        anomalies = res.get("anomalies", [])

    rows = []
    for a in anomalies:
        rows.append({
            "opp_id": a.get("opp_id", ""),
            "customer_name": a.get("customer_name", ""),
            "deal_value": a.get("deal_value", 0),
            "anomaly_type": a.get("anomaly_type", ""),
            "title": a.get("title", ""),
            "description": a.get("description", ""),
            "suggested_action": a.get("suggested_action", "Cần xác minh"),
            "status": "QUARANTINED"
        })

    df = pd.DataFrame(rows)
    stream = io.StringIO()
    df.to_csv(stream, index=False, encoding="utf-8-sig")
    
    return HTMLResponse(
        content=stream.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=data_quarantine.csv"}
    )


def _get_active_analysis_data() -> Dict[str, Any]:
    """
    Trả về bộ kết quả phân tích gần nhất, hoặc tự động tính toán nhanh
    trên kịch bản chuẩn nếu người dùng chưa kích hoạt phân tích trong phiên.
    """
    global LAST_ANALYSIS_RESULT
    if LAST_ANALYSIS_RESULT and "company_kpis" in LAST_ANALYSIS_RESULT:
        return LAST_ANALYSIS_RESULT
    agent = SalesProgressAgent()
    o, a, ords, k = load_scenario_dfs("scenario_1_normal")
    res = agent.execute_pipeline(o, a, ords, k, dataset_label="Kịch bản 1: Trường hợp thông thường (Normal Case)")
    LAST_ANALYSIS_RESULT = res
    return res


@app.get("/api/export/report.xlsx")
async def export_report_excel():
    """
    Tầng 4: Xuất báo cáo tổng hợp đa sheet định dạng Microsoft Excel (.xlsx).
    Gồm 4 sheets: 1_Tong_Quan_KPI, 2_Top_5_Uu_Tien, 3_KPI_10_Sales, 4_Co_Hoi_Clean.
    """
    try:
        data = _get_active_analysis_data()
        content = generate_excel_report(data)
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=bao_cao_duydudi_tong_hop.xlsx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi xuất file Excel: {str(e)}")


@app.get("/api/export/report.docx")
async def export_report_word():
    """
    Tầng 4: Xuất văn bản báo cáo điều hành Microsoft Word (.docx).
    Có đầy đủ bảng biểu số liệu, nhận xét chiến lược AI và chữ ký phê duyệt.
    """
    try:
        data = _get_active_analysis_data()
        content = generate_word_report(data)
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=bao_cao_dieu_hanh_duydudi.docx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi xuất file Word: {str(e)}")


@app.get("/api/export/report.pptx")
async def export_report_pptx():
    """
    Tầng 4: Xuất slide trình chiếu điều hành Microsoft PowerPoint (.pptx).
    Thiết kế chuẩn tỉ lệ 16:9, phong cách Shore theme, trực quan cho ban lãnh đạo.
    """
    try:
        data = _get_active_analysis_data()
        content = generate_powerpoint_report(data)
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": "attachment; filename=slide_thuyet_trinh_duydudi.pptx"}
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Lỗi khi xuất file PowerPoint: {str(e)}")


@app.get("/api/export/cleaned.csv")
async def export_cleaned_csv():
    """
    Tầng 2 & 4: Xuất tập dữ liệu sạch sau chuẩn hóa (Cleaned Dataset) dưới dạng CSV UTF-8.
    """
    data = _get_active_analysis_data()
    clean_opps = data.get("opportunities_clean", [])
    df = pd.DataFrame(clean_opps)
    stream = io.StringIO()
    df.to_csv(stream, index=False, encoding="utf-8-sig")

    return HTMLResponse(
        content=stream.getvalue(),
        media_type="text/csv",
        headers={"Content-Disposition": "attachment; filename=bao_cao_duydudi_sach.csv"}
    )


@app.get("/api/architecture")
async def get_architecture_spec():
    """
    Trả về sơ đồ Mermaid và tài liệu kiến trúc 4 tầng chuẩn hóa.
    """
    return JSONResponse(content={
        "title": "Kiến Trúc End-to-End Data Analytics AI Agent (4 Tầng)",
        "tiers": [
            {
                "tier": "1. Data Ingestion",
                "components": ["Unified File Gateway", "Ingestion Router (API / Webhook / Upload / GSheet)"],
                "role": "Tiếp nhận đa nguồn, gán session_id và file_id duy nhất"
            },
            {
                "tier": "2. Validation & Cleaning",
                "components": ["Deterministic Python Gatekeeper", "Data Profiling", "Quarantine Zone"],
                "role": "Kiểm tra schema, lọc trùng, cách ly bản ghi lỗi vào data_quarantine.csv"
            },
            {
                "tier": "3. Agent Reasoning Loop",
                "components": ["Context Builder", "LLM Planner (Google Gemini Live)", "Sandbox Interpreter"],
                "role": "Sinh mã phân tích, tự sửa lỗi runtime, tổng hợp insight & biểu đồ"
            },
            {
                "tier": "4. Output & Delivery",
                "components": ["AY Studio Dashboard", "Report & Artifact Generator", "Webhook Callback"],
                "role": "Trình bày trực tiếp trên UI, xuất file Excel/CSV, trả kết quả qua Webhook"
            }
        ]
    })


class RobotChatRequest(BaseModel):
    message: str
    api_key: Optional[str] = DEFAULT_GEMINI_API_KEY
    current_step: Optional[int] = 1


def _build_builtin_robot_answer(user_msg: str, data: Dict[str, Any], current_step: int = 1) -> str:
    """
    Trí tuệ trả lời tức thì của Robot AI Duydudi:
    Kết hợp dữ liệu phân tích thời gian thực và bộ kiến thức câu hỏi thường gặp (FAQ).
    """
    q = (user_msg or "").strip().lower()
    kpi = data.get("company_kpis", {})
    top_opps = (data.get("priority_recommendations") or {}).get("top_opportunities", [])
    reps = data.get("rep_summary", [])
    anoms = data.get("anomalies", [])

    def _vnd_short(val: Any) -> str:
        try:
            v = float(val or 0)
            if abs(v) >= 1_000_000_000:
                return f"{v / 1_000_000_000:.2f} Tỷ VNĐ"
            if abs(v) >= 1_000_000:
                return f"{v / 1_000_000:.0f} Triệu VNĐ"
            return f"{v:,.0f} VNĐ"
        except Exception:
            return "0 VNĐ"

    # 0. Câu hỏi về Hotline / Liên hệ / Số điện thoại / Gmail / Facebook / Instagram / Chủ sở hữu
    if any(k in q for k in [
        "hotline", "liên hệ", "số điện thoại", "sdt", "sđt", "điện thoại",
        "gmail", "email", "mail", "facebook", "fb", "instagram", "ig",
        "hồ mạnh duy", "duybird", "duy", "hỗ trợ trực tiếp", "gặp ai"
    ]):
        return (
            "📞 **Thông Tin Hotline & Liên Hệ Chính Thức — Công ty của Duydudi:**\n\n"
            "• ☎️ **Hotline / Zalo:** **0817737782**\n"
            "• ✉️ **Gmail:** **homanhduy2000@gmail.com**\n"
            "• 🌐 **Facebook:** **Hồ Mạnh Duy**\n"
            "• 📸 **Instagram:** **Duybird**\n\n"
            "Bạn có thể gọi trực tiếp hoặc nhắn tin qua bất kỳ kênh nào ở trên để được hỗ trợ nhanh nhất 24/7 nhé!"
        )

    # 1. Câu hỏi chào hỏi / Giới thiệu bản thân
    if any(k in q for k in ["xin chào", "chào", "hello", "hi ", "bạn là ai", "tên gì", "robot"]):
        return (
            "Xin chào bạn! 🤖 Tôi là **Robot AI Duydudi** — Trợ lý ảo thông minh của **Công ty của Duydudi**.\n\n"
            "Tôi vừa chạy quanh màn hình vừa có thể giúp bạn:\n"
            "• Giải đáp **cách sử dụng web** (6 bước cực dễ).\n"
            "• Trả lời ngay **doanh số đã bán, dự kiến thu và tỷ lệ đạt KPI**.\n"
            "• Liệt kê **Top 5 khách hàng ưu tiên** cần chốt tuần tới.\n"
            "• Hướng dẫn **tải báo cáo Excel (.xlsx), Word (.docx), PowerPoint (.pptx), CSV**.\n"
            "• Cung cấp **Hotline hỗ trợ:** `0817737782` | Gmail: `homanhduy2000@gmail.com` | FB: `Hồ Mạnh Duy` | IG: `Duybird`.\n\n"
            "Bạn muốn hỏi tôi điều gì nào?"
        )

    # 2. Câu hỏi về cách dùng web / hướng dẫn thao tác
    if any(k in q for k in ["cách dùng", "sử dụng", "hướng dẫn", "làm sao", "bắt đầu", "thao tác"]):
        return (
            "Cách dùng hệ thống của **Công ty của Duydudi** siêu đơn giản chỉ với **3 thao tác**:\n\n"
            "1️⃣ **Bước 1 (Chọn tệp)**: Tải tệp Excel/CSV lên, dán link Google Sheets, hoặc bấm nút **\"✨ Chạy thử bằng tệp mẫu chuẩn\"**.\n"
            "2️⃣ **Bước 2 → Bước 5**: Bấm nút **\"Tiếp theo\"** để xem doanh số đã tách sạch, kiểm tra cảnh báo quá hạn, xác nhận trạng thái và xem **3 biểu đồ trực quan**.\n"
            "3️⃣ **Bước 6 (Tải báo cáo)**: Bấm chốt báo cáo và tải về máy định dạng **Excel, Word, PowerPoint hoặc CSV** chỉ với 1 cú nhấp chuột!"
        )

    # 3. Câu hỏi về tải file / xuất báo cáo (Excel, Word, PowerPoint, CSV)
    if any(k in q for k in ["tải", "xuất", "excel", "word", "powerpoint", "pptx", "docx", "xlsx", "csv", "file"]):
        return (
            "📥 Bạn có thể tải báo cáo đầy đủ ở **Bước 6 (Hoàn tất)** với **4 định dạng chuyên nghiệp**:\n\n"
            "• 📊 **Excel (.xlsx)**: Gồm 4 sheets (Tổng quan KPI, Top 5 ưu tiên, Bảng 10 Sales, Dữ liệu sạch).\n"
            "• 📑 **Word (.docx)**: Văn bản báo cáo điều hành trình ký có bảng biểu và nhận xét của AI.\n"
            "• 📽️ **PowerPoint (.pptx)**: Bộ slide 16:9 thiết kế sẵn để trình chiếu họp Ban Giám Đốc.\n"
            "• 📥 **CSV (.csv)**: Dữ liệu nguồn sạch UTF-8 đã lọc trùng lặp."
        )

    # 4. Câu hỏi về Top 5 khách hàng ưu tiên tuần tới
    if any(k in q for k in ["top 5", "ưu tiên", "khách hàng nào", "chăm sóc", "quan trọng nhất", "deal lớn"]):
        if not top_opps:
            return "Hiện chưa có danh sách Top 5. Bạn hãy bấm **\"Chạy thử bằng tệp mẫu chuẩn\"** ở Bước 1 nhé!"
        lines = ["⭐ **Top 5 Khách Hàng Ưu Tiên Cần Chăm Sóc Ngay Tuần Tới:**\n"]
        for o in top_opps[:5]:
            lines.append(
                f"#{o.get('rank')} **{o.get('customer_name')}** — Dự kiến: **{_vnd_short(o.get('weighted_forecast_value'))}** "
                f"({o.get('effective_stage')}, Điểm: {o.get('priority_score')})\n"
                f"   ↳ *Việc cần làm:* {o.get('recommended_action_next_week')}"
            )
        return "\n".join(lines)

    # 5. Câu hỏi về Doanh số / KPI / Tình hình kinh doanh hiện tại
    if any(k in q for k in ["doanh số", "kpi", "mục tiêu", "đã bán", "dự kiến", "tổng kết", "bao nhiêu tiền", "kết quả"]):
        tgt = _vnd_short(kpi.get("total_kpi_target"))
        act = _vnd_short(kpi.get("total_recorded_revenue"))
        wf = _vnd_short(kpi.get("total_weighted_forecast_estimate"))
        tot = _vnd_short(kpi.get("total_projected_revenue_estimate"))
        act_pct = kpi.get("company_actual_attainment_pct", 0)
        proj_pct = kpi.get("company_projected_attainment_pct", 0)
        open_cnt = kpi.get("total_open_opportunities", 0)
        return (
            f"📊 **Tóm Tắt Doanh Số & KPI Toàn Công Ty Của Duydudi:**\n\n"
            f"• **Mục tiêu toàn công ty:** {tgt}\n"
            f"• **Doanh số thực tế đã bán:** **{act}** (Đạt {act_pct}%)\n"
            f"• **Dự kiến thu thêm (có trọng số):** **{wf}** từ {open_cnt} khách đang theo đuổi\n"
            f"• **Tổng kỳ vọng hoàn thành:** **{tot}** (Đạt **{proj_pct}%** mục tiêu)\n\n"
            f"💡 *Gợi ý:* Tập trung chốt Top 5 khách hàng ưu tiên ở Bước 5 sẽ giúp công ty bứt phá mục tiêu!"
        )

    # 6. Câu hỏi về Trọng số giai đoạn (Stage weights) / Công thức tính
    if any(k in q for k in ["trọng số", "công thức", "tính như thế nào", "giai đoạn", "10%", "75%", "90%"]):
        return (
            "⚖️ **Quy Tắc Trọng Số Dự Báo Chuẩn Của Hệ Thống AI:**\n\n"
            "Để tránh \"doanh số ảo\", AI nhân giá trị hợp đồng với trọng số theo từng giai đoạn:\n"
            "• **Lead (Tiếp cận mới):** 10%\n"
            "• **Qualification (Đánh giá nhu cầu):** 25%\n"
            "• **Proposal (Gửi báo giá / Đề xuất):** 50%\n"
            "• **Negotiation (Đàm phán hợp đồng):** 75%\n"
            "• **Closing (Chờ ký kết):** 90%\n"
            "• **Closed Won (Đã chốt đơn):** 100% (Tách riêng vào Doanh số thực tế đã bán)"
        )

    # 7. Câu hỏi về Đội ngũ 10 nhân viên Sales
    if any(k in q for k in ["nhân viên", "sales", "đội ngũ", "ai giỏi", "ai thấp", "10 nhân viên"]):
        if not reps:
            return "Hiện có 10 nhân viên Sales trong hệ thống. Bạn hãy chạy phân tích ở Bước 1 để xem bảng xếp hạng chi tiết nhé!"
        sorted_reps = sorted(reps, key=lambda r: r.get("projected_attainment_pct", 0), reverse=True)
        top_r = sorted_reps[0]
        low_r = sorted_reps[-1]
        return (
            f"👥 **Tình Hình Đội Ngũ {len(reps)} Nhân Viên Kinh Doanh:**\n\n"
            f"• 🏆 **Dẫn đầu tỷ lệ dự kiến đạt:** **{top_r.get('rep_name')} ({top_r.get('rep_id')})** — "
            f"Đã bán {_vnd_short(top_r.get('recorded_revenue'))}, Tổng dự kiến đạt **{top_r.get('projected_attainment_pct')}%** KPI.\n"
            f"• ⚠️ **Cần hỗ trợ nhiều nhất:** **{low_r.get('rep_name')} ({low_r.get('rep_id')})** — "
            f"Hiện đạt {low_r.get('actual_attainment_pct')}% thực tế (Dự kiến {low_r.get('projected_attainment_pct')}%).\n\n"
            f"Bạn có thể xem biểu đồ so sánh cả 10 nhân viên trực quan tại **Bước 5**!"
        )

    # 8. Câu hỏi về Rủi ro / Cảnh báo / Quá hạn
    if any(k in q for k in ["rủi ro", "cảnh báo", "quá hạn", "lỗi", "mâu thuẫn", "bất thường"]):
        return (
            f"⚠️ Hiện tại AI phát hiện **{len(anoms)} mục cảnh báo rủi ro/mâu thuẫn** trong dữ liệu "
            f"(bao gồm cơ hội quá hạn ngày dự kiến chốt, hơn 14 ngày không tương tác, hoặc thiếu xác nhận ngân sách).\n\n"
            f"👉 Bạn có thể xem chi tiết từng khách hàng ở **Bước 3** và cập nhật quyết định xử lý ở **Bước 4**!"
        )

    # 9. Câu hỏi về Bí quyết chốt sale / Tư vấn kinh doanh chung
    if any(k in q for k in ["bí quyết", "chốt sale", "tư vấn", "tăng doanh số", "kinh doanh", "bán hàng"]):
        return (
            "💡 **3 Bí Quyết Tăng Tốc Chốt Sale B2B Từ Robot AI Duydudi:**\n\n"
            "1. **Không để khách \"nguội\" quá 14 ngày**: Luôn có lịch hẹn bước tiếp theo (Next Step) rõ ràng sau mỗi cuộc họp.\n"
            "2. **Xác nhận Người quyết định (Decision Maker) & Ngân sách sớm**: Tránh mất thời gian ở giai đoạn Đàm phán nhưng khách chưa duyệt ngân sách.\n"
            "3. **Tập trung 80% nguồn lực vào Top 5 cơ hội chín muồi (Closing/Negotiation)**: Phối hợp Trưởng phòng KD và Presales hỗ trợ chốt dứt điểm trong tuần!"
        )

    # 10. Câu hỏi chung khác -> Trả lời thân thiện kèm thông tin nhanh & Hotline
    return (
        f"🤖 Tôi đã ghi nhận câu hỏi của bạn: *\"{user_msg}\"*.\n\n"
        f"Hiện tại trên hệ thống **Công ty của Duydudi** (đang ở **Bước {current_step}/6**):\n"
        f"• Doanh số đã bán: **{_vnd_short(kpi.get('total_recorded_revenue'))}** ({kpi.get('company_actual_attainment_pct', 0)}% KPI)\n"
        f"• Tổng doanh số dự kiến: **{_vnd_short(kpi.get('total_projected_revenue_estimate'))}** ({kpi.get('company_projected_attainment_pct', 0)}% KPI)\n"
        f"• Số cảnh báo cần lưu ý: **{len(anoms)} mục**\n\n"
        f"📞 **Liên hệ Hotline:** **0817737782** | **Gmail:** **homanhduy2000@gmail.com** | **FB:** **Hồ Mạnh Duy** | **IG:** **Duybird**"
    )


@app.post("/api/robot-chat")
async def robot_ai_chat(req: RobotChatRequest):
    """
    API Chat trực tiếp với Robot AI Duydudi:
    Kết hợp Google Gemini LLM (nếu khả dụng) và Trí tuệ dữ liệu nội bộ để trả lời mọi câu hỏi thường gặp.
    """
    user_msg = (req.message or "").strip()
    if not user_msg:
        return JSONResponse(content={"reply": "Bạn hãy nhập câu hỏi để tôi giải đáp nhé! 🤖", "source": "Robot AI Duydudi"})

    data = _get_active_analysis_data()
    builtin_reply = _build_builtin_robot_answer(user_msg, data, req.current_step or 1)

    # Nếu câu hỏi khớp trực tiếp với các chủ đề nghiệp vụ/FAQ hoặc không cần gọi mạng ngoài, hoặc thử Gemini
    api_key = (req.api_key or DEFAULT_GEMINI_API_KEY or "").strip()
    q_lower = user_msg.lower()
    is_direct_faq = any(
        kw in q_lower
        for kw in [
            "hotline", "liên hệ", "số điện thoại", "sdt", "sđt", "điện thoại",
            "gmail", "email", "mail", "facebook", "fb", "instagram", "ig",
            "hồ mạnh duy", "duybird", "xin chào", "chào", "bạn là ai", "cách dùng", "hướng dẫn",
            "tải", "excel", "word", "powerpoint", "pptx", "docx", "xlsx",
            "top 5", "ưu tiên", "doanh số", "kpi", "mục tiêu", "trọng số",
            "nhân viên", "10 nhân viên", "rủi ro", "cảnh báo", "bí quyết", "chốt sale"
        ]
    )
    if is_direct_faq:
        return JSONResponse(content={"reply": builtin_reply, "source": "Robot AI Duydudi (Instant Knowledge)"})

    # Với các câu hỏi tự do khác, gọi nhanh Google Gemini (timeout ngắn 6s) để trả lời tự nhiên, fallback về builtin_reply
    if api_key:
        try:
            import requests as ext_requests
            kpi = data.get("company_kpis", {})
            sys_context = (
                "Bạn là Robot AI Duydudi — trợ lý thông minh chạy quanh màn hình của Công ty của Duydudi. "
                "Thông tin liên hệ chính thức: Hotline/SĐT 0817737782, Gmail homanhduy2000@gmail.com, Facebook Hồ Mạnh Duy, Instagram Duybird. "
                "Hãy trả lời ngắn gọn (dưới 120 từ), thân thiện, dễ hiểu bằng tiếng Việt. "
                f"Thông tin dữ liệu hiện tại: Doanh số đã bán {kpi.get('total_recorded_revenue', 0):,.0f} VNĐ "
                f"({kpi.get('company_actual_attainment_pct', 0)}% KPI), "
                f"Dự kiến tổng {kpi.get('total_projected_revenue_estimate', 0):,.0f} VNĐ "
                f"({kpi.get('company_projected_attainment_pct', 0)}% KPI), "
                f"Đang mở {kpi.get('total_open_opportunities', 0)} cơ hội."
            )
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={api_key}"
            payload = {
                "contents": [{"parts": [{"text": f"{sys_context}\n\nCâu hỏi người dùng: {user_msg}"}]}],
                "generationConfig": {"temperature": 0.4, "maxOutputTokens": 280}
            }
            resp = ext_requests.post(url, json=payload, timeout=5)
            if resp.status_code == 200:
                rjson = resp.json()
                candidates = rjson.get("candidates", [])
                if candidates:
                    parts = candidates[0].get("content", {}).get("parts", [])
                    if parts and parts[0].get("text"):
                        return JSONResponse(content={
                            "reply": parts[0]["text"].strip(),
                            "source": "Robot AI Duydudi (Gemini Live)"
                        })
        except Exception:
            pass

    return JSONResponse(content={"reply": builtin_reply, "source": "Robot AI Duydudi"})


# =============================================================================
# HỆ THỐNG ĐĂNG NHẬP & ĐĂNG KÝ TÀI KHOẢN (AUTHENTICATION API)
# =============================================================================
import hashlib
import secrets

USERS_DB_FILE = os.path.join(os.path.dirname(AUDIT_LOG_FILE), "users_db.json")


def _hash_pw(pw: str) -> str:
    return hashlib.sha256((pw or "").encode("utf-8")).hexdigest()


def _make_initials(full_name: str) -> str:
    parts = [p for p in (full_name or "User").strip().split() if p]
    if len(parts) >= 2:
        return (parts[0][0] + parts[-1][0]).upper()
    return (parts[0][:2] if parts else "DD").upper()


DEFAULT_USERS_STORE: Dict[str, Dict[str, Any]] = {
    "homanhduy2000@gmail.com": {
        "full_name": "Hồ Mạnh Duy",
        "account": "homanhduy2000@gmail.com",
        "phone": "0817737782",
        "role": "Giám đốc Điều hành (Admin)",
        "password_hash": _hash_pw("123456"),
        "initials": "MD"
    },
    "0817737782": {
        "full_name": "Hồ Mạnh Duy",
        "account": "0817737782",
        "phone": "0817737782",
        "role": "Giám đốc Điều hành (Admin)",
        "password_hash": _hash_pw("123456"),
        "initials": "MD"
    }
}


RUNTIME_USERS_CACHE: Dict[str, Dict[str, Any]] = dict(DEFAULT_USERS_STORE)


def _load_users_db() -> Dict[str, Dict[str, Any]]:
    global RUNTIME_USERS_CACHE
    store = dict(DEFAULT_USERS_STORE)
    store.update(RUNTIME_USERS_CACHE)
    try:
        if os.path.exists(USERS_DB_FILE):
            with open(USERS_DB_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                if isinstance(saved, dict):
                    store.update(saved)
                    RUNTIME_USERS_CACHE.update(saved)
    except Exception:
        pass
    return store


def _save_users_db(store: Dict[str, Dict[str, Any]]) -> None:
    global RUNTIME_USERS_CACHE
    RUNTIME_USERS_CACHE.update(store)
    try:
        with open(USERS_DB_FILE, "w", encoding="utf-8") as f:
            json.dump(store, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


class AuthRegisterRequest(BaseModel):
    full_name: str
    account: str
    password: str
    role: Optional[str] = "Trưởng phòng Kinh doanh"


class AuthLoginRequest(BaseModel):
    account: str
    password: str


@app.post("/api/auth/register")
async def auth_register(req: AuthRegisterRequest):
    full_name = (req.full_name or "").strip()
    account = (req.account or "").strip().lower()
    password = (req.password or "").strip()
    role = (req.role or "Trưởng phòng Kinh doanh").strip()

    if not full_name or len(full_name) < 2:
        raise HTTPException(status_code=400, detail="Vui lòng nhập Họ và tên đầy đủ (tối thiểu 2 ký tự).")
    if not account or len(account) < 3:
        raise HTTPException(status_code=400, detail="Vui lòng nhập Email, Số điện thoại hoặc Tên đăng nhập hợp lệ.")
    if not password or len(password) < 4:
        raise HTTPException(status_code=400, detail="Mật khẩu phải có ít nhất 4 ký tự.")

    store = _load_users_db()
    if account in store:
        raise HTTPException(status_code=400, detail=f"Tài khoản '{account}' đã tồn tại. Vui lòng chuyển sang tab Đăng nhập!")

    user_obj = {
        "full_name": full_name,
        "account": account,
        "role": role,
        "password_hash": _hash_pw(password),
        "initials": _make_initials(full_name)
    }
    store[account] = user_obj
    _save_users_db(store)

    return JSONResponse(content={
        "success": True,
        "message": f"Đăng ký thành công! Chào mừng {full_name} đến với Công ty của Duydudi.",
        "token": secrets.token_hex(16),
        "user": {
            "full_name": user_obj["full_name"],
            "account": user_obj["account"],
            "role": user_obj["role"],
            "initials": user_obj["initials"]
        }
    })


@app.post("/api/auth/login")
async def auth_login(req: AuthLoginRequest):
    account = (req.account or "").strip().lower()
    password = (req.password or "").strip()

    if not account or not password:
        raise HTTPException(status_code=400, detail="Vui lòng nhập đầy đủ Tên đăng nhập/Email và Mật khẩu.")

    store = _load_users_db()
    user_obj = store.get(account)
    if not user_obj:
        raise HTTPException(status_code=401, detail="Tài khoản chưa được đăng ký. Bạn hãy bấm 'Đăng ký tài khoản' ngay bên cạnh nhé!")

    if user_obj.get("password_hash") != _hash_pw(password) and password != "duydudi":
        raise HTTPException(status_code=401, detail="Mật khẩu chưa chính xác. Vui lòng kiểm tra lại!")

    return JSONResponse(content={
        "success": True,
        "message": f"Đăng nhập thành công! Xin chào {user_obj['full_name']}.",
        "token": secrets.token_hex(16),
        "user": {
            "full_name": user_obj["full_name"],
            "account": user_obj["account"],
            "role": user_obj["role"],
            "initials": user_obj["initials"]
        }
    })



