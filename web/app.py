"""
审计底稿质检Agent - Web应用
FastAPI后端 + 前端页面
"""

import os
import sys
import json
import uuid
import shutil
from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from typing import Optional

# 将项目根目录添加到 Python 路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.app_paths import (
    AI_CONFIG_FILE,
    APP_VERSION,
    CACHE_DIR,
    LAUNCHER_CONFIG_FILE,
    LOG_DIR,
    OUTPUT_DIR as USER_OUTPUT_DIR,
    ROOT_DIR,
    UPLOAD_DIR as USER_UPLOAD_DIR,
    load_json,
    resource_dir,
    save_json,
)

# 每位测试人员的数据都保存在自己的LOCALAPPDATA目录。
UPLOAD_DIR = str(USER_UPLOAD_DIR)
OUTPUT_DIR = str(USER_OUTPUT_DIR)
CONFIG_FILE = str(AI_CONFIG_FILE)
WEB_DIR = resource_dir() / "web"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _load_config() -> dict:
    """从文件读取AI配置"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def _init_config_from_file():
    """启动时从配置文件加载AI配置到环境变量（必须在导入src模块之前调用）"""
    config = _load_config()
    if config.get("api_key"):
        os.environ["AI_PROVIDER"] = config.get("provider", "auto")
        os.environ["AI_API_KEY"] = config.get("api_key", "")
        if config.get("model"):
            os.environ["AI_MODEL"] = config["model"]
        if config.get("base_url"):
            os.environ["AI_BASE_URL"] = config["base_url"]
        os.environ["ENABLE_AI_CHECK"] = str(config.get("enable_ai_check", True)).lower()


# 启动时加载配置文件（必须在导入src模块之前，否则config变量会绑定空值）
_init_config_from_file()


def _clean_python_cache():
    """清除所有 __pycache__ 和 .pyc 文件，确保加载最新代码"""
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    deleted = 0
    for root, dirs, _files in os.walk(project_root):
        if "__pycache__" in dirs:
            cache_dir = os.path.join(root, "__pycache__")
            try:
                shutil.rmtree(cache_dir, ignore_errors=True)
                deleted += 1
            except Exception:
                pass
    # 清除残留 .pyc 文件
    for root, _dirs, files in os.walk(project_root):
        for f in files:
            if f.endswith(".pyc"):
                try:
                    os.remove(os.path.join(root, f))
                except Exception:
                    pass
    import importlib
    importlib.invalidate_caches()
    if deleted:
        print(f"[缓存清理] 已清除 {deleted} 个 __pycache__ 目录")


_clean_python_cache()

from src.checker import QualityChecker
from src.models import Severity, RuleStatus
from src.baseline import load_projects, save_project, get_project, extract_task_baseline
from src.project_config import (
    ProjectConfigValidationError,
    build_project_config_template,
    parse_project_config_workbook,
    validate_project_config,
)

# 记录启动时源码文件的最后修改时间，后续仅在有代码变动时才触发热重载
import glob as _glob_mtime
_SRC_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src")
_last_code_mtime = 0
for _pyf in _glob_mtime.glob(os.path.join(_SRC_DIR, "**", "*.py"), recursive=True):
    try:
        _mt = os.path.getmtime(_pyf)
        if _mt > _last_code_mtime:
            _last_code_mtime = _mt
    except OSError:
        pass

app = FastAPI(title="审计底稿复核系统", version=APP_VERSION)

# 挂载静态文件
app.mount("/static", StaticFiles(directory=str(WEB_DIR / "static")), name="static")

checker = QualityChecker()


def _save_config(config: dict):
    """将AI配置写入文件"""
    # 防护：如果前端传的api_key为空或包含脱敏标记***，保留旧密钥不覆盖
    new_api_key = config.get("api_key", "")
    if not new_api_key or "***" in new_api_key:
        old_config = _load_config()
        old_key = old_config.get("api_key", "")
        if old_key and "***" not in old_key:
            config["api_key"] = old_key
        elif not new_api_key and not old_key:
            pass  # 新旧都为空，保持空
        elif "***" in new_api_key:
            print(f"[配置保护] 检测到脱敏密钥，已保留旧密钥")
            config["api_key"] = old_key if old_key and "***" not in old_key else ""

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    # 同步更新环境变量，使当前进程立即生效
    os.environ["AI_PROVIDER"] = config.get("provider", "auto")
    os.environ["AI_API_KEY"] = config.get("api_key", "")
    if config.get("model"):
        os.environ["AI_MODEL"] = config["model"]
    elif "AI_MODEL" in os.environ:
        del os.environ["AI_MODEL"]
    # base_url只在非空时设置，空字符串表示使用供应商默认地址
    if config.get("base_url"):
        os.environ["AI_BASE_URL"] = config["base_url"]
    elif "AI_BASE_URL" in os.environ:
        del os.environ["AI_BASE_URL"]
    os.environ["ENABLE_AI_CHECK"] = str(config.get("enable_ai_check", True)).lower()
    # 重置checker以使用新配置
    import importlib
    import src.config as cfg
    importlib.reload(cfg)
    import src.ai_client as ac
    importlib.reload(ac)
    global checker
    checker = QualityChecker()


class AIConfigRequest(BaseModel):
    provider: str = "auto"
    api_key: str = ""
    model: str = ""
    base_url: str = ""
    enable_ai_check: bool = True


class LauncherConfigRequest(BaseModel):
    update_source: str = ""


class ProjectMappingRequest(BaseModel):
    project_code: str
    basic_info: dict = Field(default_factory=dict)
    materiality: Optional[dict] = None
    subject_assertions: Optional[dict] = None
    mapping: Optional[dict] = None


@app.get("/", response_class=HTMLResponse)
async def index():
    """返回首页"""
    html_path = WEB_DIR / "templates" / "index.html"
    with open(html_path, "r", encoding="utf-8") as f:
        return f.read()


@app.get("/api/projects")
async def list_projects():
    """列出项目mapping配置"""
    return {"projects": list(load_projects().values())}


@app.post("/api/projects")
async def save_project_mapping(req: ProjectMappingRequest):
    """保存项目层级mapping配置"""
    try:
        validated = None
        if req.materiality is not None or req.subject_assertions is not None:
            validated = validate_project_config(
                req.project_code,
                req.basic_info,
                req.materiality,
                req.subject_assertions,
            )
        record = save_project(
            req.project_code,
            req.mapping,
            validated["basic_info"] if validated else req.basic_info,
            validated["materiality"] if validated else req.materiality,
            validated["subject_assertions"] if validated else req.subject_assertions,
        )
        return {"message": "项目mapping保存成功", "project": record}
    except (ProjectConfigValidationError, ValueError) as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/projects/template")
async def download_project_config_template():
    """下载项目基础信息配置模板。"""
    content = build_project_config_template()
    headers = {
        "Content-Disposition": (
            "attachment; filename*=UTF-8''"
            "%E9%A1%B9%E7%9B%AE%E5%9F%BA%E7%A1%80%E4%BF%A1%E6%81%AF%E9%85%8D%E7%BD%AE%E6%A8%A1%E6%9D%BF.xlsx"
        )
    }
    return StreamingResponse(
        content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


@app.post("/api/projects/import/preview")
async def preview_project_config_import(file: UploadFile = File(...)):
    """解析并校验项目配置表，返回待确认的数据。"""
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="项目配置表仅支持.xlsx格式")
    try:
        project = parse_project_config_workbook(file.file)
        existing = get_project(project["project_code"])
        return {
            "message": "配置表校验通过",
            "project": project,
            "will_update_existing": existing is not None,
            "preserved_mapping": (existing or {}).get("mapping", {}),
        }
    except ProjectConfigValidationError as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/baseline/extract")
async def extract_baseline(
    project_code: str = Form(...),
    subject: str = Form("management_sales_expenses"),
    include_te_sad: bool = Form(True),
    include_a3: bool = Form(True),
    file: UploadFile = File(...),
):
    """按项目mapping从本次上传的最新版TE/SAD/A3表中提取基准值"""
    if subject not in ("fixed_assets", "management_sales_expenses"):
        raise HTTPException(status_code=400, detail="当前仅支持固定资产或费用科目")
    project = get_project(project_code)
    if not project:
        raise HTTPException(status_code=404, detail="项目mapping不存在")
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="仅支持Excel文件格式")

    task_id = str(uuid.uuid4())[:8]
    upload_path = os.path.join(UPLOAD_DIR, f"{task_id}_baseline_{file.filename}")
    with open(upload_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        baseline = extract_task_baseline(
            upload_path,
            project.get("mapping", {}),
            subject,
            include_te_sad=include_te_sad,
            include_a3=include_a3,
        )
        return {"baseline": baseline, "source_file": file.filename}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"基准信息提取失败: {str(e)}")


@app.post("/api/check")
async def check_workbook(
    file: UploadFile = File(...),
    tod_file: Optional[UploadFile] = File(None),
    subject: str = Form("management_sales_expenses"),
    project_code: Optional[str] = Form(None),
    te_sad_source: str = Form("none"),
    baseline_json: Optional[str] = Form(None),
    te_sad_enabled: bool = Form(False),
    a3_enabled: bool = Form(False),
    enable_ai: bool = Form(True),
    manual_te: Optional[float] = Form(None),
    manual_sad: Optional[float] = Form(None),
):
    """
    上传底稿并执行质检

    Args:
        file: 上传的Excel文件（主底稿，如U_exp SWP）
        tod_file: 可选的上传TOD底稿文件

    Returns:
        质检结果摘要
    """
    # 验证文件类型
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="仅支持Excel文件格式")
    if tod_file and not tod_file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="TOD文件仅支持Excel格式")
    if subject not in ("fixed_assets", "management_sales_expenses"):
        raise HTTPException(status_code=400, detail="当前仅支持固定资产或费用科目")

    # 保存上传文件
    task_id = str(uuid.uuid4())[:8]
    upload_path = os.path.join(UPLOAD_DIR, f"{task_id}_{file.filename}")
    with open(upload_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    # 保存TOD文件（如有）
    tod_upload_path = None
    if tod_file:
        tod_upload_path = os.path.join(UPLOAD_DIR, f"{task_id}_tod_{tod_file.filename}")
        with open(tod_upload_path, "wb") as f:
            shutil.copyfileobj(tod_file.file, f)

    # 执行质检
    project = None
    if project_code and project_code.strip():
        project = get_project(project_code)
        if not project:
            raise HTTPException(status_code=404, detail=f"未找到项目配置：{project_code.strip()}")

    baseline = {}
    raw_baseline = {}
    if baseline_json:
        try:
            raw_baseline = json.loads(baseline_json)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="baseline_json格式错误")
    baseline["te_sad_enabled"] = te_sad_enabled
    baseline["a3_enabled"] = a3_enabled

    project_materiality = (project or {}).get("materiality", {})
    if project_materiality.get("pm") is not None:
        baseline["pm"] = {
            "value": float(project_materiality["pm"]),
            "source": {"method": "project_config"},
        }

    if te_sad_enabled and te_sad_source == "project_config":
        if project_materiality.get("te") is None or project_materiality.get("sad") is None:
            raise HTTPException(status_code=400, detail="项目配置中未填写PM/TE/SAD")
        baseline["te"] = {
            "value": float(project_materiality["te"]),
            "source": {"method": "project_config"},
        }
        baseline["sad"] = {
            "value": float(project_materiality["sad"]),
            "source": {"method": "project_config"},
        }
    elif te_sad_enabled and isinstance(raw_baseline, dict):
        if "te" in raw_baseline:
            baseline["te"] = raw_baseline["te"]
        if "sad" in raw_baseline:
            baseline["sad"] = raw_baseline["sad"]
    if a3_enabled and isinstance(raw_baseline, dict) and "a3" in raw_baseline:
        baseline["a3"] = raw_baseline["a3"]

    if te_sad_enabled and manual_te is not None:
        baseline["te"] = {"value": float(manual_te), "source": {"method": "manual"}}
    if te_sad_enabled and manual_sad is not None:
        baseline["sad"] = {"value": float(manual_sad), "source": {"method": "manual"}}
    if te_sad_enabled and "te" not in baseline and project_materiality.get("te") is not None:
        baseline["te"] = {
            "value": float(project_materiality["te"]),
            "source": {"method": "project_config"},
        }
    if te_sad_enabled and "sad" not in baseline and project_materiality.get("sad") is not None:
        baseline["sad"] = {
            "value": float(project_materiality["sad"]),
            "source": {"method": "project_config"},
        }

    try:
        global _last_code_mtime

        # 扫描 src 源码文件 mtime，仅当有文件更新时才热重载模块
        import glob as _glob
        _current_mtime = 0
        for _pyf in _glob.glob(os.path.join(_SRC_DIR, "**", "*.py"), recursive=True):
            try:
                _mt = os.path.getmtime(_pyf)
                if _mt > _current_mtime:
                    _current_mtime = _mt
            except OSError:
                pass

        if _current_mtime > _last_code_mtime:
            import importlib
            print(f"[热重载] 检测到源码变动，重新加载模块...")
            importlib.invalidate_caches()
            to_drop = [k for k in list(sys.modules.keys()) if k.startswith("src")]
            for k in to_drop:
                del sys.modules[k]
            from src.config import reload_config; reload_config()
            _last_code_mtime = _current_mtime

        from src.checker import QualityChecker as QC
        import src.checker as _checker_mod
        check_instance = QC(enable_ai=enable_ai)
        # 记录模块加载时刻，方便调试
        import datetime as _dt
        check_instance._dbg_file = getattr(_checker_mod, '_MODULE_FILE', '?')
        check_instance._dbg_time = getattr(_checker_mod, '_MODULE_LOADED_AT', 0)
        check_instance._dbg_iso = _dt.datetime.fromtimestamp(check_instance._dbg_time).isoformat()
        # 如有TOD底稿，合并到主底稿中
        if tod_upload_path:
            from src.parser import WorkbookParser as WBP
            tod_parser = WBP()
            tod_wb = tod_parser.parse(tod_upload_path)
            # Parse main workbook
            main_parser = WBP()
            main_wb = main_parser.parse(upload_path)
            # Merge TOD sheets into main workbook
            for sn, df in tod_wb.sheets.items():
                if sn not in main_wb.sheets and not sn.startswith("DS_") and sn != "SkywindSettingSheet":
                    main_wb.sheets[sn] = df
            # Re-save merged workbook
            import openpyxl as _oxl
            merged_wb = _oxl.load_workbook(upload_path)
            src_wb = _oxl.load_workbook(tod_upload_path)
            for sn in src_wb.sheetnames:
                if sn not in merged_wb.sheetnames and not sn.startswith("DS_") and sn != "SkywindSettingSheet":
                    ws = src_wb[sn]
                    new_ws = merged_wb.create_sheet(title=sn)
                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                        for cell in row:
                            new_ws[cell.coordinate].value = cell.value
            merged_wb.save(upload_path)
            merged_wb.close()
            src_wb.close()
            print(f"[TOD合并] 已将{len(tod_wb.sheets)}个TOD Sheet合并至主底稿")

        workbook, results, execution_records = check_instance.check(
            upload_path,
            context={
                "subject": subject,
                "baseline": baseline,
                "project_config": project or {},
            }
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        os.remove(upload_path)
        raise HTTPException(status_code=500, detail=f"质检执行失败: {str(e)}")

    # 热重载后必须重新导入枚举类型，因为旧 Severity/RuleStatus 与新模块中的不是同一对象
    # Python Enum 按 identity 比较，新旧模块的 Severity.HIGH == Severity.HIGH 会返回 False
    from src.models import Severity, RuleStatus

    # 生成质检报告。每次从当前已加载模块创建实例，避免热重载后持有旧类，
    # 也避免在接口函数内给全局reporter赋值造成UnboundLocalError。
    output_filename = f"{task_id}_质检报告_{file.filename}"
    output_path = os.path.join(OUTPUT_DIR, output_filename)
    try:
        from src.reporter import QualityReporter as CurrentQualityReporter
        current_reporter = CurrentQualityReporter()
        current_reporter.generate_report(
            workbook, results, output_path, execution_records
        )
        display_results = current_reporter._aggregate_results(results)
    except Exception as e:
        import traceback
        traceback.print_exc()
        if os.path.exists(upload_path):
            os.remove(upload_path)
        raise HTTPException(status_code=500, detail=f"质检报告生成失败: {str(e)}")

    # 按状态分类统计
    passed_count = sum(1 for er in execution_records if er.status == RuleStatus.PASSED)
    failed_count = sum(1 for er in execution_records if er.status == RuleStatus.FAILED)
    skipped_count = sum(1 for er in execution_records if er.status == RuleStatus.SKIPPED)
    ai_disabled_count = sum(1 for er in execution_records if er.status == RuleStatus.AI_DISABLED)

    def _passed_record_payload(er):
        evidence_sheet, evidence_cell, _, _ = current_reporter._display_record_evidence(er)
        return {
            "rule_id": er.rule_id,
            "rule_name": er.rule_name,
            "category": er.category.value,
            "sheet": evidence_sheet,
            "cell": evidence_cell,
            "execution_detail": er.execution_detail,
        }

    # 生成结果摘要
    summary = {
        "task_id": task_id,
        "filename": file.filename,
        "total_issues": len(display_results),
        "high": sum(1 for r in display_results if r.severity == Severity.HIGH),
        "medium": sum(1 for r in display_results if r.severity == Severity.MEDIUM),
        "low": sum(1 for r in display_results if r.severity == Severity.LOW),
        "review": sum(1 for r in display_results if r.severity == Severity.REVIEW),
        "rule_summary": {
            "total": len(execution_records),
            "passed": passed_count,
            "failed": failed_count,
            "skipped": skipped_count,
            "ai_disabled": ai_disabled_count,
        },
        "issues": [
            {
                "rule_id": r.rule_id,
                "rule_name": r.rule_name,
                "category": r.category.value,
                "severity": r.severity.value,
                "sheet": current_reporter._display_location(r)[0],
                "cell": current_reporter._display_location(r)[1],
                "message": r.message,
                "expected": r.expected,
                "actual": r.actual,
                "ai_suggestion": r.ai_suggestion,
                "execution_detail": r.execution_detail,
            }
            for r in display_results
        ],
        "execution_records": {
            "passed": [
                _passed_record_payload(er)
                for er in execution_records if er.status == RuleStatus.PASSED
            ],
            "failed": [
                {"rule_id": er.rule_id, "rule_name": er.rule_name, "category": er.category.value, "issue_count": len(er.issues)}
                for er in execution_records if er.status == RuleStatus.FAILED
            ],
            "skipped": [
                {"rule_id": er.rule_id, "rule_name": er.rule_name, "category": er.category.value, "reason": er.skipped_reason}
                for er in execution_records if er.status == RuleStatus.SKIPPED
            ],
            "ai_disabled": [
                {"rule_id": er.rule_id, "rule_name": er.rule_name, "category": er.category.value, "reason": er.skipped_reason}
                for er in execution_records if er.status == RuleStatus.AI_DISABLED
            ],
        },
        "output_file": output_filename,
        "baseline": baseline,
        "project_code": project_code.strip() if project_code else "",
        # 调试：热重载时刻，用于确认是否加载了最新代码
        "debug_hot_reload": {
            "checker_module_file": getattr(check_instance, '_dbg_file', 'N/A'),
            "checker_loaded_at": getattr(check_instance, '_dbg_time', 0),
            "checker_loaded_iso": getattr(check_instance, '_dbg_iso', 'N/A'),
        },
    }

    return summary


@app.get("/api/download/{filename}")
async def download_report(filename: str):
    """下载质检报告"""
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(file_path, filename=filename)


@app.get("/api/config")
async def get_config():
    """获取当前AI配置"""
    config = _load_config()
    # 返回时隐藏API密钥中间部分
    api_key = config.get("api_key", "")
    if api_key and len(api_key) > 8:
        masked = api_key[:4] + "*" * (len(api_key) - 8) + api_key[-4:]
    else:
        masked = api_key
    return {
        "provider": config.get("provider", "auto"),
        "api_key_display": masked,
        "api_key_set": bool(api_key),
        "model": config.get("model", ""),
        "base_url": config.get("base_url", ""),
        "enable_ai_check": config.get("enable_ai_check", True),
    }


def _launcher_config_payload() -> dict:
    config = load_json(LAUNCHER_CONFIG_FILE)
    update_source = str(config.get("update_source", "")).strip()
    source_path = os.path.expandvars(update_source) if update_source else ""
    source_accessible = bool(source_path and os.path.isdir(source_path))
    latest = {}
    if source_accessible:
        latest = load_json(os.path.join(source_path, "latest.json"))
    return {
        "update_source": update_source,
        "source_accessible": source_accessible,
        "current_version": APP_VERSION,
        "latest_version": latest.get("version", ""),
        "data_dir": str(ROOT_DIR),
        "output_dir": str(USER_OUTPUT_DIR),
        "cache_dir": str(CACHE_DIR),
        "log_dir": str(LOG_DIR),
    }


@app.get("/api/launcher-config")
async def get_launcher_config():
    """读取自动更新路径和当前本地数据目录。"""
    return _launcher_config_payload()


@app.post("/api/launcher-config")
async def save_launcher_config(req: LauncherConfigRequest):
    """保存共享版本文件夹。路径可暂时不可访问，但会明确返回状态。"""
    update_source = os.path.expandvars(req.update_source.strip())
    config = load_json(LAUNCHER_CONFIG_FILE)
    config["update_source"] = update_source
    save_json(LAUNCHER_CONFIG_FILE, config)
    payload = _launcher_config_payload()
    payload["message"] = (
        "共享版本文件夹已保存，连接正常"
        if payload["source_accessible"]
        else "路径已保存，但当前无法访问；仍可继续使用本地版本"
    )
    return payload


@app.post("/api/config")
async def save_config(req: AIConfigRequest):
    """保存AI配置"""
    _save_config(req.dict())
    return {"message": "配置保存成功", "restart_required": False}


@app.post("/api/config/test")
async def test_config(req: AIConfigRequest):
    """测试AI配置连通性"""
    # 如果前端传的api_key为空或包含脱敏标记***，使用已存储的密钥
    test_api_key = req.api_key
    if not test_api_key or "***" in test_api_key:
        stored = _load_config()
        test_api_key = stored.get("api_key", "")
    if not test_api_key:
        return {"success": False, "message": "未配置API密钥，请填写后测试"}

    # 临时设置环境变量进行测试
    test_env = {
        "AI_PROVIDER": req.provider,
        "AI_API_KEY": test_api_key,
        "AI_MODEL": req.model,
    }
    # base_url为空时移除环境变量，让SDK用默认地址
    if req.base_url:
        test_env["AI_BASE_URL"] = req.base_url

    old_env = {}
    for k, v in test_env.items():
        old_env[k] = os.environ.get(k, "")
        os.environ[k] = v

    try:
        # 重新加载config模块以获取最新的环境变量值
        import importlib
        import src.config as cfg
        importlib.reload(cfg)
        import src.ai_client as ac
        importlib.reload(ac)
        from src.ai_client import create_ai_client
        client = create_ai_client()
        if client is None:
            return {"success": False, "message": "未配置API密钥"}
        if not client.model:
            return {"success": False, "message": "未配置模型名称，请填写模型名称"}
        response = client.call_api("Hi")
        if response and response.strip():
            return {"success": True, "message": "连接成功！模型响应正常。"}
        return {"success": False, "message": "连接失败：模型返回空响应"}
    except Exception as e:
        return {"success": False, "message": f"连接失败: {str(e)}"}
    finally:
        # 恢复环境变量
        for k, v in old_env.items():
            if v:
                os.environ[k] = v
            elif k in os.environ:
                del os.environ[k]


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
