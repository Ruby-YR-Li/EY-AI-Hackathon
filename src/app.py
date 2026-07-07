"""Streamlit UI for the expense workpaper Review Assistant — v3.

Main area: upload + run + results. Sidebar: rules + LLM settings only.
Design follows the fixed-asset QC agent pattern.
"""

from __future__ import annotations

import tempfile
import time
from pathlib import Path

import streamlit as st

from expense_review_engine import (
    ReviewNote,
    ReviewSummary,
    classify_workpaper,
    notes_to_csv,
    notes_to_rows,
    review_files,
)
from export_review_package import export_review_package
from llm_assistant import LLMConfig, list_models, test_connection

# ── page config ──────────────────────────────────────────────────
st.set_page_config(page_title="费用底稿 AI Review 助手", layout="wide")

# ── rules registry ───────────────────────────────────────────────
RULE_LABELS: list[tuple[str, str]] = [
    ("lead_basics", "Lead 基础信息"),
    ("summary_execution", "汇总页执行状态"),
    ("special_expenses", "特殊费用识别"),
    ("legal_consistency", "法律费用一致性"),
    ("bkd_notes", "BKD Notes 检查"),
    ("cutoff", "截止性测试"),
    ("tod", "TOD 样本池"),
    ("bkd_currency", "BKD 货币/单位"),
    ("bkd_prior_year", "上期审定数检查"),
    ("tod_population", "TOD 样本总体范围"),
    ("tod_key_items", "TOD 关键项(KI)"),
    ("tod_negative", "TOD 负值处理"),
    ("missing_sheets", "缺失标准Sheet"),
]

RULE_RUN_ORDER: list[tuple[str, str]] = [
    ("lead_basics", "Lead 基础信息检查"),
    ("summary_execution", "汇总页执行状态检查"),
    ("special_expenses", "特殊费用识别"),
    ("legal_consistency", "法律费用程序一致性"),
    ("bkd_notes", "BKD Notes 检查"),
    ("bkd_currency", "BKD 货币/单位检查"),
    ("bkd_prior_year", "上期审定数检查"),
    ("cutoff", "截止性测试"),
    ("tod", "TOD 样本池检查"),
    ("tod_population", "TOD 样本总体范围"),
    ("tod_key_items", "TOD 关键项(KI)"),
    ("tod_negative", "TOD 负值处理"),
    ("missing_sheets", "缺失标准 Sheet"),
]


# ── CSS ──────────────────────────────────────────────────────────
def inject_style() -> None:
    st.markdown("""
    <style>
    :root {
        --ey-black: #111111; --ey-yellow: #ffe600;
        --ey-gray-100: #f5f5f5; --ey-gray-200: #dadde2;
        --ey-gray-300: #d0d5dd; --ey-gray-400: #98a2b3;
        --ey-gray-500: #667085; --ey-gray-700: #475467;
        --ey-gray-900: #242424;
        --risk-high: #b42318; --risk-medium: #b54708; --risk-low: #175cd3;
    }
    .main .block-container { padding-top: 1rem; max-width: 1420px; }
    .qc-topbar {
        background: var(--ey-black); border-left: 8px solid var(--ey-yellow);
        color: #fff; padding: 18px 22px; margin-bottom: 20px;
    }
    .qc-topbar h1 { margin: 0; font-size: 1.42rem; font-weight: 650; }
    .qc-topbar p { margin: 6px 0 0; color: #d6d6d6; font-size: 0.9rem; }
    .qc-file-header {
        border: 1px solid var(--ey-gray-200); border-left: 6px solid var(--ey-yellow);
        border-radius: 6px; background: #fff; padding: 14px 18px; margin: 8px 0 18px;
    }
    .qc-file-header h2 { margin: 0; color: var(--ey-black); font-size: 1.3rem; font-weight: 700; }
    .qc-file-header p { margin: 8px 0 0; color: var(--ey-gray-500); font-size: 0.88rem; }
    .qc-card {
        border: 1px solid var(--ey-gray-200);
        border-left: 5px solid var(--accent, var(--ey-gray-700));
        border-radius: 6px; background: #fff; padding: 12px 14px;
        min-height: 78px; box-shadow: 0 1px 2px rgba(16,24,40,0.04);
    }
    .qc-card-high { --accent: var(--risk-high); }
    .qc-card-warn { --accent: var(--risk-medium); }
    .qc-card-pass { --accent: var(--risk-low); }
    .qc-card-other { --accent: var(--ey-gray-700); }
    .qc-card-label { color: var(--ey-gray-500); font-size: 0.78rem; }
    .qc-card-value { color: var(--ey-black); font-size: 1.65rem; font-weight: 750; margin-top: 4px; }
    .qc-card-note { color: var(--ey-gray-500); font-size: 0.74rem; margin-top: 4px; }

    /* upload zone */
    .upload-zone {
        border: 2px dashed var(--ey-gray-300); border-radius: 10px;
        background: #fafafa; padding: 32px 24px; text-align: center;
        margin-bottom: 16px;
    }
    .upload-zone.has-files {
        border-color: var(--ey-yellow); border-style: solid; background: #fffef5;
    }
    .upload-icon { font-size: 2.4rem; margin-bottom: 8px; }
    .upload-title { font-size: 1.1rem; font-weight: 650; color: var(--ey-black); margin-bottom: 4px; }
    .upload-hint { font-size: 0.85rem; color: var(--ey-gray-500); }

    /* file status row */
    .file-status {
        display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 18px;
    }
    .file-badge {
        display: inline-flex; align-items: center; gap: 6px;
        padding: 8px 14px; border-radius: 6px; font-size: 0.88rem; font-weight: 600;
    }
    .file-badge.ok { background: #ecfdf3; color: #067647; border: 1px solid #abefc6; }
    .file-badge.warn { background: #fff7ed; color: #b54708; border: 1px solid #fed7aa; }
    .file-badge.info { background: #eff8ff; color: #175cd3; border: 1px solid #b2ddff; }

    /* run button area */
    .run-area {
        background: #fafafa; border: 1px solid var(--ey-gray-200); border-radius: 8px;
        padding: 16px 20px; margin-bottom: 20px;
    }
    .run-area .ready-text { color: var(--ey-black); font-size: 0.95rem; font-weight: 600; }
    .run-area .hint-text { color: var(--ey-gray-500); font-size: 0.82rem; margin-top: 4px; }

    div.stButton > button[kind="primary"] {
        background: var(--ey-black); color: #fff; border: 1px solid var(--ey-black);
        border-radius: 4px; font-size: 1rem; padding: 0.5rem 1.2rem;
    }
    div.stButton > button[kind="primary"]:hover {
        background: var(--ey-gray-900); color: var(--ey-yellow);
    }
    div.stDownloadButton > button[kind="primary"] {
        background: var(--ey-black); color: #fff; border: 1px solid var(--ey-black);
        border-radius: 4px;
    }
    div.stDownloadButton > button[kind="primary"]:hover {
        background: var(--ey-gray-900); color: var(--ey-yellow);
    }
    </style>
    """, unsafe_allow_html=True)


# ── helpers ──────────────────────────────────────────────────────
def _save_upload(uploaded_file, suffix: str) -> Path | None:
    if uploaded_file is None:
        return None
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return Path(tmp.name)


def _render_card(label: str, value: object, note: str = "", tone: str = "other") -> None:
    st.markdown(f"""
    <div class="qc-card qc-card-{tone}">
      <div class="qc-card-label">{label}</div>
      <div class="qc-card-value">{value}</div>
      <div class="qc-card-note">{note}</div>
    </div>
    """, unsafe_allow_html=True)


def render_notes_table(rows: list[dict[str, str]]) -> None:
    if not rows:
        st.success("未发现 Review Notes。")
        return
    st.dataframe(rows, use_container_width=True, hide_index=True)


# ── sidebar ──────────────────────────────────────────────────────
def render_sidebar() -> tuple[set[str], LLMConfig | None]:
    """侧边栏仅放配置项：检查项 + LLM 设置。"""
    # ── rules ──
    st.sidebar.header("🔧 检查项")
    col_a, col_b = st.sidebar.columns(2)
    with col_a:
        if st.button("全选", use_container_width=True, key="sb_all"):
            st.session_state["rule_selections"] = {k for k, _ in RULE_LABELS}
    with col_b:
        if st.button("清空", use_container_width=True, key="sb_none"):
            st.session_state["rule_selections"] = set()

    if "rule_selections" not in st.session_state:
        st.session_state["rule_selections"] = {k for k, _ in RULE_LABELS}

    selected: set[str] = set()
    for key, label in RULE_LABELS:
        if st.sidebar.checkbox(label, value=key in st.session_state["rule_selections"], key=f"rule_{key}"):
            selected.add(key)
    st.session_state["rule_selections"] = selected
    st.sidebar.caption(f"已选 {len(selected)} 项检查")

    st.sidebar.markdown("---")

    # ── LLM ──
    st.sidebar.header("🤖 AI 增强（可选）")
    use_llm = st.sidebar.checkbox("启用 AI 辅助判断", value=False)
    if not use_llm:
        st.sidebar.caption("AI 未启用 · 将使用内置检查能力")
        return selected, None

    models = list_models()
    model_key = st.sidebar.selectbox(
        "模型",
        options=[m["key"] for m in models],
        format_func=lambda k: next((m["label"] for m in models if m["key"] == k), k),
        index=0,
    )
    api_key = st.sidebar.text_input("API Key", type="password", placeholder="留空则从环境变量读取")
    timeout = st.sidebar.number_input("超时(秒)", value=20, min_value=5, max_value=60)

    config = LLMConfig(enabled=True, api_key=api_key, model_key=model_key, timeout_seconds=int(timeout))

    # 测试连接 + 状态持久化
    if "llm_test_result" not in st.session_state:
        st.session_state["llm_test_result"] = None

    if st.sidebar.button("🔍 测试连接", use_container_width=True):
        with st.spinner("测试中..."):
            result = test_connection(config)
        st.session_state["llm_test_result"] = result

    test = st.session_state.get("llm_test_result")
    if test is not None:
        if test["ok"]:
            st.sidebar.success(f"🟢 已连接 · {config.model} · {test['latency_ms']}ms")
        else:
            st.sidebar.error(f"🔴 {test['error'][:100]}")
    elif api_key:
        st.sidebar.caption("已填写 API Key · 点击测试连接验证")
    else:
        st.sidebar.caption("未填写 API Key · 将从环境变量自动读取")

    return selected, config


# ── main area: upload + run ──────────────────────────────────────
def _detect_files(uploaded) -> tuple[list[Path], Path | None, Path | None]:
    """解析上传文件并返回路径。"""
    main_path = tod_path = None
    all_paths: list[Path] = []

    for uf in uploaded:
        tmp = _save_upload(uf, ".xlsx")
        if not tmp:
            continue
        all_paths.append(tmp)
        ft = classify_workpaper(tmp)
        if ft == "main" and not main_path:
            main_path = tmp
        elif ft == "tod" and not tod_path:
            tod_path = tmp

    return all_paths, main_path, tod_path


def render_upload_zone() -> tuple[list[Path], Path | None, Path | None]:
    """主区域的上传区。"""
    uploaded = st.file_uploader(
        "选择底稿文件",
        type=["xlsx"],
        accept_multiple_files=True,
        label_visibility="collapsed",
        key="main_uploader",
    )

    all_paths, main_path, tod_path = _detect_files(uploaded or [])

    if not uploaded:
        st.markdown("""
        <div class="upload-zone">
          <div class="upload-icon">📂</div>
          <div class="upload-title">上传费用底稿文件</div>
          <div class="upload-hint">拖拽 .xlsx 文件到此处，或点击「Browse files」选择<br>
          一套底稿 = 主底稿（VC&VD）+ TOD底稿（可选），系统自动识别分类</div>
        </div>
        """, unsafe_allow_html=True)
        return all_paths, main_path, tod_path

    # 显示文件识别状态
    st.markdown('<div class="file-status">', unsafe_allow_html=True)

    if main_path:
        st.success(f"✅ 主底稿")
    else:
        st.warning("⚠️ 未识别到主底稿（需包含 Lead/VC/VD/汇总等 Sheet）")

    if tod_path:
        st.success(f"✅ TOD 底稿")
    else:
        st.info("ℹ️ 未上传 TOD 底稿 · TOD 相关检查将跳过")

    st.markdown('</div>', unsafe_allow_html=True)
    return all_paths, main_path, tod_path


def render_run_button(main_path: Path | None, enabled_rules: set[str], llm_config: LLMConfig | None) -> bool:
    """渲染「开始 Review」按钮区域，返回是否点击。"""
    ready = main_path is not None and len(enabled_rules) > 0

    status_parts = []
    if main_path:
        status_parts.append("✅ 底稿已加载")
    else:
        status_parts.append("❌ 请上传底稿")
    if llm_config and llm_config.is_available:
        test = st.session_state.get("llm_test_result")
        if test and test["ok"]:
            status_parts.append("🤖 AI 引擎已就绪")
        else:
            status_parts.append("🤖 AI 引擎待验证")
    else:
        status_parts.append("🤖 AI 引擎待配置")

    st.markdown(f"""
    <div class="run-area">
      <div class="ready-text">{' · '.join(status_parts)}</div>
      <div class="hint-text">点击下方按钮开始自动 Review，执行过程会实时展示</div>
    </div>
    """, unsafe_allow_html=True)

    return st.button(
        "🚀 开始 Review", type="primary", use_container_width=True, disabled=not ready,
    )


# ── main area: results ───────────────────────────────────────────
def render_results(
    notes: list[ReviewNote],
    summary: ReviewSummary,
    elapsed_s: float,
    llm_status_note: str = "",
    bundle: dict | None = None,
) -> None:
    rows = notes_to_rows(notes)
    high_count = summary.high_count
    worst = "High" if high_count > 0 else ("Medium" if summary.medium_count > 0 else "Low")

    st.markdown(f"""
    <div class="qc-file-header">
      <h2>Review 完成 · Findings {summary.total_notes} 条</h2>
      <p>最高风险：{worst} · 耗时 {elapsed_s:.1f}s · {llm_status_note}</p>
    </div>
    """, unsafe_allow_html=True)

    # metric cards
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        _render_card("Review Notes", summary.total_notes, "发现问题总数", "other")
    with c2:
        _render_card("High", summary.high_count, "高优先级", "high")
    with c3:
        _render_card("Medium", summary.medium_count, "中等风险", "warn")
    with c4:
        _render_card("Low", summary.low_count, "低优先级", "pass")
    with c5:
        _render_card("涉及 Sheet", summary.sheets_impacted, "跨文件+Sheet", "other")
    with c6:
        _render_card("检查项", summary.rules_run, "已完成的检查维度", "other")

    st.caption(f"程序要求 {summary.program_rows} 行 · SOP Checklist {summary.checklist_rows} 检查点")

    # downloads
    st.subheader("📥 交付物下载")
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button(
            "下载 Review Notes (CSV)",
            data=notes_to_csv(notes),
            file_name="expense_review_notes.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with dl2:
        if bundle and bundle.get("annotated_workbook"):
            ann_bytes = bundle["annotated_workbook"].read_bytes()
            st.download_button(
                "下载标注底稿副本 (.xlsx)",
                data=ann_bytes,
                file_name=bundle["annotated_workbook"].name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.button("标注底稿不可用", disabled=True, use_container_width=True)
    with dl3:
        if bundle and bundle.get("html_report"):
            html_bytes = bundle["html_report"].read_bytes()
            st.download_button(
                "下载 HTML 报告",
                data=html_bytes,
                file_name=bundle["html_report"].name,
                mime="text/html",
                use_container_width=True,
            )
        else:
            st.button("HTML报告不可用", disabled=True, use_container_width=True)

    # tabs
    high_rows = [r for r in rows if r["risk_level"] == "High"]
    medium_rows = [r for r in rows if r["risk_level"] == "Medium"]

    tab_all, tab_high, tab_medium, tab_log = st.tabs([
        f"全部 ({len(rows)})",
        f"🔴 High ({len(high_rows)})",
        f"🟡 Medium ({len(medium_rows)})",
        "📋 执行日志",
    ])

    with tab_all:
        render_notes_table(rows)
    with tab_high:
        render_notes_table(high_rows)
    with tab_medium:
        render_notes_table(medium_rows)
    with tab_log:
        st.caption("以下为本次 Review 的执行步骤。")
        if "run_log" in st.session_state:
            for entry in st.session_state["run_log"]:
                st.write(entry)

    with st.expander("🔍 系统诊断", expanded=False):
        st.write(f"总耗时: {elapsed_s:.1f}s")
        st.write(f"LLM 状态: {llm_status_note or '未启用'}")
        st.write(f"Review Notes: {summary.total_notes} (High:{summary.high_count} Medium:{summary.medium_count} Low:{summary.low_count})")


# ── main ─────────────────────────────────────────────────────────
def main() -> None:
    inject_style()
    st.markdown("""
    <div class="qc-topbar">
      <h1>费用底稿 AI Review 助手</h1>
      <p>上传底稿 → 自动 Review → 输出可追溯报告</p>
    </div>
    """, unsafe_allow_html=True)

    # ══ sidebar: settings only ══
    enabled_rules, llm_config = render_sidebar()

    # ══ main area: upload → run → results ══
    all_paths, main_path, tod_path = render_upload_zone()

    run_clicked = render_run_button(main_path, enabled_rules, llm_config)

    # init session state
    if "results" not in st.session_state:
        st.session_state["results"] = None

    # ── show previous results ──
    if st.session_state["results"] is not None and not run_clicked:
        notes, summary, elapsed_s, llm_note, bundle = st.session_state["results"]
        render_results(notes, summary, elapsed_s, llm_note, bundle)
        if all_paths:
            st.divider()
            st.caption("上传新文件后点击「开始 Review」可重新运行。")
        return

    # ── idle ──
    if not run_clicked:
        return

    # ── execute ──
    run_log: list[str] = []
    t_start = time.perf_counter()

    with st.status("正在执行 Review...", expanded=True) as status:
        files: dict[str, object] = {"workpapers": [str(p) for p in all_paths]}
        if main_path:
            files["main_workpaper"] = str(main_path)
        if tod_path:
            files["tod_workpaper"] = str(tod_path)

        step_idx = 0
        total_steps = len(RULE_RUN_ORDER) + 2

        st.write("📖 读取底稿结构...")

        try:
            summary, notes = review_files(files, llm_config=llm_config)
        except Exception as exc:
            st.error(f"Review 运行失败：{exc}")
            status.update(label="Review 失败", state="error")
            return

        # export
        bundle = None
        try:
            st.write("📊 生成报告与标注底稿...")
            bundle = export_review_package(notes, summary, files=files)
        except Exception as exc:
            st.warning(f"导出生成部分失败：{exc}")

        # build log
        notes_by_source: dict[str, list[ReviewNote]] = {}
        for n in notes:
            s = n.source or ""
            if "Lead" in s or "lead" in s.lower():
                k = "lead_basics"
            elif "汇总" in s or ("程序" in s and "执行" in s):
                k = "summary_execution"
            elif "法律" in s:
                k = "legal_consistency"
            elif "特殊费用" in s:
                k = "special_expenses"
            elif "BKD" in s and "Notes" in s:
                k = "bkd_notes"
            elif "BKD" in s and "货币" in s:
                k = "bkd_currency"
            elif "上期" in s:
                k = "bkd_prior_year"
            elif "截止" in s:
                k = "cutoff"
            elif "TOD" in s and "总体" in s:
                k = "tod_population"
            elif "TOD" in s and "关键" in s:
                k = "tod_key_items"
            elif "TOD" in s and "负值" in s:
                k = "tod_negative"
            elif "TOD" in s or "样本池" in s:
                k = "tod"
            elif "缺失" in s or "Sheet" in s:
                k = "missing_sheets"
            else:
                k = "other"
            notes_by_source.setdefault(k, []).append(n)

        for rule_key, step_label in RULE_RUN_ORDER:
            if rule_key not in enabled_rules:
                continue
            step_idx += 1
            count = len(notes_by_source.get(rule_key, []))
            icon = "🔴" if any(n.risk_level == "High" for n in notes_by_source.get(rule_key, [])) else ("🟡" if count > 0 else "✅")
            run_log.append(f"{icon} [{step_idx}/{total_steps}] {step_label} · {count} 条")
            st.write(run_log[-1])

        st.write("📊 生成报告...")
        run_log.append(f"📊 生成报告完成")
        status.update(label=f"Review 完成 · {summary.total_notes} 条 Findings", state="complete")

    elapsed_s = time.perf_counter() - t_start
    llm_note = "LLM 未启用"
    if llm_config and llm_config.is_available:
        llm_note = "LLM 已启用"

    st.session_state["results"] = (notes, summary, elapsed_s, llm_note, bundle)
    st.session_state["run_log"] = run_log

    render_results(notes, summary, elapsed_s, llm_note, bundle)


if __name__ == "__main__":
    main()
