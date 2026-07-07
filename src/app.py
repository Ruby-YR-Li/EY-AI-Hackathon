"""Streamlit UI for the expense workpaper Review Assistant — v2.

Pure-upload workflow with real-time execution logging, rule toggles,
and LLM connectivity test. Design system follows the fixed-asset QC agent.
"""

from __future__ import annotations

import tempfile
import time
from dataclasses import asdict
from pathlib import Path

import streamlit as st

from expense_review_engine import (
    DEFAULT_FILES,
    ReviewNote,
    ReviewSummary,
    classify_workpaper,
    notes_to_csv,
    notes_to_rows,
    review_files,
)
from export_review_package import export_review_package
from llm_assistant import LLMConfig, test_connection

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
    ("tb_bkd_match", "TB/BKD 科目比对"),
    ("gl_scan", "序时账关键词扫描"),
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
    ("tb_bkd_match", "科目余额表 / BKD 比对"),
    ("gl_scan", "序时账关键词扫描"),
]


# ── CSS ──────────────────────────────────────────────────────────
def inject_style() -> None:
    st.markdown("""
    <style>
    :root {
        --ey-black: #111111; --ey-yellow: #ffe600;
        --ey-gray-100: #f5f5f5; --ey-gray-200: #dadde2;
        --ey-gray-300: #d0d5dd; --ey-gray-500: #667085;
        --ey-gray-700: #475467; --ey-gray-900: #242424;
        --risk-high: #b42318; --risk-medium: #b54708; --risk-low: #175cd3;
    }
    .main .block-container { padding-top: 1rem; max-width: 1420px; }
    .qc-topbar {
        background: var(--ey-black); border-left: 8px solid var(--ey-yellow);
        color: #fff; padding: 18px 22px; margin-bottom: 16px;
    }
    .qc-topbar h1 { margin: 0; font-size: 1.42rem; font-weight: 650; }
    .qc-topbar p { margin: 6px 0 0; color: #d6d6d6; font-size: 0.9rem; }
    .qc-file-header {
        border: 1px solid var(--ey-gray-200); border-left: 6px solid var(--ey-yellow);
        border-radius: 6px; background: #fff; padding: 14px 18px; margin: 8px 0 18px;
    }
    .qc-file-header h2 { margin: 0; color: var(--ey-black); font-size: 1.35rem; font-weight: 700; }
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
    .qc-card-label { color: var(--ey-gray-500); font-size: 0.8rem; }
    .qc-card-value { color: var(--ey-black); font-size: 1.72rem; font-weight: 750; line-height: 1.18; margin-top: 4px; }
    .qc-card-note { color: var(--ey-gray-500); font-size: 0.76rem; margin-top: 4px; }
    .qc-section-title { color: var(--ey-black); font-size: 1.15rem; font-weight: 700; margin: 10px 0 4px; }
    .qc-section-caption { color: var(--ey-gray-500); font-size: 0.86rem; margin: 0 0 10px; }
    .llm-status-ok {
        display: inline-block; padding: 4px 10px; border-radius: 999px;
        background: #ecfdf3; color: #067647; font-weight: 700; font-size: 0.8rem;
    }
    .llm-status-err {
        display: inline-block; padding: 4px 10px; border-radius: 999px;
        background: #fff3f0; color: #b42318; font-weight: 700; font-size: 0.8rem;
    }
    .llm-status-off {
        display: inline-block; padding: 4px 10px; border-radius: 999px;
        background: #f5f5f5; color: #667085; font-weight: 700; font-size: 0.8rem;
    }
    div.stButton > button[kind="primary"] {
        background: var(--ey-black); color: #fff; border: 1px solid var(--ey-black);
        border-radius: 4px;
    }
    div.stButton > button[kind="primary"]:hover {
        background: var(--ey-gray-900); border-color: var(--ey-gray-900); color: var(--ey-yellow);
    }
    div.stDownloadButton > button[kind="primary"] {
        background: var(--ey-black); color: #fff; border: 1px solid var(--ey-black);
        border-radius: 4px;
    }
    div.stDownloadButton > button[kind="primary"]:hover {
        background: var(--ey-gray-900); border-color: var(--ey-gray-900); color: var(--ey-yellow);
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


# ── topbar ───────────────────────────────────────────────────────
def render_topbar() -> None:
    st.markdown("""
    <div class="qc-topbar">
      <h1>费用底稿 AI Review 助手</h1>
      <p>上传底稿 → 自动 Review → 输出可追溯报告 · 15 条规则覆盖</p>
    </div>
    """, unsafe_allow_html=True)


# ── sidebar: upload ──────────────────────────────────────────────
def render_sidebar_upload() -> tuple[list[Path], Path | None, Path | None, Path | None]:
    """返回 (all_uploaded, main_path, tod_path, tb_path, gl_path)。"""
    st.sidebar.header("📂 底稿上传")
    uploaded = st.sidebar.file_uploader(
        "选择底稿文件（一套最多2个 .xlsx）",
        type=["xlsx"],
        accept_multiple_files=True,
        help="上传主底稿（VC&VD）和 TOD 底稿，系统自动识别分类。同时支持上传科目余额表和序时账作为辅助资料。",
    )

    main_path: Path | None = None
    tod_path: Path | None = None
    tb_path: Path | None = None
    gl_path: Path | None = None
    all_paths: list[Path] = []

    if uploaded:
        for uf in uploaded:
            tmp = _save_upload(uf, ".xlsx")
            if tmp:
                all_paths.append(tmp)
                file_type = classify_workpaper(tmp)
                if file_type == "main" and not main_path:
                    main_path = tmp
                elif file_type == "tod" and not tod_path:
                    tod_path = tmp
                # 科目余额表和序时账通过文件名特征识别
                name_lower = uf.name.lower()
                if "科目余额" in name_lower or "tb" in name_lower:
                    tb_path = tmp
                if "序时账" in name_lower or "gl" in name_lower or "明细账" in name_lower:
                    gl_path = tmp

        # 显示识别结果
        if main_path:
            st.sidebar.success(f"✅ 主底稿")
        else:
            st.sidebar.warning("⚠️ 未识别到主底稿")

        if tod_path:
            st.sidebar.success(f"✅ TOD 底稿")
        else:
            st.sidebar.info("ℹ️ 未上传 TOD 底稿（TOD 规则将跳过）")

        if tb_path:
            st.sidebar.success("✅ 科目余额表")
        if gl_path:
            st.sidebar.success("✅ 序时账")
    else:
        st.sidebar.info("上传 .xlsx 文件以开始 Review。")

    return all_paths, main_path, tod_path, tb_path, gl_path


# ── sidebar: rules ───────────────────────────────────────────────
def render_sidebar_rules() -> set[str]:
    st.sidebar.markdown("---")
    st.sidebar.header("🔧 检查规则")

    col_a, col_b = st.sidebar.columns(2)
    with col_a:
        if st.button("全选", use_container_width=True):
            st.session_state["rule_selections"] = {k for k, _ in RULE_LABELS}
    with col_b:
        if st.button("清空", use_container_width=True):
            st.session_state["rule_selections"] = set()

    if "rule_selections" not in st.session_state:
        st.session_state["rule_selections"] = {k for k, _ in RULE_LABELS}

    selected = set()
    for key, label in RULE_LABELS:
        if st.sidebar.checkbox(
            label, value=key in st.session_state["rule_selections"],
            key=f"rule_{key}",
        ):
            selected.add(key)
    st.session_state["rule_selections"] = selected
    return selected


# ── sidebar: LLM ─────────────────────────────────────────────────
def render_sidebar_llm() -> LLMConfig | None:
    st.sidebar.markdown("---")
    st.sidebar.header("🤖 LLM 增强（可选）")

    use_llm = st.sidebar.checkbox("启用 LLM 辅助判断", value=False)
    if not use_llm:
        st.sidebar.caption('<span class="llm-status-off">LLM 未启用</span>', unsafe_allow_html=True)
        return None

    api_key = st.sidebar.text_input(
        "DeepSeek API Key",
        type="password",
        placeholder="sk-...",
        help="留空则从 DEEPSEEK_API_KEY 环境变量读取。",
    )
    model = st.sidebar.text_input("模型", value="deepseek-chat", help="默认 deepseek-chat")
    timeout = st.sidebar.number_input("超时(秒)", value=20, min_value=5, max_value=60)

    config = LLMConfig(
        enabled=True,
        api_key=api_key,
        model=model,
        timeout_seconds=int(timeout),
    )

    if st.sidebar.button("🔍 测试连接", use_container_width=True):
        with st.sidebar:
            with st.spinner("测试中..."):
                result = test_connection(config)
            if result["ok"]:
                st.success(f"🟢 已连接 · {result['model']} · {result['latency_ms']}ms")
            else:
                st.error(f"🔴 连接失败: {result['error'][:120]}")

    # 显示当前 LLM 状态
    if api_key:
        st.sidebar.caption("已填写 API Key · 点击「测试连接」验证")
    else:
        st.sidebar.caption("未填写 API Key · 将尝试环境变量")

    return config


# ── main: results view ───────────────────────────────────────────
def render_results(
    notes: list[ReviewNote],
    summary: ReviewSummary,
    elapsed_s: float,
    llm_status_note: str = "",
    bundle: dict | None = None,
) -> None:
    """展示 Review 完成后的结果区域。"""
    rows = notes_to_rows(notes)

    # file header
    high_count = summary.high_count
    worst = "High" if high_count > 0 else ("Medium" if summary.medium_count > 0 else "Low")
    st.markdown(f"""
    <div class="qc-file-header">
      <h2>费用底稿 Review · Findings {summary.total_notes} 条</h2>
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
        _render_card("规则数", summary.rules_run, "已执行的检查规则", "other")

    st.caption(f"程序要求 {summary.program_rows} 行 · SOP Checklist {summary.checklist_rows} 检查点")

    # downloads
    st.subheader("交付物下载")
    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button(
            "📥 下载 Review Notes (CSV)",
            data=notes_to_csv(notes),
            file_name="expense_review_notes.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with dl2:
        if bundle and bundle.get("annotated_workbook"):
            ann_bytes = bundle["annotated_workbook"].read_bytes()
            st.download_button(
                "📥 下载标注底稿副本 (.xlsx)",
                data=ann_bytes,
                file_name=bundle["annotated_workbook"].name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        else:
            st.download_button(
                "📥 下载标注底稿副本",
                data=b"",
                file_name="annotated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                disabled=True,
            )
    with dl3:
        if bundle and bundle.get("html_report"):
            html_bytes = bundle["html_report"].read_bytes()
            st.download_button(
                "📥 下载 HTML 报告",
                data=html_bytes,
                file_name=bundle["html_report"].name,
                mime="text/html",
                use_container_width=True,
            )
        else:
            st.download_button(
                "📥 下载 HTML 报告",
                data=b"",
                file_name="review_report.html",
                mime="text/html",
                use_container_width=True,
                disabled=True,
            )

    # tabs
    high_rows = [r for r in rows if r["risk_level"] == "High"]
    medium_rows = [r for r in rows if r["risk_level"] == "Medium"]
    low_rows = [r for r in rows if r["risk_level"] == "Low"]

    tab_all, tab_high, tab_medium, tab_log = st.tabs([
        f"全部 ({len(rows)})",
        f"🔴 High ({len(high_rows)})",
        f"🟡 Medium ({len(medium_rows)})",
        "📋 规则执行日志",
    ])

    with tab_all:
        render_notes_table(rows)
    with tab_high:
        render_notes_table(high_rows)
    with tab_medium:
        render_notes_table(medium_rows)
    with tab_log:
        st.caption("以下为本次执行的规则步骤日志。")
        if "run_log" in st.session_state:
            for entry in st.session_state["run_log"]:
                st.write(entry)

    # diagnostics expander
    with st.expander("系统诊断", expanded=False):
        st.write(f"总耗时: {elapsed_s:.1f}s")
        st.write(f"LLM 状态: {llm_status_note or '未启用'}")
        st.write(f"Review Notes: {summary.total_notes} (H:{summary.high_count} M:{summary.medium_count} L:{summary.low_count})")
        st.write(f"规则数: {summary.rules_run} · 涉及 Sheet: {summary.sheets_impacted}")


# ── main ─────────────────────────────────────────────────────────
def main() -> None:
    inject_style()
    render_topbar()

    # sidebar
    all_paths, main_path, tod_path, tb_path, gl_path = render_sidebar_upload()
    enabled_rules = render_sidebar_rules()
    llm_config = render_sidebar_llm()

    # run button in sidebar
    st.sidebar.markdown("---")
    ready = main_path is not None and len(enabled_rules) > 0
    run_clicked = st.sidebar.button(
        "🚀 开始 Review", type="primary", use_container_width=True, disabled=not ready,
    )

    # init session state
    if "results" not in st.session_state:
        st.session_state["results"] = None

    # show previous results if available
    if st.session_state["results"] is not None and not run_clicked:
        notes, summary, elapsed_s, llm_note, bundle = st.session_state["results"]
        render_results(notes, summary, elapsed_s, llm_note, bundle)
        if main_path:
            st.divider()
            st.caption("上传新文件后点击「开始 Review」可重新运行。")
        return

    # idle state
    if not run_clicked:
        st.info("👈 在左侧边栏上传底稿、配置规则，然后点击「开始 Review」。")
        with st.expander("📖 使用说明", expanded=True):
            st.markdown("""
            1. **上传底稿**：选择主底稿（VC&VD）和 TOD 底稿（.xlsx），系统自动识别分类
            2. **配置规则**：勾选需要执行的检查规则（默认全选）
            3. **LLM 增强**（可选）：填写 DeepSeek API Key 并点击「测试连接」
            4. **开始 Review**：点击按钮，查看逐步执行日志
            5. **下载结果**：CSV 格式，可复制到底稿 Review Notes
            """)
        return

    # ── execute review ──
    run_log: list[str] = []
    t_start = time.perf_counter()

    with st.status("正在执行 Review...", expanded=True) as status:
        # Build file paths dict
        files: dict[str, object] = {"workpapers": [str(p) for p in all_paths]}
        if main_path:
            files["main_workpaper"] = str(main_path)
        if tod_path:
            files["tod_workpaper"] = str(tod_path)
        if tb_path:
            files["trial_balance"] = str(tb_path)
        if gl_path:
            files["general_ledger"] = str(gl_path)

        step_idx = 0
        total_steps = len(RULE_RUN_ORDER) + 2  # rules + read + generate

        st.write("📖 读取底稿结构...")
        step_idx += 1

        # Collect LLM statuses
        llm_statuses: list[dict] = []

        try:
            summary, notes = review_files(files, llm_config=llm_config)
        except Exception as exc:
            st.error(f"Review 运行失败：{exc}")
            status.update(label="Review 失败", state="error")
            return

        # Generate export bundle
        bundle = None
        try:
            st.write("📊 生成报告与标注底稿...")
            bundle = export_review_package(notes, summary, files=files)
        except Exception as exc:
            st.warning(f"导出生成部分失败：{exc}（Review Notes 仍可用）")

        # Build a step-by-step log based on the actual notes
        notes_by_rule: dict[str, list[ReviewNote]] = {}
        for n in notes:
            # Map note source → rule key
            rule_key = "other"
            source = n.source or ""
            if "Lead" in source or "lead" in source.lower():
                rule_key = "lead_basics"
            elif "汇总" in source or "程序" in source:
                rule_key = "summary_execution" if "执行" in source else "legal_consistency"
            elif "特殊费用" in source or "LEGAL" in n.issue_type:
                rule_key = "special_expenses"
            elif "BKD" in source and "Notes" in source:
                rule_key = "bkd_notes"
            elif "BKD" in source and "货币" in source:
                rule_key = "bkd_currency"
            elif "上期" in source:
                rule_key = "bkd_prior_year"
            elif "截止" in source:
                rule_key = "cutoff"
            elif "TOD" in source and "总体" in source:
                rule_key = "tod_population"
            elif "TOD" in source and "关键" in source:
                rule_key = "tod_key_items"
            elif "TOD" in source and "负值" in source:
                rule_key = "tod_negative"
            elif "TOD" in source or "样本池" in source:
                rule_key = "tod"
            elif "缺失" in source or "Sheet" in source:
                rule_key = "missing_sheets"
            elif "TB" in source or "科目余额" in source:
                rule_key = "tb_bkd_match"
            elif "序时账" in source:
                rule_key = "gl_scan"
            elif "法律" in source:
                rule_key = "legal_consistency"
            notes_by_rule.setdefault(rule_key, []).append(n)

        for rule_key, step_label in RULE_RUN_ORDER:
            if rule_key not in enabled_rules:
                continue
            step_idx += 1
            count = len(notes_by_rule.get(rule_key, []))
            icon = "🔴" if any(n.risk_level == "High" for n in notes_by_rule.get(rule_key, [])) else ("🟡" if count > 0 else "✅")
            run_log.append(f"{icon} [{step_idx}/{total_steps}] {step_label} · {count} 条")
            st.write(run_log[-1])

        st.write("📊 生成报告...")
        run_log.append(f"📊 生成报告完成")

        status.update(label=f"Review 完成 · {summary.total_notes} 条 Findings", state="complete")

    elapsed_s = time.perf_counter() - t_start
    llm_note = "LLM 未启用"
    if llm_config and llm_config.is_available:
        llm_note = "LLM 已启用"

    # persist results
    st.session_state["results"] = (notes, summary, elapsed_s, llm_note, bundle)
    st.session_state["run_log"] = run_log

    # render
    render_results(notes, summary, elapsed_s, llm_note, bundle)


if __name__ == "__main__":
    main()
