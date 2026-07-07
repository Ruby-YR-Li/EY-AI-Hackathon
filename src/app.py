"""Streamlit UI for the expense workpaper Review Assistant."""

from __future__ import annotations

import tempfile
from dataclasses import asdict
from pathlib import Path

import streamlit as st

from expense_review_engine import DEFAULT_FILES, notes_to_csv, notes_to_rows, review_files


st.set_page_config(page_title="费用底稿 AI Review 助手", layout="wide")


def inject_style() -> None:
    st.markdown(
        """
        <style>
        :root {
            --ey-black: #111111;
            --ey-yellow: #ffe600;
            --ey-gray-100: #f5f5f5;
            --ey-gray-300: #d0d5dd;
            --ey-gray-700: #475467;
            --risk-high: #b42318;
            --risk-medium: #b54708;
            --risk-low: #175cd3;
        }
        .main .block-container { padding-top: 1rem; max-width: 1420px; }
        .topbar {
            background: var(--ey-black);
            border-left: 8px solid var(--ey-yellow);
            color: white;
            padding: 18px 22px;
            margin-bottom: 16px;
        }
        .topbar h1 { margin: 0; font-size: 1.5rem; letter-spacing: 0; }
        .topbar p { margin: 6px 0 0 0; color: #d6d6d6; font-size: 0.92rem; }
        .section-caption { color: var(--ey-gray-700); font-size: 0.88rem; margin-bottom: 8px; }
        div[data-testid="stMetric"] {
            background: white;
            border: 1px solid var(--ey-gray-300);
            border-left: 5px solid var(--ey-yellow);
            padding: 10px 12px;
            border-radius: 6px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def save_upload(uploaded_file, suffix: str) -> Path | None:
    if uploaded_file is None:
        return None
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return Path(tmp.name)


def file_selector() -> dict[str, Path]:
    st.sidebar.subheader("资料来源")
    use_default = st.sidebar.checkbox("使用当前案例资料", value=True)
    if use_default:
        st.sidebar.caption("默认读取资料库中的比赛输入材料。")
        for label, path in DEFAULT_FILES.items():
            exists = "OK" if path.exists() else "Missing"
            st.sidebar.write(f"{exists} `{label}`")
        return dict(DEFAULT_FILES)

    st.sidebar.caption("上传文件将仅用于本次运行，不覆盖资料库。")
    main = st.sidebar.file_uploader("主底稿 U_exp", type=["xlsx"], key="main")
    tod = st.sidebar.file_uploader("TOD 底稿", type=["xlsx"], key="tod")
    program = st.sidebar.file_uploader("审计程序要求", type=["docx"], key="program")
    sop = st.sidebar.file_uploader("SOP Excel", type=["xlsx"], key="sop")

    paths = dict(DEFAULT_FILES)
    overrides = {
        "main_workpaper": save_upload(main, ".xlsx"),
        "tod_workpaper": save_upload(tod, ".xlsx"),
        "program_doc": save_upload(program, ".docx"),
        "sop_workbook": save_upload(sop, ".xlsx"),
    }
    for key, value in overrides.items():
        if value:
            paths[key] = value
    return paths


def render_notes_table(rows: list[dict[str, str]]) -> None:
    if not rows:
        st.success("本次规则未发现 Review Notes。")
        return
    st.dataframe(rows, use_container_width=True, hide_index=True)


def main() -> None:
    inject_style()
    st.markdown(
        """
        <div class="topbar">
          <h1>费用底稿 AI Review 助手</h1>
          <p>基于 SOP、审计程序要求和当前 VC&VD 底稿，自动生成可定位的 Review Notes。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    paths = file_selector()

    col_a, col_b = st.columns([0.68, 0.32])
    with col_a:
        st.subheader("Review 控制台")
        st.markdown(
            '<div class="section-caption">点击后将读取主底稿、TOD 底稿、SOP 和程序要求，执行当前 MVP 规则。</div>',
            unsafe_allow_html=True,
        )
    with col_b:
        run_clicked = st.button("开始 Review", type="primary", use_container_width=True)

    if not run_clicked:
        st.info("当前默认案例资料已就绪。点击“开始 Review”生成 Review Notes。")
        return

    try:
        summary, notes = review_files(paths)
    except Exception as exc:  # Keep demo recoverable.
        st.error(f"Review 运行失败：{exc}")
        st.stop()

    st.subheader("风险概览")
    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Review Notes", summary.total_notes)
    m2.metric("High", summary.high_count)
    m3.metric("Medium", summary.medium_count)
    m4.metric("Low", summary.low_count)
    m5.metric("涉及 Sheet", summary.sheets_impacted)
    m6.metric("规则数", summary.rules_run)

    st.caption(
        f"已读取程序要求 {summary.program_rows} 行；SOP Review checklist 识别 {summary.checklist_rows} 个检查点。"
    )

    rows = notes_to_rows(notes)
    high_rows = [row for row in rows if row["risk_level"] == "High"]
    medium_rows = [row for row in rows if row["risk_level"] == "Medium"]

    tab_all, tab_high, tab_medium = st.tabs(["全部 Review Notes", "High", "Medium"])
    with tab_all:
        render_notes_table(rows)
    with tab_high:
        render_notes_table(high_rows)
    with tab_medium:
        render_notes_table(medium_rows)

    st.download_button(
        "下载 Review Notes CSV",
        data=notes_to_csv(notes),
        file_name="expense_review_notes.csv",
        mime="text/csv",
        use_container_width=True,
    )

    with st.expander("本次规则覆盖说明", expanded=False):
        st.write(asdict(summary))
        st.markdown(
            """
            当前 MVP 优先覆盖已给案例中最稳定、最容易向评委解释的检查点：
            Lead 基础信息、汇总页执行状态、法律费用一致性、BKD Notes、截止性测试、TOD 样本池。
            """
        )


if __name__ == "__main__":
    main()
