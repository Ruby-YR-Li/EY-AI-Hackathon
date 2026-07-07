from __future__ import annotations

import html
import json
import mimetypes
import os
import re
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from llm_review import enhance_review_with_llm, final_stats
from review_engine import (
    OUTPUT_DIR,
    SAMPLE_DIR,
    UPLOAD_DIR,
    as_public_payload,
    create_demo_workbook,
    ensure_dirs,
    review_workbook,
)


BASE_DIR = Path(__file__).resolve().parent
LAST_RESULT: dict | None = None


def page(result: dict | None = None, error: str = "") -> bytes:
    result_json = json.dumps(result or {}, ensure_ascii=False)
    notes = (result or {}).get("notes", [])
    stats = (result or {}).get("stats", {})
    source = (result or {}).get("source", {})
    exports = (result or {}).get("exports", {})
    llm_meta = (result or {}).get("llm", {})
    export_buttons = ""
    if exports:
        export_buttons = f"""
        <a class="btn btn-secondary" href="/download?path={urllib.parse.quote(exports.get('xlsx',''))}">导出 Review Notes</a>
        <a class="btn btn-secondary" href="/download?path={urllib.parse.quote(exports.get('highlighted',''))}">下载高亮底稿</a>
        """
    body = f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>AI审计底稿Review助手</title>
<style>
:root {{
  --primary: #1a3a5c; --primary-light: #2a5a8c; --primary-bg: #e8eef4;
  --danger: #d32f2f; --danger-bg: #fdecea;
  --warning: #e67e22; --warning-bg: #fff8f0;
  --success: #2e7d32; --success-bg: #e8f5e9;
  --bg: #f0f2f5; --card: #ffffff; --text: #1a1a2e; --text-secondary: #6b7280;
  --border: #e5e7eb; --line: #d9dee7;
  --sidebar-w: 220px; --shadow: 0 1px 3px rgba(0,0,0,.08);
  --shadow-hover: 0 4px 12px rgba(0,0,0,.12); --radius: 8px;
  --font: "Microsoft YaHei", "PingFang SC", "Helvetica Neue", sans-serif;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: var(--font); color: var(--text); background: var(--bg); min-height: 100vh; }}
a {{ text-decoration: none; }}

/* Sidebar */
.sidebar {{
  width: var(--sidebar-w); min-height: 100vh; background: linear-gradient(180deg, #152b45, var(--primary));
  color: #fff; position: fixed; left: 0; top: 0; bottom: 0;
  display: flex; flex-direction: column; z-index: 100; overflow-y: auto;
}}
.sidebar-logo {{ padding: 24px 20px 20px; border-bottom: 1px solid rgba(255,255,255,.1); }}
.sidebar-logo h2 {{ font-size: 17px; font-weight: 700; letter-spacing: .5px; }}
.sidebar-logo .st {{ font-size: 11px; opacity: .55; margin-top: 4px; }}
.sidebar-nav {{ flex: 1; padding: 12px 0; }}
.nav-item {{
  display: flex; align-items: center; gap: 10px; padding: 11px 20px;
  color: rgba(255,255,255,.7); font-size: 13px; cursor: pointer; transition: all .2s;
  border-left: 3px solid transparent;
}}
.nav-item:hover, .nav-item.active {{ color: #fff; background: rgba(255,255,255,.08); border-left-color: #5b9bd5; }}
.nav-item .ni {{ font-size: 15px; width: 22px; text-align: center; }}
.sidebar-footer {{ padding: 16px 20px; border-top: 1px solid rgba(255,255,255,.1); font-size: 11px; opacity: .45; }}

/* Main */
.main-content {{ margin-left: var(--sidebar-w); padding: 24px 28px 40px; max-width: calc(100% - var(--sidebar-w)); }}
.section-anchor {{ scroll-margin-top: 20px; }}

/* Header */
.page-header {{ margin-bottom: 20px; }}
.page-header h1 {{ font-size: 24px; font-weight: 700; color: var(--primary); }}
.meta-row {{ display: flex; gap: 18px; font-size: 13px; color: var(--text-secondary); margin-top: 6px; flex-wrap: wrap; }}
.meta-item {{ display: flex; align-items: center; gap: 4px; }}
.meta-item strong {{ color: var(--text); }}

/* Upload area */
.upload-area {{
  background: var(--card); border-radius: var(--radius); padding: 16px 20px;
  margin-bottom: 20px; box-shadow: var(--shadow);
  display: flex; align-items: center; gap: 12px; flex-wrap: wrap;
}}
.upload-area form {{ display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }}
input[type=file] {{ border: 1.5px dashed #bbb; padding: 9px 12px; border-radius: 6px; font-size: 13px; background: #fafbfc; font-family: var(--font); }}
.btn {{
  appearance: none; border: 0; border-radius: 6px; padding: 10px 18px;
  font-weight: 600; cursor: pointer; font-size: 13px; font-family: var(--font);
  display: inline-flex; align-items: center; gap: 5px; transition: all .2s;
}}
.btn-primary {{ background: var(--primary); color: #fff; }}
.btn-primary:hover {{ background: var(--primary-light); }}
.btn-secondary {{ background: #e8edf3; color: var(--primary); border: 1px solid #cdd5df; }}
.btn-secondary:hover {{ background: #dde3eb; }}
.danger {{ color: #8a1f11; background: #fff2f0; border: 1px solid #f3b4aa; padding: 10px 14px; border-radius: 6px; font-size: 13px; }}

/* Stat cards */
.stat-cards {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-bottom: 20px; }}
.stat-card {{
  background: var(--card); border-radius: var(--radius); padding: 18px 20px;
  box-shadow: var(--shadow); transition: all .25s; border-top: 3px solid transparent; cursor: default;
}}
.stat-card:hover {{ transform: translateY(-2px); box-shadow: var(--shadow-hover); }}
.stat-card.total {{ border-top-color: var(--primary); }}
.stat-card.high-card {{ border-top-color: var(--danger); }}
.stat-card.medium-card {{ border-top-color: var(--warning); }}
.stat-card.low-card {{ border-top-color: var(--success); }}
.stat-card .sl {{ font-size: 12px; color: var(--text-secondary); display: flex; align-items: center; gap: 5px; }}
.stat-card .sv {{ font-size: 34px; font-weight: 800; margin-top: 6px; }}
.stat-card.total .sv {{ color: var(--primary); }}
.stat-card.high-card .sv {{ color: var(--danger); }}
.stat-card.medium-card .sv {{ color: var(--warning); }}
.stat-card.low-card .sv {{ color: var(--success); }}

/* Mid row */
.mid-row {{ display: grid; grid-template-columns: 280px 1fr; gap: 14px; margin-bottom: 20px; }}
.mid-card {{ background: var(--card); border-radius: var(--radius); padding: 20px; box-shadow: var(--shadow); }}
.mid-card h3 {{ font-size: 14px; font-weight: 600; margin-bottom: 14px; color: var(--text); }}
.risk-wrap {{ display: flex; align-items: center; gap: 20px; }}
.risk-ring {{ flex-shrink: 0; }}
.risk-score {{ font-size: 36px; font-weight: 800; fill: var(--warning); }}
.risk-comment {{ font-size: 12px; color: var(--text-secondary); padding: 6px 12px; background: var(--warning-bg); border-radius: 4px; font-weight: 500; text-align: center; }}

/* Program bars */
.program-bars {{ display: flex; flex-direction: column; gap: 8px; }}
.pbar {{ display: flex; align-items: center; gap: 8px; font-size: 12px; }}
.pbar-label {{ width: 100px; text-align: right; flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--text-secondary); }}
.pbar-track {{ flex: 1; height: 22px; background: #eef1f5; border-radius: 4px; overflow: hidden; }}
.pbar-fill {{ height: 100%; border-radius: 4px; transition: width .5s ease; }}
.pbar-cnt {{ width: 28px; flex-shrink: 0; font-weight: 700; color: var(--text); text-align: right; }}

/* Category cloud */
.cat-cloud {{ background: var(--card); border-radius: var(--radius); padding: 16px 20px; margin-bottom: 20px; box-shadow: var(--shadow); }}
.cat-cloud h3 {{ font-size: 13px; font-weight: 600; color: var(--text-secondary); margin-bottom: 12px; }}
.cat-tags {{ display: flex; flex-wrap: wrap; gap: 8px; }}
.cat-tag {{
  display: inline-flex; align-items: center; gap: 5px; padding: 5px 12px;
  border-radius: 16px; font-size: 12px; background: #f3f5f8; color: #4b5563;
  cursor: pointer; transition: all .2s; border: 1px solid transparent; user-select: none;
}}
.cat-tag:hover {{ background: #e2e8f0; }}
.cat-tag.active {{ background: var(--primary-bg); border-color: var(--primary-light); color: var(--primary); }}
.cat-dot {{ width: 7px; height: 7px; border-radius: 50%; flex-shrink: 0; }}
.cat-cnt {{ font-weight: 700; opacity: .65; }}

/* Filter bar */
.filter-bar {{ background: var(--card); border-radius: var(--radius); padding: 14px 18px; margin-bottom: 14px; box-shadow: var(--shadow); }}
.filter-row {{ display: flex; align-items: center; gap: 6px; flex-wrap: wrap; margin-bottom: 8px; }}
.filter-row:last-child {{ margin-bottom: 0; }}
.filter-label {{ font-size: 11px; font-weight: 700; color: var(--text-secondary); margin-right: 2px; text-transform: uppercase; letter-spacing: .3px; white-space: nowrap; }}
.filter-pill {{
  padding: 4px 12px; border-radius: 14px; font-size: 12px; font-weight: 500;
  cursor: pointer; transition: all .2s; border: 1.5px solid var(--border);
  background: #f9fafb; color: #4b5563; user-select: none;
}}
.filter-pill:hover {{ background: #eef1f5; }}
.filter-pill.active {{ background: var(--primary); color: #fff; border-color: var(--primary); }}
.filter-pill.sev-high.active {{ background: var(--danger); border-color: var(--danger); }}
.filter-pill.sev-medium.active {{ background: var(--warning); border-color: var(--warning); }}
.filter-pill.sev-low.active {{ background: var(--success); border-color: var(--success); }}
.filter-search {{ margin-left: auto; position: relative; }}
.filter-search input {{
  padding: 5px 10px 5px 30px; border: 1.5px solid var(--border);
  border-radius: 16px; font-size: 12px; width: 190px; outline: none;
  font-family: var(--font); transition: border-color .2s;
}}
.filter-search input:focus {{ border-color: var(--primary-light); }}
.filter-search .si {{ position: absolute; left: 9px; top: 50%; transform: translateY(-50%); font-size: 14px; color: #9ca3af; pointer-events: none; }}
.result-count {{ font-size: 12px; color: var(--text-secondary); margin-left: 8px; white-space: nowrap; }}

/* Clear btn */
.btn-clear {{ padding: 4px 12px; border-radius: 14px; font-size: 11px; cursor: pointer; border: 1.5px solid var(--danger); background: #fff; color: var(--danger); font-weight: 600; transition: all .2s; }}
.btn-clear:hover {{ background: var(--danger-bg); }}

/* Top 5 */
.top5-section {{
  background: linear-gradient(135deg, #fff5f5, #fff0e6);
  border: 1.5px solid #fcc; border-radius: var(--radius);
  padding: 18px 20px; margin-bottom: 14px;
}}
.top5-section h3 {{ font-size: 14px; color: var(--danger); margin-bottom: 10px; font-weight: 700; }}
.top5-list {{ list-style: none; display: flex; flex-direction: column; gap: 4px; }}
.top5-item {{
  display: flex; align-items: center; gap: 8px; padding: 7px 10px;
  border-radius: 6px; cursor: pointer; transition: background .2s; font-size: 13px;
}}
.top5-item:hover {{ background: rgba(211,47,47,.06); }}
.top5-rank {{
  width: 22px; height: 22px; border-radius: 50%; background: var(--danger);
  color: #fff; font-size: 11px; font-weight: 700; display: flex;
  align-items: center; justify-content: center; flex-shrink: 0;
}}
.top5-id {{ font-weight: 700; color: var(--danger); font-size: 12px; min-width: 56px; }}
.top5-desc {{ color: var(--text); flex: 1; }}

/* Issue cards */
.issue-list {{ display: flex; flex-direction: column; gap: 6px; margin-bottom: 20px; }}
.issue-card {{
  background: var(--card); border-radius: 6px; box-shadow: var(--shadow);
  overflow: hidden; border-left: 4px solid transparent;
  transition: all .2s; cursor: pointer;
}}
.issue-card:hover {{ box-shadow: var(--shadow-hover); }}
.issue-card.high {{ border-left-color: var(--danger); }}
.issue-card.medium {{ border-left-color: var(--warning); }}
.issue-card.low {{ border-left-color: var(--success); }}
.issue-card.expanded {{ box-shadow: var(--shadow-hover); }}
.issue-header {{
  display: flex; align-items: center; gap: 10px; padding: 12px 16px;
}}
.issue-dot {{ width: 9px; height: 9px; border-radius: 50%; flex-shrink: 0; }}
.issue-dot.high {{ background: var(--danger); }}
.issue-dot.medium {{ background: var(--warning); }}
.issue-dot.low {{ background: var(--success); }}
.issue-id {{ font-weight: 700; font-size: 12px; min-width: 52px; flex-shrink: 0; }}
.issue-id.high {{ color: var(--danger); }}
.issue-id.medium {{ color: var(--warning); }}
.issue-id.low {{ color: var(--success); }}
.issue-title {{ flex: 1; font-size: 13px; font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
.issue-badges {{ display: flex; gap: 5px; flex-shrink: 0; flex-wrap: wrap; }}
.issue-badge {{
  padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 500;
}}
.badge-prog {{ background: var(--primary-bg); color: var(--primary); }}
.badge-cat {{ background: #f3f5f8; color: #4b5563; }}
.badge-src {{ background: #fef3c7; color: #92400e; }}
.badge-src.llm {{ background: #e0f2fe; color: #0369a1; }}
.issue-chev {{ font-size: 11px; color: #9ca3af; transition: transform .2s; }}
.issue-card.expanded .issue-chev {{ transform: rotate(180deg); }}
.issue-body {{ display: none; padding: 0 16px 14px; border-top: 1px solid #f0f1f3; font-size: 13px; }}
.issue-card.expanded .issue-body {{ display: block; }}
.issue-sec {{ margin-top: 10px; }}
.issue-sec .lbl {{ font-size: 10px; font-weight: 700; color: #9ca3af; text-transform: uppercase; letter-spacing: .4px; margin-bottom: 4px; }}
.issue-sec .problem {{
  color: var(--danger); background: var(--danger-bg); padding: 10px 14px;
  border-radius: 6px; border-left: 3px solid var(--danger); line-height: 1.7;
}}
.issue-sec .suggestion {{
  color: var(--success); background: var(--success-bg); padding: 10px 14px;
  border-radius: 6px; border-left: 3px solid var(--success); line-height: 1.7;
}}
.issue-sec .risk {{
  color: var(--warning); background: var(--warning-bg); padding: 8px 12px;
  border-radius: 6px; border-left: 3px solid var(--warning); line-height: 1.6; font-size: 12px;
}}
.issue-meta {{ display: flex; gap: 18px; flex-wrap: wrap; font-size: 12px; color: var(--text-secondary); margin-top: 10px; }}

/* Bottom stats */
.bottom-stats {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin-top: 8px; }}
.bs-card {{ background: var(--card); border-radius: var(--radius); padding: 18px; box-shadow: var(--shadow); text-align: center; }}
.bs-card .bs-val {{ font-size: 26px; font-weight: 800; color: var(--primary); }}
.bs-card .bs-lbl {{ font-size: 12px; color: var(--text-secondary); margin-top: 4px; }}

/* Empty state */
.empty-state {{ text-align: center; padding: 48px 24px; color: #9ca3af; font-size: 14px; border: 2px dashed var(--border); border-radius: var(--radius); background: #fafbfc; }}

/* Flash animation */
.highlight-flash {{ animation: hpulse 1.5s ease; }}
@keyframes hpulse {{ 0% {{ background: #fff9c4; }} 100% {{ background: transparent; }} }}

/* Responsive */
@media (max-width: 1200px) {{
  .stat-cards {{ grid-template-columns: repeat(2, 1fr); }}
  .mid-row {{ grid-template-columns: 1fr; }}
}}
@media (max-width: 900px) {{
  .sidebar {{ display: none; }}
  .main-content {{ margin-left: 0; max-width: 100%; padding: 16px; }}
  .stat-cards {{ grid-template-columns: 1fr; }}
  .bottom-stats {{ grid-template-columns: repeat(2, 1fr); }}
}}
</style>
</head>
<body>

<!-- Sidebar -->
<aside class="sidebar">
  <div class="sidebar-logo">
    <h2>AI底稿Review</h2>
    <div class="st">质检报告仪表盘</div>
  </div>
  <nav class="sidebar-nav">
    <a class="nav-item active" onclick="scrollToSection('overview', event)"><span class="ni">&#9632;</span> 概览面板</a>
    <a class="nav-item" onclick="scrollToSection('upload-sec', event)"><span class="ni">&#11014;</span> 上传底稿</a>
    <a class="nav-item" onclick="scrollToSection('issues-sec', event)"><span class="ni">&#9776;</span> 问题清单</a>
    <a class="nav-item" onclick="scrollToSection('bottom-sec', event)"><span class="ni">&#9432;</span> 汇总统计</a>
  </nav>
  <div class="sidebar-footer">
    A公司 | 2025年度<br>报告生成：2026-07-07
  </div>
</aside>

<!-- Main Content -->
<main class="main-content">

  <!-- Header -->
  <div class="page-header">
    <h1>AI审计底稿质检报告</h1>
    <div class="meta-row">
      <span class="meta-item">文件：<strong>{html.escape(source.get('file','尚未上传'))}</strong></span>
      <span class="meta-item">页签：<strong>{source.get('sheet_count','-')}</strong></span>
      <span class="meta-item">公式：<strong>{source.get('formula_count','-')}</strong></span>
      <span class="meta-item">AI增强：<strong>{html.escape(llm_meta.get('message', '尚未运行'))}</strong></span>
    </div>
  </div>

  <!-- Upload -->
  <section class="upload-area section-anchor" id="upload-sec">
    <form action="/review" method="post" enctype="multipart/form-data">
      <input type="file" name="workpaper" accept=".xlsx,.xlsm" required>
      <button class="btn btn-primary" type="submit">&#9654; 开始Review</button>
      <a class="btn btn-secondary" href="/demo">使用模拟问题底稿</a>
      {export_buttons}
    </form>
    {f'<div class="danger">{html.escape(error)}</div>' if error else ''}
  </section>

  <!-- Stat Cards -->
  <div class="stat-cards section-anchor" id="overview">
    <div class="stat-card total">
      <div class="sl">&#128202; Review Notes</div>
      <div class="sv" id="statTotal">{stats.get('total', 0)}</div>
    </div>
    <div class="stat-card high-card">
      <div class="sl">&#128308; High</div>
      <div class="sv" id="statHigh">{stats.get('high', 0)}</div>
    </div>
    <div class="stat-card medium-card">
      <div class="sl">&#128999; Medium</div>
      <div class="sv" id="statMedium">{stats.get('medium', 0)}</div>
    </div>
    <div class="stat-card low-card">
      <div class="sl">&#128994; Low</div>
      <div class="sv" id="statLow">{stats.get('low', 0)}</div>
    </div>
  </div>

  <!-- Risk Score + Program Distribution -->
  <div class="mid-row">
    <div class="mid-card">
      <h3>综合风险评分</h3>
      <div class="risk-wrap">
        <div class="risk-ring">
          <svg width="140" height="140" viewBox="0 0 140 140" style="transform:rotate(-90deg)">
            <circle cx="70" cy="70" r="58" fill="none" stroke="#e5e7eb" stroke-width="10"/>
            <circle id="gaugeFill" cx="70" cy="70" r="58" fill="none" stroke="#e67e22" stroke-width="10" stroke-linecap="round" stroke-dasharray="364.43" stroke-dashoffset="102"/>
          </svg>
          <text x="70" y="70" text-anchor="middle" dy=".35em" font-size="34" font-weight="800" fill="#e67e22" transform="rotate(90, 70, 70)" id="gaugeScore">72</text>
        </div>
        <div>
          <div class="risk-comment" id="gaugeComment">需关注，建议整改后复核</div>
        </div>
      </div>
    </div>
    <div class="mid-card">
      <h3>按审计程序的问题分布</h3>
      <div class="program-bars" id="programBars"></div>
    </div>
  </div>

  <!-- Category Cloud -->
  <div class="cat-cloud">
    <h3>按问题类别分布（点击筛选）</h3>
    <div class="cat-tags" id="categoryTags"></div>
  </div>

  <!-- Filter Bar -->
  <div class="filter-bar section-anchor" id="issues-sec">
    <div class="filter-row">
      <span class="filter-label">严重度</span>
      <span class="filter-pill active" data-filter="severity" data-value="all">全部</span>
      <span class="filter-pill sev-high" data-filter="severity" data-value="High">&#128308; 高</span>
      <span class="filter-pill sev-medium" data-filter="severity" data-value="Medium">&#128999; 中</span>
      <span class="filter-pill sev-low" data-filter="severity" data-value="Low">&#128994; 低</span>
    </div>
    <div class="filter-row">
      <span class="filter-label">程序</span>
      <span class="filter-pill active" data-filter="program" data-value="all">全部</span>
      <span id="programFilters"></span>
      <button class="btn-clear" onclick="clearAllFilters()">&#10005; 清除筛选</button>
      <div class="filter-search">
        <span class="si">&#128269;</span>
        <input type="text" id="searchInput" placeholder="搜索标题或内容...">
      </div>
      <span class="result-count" id="resultCount">显示 0 条</span>
    </div>
  </div>

  <!-- Top 5 Priority -->
  <div class="top5-section" id="top5Section">
    <h3>&#128293; 最需要立即处理的前5大问题</h3>
    <ul class="top5-list" id="top5List"></ul>
  </div>

  <!-- Issue Cards -->
  <div class="issue-list" id="issueList"></div>
  <div class="empty-state" id="emptyState" style="display:none">没有匹配的问题，请调整筛选条件。</div>

  <!-- Bottom Stats -->
  <div class="bottom-stats section-anchor" id="bottom-sec">
    <div class="bs-card"><div class="bs-val" id="bsTotal">0</div><div class="bs-lbl">Review Notes 总数</div></div>
    <div class="bs-card"><div class="bs-val" style="color:#1565c0;" id="bsRule">0</div><div class="bs-lbl">规则引擎发现</div></div>
    <div class="bs-card"><div class="bs-val" style="color:#7c3aed;" id="bsLLM">0</div><div class="bs-lbl">LLM增强发现</div></div>
    <div class="bs-card"><div class="bs-val" style="color:#2e7d32;" id="bsExport">-</div><div class="bs-lbl">导出文件</div></div>
  </div>

</main>

<script>
// ---- DATA ----
var REVIEW = window.__AI_REVIEW_RESULT__ || {{}};
var notes = REVIEW.notes || [];
var stats = REVIEW.stats || {{}};
var exports = REVIEW.exports || {{}};

// ---- INIT ----
document.addEventListener("DOMContentLoaded", function() {{
  renderAll();
  bindFilters();
  document.getElementById("searchInput").addEventListener("input", renderIssues);
}});

function renderAll() {{
  renderStatCards();
  renderGauge();
  renderProgramBars();
  renderCategoryTags();
  renderProgramFilters();
  renderTop5();
  renderIssues();
  renderBottomStats();
}}

// ---- Stat Cards ----
function renderStatCards() {{
  document.getElementById("statTotal").textContent = stats.total || 0;
  document.getElementById("statHigh").textContent = stats.high || 0;
  document.getElementById("statMedium").textContent = stats.medium || 0;
  document.getElementById("statLow").textContent = stats.low || 0;
}}

// ---- Risk Gauge ----
function renderGauge() {{
  var total = stats.total || 0;
  if (total === 0) return;
  var score = Math.round((stats.high * 100 + stats.medium * 50 + stats.low * 15) / (total * 100) * 100);
  score = Math.max(5, Math.min(95, score));
  var dash = 364.43;
  var offset = dash - (score / 100) * dash;
  document.getElementById("gaugeFill").setAttribute("stroke-dashoffset", offset);
  document.getElementById("gaugeScore").textContent = score;
  var c, t;
  if (score >= 65) {{ c = "#d32f2f"; t = "需重点关注，建议立即整改"; }}
  else if (score >= 35) {{ c = "#e67e22"; t = "需关注，建议整改后复核"; }}
  else {{ c = "#2e7d32"; t = "底稿质量良好，少量优化即可"; }}
  document.getElementById("gaugeFill").setAttribute("stroke", c);
  document.getElementById("gaugeScore").setAttribute("fill", c);
  document.getElementById("gaugeComment").textContent = t;
}}

// ---- Program Bars ----
function renderProgramBars() {{
  var prog = {{}};
  notes.forEach(function(n) {{
    var s = n.sheet || "其他";
    prog[s] = (prog[s] || 0) + 1;
  }});
  var sorted = Object.entries(prog).sort(function(a, b) {{ return b[1] - a[1]; }});
  var max = sorted.length ? sorted[0][1] : 1;
  var colors = ["#3a7bd5","#5b9bd5","#2a5a8c","#4a90d9","#1a3a5c","#6db3f2","#8db8e8","#a0c8f0"];
  var html = "";
  sorted.forEach(function(e, i) {{
    var pct = Math.round((e[1] / max) * 100);
    html += '<div class="pbar"><div class="pbar-label">' + esc(e[0]) + '</div>';
    html += '<div class="pbar-track"><div class="pbar-fill" style="width:' + pct + '%;background:' + colors[i % colors.length] + ';"></div></div>';
    html += '<div class="pbar-cnt">' + e[1] + '</div></div>';
  }});
  document.getElementById("programBars").innerHTML = html || '<div style="color:#9ca3af;font-size:13px;">暂无数据</div>';
}}

// ---- Category Tags ----
var curCategory = "all";
var catColors = ["#e74c3c","#e67e22","#f39c12","#3498db","#9b59b6","#1abc9c","#34495e","#7f8c8d","#e91e63","#00bcd4"];
function renderCategoryTags() {{
  var cats = stats.categories || {{}};
  var html = "";
  var i = 0;
  for (var k in cats) {{
    html += '<span class="cat-tag" data-cat="' + esc(k) + '" onclick="filterByCategory(\'' + esc(k) + '\', event)">';
    html += '<span class="cat-dot" style="background:' + catColors[i % catColors.length] + ';"></span>';
    html += esc(k) + ' <span class="cat-cnt">' + cats[k] + '</span></span>';
    i++;
  }}
  document.getElementById("categoryTags").innerHTML = html || '<span style="color:#9ca3af;font-size:13px;">暂无数据</span>';
}}

function filterByCategory(cat, evt) {{
  if (curCategory === cat) {{
    curCategory = "all";
    document.querySelectorAll(".cat-tag").forEach(function(t) {{ t.classList.remove("active"); }});
  }} else {{
    curCategory = cat;
    document.querySelectorAll(".cat-tag").forEach(function(t) {{ t.classList.remove("active"); }});
    if (evt && evt.currentTarget) evt.currentTarget.classList.add("active");
  }}
  renderIssues();
}}

// ---- Program Filters ----
function renderProgramFilters() {{
  var progs = new Set();
  notes.forEach(function(n) {{ progs.add(n.sheet); }});
  var html = "";
  progs.forEach(function(p) {{
    if (!p) return;
    html += '<span class="filter-pill" data-filter="program" data-value="' + esc(p) + '">' + esc(p) + '</span>';
  }});
  document.getElementById("programFilters").innerHTML = html;
  // rebind
  document.querySelectorAll(".filter-pill[data-filter=program]").forEach(function(b) {{
    b.addEventListener("click", function() {{
      var g = document.querySelectorAll('.filter-pill[data-filter="' + this.dataset.filter + '"]');
      g.forEach(function(s) {{ s.classList.remove("active"); }});
      this.classList.add("active");
      renderIssues();
    }});
  }});
}}

// ---- Filters ----
function bindFilters() {{
  document.querySelectorAll(".filter-pill[data-filter=severity]").forEach(function(b) {{
    b.addEventListener("click", function() {{
      document.querySelectorAll('.filter-pill[data-filter=severity]').forEach(function(s) {{ s.classList.remove("active"); }});
      this.classList.add("active");
      renderIssues();
    }});
  }});
}}

function getFilters() {{
  var sev = document.querySelector('.filter-pill[data-filter=severity].active').dataset.value;
  var prog = document.querySelector('.filter-pill[data-filter=program].active').dataset.value;
  var q = (document.getElementById("searchInput").value || "").trim().toLowerCase();
  return {{ severity: sev, program: prog, search: q, category: curCategory }};
}}

// ---- Top 5 ----
function renderTop5() {{
  var sorted = notes.slice().sort(function(a, b) {{
    var sv = {{ High: 3, Medium: 2, Low: 1 }};
    return (sv[b.severity] || 0) - (sv[a.severity] || 0);
  }});
  var top5 = sorted.slice(0, 5);
  var html = "";
  top5.forEach(function(n, i) {{
    html += '<li class="top5-item" onclick="jumpToIssue(\'' + esc(n.id) + '\')">';
    html += '<span class="top5-rank">' + (i + 1) + '</span>';
    html += '<span class="top5-id">' + esc(n.id) + '</span>';
    html += '<span class="top5-desc">' + esc(n.issue) + '</span></li>';
  }});
  document.getElementById("top5List").innerHTML = html;
  var sec = document.getElementById("top5Section");
  if (top5.length === 0) sec.style.display = "none"; else sec.style.display = "";
}}

function jumpToIssue(id) {{
  clearAllFilters();
  setTimeout(function() {{
    var card = document.getElementById("ic-" + id);
    if (!card) return;
    document.querySelectorAll(".issue-card.expanded").forEach(function(c) {{ c.classList.remove("expanded"); }});
    card.classList.add("expanded");
    card.classList.add("highlight-flash");
    card.scrollIntoView({{ behavior: "smooth", block: "center" }});
    setTimeout(function() {{ card.classList.remove("highlight-flash"); }}, 1500);
  }}, 150);
}}

// ---- Issue Cards ----
function renderIssues() {{
  var f = getFilters();
  var filtered = notes.filter(function(n) {{
    if (f.severity !== "all" && n.severity !== f.severity) return false;
    if (f.program !== "all" && n.sheet !== f.program) return false;
    if (f.category !== "all" && n.category !== f.category) return false;
    if (f.search) {{
      var h = (n.id + " " + n.issue + " " + n.suggestion + " " + n.sheet + " " + n.category + " " + (n.location||"")).toLowerCase();
      if (h.indexOf(f.search) === -1) return false;
    }}
    return true;
  }});

  document.getElementById("resultCount").textContent = "显示 " + filtered.length + " 条";

  if (filtered.length === 0) {{
    document.getElementById("issueList").innerHTML = "";
    document.getElementById("emptyState").style.display = "block";
    return;
  }}
  document.getElementById("emptyState").style.display = "none";

  var sevLabel = {{ High: "高", Medium: "中", Low: "低" }};
  var sevEmoji = {{ High: "&#128308;", Medium: "&#128999;", Low: "&#128994;" }};
  var html = "";
  filtered.forEach(function(n) {{
    var srcClass = (n.source || "").indexOf("llm") !== -1 ? "llm" : "";
    html += '<div class="issue-card ' + n.severity.toLowerCase() + '" id="ic-' + esc(n.id) + '" onclick="toggleCard(this, event)">';
    html += '<div class="issue-header">';
    html += '<span class="issue-dot ' + n.severity.toLowerCase() + '"></span>';
    html += '<span class="issue-id ' + n.severity.toLowerCase() + '">' + esc(n.id) + '</span>';
    html += '<span class="issue-title" title="' + esc(n.issue) + '">' + esc(n.issue) + '</span>';
    html += '<div class="issue-badges">';
    html += '<span class="issue-badge badge-prog">' + esc(n.sheet||"") + '</span>';
    html += '<span class="issue-badge badge-cat">' + esc(n.category||"") + '</span>';
    html += '<span class="issue-badge badge-src ' + srcClass + '">' + esc(n.source||"rule") + '</span>';
    html += '</div>';
    html += '<span class="issue-chev">&#9660;</span>';
    html += '</div>';
    html += '<div class="issue-body">';
    html += '<div class="issue-sec"><div class="lbl">&#128308; 问题描述</div><div class="problem">' + esc(n.issue) + '</div></div>';
    html += '<div class="issue-sec"><div class="lbl">&#128161; 建议修改</div><div class="suggestion">' + esc(n.suggestion) + '</div></div>';
    if (n.audit_risk) {{
      html += '<div class="issue-sec"><div class="lbl">&#9888; 审计风险</div><div class="risk">' + esc(n.audit_risk) + '</div></div>';
    }}
    html += '<div class="issue-meta">';
    html += '<span>&#128205; 定位：' + esc(n.location||n.sheet) + '</span>';
    html += '<span>&#128193; 程序：' + esc(n.sheet||"") + '</span>';
    html += '<span>&#128220; 类别：' + esc(n.category||"") + '</span>';
    html += '<span>' + sevEmoji[n.severity] + ' ' + sevLabel[n.severity] + '</span>';
    if (n.confidence) html += '<span>&#127919; 置信度：' + Math.round(n.confidence * 100) + '%</span>';
    html += '</div>';
    html += '</div>';
    html += '</div>';
  }});
  document.getElementById("issueList").innerHTML = html;
}}

function toggleCard(card, evt) {{
  if (evt.target.tagName === "BUTTON" || evt.target.tagName === "A") return;
  var was = card.classList.contains("expanded");
  document.querySelectorAll(".issue-card.expanded").forEach(function(c) {{ c.classList.remove("expanded"); }});
  if (!was) card.classList.add("expanded");
}}

// ---- Bottom Stats ----
function renderBottomStats() {{
  document.getElementById("bsTotal").textContent = stats.total || 0;
  var ruleCount = notes.filter(function(n) {{ return (n.source||"") === "rule"; }}).length;
  var llmCount = notes.filter(function(n) {{ return (n.source||"").indexOf("llm") !== -1; }}).length;
  document.getElementById("bsRule").textContent = ruleCount;
  document.getElementById("bsLLM").textContent = llmCount;
  document.getElementById("bsExport").textContent = (exports.xlsx || exports.csv) ? "可用 &#10003;" : "暂无";
}}

// ---- Clear All Filters ----
function clearAllFilters() {{
  document.querySelectorAll(".filter-pill").forEach(function(b) {{ b.classList.remove("active"); }});
  document.querySelectorAll('.filter-pill[data-filter=severity][data-value=all]').forEach(function(b) {{ b.classList.add("active"); }});
  document.querySelectorAll('.filter-pill[data-filter=program][data-value=all]').forEach(function(b) {{ b.classList.add("active"); }});
  curCategory = "all";
  document.querySelectorAll(".cat-tag").forEach(function(t) {{ t.classList.remove("active"); }});
  document.getElementById("searchInput").value = "";
  renderIssues();
}}

// ---- Sidebar Navigation ----
function scrollToSection(id, evt) {{
  var el = document.getElementById(id);
  if (el) el.scrollIntoView({{ behavior: "smooth", block: "start" }});
  document.querySelectorAll(".nav-item").forEach(function(n) {{ n.classList.remove("active"); }});
  if (evt && evt.currentTarget) evt.currentTarget.classList.add("active");
}}

// ---- Helpers ----
function esc(str) {{
  if (!str) return "";
  var d = document.createElement("div");
  d.textContent = str;
  return d.innerHTML;
}}
</script>

</body>
</html>"""
    return body.encode("utf-8")


def parse_multipart(body: bytes, content_type: str) -> tuple[str, bytes]:
    match = re.search(r"boundary=(.+)", content_type)
    if not match:
        raise ValueError("无法识别上传边界。")
    boundary = ("--" + match.group(1)).encode()
    for part in body.split(boundary):
        if b"Content-Disposition" not in part or b"name=\"workpaper\"" not in part:
            continue
        header, _, data = part.partition(b"\r\n\r\n")
        name_match = re.search(rb'filename="([^"]+)"', header)
        filename = name_match.group(1).decode("utf-8", "ignore") if name_match else "upload.xlsx"
        data = data.rstrip(b"\r\n-")
        return filename, data
    raise ValueError("未找到上传文件。")


def run_review(path: Path) -> dict:
    rule_notes, rule_stats = review_workbook(path)
    notes, llm_meta = enhance_review_with_llm(path, rule_notes, rule_stats)
    stats = final_stats(notes)
    payload = as_public_payload(notes, stats, path)
    payload["llm"] = llm_meta
    payload["rule_stats"] = rule_stats
    return payload


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        global LAST_RESULT
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/":
            self.respond(200, page(LAST_RESULT), "text/html; charset=utf-8")
        elif parsed.path == "/demo":
            try:
                demo = create_demo_workbook()
                LAST_RESULT = run_review(demo)
                self.respond(200, page(LAST_RESULT), "text/html; charset=utf-8")
            except Exception as exc:
                self.respond(500, page(LAST_RESULT, str(exc)), "text/html; charset=utf-8")
        elif parsed.path == "/download":
            params = urllib.parse.parse_qs(parsed.query)
            requested = Path(params.get("path", [""])[0])
            try:
                resolved = requested.resolve()
                allowed_roots = [OUTPUT_DIR.resolve(), SAMPLE_DIR.resolve(), UPLOAD_DIR.resolve()]
                if not any(str(resolved).startswith(str(root)) for root in allowed_roots) or not resolved.exists():
                    raise FileNotFoundError("文件不可下载或不存在。")
                ctype = mimetypes.guess_type(str(resolved))[0] or "application/octet-stream"
                self.send_response(200)
                self.send_header("Content-Type", ctype)
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{urllib.parse.quote(resolved.name)}")
                self.end_headers()
                self.wfile.write(resolved.read_bytes())
            except Exception as exc:
                self.respond(404, page(LAST_RESULT, str(exc)), "text/html; charset=utf-8")
        else:
            self.respond(404, b"Not Found", "text/plain; charset=utf-8")

    def do_POST(self) -> None:
        global LAST_RESULT
        if self.path != "/review":
            self.respond(404, b"Not Found", "text/plain; charset=utf-8")
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            filename, content = parse_multipart(self.rfile.read(length), self.headers.get("Content-Type", ""))
            clean = Path(filename).name or "upload.xlsx"
            target = UPLOAD_DIR / clean
            target.write_bytes(content)
            LAST_RESULT = run_review(target)
            self.respond(200, page(LAST_RESULT), "text/html; charset=utf-8")
        except Exception as exc:
            self.respond(500, page(LAST_RESULT, str(exc)), "text/html; charset=utf-8")

    def respond(self, status: int, body: bytes, content_type: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def main() -> None:
    ensure_dirs()
    create_demo_workbook()
    port = int(os.environ.get("PORT", "8765"))
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    print(f"AI Review Assistant running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
