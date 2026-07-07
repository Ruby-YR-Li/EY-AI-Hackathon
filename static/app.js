const state = {
  result: null,
  selectedId: null,
  annotatedUrl: null,
  progressTimer: null,
};

const severityLabel = {
  High: "高",
  Medium: "中",
  Low: "低",
};

const els = {
  uploadForm: document.querySelector("#uploadForm"),
  workpaperInput: document.querySelector("#workpaperInput"),
  workpaperCount: document.querySelector("#workpaperCount"),
  tbInput: document.querySelector("#tbInput"),
  jeInput: document.querySelector("#jeInput"),
  programInput: document.querySelector("#programInput"),
  sampleBtn: document.querySelector("#sampleBtn"),
  heroHealthBtn: document.querySelector("#heroHealthBtn"),
  statusPill: document.querySelector("#statusPill"),
  runMeta: document.querySelector("#runMeta"),
  metricTotal: document.querySelector("#metricTotal"),
  metricHigh: document.querySelector("#metricHigh"),
  metricMedium: document.querySelector("#metricMedium"),
  metricLow: document.querySelector("#metricLow"),
  heroMetricTotal: document.querySelector("#heroMetricTotal"),
  heroMetricHigh: document.querySelector("#heroMetricHigh"),
  severityFilter: document.querySelector("#severityFilter"),
  categoryFilter: document.querySelector("#categoryFilter"),
  supportCheckToggle: document.querySelector("#supportCheckToggle"),
  llmToggle: document.querySelector("#llmToggle"),
  visibleCount: document.querySelector("#visibleCount"),
  findingsBody: document.querySelector("#findingsBody"),
  detailPanel: document.querySelector("#detailPanel"),
  detailSeverity: document.querySelector("#detailSeverity"),
  copyMdBtn: document.querySelector("#copyMdBtn"),
  downloadCsvBtn: document.querySelector("#downloadCsvBtn"),
  downloadAnnotatedBtn: document.querySelector("#downloadAnnotatedBtn"),
  aiMode: document.querySelector("#aiMode"),
  aiMessage: document.querySelector("#aiMessage"),
  toggleConfigBtn: document.querySelector("#toggleConfigBtn"),
  apiConfigPanel: document.querySelector("#apiConfigPanel"),
  apiKeyInput: document.querySelector("#apiKeyInput"),
  modelInput: document.querySelector("#modelInput"),
  endpointInput: document.querySelector("#endpointInput"),
  saveConfigBtn: document.querySelector("#saveConfigBtn"),
  clearConfigBtn: document.querySelector("#clearConfigBtn"),
  configMessage: document.querySelector("#configMessage"),
  progressText: document.querySelector("#progressText"),
  progressPercent: document.querySelector("#progressPercent"),
  progressBar: document.querySelector("#progressBar"),
  progressSteps: document.querySelector("#progressSteps"),
  spotlight: document.querySelector("#spotlight"),
  heroStage: document.querySelector(".hero-stage"),
};

function setStatus(text, mode = "ready") {
  els.statusPill.textContent = text;
  els.statusPill.dataset.mode = mode;
}

function setProgress(percent, text, steps = []) {
  els.progressBar.style.width = `${percent}%`;
  els.progressPercent.textContent = `${percent}%`;
  els.progressText.textContent = text;
  els.progressSteps.innerHTML = steps
    .map((step) => `<li class="${step.status || ""}">${escapeHtml(step.text)}</li>`)
    .join("");
}

function progressLabels(kind) {
  const supportStep = els.supportCheckToggle.checked ? "执行 TB/JE 核对" : "跳过 TB/JE 核对";
  const llmStep = els.llmToggle?.checked ? "可选 LLM 润色" : "规则结果输出";
  if (kind === "upload") {
    return ["接收上传材料", "解析 Excel 材料", "执行 VCVD 规则", supportStep, llmStep];
  }
  return ["读取样例材料", "解析 Excel 材料", "执行 VCVD 规则", supportStep, llmStep];
}

function startProgress(kind) {
  clearInterval(state.progressTimer);
  const labels = progressLabels(kind);
  let current = 0;
  setProgress(8, labels[0], labels.map((text, idx) => ({ text, status: idx === 0 ? "done" : "" })));
  state.progressTimer = setInterval(() => {
    current = Math.min(current + 1, labels.length - 1);
    const percent = Math.min(15 + current * 17, 88);
    setProgress(
      percent,
      labels[current],
      labels.map((text, idx) => ({ text, status: idx <= current ? "done" : "" }))
    );
    if (current >= labels.length - 1) {
      clearInterval(state.progressTimer);
    }
  }, 800);
}

function finishProgress(text = "分析完成") {
  clearInterval(state.progressTimer);
  const steps = [
    { text: "材料解析完成", status: "done" },
    { text: "VCVD 规则检查完成", status: "done" },
    { text: "Review Notes 已生成", status: "done" },
  ];
  if (els.llmToggle?.checked) {
    steps.push({ text: "LLM 润色已尝试", status: "done" });
  }
  setProgress(100, text, steps);
}

function failProgress(message) {
  clearInterval(state.progressTimer);
  setProgress(100, "分析失败", [{ text: message, status: "error" }]);
}

function filteredFindings() {
  const findings = state.result?.findings || [];
  const severity = els.severityFilter.value;
  const category = els.categoryFilter.value;
  return findings.filter((item) => {
    const severityMatch = severity === "all" || item.severity === severity;
    const categoryMatch = category === "all" || item.category === category;
    return severityMatch && categoryMatch;
  });
}

function locationText(item) {
  const loc = item.location || {};
  return [loc.sheet, loc.cell].filter(Boolean).join("!");
}

function syncHeroMetrics(stats) {
  if (els.heroMetricTotal) els.heroMetricTotal.textContent = stats.total;
  if (els.heroMetricHigh) els.heroMetricHigh.textContent = stats.high;
  const whyTotal = document.querySelector("#whyMetricTotal");
  const whyHigh = document.querySelector("#whyMetricHigh");
  if (whyTotal) whyTotal.textContent = stats.total;
  if (whyHigh) whyHigh.textContent = stats.high;
}

function renderAll() {
  const stats = state.result?.stats || { total: 0, high: 0, medium: 0, low: 0 };
  els.metricTotal.textContent = stats.total;
  els.metricHigh.textContent = stats.high;
  els.metricMedium.textContent = stats.medium;
  els.metricLow.textContent = stats.low;
  syncHeroMetrics(stats);
  els.runMeta.textContent = state.result
    ? `${state.result.generatedAt} · ${state.result.files.length} 个材料 · 内置 VCVD SOP`
    : "等待分析";
  renderAiStatus();
  renderAnnotatedStatus();
  renderCategoryFilter();
  renderTable();
  renderDetail();
}

function renderAnnotatedStatus() {
  state.annotatedUrl = state.result?.annotated?.url || null;
  els.downloadAnnotatedBtn.disabled = !state.annotatedUrl;
  els.downloadAnnotatedBtn.textContent = state.annotatedUrl ? "下载批注底稿（可选）" : "暂无批注底稿";
}

function renderAiStatus() {
  const ai = state.result?.ai;
  const llmEnabled = els.llmToggle?.checked;
  if (!ai) {
    els.aiMode.textContent = llmEnabled ? "规则 + LLM（待分析）" : "规则引擎";
    els.aiMessage.textContent = llmEnabled
      ? "已开启 LLM 润色；需配置 API Key 后才会生效。"
      : "默认仅使用内置 VCVD 规则与证据定位。";
    return;
  }
  if (ai.mode === "llm_enhanced") {
    els.aiMode.textContent = `规则 + LLM · ${ai.model}`;
  } else {
    els.aiMode.textContent = llmEnabled ? "规则引擎（LLM 未生效）" : "规则引擎";
  }
  els.aiMessage.textContent = ai.message;
}

async function loadConfig() {
  try {
    const response = await safeFetch("/api/config", { method: "GET" });
    const config = await response.json();
    if (config.endpoint) els.endpointInput.value = config.endpoint;
    if (config.ai?.model && config.ai.model !== "not configured") {
      els.modelInput.value = config.ai.model;
    }
    if (config.ai?.enabled && els.llmToggle) {
      els.llmToggle.checked = true;
    }
    state.result = state.result || {};
    state.result.ai = config.ai;
    renderAiStatus();
  } catch (error) {
    els.configMessage.textContent = error.message;
  }
}

async function checkHealth() {
  const response = await safeFetch("/api/health", { method: "GET" });
  const payload = await response.json();
  if (!payload.ok) throw new Error("本地分析服务未就绪。");
}

async function saveConfig(clearApiKey = false) {
  els.configMessage.textContent = "保存中...";
  const payload = {
    apiKey: els.apiKeyInput.value.trim(),
    model: els.modelInput.value.trim() || "gpt-4.1-mini",
    endpoint: els.endpointInput.value.trim() || "https://api.openai.com/v1/responses",
    clearApiKey,
  };
  const response = await safeFetch("/api/config", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const config = await response.json();
  state.result = state.result || {};
  state.result.ai = config.ai;
  renderAiStatus();
  els.apiKeyInput.value = "";
  els.configMessage.textContent = clearApiKey ? "API Key 已清除。" : "配置已保存。开启 LLM 后再次分析会尝试润色。";
}

function renderCategoryFilter() {
  const current = els.categoryFilter.value;
  const categories = [...new Set((state.result?.findings || []).map((item) => item.category))];
  els.categoryFilter.innerHTML = `<option value="all">全部</option>`;
  for (const category of categories) {
    const option = document.createElement("option");
    option.value = category;
    option.textContent = category;
    els.categoryFilter.appendChild(option);
  }
  if (categories.includes(current)) {
    els.categoryFilter.value = current;
  }
}

function renderTable() {
  const findings = filteredFindings();
  els.visibleCount.textContent = `${findings.length} 项`;
  if (!findings.length) {
    els.findingsBody.innerHTML = `<tr><td colspan="6" class="empty">没有符合筛选条件的 Review Note</td></tr>`;
    return;
  }
  els.findingsBody.innerHTML = "";
  for (const item of findings) {
    const row = document.createElement("tr");
    row.className = item.id === state.selectedId ? "active" : "";
    row.innerHTML = `
      <td>${item.id}</td>
      <td><span class="badge ${item.severity}">${severityLabel[item.severity] || item.severity}</span></td>
      <td>${escapeHtml(item.category)}</td>
      <td><span class="problem-text">${escapeHtml(item.title)}</span></td>
      <td><span class="risk-text">${escapeHtml(item.risk)}</span></td>
      <td><span class="location">${escapeHtml(locationText(item))}</span></td>
    `;
    row.addEventListener("click", () => {
      state.selectedId = item.id;
      renderTable();
      renderDetail();
      document.querySelector("#results")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    els.findingsBody.appendChild(row);
  }
  if (!state.selectedId && findings[0]) {
    state.selectedId = findings[0].id;
    renderTable();
    renderDetail();
  }
}

function renderDetail() {
  const item = (state.result?.findings || []).find((finding) => finding.id === state.selectedId);
  if (!item) {
    els.detailSeverity.textContent = "未选择";
    els.detailPanel.innerHTML = `
      <div class="panel-head">
        <h3>问题详情</h3>
        <span id="detailSeverity">未选择</span>
      </div>
      <div class="detail-empty">选择一条 Review Note 查看证据、风险与建议修改。</div>
    `;
    els.detailSeverity = document.querySelector("#detailSeverity");
    return;
  }
  els.detailPanel.innerHTML = `
    <div class="panel-head">
      <h3>${escapeHtml(item.id)} · ${escapeHtml(item.category)}</h3>
      <span class="badge ${item.severity}">${severityLabel[item.severity] || item.severity}</span>
    </div>
    <div class="detail-body">
      <section class="detail-section">
        <h4>位置</h4>
        <p class="location">${escapeHtml(item.location.file || "")}<br>${escapeHtml(locationText(item))}</p>
      </section>
      <section class="detail-section">
        <h4>问题发现</h4>
        <p class="problem-text">${escapeHtml(item.title)}</p>
      </section>
      <section class="detail-section">
        <h4>Review Note</h4>
        <p>${escapeHtml(item.reviewNote)}</p>
      </section>
      <section class="detail-section">
        <h4>审计风险</h4>
        <p class="risk-text">${escapeHtml(item.risk)}</p>
      </section>
      <section class="detail-section">
        <h4>建议修改</h4>
        <p>${escapeHtml(item.recommendation)}</p>
      </section>
      <section class="detail-section">
        <h4>定位证据</h4>
        <p class="evidence">${escapeHtml(item.evidence)}</p>
      </section>
    </div>
  `;
}

async function analyzeSample() {
  setStatus("Analyzing", "busy");
  startProgress("sample");
  const support = els.supportCheckToggle.checked ? "1" : "0";
  const response = await safeFetch(`/api/analyze-sample?support=${support}`, { method: "POST" });
  state.result = await response.json();
  if (state.result.error) throw new Error(state.result.error);
  state.selectedId = state.result.findings[0]?.id || null;
  setStatus("Done", "done");
  renderAll();
  finishProgress("样例分析完成");
}

async function analyzeUpload() {
  const form = new FormData();
  for (const file of els.workpaperInput.files) {
    form.append("workpapers", file);
  }
  for (const file of els.programInput.files) {
    form.append("programs", file);
  }
  if (els.tbInput.files[0]) {
    form.append("tb", els.tbInput.files[0]);
  }
  if (els.jeInput.files[0]) {
    form.append("je", els.jeInput.files[0]);
  }
  form.append("supportChecks", els.supportCheckToggle.checked ? "1" : "0");
  setStatus("Analyzing", "busy");
  startProgress("upload");
  const response = await safeFetch("/api/analyze", { method: "POST", body: form });
  state.result = await response.json();
  if (state.result.error) throw new Error(state.result.error);
  state.selectedId = state.result.findings[0]?.id || null;
  setStatus("Done", "done");
  renderAll();
  finishProgress("上传分析完成");
  document.querySelector("#results")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function safeFetch(url, options) {
  try {
    const response = await fetch(url, options);
    if (!response.ok) {
      let message = `${response.status} ${response.statusText}`;
      try {
        const payload = await response.clone().json();
        message = payload.error || message;
      } catch (_) {
        message = await response.text();
      }
      throw new Error(message || "请求失败");
    }
    return response;
  } catch (error) {
    if (error instanceof TypeError) {
      throw new Error("无法连接本地分析服务。请确认使用 http://127.0.0.1:8765/ 打开，并且服务正在运行。");
    }
    throw error;
  }
}

function markdownNotes() {
  const findings = state.result?.findings || [];
  if (!findings.length) return "";
  return findings
    .map((item) => {
      const loc = locationText(item);
      return `### ${item.id} ${item.title}
- 等级：${severityLabel[item.severity] || item.severity}
- 类型：${item.category}
- 位置：${item.location.file} / ${loc}
- Review Note：${item.reviewNote}
- 建议修改：${item.recommendation}
- 审计风险：${item.risk}`;
    })
    .join("\n\n");
}

function csvNotes() {
  const rows = [["编号", "等级", "类型", "文件", "位置", "问题", "Review Note", "建议修改", "审计风险"]];
  for (const item of state.result?.findings || []) {
    rows.push([
      item.id,
      severityLabel[item.severity] || item.severity,
      item.category,
      item.location.file || "",
      locationText(item),
      item.title,
      item.reviewNote,
      item.recommendation,
      item.risk,
    ]);
  }
  return rows
    .map((row) => row.map((cell) => `"${String(cell).replaceAll('"', '""')}"`).join(","))
    .join("\n");
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function initScrollReveal() {
  const nodes = document.querySelectorAll(".scroll-reveal");
  if (!("IntersectionObserver" in window)) {
    nodes.forEach((node) => node.classList.add("is-visible"));
    return;
  }
  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("is-visible");
          observer.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.14, rootMargin: "0px 0px -8% 0px" }
  );
  nodes.forEach((node) => observer.observe(node));
}

function initSpotlight() {
  if (!els.spotlight || window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  let visible = false;
  const show = () => {
    visible = true;
    els.spotlight.style.opacity = "1";
  };
  const hide = () => {
    visible = false;
    els.spotlight.style.opacity = "0";
  };
  document.addEventListener(
    "mousemove",
    (event) => {
      if (!visible) show();
      els.spotlight.style.left = `${event.clientX}px`;
      els.spotlight.style.top = `${event.clientY}px`;
    },
    { passive: true }
  );
  document.addEventListener("mouseleave", hide, { passive: true });
}

function initParallax() {
  if (window.matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  const stage = els.heroStage;
  if (!stage) return;
  const cards = stage.querySelectorAll(".floating-card");
  window.addEventListener(
    "scroll",
    () => {
      const rect = stage.getBoundingClientRect();
      const progress = Math.min(Math.max(1 - rect.top / window.innerHeight, 0), 1);
      stage.style.transform = `translateY(${progress * 18}px)`;
      cards.forEach((card) => {
        const depth = Number(card.dataset.depth || 1);
        card.style.setProperty("--parallax-y", `${progress * 22 * depth}px`);
      });
    },
    { passive: true }
  );
}

async function runHealthCheck() {
  await checkHealth();
  setStatus("Ready", "done");
  setProgress(100, "本地分析服务已连接", [{ text: "服务在线，可以上传底稿开始 Review", status: "done" }]);
  els.runMeta.textContent = "服务在线，请选择底稿、TB/JE 后点击分析上传。";
}

els.workpaperInput.addEventListener("change", () => {
  const count = els.workpaperInput.files.length;
  els.workpaperCount.textContent = count ? `${count} 个底稿已选择` : "可多选 U_exp、TOD 等 Excel";
});

els.uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!els.workpaperInput.files.length) {
    setStatus("Need files", "error");
    const message = "请先在「选择底稿」区域上传至少一个 Excel 底稿。";
    els.runMeta.textContent = message;
    failProgress(message);
    return;
  }
  try {
    await analyzeUpload();
  } catch (error) {
    setStatus("Error", "error");
    els.runMeta.textContent = error.message;
    failProgress(error.message);
  }
});

els.sampleBtn.addEventListener("click", async () => {
  try {
    await runHealthCheck();
  } catch (error) {
    setStatus("Error", "error");
    els.runMeta.textContent = error.message;
    failProgress(error.message);
  }
});

els.heroHealthBtn?.addEventListener("click", async () => {
  try {
    await runHealthCheck();
    document.querySelector("#workspace")?.scrollIntoView({ behavior: "smooth" });
  } catch (error) {
    setStatus("Error", "error");
    els.runMeta.textContent = error.message;
    failProgress(error.message);
  }
});

els.severityFilter.addEventListener("change", () => {
  state.selectedId = null;
  renderTable();
  renderDetail();
});

els.categoryFilter.addEventListener("change", () => {
  state.selectedId = null;
  renderTable();
  renderDetail();
});

els.llmToggle?.addEventListener("change", () => renderAiStatus());

els.copyMdBtn.addEventListener("click", async () => {
  const notes = markdownNotes();
  if (!notes) return;
  await navigator.clipboard.writeText(notes);
  setStatus("Copied", "done");
});

els.downloadCsvBtn.addEventListener("click", () => {
  const csv = "\ufeff" + csvNotes();
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "review-notes.csv";
  link.click();
  URL.revokeObjectURL(url);
});

els.downloadAnnotatedBtn.addEventListener("click", () => {
  if (!state.annotatedUrl) return;
  const link = document.createElement("a");
  link.href = state.annotatedUrl;
  link.download = "annotated_review_workpapers.zip";
  link.click();
});

els.toggleConfigBtn.addEventListener("click", () => {
  els.apiConfigPanel.hidden = !els.apiConfigPanel.hidden;
});

els.saveConfigBtn.addEventListener("click", async () => {
  try {
    await saveConfig(false);
  } catch (error) {
    els.configMessage.textContent = error.message;
  }
});

els.clearConfigBtn.addEventListener("click", async () => {
  try {
    await saveConfig(true);
  } catch (error) {
    els.configMessage.textContent = error.message;
  }
});

initScrollReveal();
initSpotlight();
initParallax();

if (window.location.protocol === "file:") {
  setStatus("Open URL", "ready");
  els.runMeta.textContent = "请通过 http://127.0.0.1:8765/ 打开，上传和分析功能需要本地服务。";
} else {
  checkHealth()
    .then(() => loadConfig())
    .then(() => {
      setStatus("Ready", "done");
      setProgress(0, "请选择底稿后开始分析", []);
      els.runMeta.textContent = "服务在线，内置 VCVD SOP 已就绪。";
    })
    .catch((error) => {
      setStatus("Error", "error");
      els.runMeta.textContent = error.message;
      failProgress(error.message);
    });
}
