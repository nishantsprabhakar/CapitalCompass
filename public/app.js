const form = document.getElementById("dealForm");
const authHome = document.getElementById("authHome");
const appShell = document.getElementById("appShell");
const loginForm = document.getElementById("loginForm");
const signupForm = document.getElementById("signupForm");
const authMessage = document.getElementById("authMessage");
const paymentLink = document.getElementById("paymentLink");
const premiumButton = document.getElementById("premiumButton");
const promoCode = document.getElementById("promoCode");
const logoutButton = document.getElementById("logoutButton");
const accountName = document.getElementById("accountName");
const accountAccess = document.getElementById("accountAccess");
const adminNav = document.getElementById("adminNav");
const statusEl = document.getElementById("runStatus");
const downloads = document.getElementById("downloads");
const thesis = document.getElementById("thesis");
const summary = document.getElementById("summary");
const templateBox = document.getElementById("templateBox");
const enterpriseReview = document.getElementById("enterpriseReview");
const pipelineOwnerEmail = document.getElementById("pipelineOwnerEmail");
const pipelineImportForm = document.getElementById("pipelineImportForm");
const pipelineManualForm = document.getElementById("pipelineManualForm");
const pipelineDeals = document.getElementById("pipelineDeals");
const pipelineSummary = document.getElementById("pipelineSummary");
const refreshPipeline = document.getElementById("refreshPipeline");
const pipelineSearch = document.getElementById("pipelineSearch");
const pipelineStageFilter = document.getElementById("pipelineStageFilter");
const pipelineStatusFilter = document.getElementById("pipelineStatusFilter");
const pipelineStatusChart = document.getElementById("pipelineStatusChart");
const pipelineStageChart = document.getElementById("pipelineStageChart");
const pipelineSectorChart = document.getElementById("pipelineSectorChart");
const pipelineEconomicsChart = document.getElementById("pipelineEconomicsChart");
let currentUser = null;
let csrfToken = "";
let currentPaymentLink = "#";
let currentPipelineOwner = "nishant.p@skegen.com";
let pipelineState = { deals: [], summary: {}, ownerEmail: "" };

initInterfaceEffects();
initAuth();

fetch("/api/template")
  .then((r) => r.json())
  .then((t) => {
    templateBox.classList.add("template-hidden");
    templateBox.textContent = "";
  })
  .catch(() => { templateBox.textContent = "Template metadata unavailable."; });

document.querySelectorAll(".auth-tab").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".auth-tab").forEach((x) => x.classList.remove("active"));
    button.classList.add("active");
    const mode = button.dataset.authMode;
    loginForm.classList.toggle("hidden", mode !== "login");
    signupForm.classList.toggle("hidden", mode !== "signup");
    authMessage.textContent = "";
  });
});

loginForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  await authSubmit("/api/login", loginForm, "Login failed");
});

signupForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  await authSubmit("/api/signup", signupForm, "Signup failed");
});

logoutButton.addEventListener("click", async () => {
  await apiFetch("/api/logout", { method: "POST" });
  currentUser = null;
  csrfToken = "";
  appShell.classList.add("hidden");
  authHome.classList.remove("hidden");
  authMessage.textContent = "Logged out securely.";
});

promoCode.addEventListener("input", () => refreshPaymentLink());
refreshPipeline.addEventListener("click", () => loadPipeline());
[pipelineSearch, pipelineStageFilter, pipelineStatusFilter].forEach((control) => {
  control?.addEventListener("input", () => renderPipeline(pipelineState));
  control?.addEventListener("change", () => renderPipeline(pipelineState));
});
document.querySelectorAll(".funnel-stage").forEach((stage) => {
  stage.addEventListener("click", () => {
    const value = stage.dataset.stage || "";
    pipelineStageFilter.value = pipelineStageFilter.value === value ? "" : value;
    renderPipeline(pipelineState);
  });
});
pipelineImportForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!currentUser) return showAuthMessage("Please login before importing a pipeline.", true);
  const data = new FormData(pipelineImportForm);
  const response = await apiFetch(`/api/pipeline/import?ownerEmail=${encodeURIComponent(pipelineOwnerEmail.value || currentPipelineOwner)}`, { method: "POST", body: data });
  const result = await response.json();
  if (!response.ok) return renderPipelineError(result.error || "Pipeline import failed.");
  pipelineImportForm.reset();
  pipelineOwnerEmail.value = result.ownerEmail || currentPipelineOwner;
  renderPipeline(result);
});
pipelineManualForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!currentUser) return showAuthMessage("Please login before adding a deal.", true);
  const body = Object.fromEntries(new FormData(pipelineManualForm).entries());
  body.ownerEmail = pipelineOwnerEmail.value || currentPipelineOwner;
  const response = await apiFetch(`/api/pipeline/deals?ownerEmail=${encodeURIComponent(body.ownerEmail)}`, { method: "POST", body: JSON.stringify(body) });
  const result = await response.json();
  if (!response.ok) return renderPipelineError(result.error || "Could not add deal.");
  pipelineManualForm.reset();
  await loadPipeline();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!currentUser) return showAuthMessage("Please login before running diligence.", true);
  const button = form.querySelector("button[type=submit]");
  button.disabled = true;
  statusEl.textContent = "Running";
  downloads.textContent = "Generating files...";
  thesis.textContent = "Extracting materials, building diligence view, and preparing exports.";
  enterpriseReview.textContent = "Benchmarking enterprise readiness, source quality, audit trail, and strategic-buyer fit.";

  try {
    const body = buildSubmissionBody();
    const response = await apiFetch("/api/analyze", { method: "POST", body });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "Analysis failed");
    renderAnalysis(data.analysis, data.files);
    statusEl.textContent = "Complete";
  } catch (error) {
    statusEl.textContent = "Error";
    thesis.textContent = error.message;
    downloads.textContent = "No files generated.";
    enterpriseReview.textContent = "Enterprise review unavailable because analysis failed.";
  } finally {
    button.disabled = false;
  }
});

async function initAuth() {
  try {
    const response = await fetch("/api/me", { credentials: "same-origin" });
    const data = await response.json();
    csrfToken = data.csrfToken || "";
    currentPaymentLink = data.paymentLink || "#";
    if (data.user) {
      currentUser = data.user;
      showWorkspace();
    } else {
      authHome.classList.remove("hidden");
      appShell.classList.add("hidden");
    }
    await refreshPaymentLink();
  } catch {
    authHome.classList.remove("hidden");
  }
}

async function authSubmit(url, sourceForm, fallback) {
  const button = sourceForm.querySelector("button");
  button.disabled = true;
  showAuthMessage("Securing session...");
  try {
    const payload = Object.fromEntries(new FormData(sourceForm).entries());
    const response = await fetch(url, {
      method: "POST",
      credentials: "same-origin",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || fallback);
    currentUser = data.user;
    csrfToken = data.csrfToken || "";
    currentPaymentLink = data.paymentLink || "#";
    sourceForm.reset();
    showWorkspace();
    await refreshPaymentLink();
  } catch (error) {
    showAuthMessage(error.message, true);
  } finally {
    button.disabled = false;
  }
}

async function apiFetch(url, options = {}) {
  const headers = { ...(options.headers || {}) };
  const isFormData = options.body instanceof FormData;
  if (!isFormData && options.body && !headers["Content-Type"]) headers["Content-Type"] = "application/json";
  if ((options.method || "GET").toUpperCase() !== "GET" && csrfToken) headers["X-CSRF-Token"] = csrfToken;
  return fetch(url, { ...options, credentials: "same-origin", headers });
}

function showWorkspace() {
  authHome.classList.add("hidden");
  appShell.classList.remove("hidden");
  const access = currentUser.featureAccess || {};
  currentPipelineOwner = currentUser.role === "admin" ? "nishant.p@skegen.com" : currentUser.email;
  if (pipelineOwnerEmail) {
    pipelineOwnerEmail.value = currentPipelineOwner;
    pipelineOwnerEmail.disabled = currentUser.role !== "admin";
  }
  accountName.textContent = `${currentUser.name || currentUser.email}`;
  accountAccess.textContent = `${currentUser.role.toUpperCase()} | ${currentUser.plan.toUpperCase()} plan | Deep dive ${access.deepDive ? "enabled" : "locked"} | AI review ${access.aiReview ? "enabled" : "locked"}`;
  adminNav.classList.toggle("hidden", currentUser.role !== "admin");
  loadPipeline();
}

function showAuthMessage(message, isError = false) {
  authMessage.textContent = message;
  authMessage.classList.toggle("error", Boolean(isError));
}

async function refreshPaymentLink() {
  const code = encodeURIComponent(promoCode?.value?.trim() || "");
  try {
    const response = currentUser ? await fetch(`/api/payment-link${code ? `?promo=${code}` : ""}`, { credentials: "same-origin" }) : null;
    const data = response?.ok ? await response.json() : { paymentLink: currentPaymentLink };
    currentPaymentLink = data.paymentLink || currentPaymentLink || "#";
  } catch {}
  if (paymentLink) paymentLink.href = currentPaymentLink || "#";
  if (premiumButton) premiumButton.href = currentPaymentLink || "#";
}

async function loadPipeline() {
  if (!currentUser) return;
  const owner = pipelineOwnerEmail?.value || currentPipelineOwner;
  const response = await apiFetch(`/api/pipeline?ownerEmail=${encodeURIComponent(owner)}`);
  const data = await response.json();
  if (!response.ok) return renderPipelineError(data.error || "Could not load pipeline.");
  currentPipelineOwner = data.ownerEmail || owner;
  if (pipelineOwnerEmail) pipelineOwnerEmail.value = currentPipelineOwner;
  renderPipeline(data);
}

function renderPipeline(data) {
  pipelineState = data || pipelineState;
  const allDeals = pipelineState.deals || [];
  const deals = filterPipelineDeals(allDeals);
  const s = buildPipelineSummary(deals);
  pipelineSummary.innerHTML = [
    pipelineMetric("Total deals", s.total ?? 0),
    pipelineMetric("Active / DD", s.active ?? 0),
    pipelineMetric("Rejected", s.rejected ?? 0),
    pipelineMetric("On hold", s.onHold ?? 0),
    pipelineMetric("Avg margin", `${((s.avgMargin || 0) * 100).toFixed(1)}%`),
    pipelineMetric("Revenue pool", `${Math.round(s.totalRevenue || 0).toLocaleString()} cr`)
  ].join("");
  (s.byStage || []).forEach((stage, idx) => {
    const el = document.getElementById(`stage${idx + 1}Count`);
    if (el) {
      el.textContent = stage.count;
      const block = el.closest(".funnel-stage");
      block?.style.setProperty("--stage-scale", String(Math.max(.18, Math.min(1, (stage.count || 0) / Math.max(1, s.total || 1)))));
      block?.classList.toggle("selected", normalizeStage(pipelineStageFilter.value) === stage.stage);
    }
  });
  renderPipelineCharts(deals, s);
  if (!deals.length) {
    pipelineDeals.innerHTML = `<div class="pipeline-empty">No deals match the current pipeline filters.</div>`;
    return;
  }
  pipelineDeals.innerHTML = deals.map((d, idx) => `
    <article class="deal-card ${pipelineStatusClass(d.status)}" style="--delay:${idx * 35}ms" data-deal-id="${escapeHtml(d.id)}">
      <div>
        <span>${escapeHtml(d.stage || "Pipeline")}</span>
        <h3>${escapeHtml(d.company)}</h3>
        <p>${escapeHtml(d.sector || "Sector not tagged")}</p>
      </div>
      <div class="deal-metrics">
        <strong>${formatPipelineValue(d.revenue)}</strong><span>Revenue</span>
        <strong>${formatPipelineValue(d.ebitda)}</strong><span>EBITDA</span>
        <strong>${typeof d.margin === "number" ? `${(d.margin * 100).toFixed(1)}%` : escapeHtml(d.margin || "-")}</strong><span>Margin</span>
      </div>
      <p class="deal-brief">${escapeHtml(d.brief || d.notes || "No brief logged.")}</p>
      <div class="deal-foot">
        <strong>${escapeHtml(d.status || "Active")}</strong>
        <span>${escapeHtml(d.ask || "Ask TBD")}</span>
        <button type="button" class="secondary" data-analyze-deal="${escapeHtml(d.id)}">Analyze</button>
        <button type="button" class="danger" data-delete-deal="${escapeHtml(d.id)}">Delete</button>
      </div>
      <details>
        <summary>Deal notes</summary>
        <p>${escapeHtml(d.notes || d.rejectionReason || d.brief || "No additional notes.")}</p>
      </details>
    </article>
  `).join("");
  pipelineDeals.querySelectorAll("[data-delete-deal]").forEach((button) => {
    button.addEventListener("click", async () => {
      const response = await apiFetch(`/api/pipeline/deals/${encodeURIComponent(button.dataset.deleteDeal)}?ownerEmail=${encodeURIComponent(currentPipelineOwner)}`, { method: "DELETE" });
      if (!response.ok) return renderPipelineError((await response.json()).error || "Could not delete deal.");
      await loadPipeline();
    });
  });
  pipelineDeals.querySelectorAll("[data-analyze-deal]").forEach((button) => {
    button.addEventListener("click", () => {
      const deal = allDeals.find((d) => d.id === button.dataset.analyzeDeal);
      if (!deal) return;
      const company = form.querySelector('[name="companyName"]');
      const sector = form.querySelector('[name="sector"]');
      if (company) company.value = deal.company || "";
      if (sector) sector.value = deal.sector || "";
      document.getElementById("inputs")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
  });
}

function pipelineMetric(label, value) {
  return `<article><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></article>`;
}

function filterPipelineDeals(deals) {
  const q = (pipelineSearch?.value || "").trim().toLowerCase();
  const stage = pipelineStageFilter?.value || "";
  const status = pipelineStatusFilter?.value || "";
  return deals.filter((d) => {
    const haystack = [d.company, d.sector, d.brief, d.notes, d.ask, d.rejectionReason].join(" ").toLowerCase();
    return (!q || haystack.includes(q)) && (!stage || normalizeStage(d.stage) === normalizeStage(stage)) && (!status || String(d.status || "").toLowerCase() === status.toLowerCase());
  });
}

function buildPipelineSummary(deals) {
  const stages = ["1. Deal Sourcing", "2. Initial Screening", "3. Preliminary DD", "4. IC Approval - Prelim", "5. Full Due Diligence"];
  const byStage = stages.map((stage) => ({ stage, count: deals.filter((d) => normalizeStage(d.stage) === stage).length }));
  const active = deals.filter((d) => /active|progress/i.test(d.status || "")).length;
  const rejected = deals.filter((d) => /reject/i.test(d.status || "")).length;
  const onHold = deals.filter((d) => /hold/i.test(d.status || "")).length;
  const totalRevenue = deals.reduce((sum, d) => sum + (typeof d.revenue === "number" ? d.revenue : 0), 0);
  const marginDeals = deals.filter((d) => typeof d.margin === "number");
  const avgMargin = marginDeals.length ? marginDeals.reduce((sum, d) => sum + d.margin, 0) / marginDeals.length : 0;
  return { total: deals.length, active, rejected, onHold, byStage, totalRevenue, avgMargin };
}

function renderPipelineCharts(deals, summary) {
  renderStatusDonut(summary);
  renderStageBars(summary);
  renderSectorBars(deals);
  renderEconomicsChart(deals);
}

function renderStatusDonut(summary) {
  const active = summary.active || 0;
  const rejected = summary.rejected || 0;
  const hold = summary.onHold || 0;
  const total = Math.max(1, active + rejected + hold);
  const activePct = active / total * 100;
  const rejectedPct = rejected / total * 100;
  const rejectedEnd = activePct + rejectedPct;
  pipelineStatusChart.innerHTML = `
    <div class="donut" style="--active:${activePct};--rejected-end:${rejectedEnd};"></div>
    <div class="chart-legend">
      <span><i class="dot active"></i>Active ${active}</span>
      <span><i class="dot rejected"></i>Rejected ${rejected}</span>
      <span><i class="dot hold"></i>Hold ${hold}</span>
    </div>
  `;
}

function renderStageBars(summary) {
  const max = Math.max(1, ...summary.byStage.map((x) => x.count));
  pipelineStageChart.innerHTML = summary.byStage.map((x, idx) => `
    <button type="button" class="stage-bar ${pipelineStageFilter.value === x.stage ? "selected" : ""}" data-stage-filter="${escapeHtml(x.stage)}">
      <span>${idx + 1}</span>
      <strong>${escapeHtml(x.stage.replace(/^\d+\.\s*/, ""))}</strong>
      <em style="--bar:${Math.max(4, x.count / max * 100)}%"></em>
      <b>${x.count}</b>
    </button>
  `).join("");
  pipelineStageChart.querySelectorAll("[data-stage-filter]").forEach((button) => {
    button.addEventListener("click", () => {
      const value = button.dataset.stageFilter;
      pipelineStageFilter.value = normalizeStage(pipelineStageFilter.value) === value ? "" : value;
      renderPipeline(pipelineState);
    });
  });
}

function renderSectorBars(deals) {
  const sectors = Object.entries(deals.reduce((acc, d) => {
    const key = d.sector || "Unclassified";
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {})).sort((a, b) => b[1] - a[1]).slice(0, 6);
  const max = Math.max(1, ...sectors.map((x) => x[1]));
  pipelineSectorChart.innerHTML = sectors.length ? sectors.map(([sector, count]) => `
    <div class="sector-row"><span>${escapeHtml(sector)}</span><em style="--bar:${Math.max(6, count / max * 100)}%"></em><strong>${count}</strong></div>
  `).join("") : `<p class="fine-print">No sector exposure yet.</p>`;
}

function renderEconomicsChart(deals) {
  const economics = deals
    .filter((d) => typeof d.revenue === "number" || typeof d.ebitda === "number")
    .sort((a, b) => (b.revenue || 0) - (a.revenue || 0))
    .slice(0, 8);
  const max = Math.max(1, ...economics.flatMap((d) => [Math.abs(d.revenue || 0), Math.abs(d.ebitda || 0)]));
  pipelineEconomicsChart.innerHTML = economics.length ? economics.map((d) => `
    <div class="econ-row">
      <span>${escapeHtml(d.company)}</span>
      <em class="rev" style="--bar:${Math.max(3, Math.abs(d.revenue || 0) / max * 100)}%"></em>
      <em class="ebitda ${Number(d.ebitda || 0) < 0 ? "negative" : ""}" style="--bar:${Math.max(3, Math.abs(d.ebitda || 0) / max * 100)}%"></em>
      <strong>${formatPipelineValue(d.revenue)}</strong>
    </div>
  `).join("") : `<p class="fine-print">No revenue/EBITDA values yet.</p>`;
}

function normalizeStage(stage) {
  return String(stage || "").replace("–", "-").replace(/\s+/g, " ").trim();
}

function pipelineStatusClass(status) {
  if (/reject/i.test(status || "")) return "deal-rejected";
  if (/hold/i.test(status || "")) return "deal-hold";
  return "deal-active";
}

function formatPipelineValue(value) {
  if (typeof value === "number") return value >= 1000 ? `${Math.round(value).toLocaleString()} cr` : `${value.toLocaleString(undefined, { maximumFractionDigits: 1 })} cr`;
  return escapeHtml(value || "-");
}

function renderPipelineError(message) {
  pipelineDeals.innerHTML = `<div class="pipeline-empty error">${escapeHtml(message)}</div>`;
}

function buildSubmissionBody() {
  const fileInput = form.querySelector('input[type="file"]');
  const templateInput = form.querySelector('input[name="templateFile"]');
  const folderPath = form.querySelector('input[name="folderPath"]').value.trim();
  const files = [...(fileInput.files || [])];
  const totalBytes = files.reduce((sum, file) => sum + file.size, 0);
  const tooLargeForBrowser = totalBytes > 120 * 1024 * 1024;
  const data = new FormData();

  for (const field of form.elements) {
    if (!field.name || field.type === "file" || field.tagName === "BUTTON") continue;
    data.append(field.name, field.value);
  }

  if (tooLargeForBrowser && folderPath) {
    thesis.textContent = "Large folder detected. Reading from the local folder path instead of uploading every file through the browser.";
    return data;
  }

  if (tooLargeForBrowser && !folderPath) {
    throw new Error("The selected folder is too large for browser upload. Paste the local folder path in the folder path field, then run diligence again.");
  }

  for (const file of files) {
    if (/\.(pptx|docx|xlsx|csv|txt|md)$/i.test(file.name)) data.append("files", file, file.webkitRelativePath || file.name);
  }
  const template = templateInput?.files?.[0];
  if (template && /\.pptx$/i.test(template.name)) data.append("templateFile", template, template.name);
  return data;
}

function renderAnalysis(a, files) {
  const included = a.docsSummary.filter((d) => d.status === "included");
  const excluded = a.docsSummary.filter((d) => d.status === "excluded");
  const companyInput = form.querySelector('input[name="companyName"]');
  if (companyInput && !companyInput.value.trim() && a.companyName && a.companyName !== "Target Company") {
    companyInput.value = a.companyName;
  }
  summary.innerHTML = [
    metric("Company", a.companyName),
    metric("Recommendation", a.recommendation),
    metric("Score", `${a.scorecard.total}/100`),
    metric("Confidence", `${a.scorecard.confidence || 0}%`),
    metric("Platform", `${a.enterpriseReadiness?.score || 0}/100`),
    metric("Deal evidence", `${a.enterpriseReadiness?.dealEvidenceQuality || 0}/100`),
    metric("Docs used", `${included.length}/${a.docsSummary.length}`),
    metric("Sources", String(a.research.length))
  ].join("");
  thesis.innerHTML = `
    <div class="analysis-layout">
      <section class="analysis-card wide">
        <div class="card-label">Risk-first investment view</div>
        <ul class="clean-list">${a.thesis.map((x) => `<li>${escapeHtml(x)}</li>`).join("")}</ul>
      </section>
      <section class="analysis-card">
        <div class="card-label">Evidence gaps</div>
        <ul class="compact-list">${a.evidence.missingEvidence.slice(0, 10).map((x) => `<li>${escapeHtml(x)}</li>`).join("")}</ul>
      </section>
      <section class="analysis-card">
        <div class="card-label">AI enrichment</div>
        <div class="doc-control">
          <strong>${escapeHtml(a.aiReview?.status || "disabled")}</strong>
          <span>${escapeHtml(a.aiReview?.provider || "No AI API configured")}${a.aiReview?.model ? ` | ${escapeHtml(a.aiReview.model)}` : ""}</span>
        </div>
        ${a.aiReview?.summary ? `<p class="ai-summary">${escapeHtml(a.aiReview.summary).slice(0, 900)}</p>` : `<p class="ai-summary muted">Optional. Add any AI API endpoint, model, and key in Deal Setup.</p>`}
      </section>
      <section class="analysis-card wide">
        <div class="card-label">Capital Compass IC readiness score</div>
        <div class="score-grid">
          ${Object.entries(a.scorecard.components).map(([key, component]) => `
            <article class="score-pillar">
              <div>
                <strong>${escapeHtml(titleCase(key))}</strong>
                <span>${escapeHtml(component.score)}/20</span>
              </div>
              <p>${escapeHtml(component.rationale)}</p>
              <ul>
                ${(component.subFactors || []).slice(0, 4).map((f) => `<li><b>${escapeHtml(f.score)}/100</b> ${escapeHtml(f.name)} <em>${escapeHtml(f.evidenceTier)}</em></li>`).join("")}
              </ul>
            </article>
          `).join("")}
        </div>
        <div class="score-footnote">
          <span>${escapeHtml(a.scorecard.methodology || "Risk-adjusted PE diligence score.")}</span>
          <span>Penalties: critical ${escapeHtml(a.scorecard.penalties?.criticalRisk ?? 0)}, high ${escapeHtml(a.scorecard.penalties?.highRisk ?? 0)}, evidence ${escapeHtml(a.scorecard.penalties?.evidence ?? 0)}, gates ${escapeHtml(a.scorecard.penalties?.gating ?? 0)}</span>
        </div>
        ${(a.scorecard.gates || []).length ? `<div class="gate-list">${a.scorecard.gates.map((g) => `<div><strong>${escapeHtml(g.title)}</strong><span>${escapeHtml(g.reason)}</span></div>`).join("")}</div>` : ""}
      </section>
      <section class="analysis-card">
        <div class="card-label">Document control</div>
        <div class="doc-control">
          <strong>${included.length} included</strong>
          <span>${excluded.length} excluded as low relevance</span>
          <span>${a.template.uploadedName ? `IC template: ${escapeHtml(a.template.uploadedName)}` : "IC template: CapitalCompass PE style"}</span>
        </div>
        ${excluded.length ? `<details><summary>Show excluded documents</summary><ul class="compact-list">${excluded.slice(0, 20).map((d) => `<li>${escapeHtml(d.name)} <span>score ${d.relevanceScore}</span></li>`).join("")}</ul></details>` : ""}
      </section>
      <section class="analysis-card wide">
        <div class="card-label">Red flag register</div>
        <div class="risk-table">
          ${a.riskRegister.map((r) => `
            <article class="risk-row severity-${r.severity.toLowerCase()}">
              <div><span class="severity">${escapeHtml(r.severity)}</span><strong>${escapeHtml(r.title)}</strong></div>
              <p>${escapeHtml(r.whyItMatters)}</p>
              <small>${escapeHtml(r.diligenceRequired)}</small>
            </article>
          `).join("")}
        </div>
      </section>
      <section class="analysis-card wide">
        <div class="card-label">Top diligence questions</div>
        <div class="question-grid">
          ${a.questions.map((s) => `
            <article>
              <h3>${escapeHtml(s.title)}</h3>
              <ul>${s.items.slice(0, 4).map((q) => `<li>${escapeHtml(q)}</li>`).join("")}</ul>
            </article>
          `).join("")}
        </div>
      </section>
    </div>
  `;
  renderEnterpriseReview(a);
  downloads.innerHTML = files.map((f) => `<a href="${f.href}">${f.label}</a>`).join("");
}

function renderEnterpriseReview(a) {
  const er = a.enterpriseReadiness;
  const sq = a.sourceQuality;
  enterpriseReview.innerHTML = `
    <div class="enterprise-layout">
      <section class="analysis-card wide readiness-hero">
        <div>
          <div class="card-label">Acquisition readiness verdict</div>
          <h3>${escapeHtml(er.verdict)}</h3>
          <p>${escapeHtml(er.acquisitionCase)}</p>
        </div>
        <div class="readiness-score">
          <strong>${escapeHtml(er.score)}</strong>
          <span>/100</span>
        </div>
      </section>
      <section class="analysis-card wide">
        <div class="card-label">Score separation</div>
        <div class="separation-grid">
          <article><span>Platform architecture</span><strong>${escapeHtml(er.score)}/100</strong><p>Product capability score for acquisition-readiness architecture.</p></article>
          <article><span>Deal evidence quality</span><strong>${escapeHtml(er.dealEvidenceQuality)}/100</strong><p>Quality of this specific uploaded target-company evidence package.</p></article>
          <article><span>Commercial proof</span><strong>External</strong><p>Revenue, traction, retention, references, and valuation must be proven outside a local code build.</p></article>
        </div>
      </section>
      <section class="analysis-card">
        <div class="card-label">Source reliability</div>
        <div class="source-meter">
          <strong>${escapeHtml(sq.score)}/100</strong>
          <span>${escapeHtml(sq.verdict)}</span>
        </div>
        <ul class="compact-list">
          <li>${escapeHtml(sq.usableSources)} usable external sources</li>
          <li>${escapeHtml(sq.authoritativeSources)} authoritative sources</li>
          <li>${escapeHtml(sq.institutionalSources)} institutional sources</li>
          <li>${escapeHtml(sq.failedSources)} failed fetches</li>
        </ul>
      </section>
      <section class="analysis-card">
        <div class="card-label">Audit trail</div>
        <div class="audit-grid">
          <span>Session</span><strong>${escapeHtml(a.auditTrail.sessionId)}</strong>
          <span>Documents</span><strong>${escapeHtml(a.auditTrail.documentsIncluded)}/${escapeHtml(a.auditTrail.documentsReviewed)}</strong>
          <span>Score version</span><strong>${escapeHtml(a.auditTrail.scoreVersion)}</strong>
          <span>Record hash</span><strong>${escapeHtml(a.auditTrail.recordHash.slice(0, 16))}...</strong>
        </div>
      </section>
      <section class="analysis-card wide">
        <div class="card-label">Enterprise readiness scorecard</div>
        <div class="enterprise-scorecard">
          ${er.dimensions.map((d) => `
            <article>
              <div><strong>${escapeHtml(d.name)}</strong><span>${escapeHtml(d.score)}/100</span></div>
              <small>${escapeHtml(d.status)} | Weight ${escapeHtml(d.weight)}%</small>
              <p>${escapeHtml(d.rationale)}</p>
            </article>
          `).join("")}
        </div>
      </section>
      <section class="analysis-card wide">
        <div class="card-label">Benchmark against global platforms</div>
        <div class="benchmark-grid">
          ${a.benchmark.map((b) => `
            <article>
              <h3>${escapeHtml(b.platform)}</h3>
              <p><b>Benchmark:</b> ${escapeHtml(b.benchmarkStrength)}</p>
              <p><b>CapitalCompass:</b> ${escapeHtml(b.capitalCompassPosition)}</p>
              <span>${escapeHtml(b.acquisitionImplication)}</span>
            </article>
          `).join("")}
        </div>
      </section>
      <section class="analysis-card wide">
        <div class="card-label">External diligence still required for USD 100mn strategic process</div>
        <ul class="clean-list">${er.mustFix.map((x) => `<li>${escapeHtml(x)}</li>`).join("")}</ul>
      </section>
      <section class="analysis-card wide">
        <div class="card-label">Completed platform controls</div>
        <ul class="clean-list">${(er.completedControls || []).map((x) => `<li>${escapeHtml(x)}</li>`).join("")}</ul>
      </section>
    </div>
  `;
}

function metric(label, value) {
  return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`;
}

function escapeHtml(value) {
  return String(value).replace(/[<>&'"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&#39;", '"': "&quot;" }[c]));
}

function titleCase(value) {
  return String(value).replace(/([A-Z])/g, " $1").replace(/^./, (m) => m.toUpperCase());
}

function initInterfaceEffects() {
  document.documentElement.classList.add("js-ready");
  const progress = document.querySelector(".scroll-progress");
  const navLinks = [...document.querySelectorAll("nav a[href^='#']")];
  const sections = navLinks
    .map((link) => document.querySelector(link.getAttribute("href")))
    .filter(Boolean);

  function updateProgress() {
    const max = document.documentElement.scrollHeight - window.innerHeight;
    const pct = max > 0 ? Math.min(100, Math.max(0, window.scrollY / max * 100)) : 0;
    if (progress) progress.style.width = `${pct}%`;
  }
  updateProgress();
  window.addEventListener("scroll", updateProgress, { passive: true });

  const revealObserver = new IntersectionObserver((entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) entry.target.classList.add("in-view");
    });
  }, { threshold: 0.14 });
  document.querySelectorAll(".reveal").forEach((el) => revealObserver.observe(el));
  requestAnimationFrame(() => {
    document.querySelectorAll(".reveal").forEach((el) => {
      const rect = el.getBoundingClientRect();
      if (rect.top < window.innerHeight * 0.92) el.classList.add("in-view");
    });
  });

  const navObserver = new IntersectionObserver((entries) => {
    const visible = entries
      .filter((entry) => entry.isIntersecting)
      .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
    if (!visible) return;
    navLinks.forEach((link) => link.classList.toggle("active", link.getAttribute("href") === `#${visible.target.id}`));
  }, { rootMargin: "-20% 0px -60% 0px", threshold: [0.12, 0.25, 0.5] });
  sections.forEach((section) => navObserver.observe(section));

  const tabCopy = {
    commercial: ["Commercial lens", "Validate whether revenue is durable enough to underwrite.", "Prioritize customer-level revenue, contract enforceability, retention cohorts, pricing power, top-account concentration, references, and pipeline conversion evidence."],
    financial: ["Financial lens", "Rebuild the case from source-backed drivers, not management-case ambition.", "Tie revenue recognition to invoices and cash, bridge gross margin and EBITDA, quantify working capital, and stress capex, liquidity, and downside case funding."],
    market: ["Market lens", "Separate attractive category narrative from budget-backed demand.", "Triangulate TAM/SAM/SOM, buyer urgency, competitive substitution, regulatory exposure, public/private comps, and exit-buyer depth."],
    ic: ["IC lens", "Convert diligence findings into a decision-ready sponsor narrative.", "Frame thesis, red flags, mitigants, valuation sensitivity, return bridge, 100-day plan, and explicit conditions precedent for investment committee approval."]
  };
  const stage = document.getElementById("worktabStage");
  document.querySelectorAll(".worktab").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".worktab").forEach((x) => x.classList.remove("active"));
      button.classList.add("active");
      const copy = tabCopy[button.dataset.tab] || tabCopy.commercial;
      if (stage) {
        stage.classList.remove("tab-pulse");
        stage.offsetHeight;
        stage.innerHTML = `<span>${escapeHtml(copy[0])}</span><strong>${escapeHtml(copy[1])}</strong><p>${escapeHtml(copy[2])}</p>`;
        stage.classList.add("tab-pulse");
      }
    });
  });
}
