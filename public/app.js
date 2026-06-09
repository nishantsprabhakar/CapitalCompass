const form = document.getElementById("dealForm");
const statusEl = document.getElementById("runStatus");
const downloads = document.getElementById("downloads");
const thesis = document.getElementById("thesis");
const summary = document.getElementById("summary");
const templateBox = document.getElementById("templateBox");
const enterpriseReview = document.getElementById("enterpriseReview");

initInterfaceEffects();

fetch("/api/template")
  .then((r) => r.json())
  .then((t) => {
    templateBox.classList.add("template-hidden");
    templateBox.textContent = "";
  })
  .catch(() => { templateBox.textContent = "Template metadata unavailable."; });

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const button = form.querySelector("button[type=submit]");
  button.disabled = true;
  statusEl.textContent = "Running";
  downloads.textContent = "Generating files...";
  thesis.textContent = "Extracting materials, building diligence view, and preparing exports.";
  enterpriseReview.textContent = "Benchmarking enterprise readiness, source quality, audit trail, and strategic-buyer fit.";

  try {
    const body = buildSubmissionBody();
    const response = await fetch("/api/analyze", { method: "POST", body });
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
