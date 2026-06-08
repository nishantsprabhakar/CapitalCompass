const form = document.getElementById("dealForm");
const statusEl = document.getElementById("runStatus");
const downloads = document.getElementById("downloads");
const thesis = document.getElementById("thesis");
const summary = document.getElementById("summary");
const templateBox = document.getElementById("templateBox");

fetch("/api/template")
  .then((r) => r.json())
  .then((t) => {
    templateBox.innerHTML = `<strong>Default IC template</strong><br>${t.name}<br>${t.slides || 0} slides detected<br>Palette: ${(t.colors || []).slice(0, 4).map((c) => `<span style="display:inline-block;width:12px;height:12px;background:#${c};border:1px solid rgba(255,255,255,.35);vertical-align:-2px;margin-right:3px"></span>`).join("")}`;
  })
  .catch(() => { templateBox.textContent = "Template metadata unavailable."; });

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const button = form.querySelector("button[type=submit]");
  button.disabled = true;
  statusEl.textContent = "Running";
  downloads.textContent = "Generating files...";
  thesis.textContent = "Extracting materials, building diligence view, and preparing exports.";

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
  summary.innerHTML = [
    metric("Recommendation", a.recommendation),
    metric("Score", `${a.scorecard.total}/100`),
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
        <div class="card-label">Document control</div>
        <div class="doc-control">
          <strong>${included.length} included</strong>
          <span>${excluded.length} excluded as low relevance</span>
          <span>IC template: ${escapeHtml(a.template.uploadedName || a.template.name || "Default")}</span>
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
  downloads.innerHTML = files.map((f) => `<a href="${f.href}">${f.label}</a>`).join("");
}

function metric(label, value) {
  return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`;
}

function escapeHtml(value) {
  return String(value).replace(/[<>&'"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&#39;", '"': "&quot;" }[c]));
}
