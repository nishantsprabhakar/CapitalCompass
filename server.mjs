import http from "node:http";
import https from "node:https";
import fs from "node:fs/promises";
import fssync from "node:fs";
import path from "node:path";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} from "docx";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const PptxGenJS = require("pptxgenjs");
const PUBLIC_DIR = path.join(__dirname, "public");
const WORK_DIR = path.join(__dirname, "work");
const OUTPUT_DIR = path.join(__dirname, "outputs");
const UPLOAD_DIR = path.join(WORK_DIR, "uploads");
const DEFAULT_TEMPLATE = path.join(WORK_DIR, "Pixxel Analysis_working v2.pptx");
const PORT = Number(process.env.PORT || 4174);

await fs.mkdir(PUBLIC_DIR, { recursive: true });
await fs.mkdir(UPLOAD_DIR, { recursive: true });
await fs.mkdir(OUTPUT_DIR, { recursive: true });

const stopWords = new Set("the and for with from into that this are was were have has had company business market revenue margin growth product customers customer management investment financial model data source sector in on of to a an by as is be or at its it their they".split(" "));

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);
    if (req.method === "GET" && url.pathname === "/") return serveFile(res, path.join(PUBLIC_DIR, "index.html"));
    if (req.method === "GET" && url.pathname.startsWith("/assets/")) return serveFile(res, path.join(PUBLIC_DIR, url.pathname.replace("/assets/", "")));
    if (req.method === "GET" && url.pathname === "/api/template") return json(res, await inspectTemplate(DEFAULT_TEMPLATE));
    if (req.method === "POST" && url.pathname === "/api/analyze") return handleAnalyze(req, res);
    if (req.method === "POST" && url.pathname === "/api/research") return handleResearch(req, res);
    if (req.method === "GET" && url.pathname.startsWith("/download/")) {
      const name = path.basename(decodeURIComponent(url.pathname.replace("/download/", "")));
      return serveFile(res, path.join(OUTPUT_DIR, name), true);
    }
    res.writeHead(404).end("Not found");
  } catch (error) {
    console.error(error);
    json(res, { error: error.message || String(error) }, 500);
  }
});

server.listen(PORT, () => {
  console.log(`Capital Compass running on http://localhost:${PORT}`);
});

async function handleAnalyze(req, res) {
  const contentType = req.headers["content-type"] || "";
  const contentLength = Number(req.headers["content-length"] || 0);
  if (contentLength > 150 * 1024 * 1024) {
    return json(res, {
      error: "Upload is too large for browser transfer. Paste the local diligence folder path instead, or upload a smaller selected set of key files."
    }, 413);
  }
  let fields = {};
  let uploaded = [];
  if (contentType.includes("multipart/form-data")) {
    const body = await readBody(req);
    const boundary = contentType.match(/boundary=(.+)$/)?.[1];
    if (!boundary) throw new Error("Missing multipart boundary.");
    ({ fields, files: uploaded } = await parseMultipart(body, boundary));
  } else {
    fields = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  }

  const sessionId = crypto.randomBytes(5).toString("hex");
  const companyName = clean(fields.companyName || "Target Company");
  const stage = fields.stage || "screening";
  const sourceUrls = splitLines(fields.sourceUrls || "");
  const folderPath = fields.folderPath ? String(fields.folderPath).trim() : "";
  const folderFiles = folderPath ? await listReadableFiles(folderPath) : [];
  const allFiles = [...uploaded, ...folderFiles]
    .filter(Boolean)
    .filter((file) => /\.(pptx|docx|xlsx|csv|txt|md)$/i.test(file.name || file.path))
    .slice(0, 140);
  const extracted = await extractCorpus(allFiles);
  const research = await researchUrls(sourceUrls);
  const template = await inspectTemplate(DEFAULT_TEMPLATE);
  const analysis = buildInvestmentAnalysis({ companyName, stage, extracted, research, template, fields });
  const files = [];

  const safe = slug(companyName);
  const notePath = path.join(OUTPUT_DIR, `${safe}-initial-screening-note-${sessionId}.docx`);
  await buildScreeningDocx(notePath, analysis);
  files.push({ label: "Initial Screening Note and Diligence Questions (.docx)", href: `/download/${path.basename(notePath)}` });

  if (stage === "deepDive" || stage === "full") {
    const deckPath = path.join(OUTPUT_DIR, `${safe}-pe-ic-memo-${sessionId}.pptx`);
    const modelPath = path.join(OUTPUT_DIR, `${safe}-financial-model-${sessionId}.xlsx`);
    await buildIcDeck(deckPath, analysis, template);
    await buildFinancialModel(modelPath, analysis);
    files.push({ label: "PE Grade IC Memo (.pptx)", href: `/download/${path.basename(deckPath)}` });
    files.push({ label: "Financial Model (.xlsx)", href: `/download/${path.basename(modelPath)}` });
  }

  json(res, { analysis, files });
}

async function handleResearch(req, res) {
  const payload = JSON.parse((await readBody(req)).toString("utf8") || "{}");
  const q = encodeURIComponent(payload.query || "");
  if (!q) return json(res, { results: [] });
  const ddg = await fetchText(`https://duckduckgo.com/html/?q=${q}`);
  const results = [...ddg.matchAll(/<a rel="nofollow" class="result__a" href="([^"]+)">([\s\S]*?)<\/a>/g)]
    .slice(0, 8)
    .map((m) => ({ url: decodeHtml(m[1]), title: stripTags(m[2]) }));
  json(res, { results });
}

async function parseMultipart(buffer, boundary) {
  const delimiter = Buffer.from(`--${boundary}`);
  const files = [];
  const fields = {};
  const parts = buffer.toString("binary").split(delimiter.toString("binary")).slice(1, -1);
  const uploadRoot = path.join(UPLOAD_DIR, crypto.randomBytes(6).toString("hex"));
  await fs.mkdir(uploadRoot, { recursive: true });
  for (const raw of parts) {
    const idx = raw.indexOf("\r\n\r\n");
    if (idx < 0) continue;
    const header = raw.slice(0, idx);
    let body = raw.slice(idx + 4);
    if (body.endsWith("\r\n")) body = body.slice(0, -2);
    const name = header.match(/name="([^"]+)"/)?.[1];
    const filename = header.match(/filename="([^"]*)"/)?.[1];
    if (!name) continue;
    if (filename) {
      const safeName = path.basename(filename.replaceAll("\\", "/"));
      if (!safeName) continue;
      const filePath = path.join(uploadRoot, safeName);
      await fs.writeFile(filePath, Buffer.from(body, "binary"));
      files.push({ path: filePath, name: safeName });
    } else {
      fields[name] = Buffer.from(body, "binary").toString("utf8").trim();
    }
  }
  return { fields, files };
}

async function listReadableFiles(folderPath) {
  const root = folderPath.replace(/^"+|"+$/g, "");
  const out = [];
  async function walk(dir) {
    const resolvedDir = path.resolve(dir);
    if (resolvedDir === path.resolve(UPLOAD_DIR)) return;
    if (resolvedDir.startsWith(path.resolve(UPLOAD_DIR) + path.sep)) return;
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const p = path.join(dir, entry.name);
      if (entry.isDirectory()) await walk(p);
      else if (/\.(pptx|docx|xlsx|csv|txt|md)$/i.test(entry.name)) out.push({ path: p, name: entry.name });
    }
  }
  await walk(root);
  return out.slice(0, 120);
}

async function extractCorpus(files) {
  const docs = [];
  for (const file of files) {
    try {
      const ext = path.extname(file.name || file.path).toLowerCase();
      const buf = await fs.readFile(file.path);
      let text = "";
      if (ext === ".pptx") text = await extractPptxText(buf);
      else if (ext === ".docx") text = await extractDocxText(buf);
      else if (ext === ".xlsx") text = await extractXlsxText(buf);
      else text = buf.toString("utf8");
      docs.push({ name: file.name || path.basename(file.path), type: ext.replace(".", ""), chars: text.length, text: normalize(text).slice(0, 90000) });
    } catch (error) {
      docs.push({ name: file.name || path.basename(file.path), type: "error", chars: 0, text: `Extraction failed: ${error.message}` });
    }
  }
  return docs;
}

async function extractPptxText(buf) {
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).sort(naturalSort);
  const chunks = [];
  for (const name of names) {
    const xml = await zip.file(name).async("text");
    const text = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeHtml(m[1])).join(" ");
    if (text.trim()) chunks.push(`${path.basename(name, ".xml")}: ${text}`);
  }
  return chunks.join("\n");
}

async function extractDocxText(buf) {
  const zip = await JSZip.loadAsync(buf);
  const file = zip.file("word/document.xml");
  if (!file) return "";
  const xml = await file.async("text");
  return [...xml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)].map((m) => decodeHtml(m[1])).join(" ");
}

async function extractXlsxText(buf) {
  const zip = await JSZip.loadAsync(buf);
  const shared = [];
  const ss = zip.file("xl/sharedStrings.xml");
  if (ss) {
    const xml = await ss.async("text");
    for (const m of xml.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)) shared.push(decodeHtml(m[1]));
  }
  const sheets = Object.keys(zip.files).filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n)).sort(naturalSort);
  const rows = [];
  for (const sheet of sheets) {
    const xml = await zip.file(sheet).async("text");
    rows.push(path.basename(sheet, ".xml"));
    for (const m of xml.matchAll(/<c[^>]*(?:t="s")?[^>]*>([\s\S]*?)<\/c>/g)) {
      const v = m[1].match(/<v>([\s\S]*?)<\/v>/)?.[1];
      if (v === undefined) continue;
      rows.push(shared[Number(v)] || v);
    }
  }
  return rows.join(" ");
}

async function inspectTemplate(templatePath) {
  if (!fssync.existsSync(templatePath)) return { available: false, name: "No template loaded", colors: ["17365D", "B48A2C", "F3F5F7", "2F3A45"], slides: 0 };
  const buf = await fs.readFile(templatePath);
  const zip = await JSZip.loadAsync(buf);
  const slideCount = Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
  const themeXml = zip.file("ppt/theme/theme1.xml") ? await zip.file("ppt/theme/theme1.xml").async("text") : "";
  const colors = [...themeXml.matchAll(/<a:srgbClr val="([0-9A-Fa-f]{6})"/g)].map((m) => m[1].toUpperCase());
  return {
    available: true,
    name: path.basename(templatePath),
    slides: slideCount,
    colors: [...new Set(colors)].slice(0, 8),
    defaultPath: templatePath
  };
}

function buildInvestmentAnalysis({ companyName, stage, extracted, research, template, fields }) {
  const rankedDocs = rankDocuments(companyName, extracted);
  const relevantDocs = rankedDocs.filter((d) => d.relevanceScore >= 3);
  const scopedDocs = relevantDocs.length ? relevantDocs : rankedDocs.slice(0, Math.min(5, rankedDocs.length));
  const corpus = scopedDocs.map((d) => d.text).join("\n");
  const words = keywordList(corpus, 18);
  const metrics = findMetrics(corpus);
  const docsSummary = rankedDocs.map((d) => ({
    name: d.name,
    type: d.type,
    chars: d.chars,
    relevanceScore: d.relevanceScore,
    status: scopedDocs.includes(d) ? "included" : "excluded"
  }));
  const excludedDocs = docsSummary.filter((d) => d.status === "excluded");
  const sector = clean(fields.sector || inferSector(corpus, words));
  const geography = clean(fields.geography || inferGeography(corpus));
  const evidence = extractEvidence(corpus, metrics);
  const thesis = buildThesis(companyName, corpus, words, sector, evidence);
  const risks = buildRisks(corpus, sector, evidence);
  const riskRegister = buildRiskRegister(companyName, sector, corpus, metrics, evidence);
  const questions = diligenceQuestions(companyName, sector, metrics, corpus);
  const scorecard = scoreCompany(corpus, metrics, research, evidence, riskRegister);
  const recommendation = scorecard.total >= 72
    ? "Proceed only to risk-focused confirmatory diligence"
    : scorecard.total >= 58
      ? "Proceed selectively; no IC approval until red flags are cleared"
      : "Hold pending material diligence evidence";
  return {
    generatedAt: new Date().toISOString(),
    companyName,
    stage,
    sector,
    geography,
    template,
    docsSummary,
    excludedDocs,
    research,
    words,
    metrics,
    evidence,
    thesis,
    risks,
    riskRegister,
    questions,
    scorecard,
    recommendation,
    memoSections: buildMemoSections(companyName, sector, geography, thesis, risks, metrics, scorecard, research, evidence, riskRegister)
  };
}

function buildThesis(companyName, corpus, words, sector, evidence) {
  return [
    `The current investment posture on ${companyName} should be skeptical-positive rather than approval-oriented: the materials suggest a potentially differentiated ${sector} asset, but the available evidence is not yet sufficient to underwrite revenue durability, margin quality, capital intensity, or exit depth.`,
    `The diligence burden is concentrated in ${evidence.missingEvidence.slice(0, 4).join(", ")}. These items should be treated as gating IC issues, not ordinary confirmatory diligence.`,
    `A PE process should prioritize downside protection first: customer concentration, contract enforceability, revenue recognition, satellite/manufacturing execution, regulatory exposure, and financing needs must be quantified before valuation work is decision-useful.`
  ];
}

function buildRisks(corpus, sector, evidence = { missingEvidence: [] }) {
  const base = [
    ["Forecast credibility", "Management case may rely on bookings conversion, pipeline timing, or margin expansion that has not yet been independently validated."],
    ["Customer concentration", "Revenue durability needs proof through cohort retention, renewal history, contract terms, and dependence on a small set of customers."],
    ["Technology differentiation", `For a ${sector} business, technical claims require third-party validation, product benchmarks, and evidence of switching costs.`],
    ["Working capital and cash needs", "Growth could consume cash through receivables, inventory, capex, milestone timing, or delayed customer collections."],
    ["Exit depth", "Strategic and sponsor exit appetite must be mapped against scale, profitability, public comps, and precedent transaction evidence."]
  ];
  if (/regulat|license|defen[cs]e|government|satellite|space/i.test(corpus)) base.push(["Regulatory and government exposure", "Government procurement, export controls, licenses, and national security considerations may affect timing, eligibility, and buyer universe."]);
  return base;
}

function buildRiskRegister(companyName, sector, corpus, metrics, evidence) {
  const riskRows = [
    risk("Critical", "Revenue quality and backlog conversion", "Management materials may show pipeline, customers, or contracted demand, but PE underwriting needs customer-level revenue, signed contract terms, renewal rights, cancellation rights, revenue recognition policy, and cash collection evidence.", "Request customer-level revenue bridge, top customer contracts, invoices, collections, deferred/unbilled revenue, churn/expansion cohorts, and pipeline conversion history."),
    risk("Critical", "Gross margin and EBITDA bridge", `The extracted metrics include ${Object.entries(metrics).map(([k, v]) => `${k}: ${v}`).join("; ") || "limited explicit financial metrics"}, but headline margins can be distorted by capitalization, launch costs, R&D treatment, one-offs, utilization, and support obligations.`, "Rebuild gross margin from raw COGS, satellite operations cost, cloud/processing cost, support cost, headcount allocation, capitalization policy, and normalized EBITDA adjustments."),
    risk("High", "Technology and product performance", `For a ${sector} company, technical differentiation must be independently validated rather than accepted from the company deck.`, "Run expert calls on hyperspectral resolution, revisit rates, calibration, latency, data quality, competing constellations, substitute datasets, and buyer willingness to pay."),
    risk("High", "Manufacturing, launch, and capex execution", "Growth may require satellite builds, launch slots, ground infrastructure, data-processing capacity, and working capital ahead of revenue realization.", "Validate build schedule, supplier dependencies, launch contracts, capex budget, failure rates, insurance, and contingency funding."),
    risk("High", "Customer concentration and procurement cyclicality", "Government, enterprise, and strategic customers can create lumpy revenue, long procurement cycles, pilot-to-production risk, and concentration risk.", "Ask for top 20 customers, customer references, procurement status, renewal dates, expansion pipeline, budget owner, and loss analysis."),
    risk("Medium", "Regulatory, licensing, and national security exposure", "Space, imagery, government, and cross-border data use can create approvals, export controls, data restrictions, or buyer-universe constraints.", "Map licenses, remote-sensing rules, export controls, data localization, government approvals, and change-of-control implications."),
    risk("Medium", "Exit and valuation support", "A differentiated story does not automatically create sponsor exit liquidity; exit depends on scale, profitability, strategic buyer appetite, public market comparables, and defensible growth.", "Build buyer universe, precedent transactions, public comps, scale thresholds, synergy rationale, and valuation sensitivity to exit multiple and EBITDA margin.")
  ];
  if (!/customer|contract|retention|renewal/i.test(corpus)) riskRows.unshift(risk("Critical", "Missing commercial proof", `${companyName} materials do not yet provide enough direct evidence of retention, renewal, customer concentration, and contract enforceability.`, "Make customer calls and customer-level data a gating item before any IC recommendation."));
  if (evidence.missingEvidence.length) riskRows.push(risk("High", "Evidence gaps in provided materials", `The current corpus lacks or under-discloses: ${evidence.missingEvidence.join(", ")}.`, "Treat each missing item as an explicit diligence request with owner, source document, and required deadline."));
  return riskRows;
}

function diligenceQuestions(companyName, sector, metrics, corpus) {
  const revenueMention = metrics.revenue || "not disclosed";
  return [
    section("Business Model and Market", [
      `What is the exact revenue model by product, customer segment, geography, contract type, and pricing unit? Current disclosed revenue reference: ${revenueMention}.`,
      "What is the TAM/SAM/SOM methodology, source base, bottom-up customer count, and evidence of budget availability?",
      "Which customer pain point is severe enough to support recurring spend, premium pricing, and long-term vendor retention?",
      "What portion of growth is market expansion versus share gain, price uplift, new product launch, or channel expansion?",
      "Which competitors are shortlisted in live deals, and what are win/loss reasons over the last 24 months?"
    ]),
    section("Commercial Diligence", [
      "Provide customer-level revenue for the last 36 months with churn, expansion, new logo, price, volume, and one-off revenue flags.",
      "List top 20 customers with ARR/revenue, start date, renewal date, gross margin, payment terms, contract length, termination rights, and pipeline expansion.",
      "Provide full sales funnel by stage, source, probability, expected close date, ACV, gross margin, and historical conversion rate.",
      "Share cohort retention, net revenue retention, gross revenue retention, and logo retention by vintage.",
      "Which reference customers can verify ROI, integration effort, switching cost, and competitive displacement?"
    ]),
    section("Financial Quality", [
      "Provide audited financials, monthly management accounts, trial balance, revenue waterfall, and EBITDA bridge for the last three years.",
      "Reconcile revenue recognition policy to contracts, invoices, cash collections, deferred revenue, unbilled revenue, and credit notes.",
      "Bridge gross margin by product, customer, region, utilization, input costs, warranty/support costs, and one-time items.",
      "Provide normalized EBITDA adjustments with source evidence, recurrence assessment, cash impact, and auditor treatment.",
      "Detail working capital seasonality, overdue receivables, inventory provisioning, customer advances, and supplier concentration."
    ]),
    section("Operations and Technology", [
      `What are the core proprietary assets, IP ownership, product roadmap, and build-versus-buy dependencies for ${companyName}?`,
      "Provide uptime, delivery accuracy, defect rates, SLA performance, implementation timeline, support backlog, and product usage metrics.",
      "Which suppliers, platforms, licenses, datasets, or infrastructure providers create single points of failure?",
      "Provide org chart, attrition, hiring plan, founder dependency, critical roles, compensation benchmarks, and succession risk.",
      "What capex, R&D, and implementation capacity is required to hit the plan?"
    ]),
    section("Legal, Regulatory, ESG", [
      "List all licenses, permits, export controls, government approvals, data privacy obligations, and pending renewals.",
      "Provide material contracts, litigation, related-party transactions, contingent liabilities, and change-of-control restrictions.",
      "Identify ESG, safety, security, environmental, labor, and supply-chain risks that could affect diligence, financing, or exit."
    ]),
    section("Deal, Valuation, and Exit", [
      "What transaction perimeter is contemplated: primary, secondary, control, minority, earn-out, rollover, option pool, and use of proceeds?",
      "What valuation expectation is implied by management, prior rounds, comparable transactions, DCF, and sponsor return requirements?",
      "Which value creation levers are under PE control in the first 100 days, 12 months, and hold period?",
      "What are credible exit routes, likely buyers, timing constraints, and valuation sensitivities?"
    ])
  ];
}

function buildMemoSections(companyName, sector, geography, thesis, risks, metrics, scorecard, research, evidence, riskRegister) {
  return [
    { title: "Executive Recommendation", body: `Recommendation: ${scorecard.total >= 72 ? "Proceed to focused diligence" : scorecard.total >= 58 ? "Proceed selectively" : "Hold"}. This is not IC-approval ready; the case requires evidence on revenue quality, customer durability, margin bridge, capital intensity, and exit depth before approval.`, bullets: thesis },
    { title: "Company Overview", body: `${companyName} operates in ${sector}${geography ? ` with relevance to ${geography}` : ""}. The company materials should be triangulated with customer references, third-party market work, and source-backed financials.`, bullets: ["Products and segments to be confirmed", "Revenue model and contract structure to be reconciled", "Management plan to be sensitized against downside case"] },
    { title: "Market Attractiveness", body: "Assess top-down market evidence only after bottom-up budget ownership, procurement cycle, and customer ROI have been validated.", bullets: ["TAM/SAM/SOM bridge", "Growth drivers and cyclicality", "Regulatory tailwinds and friction points", "Competitive intensity"] },
    { title: "Commercial Quality", body: "Commercial diligence should focus on repeatability of demand and proof that growth is not concentrated in one-off projects or early adopter customers.", bullets: ["Customer concentration", "Retention and expansion", "Pipeline conversion", "Pricing power"] },
    { title: "Financial Performance", body: `Extracted financial references include: ${Object.entries(metrics).slice(0, 8).map(([k, v]) => `${k}: ${v}`).join("; ") || "limited explicit metrics in provided materials"}.`, bullets: ["Quality of revenue", "Gross margin bridge", "EBITDA normalization", "Working capital and cash conversion"] },
    { title: "Risks and Mitigants", body: "The IC memo should treat risk as underwritable only when the diligence evidence shows magnitude, probability, owner, mitigation, and valuation impact.", bullets: riskRegister.slice(0, 7).map((r) => `${r.severity} - ${r.title}: ${r.whyItMatters}`) },
    { title: "Valuation and Returns", body: "Build valuation from entry multiple, DCF, precedent transactions, public comps, and sponsor return math. Sensitize entry valuation, revenue CAGR, EBITDA margin, leverage, and exit multiple.", bullets: ["Base/downside/upside cases", "IRR and MOIC bridge", "Debt capacity and covenant headroom", "Exit buyer universe"] },
    { title: "Evidence Gaps", body: "The following gaps should be assigned to diligence owners and tracked as explicit IC gating items.", bullets: evidence.missingEvidence },
    { title: "Source Research", body: research.length ? "External sources were captured for triangulation and citation." : "No external URLs were supplied. Use the research module to pull company site, filings, industry reports, customer evidence, and news.", bullets: research.slice(0, 6).map((r) => `${r.title || r.url}: ${r.summary}`) }
  ];
}

async function buildScreeningDocx(outPath, a) {
  const rows = [
    ["Recommendation", a.recommendation],
    ["Score", `${a.scorecard.total}/100`],
    ["Sector", a.sector],
    ["Geography", a.geography || "Not specified"],
    ["Documents included", a.docsSummary.filter((d) => d.status === "included").map((d) => `${d.name} (${d.type}, relevance ${d.relevanceScore})`).join("; ") || "None"],
    ["Documents excluded", a.excludedDocs?.length ? a.excludedDocs.map((d) => `${d.name} (relevance ${d.relevanceScore})`).join("; ") : "None"]
  ];
  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 720 } } },
      children: [
        para(a.companyName, 30, true, "17365D"),
        para("Initial Screening Note and Diligence Question List", 16, false, "666666"),
        para(`Generated ${new Date(a.generatedAt).toLocaleString()}`, 9, false, "777777"),
        table(rows),
        heading("0. IC Decision Gate"),
        table([
          ["Gate", "Current status", "Required evidence before IC approval"],
          ["Revenue quality", "Open", "Customer-level revenue bridge, signed contracts, churn/expansion cohorts, collections and revenue-recognition support."],
          ["Margin quality", "Open", "Gross margin bridge from raw COGS, capitalization policy, satellite ops cost, cloud/data cost, support cost, and normalized EBITDA adjustments."],
          ["Technical moat", "Open", "Expert validation of hyperspectral performance, revisit rate, calibration, latency, data quality, substitution risk, and customer willingness to pay."],
          ["Funding and capex", "Open", "Satellite build schedule, launch contracts, capex plan, insurance, contingency funding, and cash runway under downside case."],
          ["Exit support", "Open", "Strategic buyer universe, precedent transactions, public comps, scale thresholds, and valuation sensitivity."]
        ]),
        heading("1. Investment View"),
        ...a.thesis.map((x) => bullet(x)),
        heading("2. Scorecard"),
        table(Object.entries(a.scorecard.components).map(([k, v]) => [titleCase(k), `${v.score}/20`, v.rationale])),
        heading("3. Risk Register and Red Flags"),
        table([["Severity", "Risk", "Why it matters", "Diligence required"], ...a.riskRegister.map((r) => [r.severity, r.title, r.whyItMatters, r.diligenceRequired])]),
        heading("4. Evidence Gaps"),
        ...a.evidence.missingEvidence.map((x) => bullet(x)),
        heading("5. Immediate Diligence Workplan"),
        ...["Validate revenue quality and customer concentration", "Run customer calls and competitive referencing", "Rebuild management model from raw drivers", "Prepare legal/regulatory review", "Benchmark valuation and exit universe"].map((x, i) => bullet(`${i + 1}. ${x}`)),
        heading("6. Detailed Diligence Questions"),
        ...a.questions.flatMap((s) => [subheading(s.title), ...s.items.map((q) => bullet(q))]),
        heading("7. Source Research Captured"),
        ...(a.research.length ? a.research.map((r) => bullet(`${r.title || r.url}: ${r.summary}`)) : [bullet("No source URLs supplied yet.")])
      ]
    }]
  });
  await Packer.toBuffer(doc).then((b) => fs.writeFile(outPath, b));
}

async function buildIcDeck(outPath, a, template) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Capital Compass";
  pptx.subject = `${a.companyName} IC memo`;
  pptx.title = `${a.companyName} PE IC memo`;
  pptx.company = "Generated locally";
  pptx.theme = {
    headFontFace: "Calibri",
    bodyFontFace: "Calibri",
    lang: "en-US"
  };
  const assets = await getTemplateAssets(template.defaultPath);
  const C = {
    navy: "1F497D",
    gold: "CD9649",
    goldDark: "AF803E",
    pale: "F8F7F1",
    ink: "263238",
    grey: "6B7280",
    red: "A33A3A",
    green: "2F7D4C",
    asset: assets.hero
  };
  slideTitle(pptx, C, a.companyName, "Capital Compass IC Memorandum", [
    `Recommendation: ${a.recommendation}`,
    `Score: ${a.scorecard.total}/100`,
    `Template reference: ${template.name || "custom"}`
  ]);
  addAgenda(pptx, C);
  addExecSummary(pptx, C, a);
  addRiskRegisterDeck(pptx, C, a);
  addEvidenceGapDeck(pptx, C, a);
  addScorecard(pptx, C, a);
  for (const section of a.memoSections.slice(1, 7)) addTextSlide(pptx, C, section.title, section.body, section.bullets);
  addQuestionsSlide(pptx, C, a);
  addModelSlide(pptx, C, a);
  addSourcesSlide(pptx, C, a);
  await pptx.writeFile({ fileName: outPath });
}

function slideBase(pptx, C, title) {
  const s = pptx.addSlide();
  s.background = { color: "FFFFFF" };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.1, h: 7.5, fill: { color: C.gold }, line: { color: C.gold } });
  s.addShape(pptx.ShapeType.rect, { x: 0.35, y: 0.88, w: 8.95, h: 0.03, fill: { color: C.gold }, line: { color: C.gold } });
  for (let i = 0; i < 10; i++) {
    s.addShape(pptx.ShapeType.ellipse, { x: 0.65 + i * 1.08, y: 0.28 + (i % 3) * 0.08, w: 0.04, h: 0.04, fill: { color: C.gold }, line: { color: C.gold } });
  }
  s.addText(title, { x: 0.45, y: 0.42, w: 9.0, h: 0.34, fontFace: "Calibri", fontSize: 18, bold: true, color: C.goldDark, margin: 0 });
  s.addText("Capital Compass IC Analysis", { x: 10.1, y: 0.47, w: 2.5, h: 0.22, fontSize: 7.5, color: C.gold, align: "right", margin: 0 });
  s.addText("Capital in the Shadows | Nishant Prabhakar", { x: 9.4, y: 6.95, w: 3.2, h: 0.18, fontSize: 6.8, color: "9A9A9A", align: "right", margin: 0 });
  return s;
}

function slideTitle(pptx, C, title, subtitle, bullets) {
  const s = pptx.addSlide();
  s.background = { color: "FFFFFF" };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.1, h: 7.5, fill: { color: C.gold }, line: { color: C.gold } });
  if (C.asset) {
    s.addImage({ path: C.asset, x: 8.15, y: 0, w: 5.2, h: 7.5, transparency: 8 });
    s.addShape(pptx.ShapeType.rect, { x: 8.15, y: 0, w: 5.2, h: 7.5, fill: { color: "FFFFFF", transparency: 18 }, line: { color: "FFFFFF", transparency: 100 } });
  }
  for (let i = 0; i < 14; i++) s.addShape(pptx.ShapeType.ellipse, { x: 0.72 + i * 0.82, y: 0.36 + (i % 4) * 0.16, w: 0.04, h: 0.04, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText("CAPITAL COMPASS | INVESTMENT COMMITTEE ANALYSIS", { x: 0.5, y: 1.1, w: 6.5, h: 0.25, fontSize: 13, color: C.gold, margin: 0 });
  s.addShape(pptx.ShapeType.line, { x: 0.5, y: 1.68, w: 6.6, h: 0, line: { color: C.gold, width: 1.8 } });
  s.addText(title, { x: 0.5, y: 1.78, w: 7.2, h: 0.9, fontFace: "Calibri", fontSize: 34, bold: true, color: C.goldDark, margin: 0 });
  s.addText(subtitle, { x: 0.52, y: 2.86, w: 6.6, h: 0.35, fontSize: 18, color: C.gold, margin: 0 });
  bullets.forEach((b, i) => s.addText(b, { x: 0.55, y: 3.55 + i * 0.36, w: 6.8, h: 0.25, fontSize: 11, color: C.goldDark, breakLine: false, margin: 0 }));
  [["Shadow map", "risk-led"], ["PE model", "scenario-ready"], ["Compass", "IC direction"]].forEach((m, i) => {
    s.addShape(pptx.ShapeType.rect, { x: 0.28 + i * 2.35, y: 5.62, w: 1.95, h: 0.82, fill: { color: "FFFFFF" }, line: { color: C.gold, width: 0.75 } });
    s.addText(m[0], { x: 0.36 + i * 2.35, y: 5.78, w: 1.8, h: 0.2, fontSize: 13, bold: true, color: C.goldDark, align: "center", margin: 0 });
    s.addText(m[1], { x: 0.36 + i * 2.35, y: 6.12, w: 1.8, h: 0.16, fontSize: 7.5, color: C.gold, align: "center", margin: 0 });
  });
}

function addAgenda(pptx, C) {
  const s = slideBase(pptx, C, "IC memo structure");
  ["Recommendation and scorecard", "Company and market overview", "Commercial and financial diligence", "Risks, mitigants, valuation, returns", "Diligence questions and source appendix"].forEach((x, i) => {
    s.addShape(pptx.ShapeType.rect, { x: 0.75, y: 1.35 + i * 0.82, w: 0.45, h: 0.45, fill: { color: i === 0 ? C.gold : C.navy }, line: { color: "FFFFFF" } });
    s.addText(String(i + 1), { x: 0.75, y: 1.46 + i * 0.82, w: 0.45, h: 0.16, fontSize: 10, bold: true, color: "FFFFFF", align: "center", margin: 0 });
    s.addText(x, { x: 1.42, y: 1.41 + i * 0.82, w: 9.8, h: 0.28, fontSize: 16, color: C.ink, margin: 0 });
  });
}

function addExecSummary(pptx, C, a) {
  const s = slideBase(pptx, C, "Executive recommendation");
  card(s, C, 0.55, 1.2, 3.0, 1.35, "Recommendation", a.recommendation, C.gold);
  card(s, C, 3.85, 1.2, 2.1, 1.35, "Score", `${a.scorecard.total}/100`, scoreColor(a.scorecard.total, C));
  card(s, C, 6.25, 1.2, 2.6, 1.35, "Sector", a.sector, C.navy);
  card(s, C, 9.15, 1.2, 3.0, 1.35, "Current bias", a.scorecard.total >= 58 ? "Investable, diligence-led" : "Evidence gap", a.scorecard.total >= 58 ? C.green : C.red);
  s.addText("Investment thesis", { x: 0.65, y: 3.0, w: 2.5, h: 0.25, fontSize: 12, bold: true, color: C.ink, margin: 0 });
  s.addText(a.thesis.map((x) => `- ${x}`).join("\n"), { x: 0.65, y: 3.38, w: 5.8, h: 2.55, fontSize: 10.5, color: C.ink, valign: "top", fit: "shrink" });
  s.addText("Priority diligence burden", { x: 7.0, y: 3.0, w: 3.2, h: 0.25, fontSize: 12, bold: true, color: C.ink, margin: 0 });
  s.addText(a.risks.slice(0, 5).map((r) => `- ${r[0]}: ${r[1]}`).join("\n"), { x: 7.0, y: 3.38, w: 5.4, h: 2.55, fontSize: 10.2, color: C.ink, valign: "top", fit: "shrink" });
}

function addScorecard(pptx, C, a) {
  const s = slideBase(pptx, C, "Investment screen scorecard");
  const entries = Object.entries(a.scorecard.components);
  entries.forEach(([k, v], i) => {
    const y = 1.24 + i * 0.82;
    s.addText(titleCase(k), { x: 0.7, y, w: 2.4, h: 0.24, fontSize: 11, bold: true, color: C.ink, margin: 0 });
    s.addShape(pptx.ShapeType.rect, { x: 3.2, y: y + 0.03, w: 5.2, h: 0.16, fill: { color: "E7EBF1" }, line: { color: "E7EBF1" } });
    s.addShape(pptx.ShapeType.rect, { x: 3.2, y: y + 0.03, w: 5.2 * (v.score / 20), h: 0.16, fill: { color: scoreColor(v.score * 5, C) }, line: { color: scoreColor(v.score * 5, C) } });
    s.addText(`${v.score}/20`, { x: 8.62, y: y - 0.02, w: 0.65, h: 0.2, fontSize: 10, bold: true, color: C.ink, align: "right", margin: 0 });
    s.addText(v.rationale, { x: 9.55, y: y - 0.08, w: 2.75, h: 0.36, fontSize: 8.2, color: C.grey, fit: "shrink", margin: 0.02 });
  });
  s.addText("Screening interpretation", { x: 0.7, y: 5.65, w: 2.4, h: 0.22, fontSize: 11, bold: true, color: C.ink, margin: 0 });
  s.addText(a.recommendation, { x: 3.2, y: 5.62, w: 7.6, h: 0.28, fontSize: 13, bold: true, color: C.navy, margin: 0 });
}

function addRiskRegisterDeck(pptx, C, a) {
  const s = slideBase(pptx, C, "Red flag register");
  s.addText("Underwriting should remain risk-led until each red flag has a source-backed answer and quantified valuation impact.", { x: 0.55, y: 1.08, w: 11.5, h: 0.32, fontSize: 11.5, color: C.grey, margin: 0 });
  const rows = [["Severity", "Risk", "Why it matters", "Diligence required"], ...a.riskRegister.slice(0, 7).map((r) => [r.severity, r.title, r.whyItMatters, r.diligenceRequired])];
  s.addTable(rows, {
    x: 0.55, y: 1.55, w: 12.1, h: 4.85,
    border: { type: "solid", color: "E1D3BC", pt: 0.55 },
    fill: "FFFFFF",
    color: C.ink,
    fontFace: "Calibri",
    fontSize: 7.5,
    bold: false,
    valign: "mid",
    colW: [0.85, 2.1, 4.25, 4.9],
    margin: 0.05,
    fit: "shrink"
  });
  s.addShape(pptx.ShapeType.rect, { x: 0.55, y: 6.62, w: 4.2, h: 0.06, fill: { color: C.gold }, line: { color: C.gold } });
  s.addText("IC standard: risk is underwritable only with magnitude, probability, owner, mitigation, and valuation impact.", { x: 0.55, y: 6.78, w: 10.8, h: 0.18, fontSize: 8.5, color: C.goldDark, margin: 0 });
}

function addEvidenceGapDeck(pptx, C, a) {
  const s = slideBase(pptx, C, "Evidence gap tracker");
  const gaps = a.evidence.missingEvidence.slice(0, 10);
  gaps.forEach((gap, i) => {
    const x = i % 2 === 0 ? 0.65 : 6.55;
    const y = 1.22 + Math.floor(i / 2) * 0.82;
    s.addShape(pptx.ShapeType.rect, { x, y, w: 5.25, h: 0.56, fill: { color: i < 4 ? "FFF5EC" : "FFFFFF" }, line: { color: "E1D3BC", width: 0.65 } });
    s.addShape(pptx.ShapeType.rect, { x, y, w: 0.08, h: 0.56, fill: { color: i < 4 ? C.red : C.gold }, line: { color: i < 4 ? C.red : C.gold } });
    s.addText(i < 4 ? "GATING" : "OPEN", { x: x + 0.18, y: y + 0.12, w: 0.74, h: 0.12, fontSize: 6.2, bold: true, color: i < 4 ? C.red : C.goldDark, margin: 0 });
    s.addText(gap, { x: x + 1.02, y: y + 0.1, w: 4.0, h: 0.2, fontSize: 9.3, bold: true, color: C.ink, margin: 0, fit: "shrink" });
    s.addText("Required before IC approval", { x: x + 1.02, y: y + 0.36, w: 3.7, h: 0.12, fontSize: 6.8, color: C.grey, margin: 0 });
  });
  s.addText("Use this tracker as the diligence workplan: assign owner, source document, evidence standard, and deadline for each gap.", { x: 0.65, y: 6.18, w: 11.2, h: 0.35, fontSize: 11, color: C.goldDark, bold: true, margin: 0 });
}

function addTextSlide(pptx, C, title, body, bullets) {
  const s = slideBase(pptx, C, title);
  s.addText(body, { x: 0.68, y: 1.25, w: 11.8, h: 0.7, fontSize: 12, color: C.ink, breakLine: false, fit: "shrink", margin: 0.02 });
  bullets.slice(0, 7).forEach((b, i) => {
    const y = 2.22 + i * 0.56;
    s.addShape(pptx.ShapeType.rect, { x: 0.72, y: y + 0.03, w: 0.12, h: 0.12, fill: { color: C.gold }, line: { color: C.gold } });
    s.addText(String(b), { x: 1.02, y: y - 0.04, w: 10.9, h: 0.36, fontSize: 10.5, color: C.ink, fit: "shrink", margin: 0.01 });
  });
}

function addQuestionsSlide(pptx, C, a) {
  const s = slideBase(pptx, C, "Diligence question bank");
  a.questions.slice(0, 6).forEach((q, i) => {
    const x = i % 2 === 0 ? 0.65 : 6.75;
    const y = 1.18 + Math.floor(i / 2) * 1.65;
    s.addText(q.title, { x, y, w: 5.3, h: 0.22, fontSize: 10.5, bold: true, color: C.navy, margin: 0 });
    s.addText(q.items.slice(0, 3).map((item) => `- ${item}`).join("\n"), { x, y: y + 0.3, w: 5.4, h: 1.05, fontSize: 7.8, color: C.ink, fit: "shrink", margin: 0.02 });
  });
}

function addModelSlide(pptx, C, a) {
  const s = slideBase(pptx, C, "Financial model architecture");
  const blocks = [
    ["Inputs", "Transaction date, entry valuation, leverage, use of proceeds, operating drivers"],
    ["Revenue Build", "Segment revenue, volume, price, churn, pipeline conversion, cohort expansion"],
    ["Cost Build", "COGS, gross margin, opex, headcount, R&D, S&M, G&A"],
    ["Cash Flow", "EBITDA, capex, working capital, tax, FCF, debt paydown"],
    ["Returns", "Exit multiple, debt balance, equity value, MOIC, IRR"],
    ["Sensitivities", "Entry/exit multiple, revenue CAGR, EBITDA margin, leverage, downside case"]
  ];
  blocks.forEach((b, i) => card(s, C, 0.7 + (i % 3) * 4.05, 1.35 + Math.floor(i / 3) * 2.05, 3.55, 1.35, b[0], b[1], i === 0 ? C.gold : C.navy));
}

function addSourcesSlide(pptx, C, a) {
  const s = slideBase(pptx, C, "Sources and evidence register");
  const rows = a.research.length ? a.research.slice(0, 8).map((r, i) => [`${i + 1}`, r.title || r.url, r.url]) : [["1", "No external URLs supplied", "Use the research panel to add company, filing, market, customer, and news sources."]];
  s.addTable([["#", "Source", "URL"], ...rows], {
    x: 0.65, y: 1.25, w: 12, h: 4.8,
    border: { type: "solid", color: "DCE1EA", pt: 0.5 },
    fill: "FFFFFF",
    fontSize: 8,
    color: C.ink,
    fit: "shrink",
    colW: [0.45, 4.8, 6.75],
    margin: 0.04
  });
}

async function getTemplateAssets(templatePath) {
  if (!templatePath || !fssync.existsSync(templatePath)) return {};
  const assetDir = path.join(WORK_DIR, "template-assets");
  await fs.mkdir(assetDir, { recursive: true });
  const zip = await JSZip.loadAsync(await fs.readFile(templatePath));
  const media = Object.keys(zip.files)
    .filter((n) => /^ppt\/media\/image\d+\.(png|jpg|jpeg)$/i.test(n))
    .map((n) => ({ name: n, size: zip.files[n]._data?.uncompressedSize || 0 }))
    .sort((a, b) => b.size - a.size);
  if (!media.length) return {};
  const hero = media[0].name;
  const ext = path.extname(hero) || ".png";
  const out = path.join(assetDir, `template-hero${ext}`);
  await fs.writeFile(out, await zip.file(hero).async("nodebuffer"));
  return { hero: out };
}

async function buildFinancialModel(outPath, a) {
  const years = ["Actual Y-2", "Actual Y-1", "Actual Y0", "FY1", "FY2", "FY3", "FY4", "FY5"];
  const h = ["Metric", ...years];
  const sheets = [
    sheet("README", [
      ["Capital Compass Financial Model - Detailed Sponsor Model"],
      [`Company: ${a.companyName}`],
      ["Purpose: PE evaluation model with operating drivers, LBO returns, valuation, sensitivities, source register, risk register, and checks."],
      ["Instruction: replace assumption cells with source company model / diligence data; formulas are linked across tabs."],
      ["Recommendation", a.recommendation]
    ]),
    sheet("Source_Register", [["Source / item", "Status", "Use in model", "Owner / request"], ...a.docsSummary.map((d) => [d.name, d.status || "included", `${d.type}; relevance ${d.relevanceScore ?? ""}; ${d.chars} chars`, "Archive and reconcile"]), ...a.evidence.missingEvidence.map((g) => [g, "Missing", "Required for underwriting", "Diligence request"])]),
    sheet("Risk_Register", [["Severity", "Risk", "Why it matters", "Diligence required"], ...a.riskRegister.slice(0, 12).map((r) => [r.severity, r.title, r.whyItMatters, r.diligenceRequired])]),
    sheet("Control_Panel", [["Input", "Value", "Comment"], ["Entry EBITDA multiple", 12.0, "Replace with process valuation"], ["Exit EBITDA multiple", 11.0, "Base case exit"], ["Debt / EBITDA at close", 3.0, "Opening leverage"], ["Cash interest rate", 0.11, "Blended cost of debt"], ["Tax rate", 0.25, "Cash tax"], ["Minimum cash", 5.0, "Cash floor"], ["Sponsor fees % EV", 0.025, "Transaction costs"], ["Case selector", "Base", "Base / Downside / Upside"]]),
    sheet("Assumptions", [["Driver", ...years], ["Customers / logos", 20, 35, 50, 72, 98, 128, 160, 195], ["Revenue / customer", 1.1, 1.35, 1.7, 2.1, 2.55, 3.0, 3.35, 3.65], ["Gross margin", 0.35, 0.44, 0.52, 0.56, 0.60, 0.63, 0.65, 0.66], ["R&D % revenue", 0.34, 0.30, 0.24, 0.20, 0.17, 0.15, 0.13, 0.12], ["S&M % revenue", 0.26, 0.24, 0.22, 0.20, 0.18, 0.16, 0.15, 0.14], ["G&A % revenue", 0.18, 0.16, 0.14, 0.13, 0.12, 0.11, 0.10, 0.095], ["Capex % revenue", 0.45, 0.38, 0.30, 0.24, 0.20, 0.17, 0.14, 0.12], ["NWC % revenue", 0.24, 0.23, 0.22, 0.21, 0.20, 0.19, 0.18, 0.18]]),
    sheet("Revenue_Build", [h, ["Customers / logos", "=Assumptions!B2", "=Assumptions!C2", "=Assumptions!D2", "=Assumptions!E2", "=Assumptions!F2", "=Assumptions!G2", "=Assumptions!H2", "=Assumptions!I2"], ["Revenue / customer", "=Assumptions!B3", "=Assumptions!C3", "=Assumptions!D3", "=Assumptions!E3", "=Assumptions!F3", "=Assumptions!G3", "=Assumptions!H3", "=Assumptions!I3"], ["Core imagery revenue", "=B2*B3", "=C2*C3", "=D2*D3", "=E2*E3", "=F2*F3", "=G2*G3", "=H2*H3", "=I2*I3"], ["Analytics attach %", 0.10, 0.13, 0.17, 0.22, 0.27, 0.31, 0.34, 0.36], ["Analytics revenue", "=B4*B5", "=C4*C5", "=D4*D5", "=E4*E5", "=F4*F5", "=G4*G5", "=H4*H5", "=I4*I5"], ["Government / strategic projects", 8, 12, 18, 25, 31, 36, 40, 45], ["Total revenue", "=B4+B6+B7", "=C4+C6+C7", "=D4+D6+D7", "=E4+E6+E7", "=F4+F6+F7", "=G4+G6+G7", "=H4+H6+H7", "=I4+I6+I7"], ["Revenue growth", "", "=C8/B8-1", "=D8/C8-1", "=E8/D8-1", "=F8/E8-1", "=G8/F8-1", "=H8/G8-1", "=I8/H8-1"], ["Contracted revenue %", 0.35, 0.42, 0.50, 0.56, 0.61, 0.65, 0.68, 0.70], ["Contracted revenue", "=B8*B10", "=C8*C10", "=D8*D10", "=E8*E10", "=F8*F10", "=G8*G10", "=H8*H10", "=I8*I10"]]),
    sheet("COGS_Margin", [h, ["Revenue", "=Revenue_Build!B8", "=Revenue_Build!C8", "=Revenue_Build!D8", "=Revenue_Build!E8", "=Revenue_Build!F8", "=Revenue_Build!G8", "=Revenue_Build!H8", "=Revenue_Build!I8"], ["Satellite ops cost", "=B2*18%", "=C2*16%", "=D2*14%", "=E2*13%", "=F2*12%", "=G2*11%", "=H2*10%", "=I2*10%"], ["Cloud/data processing", "=B2*12%", "=C2*11%", "=D2*10%", "=E2*9%", "=F2*8%", "=G2*8%", "=H2*7%", "=I2*7%"], ["Support / delivery COGS", "=B2*35%", "=C2*29%", "=D2*24%", "=E2*22%", "=F2*20%", "=G2*18%", "=H2*18%", "=I2*17%"], ["Total COGS", "=SUM(B3:B5)", "=SUM(C3:C5)", "=SUM(D3:D5)", "=SUM(E3:E5)", "=SUM(F3:F5)", "=SUM(G3:G5)", "=SUM(H3:H5)", "=SUM(I3:I5)"], ["Gross profit", "=B2-B6", "=C2-C6", "=D2-D6", "=E2-E6", "=F2-F6", "=G2-G6", "=H2-H6", "=I2-I6"], ["Gross margin", "=B7/B2", "=C7/C2", "=D7/D2", "=E7/E2", "=F7/F2", "=G7/G2", "=H7/H2", "=I7/I2"]]),
    sheet("Opex_Headcount", [h, ["Revenue", "=Revenue_Build!B8", "=Revenue_Build!C8", "=Revenue_Build!D8", "=Revenue_Build!E8", "=Revenue_Build!F8", "=Revenue_Build!G8", "=Revenue_Build!H8", "=Revenue_Build!I8"], ["R&D", "=B2*Assumptions!B5", "=C2*Assumptions!C5", "=D2*Assumptions!D5", "=E2*Assumptions!E5", "=F2*Assumptions!F5", "=G2*Assumptions!G5", "=H2*Assumptions!H5", "=I2*Assumptions!I5"], ["Sales & marketing", "=B2*Assumptions!B6", "=C2*Assumptions!C6", "=D2*Assumptions!D6", "=E2*Assumptions!E6", "=F2*Assumptions!F6", "=G2*Assumptions!G6", "=H2*Assumptions!H6", "=I2*Assumptions!I6"], ["G&A", "=B2*Assumptions!B7", "=C2*Assumptions!C7", "=D2*Assumptions!D7", "=E2*Assumptions!E7", "=F2*Assumptions!F7", "=G2*Assumptions!G7", "=H2*Assumptions!H7", "=I2*Assumptions!I7"], ["Total opex", "=SUM(B3:B5)", "=SUM(C3:C5)", "=SUM(D3:D5)", "=SUM(E3:E5)", "=SUM(F3:F5)", "=SUM(G3:G5)", "=SUM(H3:H5)", "=SUM(I3:I5)"], ["Employees", 120, 180, 250, 310, 365, 420, 465, 500], ["Revenue / employee", "=B2/B7", "=C2/C7", "=D2/D7", "=E2/E7", "=F2/F7", "=G2/G7", "=H2/H7", "=I2/I7"]]),
    sheet("Capex_NWC", [h, ["Revenue", "=Revenue_Build!B8", "=Revenue_Build!C8", "=Revenue_Build!D8", "=Revenue_Build!E8", "=Revenue_Build!F8", "=Revenue_Build!G8", "=Revenue_Build!H8", "=Revenue_Build!I8"], ["Growth capex", "=B2*Assumptions!B8", "=C2*Assumptions!C8", "=D2*Assumptions!D8", "=E2*Assumptions!E8", "=F2*Assumptions!F8", "=G2*Assumptions!G8", "=H2*Assumptions!H8", "=I2*Assumptions!I8"], ["Maintenance capex", "=B2*4%", "=C2*4%", "=D2*4%", "=E2*4%", "=F2*4%", "=G2*4%", "=H2*4%", "=I2*4%"], ["Total capex", "=B3+B4", "=C3+C4", "=D3+D4", "=E3+E4", "=F3+F4", "=G3+G4", "=H3+H4", "=I3+I4"], ["NWC balance", "=B2*Assumptions!B9", "=C2*Assumptions!C9", "=D2*Assumptions!D9", "=E2*Assumptions!E9", "=F2*Assumptions!F9", "=G2*Assumptions!G9", "=H2*Assumptions!H9", "=I2*Assumptions!I9"], ["Change in NWC", "=0", "=C6-B6", "=D6-C6", "=E6-D6", "=F6-E6", "=G6-F6", "=H6-G6", "=I6-H6"]]),
    sheet("Income_Statement", [h, ["Revenue", "=Revenue_Build!B8", "=Revenue_Build!C8", "=Revenue_Build!D8", "=Revenue_Build!E8", "=Revenue_Build!F8", "=Revenue_Build!G8", "=Revenue_Build!H8", "=Revenue_Build!I8"], ["Gross profit", "=COGS_Margin!B7", "=COGS_Margin!C7", "=COGS_Margin!D7", "=COGS_Margin!E7", "=COGS_Margin!F7", "=COGS_Margin!G7", "=COGS_Margin!H7", "=COGS_Margin!I7"], ["Opex", "=Opex_Headcount!B6", "=Opex_Headcount!C6", "=Opex_Headcount!D6", "=Opex_Headcount!E6", "=Opex_Headcount!F6", "=Opex_Headcount!G6", "=Opex_Headcount!H6", "=Opex_Headcount!I6"], ["EBITDA", "=B3-B4", "=C3-C4", "=D3-D4", "=E3-E4", "=F3-F4", "=G3-G4", "=H3-H4", "=I3-I4"], ["EBITDA margin", "=B5/B2", "=C5/C2", "=D5/D2", "=E5/E2", "=F5/F2", "=G5/G2", "=H5/H2", "=I5/I2"], ["D&A", "=Capex_NWC!B5*18%", "=Capex_NWC!C5*18%", "=Capex_NWC!D5*18%", "=Capex_NWC!E5*18%", "=Capex_NWC!F5*18%", "=Capex_NWC!G5*18%", "=Capex_NWC!H5*18%", "=Capex_NWC!I5*18%"], ["EBIT", "=B5-B7", "=C5-C7", "=D5-D7", "=E5-E7", "=F5-F7", "=G5-G7", "=H5-H7", "=I5-I7"], ["Cash tax", "=MAX(0,B8*Control_Panel!B6)", "=MAX(0,C8*Control_Panel!B6)", "=MAX(0,D8*Control_Panel!B6)", "=MAX(0,E8*Control_Panel!B6)", "=MAX(0,F8*Control_Panel!B6)", "=MAX(0,G8*Control_Panel!B6)", "=MAX(0,H8*Control_Panel!B6)", "=MAX(0,I8*Control_Panel!B6)"]]),
    sheet("Cash_Flow", [h, ["EBITDA", "=Income_Statement!B5", "=Income_Statement!C5", "=Income_Statement!D5", "=Income_Statement!E5", "=Income_Statement!F5", "=Income_Statement!G5", "=Income_Statement!H5", "=Income_Statement!I5"], ["Cash tax", "=Income_Statement!B9", "=Income_Statement!C9", "=Income_Statement!D9", "=Income_Statement!E9", "=Income_Statement!F9", "=Income_Statement!G9", "=Income_Statement!H9", "=Income_Statement!I9"], ["Capex", "=Capex_NWC!B5", "=Capex_NWC!C5", "=Capex_NWC!D5", "=Capex_NWC!E5", "=Capex_NWC!F5", "=Capex_NWC!G5", "=Capex_NWC!H5", "=Capex_NWC!I5"], ["Change in NWC", "=Capex_NWC!B7", "=Capex_NWC!C7", "=Capex_NWC!D7", "=Capex_NWC!E7", "=Capex_NWC!F7", "=Capex_NWC!G7", "=Capex_NWC!H7", "=Capex_NWC!I7"], ["Unlevered FCF", "=B2-B3-B4-B5", "=C2-C3-C4-C5", "=D2-D3-D4-D5", "=E2-E3-E4-E5", "=F2-F3-F4-F5", "=G2-G3-G4-G5", "=H2-H3-H4-H5", "=I2-I3-I4-I5"]]),
    sheet("Debt_Schedule", [h, ["Opening debt", 0, 0, 0, "=Valuation_Returns!B8", "=MAX(0,E2+E3-E4-E5)", "=MAX(0,F2+F3-F4-F5)", "=MAX(0,G2+G3-G4-G5)", "=MAX(0,H2+H3-H4-H5)"], ["Interest expense", "=B2*Control_Panel!B5", "=C2*Control_Panel!B5", "=D2*Control_Panel!B5", "=E2*Control_Panel!B5", "=F2*Control_Panel!B5", "=G2*Control_Panel!B5", "=H2*Control_Panel!B5", "=I2*Control_Panel!B5"], ["Mandatory amortization", 0, 0, 0, "=E2*5%", "=F2*5%", "=G2*5%", "=H2*5%", "=I2*5%"], ["Cash sweep", 0, 0, 0, "=MAX(0,Cash_Flow!E6-E4)", "=MAX(0,Cash_Flow!F6-F4)", "=MAX(0,Cash_Flow!G6-G4)", "=MAX(0,Cash_Flow!H6-H4)", "=MAX(0,Cash_Flow!I6-I4)"], ["Ending debt", "=MAX(0,B2+B3-B4-B5)", "=MAX(0,C2+C3-C4-C5)", "=MAX(0,D2+D3-D4-D5)", "=MAX(0,E2+E3-E4-E5)", "=MAX(0,F2+F3-F4-F5)", "=MAX(0,G2+G3-G4-G5)", "=MAX(0,H2+H3-H4-H5)", "=MAX(0,I2+I3-I4-I5)"], ["Net leverage", "=IFERROR(B6/Income_Statement!B5,0)", "=IFERROR(C6/Income_Statement!C5,0)", "=IFERROR(D6/Income_Statement!D5,0)", "=IFERROR(E6/Income_Statement!E5,0)", "=IFERROR(F6/Income_Statement!F5,0)", "=IFERROR(G6/Income_Statement!G5,0)", "=IFERROR(H6/Income_Statement!H5,0)", "=IFERROR(I6/Income_Statement!I5,0)"]]),
    sheet("Valuation_Returns", [["Input", "Value"], ["Entry EBITDA", "=Income_Statement!D5"], ["Entry multiple", "=Control_Panel!B2"], ["Enterprise value", "=B2*B3"], ["Net debt / cash", 0], ["Equity purchase price", "=B4-B5"], ["Debt at close", "=B2*Control_Panel!B4"], ["Sponsor equity", "=B6-B7"], ["Exit EBITDA", "=Income_Statement!I5"], ["Exit multiple", "=Control_Panel!B3"], ["Exit enterprise value", "=B9*B10"], ["Debt at exit", "=Debt_Schedule!I6"], ["Exit equity value", "=B11-B12"], ["MOIC", "=B13/B8"], ["IRR", "=POWER(B14,1/5)-1"], ["Entry EV / Revenue", "=B4/Revenue_Build!D8"], ["Exit EV / Revenue", "=B11/Revenue_Build!I8"]]),
    sheet("Sensitivity", [["Exit Multiple / FY5 EBITDA Margin", "10.0x", "11.0x", "12.0x", "13.0x", "14.0x"], ["18%", "=(B$1*Revenue_Build!I8*18%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(C$1*Revenue_Build!I8*18%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(D$1*Revenue_Build!I8*18%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(E$1*Revenue_Build!I8*18%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(F$1*Revenue_Build!I8*18%-Valuation_Returns!B12)/Valuation_Returns!B8"], ["22%", "=(B$1*Revenue_Build!I8*22%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(C$1*Revenue_Build!I8*22%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(D$1*Revenue_Build!I8*22%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(E$1*Revenue_Build!I8*22%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(F$1*Revenue_Build!I8*22%-Valuation_Returns!B12)/Valuation_Returns!B8"], ["26%", "=(B$1*Revenue_Build!I8*26%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(C$1*Revenue_Build!I8*26%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(D$1*Revenue_Build!I8*26%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(E$1*Revenue_Build!I8*26%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(F$1*Revenue_Build!I8*26%-Valuation_Returns!B12)/Valuation_Returns!B8"], ["30%", "=(B$1*Revenue_Build!I8*30%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(C$1*Revenue_Build!I8*30%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(D$1*Revenue_Build!I8*30%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(E$1*Revenue_Build!I8*30%-Valuation_Returns!B12)/Valuation_Returns!B8", "=(F$1*Revenue_Build!I8*30%-Valuation_Returns!B12)/Valuation_Returns!B8"]]),
    sheet("Dashboard", [["Metric", "Value", "Comment"], ["Recommendation", a.recommendation, "Risk-adjusted process view"], ["Screen score", a.scorecard.total, "Capped for evidence gaps"], ["FY5 revenue", "=Revenue_Build!I8", "Model output"], ["FY5 EBITDA", "=Income_Statement!I5", "Model output"], ["FY5 EBITDA margin", "=Income_Statement!I6", "Model output"], ["MOIC", "=Valuation_Returns!B14", "Sponsor return"], ["IRR", "=Valuation_Returns!B15", "Sponsor return"], ["Critical risks", a.riskRegister.filter((r) => r.severity === "Critical").length, "Must clear before IC"]]),
    sheet("Checks", [["Check", "Status"], ["Revenue positive", '=IF(MIN(Revenue_Build!B8:I8)>0,"OK","Check")'], ["Gross profit ties", '=IF(ABS(COGS_Margin!I7-(Revenue_Build!I8-COGS_Margin!I6))<0.01,"OK","Check")'], ["EBITDA ties", '=IF(ABS(Income_Statement!I5-(COGS_Margin!I7-Opex_Headcount!I6))<0.01,"OK","Check")'], ["MOIC computable", '=IF(ISERROR(Valuation_Returns!B14),"Check","OK")'], ["Exit equity positive", '=IF(Valuation_Returns!B13>0,"OK","Check")'], ["Evidence gaps", a.evidence.missingEvidence.length ? "Open diligence gaps" : "OK"]])
  ];
  await writeXlsx(outPath, sheets);
}

async function writeXlsx(outPath, sheets) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>${sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  zip.folder("_rels").file(".rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  zip.folder("xl").file("workbook.xml", `<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${xml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`);
  zip.folder("xl").folder("_rels").file("workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  zip.folder("xl").file("styles.xml", `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font><sz val="10"/><name val="Aptos"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font><font><b/><sz val="12"/><color rgb="FF17365D"/><name val="Aptos"/></font></fonts><fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFD6DCE6"/></left><right style="thin"><color rgb="FFD6DCE6"/></right><top style="thin"><color rgb="FFD6DCE6"/></top><bottom style="thin"><color rgb="FFD6DCE6"/></bottom></border></borders><cellXfs count="4"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/><xf fontId="0" fillId="3" borderId="1" applyFill="1" applyBorder="1"/><xf fontId="2" fillId="0" borderId="0" applyFont="1"/></cellXfs></styleSheet>`);
  const ws = zip.folder("xl").folder("worksheets");
  sheets.forEach((s, i) => ws.file(`sheet${i + 1}.xml`, worksheetXml(s.rows)));
  await fs.writeFile(outPath, await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }));
}

function worksheetXml(rows) {
  return `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${Array.from({ length: 10 }, (_, i) => `<col min="${i + 1}" max="${i + 1}" width="${i === 0 ? 28 : 14}" customWidth="1"/>`).join("")}</cols><sheetData>${rows.map((row, r) => `<row r="${r + 1}">${row.map((v, c) => cellXml(v, r, c)).join("")}</row>`).join("")}</sheetData></worksheet>`;
}

function cellXml(value, r, c) {
  const ref = `${col(c)}${r + 1}`;
  const style = r === 0 ? 1 : c === 0 ? 2 : 0;
  if (typeof value === "number") return `<c r="${ref}" s="${style}"><v>${value}</v></c>`;
  if (typeof value === "string" && value.startsWith("=")) return `<c r="${ref}" s="${style}"><f>${xml(value.slice(1))}</f></c>`;
  return `<c r="${ref}" t="inlineStr" s="${style}"><is><t>${xml(value ?? "")}</t></is></c>`;
}

async function researchUrls(urls) {
  const out = [];
  for (const url of urls.slice(0, 12)) {
    try {
      const text = await fetchText(url);
      const title = stripTags((text.match(/<title[^>]*>([\s\S]*?)<\/title>/i)?.[1] || url).trim()).slice(0, 120);
      const cleanText = normalize(stripTags(text)).slice(0, 1200);
      out.push({ url, title, summary: cleanText.slice(0, 420) });
    } catch (error) {
      out.push({ url, title: "Fetch failed", summary: error.message });
    }
  }
  return out;
}

function fetchText(url) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith("https") ? https : http;
    const req = lib.get(url, { headers: { "User-Agent": "Mozilla/5.0 Capital-Compass" }, timeout: 12000 }, (r) => {
      if (r.statusCode >= 300 && r.statusCode < 400 && r.headers.location) return resolve(fetchText(new URL(r.headers.location, url).toString()));
      let data = "";
      r.setEncoding("utf8");
      r.on("data", (d) => { data += d; if (data.length > 700000) req.destroy(); });
      r.on("end", () => resolve(data));
    });
    req.on("error", reject);
    req.on("timeout", () => req.destroy(new Error("Request timed out")));
  });
}

function scoreCompany(corpus, metrics, research, evidence, riskRegister) {
  const components = {
    market: componentScore(corpus, ["tam", "market", "growth", "demand", "industry"], research.length ? 11 : 8, "Market claims need third-party sizing and budget validation."),
    differentiation: componentScore(corpus, ["proprietary", "patent", "technology", "platform", "unique", "moat"], 9, "Differentiation requires technical and customer proof."),
    commercial: componentScore(corpus, ["customer", "contract", "pipeline", "retention", "renewal", "arr"], 7, "Commercial quality depends on cohort and customer-level evidence."),
    financial: componentScore(corpus, ["revenue", "ebitda", "gross margin", "cash", "profit", "working capital"], Object.keys(metrics).length ? 10 : 6, "Financial disclosure needs source reconciliation."),
    exit: componentScore(corpus, ["acquisition", "strategic", "ipo", "exit", "valuation", "multiple"], 6, "Exit depth is not yet fully evidenced.")
  };
  const evidencePenalty = Math.min(18, Math.max(0, evidence.missingEvidence.length - 3) * 3);
  const criticalPenalty = riskRegister.filter((r) => r.severity === "Critical").length * 2;
  const total = Math.max(25, Object.values(components).reduce((s, c) => s + c.score, 0) - evidencePenalty - criticalPenalty);
  return { total, components };
}

function componentScore(corpus, needles, base, fallback) {
  const hits = needles.reduce((n, k) => n + (new RegExp(k, "i").test(corpus) ? 1 : 0), 0);
  return { score: Math.max(4, Math.min(18, base + hits * 1.5)), rationale: hits ? `Mentions found for ${hits}/${needles.length} screen factors, but score remains capped until source-backed diligence evidence is provided.` : fallback };
}

function rankDocuments(companyName, docs) {
  const tokens = companyTokens(companyName);
  return docs.map((doc) => {
    const haystack = `${doc.name} ${doc.text}`.toLowerCase();
    let score = 0;
    for (const token of tokens) {
      const matches = haystack.match(new RegExp(escapeRegex(token), "g"))?.length || 0;
      score += Math.min(8, matches * 2);
    }
    if (tokens.some((t) => doc.name.toLowerCase().includes(t))) score += 8;
    if (/\b(ic|investment|analysis|memo|deck|model|financial|cim|teaser)\b/i.test(doc.name)) score += 1;
    if (/quanta|vaccine|bharat biotech|ibm/i.test(haystack) && !/pixxel/i.test(haystack)) score -= 10;
    return { ...doc, relevanceScore: score };
  }).sort((a, b) => b.relevanceScore - a.relevanceScore || b.chars - a.chars);
}

function companyTokens(companyName) {
  return clean(companyName).toLowerCase().split(/[^a-z0-9]+/).filter((t) => t.length > 2);
}

function extractEvidence(corpus, metrics) {
  const checks = [
    ["customer-level revenue", /top \d+ customers|customer[- ]level|cohort|retention|renewal|nrr|grr/i],
    ["signed contracts and contract terms", /contract|signed|renewal|termination|purchase order|po\b|msa|sla/i],
    ["audited historical financials", /audited|audit|income statement|balance sheet|cash flow|trial balance/i],
    ["gross margin bridge", /gross margin|cogs|cost of goods|contribution margin|unit economics/i],
    ["working capital and cash conversion", /working capital|receivable|payable|inventory|cash conversion|collections/i],
    ["capex and funding plan", /capex|capital expenditure|launch cost|satellite build|funding|cash runway/i],
    ["market sizing support", /tam|sam|som|market size|market sizing|cagr/i],
    ["competitive win/loss evidence", /competitor|competition|win.?loss|win rate|substitute/i],
    ["regulatory/license mapping", /license|regulatory|approval|export control|government|remote sensing/i],
    ["exit comparables", /precedent|comparable|public comp|strategic buyer|exit multiple|acquisition/i]
  ];
  const present = checks.filter(([, re]) => re.test(corpus)).map(([label]) => label);
  const missingEvidence = checks.filter(([, re]) => !re.test(corpus)).map(([label]) => label);
  if (!Object.keys(metrics).length) missingEvidence.unshift("source-backed financial metrics");
  return { present, missingEvidence };
}

function risk(severity, title, whyItMatters, diligenceRequired) {
  return { severity, title, whyItMatters, diligenceRequired };
}

function findMetrics(text) {
  const out = {};
  const patterns = {
    revenue: /(revenue|sales|arr)[^\n]{0,50}?((?:\$|INR|Rs\.?|USD)?\s?\d+(?:\.\d+)?\s?(?:m|mn|million|cr|crore|bn|billion)?)/i,
    margin: /(gross margin|ebitda margin|margin)[^\n]{0,50}?(\d+(?:\.\d+)?\s?%)/i,
    growth: /(growth|cagr)[^\n]{0,50}?(\d+(?:\.\d+)?\s?%)/i,
    customers: /(customers|clients)[^\n]{0,50}?(\d+(?:,\d{3})*)/i,
    valuation: /(valuation|enterprise value|ev)[^\n]{0,50}?((?:\$|INR|Rs\.?|USD)?\s?\d+(?:\.\d+)?\s?(?:m|mn|million|cr|crore|bn|billion)?)/i
  };
  for (const [k, re] of Object.entries(patterns)) {
    const m = text.match(re);
    if (m) out[k] = `${m[1]} ${m[2]}`.trim();
  }
  return out;
}

function keywordList(text, limit) {
  const counts = new Map();
  normalize(text).toLowerCase().match(/[a-z][a-z0-9-]{3,}/g)?.forEach((w) => {
    if (!stopWords.has(w)) counts.set(w, (counts.get(w) || 0) + 1);
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, limit).map(([word, count]) => ({ word, count }));
}

function findSentences(text, needles) {
  return normalize(text).split(/(?<=[.!?])\s+/).filter((s) => needles.some((n) => s.toLowerCase().includes(n))).map((s) => s.slice(0, 280));
}

function inferSector(corpus, words) {
  if (/satellite|space|earth observation|geospatial|imagery/i.test(corpus)) return "space technology / geospatial intelligence";
  if (/saas|software|platform|subscription/i.test(corpus)) return "software / technology";
  if (/manufactur|factory|supply chain/i.test(corpus)) return "industrial technology";
  return words[0]?.word ? `${words[0].word} ecosystem` : "target sector";
}

function inferGeography(corpus) {
  if (/india|indian|bangalore|mumbai|delhi/i.test(corpus)) return "India";
  if (/united states|u\.s\.|usa|america/i.test(corpus)) return "United States";
  if (/europe|eu|uk|germany|france/i.test(corpus)) return "Europe";
  return "";
}

function section(title, items) { return { title, items }; }
function sheet(name, rows) { return { name, rows }; }
function para(text, size = 10, bold = false, color = "222222") { return new Paragraph({ children: [new TextRun({ text, size: size * 2, bold, color, font: "Aptos" })], spacing: { after: 130 } }); }
function heading(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 } }); }
function subheading(text) { return new Paragraph({ text, heading: HeadingLevel.HEADING_2, spacing: { before: 180, after: 80 } }); }
function bullet(text) { return new Paragraph({ text, bullet: { level: 0 }, spacing: { after: 80 } }); }
function table(rows) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map((r, i) => new TableRow({ children: r.map((c) => new TableCell({ margins: { top: 80, bottom: 80, left: 100, right: 100 }, shading: i === 0 ? { fill: "EAF0F7" } : undefined, borders: lightBorders(), children: [new Paragraph({ children: [new TextRun({ text: String(c || ""), bold: i === 0, size: 18, font: "Aptos" })] })] })) }))
  });
}
function lightBorders() { return { top: { style: BorderStyle.SINGLE, size: 1, color: "D9DEE7" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "D9DEE7" }, left: { style: BorderStyle.SINGLE, size: 1, color: "D9DEE7" }, right: { style: BorderStyle.SINGLE, size: 1, color: "D9DEE7" } }; }
function card(s, C, x, y, w, h, label, value, accent) {
  s.addShape("roundRect", { x, y, w, h, rectRadius: 0.05, fill: { color: "FFFFFF" }, line: { color: "DCE1EA", width: 0.7 } });
  s.addShape(pptxShapeRect(), { x, y, w: 0.08, h, fill: { color: accent }, line: { color: accent } });
  s.addText(label, { x: x + 0.22, y: y + 0.18, w: w - 0.4, h: 0.18, fontSize: 7.5, color: C.grey, bold: true, margin: 0 });
  s.addText(value, { x: x + 0.22, y: y + 0.5, w: w - 0.35, h: h - 0.62, fontSize: String(value).length > 42 ? 9 : 14, color: C.ink, bold: true, fit: "shrink", margin: 0.02 });
}
function pptxShapeRect() { return "rect"; }
function scoreColor(score, C) { return score >= 70 ? C.green : score >= 55 ? C.gold : C.red; }
function normalizePalette(colors = []) { const p = colors.filter(Boolean); return [p[0] || "17365D", p[1] || "B48A2C", ...p]; }
function splitLines(text) { return String(text || "").split(/\r?\n|,/).map((s) => s.trim()).filter((s) => /^https?:\/\//i.test(s)); }
function stripTags(html) { return decodeHtml(String(html).replace(/<script[\s\S]*?<\/script>/gi, " ").replace(/<style[\s\S]*?<\/style>/gi, " ").replace(/<[^>]+>/g, " ")); }
function decodeHtml(s) { return String(s).replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&#x2F;/g, "/"); }
function normalize(s) { return decodeHtml(String(s || "")).replace(/\s+/g, " ").trim(); }
function clean(s) { return normalize(s).slice(0, 160); }
function slug(s) { return clean(s).toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "") || "target"; }
function titleCase(s) { return s.replace(/([A-Z])/g, " $1").replace(/^./, (m) => m.toUpperCase()); }
function xml(v) { return String(v ?? "").replace(/[<>&'"]/g, (c) => ({ "<": "&lt;", ">": "&gt;", "&": "&amp;", "'": "&apos;", '"': "&quot;" }[c])); }
function escapeRegex(v) { return String(v).replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
function col(n) { let s = ""; for (n += 1; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(((n - 1) % 26) + 65) + s; return s; }
function naturalSort(a, b) { return a.localeCompare(b, undefined, { numeric: true }); }
function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}
async function serveFile(res, file, download = false) {
  if (!path.resolve(file).startsWith(path.resolve(__dirname)) && !path.resolve(file).startsWith(path.resolve(OUTPUT_DIR))) return res.writeHead(403).end("Forbidden");
  const data = await fs.readFile(file);
  const ext = path.extname(file).toLowerCase();
  const types = { ".html": "text/html", ".css": "text/css", ".js": "text/javascript", ".svg": "image/svg+xml", ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
  res.writeHead(200, { "Content-Type": types[ext] || "application/octet-stream", ...(download ? { "Content-Disposition": `attachment; filename="${path.basename(file)}"` } : {}) });
  res.end(data);
}
function json(res, payload, status = 200) {
  res.writeHead(status, { "Content-Type": "application/json" });
  res.end(JSON.stringify(payload));
}
