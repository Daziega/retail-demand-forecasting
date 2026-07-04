const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaChartLine, FaBoxes, FaExclamationTriangle, FaBrain, FaDatabase,
  FaCogs, FaBalanceScale, FaCheckCircle, FaTimesCircle, FaLightbulb,
  FaClipboardCheck, FaRoad, FaFlask, FaLayerGroup, FaShieldAlt,
  FaChartBar, FaSearchDollar, FaStore, FaUsers, FaChartPie, FaListOl,
  FaBullseye, FaSitemap, FaTools, FaHandshake
} = require("react-icons/fa");

const FIG = "/Users/desmond/Capstone Project/retail-demand-forecasting/figures";

// ---- Palette: Midnight Executive ----
const NAVY = "1E2761";
const NAVY_DARK = "141B4D";
const ICE = "CADCFC";
const ICE_SOFT = "EAF0FE";
const WHITE = "FFFFFF";
const SLATE = "4A5578";
const ACCENT_TEAL = "1C9C8E";
const ACCENT_CORAL = "E8654F";
const ACCENT_AMBER = "E0A82E";
const TEXT_DARK = "1A1F3D";
const TEXT_MUTE = "5B6482";

function makeShadow(opacity = 0.12) {
  return { type: "outer", color: "000000", blur: 8, offset: 3, angle: 45, opacity };
}

async function iconPng(IconComponent, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(IconComponent, { color, size: String(size) })
  );
  const pngBuffer = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + pngBuffer.toString("base64");
}

function iconCircle(slide, icon, x, y, d, bg, fg) {
  slide.addShape("ellipse", { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" } });
  const pad = d * 0.26;
  slide.addImage({ data: icon, x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}

function pageNum(slide, n, total, dark = false) {
  slide.addText(`${n} / ${total}`, {
    x: 9.15, y: 5.32, w: 0.75, h: 0.25, fontSize: 9, color: dark ? ICE : TEXT_MUTE,
    align: "right", fontFace: "Calibri", margin: 0,
  });
}

function kicker(slide, text, color = ACCENT_TEAL) {
  slide.addText(text.toUpperCase(), {
    x: 0.55, y: 0.32, w: 8, h: 0.3, fontSize: 12, bold: true, color,
    fontFace: "Calibri", charSpacing: 2, margin: 0,
  });
}

function title(slide, text, opts = {}) {
  slide.addText(text, {
    x: 0.55, y: 0.6, w: opts.w || 8.9, h: opts.h || 0.7, fontSize: opts.fontSize || 30,
    bold: true, color: opts.color || TEXT_DARK, fontFace: "Cambria", margin: 0,
  });
}

async function main() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Desmond Korbla Aziega, Theresa Korlekuor Apla-Kweku, Abdul-Razak Seidu";
  pres.title = "Demand Forecasting and Inventory Optimisation Using Machine Learning";

  // Pre-render icons once
  const ic = {};
  const specs = [
    ["chart", FaChartLine, WHITE], ["boxes", FaBoxes, WHITE],
    ["warn", FaExclamationTriangle, WHITE], ["brain", FaBrain, WHITE],
    ["db", FaDatabase, WHITE], ["cogs", FaCogs, WHITE],
    ["scale", FaBalanceScale, WHITE], ["check", FaCheckCircle, WHITE],
    ["cross", FaTimesCircle, WHITE], ["bulb", FaLightbulb, WHITE],
    ["clip", FaClipboardCheck, WHITE], ["road", FaRoad, WHITE],
    ["flask", FaFlask, WHITE], ["layers", FaLayerGroup, WHITE],
    ["shield", FaShieldAlt, WHITE], ["bar", FaChartBar, WHITE],
    ["search", FaSearchDollar, WHITE], ["store", FaStore, WHITE],
    ["users", FaUsers, WHITE], ["pie", FaChartPie, WHITE],
    ["list", FaListOl, WHITE], ["target", FaBullseye, WHITE],
    ["sitemap", FaSitemap, WHITE], ["tools", FaTools, WHITE],
    ["handshake", FaHandshake, WHITE],
  ];
  for (const [key, comp, color] of specs) {
    ic[key] = await iconPng(comp, color, 256);
  }
  // Navy-colored variants for use on light backgrounds
  const icNavy = {};
  for (const [key, comp] of specs) {
    icNavy[key] = await iconPng(comp, NAVY, 256);
  }

  const TOTAL = 25;

  // ============================================================
  // SLIDE 1 — TITLE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    // subtle geometric accent - large soft circle, not a stripe
    s.addShape("ellipse", { x: 6.9, y: -1.6, w: 4.6, h: 4.6, fill: { color: NAVY_DARK }, line: { type: "none" } });
    s.addShape("ellipse", { x: 8.6, y: 3.2, w: 2.6, h: 2.6, fill: { color: "263082" }, line: { type: "none" } });

    s.addText("MASTER'S THESIS DEFENSE", {
      x: 0.7, y: 1.05, w: 6, h: 0.35, fontSize: 13, bold: true, color: ACCENT_TEAL,
      fontFace: "Calibri", charSpacing: 3, margin: 0,
    });
    s.addText("Demand Forecasting and\nInventory Optimisation\nUsing Machine Learning", {
      x: 0.7, y: 1.5, w: 8.2, h: 2.1, fontSize: 38, bold: true, color: WHITE,
      fontFace: "Cambria", margin: 0, lineSpacingMultiple: 1.05,
    });
    s.addText("A simulation-based evaluation on the M5-Forecasting (Walmart) dataset, with an SME-accessible Power BI dashboard", {
      x: 0.7, y: 3.75, w: 7.3, h: 0.65, fontSize: 15, italic: true, color: ICE,
      fontFace: "Calibri", margin: 0,
    });

    s.addShape("line", { x: 0.7, y: 4.5, w: 3.4, h: 0, line: { color: "3A4590", width: 1 } });
    s.addText([
      { text: "Desmond Korbla Aziega", options: { bold: true, breakLine: true } },
      { text: "Theresa Korlekuor Apla-Kweku", options: { bold: true, breakLine: true } },
      { text: "Abdul-Razak Seidu", options: { bold: true, breakLine: true } },
      { text: "Supervisor: Prof. Rocío González Martínez", options: { breakLine: true } },
      { text: "Master's Programme in Data Science  |  2026", options: {} },
    ], { x: 0.7, y: 4.62, w: 6.5, h: 1.0, fontSize: 10.5, color: ICE, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.15 });
  }

  // ============================================================
  // SLIDE 2 — AGENDA
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Roadmap");
    title(s, "What This Defense Covers");

    const items = [
      ["store", "The Business Problem", "Why forecasting alone doesn't solve inventory decisions"],
      ["flask", "Methodology", "Five model tiers, OUTL policy, and forward simulation"],
      ["chart", "Results", "Forecast accuracy and the central inventory finding"],
      ["target", "Recommendations", "A diagnostic tool for SME decision-makers"],
    ];
    const colW = 2.15, gap = 0.18, startX = 0.55, y = 1.75, h = 2.85;
    items.forEach((it, i) => {
      const x = startX + i * (colW + gap);
      s.addShape("roundRect", {
        x, y, w: colW, h, rectRadius: 0.08, fill: { color: ICE_SOFT }, line: { type: "none" },
        shadow: makeShadow(0.08),
      });
      iconCircle(s, icNavy[it[0]], x + colW / 2 - 0.32, y + 0.3, 0.64, WHITE, NAVY);
      s.addText(`0${i + 1}`, { x: x + 0.15, y: y + 0.28, w: 0.8, h: 0.4, fontSize: 12, bold: true, color: ICE, fontFace: "Cambria", margin: 0 });
      s.addText(it[1], { x: x + 0.15, y: y + 1.15, w: colW - 0.3, h: 0.6, fontSize: 14, bold: true, color: TEXT_DARK, fontFace: "Cambria", align: "center", margin: 0 });
      s.addText(it[2], { x: x + 0.15, y: y + 1.75, w: colW - 0.3, h: 0.95, fontSize: 10.5, color: TEXT_MUTE, fontFace: "Calibri", align: "center", margin: 0 });
    });
    pageNum(s, 2, TOTAL);
  }

  // ============================================================
  // SLIDE 3 — THE BUSINESS DECISION
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "The Decision-Making Context");
    title(s, "Every Retailer Faces the Same Trade-off");

    s.addText([
      { text: "Too little stock ", options: { bold: true, color: ACCENT_CORAL, breakLine: true } },
      { text: "→ lost sales, eroded customer trust", options: { breakLine: true, color: TEXT_MUTE } },
      { text: " ", options: { breakLine: true, fontSize: 6 } },
      { text: "Too much stock ", options: { bold: true, color: ACCENT_AMBER, breakLine: true } },
      { text: "→ tied-up cash, holding cost, spoilage risk", options: { color: TEXT_MUTE } },
    ], { x: 0.55, y: 1.7, w: 4.5, h: 1.7, fontSize: 15, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3 });

    s.addText(
      "Large retailers solve this with dedicated data-science teams and enterprise software. Small and medium-sized enterprises (SMEs) typically have neither — yet face an identical operational problem.",
      { x: 0.55, y: 3.35, w: 4.5, h: 1.4, fontSize: 13, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3 }
    );

    // Right: big stat callout card
    s.addShape("roundRect", {
      x: 5.5, y: 1.65, w: 4.0, h: 3.35, rectRadius: 0.1, fill: { color: NAVY }, line: { type: "none" },
      shadow: makeShadow(0.15),
    });
    s.addText("Published claims say:", { x: 5.8, y: 1.95, w: 3.4, h: 0.35, fontSize: 12, color: ICE, fontFace: "Calibri", margin: 0 });
    s.addText("10–25%", { x: 5.8, y: 2.3, w: 3.4, h: 0.75, fontSize: 42, bold: true, color: WHITE, fontFace: "Cambria", margin: 0 });
    s.addText("forecast accuracy improvement", { x: 5.8, y: 3.0, w: 3.4, h: 0.4, fontSize: 11.5, color: ICE, fontFace: "Calibri", margin: 0 });
    s.addShape("line", { x: 5.8, y: 3.5, w: 3.4, h: 0, line: { color: "3A4590", width: 1 } });
    s.addText("5–15%", { x: 5.8, y: 3.65, w: 3.4, h: 0.6, fontSize: 34, bold: true, color: WHITE, fontFace: "Cambria", margin: 0 });
    s.addText("inventory cost reduction — from published ML forecasting studies", { x: 5.8, y: 4.25, w: 3.4, h: 0.6, fontSize: 11, color: ICE, fontFace: "Calibri", margin: 0 });
    pageNum(s, 3, TOTAL);
  }

  // ============================================================
  // SLIDE 4 — NATURE OF THE PROBLEM (4 gap cards)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Why This Is Genuinely Hard");
    title(s, "Four Structural Gaps in Current Practice");

    const gaps = [
      ["chart", "Forecast ≠ Decision", "Most studies report accuracy and stop — they never translate forecasts into inventory outcomes.", ACCENT_TEAL],
      ["cross", "No Live Validation", "Inventory performance is usually assumed analytically, not measured by running the policy forward.", ACCENT_CORAL],
      ["store", "SMEs Left Behind", "Enterprise tools assume infrastructure and staffing SMEs simply don't have.", ACCENT_AMBER],
      ["scale", "Uncertainty Ignored", "Textbook safety-stock formulas assume symmetric errors — real retail demand is skewed and intermittent.", NAVY],
    ];
    const colW = 2.15, gap = 0.18, startX = 0.55, y = 1.7;
    gaps.forEach((g, i) => {
      const x = startX + i * (colW + gap);
      s.addShape("roundRect", { x, y, w: colW, h: 3.0, rectRadius: 0.08, fill: { color: ICE_SOFT }, line: { type: "none" }, shadow: makeShadow(0.08) });
      iconCircle(s, ic[g[0]], x + colW / 2 - 0.28, y + 0.28, 0.56, g[3], WHITE);
      s.addText(`GAP ${i + 1}`, { x: x + 0.15, y: y + 0.98, w: colW - 0.3, h: 0.25, fontSize: 10, bold: true, color: g[3], fontFace: "Calibri", align: "center", charSpacing: 1, margin: 0 });
      s.addText(g[1], { x: x + 0.12, y: y + 1.24, w: colW - 0.24, h: 0.55, fontSize: 13, bold: true, color: TEXT_DARK, fontFace: "Cambria", align: "center", margin: 0 });
      s.addText(g[2], { x: x + 0.15, y: y + 1.82, w: colW - 0.3, h: 1.1, fontSize: 10, color: TEXT_MUTE, fontFace: "Calibri", align: "center", margin: 0 });
    });
    pageNum(s, 4, TOTAL);
  }

  // ============================================================
  // SLIDE 5 — RESEARCH OBJECTIVES
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Research Objectives");
    title(s, "Five Objectives Guided This Research");

    const objs = [
      "Comparatively evaluate forecasting approaches — classical, ML, deep learning, ensemble, and probabilistic — on a hierarchical retail benchmark.",
      "Develop a systematic feature-engineering strategy capturing the principal drivers of retail demand.",
      "Integrate the best-performing forecast with the Order-Up-To-Level (OUTL) inventory policy.",
      "Test whether literature-reported gains (10–25% accuracy, 5–15% cost) transfer to SKU-daily granularity.",
      "Deliver a Power BI dashboard translating outputs into an accessible tool for SME decision-makers.",
    ];
    let y = 1.65;
    objs.forEach((o, i) => {
      const rowH = 0.62;
      s.addShape("roundRect", { x: 0.55, y, w: 0.5, h: 0.5, rectRadius: 0.06, fill: { color: NAVY }, line: { type: "none" } });
      s.addText(String(i + 1), { x: 0.55, y, w: 0.5, h: 0.5, fontSize: 18, bold: true, color: WHITE, fontFace: "Cambria", align: "center", valign: "middle", margin: 0 });
      s.addText(o, { x: 1.25, y: y - 0.02, w: 8.2, h: 0.58, fontSize: 12.5, color: TEXT_DARK, fontFace: "Calibri", valign: "middle", margin: 0 });
      y += rowH + 0.14;
    });
    pageNum(s, 5, TOTAL);
  }

  // ============================================================
  // SLIDE 6 — LITERATURE / GAP TABLE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Literature Review — 52 Sources Synthesised");
    title(s, "A Consistent Finding, With One Critical Caveat");

    s.addText([
      { text: "ML and ensemble methods consistently outperform classical statistics ", options: { bold: true, breakLine: true } },
      { text: "across retail forecasting studies — this is well replicated.", options: { color: TEXT_MUTE, breakLine: true } },
    ], { x: 0.55, y: 1.55, w: 8.9, h: 0.65, fontSize: 13, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.2 });

    s.addShape("roundRect", { x: 0.55, y: 2.35, w: 8.9, h: 1.55, rectRadius: 0.08, fill: { color: "FFF4E8" }, line: { type: "none" }, shadow: makeShadow(0.06) });
    iconCircle(s, ic["warn"], 0.8, 2.6, 0.55, ACCENT_AMBER, WHITE);
    s.addText([
      { text: "The caveat that matters most: ", options: { bold: true, color: TEXT_DARK, breakLine: true } },
      { text: "The published 10–25% / 5–15% benchmarks are measured on data aggregated to the weekly or category level — not the daily, individual-product level where real inventory decisions are made.", options: { color: TEXT_DARK } },
    ], { x: 1.55, y: 2.55, w: 7.65, h: 1.2, fontSize: 12.5, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.25 });

    s.addText(
      "This thesis was designed to directly test whether those benchmark figures survive the move to SKU-daily granularity, using real product-store-day data and inventory outcomes measured — not assumed.",
      { x: 0.55, y: 4.15, w: 8.9, h: 0.85, fontSize: 12.5, italic: true, color: SLATE, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.25 }
    );
    pageNum(s, 6, TOTAL);
  }

  // ============================================================
  // SLIDE 7 — DATA (stat callouts)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    kicker(s, "The Data", ICE);
    s.addText("M5-Forecasting (Walmart) Dataset", { x: 0.55, y: 0.6, w: 8.9, h: 0.7, fontSize: 30, bold: true, color: WHITE, fontFace: "Cambria", margin: 0 });

    const stats = [
      ["3,049", "products"], ["10", "stores, 3 US states"], ["1,913", "days (2011–2016)"],
      ["502", "series stratified\nsubsample used"], ["56", "day out-of-sample\ntest window"],
    ];
    const colW = 1.72, gap = 0.12, startX = 0.55, y = 1.85;
    stats.forEach((st, i) => {
      const x = startX + i * (colW + gap);
      s.addShape("roundRect", { x, y, w: colW, h: 2.1, rectRadius: 0.08, fill: { color: NAVY_DARK }, line: { type: "none" } });
      s.addText(st[0], { x: x + 0.08, y: y + 0.35, w: colW - 0.16, h: 0.75, fontSize: 30, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", align: "center", margin: 0 });
      s.addText(st[1], { x: x + 0.1, y: y + 1.15, w: colW - 0.2, h: 0.8, fontSize: 11, color: ICE, fontFace: "Calibri", align: "center", margin: 0 });
    });
    s.addText(
      "Chosen for scale, hierarchical structure mirroring real SME retail operations, and status as an established academic benchmark. A stratified subsample (proportional across category × state) was used for computational tractability.",
      { x: 0.55, y: 4.25, w: 8.9, h: 0.85, fontSize: 12, italic: true, color: ICE, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.25 }
    );
    pageNum(s, 7, TOTAL, true);
  }

  // ============================================================
  // SLIDE 8 — METHODOLOGY: 5 MODEL TIERS (pipeline)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Methodology");
    title(s, "Five Model Tiers, Compared Head-to-Head");

    const tiers = [
      ["list", "Tier 1", "Classical", "MA, SES, AutoARIMA,\nCroston, Croston-SBA,\nSeasonal Naive", SLATE],
      ["cogs", "Tier 2", "Standard ML", "Random Forest,\nXGBoost, LightGBM\n(Tweedie loss)", ACCENT_TEAL],
      ["brain", "Tier 3", "Deep Learning", "LSTM — 128/64\nstacked units,\ndropout 0.2", NAVY],
      ["layers", "Tier 4", "Ensemble", "Stacking: LGBM + XGB\n+ RF → Ridge\nmeta-learner", ACCENT_CORAL],
      ["pie", "Tier 5", "Probabilistic", "LightGBM quantile\nregression\n(P10 – P99)", ACCENT_AMBER],
    ];
    const colW = 1.72, gap = 0.12, startX = 0.55, y = 1.75;
    tiers.forEach((t, i) => {
      const x = startX + i * (colW + gap);
      s.addShape("roundRect", { x, y, w: colW, h: 2.9, rectRadius: 0.08, fill: { color: ICE_SOFT }, line: { type: "none" }, shadow: makeShadow(0.07) });
      iconCircle(s, ic[t[0]], x + colW / 2 - 0.26, y + 0.22, 0.52, t[4], WHITE);
      s.addText(t[1], { x: x + 0.1, y: y + 0.85, w: colW - 0.2, h: 0.28, fontSize: 10.5, bold: true, color: t[4], fontFace: "Calibri", align: "center", charSpacing: 1, margin: 0 });
      s.addText(t[2], { x: x + 0.08, y: y + 1.12, w: colW - 0.16, h: 0.42, fontSize: 13, bold: true, color: TEXT_DARK, fontFace: "Cambria", align: "center", margin: 0 });
      s.addText(t[3], { x: x + 0.1, y: y + 1.62, w: colW - 0.2, h: 1.15, fontSize: 9.5, color: TEXT_MUTE, fontFace: "Calibri", align: "center", margin: 0, lineSpacingMultiple: 1.15 });
    });
    s.addText("Trained strictly on historical data. Evaluated exclusively on a later, entirely unseen test period.", {
      x: 0.55, y: 4.85, w: 8.9, h: 0.4, fontSize: 11.5, italic: true, color: SLATE, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 8, TOTAL);
  }

  // ============================================================
  // SLIDE 9 — OUTL POLICY
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Methodology");
    title(s, "Translating Forecasts Into Inventory Decisions");

    s.addText("The Order-Up-To-Level (R, s, S) Policy", { x: 0.55, y: 1.55, w: 5, h: 0.4, fontSize: 15, bold: true, color: NAVY, fontFace: "Cambria", margin: 0 });
    s.addText(
      "A periodic review policy: every R days, if the inventory position falls at or below reorder point s, an order is placed to raise stock to order-up-to level S.",
      { x: 0.55, y: 2.0, w: 4.5, h: 1.0, fontSize: 12.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3 }
    );
    s.addText("Three policy variants compared — differing only in how safety stock is calculated:", {
      x: 0.55, y: 3.05, w: 4.5, h: 0.5, fontSize: 12, bold: true, color: TEXT_DARK, fontFace: "Calibri", margin: 0,
    });
    s.addText([
      { text: "Classical  ", options: { bold: true, color: SLATE, breakLine: true } },
      { text: "— Gaussian formula fed by classical forecast", options: { color: TEXT_MUTE, breakLine: true } },
      { text: "ML-Gaussian  ", options: { bold: true, color: ACCENT_TEAL, breakLine: true } },
      { text: "— Gaussian formula fed by ML ensemble forecast", options: { color: TEXT_MUTE, breakLine: true } },
      { text: "ML-Empirical-Quantile  ", options: { bold: true, color: ACCENT_CORAL, breakLine: true } },
      { text: "— distribution-aware formula from quantile model", options: { color: TEXT_MUTE } },
    ], { x: 0.55, y: 3.5, w: 4.5, h: 1.4, fontSize: 11, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.2 });

    // Right: formula card
    s.addShape("roundRect", { x: 5.4, y: 1.55, w: 4.1, h: 3.35, rectRadius: 0.1, fill: { color: NAVY }, line: { type: "none" }, shadow: makeShadow(0.12) });
    s.addText("Safety Stock Formulas", { x: 5.7, y: 1.8, w: 3.5, h: 0.35, fontSize: 12, bold: true, color: ICE, fontFace: "Calibri", margin: 0 });
    s.addText("SS = z · σ · √(R + L)", { x: 5.7, y: 2.2, w: 3.5, h: 0.5, fontSize: 17, bold: true, color: WHITE, fontFace: "Cambria", margin: 0 });
    s.addText("Gaussian (textbook) — assumes symmetric forecast-error distribution", { x: 5.7, y: 2.72, w: 3.5, h: 0.6, fontSize: 10.5, italic: true, color: ICE, fontFace: "Calibri", margin: 0 });
    s.addShape("line", { x: 5.7, y: 3.35, w: 3.5, h: 0, line: { color: "3A4590", width: 1 } });
    s.addText("SS = (P95 − P50) · √(R + L)", { x: 5.7, y: 3.5, w: 3.5, h: 0.5, fontSize: 17, bold: true, color: WHITE, fontFace: "Cambria", margin: 0 });
    s.addText("Empirical quantile — reflects the true, skewed shape of retail demand", { x: 5.7, y: 4.05, w: 3.5, h: 0.6, fontSize: 10.5, italic: true, color: ICE, fontFace: "Calibri", margin: 0 });
    pageNum(s, 9, TOTAL);
  }

  // ============================================================
  // SLIDE 10 — THE CRITICAL METHODOLOGICAL CHOICE
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "The Pivotal Methodological Decision", ACCENT_CORAL);
    title(s, "Measuring Performance, Not Assuming It");

    const colW = 4.3, gap = 0.3, y = 1.75, h = 2.9;
    // Left card: assumed
    s.addShape("roundRect", { x: 0.55, y, w: colW, h, rectRadius: 0.1, fill: { color: "FDF0EE" }, line: { type: "none" }, shadow: makeShadow(0.06) });
    iconCircle(s, ic["cross"], 0.55 + colW / 2 - 0.3, y + 0.25, 0.6, ACCENT_CORAL, WHITE);
    s.addText("Most Published Studies", { x: 0.75, y: y + 1.0, w: colW - 0.4, h: 0.35, fontSize: 13, bold: true, color: ACCENT_CORAL, fontFace: "Cambria", align: "center", margin: 0 });
    s.addText(
      "Calculate expected cost and service level analytically — using a formula to project what should happen under an assumed stockout rate.",
      { x: 0.8, y: y + 1.4, w: colW - 0.5, h: 1.3, fontSize: 12, color: TEXT_DARK, fontFace: "Calibri", align: "center", margin: 0, lineSpacingMultiple: 1.3 }
    );

    // Right card: measured
    const x2 = 0.55 + colW + gap;
    s.addShape("roundRect", { x: x2, y, w: colW, h, rectRadius: 0.1, fill: { color: "E9F7F4" }, line: { type: "none" }, shadow: makeShadow(0.06) });
    iconCircle(s, ic["check"], x2 + colW / 2 - 0.3, y + 0.25, 0.6, ACCENT_TEAL, WHITE);
    s.addText("This Thesis", { x: x2 + 0.2, y: y + 1.0, w: colW - 0.4, h: 0.35, fontSize: 13, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", align: "center", margin: 0 });
    s.addText(
      "Runs each policy forward, day by day, against real 56-day test demand for all 502 products — genuinely tracking stock, genuinely recording every stockout.",
      { x: x2 + 0.25, y: y + 1.4, w: colW - 0.5, h: 1.3, fontSize: 12, color: TEXT_DARK, fontFace: "Calibri", align: "center", margin: 0, lineSpacingMultiple: 1.3 }
    );
    pageNum(s, 10, TOTAL);
  }

  // ============================================================
  // SLIDE 11 — ENGINE VALIDATION (falsification)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Validating the Simulation Engine");
    title(s, "Falsification Test: Can It Detect a Stockout?");
    s.addText(
      "Before trusting the simulation, it was deliberately stress-tested under conditions engineered to force stockouts.",
      { x: 0.55, y: 1.35, w: 8.9, h: 0.4, fontSize: 12.5, color: TEXT_MUTE, fontFace: "Calibri", margin: 0 }
    );

    const rows = [
      ["Real buffers (baseline)", "100.0%", "$0", ACCENT_TEAL],
      ["Safety stock set to zero", "74.0%", "$124,338", ACCENT_AMBER],
      ["Demand scaled ×3", "94.1%", "$87,691", ACCENT_AMBER],
      ["Replenishment suppressed", "0.0%", "$469,556", ACCENT_CORAL],
    ];
    let y = 2.0;
    s.addText("Stress Condition", { x: 0.55, y, w: 4.0, h: 0.35, fontSize: 11, bold: true, color: SLATE, fontFace: "Calibri", margin: 0 });
    s.addText("Realised SL", { x: 5.3, y, w: 1.7, h: 0.35, fontSize: 11, bold: true, color: SLATE, fontFace: "Calibri", align: "center", margin: 0 });
    s.addText("Stockout Cost", { x: 7.1, y, w: 2.3, h: 0.35, fontSize: 11, bold: true, color: SLATE, fontFace: "Calibri", align: "center", margin: 0 });
    y += 0.45;
    rows.forEach((r) => {
      s.addShape("roundRect", { x: 0.55, y, w: 8.9, h: 0.55, rectRadius: 0.05, fill: { color: ICE_SOFT }, line: { type: "none" } });
      s.addText(r[0], { x: 0.75, y, w: 4.0, h: 0.55, fontSize: 12, color: TEXT_DARK, fontFace: "Calibri", valign: "middle", margin: 0 });
      s.addText(r[1], { x: 5.3, y, w: 1.7, h: 0.55, fontSize: 13, bold: true, color: r[3], fontFace: "Cambria", align: "center", valign: "middle", margin: 0 });
      s.addText(r[2], { x: 7.1, y, w: 2.3, h: 0.55, fontSize: 13, bold: true, color: r[3], fontFace: "Cambria", align: "center", valign: "middle", margin: 0 });
      y += 0.65;
    });
    s.addText(
      "Every stress condition correctly registered degraded service and rising cost — confirming the 100% baseline result reflects demand genuinely being met, not a counter that never fires.",
      { x: 0.55, y: y + 0.05, w: 8.9, h: 0.5, fontSize: 11, italic: true, color: SLATE, fontFace: "Calibri", margin: 0 }
    );
    pageNum(s, 11, TOTAL);
  }

  // ============================================================
  // SLIDE 12 — RESULTS: FORECAST ACCURACY
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Results — Forecast Accuracy");
    title(s, "The Ensemble Wins, But the Margin Is Modest");

    s.addChart(pres.charts.BAR, [{
      name: "MAE",
      labels: ["Seasonal\nNaive", "MA", "Croston-\nSBA", "AutoARIMA", "LightGBM", "Random\nForest", "XGBoost", "Stacking\nEnsemble"],
      values: [1.145, 1.080, 1.021, 0.997, 0.976, 0.977, 0.970, 0.952],
    }], {
      x: 0.55, y: 1.5, w: 5.7, h: 3.55, barDir: "col",
      chartColors: [SLATE, SLATE, SLATE, SLATE, ACCENT_TEAL, ACCENT_TEAL, ACCENT_TEAL, NAVY],
      chartArea: { fill: { color: WHITE } },
      catAxisLabelColor: TEXT_MUTE, valAxisLabelColor: TEXT_MUTE, catAxisLabelFontSize: 9,
      valGridLine: { color: "E2E8F0", size: 0.5 }, catGridLine: { style: "none" },
      showValue: true, dataLabelPosition: "outEnd", dataLabelColor: TEXT_DARK, dataLabelFontSize: 8,
      dataLabelFormatCode: "0.00",
      valAxisMinVal: 0,
      showLegend: false, showTitle: true, title: "MAE by Model (lower = better)", titleFontSize: 12, titleColor: TEXT_DARK,
    });

    // Right stat callouts
    const statY = 1.6;
    s.addShape("roundRect", { x: 6.5, y: statY, w: 3.0, h: 1.5, rectRadius: 0.08, fill: { color: ICE_SOFT }, line: { type: "none" } });
    s.addText("+4.5%", { x: 6.6, y: statY + 0.15, w: 2.8, h: 0.7, fontSize: 30, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", margin: 0 });
    s.addText("MAE improvement vs best classical baseline", { x: 6.6, y: statY + 0.95, w: 2.8, h: 0.5, fontSize: 10.5, color: TEXT_MUTE, fontFace: "Calibri", margin: 0 });

    s.addShape("roundRect", { x: 6.5, y: statY + 1.7, w: 3.0, h: 1.5, rectRadius: 0.08, fill: { color: ICE_SOFT }, line: { type: "none" } });
    s.addText("+9.5%", { x: 6.6, y: statY + 1.85, w: 2.8, h: 0.7, fontSize: 30, bold: true, color: NAVY, fontFace: "Cambria", margin: 0 });
    s.addText("RMSE improvement — large-error / spike sensitivity", { x: 6.6, y: statY + 2.65, w: 2.8, h: 0.5, fontSize: 10.5, color: TEXT_MUTE, fontFace: "Calibri", margin: 0 });

    s.addText("MAPE improvement was essentially flat — below the 10–25% literature benchmark, as anticipated (Slide 6).", {
      x: 0.55, y: 5.1, w: 8.9, h: 0.35, fontSize: 10.5, italic: true, color: SLATE, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 12, TOTAL);
  }

  // ============================================================
  // SLIDE 13 — SHAP INTERPRETABILITY
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Model Interpretability");
    title(s, "What Actually Drives the Forecast");
    s.addImage({ path: `${FIG}/shap_importance.png`, x: 1.3, y: 1.5, w: 7.4, h: 3.55, sizing: { type: "contain", w: 7.4, h: 3.55 } });
    s.addText("Top 5 features explain 78.6% of model behaviour — dominated by recent sales trend and item-level seasonality, not opaque signals.", {
      x: 0.55, y: 5.12, w: 8.9, h: 0.35, fontSize: 11, italic: true, color: SLATE, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 13, TOTAL);
  }

  // ============================================================
  // SLIDE 14 — THE HEADLINE FINDING (dramatic)
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    s.addShape("ellipse", { x: -1.5, y: 2.8, w: 4, h: 4, fill: { color: NAVY_DARK }, line: { type: "none" } });

    s.addText("THE CENTRAL FINDING", { x: 0.7, y: 0.55, w: 6, h: 0.35, fontSize: 13, bold: true, color: ACCENT_AMBER, fontFace: "Calibri", charSpacing: 3, margin: 0 });
    s.addText("100%", { x: 0.6, y: 0.95, w: 5, h: 1.6, fontSize: 90, bold: true, color: WHITE, fontFace: "Cambria", margin: 0 });
    s.addText("Realised Service Level", { x: 0.7, y: 2.5, w: 5.5, h: 0.5, fontSize: 20, bold: true, color: ICE, fontFace: "Cambria", margin: 0 });
    s.addText(
      "— achieved by ALL THREE policies (Classical, ML-Gaussian, ML-Quantile) across every one of the 502 products, over the full 56-day test window. No policy ever stocked out.",
      { x: 0.7, y: 3.1, w: 5.6, h: 1.3, fontSize: 13.5, color: ICE, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3 }
    );

    s.addShape("roundRect", { x: 6.6, y: 1.1, w: 3.0, h: 3.85, rectRadius: 0.1, fill: { color: "263082" }, line: { type: "none" } });
    s.addText("Consequence:", { x: 6.85, y: 1.35, w: 2.5, h: 0.35, fontSize: 12, bold: true, color: ACCENT_AMBER, fontFace: "Calibri", margin: 0 });
    s.addText(
      "With zero stockouts, there was no stockout cost left for a smarter forecast or a smarter safety-stock formula to save.",
      { x: 6.85, y: 1.72, w: 2.5, h: 1.05, fontSize: 12, color: WHITE, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.25 }
    );
    s.addShape("line", { x: 6.85, y: 2.85, w: 2.5, h: 0, line: { color: "3A4590", width: 1 } });
    s.addText("ML-Quantile policy cost:", { x: 6.85, y: 3.0, w: 2.5, h: 0.32, fontSize: 11, color: ICE, fontFace: "Calibri", margin: 0 });
    s.addText("+1.5%", { x: 6.85, y: 3.28, w: 2.5, h: 0.65, fontSize: 30, bold: true, color: ACCENT_CORAL, fontFace: "Cambria", margin: 0 });
    s.addText("more than the classical baseline — unnecessary safety stock, no offsetting benefit", { x: 6.85, y: 3.95, w: 2.5, h: 0.85, fontSize: 9.5, color: ICE, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.2 });
    pageNum(s, 14, TOTAL, true);
  }

  // ============================================================
  // SLIDE 15 — WHERE THE "21%" WENT
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "An Honest Correction", ACCENT_CORAL);
    title(s, "Why the Earlier 21% Estimate Did Not Survive", { fontSize: 28 });

    const colW = 4.3, gap = 0.3, y = 1.7, h = 3.05;
    s.addShape("roundRect", { x: 0.55, y, w: colW, h, rectRadius: 0.1, fill: { color: "FDF0EE" }, line: { type: "none" } });
    s.addText("Earlier (Analytical) Draft", { x: 0.75, y: y + 0.25, w: colW - 0.4, h: 0.35, fontSize: 13, bold: true, color: ACCENT_CORAL, fontFace: "Cambria", margin: 0 });
    s.addText("21%", { x: 0.75, y: y + 0.65, w: colW - 0.4, h: 0.9, fontSize: 48, bold: true, color: ACCENT_CORAL, fontFace: "Cambria", margin: 0 });
    s.addText(
      "Assumed the Gaussian policy achieved ~83% service and the quantile policy ~95% \"by calibration,\" then computed cost from that gap.",
      { x: 0.8, y: y + 1.65, w: colW - 0.5, h: 1.2, fontSize: 11.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3 }
    );
    s.addText("The quantile policy won by construction — before any simulation ran.", {
      x: 0.8, y: y + 2.6, w: colW - 0.5, h: 0.4, fontSize: 10.5, italic: true, color: ACCENT_CORAL, fontFace: "Calibri", margin: 0,
    });

    const x2 = 0.55 + colW + gap;
    s.addShape("roundRect", { x: x2, y, w: colW, h, rectRadius: 0.1, fill: { color: "E9F7F4" }, line: { type: "none" } });
    s.addText("This Thesis (Simulated)", { x: x2 + 0.2, y: y + 0.25, w: colW - 0.4, h: 0.35, fontSize: 13, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", margin: 0 });
    s.addText("−1.5%", { x: x2 + 0.2, y: y + 0.65, w: colW - 0.4, h: 0.9, fontSize: 48, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", margin: 0 });
    s.addText(
      "Both policies actually achieve 100% service — the quantile buffer protects against stockouts that never occur.",
      { x: x2 + 0.25, y: y + 1.65, w: colW - 0.5, h: 1.2, fontSize: 11.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3 }
    );
    s.addText("The Ch1 5–15% inventory-cost target is not achieved on this dataset — treated as a finding, not a failure.", {
      x: x2 + 0.25, y: y + 2.55, w: colW - 0.5, h: 0.45, fontSize: 10.5, italic: true, color: ACCENT_TEAL, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 15, TOTAL);
  }

  // ============================================================
  // SLIDE 16 — COST DECOMPOSITION
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Results — Cost Decomposition");
    title(s, "Holding Cost Dominates; Stockout Cost Is Zero");

    s.addChart(pres.charts.BAR, [
      { name: "Holding Cost", labels: ["Classical", "ML-Gaussian", "ML-Quantile"], values: [63256, 63243, 64241] },
      { name: "Ordering Cost", labels: ["Classical", "ML-Gaussian", "ML-Quantile"], values: [332, 332, 332] },
      { name: "Stockout Cost", labels: ["Classical", "ML-Gaussian", "ML-Quantile"], values: [0, 0, 0] },
    ], {
      x: 0.55, y: 1.55, w: 8.9, h: 3.55, barDir: "bar", barGrouping: "stacked",
      chartColors: [NAVY, ACCENT_TEAL, ACCENT_CORAL],
      chartArea: { fill: { color: WHITE } },
      catAxisLabelColor: TEXT_DARK, valAxisLabelColor: TEXT_MUTE,
      valGridLine: { color: "E2E8F0", size: 0.5 }, catGridLine: { style: "none" },
      valAxisMinVal: 0,
      showLegend: true, legendPos: "b", legendColor: TEXT_MUTE, legendFontSize: 11,
      showTitle: true, title: "Annual Cost Composition (L = 14 days, SL = 95%)", titleFontSize: 12, titleColor: TEXT_DARK,
    });
    s.addText("Bars start at $0 — the ML-Quantile difference is a genuine but small ~1.5% increase, not the large gap a truncated axis would suggest.", {
      x: 0.55, y: 5.15, w: 8.9, h: 0.32, fontSize: 10, italic: true, color: SLATE, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 16, TOTAL);
  }

  // ============================================================
  // SLIDE 17 — SENSITIVITY / ROBUSTNESS
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Sensitivity Analysis");
    title(s, "The Finding Is Robust — Not a One-Off Result");

    const cards = [
      ["scale", "Across Lead Time (7–21 days)", "−1.3% to −1.7%", "Quantile policy consistently more expensive, never favourable"],
      ["search", "Across Stockout Multiplier (0.4×–2.0×)", "No change", "Multiplier never engages — zero stockouts to price at any level"],
    ];
    const colW = 4.3, gap = 0.3, startX = 0.55, y = 1.75;
    cards.forEach((c, i) => {
      const x = startX + i * (colW + gap);
      s.addShape("roundRect", { x, y, w: colW, h: 2.85, rectRadius: 0.09, fill: { color: ICE_SOFT }, line: { type: "none" }, shadow: makeShadow(0.07) });
      iconCircle(s, icNavy[c[0]], x + colW / 2 - 0.3, y + 0.3, 0.6, WHITE, NAVY);
      s.addText(c[1], { x: x + 0.2, y: y + 1.05, w: colW - 0.4, h: 0.55, fontSize: 13.5, bold: true, color: TEXT_DARK, fontFace: "Cambria", align: "center", margin: 0 });
      s.addText(c[2], { x: x + 0.2, y: y + 1.62, w: colW - 0.4, h: 0.5, fontSize: 17, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", align: "center", margin: 0 });
      s.addText(c[3], { x: x + 0.2, y: y + 2.18, w: colW - 0.4, h: 0.6, fontSize: 11, color: TEXT_MUTE, fontFace: "Calibri", align: "center", margin: 0, lineSpacingMultiple: 1.15 });
    });
    s.addText("A threefold demand-shock stress test (Slide 11) shows even that is not enough, on this dataset, to make the quantile policy clearly superior.", {
      x: 0.55, y: 4.65, w: 8.9, h: 0.5, fontSize: 11, italic: true, color: SLATE, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 17, TOTAL);
  }

  // ============================================================
  // SLIDE 18 — THE BOUNDARY CONDITION
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "The Real Contribution");
    title(s, "A Boundary Condition, Not a Blanket Claim");

    s.addText(
      "The value of ML forecasting and distributional safety stock is conditional on the demand regime:",
      { x: 0.55, y: 1.5, w: 8.9, h: 0.45, fontSize: 13.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0 }
    );

    const y = 2.15, colW = 4.3, gap = 0.3, h = 2.6;
    s.addShape("roundRect", { x: 0.55, y, w: colW, h, rectRadius: 0.1, fill: { color: "E9F7F4" }, line: { type: "none" } });
    iconCircle(s, ic["check"], 0.8, y + 0.25, 0.55, ACCENT_TEAL, WHITE);
    s.addText("Genuine stockout risk exists", { x: 1.55, y: y + 0.28, w: colW - 1.0, h: 0.55, fontSize: 13, bold: true, color: TEXT_DARK, fontFace: "Cambria", margin: 0 });
    s.addText("→ the larger, empirically calibrated buffer earns back its holding cost by preventing real stockouts.", {
      x: 0.8, y: y + 1.0, w: colW - 0.5, h: 1.4, fontSize: 12, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3,
    });

    const x2 = 0.55 + colW + gap;
    s.addShape("roundRect", { x: x2, y, w: colW, h, rectRadius: 0.1, fill: { color: "FDF0EE" }, line: { type: "none" } });
    iconCircle(s, ic["cross"], x2 + 0.25, y + 0.25, 0.55, ACCENT_CORAL, WHITE);
    s.addText("Service is already saturated", { x: x2 + 1.0, y: y + 0.28, w: colW - 1.0, h: 0.55, fontSize: 13, bold: true, color: TEXT_DARK, fontFace: "Cambria", margin: 0 });
    s.addText("→ the buffer protects against stockouts that don't occur — pure added cost, no benefit. (The M5 test regime.)", {
      x: x2 + 0.25, y: y + 1.0, w: colW - 0.5, h: 1.4, fontSize: 12, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.3,
    });
    pageNum(s, 18, TOTAL);
  }

  // ============================================================
  // SLIDE 19 — DASHBOARD SHOWCASE 1
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "The Practical Deliverable");
    title(s, "An Accessible Power BI Dashboard for SMEs");
    // KPI strip (native slide styling — avoids embedding Power BI's own toolbar chrome)
    const kpis = [["0.952", "MAE Ensemble"], ["100.0%", "SL Quantile"], ["−1.5%", "Cost Chg. Quantile"]];
    const kpiW = 2.7, kpiGap = 0.15, kpiStartX = 0.75;
    kpis.forEach((k, i) => {
      const kx = kpiStartX + i * (kpiW + kpiGap);
      s.addShape("roundRect", { x: kx, y: 1.45, w: kpiW, h: 0.85, rectRadius: 0.06, fill: { color: ICE_SOFT }, line: { type: "none" } });
      s.addText(k[0], { x: kx + 0.15, y: 1.5, w: kpiW - 0.3, h: 0.45, fontSize: 20, bold: true, color: NAVY, fontFace: "Cambria", margin: 0 });
      s.addText(k[1], { x: kx + 0.15, y: 1.95, w: kpiW - 0.3, h: 0.3, fontSize: 9.5, color: TEXT_MUTE, fontFace: "Calibri", margin: 0 });
    });
    s.addImage({ path: `${FIG}/dashboard_forecast_overview_sku_chartonly.png`, x: 0.9, y: 2.45, w: 8.2, h: 2.55, sizing: { type: "contain", w: 8.2, h: 2.55 } });
    s.addText("Forecast Overview — actual vs. ML-ensemble vs. classical forecast, with KPI tiles and per-SKU drilldown", {
      x: 0.55, y: 5.12, w: 8.9, h: 0.35, fontSize: 11, italic: true, color: SLATE, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 19, TOTAL);
  }

  // ============================================================
  // SLIDE 20 — DASHBOARD SHOWCASE 2
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "The Practical Deliverable");
    title(s, "Cost Analysis — The Finding, Made Visible");
    s.addImage({ path: `${FIG}/dashboard_cost_analysis.png`, x: 0.75, y: 1.5, w: 8.5, h: 3.55, sizing: { type: "contain", w: 8.5, h: 3.55 } });
    s.addText("Free (Power BI free tier). Runs on a standard laptop. Requires only data a typical POS system already produces.", {
      x: 0.55, y: 5.12, w: 8.9, h: 0.35, fontSize: 11, italic: true, color: SLATE, fontFace: "Calibri", margin: 0,
    });
    pageNum(s, 20, TOTAL);
  }

  // ============================================================
  // SLIDE 21 — CONTRIBUTIONS
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Contributions");
    title(s, "Theoretical and Practical Contributions");

    const theo = [
      "First simulation-based measurement of the Gaussian-vs-empirical safety-stock choice on M5",
      "Demonstrates ML forecasting's inventory value is regime-conditional, not automatic",
      "Shows the inventory comparison is robust to the stockout-cost multiplier (zero stockouts at 0.4×–2.0×)",
    ];
    const prac = [
      "End-to-end, reproducible, free-and-open-source pipeline (< 3 hrs on a laptop)",
      "A forward-simulation engine usable as a pre-investment diagnostic",
      "An interactive Power BI dashboard for non-technical SME managers",
    ];
    const colW = 4.3, gap = 0.3, y = 1.65;
    s.addShape("roundRect", { x: 0.55, y, w: colW, h: 3.15, rectRadius: 0.1, fill: { color: ICE_SOFT }, line: { type: "none" }, shadow: makeShadow(0.06) });
    s.addText("Theoretical", { x: 0.8, y: y + 0.22, w: colW - 0.5, h: 0.4, fontSize: 14, bold: true, color: NAVY, fontFace: "Cambria", margin: 0 });
    theo.forEach((t, i) => {
      s.addText([{ text: t, options: { bullet: { code: "2022" } } }], {
        x: 0.8, y: y + 0.75 + i * 0.78, w: colW - 0.55, h: 0.75, fontSize: 10.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.15,
      });
    });

    const x2 = 0.55 + colW + gap;
    s.addShape("roundRect", { x: x2, y, w: colW, h: 3.15, rectRadius: 0.1, fill: { color: ICE_SOFT }, line: { type: "none" }, shadow: makeShadow(0.06) });
    s.addText("Practical", { x: x2 + 0.25, y: y + 0.22, w: colW - 0.5, h: 0.4, fontSize: 14, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", margin: 0 });
    prac.forEach((t, i) => {
      s.addText([{ text: t, options: { bullet: { code: "2022" } } }], {
        x: x2 + 0.25, y: y + 0.75 + i * 0.78, w: colW - 0.55, h: 0.75, fontSize: 10.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.15,
      });
    });
    pageNum(s, 21, TOTAL);
  }

  // ============================================================
  // SLIDE 22 — RECOMMENDATIONS FOR SMES
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Recommendations");
    title(s, "For SMEs: Diagnose Before You Invest");

    const steps = [
      ["Don't trust published % claims alone", "Ask: at what data granularity, and was cost measured or assumed?"],
      ["Run the diagnostic on your own data first", "The delivered simulation engine tells you if you face real stockout risk"],
      ["Invest where the diagnostic says risk is real", "High-margin, fast-moving, or highly seasonal categories are the best candidates"],
      ["Deploy the dashboard regardless", "Free, low-risk visibility — independent of which forecasting method feeds it"],
    ];
    let y = 1.65;
    steps.forEach((st, i) => {
      s.addShape("roundRect", { x: 0.55, y, w: 0.55, h: 0.55, rectRadius: 0.28, fill: { color: NAVY }, line: { type: "none" } });
      s.addText(String(i + 1), { x: 0.55, y, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: WHITE, fontFace: "Cambria", align: "center", valign: "middle", margin: 0 });
      s.addText(st[0], { x: 1.3, y: y - 0.02, w: 8.1, h: 0.4, fontSize: 13.5, bold: true, color: TEXT_DARK, fontFace: "Cambria", margin: 0 });
      s.addText(st[1], { x: 1.3, y: y + 0.35, w: 8.1, h: 0.4, fontSize: 11, color: TEXT_MUTE, fontFace: "Calibri", margin: 0 });
      y += 0.85;
    });
    pageNum(s, 22, TOTAL);
  }

  // ============================================================
  // SLIDE 23 — LIMITATIONS & FUTURE WORK
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: WHITE };
    kicker(s, "Limitations & Future Work");
    title(s, "What This Study Does Not Claim");

    const lims = [
      "Single 56-day test window — a longer or differently-timed window could expose real stockouts",
      "LSTM evaluated at one architecture only — more elaborate sequence models untested",
      "Stockout cost modelled as linear — real-world costs (e.g. churn) may be non-linear",
    ];
    const future = [
      "Re-test on retail benchmarks with more volatile demand (Rossmann, Favorita)",
      "Validate the boundary condition in a live SME deployment with genuine stockout exposure",
      "Extend the simulation engine to non-linear stockout costs and lead-time uncertainty",
    ];
    const colW = 4.3, gap = 0.3, y = 1.65;
    s.addShape("roundRect", { x: 0.55, y, w: colW, h: 3.15, rectRadius: 0.1, fill: { color: "FDF0EE" }, line: { type: "none" } });
    s.addText("Limitations", { x: 0.8, y: y + 0.22, w: colW - 0.5, h: 0.4, fontSize: 14, bold: true, color: ACCENT_CORAL, fontFace: "Cambria", margin: 0 });
    lims.forEach((t, i) => {
      s.addText([{ text: t, options: { bullet: { code: "2022" } } }], {
        x: 0.8, y: y + 0.75 + i * 0.85, w: colW - 0.55, h: 0.8, fontSize: 10.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.2,
      });
    });
    const x2 = 0.55 + colW + gap;
    s.addShape("roundRect", { x: x2, y, w: colW, h: 3.15, rectRadius: 0.1, fill: { color: "E9F7F4" }, line: { type: "none" } });
    s.addText("Future Work", { x: x2 + 0.25, y: y + 0.22, w: colW - 0.5, h: 0.4, fontSize: 14, bold: true, color: ACCENT_TEAL, fontFace: "Cambria", margin: 0 });
    future.forEach((t, i) => {
      s.addText([{ text: t, options: { bullet: { code: "2022" } } }], {
        x: x2 + 0.25, y: y + 0.75 + i * 0.85, w: colW - 0.55, h: 0.8, fontSize: 10.5, color: TEXT_DARK, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.2,
      });
    });
    pageNum(s, 23, TOTAL);
  }

  // ============================================================
  // SLIDE 24 — CONCLUSION
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    s.addShape("ellipse", { x: 7.3, y: -1.8, w: 4.2, h: 4.2, fill: { color: NAVY_DARK }, line: { type: "none" } });
    kicker(s, "Conclusion", ACCENT_TEAL);
    s.addText("The Value of ML Forecasting Is Real —\nBut Conditional, Not Automatic", {
      x: 0.6, y: 1.1, w: 8.6, h: 1.5, fontSize: 30, bold: true, color: WHITE, fontFace: "Cambria", margin: 0, lineSpacingMultiple: 1.1,
    });
    s.addText(
      "This thesis delivers both the evidence for that conclusion and a practical, free, open-source tool that lets any organisation test which side of that condition it falls on — using its own data, before spending anything further.",
      { x: 0.6, y: 2.75, w: 8.2, h: 1.1, fontSize: 15, color: ICE, fontFace: "Calibri", margin: 0, lineSpacingMultiple: 1.35 }
    );
    s.addText(
      "The barrier this research removes is not financial. It is informational.",
      { x: 0.6, y: 4.05, w: 8.2, h: 0.6, fontSize: 16, italic: true, bold: true, color: ACCENT_AMBER, fontFace: "Cambria", margin: 0 }
    );
    pageNum(s, 24, TOTAL, true);
  }

  // ============================================================
  // SLIDE 25 — THANK YOU / Q&A
  // ============================================================
  {
    const s = pres.addSlide();
    s.background = { color: NAVY };
    s.addShape("ellipse", { x: -1.8, y: -1.8, w: 4.5, h: 4.5, fill: { color: NAVY_DARK }, line: { type: "none" } });
    s.addShape("ellipse", { x: 8.4, y: 3.4, w: 3.0, h: 3.0, fill: { color: "263082" }, line: { type: "none" } });
    s.addText("Thank You", { x: 0, y: 2.0, w: 10, h: 1.1, fontSize: 48, bold: true, color: WHITE, fontFace: "Cambria", align: "center", margin: 0 });
    s.addText("Questions & Discussion", { x: 0, y: 3.05, w: 10, h: 0.6, fontSize: 18, color: ICE, fontFace: "Calibri", align: "center", margin: 0 });
    s.addText("Desmond Korbla Aziega  •  Theresa Korlekuor Apla-Kweku  •  Abdul-Razak Seidu", {
      x: 0, y: 4.85, w: 10, h: 0.35, fontSize: 11, color: ICE, fontFace: "Calibri", align: "center", margin: 0,
    });
    s.addText("Supervisor: Prof. Rocío González Martínez", {
      x: 0, y: 5.15, w: 10, h: 0.3, fontSize: 10, italic: true, color: ICE, fontFace: "Calibri", align: "center", margin: 0,
    });
  }

  // Speaker notes — full talking-point script, one entry per slide (1-indexed)
  const notes = {
    1: "Welcome the committee, state your name and the thesis title, and give a one-sentence " +
       "framing: this is a simulation-based test of whether published ML-forecasting inventory " +
       "gains actually materialise at the SKU-daily level an SME would use. Keep this under 30 " +
       "seconds — the agenda slide does the roadmap work.",
    2: "Walk through the four sections in order: business problem, methodology, results, " +
       "recommendations. Signal explicitly that the results section contains a correction to an " +
       "earlier estimate — this primes the committee rather than surprising them cold on slide 14.",
    3: "Anchor the stakes in plain trade-off language before any technical content. The right-hand " +
       "card's 10–25% / 5–15% figures are the literature benchmarks this thesis exists to test — " +
       "say so explicitly, since slide 15 revisits these exact numbers.",
    4: "Each gap maps to a research objective on the next slide — Gap 1→Obj 1, Gap 2→Obj 3/4, " +
       "Gap 3→Obj 5, Gap 4→Obj 3. If asked to justify the gaps: they come from the 52-source " +
       "literature synthesis (slide 6), not personal opinion.",
    5: "Read these briskly — the committee has already seen them in the written thesis. Flag " +
       "Objective 4 explicitly: it was reframed, with supervisor input, from 'achieve X% gains' to " +
       "'test whether gains transfer' — this is the objective that licenses the disconfirmation " +
       "finding as a valid outcome, not a failure.",
    6: "This is the thesis's central research question, stated plainly: do the 10–25% / 5–15% " +
       "benchmarks survive the move from weekly/category aggregation to daily/SKU granularity? " +
       "Everything through slide 15 builds toward answering it — just plant the question here, " +
       "don't resolve it yet.",
    7: "Standard dataset justification. If asked why only 502 of 3,049 series: computational " +
       "tractability, with the subsample stratified proportionally across category × state to " +
       "preserve representativeness — not cherry-picked. If challenged on using Walmart data to " +
       "make SME claims: the transferability argument is about the pipeline and data requirements, " +
       "not the specific retailer — §1.2.3 specifies this concretely: three data sources (daily unit " +
       "sales, calendar/promotional data, weekly pricing) that any standard POS or ERP system " +
       "already produces, open-source tooling only, and one analyst with intermediate Python " +
       "competency. M5 is used as a large-scale, hierarchically realistic proxy, not a claim that " +
       "Walmart and a corner retailer face identical demand patterns.",
    8: "Move quickly — five tiers, all standard architectures, nothing novel claimed here. The " +
       "methodological contribution is downstream in the simulation, not in model selection. If " +
       "pressed on the LSTM specifically: 128/64 stacked units with dropout 0.2 was tuned via " +
       "validation-set MAPE, detailed in the appendix.",
    9: "Explain the (R,s,S) mechanics briefly, then land on the two safety-stock formulas — that's " +
       "the crux. Gaussian assumes symmetric forecast errors; empirical-quantile does not. This " +
       "distinction is exactly what the inventory experiment tests — say that explicitly before " +
       "moving on.",
    10: "The most important methodological slide in the deck — slow down here. State plainly: an " +
        "earlier draft used the left-hand (assumed/analytical) approach, produced a 21% cost-saving " +
        "figure, and a supervisor review caught that this was circular — the quantile policy was " +
        "assigned a higher service level by construction, then credited for 'achieving' it. The " +
        "right-hand approach is what actually got measured and reported. Naming this transparently " +
        "before showing the result builds credibility rather than undermining it.",
    11: "Before trusting a simulation that reports zero stockouts, you have to prove it CAN detect " +
        "a stockout — that's this table. All three stress conditions correctly degrade service and " +
        "spike cost, ruling out 'your simulator is just broken / always reports success.' This is " +
        "shown before the headline number precisely to pre-empt that objection.",
    12: "The ensemble modestly beats classical baselines on MAE (+4.5%) and RMSE (+9.5%), but MAPE " +
        "improvement is flat — well short of the 10–25% literature benchmark. Flag this now: better " +
        "forecasts existed, but as slide 14 shows, better forecasts didn't translate into better " +
        "inventory outcomes in this regime. That's the whole thesis in one sentence.",
    13: "Quick interpretability check — recent rolling means and item-level seasonality dominate, " +
        "not some opaque signal. Useful mainly to pre-empt a 'is this a black box' question; don't " +
        "dwell unless asked. If asked why SHAP is run on LightGBM rather than the winning stacking " +
        "ensemble: stacking ensembles combine three base learners through a meta-learner, which " +
        "resists clean, single-model attribution — LightGBM is used as the standard, well-supported " +
        "proxy for interpretability, consistent with how tree-based SHAP is normally applied in this " +
        "literature. The ensemble's own behaviour is dominated by the same base learners LightGBM " +
        "represents.",
    14: "This is the headline slide. Pause here. Emphasise: this was measured, not assumed — the " +
        "simulation was independently validated (falsification results, slide 11). If asked why " +
        "this contradicts the abstract's framing: it doesn't — it IS the framing. The abstract " +
        "explicitly states the apparent saving does not survive measurement.",
    15: "If asked 'so was the thesis a failure?' — No. Disconfirmation of a specific hypothesis " +
        "under specific conditions is a valid and valuable research outcome, especially when it " +
        "corrects a methodological flaw (assumed vs measured performance) common across the " +
        "literature this thesis reviewed. This slide names the correction directly rather than " +
        "burying it.",
    16: "This chart is the mechanical explanation for slides 14–15: holding cost is ~99.5% of total " +
        "cost and nearly identical across all three policies; ordering cost is a rounding error; " +
        "stockout cost is exactly zero for all three. There is structurally no room for a " +
        "safety-stock formula to differentiate cost when the thing it protects against never " +
        "happens.",
    17: "Anticipate the 'is this just a fluke of your parameter choices' question — walk through " +
        "both sensitivity axes actually reported in Chapter 5 §5.7 (lead time and stockout " +
        "multiplier) and note neither flips the conclusion. The stockout-multiplier result is the " +
        "strongest point: even inflating stockout cost up to 2× doesn't matter, because the " +
        "multiplier never engages against a real stockout. If asked about forecast-accuracy (MAPE) " +
        "sensitivity specifically: that sweep belonged to the earlier, retracted analytical " +
        "methodology and was not reproduced under simulation — say so directly rather than implying " +
        "it was tested. If pushed on why one 56-day window is enough to generalise a boundary " +
        "condition: it isn't, on its own — that's exactly why the deliverable is a diagnostic tool " +
        "rather than a universal percentage. The claim isn't 'this window proves ML forecasting " +
        "never helps'; it's 'here is a validated way to test your own window, your own data, before " +
        "investing.' Window-dependence is a limitation of this specific measurement, not of the " +
        "method being offered.",
    18: "This is the actual contribution: not a cost-saving percentage, but a boundary condition " +
        "and a reusable diagnostic. Committee may ask for the mechanism — explain via the two-card " +
        "comparison on this slide if asked to elaborate.",
    19: "Pivot tone here — from 'here's what we found' to 'here's what we built regardless.' The " +
        "dashboard has value independent of which policy wins, because visibility into " +
        "forecast-vs-actual is useful to an SME manager either way. Mention it runs on Power BI's " +
        "free tier.",
    20: "This page makes the finding SME-legible without requiring anyone to read the thesis — the " +
        "cost breakdown by policy, visually, is slide 16's chart made operational. Reiterate: free, " +
        "laptop-grade, uses data a typical POS system already produces.",
    21: "Two distinct audiences, two columns — don't let the committee conflate them. The " +
        "theoretical contribution is the finding itself (regime-conditionality); the practical " +
        "contribution is the reusable diagnostic tool. Both stand on their own even though the " +
        "headline number didn't survive.",
    22: "This is the advice an SME owner would actually walk away with. Step 2 is the operational " +
        "core: run the diagnostic on your OWN data before spending on either better forecasting or " +
        "better safety-stock formulas, because the answer is regime-dependent, not universal. If " +
        "asked how technical a team needs to be to actually run this: §1.2.3 is specific — one " +
        "analyst with roughly 12-18 months of Python scripting experience, using only free and " +
        "open-source libraries, on a standard laptop. No dedicated data science hire or paid " +
        "infrastructure is required, which is the whole point of the SME-accessibility claim.",
    23: "Be candid and unhurried here — a 56-day single test window is the most defensible line of " +
        "questioning to expect, and 'a longer or differently-timed window could expose real " +
        "stockouts' is the honest answer, not a hedge. Each limitation is paired row-by-row with " +
        "its future-work fix.",
    24: "Land on the reframe, not the retracted number: ML forecasting's inventory value is real " +
        "but conditional, and the contribution is a way to test which condition you're in before " +
        "investing. The closing line — 'the barrier is informational, not financial' — is the " +
        "sentence to leave the committee with.",
    25: "Stop talking. Invite questions. If the committee opens with a challenge to the 21%→−1.5% " +
        "reversal, don't get defensive — point back to slide 10 (the methodological pivot) and " +
        "slide 11 (falsification validation) as the two slides that pre-empt it.",
  };
  Object.entries(notes).forEach(([slideNum, text]) => {
    pres.slides[Number(slideNum) - 1].addNotes(text);
  });

  await pres.writeFile({ fileName: "/Users/desmond/Capstone Project/retail-demand-forecasting/defense_slides/Thesis_Defense.pptx" });
  console.log("Saved: Thesis_Defense.pptx");
}

main().catch((e) => { console.error(e); process.exit(1); });
