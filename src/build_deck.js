// Build deliverables/airline_cx_study.pptx from results/summary.json (structured deck, native charts).
const pptxgen = require("pptxgenjs");
const fs = require("fs"), path = require("path");
const SK = "/root/.claude/skills/synced/0a0f9aba-cc6f-4ef9-a0e6-3131eb626fab_b4cd41dc-8108-42f7-b006-6530c8bf94cb/pptx";
const { applyTheme } = require(SK + "/scripts/apply_theme.js");
const ROOT = path.resolve(__dirname, "..");
const S = JSON.parse(fs.readFileSync(path.join(ROOT, "results/summary.json")));
const csv = f => fs.readFileSync(path.join(ROOT, "results", f), "utf8").trim().split("\n").map(l => l.split(","));

const THEME = { name: "Skyline Teal", headFontFace: "Cambria", bodyFontFace: "Calibri",
  colors: { dk1: "1F2A2B", lt1: "FFFFFF", dk2: "0F3D3E", lt2: "E9F2F1", accent1: "0F3D3E", accent2: "F2A541",
            accent3: "5FA8A0", accent4: "C8553D", accent5: "9DB8B5", accent6: "6B7A7B", hlink: "0F6E6F", folHlink: "6B7A7B" } };
const HEX = THEME.colors;
const pres = new pptxgen(); pres.layout = "LAYOUT_16x9"; pres.author = "Aditya Singh";
pres.title = "Airline passenger experience study"; pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;
const pct = x => Math.round(x * 1000) / 10;

pres.defineSlideMaster({ title: "DARK", background: { color: C.text2 }, objects: [],
  placeholders: undefined });
pres.defineSlideMaster({ title: "TITLE_ONLY", background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 0.6, fontSize: 24, bold: true, color: C.text2, fontFace: THEME.headFontFace, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: "Source: Skytrax user reviews (airlinequality.com), 41,396 reviews of 362 airlines, to Aug 2015 · CC0 dataset",
      options: { x: 0.5, y: 5.25, w: 7.6, h: 0.25, fontSize: 9, color: C.accent6, margin: 0 } } },
  ], slideNumber: { x: 9.1, y: 5.22, w: 0.4, h: 0.25, fontSize: 9, color: C.accent6 } });

const chartBase = () => ({ catAxisLabelColor: HEX.accent6, valAxisLabelColor: HEX.accent6, catAxisLabelFontFace: "+mn-lt",
  valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  dataLabelFontSize: 11, dataLabelColor: HEX.dk1, valGridLine: { color: "E3E8E8", size: 0.5 }, catGridLine: { style: "none" },
  showLegend: false, titleFontFace: "+mn-lt" });
const note = (s, t, opts) => s.addText(t, Object.assign({ x: 0.5, y: 4.85, w: 9, h: 0.3, fontSize: 10, italic: true, color: C.accent6, margin: 0, isTextBox: true }, opts || {}));
const card = (s, x, y, w, h, name) => s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: C.background2 }, line: { color: C.background2 }, objectName: name });

// 1 title
pres.addSection({ title: "Opening" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Opening" });
s.addText("Customer experience study", { x: 0.6, y: 1.0, w: 8.8, h: 0.4, fontSize: 14, color: C.accent2, bold: true, margin: 0, isTextBox: true, charSpacing: 2 });
s.addText("What makes a passenger recommend an airline?", { x: 0.6, y: 1.45, w: 8.8, h: 1.3, fontSize: 38, bold: true, color: C.background1, fontFace: THEME.headFontFace, margin: 0, isTextBox: true, valign: "top" });
s.addText(`Key-driver analysis, India benchmark and a sample client deep dive, from ${S.reviews.toLocaleString("en-US")} real passenger reviews of ${S.airlines} airlines`,
  { x: 0.6, y: 2.95, w: 7.8, h: 0.7, fontSize: 16, color: "D7E6E4", margin: 0, isTextBox: true });
s.addText("Aditya Singh · BITS Pilani · November 2026", { x: 0.6, y: 4.7, w: 6, h: 0.3, fontSize: 12, color: "9DB8B5", margin: 0, isTextBox: true });

// 2 exec summary
pres.addSection({ title: "Summary" });
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Summary" });
s.addText("Value and cabin staff decide most recommendations", { placeholder: "title" });
const cards = [
  [`${pct(S.top2_driver_share)}%`, "of explained recommendation comes from value for money and cabin staff", "Out of 5 rated touchpoints; the ranking holds on every split we tested"],
  [`${S.theme_lift_not_vs_recommended["Staff attitude"].toFixed(1)}x`, "more often detractors mention rude or unhelpful staff than promoters do", "Staff attitude is the sharpest complaint signal in review text"],
  [`+${(S.client.whatif_close_all_gaps_pts * 100).toFixed(1)} pts`, "recommend rate for Jet Airways if it matched the benchmark on all 5 ratings", "Cabin staff is its top priority; delays explain much of the rest"],
];
cards.forEach(([big, line, sub], i) => {
  const x = 0.5 + i * 3.08;
  card(s, x, 1.3, 2.84, 3.3, `card${i}`);
  s.addText(big, { x: x + 0.2, y: 1.5, w: 2.45, h: 0.95, fontSize: 44, bold: true, color: C.text2, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
  s.addText(line, { x: x + 0.2, y: 2.5, w: 2.45, h: 1.05, fontSize: 15, color: C.text1, margin: 0, isTextBox: true, valign: "top" });
  s.addText(sub, { x: x + 0.2, y: 3.65, w: 2.45, h: 0.8, fontSize: 11, italic: true, color: C.accent6, margin: 0, isTextBox: true, valign: "top" });
});

// 3 approach
pres.addSection({ title: "Approach" });
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Approach" });
s.addText("Four steps from raw reviews to a client action list", { placeholder: "title" });
const steps = [["1", "Clean", `${S.reviews.toLocaleString("en-US")} reviews, 362 airlines; 0-ratings treated as not rated`],
  ["2", "Driver model", `Shapley split of R² on ${S.complete_case_reviews.toLocaleString("en-US")} complete reviews (R² ${S.model_r2.toFixed(2)})`],
  ["3", "Benchmark", `${S.india_reviews.toLocaleString("en-US")} reviews of 5 Indian carriers vs all airlines, with 95% CIs`],
  ["4", "Text themes", "7 complaint themes searched in review text, spot-checked by reading matches"]];
steps.forEach(([n, h, t], i) => {
  const x = 0.5 + i * 2.32;
  s.addShape(pres.shapes.OVAL, { x: x, y: 1.45, w: 0.62, h: 0.62, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: `step${n}` });
  s.addText(n, { x: x, y: 1.45, w: 0.62, h: 0.62, fontSize: 20, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true });
  if (i < 3) s.addShape(pres.shapes.LINE, { x: x + 0.75, y: 1.76, w: 1.45, h: 0, line: { color: HEX.accent5, width: 1.5, endArrowType: "triangle" } });
  s.addText(h, { x: x, y: 2.25, w: 2.1, h: 0.4, fontSize: 18, bold: true, color: C.text2, margin: 0, isTextBox: true });
  s.addText(t, { x: x, y: 2.7, w: 2.05, h: 1.3, fontSize: 13, color: C.text1, margin: 0, isTextBox: true, valign: "top" });
});
card(s, 0.5, 4.1, 9, 0.62, "check");
s.addText(`Quality check: driver ranking identical on random halves and on 2008–13 vs 2014–15 reviews (Spearman ${S.stability_spearman_random_halves.toFixed(1)})`,
  { x: 0.7, y: 4.1, w: 8.6, h: 0.62, fontSize: 13, color: C.text2, valign: "middle", margin: 0, isTextBox: true });

// 4 drivers
pres.addSection({ title: "Findings" });
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Findings" });
s.addText("Value for money is the biggest driver of recommendation", { placeholder: "title" });
const imp = Object.entries(S.importance).reverse();
s.addChart(pres.charts.BAR, [{ name: "Importance", labels: imp.map(e => e[0]), values: imp.map(e => e[1]) }],
  Object.assign(chartBase(), { x: 0.4, y: 1.15, w: 5.7, h: 3.6, barDir: "bar", chartColors: [HEX.accent1], showValue: true,
    dataLabelPosition: "outEnd", dataLabelFormatCode: "0%", valAxisHidden: true, valGridLine: { style: "none" },
    showTitle: true, title: "Share of explained recommendation (Shapley)", titleFontSize: 12, titleColor: HEX.accent6 }));
card(s, 6.4, 1.25, 3.1, 3.4, "insight");
s.addText([{ text: "What it means", options: { bold: true, fontSize: 15, color: C.text2, breakLine: true } },
  { text: "Passengers judge the trip against what they paid. Value and cabin staff together carry " + pct(S.top2_driver_share) + "% of what the ratings explain.", options: { fontSize: 13, breakLine: true } },
  { text: " ", options: { fontSize: 8, breakLine: true } },
  { text: "Entertainment matters least: fix it last.", options: { fontSize: 13 } }],
  { x: 6.6, y: 1.4, w: 2.75, h: 3.1, color: C.text1, valign: "top", margin: 0, isTextBox: true });

// 5 swing
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Findings" });
s.addText("Rate value poorly and passengers rarely recommend", { placeholder: "title" });
const gap = S.recommend_if_rated_4_5_vs_1_3, gk = Object.keys(gap);
s.addChart(pres.charts.BAR, [{ name: "Rated 4–5", labels: gk, values: gk.map(k => gap[k][0]) }, { name: "Rated 1–3", labels: gk, values: gk.map(k => gap[k][1]) }],
  Object.assign(chartBase(), { x: 0.4, y: 1.1, w: 9.2, h: 3.65, barDir: "col", barGrouping: "clustered", chartColors: [HEX.accent1, HEX.accent2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0%", valAxisMaxVal: 1, valAxisLabelFormatCode: "0%",
    showLegend: true, legendPos: "t", legendFontFace: "+mn-lt", legendFontSize: 11, legendColor: HEX.dk1 }));
note(s, `Recommend rate by rating band, complete reviews. Value: ${pct(gap["Value for money"][0])}% vs ${pct(gap["Value for money"][1])}%, the widest swing. Descriptive, not causal.`);

// 6 IPA
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Findings" });
s.addText("Importance vs performance: protect value and staff", { placeholder: "title" });
const dr = csv("drivers.csv").slice(1).map(r => ({ n: r[0], imp: +r[1], perf: +r[2], q: r[4] }));
s.addChart(pres.charts.SCATTER, [{ name: "Performance", values: dr.map(d => d.perf) }].concat(dr.map(d => ({ name: d.n, values: dr.map(e => e === d ? Math.round(d.imp * 1000) / 10 : null) }))),
  Object.assign(chartBase(), { x: 0.4, y: 1.1, w: 5.9, h: 3.7, lineSize: 0, lineDataSymbolSize: 11, chartColors: [HEX.accent1, HEX.accent4, HEX.accent2, HEX.accent3, HEX.accent6],
    showLegend: true, legendPos: "b", legendFontFace: "+mn-lt", legendFontSize: 10, legendColor: HEX.dk1,
    valAxisLabelFormatCode: "0.0", valAxisMinVal: 0, valAxisMaxVal: 35, catAxisMinVal: 2.7, catAxisMaxVal: 3.5, catAxisLabelFormatCode: "0.0",
    showValAxisTitle: true, valAxisTitle: "Importance (% of explained)", valAxisTitleFontSize: 10, valAxisTitleColor: HEX.accent6,
    showCatAxisTitle: true, catAxisTitle: "Mean score (1–5)", catAxisTitleFontSize: 10, catAxisTitleColor: HEX.accent6 }));
const qs = [["Protect", "Value for money, Cabin staff", "High importance, above-average scores"],
            ["Maintain", "Seat comfort", "Lower importance, above-average score"],
            ["Monitor", "Food & beverages, Entertainment", "Lower importance, below-average scores"]];
qs.forEach(([q, d, t], i) => {
  card(s, 6.55, 1.2 + i * 1.18, 2.95, 1.05, "q" + i);
  s.addText([{ text: q, options: { bold: true, fontSize: 14, color: C.text2, breakLine: true } }, { text: d, options: { fontSize: 12, breakLine: true } }, { text: t, options: { fontSize: 10, italic: true, color: C.accent6 } }],
    { x: 6.7, y: 1.25 + i * 1.18, w: 2.7, h: 0.95, color: C.text1, margin: 0, isTextBox: true, valign: "middle" });
});
note(s, "No driver falls in 'Fix first' across all airlines; for a single client the same matrix uses its own scores (slide 9).", { y: 4.88 });

// 7 themes
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Findings" });
s.addText("Detractors write about staff attitude, fees and delays", { placeholder: "title" });
const lift = Object.entries(S.theme_lift_not_vs_recommended).sort((a, b) => a[1] - b[1]);
s.addChart(pres.charts.BAR, [{ name: "Lift", labels: lift.map(e => e[0]), values: lift.map(e => e[1]) }],
  Object.assign(chartBase(), { x: 0.4, y: 1.1, w: 5.9, h: 3.65, barDir: "bar", chartColors: [HEX.accent4], showValue: true, dataLabelPosition: "outEnd",
    dataLabelFormatCode: '0.0"x"', valAxisHidden: true, valGridLine: { style: "none" }, showTitle: true,
    title: "Mention rate: not-recommended ÷ recommended reviews", titleFontSize: 12, titleColor: HEX.accent6 }));
const d = S.themes["Delays & cancellations"], st = S.themes["Staff attitude"];
card(s, 6.55, 1.2, 2.95, 3.45, "themes");
s.addText([{ text: `${pct(d.not_recommended)}%`, options: { fontSize: 34, bold: true, color: C.text2, fontFace: THEME.headFontFace, breakLine: true } },
  { text: `of not-recommended reviews mention delays or cancellations, vs ${pct(d.recommended)}% of recommended ones`, options: { fontSize: 13, breakLine: true } },
  { text: " ", options: { fontSize: 8, breakLine: true } },
  { text: `Staff attitude: ${pct(st.not_recommended)}% vs ${pct(st.recommended)}%. Food is mentioned more by happy passengers.`, options: { fontSize: 12, italic: true, color: C.accent6 } }],
  { x: 6.72, y: 1.35, w: 2.65, h: 3.2, color: C.text1, valign: "top", margin: 0, isTextBox: true });

// 8 India
pres.addSection({ title: "India" });
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "India" });
s.addText("In India, low-cost IndiGo beats the full-service carriers", { placeholder: "title" });
const ci = Object.entries(S.india_recommend_ci95).sort((a, b) => b[1][1] - a[1][1]);
const lab = ci.map(e => e[0]).concat(["All airlines"]), val = ci.map(e => e[1][1]).concat([S.recommend_rate]);
s.addChart(pres.charts.BAR, [{ name: "Recommend", labels: lab, values: val }],
  Object.assign(chartBase(), { x: 0.4, y: 1.1, w: 5.9, h: 3.6, barDir: "col", chartColors: [HEX.accent1, HEX.accent1, HEX.accent1, HEX.accent1, HEX.accent4, HEX.accent5],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0%", valAxisMaxVal: 1, valAxisLabelFormatCode: "0%" }));
card(s, 6.55, 1.2, 2.95, 3.45, "india");
s.addText([{ text: `${pct(S.india_recommend_low_cost)}% vs ${pct(S.india_recommend_full_service)}%`, options: { fontSize: 26, bold: true, color: C.text2, fontFace: THEME.headFontFace, breakLine: true } },
  { text: `recommend rate, low-cost (${S.india_reviews_low_cost} reviews) vs full-service (${S.india_reviews_full_service.toLocaleString("en-US")})`, options: { fontSize: 13, breakLine: true } },
  { text: " ", options: { fontSize: 8, breakLine: true } },
  { text: `IndiGo ${pct(S.india_recommend_ci95.IndiGo[1])}% ± ${pct(S.india_recommend_ci95.IndiGo[2])} pts; Jet Airways ${pct(S.client.recommend)}% ± ${pct(S.client.recommend_ci95)} pts, below the ${pct(S.recommend_rate)}% benchmark.`, options: { fontSize: 12, breakLine: true } },
  { text: "GoAir (30) and SpiceJet (61): small samples, directional only.", options: { fontSize: 11, italic: true, color: C.accent6 } }],
  { x: 6.72, y: 1.35, w: 2.65, h: 3.2, color: C.text1, valign: "top", margin: 0, isTextBox: true });

// 9 client
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "India" });
s.addText("Sample client, Jet Airways: fix cabin staff first", { placeholder: "title" });
const cl = S.client, names = Object.keys(cl.priority_score);
const hdr = ["Driver", "Gap vs benchmark", "Importance", "Priority", "What-if (pts)"].map(t => ({ text: t, options: { bold: true, color: HEX.lt1, fill: { color: HEX.dk2 }, fontSize: 11 } }));
const rows = names.map((n, i) => [n, cl.gap_vs_benchmark[n].toFixed(2), pct(S.importance[n]) + "%", cl.priority_score[n].toFixed(3), "+" + (cl.whatif_close_gaps_pts[n] * 100).toFixed(1)]
  .map((t, j) => ({ text: String(t), options: { fontSize: 12, bold: i === 0, color: HEX.dk1, fill: { color: i === 0 ? "FCEBCF" : HEX.lt1 }, align: j ? "center" : "left" } })));
s.addTable([hdr].concat(rows), { x: 0.5, y: 1.2, w: 5.9, colW: [1.75, 1.2, 1.1, 0.8, 1.05], border: { type: "solid", pt: 0.5, color: "D5DEDD" }, rowH: 0.42, fontFace: "Calibri", objectName: "clientTable" });
note(s, "Priority = importance × shortfall vs the 362-airline average. What-if from a linear model of recommendation on the 5 ratings: correlational.", { y: 3.85, w: 5.9, h: 0.5 });
card(s, 6.75, 1.2, 2.75, 3.45, "client");
s.addText([{ text: `${pct(cl.recommend)}%`, options: { fontSize: 34, bold: true, color: C.text2, fontFace: THEME.headFontFace, breakLine: true } },
  { text: `recommend rate from ${cl.reviews} reviews, ${Math.abs((cl.recommend - S.recommend_rate) * 100).toFixed(1)} pts below benchmark`, options: { fontSize: 13, breakLine: true } },
  { text: " ", options: { fontSize: 8, breakLine: true } },
  { text: `Closing all 5 rating gaps adds about ${(cl.whatif_close_all_gaps_pts * 100).toFixed(1)} pts. The rest sits outside the ratings: ${pct(cl.delay_mention_in_not_recommended)}% of its detractor reviews mention delays.`, options: { fontSize: 12 } }],
  { x: 6.92, y: 1.35, w: 2.45, h: 3.2, color: C.text1, valign: "top", margin: 0, isTextBox: true });

// 10 recommendations
pres.addSection({ title: "Close" });
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Close" });
s.addText("Recommendations and next steps", { x: 0.6, y: 0.4, w: 8.8, h: 0.7, fontSize: 30, bold: true, color: C.background1, fontFace: THEME.headFontFace, margin: 0, isTextBox: true });
const recs = [["Lead with value, not perks", "Make fare value visible (what's included, no surprise fees): value is the top driver and fees show a 2.4x complaint lift."],
  ["Train for attitude, not just process", "Cabin staff is the #2 driver and Jet's top priority; rude-staff mentions are 5.4x higher among detractors."],
  ["Treat delays as a CX metric", "Delays appear in 41% of detractor reviews: tie delay communication and recovery to the experience scorecard."]];
recs.forEach(([h, t], i) => {
  const y = 1.35 + i * 1.05;
  s.addShape(pres.shapes.OVAL, { x: 0.6, y, w: 0.55, h: 0.55, fill: { color: C.accent2 }, line: { color: C.accent2 }, objectName: "rec" + i });
  s.addText(String(i + 1), { x: 0.6, y, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.text2, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText([{ text: h, options: { bold: true, fontSize: 16, color: HEX.lt1, breakLine: true } }, { text: t, options: { fontSize: 13, color: "D7E6E4" } }],
    { x: 1.35, y: y - 0.08, w: 8.1, h: 0.85, valign: "top", margin: 0, isTextBox: true });
});
s.addText("Next: validate with a primary survey of Indian flyers by route and fare type, and join on-time data to size the delay effect.",
  { x: 0.6, y: 4.6, w: 8.8, h: 0.5, fontSize: 13, italic: true, color: C.accent2, margin: 0, isTextBox: true });

// 11 appendix
s = pres.addSlide({ masterName: "TITLE_ONLY", sectionTitle: "Close" });
s.addText("Appendix: method and limits", { placeholder: "title" });
const lim = [["Method", `Shapley (LMG) split of R² from a linear model of 'recommended' on 5 ratings, ${S.complete_case_reviews.toLocaleString("en-US")} complete reviews; ranks stable on every split. Text themes by keyword search, spot-checked.`],
  ["Sample", "Self-selected online reviewers, not a random survey: likely more extreme opinions than the average passenger."],
  ["Timing", "Reviews up to Aug 2015: Jet Airways stopped flying in 2019, so it serves as a historical sample client."],
  ["Causality", "Importance, gaps and the what-if are associations. A designed survey or a field test is needed to prove cause."]];
lim.forEach(([h, t], i) => {
  const y = 1.2 + i * 0.9;
  card(s, 0.5, y, 9, 0.78, "lim" + i);
  s.addText(h, { x: 0.7, y, w: 1.5, h: 0.78, fontSize: 14, bold: true, color: C.text2, valign: "middle", margin: 0, isTextBox: true });
  s.addText(t, { x: 2.2, y, w: 7.1, h: 0.78, fontSize: 12, color: C.text1, valign: "middle", margin: 0, isTextBox: true });
});

(async () => {
  const out = path.join(ROOT, "deliverables/airline_cx_study.pptx");
  await pres.writeFile({ fileName: out });
  await applyTheme(out, THEME);
  console.log("saved", out);
})();
