const pptxgen = require("pptxgenjs");
const fs = require("fs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "朱业拓 Toky";
pres.company = "HICC Pet";
pres.title = "转正述职报告";

// ---- HICC brand palette (from HICC PPT Template 2023) ----
const PURPLE = "6A51B7";
const PURPLE_DK = "452D8D";
const PURPLE_MD = "846CCF";
const PURPLE_LT = "D5D0F7";
const PURPLE_XLT = "F3F1FF";
const ORANGE = "FCC279";
const ORANGE_DK = "F7BA59";
const YELLOW = "FFD868";
const WHITE = "FFFFFF";
const INK = "2B2340";
const MUTED = "6E6883";
const F = "Arial";

const LOGO = "assets/logo.png";
const W = 13.33, H = 7.5;

// ---------- helpers ----------
function darkBg(slide) {
  slide.background = { color: PURPLE_DK };
  slide.addShape(pres.ShapeType.ellipse, {
    x: 10.6, y: -1.5, w: 4.6, h: 4.6, fill: { color: PURPLE, transparency: 55 },
  });
  slide.addShape(pres.ShapeType.ellipse, {
    x: 6.55, y: 6.35, w: 2.1, h: 2.1, fill: { color: PURPLE_MD, transparency: 70 },
  });
}

function lightBg(slide) {
  slide.background = { color: WHITE };
}

// page title (brand: small purple eyebrow + strong headline)
function pageTitle(slide, eyebrow, headline) {
  slide.addText(eyebrow, {
    x: 0.62, y: 0.42, w: 8, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE_DK, charSpacing: 2,
  });
  slide.addText(headline, {
    x: 0.6, y: 0.72, w: 10.6, h: 0.62, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 25, bold: true, color: PURPLE_DK,
  });
  slide.addImage({ path: LOGO, x: 11.62, y: 0.42, w: 1.12, h: 0.29 });
}

function pageNum(slide, n) {
  slide.addText(String(n).padStart(2, "0"), {
    x: 12.5, y: 6.92, w: 0.5, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10, color: PURPLE_LT, align: "right",
  });
}

// readable foreground for a given fill (white is illegible on the light brand oranges)
const LIGHT_FILLS = new Set([ORANGE, ORANGE_DK, YELLOW, PURPLE_LT, PURPLE_XLT, WHITE]);
function onColor(bg) { return LIGHT_FILLS.has(bg) ? PURPLE_DK : WHITE; }

// rounded card
function card(slide, o) {
  slide.addShape(pres.ShapeType.roundRect, {
    x: o.x, y: o.y, w: o.w, h: o.h, rectRadius: 0.1,
    fill: { color: o.fill || PURPLE_XLT },
    line: o.line ? { color: o.line, width: 1 } : { type: "none" },
  });
}

// numbered badge circle
function badge(slide, x, y, d, text, bg, fg, size) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: bg } });
  slide.addText(text, {
    x, y, w: d, h: d, isTextBox: true, margin: 0,
    fontFace: F, fontSize: size || 13, bold: true, color: onColor(bg), align: "center", valign: "middle",
  });
}

// =====================================================================
// S1 — Cover
// =====================================================================
{
  const s = pres.addSlide();
  darkBg(s);
  s.addImage({ path: LOGO, x: 0.85, y: 0.72, w: 2.05, h: 0.53 });

  s.addText("转正述职报告", {
    x: 0.85, y: 2.28, w: 9.4, h: 1.0, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 42, bold: true, color: WHITE,
  });
  s.addText("智能宠物项圈 · 犬只行为识别与皮肤健康评估算法研发", {
    x: 0.87, y: 3.36, w: 9.6, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, color: PURPLE_LT,
  });

  // identity block
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.85, y: 4.24, w: 5.55, h: 1.22, rectRadius: 0.1,
    fill: { color: PURPLE, transparency: 35 },
  });
  s.addText("朱业拓  Toky", {
    x: 1.12, y: 4.44, w: 3.2, h: 0.4, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 19, bold: true, color: WHITE,
  });
  s.addText("AI 算法工程师", {
    x: 1.14, y: 4.9, w: 3.2, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: PURPLE_LT,
  });
  s.addText("述职日期", {
    x: 4.62, y: 4.46, w: 1.6, h: 0.26, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10, color: PURPLE_LT,
  });
  s.addText("2026.10.08", {
    x: 4.6, y: 4.76, w: 1.7, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 16, bold: true, color: ORANGE,
  });

  s.addText("Health Innovation for a Clean & Comfortable Life", {
    x: 0.87, y: 6.5, w: 8, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: PURPLE_MD, italic: true,
  });

  // brand motif: + ↗ ♥ (the glyphs inside the HICC logo)
  const motif = [
    { t: "+", x: 10.05, y: 5.28, sz: 34 },
    { t: "↗", x: 10.85, y: 5.28, sz: 34 },
    { t: "♥", x: 11.65, y: 5.32, sz: 30 },
  ];
  motif.forEach(m => {
    s.addText(m.t, {
      x: m.x, y: m.y, w: 0.8, h: 0.88, isTextBox: true, margin: 0,
      fontFace: F, fontSize: m.sz, bold: true, color: ORANGE, align: "center", valign: "middle",
    });
  });
  s.addNotes("开场：各位领导、同事好，我是算法岗的朱业拓，英文名 Toky。今天向大家汇报我在试用期内的工作与成果。");
}

// =====================================================================
// S2 — 目录
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "CONTENTS", "目录");

  const items = [
    ["01", "自我介绍与岗位理解", "我是谁，以及我如何理解这个岗位"],
    ["02", "试用期关键成果", "算法指标 · 工程平台 · 效果演示 · 数据资源 · 跨团队协作"],
    ["03", "持续产出的验证", "进度曲线与迭代节奏，证明可持续交付"],
    ["04", "个人成长与沉淀", "方法论沉淀与认知提升"],
    ["05", "待提升点与下一步", "短板认知与后续规划"],
  ];

  items.forEach((it, i) => {
    const y = 1.72 + i * 1.03;
    card(s, { x: 0.6, y, w: 12.13, h: 0.86, fill: i === 1 ? PURPLE_LT : PURPLE_XLT });
    badge(s, 0.86, y + 0.15, 0.56, it[0], i === 1 ? PURPLE : PURPLE_MD, WHITE, 15);
    s.addText(it[1], {
      x: 1.66, y: y + 0.16, w: 4.2, h: 0.32, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 15, bold: true, color: PURPLE_DK,
    });
    s.addText(it[2], {
      x: 1.68, y: y + 0.5, w: 8.6, h: 0.26, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, color: MUTED,
    });
  });
  pageNum(s, 2);
  s.addNotes("汇报分五部分，重点是第二部分试用期关键成果。");
}

// =====================================================================
// S3 — 自我介绍 & 岗位职责理解
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "01  SELF INTRODUCTION", "自我介绍与岗位职责理解");

  // left: profile
  card(s, { x: 0.6, y: 1.66, w: 4.5, h: 4.9, fill: PURPLE_DK });
  s.addText("朱业拓 / Toky", {
    x: 0.92, y: 1.98, w: 3.9, h: 0.52, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 21, bold: true, color: WHITE,
  });
  s.addText("AI 算法工程师", {
    x: 0.94, y: 2.46, w: 3.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, color: ORANGE,
  });

  const facts = [
    ["入职时间", "2026 年 4 月中旬"],
    ["试用期", "至 2026.10.12"],
    ["所属项目", "智能宠物项圈"],
    ["负责方向", "犬只行为识别算法\n皮肤健康评估算法\n算法工程与数据平台"],
  ];
  let fy = 3.06;
  facts.forEach(f => {
    s.addText(f[0], {
      x: 0.94, y: fy, w: 3.9, h: 0.24, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10, color: PURPLE_MD,
    });
    const lines = f[1].split("\n").length;
    s.addText(f[1], {
      x: 0.94, y: fy + 0.24, w: 3.9, h: 0.28 * lines, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, bold: true, color: WHITE, lineSpacingMultiple: 1.15,
    });
    fy += 0.32 + 0.28 * lines;
  });

  // right: role understanding
  s.addText("我如何理解这个岗位", {
    x: 5.42, y: 1.72, w: 7.3, h: 0.34, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 16, bold: true, color: PURPLE_DK,
  });
  s.addText("算法岗不止是训模型 —— 是把「传感器信号」变成「可信的健康结论」的全链路负责人", {
    x: 5.44, y: 2.1, w: 7.3, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: MUTED,
  });

  const roles = [
    ["数据", "从 0 建立采集—标注—验证闭环，数据质量是算法上限", ORANGE],
    ["模型", "行为识别与健康评估建模，用指标与真实场景双重验证", PURPLE_MD],
    ["工程", "把模型做成稳定可调用的服务，能上线才算落地", PURPLE],
    ["协作", "主动对接产品、后端、硬件、兽医，推动资源到位", ORANGE_DK],
  ];
  roles.forEach((r, i) => {
    const y = 2.62 + i * 1.0;
    card(s, { x: 5.42, y, w: 7.31, h: 0.84, fill: PURPLE_XLT });
    badge(s, 5.66, y + 0.16, 0.52, r[0], r[2], WHITE, 12);
    s.addText(r[1], {
      x: 6.36, y: y + 0.14, w: 6.2, h: 0.56, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12, color: INK, valign: "middle", lineSpacingMultiple: 1.2,
    });
  });
  pageNum(s, 3);
  s.addNotes("我把算法岗理解为四件事的闭环：数据、模型、工程、协作。缺任何一环，算法都落不了地。");
}

// =====================================================================
// S4 — 成果总览（数据看板）
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02  KEY RESULTS", "试用期关键成果总览");

  s.addText("从「无数据、无模型、无平台」起步，到算法能力成型、服务上线、数据体系跑通", {
    x: 0.62, y: 1.44, w: 11.4, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: MUTED,
  });

  const stats = [
    ["45% → 87%", "算法整体进度", "覆盖 16 周持续推进", PURPLE],
    ["70% → 90%", "抓挠识别 F1", "8.31 目标 ≥85% 提前达成", ORANGE_DK],
    ["4 类", "行为识别能力", "抓挠 / 活动 / 睡觉 / 未佩戴", PURPLE_MD],
    ["< 0.1s", "视频-IMU 同步误差", "攻克标注核心阻塞点", PURPLE],
    ["2 套", "皮肤评估方案", "规则统计版 + ML 双模型版", ORANGE_DK],
    ["0 → 4+6", "可采集犬只资源", "影棚 4 只 + 狗场 6 只落地", PURPLE_MD],
  ];

  stats.forEach((st, i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.6 + col * 4.11;
    const y = 1.94 + row * 2.36;
    card(s, { x, y, w: 3.91, h: 2.16, fill: PURPLE_XLT });
    s.addText(st[0], {
      x: x + 0.28, y: y + 0.28, w: 3.4, h: 0.66, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 30, bold: true, color: st[3],
    });
    s.addText(st[1], {
      x: x + 0.3, y: y + 1.04, w: 3.4, h: 0.32, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, bold: true, color: PURPLE_DK,
    });
    s.addText(st[2], {
      x: x + 0.3, y: y + 1.42, w: 3.4, h: 0.5, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: MUTED, lineSpacingMultiple: 1.2,
    });
  });
  pageNum(s, 4);
  s.addNotes("这一页是总览。六个数字概括试用期：进度、模型指标、能力覆盖、工程攻坚、方案产出、数据资源。");
}

// =====================================================================
// S5 — 抓挠识别：从 0 到达标
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.1  CORE MODEL", "核心成果一：抓挠识别模型，从 0 到超额达标");

  const chartData = [
    { name: "抓挠识别 F1", labels: ["7.20 初版", "7.27 迭代", "8.03 达标", "8.24 全量复核"], values: [72, 80, 85, 90] },
    { name: "8.31 目标线 85%", labels: ["7.20 初版", "7.27 迭代", "8.03 达标", "8.24 全量复核"], values: [85, 85, 85, 85] },
  ];
  s.addChart(pres.ChartType.line, chartData, {
    x: 0.6, y: 1.66, w: 7.45, h: 3.48,
    chartColors: [PURPLE, ORANGE],
    lineDataSymbol: "circle", lineDataSymbolSize: 7, lineSize: 3,
    showTitle: false, showLegend: true, legendPos: "b", legendFontSize: 10, legendColor: MUTED,
    showValue: true, dataLabelPosition: "t", dataLabelFontSize: 11,
    dataLabelColor: PURPLE_DK, dataLabelFormatCode: '0"%"',
    valAxisMinVal: 60, valAxisMaxVal: 100,
    catAxisLabelColor: MUTED, catAxisLabelFontSize: 10,
    valAxisLabelColor: MUTED, valAxisLabelFontSize: 10,
    valGridLine: { color: "EEEBF7", size: 1 },
    catGridLine: { style: "none" },
  });

  // right rail milestones
  const ms = [
    ["起点", "7.20 初版模型 F1 70-75%", "首次产出可用模型，拿到真实误检分布"],
    ["方法", "按置信度分层复核误检 / 漏检", "以模型结果反哺数据集，定向补负样本"],
    ["达标", "8.03 F1 ≥85%，提前 4 周", "8.31 目标提前完成"],
    ["深化", "8.24 全量复核后训练犬达 90%", "工程化检测人工标注错误并重标注"],
  ];
  ms.forEach((m, i) => {
    const y = 1.66 + i * 0.9;
    card(s, { x: 8.32, y, w: 4.41, h: 0.76, fill: i === 2 ? PURPLE_LT : PURPLE_XLT });
    s.addShape(pres.ShapeType.roundRect, {
      x: 8.54, y: y + 0.1, w: 0.72, h: 0.24, rectRadius: 0.06,
      fill: { color: i === 2 ? ORANGE_DK : PURPLE_MD },
    });
    s.addText(m[0], {
      x: 8.54, y: y + 0.1, w: 0.72, h: 0.24, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 9.5, bold: true, color: i === 2 ? PURPLE_DK : WHITE, align: "center", valign: "middle",
    });
    s.addText(m[1], {
      x: 9.4, y: y + 0.09, w: 3.15, h: 0.26, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true, color: PURPLE_DK,
    });
    s.addText(m[2], {
      x: 8.56, y: y + 0.4, w: 4.0, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 9.5, color: i === 2 ? PURPLE_DK : MUTED,
    });
  });

  // bottom insight strip
  card(s, { x: 0.6, y: 5.42, w: 12.13, h: 1.06, fill: PURPLE_DK });
  s.addText("关键做法", {
    x: 0.88, y: 5.62, w: 1.5, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE,
  });
  s.addText("不盲目堆数据 —— 用模型输出反推数据缺口：按置信度区间分析误识别（走路、跑、扒门被误判），定向补充负样本；再用工程化手段批量检出人工标注错误并重标注。数据质量的提升比数据量的增长更能带来实质效果。", {
    x: 2.4, y: 5.58, w: 10.1, h: 0.74, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: WHITE, valign: "middle", lineSpacingMultiple: 1.25,
  });
  pageNum(s, 5);
  s.addNotes("这是试用期最核心的算法成果。8月31日的目标是F1≥85%，我在8月3日提前四周达成，随后通过全量复核把训练个体上的准确率推到90%。");
}

// =====================================================================
// S6 — 算法能力矩阵
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.2  CAPABILITY", "核心成果二：算法能力矩阵与皮肤健康评估方案");

  // left: 4-class behavior
  s.addText("行为识别：4 分类能力已成型", {
    x: 0.62, y: 1.62, w: 6, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: PURPLE_DK,
  });
  const beh = [
    ["抓挠识别", "已达标，训练个体 F1 90%", ORANGE_DK, "达标"],
    ["活动检测", "基础可用 80%+", PURPLE_MD, "可用"],
    ["睡觉 / 休息", "基础可用 80%+", PURPLE_MD, "可用"],
    ["未佩戴检测", "已纳入 4 分类体系并完成测试", PURPLE, "完成"],
  ];
  beh.forEach((b, i) => {
    const y = 2.06 + i * 0.83;
    card(s, { x: 0.6, y, w: 6.06, h: 0.7, fill: PURPLE_XLT });
    s.addText(b[0], {
      x: 0.88, y: y + 0.09, w: 2.1, h: 0.26, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
    });
    s.addText(b[1], {
      x: 0.9, y: y + 0.38, w: 3.9, h: 0.24, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10, color: MUTED,
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: 5.62, y: y + 0.21, w: 0.78, h: 0.28, rectRadius: 0.07, fill: { color: b[2] },
    });
    s.addText(b[3], {
      x: 5.62, y: y + 0.21, w: 0.78, h: 0.28, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 9.5, bold: true, color: onColor(b[2]), align: "center", valign: "middle",
    });
  });

  // right: skin assessment dual track
  s.addText("皮肤健康评估：双方案并行推进", {
    x: 6.94, y: 1.62, w: 6, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: PURPLE_DK,
  });

  card(s, { x: 6.92, y: 2.06, w: 5.81, h: 1.62, fill: PURPLE_XLT });
  s.addText("规则统计版（SBS）", {
    x: 7.18, y: 2.22, w: 3.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
  });
  s.addText([
    { text: "前后端 Web 系统完成，兽医问答交互可用", options: { bullet: true, breakLine: true } },
    { text: "联通每日抓挠数据，驱动规则引擎计算", options: { bullet: true, breakLine: true } },
    { text: "可解释性强，供产品与兽医直接验证反馈", options: { bullet: true } },
  ], {
    x: 7.2, y: 2.56, w: 5.3, h: 1.0, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 10.5, color: INK, paraSpaceAfter: 4,
  });

  card(s, { x: 6.92, y: 3.84, w: 5.81, h: 1.86, fill: PURPLE_DK });
  s.addText("机器学习版（双模型架构）", {
    x: 7.18, y: 4.0, w: 3.6, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: ORANGE,
  });
  s.addText([
    { text: "模型 A：特征表 → C0 / C1 / C2 皮肤状态分级", options: { bullet: true, breakLine: true } },
    { text: "模型 B：特征 + A 输出 + 用户问答 → S0 / S1 / S2", options: { bullet: true, breakLine: true } },
    { text: "合成数据模拟 86 个场景，F1 达 93%+", options: { bullet: true, breakLine: true } },
    { text: "工程框架已建立，可随真实数据持续迭代", options: { bullet: true } },
  ], {
    x: 7.2, y: 4.34, w: 5.3, h: 1.28, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 10.5, color: WHITE, paraSpaceAfter: 4,
  });

  // bottom insight
  card(s, { x: 0.6, y: 5.86, w: 12.13, h: 0.86, fill: PURPLE_LT });
  s.addText("真实验证发现的关键问题", {
    x: 0.88, y: 6.0, w: 2.5, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: PURPLE_DK,
  });
  s.addText("4 只犬 5 天测试中，两只犬日均抓挠 20+ 次触发 C2，但皮肤无实际病变 —— 说明个体「正常抓挠水平」本就存在差异。据此提出引入个体基线机制，比单纯调权重更根本。", {
    x: 3.5, y: 5.98, w: 9.0, h: 0.62, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: INK, valign: "middle", lineSpacingMultiple: 1.2,
  });
  pageNum(s, 6);
  s.addNotes("行为识别四分类已成型；皮肤评估采取规则版与ML版双线并行，在数据不充足时降低单线失败风险。");
}

// =====================================================================
// S7 — 工程与平台建设
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.3  ENGINEERING", "核心成果三：算法工程与数据平台建设");

  s.addText("模型能跑通只是第一步 —— 试用期内把算法做成了可采集、可标注、可部署、可调用的完整链路", {
    x: 0.62, y: 1.44, w: 11.6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  const eng = [
    ["01", "采集与同步工程", [
      "攻克视频-IMU 时间戳对齐，误差控制在 0.1s 以内",
      "RTSP 方案验证不可行后果断放弃，收敛至 USB 本地录制",
      "采样率 25Hz→16Hz，依奈奎斯特论证降功耗与存储",
    ]],
    ["02", "实时推理工程", [
      "WitMotion-IMU / HICC-IMU 蓝牙接收实时推理",
      "现场采集即可实时验证模型效果",
      "修复线上线下推理不一致（特征处理流程未对齐）",
    ]],
    ["03", "数据标注平台", [
      "Label Studio 部署改造 + 自研平台建设",
      "NAS 原始数据挂载、成员分级权限与数据隔离",
      "定位并修复卡顿瓶颈，保障多人协同标注效率",
    ]],
    ["04", "服务化与部署", [
      "模型 pipeline + 服务 pipeline，合并调度简化运维",
      "模型服务打包并完成外网部署，后端对接打通",
      "模拟 24 场景 / 180 天数据，压测 1 万台设备并发",
    ]],
  ];

  eng.forEach((e, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.19;
    const y = 1.88 + row * 2.44;
    card(s, { x, y, w: 5.94, h: 2.24, fill: PURPLE_XLT });
    badge(s, x + 0.26, y + 0.24, 0.5, e[0], PURPLE, WHITE, 12);
    s.addText(e[1], {
      x: x + 0.92, y: y + 0.3, w: 4.7, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 13.5, bold: true, color: PURPLE_DK,
    });
    s.addText(e[2].map((t, k) => ({
      text: t, options: { bullet: true, breakLine: k !== e[2].length - 1 },
    })), {
      x: x + 0.3, y: y + 0.82, w: 5.36, h: 1.28, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10.5, color: INK, paraSpaceAfter: 5, lineSpacingMultiple: 1.1,
    });
  });
  pageNum(s, 7);
  s.addNotes("工程能力是我认为算法岗容易被忽视但很关键的部分。模型服务外网部署与后端打通，意味着算法结果真正进入了产品使用流程。");
}

// =====================================================================
// S7b — 算法效果与平台演示（截图存在则自动嵌入，否则降级为演示面板）
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.6  LIVE DEMO", "算法效果与平台演示");

  s.addText("以下三项将在述职现场做实际演示，此处为效果留档", {
    x: 0.62, y: 1.44, w: 11.6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  const demos = [
    ["media/scratch.png", "抓挠检测效果", "IMU 信号实时推理，视频与识别结果同步比对，可直观看到抓挠片段的检出"],
    ["media/platform.png", "数据标注平台", "NAS 数据挂载、成员分级权限、视频标注流畅度优化后的实际操作界面"],
    ["media/tartar.png", "口腔牙结石识别", "基于图像识别犬只口腔牙齿的牙结石情况，为口腔健康评估提供基础能力"],
  ];

  demos.forEach((dm, i) => {
    const x = 0.6 + i * 4.11;
    card(s, { x, y: 1.9, w: 3.91, h: 4.6, fill: PURPLE_XLT });

    const frameX = x + 0.22, frameY = 2.12, frameW = 3.47, frameH = 2.34;
    const exists = fs.existsSync(dm[0]);
    if (exists) {
      s.addImage({ path: dm[0], x: frameX, y: frameY, w: frameW, h: frameH, sizing: { type: "cover", w: frameW, h: frameH } });
    } else {
      s.addShape(pres.ShapeType.roundRect, {
        x: frameX, y: frameY, w: frameW, h: frameH, rectRadius: 0.08,
        fill: { color: PURPLE_DK },
      });
      s.addText("▶", {
        x: frameX, y: frameY + 0.72, w: frameW, h: 0.6, isTextBox: true, margin: 0,
        fontFace: F, fontSize: 28, bold: true, color: ORANGE, align: "center", valign: "middle",
      });
      s.addText("现场演示", {
        x: frameX, y: frameY + 1.34, w: frameW, h: 0.32, isTextBox: true, margin: 0,
        fontFace: F, fontSize: 12, bold: true, color: WHITE, align: "center", valign: "middle",
      });
    }

    s.addText(dm[1], {
      x: x + 0.24, y: 4.62, w: 3.44, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, bold: true, color: PURPLE_DK,
    });
    s.addText(dm[2], {
      x: x + 0.24, y: 5.04, w: 3.44, h: 1.2, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10.5, color: INK, lineSpacingMultiple: 1.3,
    });
  });
  pageNum(s, 8);
  s.addNotes("这一页配合现场演示：先放抓挠检测的实际效果，再演示标注平台的操作，最后展示口腔牙结石识别。演示时长控制在 2 分钟内。");
}

// =====================================================================
// S8 — 数据资源从 0 到 1
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.4  DATA ASSETS", "核心成果四：数据资源体系从 0 到 1");

  s.addText("数据是算法的上限。试用期内主动推动，把「无数据可用」变成「稳定可持续的采集供给」", {
    x: 0.62, y: 1.44, w: 11.6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  // timeline
  const tl = [
    ["5 月", "模拟数据起步", "自研 IMU 数据模拟生成器，跑通端到端流程", PURPLE_MD],
    ["6 月", "首批真实数据", "影棚落地采集，攻克时间戳对齐，标注工具就绪", PURPLE],
    ["7 月", "采集标准化", "方案定型：IMU + 双视角 USB 摄像头；犬只扩至 3 只", PURPLE],
    ["8 月", "规模化铺路", "输出 2 份采集方案供商务洽谈；沙井成本过高转向龙岗", ORANGE_DK],
    ["9 月", "狗场落地", "龙岗实地考察通过，6 间 6 只具备正式采集条件", ORANGE_DK],
  ];

  const tlY = 2.0;
  tl.forEach((t, i) => {
    const x = 0.6 + i * 2.46;
    card(s, { x, y: tlY, w: 2.32, h: 2.7, fill: PURPLE_XLT });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.24, y: tlY + 0.22, w: 0.96, h: 0.34, rectRadius: 0.08, fill: { color: t[3] },
    });
    s.addText(t[0], {
      x: x + 0.24, y: tlY + 0.22, w: 0.96, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true, color: onColor(t[3]), align: "center", valign: "middle",
    });
    s.addText(t[1], {
      x: x + 0.24, y: tlY + 0.72, w: 1.9, h: 0.5, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK, lineSpacingMultiple: 1.1,
    });
    s.addText(t[2], {
      x: x + 0.24, y: tlY + 1.24, w: 1.88, h: 1.2, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10, color: MUTED, lineSpacingMultiple: 1.25,
    });
  });

  // bottom: two outcome cards
  card(s, { x: 0.6, y: 4.92, w: 5.94, h: 1.62, fill: PURPLE_DK });
  s.addText("可采集犬只资源", {
    x: 0.88, y: 5.1, w: 3.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, bold: true, color: ORANGE,
  });
  s.addText("影棚 0 → 4 只（大 / 中 / 小体型覆盖）  ·  龙岗狗场 6 间 6 只已谈妥", {
    x: 0.9, y: 5.44, w: 5.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: WHITE,
  });
  s.addText("长远按 30 → 50 → 70 → 100 → 150 只分阶段扩大，先跑通小规模流程再有序放量", {
    x: 0.9, y: 5.8, w: 5.4, h: 0.5, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10, color: PURPLE_LT, lineSpacingMultiple: 1.2,
  });

  card(s, { x: 6.79, y: 4.92, w: 5.94, h: 1.62, fill: PURPLE_LT });
  s.addText("主动推动的关键判断", {
    x: 7.07, y: 5.1, w: 3.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, bold: true, color: PURPLE_DK,
  });
  s.addText("沙井狗场报价 1000 元/犬/天，判断成本不经济果断放弃；主动输出采集需求方案供商务对接，转向龙岗并推动落地，避免了资源与时间的无效投入。", {
    x: 7.09, y: 5.44, w: 5.4, h: 0.9, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: INK, lineSpacingMultiple: 1.25,
  });
  pageNum(s, 9);
  s.addNotes("数据资源不能被动等待。从模拟数据起步，到影棚落地，再到狗场合作谈成，这条线是我主动推动的。");
}

// =====================================================================
// S9 — 跨团队协作
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.5  COLLABORATION", "核心成果五：跨团队协作与主动推动");

  const collab = [
    ["后端", "数据库架构调整（IMU 六轴迁移 TDengine）、服务联调、模型服务外网对接打通", PURPLE],
    ["产品", "PRD 对接、明确算法实现边界、皮肤评估方案共同推敲与验证反馈", PURPLE_MD],
    ["硬件厂商", "定位并反馈 TF 卡周期性丢数据、蓝牙遮挡断传等问题，推动修复与设备交付", ORANGE_DK],
    ["兽医 / 产品", "示范标注工具用法，建立多人协同标注分工，推动兽医参与皮肤评估", PURPLE_MD],
    ["测试 / 项目", "预留测试接口；推动算法研发在 ONES 单独立项，建立规范化管理", PURPLE],
  ];

  collab.forEach((c, i) => {
    const y = 1.7 + i * 0.86;
    card(s, { x: 0.6, y, w: 7.9, h: 0.74, fill: PURPLE_XLT });
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.84, y: y + 0.2, w: 1.28, h: 0.34, rectRadius: 0.08, fill: { color: c[0 + 2] },
    });
    s.addText(c[0], {
      x: 0.84, y: y + 0.2, w: 1.28, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true, color: onColor(c[2]), align: "center", valign: "middle",
    });
    s.addText(c[1], {
      x: 2.28, y: y + 0.1, w: 6.0, h: 0.54, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, color: INK, valign: "middle", lineSpacingMultiple: 1.2,
    });
  });

  // right: proactive proposals
  card(s, { x: 8.76, y: 1.7, w: 3.97, h: 4.3, fill: PURPLE_DK });
  s.addText("主动提出的建议", {
    x: 9.02, y: 1.92, w: 3.4, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: ORANGE,
  });
  const props = [
    ["质疑临床阈值", "实测博美一晚抓挠 10-20 次，推算全天极可能超 30 次但皮肤无异常，提出需重新调研基线依据"],
    ["建议竞品对标", "建议采购 FitBark / Tractive 等竞品实测，同步佩戴对比数据与识别效果"],
    ["推动规范立项", "即便算法仅一人，也应在 ONES 立项，明确阶段目标与验收标准"],
  ];
  props.forEach((p, i) => {
    const y = 2.36 + i * 1.24;
    s.addText(p[0], {
      x: 9.02, y, w: 3.5, h: 0.28, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11.5, bold: true, color: WHITE,
    });
    s.addText(p[1], {
      x: 9.04, y: y + 0.3, w: 3.46, h: 0.86, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 9.5, color: PURPLE_LT, lineSpacingMultiple: 1.25,
    });
  });
  pageNum(s, 10);
  s.addNotes("算法虽然只有我一个人，但工作是高度依赖协作的。我也会主动提出自己的判断，而不只是执行。");
}

// =====================================================================
// S10 — 持续产出的验证
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "03  CONSISTENCY", "我如何证明能够持续产出");

  const prog = [{
    name: "算法整体进度",
    labels: ["5.18", "5.25", "7.27", "8.03", "8.10", "8.17", "8.24", "8.31"],
    values: [45, 55, 70, 75, 76, 83, 85, 87],
  }];
  s.addChart(pres.ChartType.line, prog, {
    x: 0.6, y: 1.72, w: 7.45, h: 3.3,
    chartColors: [PURPLE],
    lineDataSymbol: "circle", lineDataSymbolSize: 6, lineSize: 3,
    showTitle: false, showLegend: false,
    showValue: true, dataLabelPosition: "t", dataLabelFontSize: 10,
    dataLabelColor: PURPLE_DK, dataLabelFormatCode: '0"%"',
    valAxisMinVal: 30, valAxisMaxVal: 100,
    catAxisLabelColor: MUTED, catAxisLabelFontSize: 10,
    valAxisLabelColor: MUTED, valAxisLabelFontSize: 10,
    valGridLine: { color: "EEEBF7", size: 1 },
    catGridLine: { style: "none" },
  });
  s.addText("算法整体进度：45% → 87%（16 周持续推进，无停滞周）", {
    x: 0.62, y: 5.06, w: 7.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: PURPLE_DK,
  });

  const proof = [
    ["16 篇", "连续周报", "每周输出完成项、进度、复盘、风险与计划，节奏无中断", ORANGE_DK],
    ["每项有结论", "汇报习惯", "每个工作项都给出明确结论，而非只罗列过程", PURPLE],
    ["问题必闭环", "工作方式", "阻塞点定位到根因再解决：时间戳对齐、推理不一致、标注错误", PURPLE],
  ];
  proof.forEach((p, i) => {
    const y = 1.72 + i * 1.16;
    card(s, { x: 8.32, y, w: 4.41, h: 1.02, fill: PURPLE_XLT });
    s.addText(p[0], {
      x: 8.56, y: y + 0.12, w: 2.0, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 16, bold: true, color: p[3],
    });
    s.addText(p[1], {
      x: 10.62, y: y + 0.17, w: 1.9, h: 0.26, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: MUTED, align: "right",
    });
    s.addText(p[2], {
      x: 8.58, y: y + 0.5, w: 3.9, h: 0.44, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10, color: INK, lineSpacingMultiple: 1.2,
    });
  });

  card(s, { x: 0.6, y: 5.5, w: 12.13, h: 0.98, fill: PURPLE_DK });
  s.addText("交付节奏", {
    x: 0.88, y: 5.68, w: 1.5, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE,
  });
  s.addText("模型从初版到达标用了 6 周、皮肤评估从方案到 Web 系统上线用了 3 周、数据平台从 0 到多人协同可用用了 4 周 —— 每个模块都有明确的起点、节奏与交付结果。", {
    x: 2.4, y: 5.64, w: 10.1, h: 0.7, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: WHITE, valign: "middle", lineSpacingMultiple: 1.25,
  });
  pageNum(s, 11);
  s.addNotes("持续产出不是靠某一次冲刺，而是稳定的节奏。16周周报、进度从45%到87%，每周都有实质推进。");
}

// =====================================================================
// S11 — 个人成长与沉淀
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "04  GROWTH", "个人成长与方法论沉淀");

  const growth = [
    ["数据质量 > 数据量", "全量复核并修正标注错误后训练的模型，效果优于单纯增加数据量。定期全量复核与工程化错误检测，比盲目堆数据更有效。", ORANGE_DK],
    ["用模型输出反推数据缺口", "按置信度区间分层分析误识别与漏识别，定向补充负样本，形成「测试发现问题→补充数据→迭代」的闭环，比增量数据更高效。", PURPLE],
    ["先分清是算法问题还是定义问题", "活动与睡觉混淆的根因是业务边界模糊，不是模型能力不足。算法无法自行划定业务语义，须先对齐定义再训练。", PURPLE_MD],
    ["工程能力决定算法能否落地", "服务化部署、接口稳定、与后端联通，才是算法进入产品流程的标志。效果好只是研发侧的里程碑。", PURPLE],
    ["及时验证、及时放弃", "RTSP 方案在验证发现网络抖动导致延迟不可预测后果断放弃，收敛至可控方案，避免沉没成本继续扩大。", ORANGE_DK],
  ];

  growth.forEach((g, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.19;
    const y = 1.66 + row * 1.72;
    const w = i === 4 ? 12.13 : 5.94;
    card(s, { x, y, w, h: 1.54, fill: PURPLE_XLT });
    badge(s, x + 0.26, y + 0.22, 0.46, String(i + 1), g[2], WHITE, 12);
    s.addText(g[0], {
      x: x + 0.86, y: y + 0.26, w: w - 1.2, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 13.5, bold: true, color: PURPLE_DK,
    });
    s.addText(g[1], {
      x: x + 0.3, y: y + 0.74, w: w - 0.6, h: 0.68, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10.5, color: INK, lineSpacingMultiple: 1.25,
    });
  });
  pageNum(s, 12);
  s.addNotes("这五条是我试用期最重要的方法论沉淀，也是我认为可以持续复用到后续工作中的东西。");
}

// =====================================================================
// S12 — 待提升点与改进计划
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "05  GAPS & PLAN", "待提升点与改进计划");

  s.addText("客观看待当前短板，每一项都有明确的改进路径", {
    x: 0.62, y: 1.44, w: 11.6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  // header row
  const cols = [[0.6, 3.2, "待提升点"], [3.9, 4.3, "现状与原因"], [8.3, 4.43, "改进计划"]];
  cols.forEach(c => {
    s.addText(c[2], {
      x: c[0] + 0.24, y: 1.9, w: c[1] - 0.4, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true, color: PURPLE_MD,
    });
  });

  const gaps = [
    ["跨个体泛化能力不足", "新个体马尔济斯 F1 仅 70%，训练个体达 90%。根因是训练覆盖的犬只个体数太少，品种、体型、运动习惯差异大", "推动狗场采集落地，将训练个体扩至 10 只以上，再以全新个体验证泛化，目标 85%+"],
    ["松动检测尚未启动", "优先级让位于抓挠识别与皮肤评估；且松动判定边界本身需与产品、兽医对齐才能定义清楚", "与产品、兽医对齐判定标准后启动开发，与未佩戴检测共同构成全天有效佩戴时间统计"],
    ["皮肤评估依赖合成数据", "ML 版仍基于合成数据训练，真实病例样本积累不足，个体基线机制尚未落地", "持续推动兽医参与评估标注，积累真实数据做二次训练；优先开发个体基线机制"],
    ["业务理解仍需加深", "对犬只行为的业务语义、兽医临床判断标准的理解，仍主要依赖他人输入，独立判断力有待加强", "主动向兽医与产品请教，系统补充宠物皮肤健康领域知识，减少反复对齐成本"],
  ];

  gaps.forEach((g, i) => {
    const y = 2.28 + i * 1.14;
    card(s, { x: 0.6, y, w: 12.13, h: 1.02, fill: i % 2 === 0 ? PURPLE_XLT : WHITE, line: i % 2 === 0 ? null : PURPLE_LT });
    badge(s, 0.84, y + 0.32, 0.4, String(i + 1), PURPLE_MD, WHITE, 11);
    s.addText(g[0], {
      x: 1.36, y: y + 0.34, w: 2.4, h: 0.36, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12, bold: true, color: PURPLE_DK, valign: "middle",
    });
    s.addText(g[1], {
      x: 4.14, y: y + 0.12, w: 4.0, h: 0.8, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10, color: MUTED, valign: "middle", lineSpacingMultiple: 1.2,
    });
    s.addText(g[2], {
      x: 8.54, y: y + 0.12, w: 4.0, h: 0.8, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10, color: INK, valign: "middle", lineSpacingMultiple: 1.2,
    });
  });
  pageNum(s, 13);
  s.addNotes("这一页是我对自己短板的判断。泛化能力是当前最需要解决的问题，其他三项也都有明确路径。");
}

// =====================================================================
// S13 — 下一步规划
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "05  NEXT STEPS", "下一步：与团队共同达成目标");

  const plans = [
    ["近期", "1-2 个月", [
      "推动龙岗狗场正式采集落地，训练个体扩至 10 只以上",
      "数据平台上线 AI 自动标注，加速人工审核效率",
      "完成皮肤评估个体基线机制，解决 C2 误触发",
      "与产品、兽医对齐活动/睡觉业务边界与松动判定标准",
    ], PURPLE],
    ["中期", "3-6 个月", [
      "抓挠识别跨个体泛化达到 85%+，模型具备真正通用性",
      "皮肤健康评估完成真实数据二次训练并上线验证",
      "补齐松动检测，形成完整的全天有效佩戴时间统计",
      "采集规模按 30 → 50 → 70 只分阶段有序扩大",
    ], PURPLE_MD],
    ["长期", "持续投入", [
      "沉淀可复用的采集—标注—训练—部署标准流程",
      "推动算法能力从「能识别」走向「能给出可信健康结论」",
      "建立竞品对标机制，明确我方算法能力边界与优势",
      "随团队规模扩大，输出规范让算法工作可协同、可交接",
    ], ORANGE_DK],
  ];

  plans.forEach((p, i) => {
    const x = 0.6 + i * 4.11;
    card(s, { x, y: 1.7, w: 3.91, h: 3.34, fill: i === 0 ? PURPLE_DK : PURPLE_XLT });
    const isDark = i === 0;
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.28, y: 1.96, w: 1.0, h: 0.36, rectRadius: 0.08, fill: { color: p[3] },
    });
    s.addText(p[0], {
      x: x + 0.28, y: 1.96, w: 1.0, h: 0.36, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12, bold: true, color: onColor(p[3]), align: "center", valign: "middle",
    });
    s.addText(p[1], {
      x: x + 1.42, y: 2.02, w: 2.2, h: 0.26, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: isDark ? PURPLE_LT : MUTED,
    });
    s.addText(p[2].map((t, k) => ({
      text: t, options: { bullet: true, breakLine: k !== p[2].length - 1 },
    })), {
      x: x + 0.3, y: 2.5, w: 3.32, h: 2.4, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10.5, color: isDark ? WHITE : INK,
      paraSpaceAfter: 9, lineSpacingMultiple: 1.25,
    });
  });

  card(s, { x: 0.6, y: 5.32, w: 12.13, h: 1.34, fill: PURPLE_DK });
  s.addText("共同目标", {
    x: 0.9, y: 5.5, w: 2.0, h: 0.26, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE,
  });
  s.addText("让算法从「跑得通」走向「靠得住」—— 成为产品可以放心依赖的健康判断能力", {
    x: 0.92, y: 5.8, w: 11.4, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15.5, bold: true, color: WHITE,
  });
  s.addText("这需要数据规模、算法能力与业务定义三者同步推进，也需要产品、兽医、后端、硬件各方持续配合，我会主动承担起推动这条链路的责任。", {
    x: 0.92, y: 6.22, w: 11.4, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: PURPLE_LT,
  });
  pageNum(s, 14);
  s.addNotes("下一步我最想解决的是泛化能力问题，这决定了产品能不能面向所有用户的狗。");
}

// =====================================================================
// S14 — Thank you
// =====================================================================
{
  const s = pres.addSlide();
  darkBg(s);
  s.addImage({ path: LOGO, x: 0.85, y: 0.72, w: 2.05, h: 0.53 });

  s.addText("感谢聆听 · 恳请指正", {
    x: 0.85, y: 2.72, w: 9.0, h: 0.8, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 36, bold: true, color: WHITE,
  });
  s.addText("THANK YOU", {
    x: 0.87, y: 3.6, w: 7.0, h: 0.44, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 19, color: ORANGE, charSpacing: 5,
  });
  s.addText("朱业拓  Toky   |   AI 算法工程师   |   2026.10.08", {
    x: 0.87, y: 4.42, w: 8.0, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: PURPLE_LT,
  });
  s.addText("Health Innovation for a Clean & Comfortable Life", {
    x: 0.87, y: 6.5, w: 8, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: PURPLE_MD, italic: true,
  });

  const motif = [
    { t: "+", x: 10.05, y: 5.28, sz: 34 },
    { t: "↗", x: 10.85, y: 5.28, sz: 34 },
    { t: "♥", x: 11.65, y: 5.32, sz: 30 },
  ];
  motif.forEach(m => {
    s.addText(m.t, {
      x: m.x, y: m.y, w: 0.8, h: 0.88, isTextBox: true, margin: 0,
      fontFace: F, fontSize: m.sz, bold: true, color: ORANGE, align: "center", valign: "middle",
    });
  });
  s.addNotes("感谢各位，欢迎提问。");
}

pres.writeFile({ fileName: "转正述职报告-朱业拓.pptx" }).then(f => console.log("written:", f));
