const pptxgen = require("pptxgenjs");
const fs = require("fs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.author = "Toky 朱业拓";
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

// fit an image inside a panel without distorting it; letterbox on a dark panel
function fitImage(slide, imgPath, px, py, pw, ph) {
  const dim = require("image-size").imageSize(fs.readFileSync(imgPath));
  const ar = dim.width / dim.height;
  let w = pw, h = pw / ar;
  if (h > ph) { h = ph; w = ph * ar; }
  slide.addImage({ path: imgPath, x: px + (pw - w) / 2, y: py + (ph - h) / 2, w, h });
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
  s.addText("Wardyn 项目 · 智能宠物项圈 · 犬只行为识别与皮肤健康评估算法研发", {
    x: 0.87, y: 3.36, w: 9.6, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, color: PURPLE_LT,
  });

  // identity block
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.85, y: 4.24, w: 5.55, h: 1.22, rectRadius: 0.1,
    fill: { color: PURPLE, transparency: 35 },
  });
  s.addText([
    { text: "Toky", options: { fontSize: 20, bold: true, color: WHITE } },
    { text: "  朱业拓", options: { fontSize: 13, bold: false, color: PURPLE_LT } },
  ], {
    x: 1.12, y: 4.4, w: 3.4, h: 0.52, isTextBox: true, margin: 0, fontFace: F,
  });
  s.addText("算法工程师", {
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
    ["02", "试用期关键成果", "算法能力 · 系统平台 · 效果演示 · 数据资源 · 跨团队协作"],
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
  s.addText([
    { text: "Toky", options: { fontSize: 22, bold: true, color: WHITE } },
    { text: "  朱业拓", options: { fontSize: 14, bold: false, color: PURPLE_LT } },
  ], {
    x: 0.92, y: 1.94, w: 3.9, h: 0.62, isTextBox: true, margin: 0, fontFace: F,
  });
  s.addText("算法工程师", {
    x: 0.94, y: 2.46, w: 3.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, color: ORANGE,
  });

  const facts = [
    ["所属部门", "用户增长研发部"],
    ["入职时间", "2026 年 4 月中旬（试用期至 10.12）"],
    ["所属项目", "Wardyn 项目 · 智能宠物项圈"],
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
// S4 — 成果总览：从 0 到 1
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02  KEY RESULTS", "试用期关键成果：从 0 到 1");

  s.addText("入职时算法侧没有数据、没有模型、没有平台 —— 这六件事都是从零跑通的", {
    x: 0.62, y: 1.44, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  // header row
  const cols = [[0.6, 3.05, "能力项"], [3.75, 3.1, "入职时"], [7.0, 5.73, "现在"]];
  cols.forEach(c => {
    s.addText(c[2], {
      x: c[0] + 0.26, y: 1.9, w: c[1] - 0.4, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true, color: PURPLE_MD,
    });
  });

  const rows = [
    ["行为识别", "无模型，无数据", "四类行为识别可用，稳定版 v2"],
    ["皮肤健康评估", "只有一份待定方案", "规则版与 ML 版闭环跑通，每日自动出结果"],
    ["数据采集", "无设备，无场地", "6 机位 × 12 设备无人值守，视频与 IMU 帧级同步"],
    ["标注体系", "无工具，无流程", "自研平台上线，AI 预标注 + 多人协同审核"],
    ["在线服务", "无", "已交付后端，App → 后端 → 算法跑通"],
    ["可采集犬只", "0 只", "影棚 4 只，龙岗狗场 6 只待进场"],
  ];

  rows.forEach((r, i) => {
    const y = 2.28 + i * 0.74;
    card(s, { x: 0.6, y, w: 12.13, h: 0.64, fill: i % 2 === 0 ? PURPLE_XLT : WHITE, line: i % 2 === 0 ? null : PURPLE_LT });
    s.addText(r[0], {
      x: 0.86, y: y + 0.16, w: 2.7, h: 0.32, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
    });
    s.addText(r[1], {
      x: 4.0, y: y + 0.18, w: 2.85, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: MUTED,
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: 6.86, y: y + 0.24, w: 0.24, h: 0.16, rectRadius: 0.04, fill: { color: ORANGE },
    });
    s.addText(r[2], {
      x: 7.26, y: y + 0.14, w: 5.3, h: 0.38, isTextBox: true, margin: 0, valign: "middle",
      fontFace: F, fontSize: 10.5, color: INK, lineSpacingMultiple: 1.15,
    });
  });

  s.addText("算法整体进度 45% → 87%，16 周无停滞。", {
    x: 0.62, y: 6.78, w: 12.0, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: MUTED,
  });
  pageNum(s, 4);
  s.addNotes("这一页是全篇的骨架。我想强调的不是某个具体分数，而是这六件事在我来之前都不存在，现在都跑通了，而且是彼此打通的一条链路。");
}

// =====================================================================
// S5 — 行为识别
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.1  CORE MODEL", "核心成果一：行为识别从 0 到可用");

  s.addText("四类行为识别成型，并建立以新个体检验泛化的方法", {
    x: 0.62, y: 1.42, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: MUTED,
  });

  // company-set milestone and its outcome
  card(s, { x: 0.6, y: 1.8, w: 12.13, h: 0.6, fill: PURPLE_DK });
  s.addText("8.31 目标", {
    x: 0.88, y: 1.94, w: 1.5, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE, valign: "middle",
  });
  s.addText("抓挠识别 F1 ≥ 85%", {
    x: 2.5, y: 1.94, w: 3.2, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, color: WHITE, valign: "middle",
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.86, y: 1.94, w: 0.86, h: 0.32, rectRadius: 0.07, fill: { color: ORANGE },
  });
  s.addText("已达成", {
    x: 5.86, y: 1.94, w: 0.86, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: PURPLE_DK, align: "center", valign: "middle",
  });
  s.addText("EVT → DVT 阶段评审节点，算法侧均按期完成对接", {
    x: 7.1, y: 1.94, w: 5.4, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: PURPLE_LT, valign: "middle",
  });

  const steps = [
    ["01", "初版模型", "影棚 3 只犬数据训练", "训练个体准确率约 85%", PURPLE_MD],
    ["02", "稳定版 v2", "补入新个体数据重训", "4 只犬准确率约 85%，误报收敛", PURPLE],
    ["03", "泛化验证方法", "每引入新个体即检验一次", "用新样本测通用性与鲁棒性", ORANGE_DK],
  ];

  steps.forEach((st, i) => {
    const x = 0.6 + i * 4.11;
    card(s, { x, y: 2.58, w: 3.91, h: 2.28, fill: i === 2 ? PURPLE_DK : PURPLE_XLT });
    const dark = i === 2;
    badge(s, x + 0.26, 2.8, 0.46, st[0], st[4], WHITE, 12);
    s.addText(st[1], {
      x: x + 0.86, y: 2.84, w: 2.9, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, bold: true, color: dark ? WHITE : PURPLE_DK,
    });
    s.addText(st[2], {
      x: x + 0.3, y: 3.4, w: 3.32, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: dark ? PURPLE_LT : MUTED,
    });
    s.addText(st[3], {
      x: x + 0.3, y: 3.82, w: 3.32, h: 0.84, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 12, bold: true, color: dark ? ORANGE : PURPLE_DK,
      lineSpacingMultiple: 1.25,
    });
  });

  card(s, { x: 0.6, y: 5.08, w: 12.13, h: 1.06, fill: PURPLE_LT });
  s.addText("下一步", {
    x: 0.88, y: 5.32, w: 1.5, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: PURPLE_DK,
  });
  s.addText("狗场先加 6 只（1-2 周），跑完再逐步扩大，每一轮都用新个体重新检验。", {
    x: 2.4, y: 5.28, w: 10.1, h: 0.66, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 13, color: INK,
  });

  s.addText("以上为算法侧数据集测试结果。", {
    x: 0.62, y: 6.36, w: 12.0, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10, color: MUTED, italic: true,
  });
  pageNum(s, 5);
  s.addNotes("先交代目标：8月31日要求抓挠识别 F1 不低于 85%，已达成。然后讲泛化。三只狗训出来的模型在训练个体上约 85%，换一只没见过的马尔济斯会掉到 75% 左右且误报变多；把它的数据补进去重训，四只又回到 85%。所以我建立的做法是：每引入新个体就做一次泛化检验。狗场先加 6 只。数字是算法侧数据集测的，还没经测试同事验收 —— 这点在待提升页会讲。");
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
    ["抓挠识别", "稳定版 v2，4 只犬验证可用", PURPLE, "可用"],
    ["活动检测", "基础可用 80%+", PURPLE_MD, "可用"],
    ["睡觉 / 休息", "基础可用 80%+", PURPLE_MD, "可用"],
    ["未佩戴检测", "已训练为独立类别", PURPLE_MD, "已训练"],
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
  s.addText("皮肤健康评估：从 0 到跑通闭环", {
    x: 6.94, y: 1.62, w: 6, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: PURPLE_DK,
  });

  card(s, { x: 6.92, y: 2.06, w: 5.81, h: 1.62, fill: PURPLE_XLT });
  s.addText("规则统计版（SBS）", {
    x: 7.18, y: 2.22, w: 3.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
  });
  s.addText([
    { text: "前后端 Web 系统完成，兽医问答可用", options: { bullet: true, breakLine: true } },
    { text: "联通每日抓挠数据，驱动规则计算", options: { bullet: true, breakLine: true } },
    { text: "可解释性强，供产品与兽医验证", options: { bullet: true } },
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
    { text: "模型 A：特征 → C0 / C1 / C2，每日自动运行", options: { bullet: true, breakLine: true } },
    { text: "模型 B：含问答特征 → S0 / S1 / S2", options: { bullet: true, breakLine: true } },
    { text: "方案可行性已验证，待真实数据二次训练", options: { bullet: true, breakLine: true } },
    { text: "个体基线已落地，内部每日运行", options: { bullet: true } },
  ], {
    x: 7.2, y: 4.34, w: 5.3, h: 1.28, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 10.5, color: WHITE, paraSpaceAfter: 4,
  });

  // bottom insight
  card(s, { x: 0.6, y: 5.86, w: 12.13, h: 0.86, fill: PURPLE_LT });
  s.addText("当前状态", {
    x: 0.88, y: 6.06, w: 2.0, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: PURPLE_DK,
  });
  s.addText("链路跑通，每日自动出结果，方案路径已验证可行。", {
    x: 2.9, y: 6.02, w: 9.6, h: 0.4, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, color: INK, valign: "middle",
  });
  pageNum(s, 6);
  s.addNotes("行为识别四分类已成型；皮肤评估采取规则版与ML版双线并行，在数据不充足时降低单线失败风险。");
}

// =====================================================================
// S7 — 四个自建系统
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.3  SYSTEMS", "核心成果三：四套系统自建，全链路打通");

  s.addText("采集、平台、训练、服务四套系统彼此打通 —— 算法能自己转起来", {
    x: 0.62, y: 1.44, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  const sys = [
    ["采集与同步", "witmotion_imu", "数据入口", [
      "6 机位 × 12 设备无人值守采集",
      "单一时钟 + 事件驱动取帧，同步由构造保证",
      "自研项圈协议解析，兼容 WitMotion 双帧格式",
    ], true],
    ["标注与训练平台", "label_infra / smart-label", "数据加工", [
      "任务分发、认领、审核、复核全流程",
      "自研平台替代 Label Studio，历史数据完整迁移",
      "AI 预标注 → 审核 → 训练 → 模型热切换",
    ], false],
    ["模型训练与皮肤评估", "imu_train", "模型产出", [
      "多种模型架构横向对比，评测流程标准化",
      "事件级指标自研，可发现片段被切碎这类错误",
      "皮肤评估三条技术路线并行",
    ], false],
    ["在线推理与评估", "algo_service", "线上落地", [
      "15 秒增量推理 + 每日批量评估 + 基线更新",
      "TDengine + MySQL 分表，按用户时区切天",
      "核心逻辑有单元测试覆盖",
      "已交付后端，全链路跑通",
    ], false],
  ];

  sys.forEach((sy, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.6 + col * 6.19;
    const y = 1.86 + row * 2.4;
    const dark = sy[4];
    card(s, { x, y, w: 5.94, h: 2.24, fill: dark ? PURPLE_DK : PURPLE_XLT });

    s.addText(sy[0], {
      x: x + 0.3, y: y + 0.2, w: 3.3, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 13.5, bold: true, color: dark ? WHITE : PURPLE_DK,
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 4.5, y: y + 0.2, w: 1.18, h: 0.3, rectRadius: 0.07,
      fill: { color: dark ? ORANGE : PURPLE_MD },
    });
    s.addText(sy[2], {
      x: x + 4.5, y: y + 0.2, w: 1.18, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10, bold: true,
      color: dark ? PURPLE_DK : WHITE, align: "center", valign: "middle",
    });
    s.addText(sy[1], {
      x: x + 0.31, y: y + 0.52, w: 4.2, h: 0.24, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 9.5, color: dark ? PURPLE_LT : MUTED,
    });
    s.addText(sy[3].map((t, k) => ({
      text: t, options: { bullet: true, breakLine: k !== sy[3].length - 1 },
    })), {
      x: x + 0.3, y: y + 0.84, w: 5.36, h: 1.3, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10, color: dark ? WHITE : INK,
      paraSpaceAfter: 5, lineSpacingMultiple: 1.12,
    });
  });

  s.addText("四套系统均配有设计文档", {
    x: 0.62, y: 6.66, w: 12.0, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: MUTED,
  });
  pageNum(s, 7);
  s.addNotes("这四套系统是试用期最实在的产出。采集解决数据从哪来，平台解决数据怎么变成标注，训练解决模型怎么出来，服务解决模型怎么用起来。四套打通，算法才能自己转起来。");
}

// =====================================================================
// S7b — 算法平台全链路
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.4  PLATFORM", "核心成果四：算法平台全链路自建");

  s.addText("不只是标注工具，而是支撑算法持续迭代的基础设施", {
    x: 0.62, y: 1.44, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  // pipeline
  const stages = ["数据采集", "AI 预标注", "分发与审核", "模型训练", "模型测试", "多算法服务"];
  stages.forEach((st, i) => {
    const x = 0.6 + i * 2.03;
    const hot = i === 1 || i === 2;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 1.88, w: 1.88, h: 0.56, rectRadius: 0.09,
      fill: { color: hot ? PURPLE : PURPLE_LT },
    });
    s.addText(st, {
      x, y: 1.88, w: 1.88, h: 0.56, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11.5, bold: true,
      color: hot ? WHITE : PURPLE_DK, align: "center", valign: "middle",
    });
    if (i < stages.length - 1) {
      s.addText("→", {
        x: x + 1.85, y: 1.88, w: 0.21, h: 0.56, isTextBox: true, margin: 0,
        fontFace: F, fontSize: 11, bold: true, color: PURPLE_MD, align: "center", valign: "middle",
      });
    }
  });

  // screenshot of the review workspace
  const shot = "media/review.png";
  const sx = 0.6, sy = 2.68, sw = 6.6, sh = 3.13;
  if (fs.existsSync(shot)) {
    s.addImage({ path: shot, x: sx, y: sy, w: sw, h: sh });
  } else {
    card(s, { x: sx, y: sy, w: sw, h: sh, fill: PURPLE_DK });
  }
  s.addText("标注审核工作台：三路摄像头同步 + IMU 波形对齐", {
    x: sx, y: sy + sh + 0.08, w: sw, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 9.5, color: MUTED,
  });

  // mechanism callouts
  const mech = [
    ["AI 预标注，人工只做确认", "模型先跑候选片段，人工从「从头标」变为「审核修正」"],
    ["疑似片段主动召回", "低置信与频谱异常片段自动推送复核，捕捉模型最不确定的样本"],
    ["多人协同审核流", "任务认领、整份通过、分级权限与数据隔离"],
  ];
  mech.forEach((m, i) => {
    const y = sy + i * 1.09;
    card(s, { x: 7.44, y, w: 5.29, h: 0.95, fill: PURPLE_XLT });
    badge(s, 7.66, y + 0.26, 0.42, String(i + 1), PURPLE_MD, WHITE, 11);
    s.addText(m[0], {
      x: 8.2, y: y + 0.12, w: 4.4, h: 0.28, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12, bold: true, color: PURPLE_DK,
    });
    s.addText(m[1], {
      x: 8.2, y: y + 0.42, w: 4.36, h: 0.46, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 9.5, color: INK, lineSpacingMultiple: 1.2,
    });
  });

  card(s, { x: 0.6, y: 6.3, w: 12.13, h: 0.72, fill: PURPLE_DK });
  s.addText("平台承载的算法", {
    x: 0.88, y: 6.52, w: 2.1, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE,
  });
  s.addText("行为识别（抓挠 / 活动 / 睡觉 / 未佩戴）　·　皮肤健康评估（规则统计版 + ML 版）　·　口腔牙结石识别", {
    x: 3.0, y: 6.5, w: 9.5, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: WHITE, valign: "middle",
  });

  pageNum(s, 8);
  s.addNotes("平台是我认为试用期最有复用价值的产出。它把采集、标注、训练、测试、服务串成一条链路，三套算法都跑在上面。最关键的两个机制是 AI 预标注和疑似片段主动召回 —— 前者把人从重复劳动里解放出来，后者让人力集中在模型最不确定的地方。");
}

// =====================================================================
// S9b — 视频演示：抓挠检出与人工确认（视频内嵌，放映时点击播放）
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.5  LIVE DEMO", "算法效果演示");

  const vw = 7.35, vh = 7.35 / 2.11;
  const vx = 0.6, vy = 1.8;
  s.addShape(pres.ShapeType.roundRect, {
    x: vx - 0.06, y: vy - 0.06, w: vw + 0.12, h: vh + 0.12, rectRadius: 0.08,
    fill: { color: PURPLE_DK },
  });
  if (fs.existsSync("media/scratch.mp4")) {
    s.addMedia({ type: "video", path: "media/scratch.mp4", x: vx, y: vy, w: vw, h: vh });
  } else {
    fitImage(s, "media/scratch.png", vx, vy, vw, vh);
  }
  s.addText("放映时点击画面播放", {
    x: vx, y: vy + vh + 0.16, w: vw, h: 0.26, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 9.5, color: MUTED,
  });

  s.addText("抓挠检出与人工确认  —  请看四处", {
    x: 8.28, y: 1.8, w: 4.5, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: PURPLE_DK,
  });

  const pts = [
    ["三路画面同步", "三机位与 IMU 共用一条时间轴"],
    ["AI 预标片段", "模型标出 11.7 秒抓挠，置信度 88%"],
    ["疑似片段召回", "低置信片段单独挑出，等人工裁决"],
    ["审核动作", "认领修改 / 整份通过，支撑多人协同"],
  ];
  pts.forEach((p, i) => {
    const y = 2.2 + i * 0.79;
    card(s, { x: 8.26, y, w: 4.47, h: 0.7, fill: i === 2 ? PURPLE_LT : PURPLE_XLT });
    badge(s, 8.46, y + 0.16, 0.38, String(i + 1), i === 2 ? PURPLE : PURPLE_MD, WHITE, 11);
    s.addText(p[0], {
      x: 8.96, y: y + 0.05, w: 3.6, h: 0.26, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11.5, bold: true, color: PURPLE_DK,
    });
    s.addText(p[1], {
      x: 8.96, y: y + 0.33, w: 3.6, h: 0.3, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 9.5, color: INK,
    });
  });

  const more = [
    ["media/platform.png", "皮肤评估每日跟踪", "每犬每日自动出有效佩戴、抓挠统计与评分"],
    ["media/tartar.png", "口腔牙齿检测", "图片 / 视频 / 实时三种模式，正常与异常识别"],
  ];
  more.forEach((m, i) => {
    const x = 0.6 + i * 6.19;
    card(s, { x, y: 5.48, w: 5.94, h: 1.42, fill: PURPLE_XLT });
    const tx = x + 0.18, ty = 5.6, tw = 2.06, th = 1.18;
    s.addShape(pres.ShapeType.roundRect, {
      x: tx, y: ty, w: tw, h: th, rectRadius: 0.06, fill: { color: PURPLE_DK },
    });
    if (fs.existsSync(m[0])) fitImage(s, m[0], tx, ty, tw, th);
    s.addText(m[1], {
      x: x + 2.42, y: 5.76, w: 3.36, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
    });
    s.addText(m[2], {
      x: x + 2.42, y: 6.1, w: 3.36, h: 0.5, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10, color: INK, lineSpacingMultiple: 1.2,
    });
  });

  pageNum(s, 9);
  s.addNotes("现场点开视频播放，另外两项在浏览器里现场演示。重点讲两个机制：AI 预标注让人从「从头标」变成「审核修正」；疑似片段召回把模型最不确定的样本主动推给人工，这两条是标注效率和数据质量的关键。播放控制在 1-2 分钟。");
}

// =====================================================================
// S8 — 数据资源从 0 到 1
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.6  DATA ASSETS", "核心成果五：数据资源体系从 0 到 1");

  s.addText("把「无数据可用」变成「稳定可持续的采集供给」", {
    x: 0.62, y: 1.44, w: 11.6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  // timeline
  const tl = [
    ["5 月", "模拟数据起步", "自研 IMU 模拟生成器，跑通端到端流程", PURPLE_MD],
    ["6 月", "首批真实数据", "影棚落地采集，攻克时间戳对齐", PURPLE],
    ["7 月", "采集标准化", "方案定型，犬只扩至 3 只", PURPLE],
    ["8 月", "规模化铺路", "输出采集方案供商务洽谈，转向龙岗", ORANGE_DK],
    ["9 月", "狗场落地", "龙岗考察通过，6 间 6 只可采集", ORANGE_DK],
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
  s.addText("影棚 0 → 4 只  ·  龙岗狗场 6 只已谈妥", {
    x: 0.9, y: 5.44, w: 5.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: WHITE,
  });
  s.addText("长远按 30 → 50 → 70 → 100 → 150 只分阶段扩大", {
    x: 0.9, y: 5.8, w: 5.4, h: 0.5, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10, color: PURPLE_LT, lineSpacingMultiple: 1.2,
  });

  card(s, { x: 6.79, y: 4.92, w: 5.94, h: 1.62, fill: PURPLE_LT });
  s.addText("主动推动的关键判断", {
    x: 7.07, y: 5.1, w: 3.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, bold: true, color: PURPLE_DK,
  });
  s.addText("沙井报价 1000 元/犬/天，判断不经济果断放弃；主动输出采集方案供商务对接，转向龙岗并推动落地。", {
    x: 7.09, y: 5.44, w: 5.4, h: 0.9, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: INK, lineSpacingMultiple: 1.25,
  });
  pageNum(s, 10);
  s.addNotes("数据资源不能被动等待。从模拟数据起步，到影棚落地，再到狗场合作谈成，这条线是我主动推动的。");
}

// =====================================================================
// S9 — 跨团队协作
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.7  COLLABORATION", "核心成果六：跨团队协作与主动推动");

  const collab = [
    ["后端", "数据库架构调整；交付算法服务工程，后端完成外网部署，全链路跑通", PURPLE],
    ["产品", "PRD 对接、明确算法边界、皮肤评估方案共同推敲", PURPLE_MD],
    ["硬件厂商", "定位并反馈丢数据、蓝牙断传问题，推动修复与交付", ORANGE_DK],
    ["兽医 / 产品", "示范标注工具，建立协同分工，推动兽医参与评估", PURPLE_MD],
    ["测试 / 项目", "预留测试接口；推动算法在 ONES 单独立项", PURPLE],
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
  s.addText("主动响应与建议", {
    x: 9.02, y: 1.9, w: 3.4, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: ORANGE,
  });
  const props = [
    ["1 天完成口腔识别可行性验证", "产品提出方向后当天做出识别原型，含图片、视频、实时三种模式，验证可行性与落点效果", true],
    ["质疑临床抓挠阈值", "实测一晚抓挠 10-20 次但皮肤无异常，提出需重新调研基线依据", false],
    ["建议竞品对标", "建议采购竞品实测，同步佩戴对比识别效果", false],
    ["推动规范立项", "算法仅一人也应立项，明确阶段目标与验收标准", false],
  ];
  props.forEach((p, i) => {
    const y = 2.32 + i * 0.94;
    s.addText(p[0], {
      x: 9.02, y, w: 3.5, h: 0.26, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true, color: p[2] ? ORANGE : WHITE,
    });
    s.addText(p[1], {
      x: 9.04, y: y + 0.28, w: 3.46, h: 0.62, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 9, color: PURPLE_LT, lineSpacingMultiple: 1.2,
    });
  });
  pageNum(s, 11);
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
    ["16 篇", "连续周报", "完成项、进度、复盘、风险、计划，节奏无中断", ORANGE_DK],
    ["每项有结论", "汇报习惯", "每项都给出明确结论，而非罗列过程", PURPLE],
    ["问题必闭环", "工作方式", "阻塞点定位到根因再解决", PURPLE],
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
  s.addText("模型初版 6 周、皮肤评估 Web 系统 3 周、数据平台多人协同可用 4 周 —— 每个模块都有明确的起点与交付。", {
    x: 2.4, y: 5.64, w: 10.1, h: 0.7, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: WHITE, valign: "middle", lineSpacingMultiple: 1.25,
  });
  pageNum(s, 12);
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
    ["数据质量 > 数据量", "训练集导出时把「存疑」时段从重叠标注中挖掉 —— 不挖的话，一段大概率是抓挠的区间会被当作负样本喂给模型，比没有这条数据更糟。", ORANGE_DK],
    ["用模型输出反推数据缺口", "把模型低置信片段与频谱线索单独建表推送人工复核：确认的是漏检、否决的是误检，两者都是重训练最有价值的样本。", PURPLE],
    ["先分清是算法问题还是定义问题", "活动与睡觉混淆的根因是业务边界模糊，不是模型能力不足。算法无法自行划定业务语义，须先对齐定义再训练。", PURPLE_MD],
    ["发现自己的错误要主动报出来", "自查发现合成数据里特征与标签共用了同一个生成表达式，虚高的宏 F1 0.957 修正后降到 0.886，写进文档公开。宁可报低，不留隐患。", PURPLE],
    ["及时验证、及时放弃", "RTSP 方案验证发现网络抖动导致延迟不可预测后果断放弃，收敛至 USB 本地录制；实测两种重采样算法差异达信号标准差的 6-8%，也是靠实测而非推理定论。", ORANGE_DK],
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
  pageNum(s, 13);
  s.addNotes("这五条是我试用期最重要的方法论沉淀，也是我认为可以持续复用到后续工作中的东西。");
}

// =====================================================================
// S12 — 待提升点与改进计划
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "05  GAPS & PLAN", "待提升点与改进计划");

  s.addText("每一项都有明确的改进路径", {
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
    ["个体覆盖不足，泛化未验证", "训练仅覆盖 4 只犬，新个体上准确率下降、误报增多", "狗场先加 6 只，逐步扩大，每轮用新个体重新检验"],
    ["尚未经测试与线上验证", "当前结果均为算法侧数据集测试", "推动测试同事介入，进入线上验证流程"],
    ["皮肤评估阈值待校准", "抓挠多但皮肤无病变时仍会误触发", "推动兽医标注，用临床数据校准映射关系"],
    ["松动检测尚未启动", "优先级让位，判定边界也需与产品兽医对齐", "对齐标准后启动，补齐有效佩戴时间统计"],
    ["业务理解仍需加深", "对兽医临床判断标准仍依赖他人输入", "系统补充宠物皮肤健康领域知识"],
  ];

  gaps.forEach((g, i) => {
    const y = 2.24 + i * 0.92;
    card(s, { x: 0.6, y, w: 12.13, h: 0.82, fill: i % 2 === 0 ? PURPLE_XLT : WHITE, line: i % 2 === 0 ? null : PURPLE_LT });
    badge(s, 0.84, y + 0.22, 0.38, String(i + 1), PURPLE_MD, WHITE, 11);
    s.addText(g[0], {
      x: 1.34, y: y + 0.24, w: 2.7, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12, bold: true, color: PURPLE_DK, valign: "middle",
    });
    s.addText(g[1], {
      x: 4.3, y: y + 0.24, w: 4.1, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: MUTED, valign: "middle",
    });
    s.addText(g[2], {
      x: 8.62, y: y + 0.24, w: 4.0, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: INK, valign: "middle",
    });
  });
  pageNum(s, 14);
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
      "推动龙岗狗场正式采集落地，先补 6 只犬并完成一轮泛化检验",
      "数据平台上线 AI 自动标注，加速人工审核效率",
      "个体基线持续调优并推上线，解决 C2 误触发",
      "与产品、兽医对齐活动/睡觉业务边界与松动判定标准",
    ], PURPLE],
    ["中期", "3-6 个月", [
      "个体覆盖扩大后，模型在全新犬只上表现稳定，具备推向用户的条件",
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
  pageNum(s, 15);
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
  s.addText([
    { text: "Toky", options: { fontSize: 13, bold: true, color: WHITE } },
    { text: " 朱业拓", options: { fontSize: 10.5, color: PURPLE_LT } },
    { text: "　　用户增长研发部 · 算法工程师　　2026.10.08", options: { fontSize: 11.5, color: PURPLE_LT } },
  ], {
    x: 0.87, y: 4.4, w: 9.0, h: 0.34, isTextBox: true, margin: 0, fontFace: F,
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

pres.writeFile({ fileName: "转正述职报告-Toky.pptx" }).then(f => console.log("written:", f));
