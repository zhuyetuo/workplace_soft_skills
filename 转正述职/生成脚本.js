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
function fitImage(slide, imgPath, px, py, pw, ph, hyperlink) {
  const dim = require("image-size").imageSize(fs.readFileSync(imgPath));
  const ar = dim.width / dim.height;
  let w = pw, h = pw / ar;
  if (h > ph) { h = ph; w = ph * ar; }
  const o = { path: imgPath, x: px + (pw - w) / 2, y: py + (ph - h) / 2, w, h };
  if (hyperlink) o.hyperlink = hyperlink;
  slide.addImage(o);
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
    ["负责方向", "狗的行为识别\n皮肤健康评估\n算法工程与数据平台"],
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
  s.addText("把项圈采到的数据，变成主人能看懂、也能放心的健康结论", {
    x: 5.44, y: 2.1, w: 7.3, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: MUTED,
  });

  const roles = [
    ["数据", "从零建立采集与标注流程，数据质量决定算法上限", ORANGE],
    ["模型", "教会算法识别狗的行为，并据此判断健康状况", PURPLE_MD],
    ["工程", "把算法做成产品能直接调用的服务，能上线才算数", PURPLE],
    ["协作", "对接产品、后端、硬件、兽医，推动资源到位", ORANGE_DK],
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

  s.addText("入职时这六件事一件都没有，现在都跑通了", {
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
    ["看懂狗在干什么", "没有模型，没有数据", "能分辨抓挠、活动、睡觉、没戴项圈"],
    ["判断皮肤健不健康", "没有方案，也没有判断标准", "每天自动给出每只狗的健康评分"],
    ["把数据采回来", "没有设备，没有场地", "6 个摄像头 + 12 个项圈，无人值守自动采集"],
    ["把数据整理好", "没有工具，没有流程", "自建标注平台，AI 先标、多人协同核对"],
    ["让产品用起来", "无", "已交付后端上线，App 里能看到算法结果"],
    ["可采集的狗", "影棚 2 只，但没有采集流程", "影棚 4 只，狗场 6 只待进场"],
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
  pageTitle(s, "02.1  CORE MODEL", "成果一：让项圈看懂狗在干什么");

  s.addText("能分辨抓挠、活动、睡觉、没戴项圈，并建立了一套「换只狗还准不准」的检验方法", {
    x: 0.62, y: 1.42, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, color: MUTED,
  });

  // company-set milestone and its outcome
  card(s, { x: 0.6, y: 1.8, w: 12.13, h: 0.6, fill: PURPLE_DK });
  s.addText("8.31 目标", {
    x: 0.88, y: 1.94, w: 1.5, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE, valign: "middle",
  });
  s.addText("抓挠识别准确率 ≥ 85%", {
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
    ["01", "第一版模型", "用影棚 3 只狗的数据训练", "这 3 只狗准确率约 85%", PURPLE_MD],
    ["02", "稳定版 v2", "把第 4 只狗的数据也补进去重训", "4 只狗准确率都到 85%，误报变少", PURPLE],
    ["03", "怎么验证靠不靠谱", "来一只就单独测，来一批就整批测", "换没见过的狗，看还准不准", ORANGE_DK],
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
  s.addText("狗场先加 6 只（1-2 周）做一次批量验证，之后按批次逐步扩大；单只新增就单独测，看具体情况定。", {
    x: 2.4, y: 5.28, w: 10.1, h: 0.66, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 13, color: INK,
  });

  s.addText("以上为算法自测结果，尚未经测试同事验收。", {
    x: 0.62, y: 6.36, w: 12.0, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10, color: MUTED, italic: true,
  });
  pageNum(s, 5);
  s.addNotes("先交代目标：8月31日要求抓挠识别准确率不低于 85%，已达成。然后讲验证方法。三只狗训出来的模型在训练个体上约 85%，换一只没见过的马尔济斯会掉到 75% 左右且误报变多；把它的数据补进去重训，四只又回到 85%。所以我的做法是分两条路：单只新增就单独精测，一次来一批就整批测，看具体情况定。狗场先加 6 只做批量验证。数字是算法侧数据集测的，还没经测试同事验收 —— 这点在待提升页会讲。");
}

// =====================================================================
// S6 — 算法能力矩阵
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.2  CAPABILITY", "成果二：从「抓了多少次」到「皮肤健不健康」");

  // left: 4-class behavior
  s.addText("能识别的四种状态", {
    x: 0.62, y: 1.62, w: 6, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: PURPLE_DK,
  });
  const beh = [
    ["抓挠", "稳定版 v2，4 只狗验证可用", PURPLE, "可用"],
    ["活动", "基础可用", PURPLE_MD, "可用"],
    ["睡觉 / 休息", "基础可用", PURPLE_MD, "可用"],
    ["没戴项圈", "基础可用，线上待接入", PURPLE_MD, "可用"],
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
  s.addText("两种打分方式，都已跑通", {
    x: 6.94, y: 1.62, w: 6, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: PURPLE_DK,
  });

  card(s, { x: 6.92, y: 2.06, w: 5.81, h: 1.62, fill: PURPLE_XLT });
  s.addText("规则版：按兽医定的规则打分", {
    x: 7.18, y: 2.22, w: 3.4, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
  });
  s.addText([
    { text: "网页系统已完成，兽医可直接填写问诊记录", options: { bullet: true, breakLine: true } },
    { text: "自动接入每天的抓挠统计，算出健康等级", options: { bullet: true, breakLine: true } },
    { text: "每一分怎么来的都能说清楚，方便兽医核对", options: { bullet: true } },
  ], {
    x: 7.2, y: 2.56, w: 5.3, h: 1.0, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 10.5, color: INK, paraSpaceAfter: 4,
  });

  card(s, { x: 6.92, y: 3.84, w: 5.81, h: 1.86, fill: PURPLE_DK });
  s.addText("学习版：让模型自己从数据里学", {
    x: 7.18, y: 4.0, w: 3.6, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: ORANGE,
  });
  s.addText([
    { text: "第一步：由行为数据判断皮肤等级，每天自动跑", options: { bullet: true, breakLine: true } },
    { text: "第二步：结合主人的问答，给出综合评估", options: { bullet: true, breakLine: true } },
    { text: "方案已验证可行，等真实病例数据再训练一轮", options: { bullet: true, breakLine: true } },
    { text: "每只狗有自己的「正常水平」作参照，已在内部每天运行", options: { bullet: true } },
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
  s.addText("整条路已经跑通，每天自动出结果，方向验证可行。", {
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
  pageTitle(s, "02.3  SYSTEMS", "成果三：把整条流水线搭起来");

  s.addText("采数据、理数据、练模型、上线用，四个环节都自己搭，而且是打通的", {
    x: 0.62, y: 1.44, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  const sys = [
    ["① 把数据采回来", "摄像头 + 项圈同步录制", "数据入口", [
      "6 个摄像头 + 12 个项圈，开机自动录、每天自动归档",
      "画面和项圈数据严丝合缝对上，这是能标注的前提",
      "打通了自研项圈的数据读取，也兼容外购设备",
    ], true],
    ["② 把数据整理好", "自建标注平台", "数据加工", [
      "任务分发、领取、审核、复核，多人可同时干活",
      "自己搭的平台替换了原来的开源工具，老数据全部迁过来",
      "AI 先标 → 人工核对 → 训练模型 → 直接换上新模型",
    ], false],
    ["③ 把模型练出来", "多种方案横向比较", "模型产出", [
      "多种模型方案横向比较，挑效果最好的用",
      "自己写了评测方法，能发现「一次抓挠被切成好几段」这类错",
      "皮肤评分同时试了三条路线，避免押注单一方案",
    ], false],
    ["④ 让产品用起来", "每天自动出结果", "线上落地", [
      "每 15 秒处理一批新数据，每天凌晨自动出评估结果",
      "按用户所在时区切分「一天」，避免时差算错",
      "核心逻辑都有自动化测试兜底",
      "已交付后端并完成外网部署，App 里能看到结果",
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

  s.addText("四套系统均配有完整设计文档", {
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
  pageTitle(s, "02.4  PLATFORM", "成果四：AI 先干粗活，人只做把关");

  s.addText("早期没有可用模型，只能纯人工标注 —— 一只狗一天录 24 段、每段 1 小时，靠人从头看到尾，费时又费人", {
    x: 0.62, y: 1.44, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  // pipeline
  const stages = ["采集数据", "AI 先标", "人工核对", "训练模型", "测试效果", "上线使用"];
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
  s.addText("标注核对界面：三路画面与项圈数据同步回看", {
    x: sx, y: sy + sh + 0.08, w: sw, h: 0.28, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 9.5, color: MUTED,
  });

  // mechanism callouts
  const mech = [
    ["AI 先标，人只做核对", "模型先给出候选片段，人从「一帧帧从头标」变成「核对和修正」"],
    ["可疑的也一并挑出来", "可能认错的、可能漏掉的都单独列出，人重点看这些，不用整段翻"],
    ["多人可以一起干", "任务分发领取、逐条或整份通过，不同角色只看该看的数据"],
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
  s.addText("平台上同时跑着：行为识别　·　皮肤健康评估　·　口腔牙齿检测", {
    x: 3.0, y: 6.5, w: 9.5, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: WHITE, valign: "middle",
  });

  pageNum(s, 8);
  s.addNotes("这里先说清楚起点：一开始没有可用模型，标注全靠人工。一只狗一天录 24 段、每段一小时，人要从头看到尾，根本看不完。现在 AI 先把候选片段标出来，还会把可能认错、可能漏掉的单独挑出来，人只看这些重点。平台是我认为试用期最有复用价值的产出。它把采集、标注、训练、测试、服务串成一条链路，三套算法都跑在上面。最关键的两个机制是 AI 预标注和疑似片段主动召回 —— 前者把人从重复劳动里解放出来，后者让人力集中在模型最不确定的地方。");
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

  s.addText("抓挠检出与人工核对  —  请看四处", {
    x: 8.28, y: 1.8, w: 4.5, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: PURPLE_DK,
  });

  const pts = [
    ["三路画面同步", "三个摄像头与项圈数据共用一条时间轴"],
    ["AI 标出的片段", "模型标出 11.7 秒抓挠，把握 88%"],
    ["拿不准的片段", "AI 没把握的地方单独挑出，等人判断"],
    ["人工核对", "认领修改 / 整份通过，多人可同时进行"],
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
    ["media/platform.png", "皮肤评估每日跟踪", "每只狗每天自动出佩戴时长、抓挠统计与健康评分"],
    ["media/tartar.png", "口腔牙齿检测", "上传照片即可判断正常或异常"],
  ];
  more.forEach((m, i) => {
    const x = 0.6 + i * 6.19;
    card(s, { x, y: 5.48, w: 5.94, h: 1.42, fill: PURPLE_XLT });
    const tx = x + 0.18, ty = 5.6, tw = 2.06, th = 1.18;
    s.addShape(pres.ShapeType.roundRect, {
      x: tx, y: ty, w: tw, h: th, rectRadius: 0.06, fill: { color: PURPLE_DK },
    });
    if (fs.existsSync(m[0])) fitImage(s, m[0], tx, ty, tw, th, { slide: 17 + i, tooltip: "点击放大" });
    // transparent hit area so the whole thumbnail is clickable, not just the image
    s.addShape(pres.ShapeType.roundRect, {
      x: tx, y: ty, w: tw, h: th, rectRadius: 0.06,
      fill: { color: WHITE, transparency: 100 }, line: { type: "none" },
      hyperlink: { slide: 17 + i, tooltip: "点击放大" },
    });
    s.addText("点击放大", {
      x: x + 2.42, y: 6.58, w: 3.36, h: 0.24, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 9, color: PURPLE_MD,
    });
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
  pageTitle(s, "02.6  DATA ASSETS", "成果五：把数据采集从零跑起来");

  s.addText("从没有一条数据可用，到每天都有稳定的数据进来", {
    x: 0.62, y: 1.44, w: 11.6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  // timeline
  const tl = [
    ["5 月", "先用模拟数据", "真实数据还没有，自己造数据把流程跑通", PURPLE_MD],
    ["6 月", "拿到第一批真数据", "影棚开始采集，解决了画面与数据对不齐的老问题", PURPLE],
    ["7 月", "采集流程定型", "方案固定下来，狗增加到 3 只", PURPLE],
    ["8 月", "为规模化铺路", "输出采集需求方案供商务对接；实地考察沙井，判断成本过高", ORANGE_DK],
    ["9 月", "狗场落实", "商务谈成龙岗，我赴现场考察，确认 6 间犬舍 6 只狗可采", ORANGE_DK],
  ];

  const tlY = 1.88;
  tl.forEach((t, i) => {
    const x = 0.6 + i * 2.46;
    card(s, { x, y: tlY, w: 2.32, h: 2.34, fill: PURPLE_XLT });
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

  // bottom: the equipment struggle, and the one piece of initiative that was mine
  card(s, { x: 0.6, y: 4.44, w: 5.94, h: 2.08, fill: PURPLE_DK });
  s.addText("设备是自己一步步试出来的", {
    x: 0.88, y: 4.6, w: 5.4, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: ORANGE,
  });
  s.addText([
    { text: "自研项圈至今没有成品，只能用问题很多的老款设备", options: { bullet: true, breakLine: true } },
    { text: "自己选型、反复退换，并改用电脑端时间做基准，才把采集跑稳", options: { bullet: true } },
  ], {
    x: 0.9, y: 5.0, w: 5.4, h: 1.0, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 11, color: WHITE, paraSpaceAfter: 5, lineSpacingMultiple: 1.2,
  });

  card(s, { x: 6.79, y: 4.44, w: 5.94, h: 2.08, fill: PURPLE_LT });
  s.addText("我主动推动的一件事", {
    x: 7.07, y: 4.6, w: 5.4, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
  });
  s.addText("影棚两位同事各养了一只狗，我提议带来参与测试 —— 公司有补贴政策，几乎零成本把可采集犬只从 2 只增加到 4 只。", {
    x: 7.09, y: 5.0, w: 5.4, h: 0.86, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 11, color: INK, lineSpacingMultiple: 1.25,
  });
  s.addText("影棚 2 → 4 只　·　龙岗狗场 6 只已谈妥　·　长远按 30 → 50 → 70 → 100 → 150 只分阶段扩大", {
    x: 7.09, y: 6.02, w: 5.4, h: 0.4, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 9.5, color: PURPLE_DK, lineSpacingMultiple: 1.2,
  });
  pageNum(s, 10);
  s.addNotes("设备这块可以多讲两句：公司自研项圈因为加了 GPS 和 LTE 一直没出成品，我手上只有老款 TF 卡设备 —— 数据要第二天才看得到，电量撑不过 12 小时得频繁换，时间戳一天还会漂将近 9 秒，画面和数据对不齐就没法标注。后来我自己上电商平台挑设备，来回退换了好几次才找到能稳定跑满 24 小时的；真正解决漂移的办法是改用电脑端时间做基准，不用设备自己的时钟。摄像头也从 TP-Link、小米监控换到 USB 直连才对齐。分工上：狗场是商务去找、去谈的，我提的是数据需求、输出采集方案，并到现场考察确认可行。我自己主动推动的是影棚这条线 —— 发现两位同事各养了一只狗，提议带来参与测试，公司有补贴政策，几乎零成本把犬只从 2 只加到 4 只。");
}

// =====================================================================
// S9 — 跨团队协作
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.7  COLLABORATION", "成果六：把事情当自己的事推着走");

  const collab = [
    ["后端", "配合调整数据存储方式；交付算法服务，后端完成部署，全链路跑通", PURPLE],
    ["产品", "PRD 对接、明确算法边界、皮肤评估方案共同推敲", PURPLE_MD],
    ["硬件厂商", "查出并反馈设备丢数据、蓝牙断连的问题，推动修复与交付", ORANGE_DK],
    ["兽医 / 产品", "示范标注工具，建立协同分工，推动兽医参与评估", PURPLE_MD],
    ["测试 / 项目", "预留测试接口；推动算法在 ONES 单独立项", PURPLE],
  ];

  s.addText("算法岗只有我一人，很多事没人催也没人推 —— 能推到哪一步，取决于我主动做到哪一步", {
    x: 0.62, y: 1.42, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  collab.forEach((c, i) => {
    const y = 1.86 + i * 0.84;
    card(s, { x: 0.6, y, w: 7.9, h: 0.72, fill: PURPLE_XLT });
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
  card(s, { x: 8.76, y: 1.86, w: 3.97, h: 4.2, fill: PURPLE_DK });
  s.addText("我主动做的几件事", {
    x: 9.02, y: 2.02, w: 3.4, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: ORANGE,
  });
  const props = [
    ["1 天做出口腔识别原型", "产品提出这个方向后，当天做出可用原型，验证了这条路走不走得通", true],
    ["质疑现有判断标准", "实测一晚抓挠 10-20 次但皮肤没问题，提出标准需要重新调研", false],
    ["建议对标竞品", "建议买竞品实测，让狗同时戴上对比效果", false],
    ["推动规范立项", "算法只有一人也应正式立项，明确目标与验收标准", false],
  ];
  props.forEach((p, i) => {
    const y = 2.46 + i * 0.9;
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
  s.addNotes("这一页想说的是我怎么对待这个项目。算法岗只有我一人，没有人会来推着我做，所以很多事我是当成自己的事在推：设备不行就自己去找设备，判断标准存疑就提出来重新调研，需要别人配合就主动去谈。同样重要的是主动开口求助 —— 16 篇周报每篇都写了资源诉求，需要什么就明确提，不闷头硬扛。");
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
    ["16 篇", "连续周报", "每周写完成了什么、卡在哪、下周做什么，从没断过", ORANGE_DK],
    ["每项有结论", "汇报习惯", "每件事都给一句明确结论，不只罗列过程", PURPLE],
    ["问题必闭环", "工作方式", "卡住的地方一定找到根本原因再解决", PURPLE],
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
  s.addText("第一版模型 6 周、皮肤评估网页系统 3 周、标注平台可多人协同 4 周 —— 每块都有明确的开始和交付。", {
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
  pageTitle(s, "04  GROWTH", "个人成长：这几个月我想明白的五件事");

  const growth = [
    ["数据质量比数量更重要", "把标注时拿不准的片段先剔掉再拿去训练 —— 留着反而会教错模型，比没这条数据更糟。", ORANGE_DK],
    ["让模型告诉我该补什么数据", "把 AI 没把握的片段挑给人看：人说「是」的是它漏掉的，人说「不是」的是它认错的，两种都最值得再训练。", PURPLE],
    ["先分清是技术问题还是定义问题", "「活动」和「睡觉」老是混，根子在于这两件事本身就没定义清楚，不是模型不行。得先把标准对齐。", PURPLE_MD],
    ["自己发现的问题要主动说", "自查时发现造的测试数据有漏洞，会让分数虚高。主动把分数改低并写进文档 —— 宁可报低，不留隐患。", PURPLE],
    ["该放弃的要早点放弃", "网络传视频方案验证下来延迟不稳定，果断改回本地录制；有争议的地方靠实测下结论，不靠猜。", ORANGE_DK],
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
    ["用来训练的狗太少", "只用了 4 只狗训练，换只没见过的狗准确率会下降", "狗场先加 6 只，逐步扩大，每轮拿新狗重新检验"],
    ["还没经过测试和线上验证", "目前都是我自己测出来的结果", "推动测试同事介入，进入线上验证流程"],
    ["皮肤评分标准还要校准", "有的狗抓得多但皮肤其实没问题，会误报", "请兽医标注真实病例，用临床数据校准标准"],
    ["项圈松动检测还没做", "优先级往后排，判定标准也要先和产品、兽医对齐", "标准定了就启动，补齐有效佩戴时长统计"],
    ["业务理解还不够深", "兽医怎么判断皮肤问题，我还主要靠别人告诉我", "系统补充宠物皮肤健康方面的知识"],
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
      "推动龙岗狗场正式采集，先补 6 只狗并完成一轮检验",
      "标注平台上线 AI 自动标注，进一步提高核对效率",
      "皮肤评分标准持续校准并推上线，解决误报问题",
      "和产品、兽医对齐行为定义与松动判定标准",
    ], PURPLE],
    ["中期", "3-6 个月", [
      "狗的数量上来后，模型换新狗也稳定，具备推给用户的条件",
      "皮肤评估用真实病例再训练一轮，并上线验证",
      "补齐松动检测，形成完整的全天佩戴时长统计",
      "采集规模按 30 → 50 → 70 只分阶段扩大",
    ], PURPLE_MD],
    ["长期", "持续投入", [
      "把采集、标注、训练、上线这套流程固化成标准做法",
      "让算法从「能认出行为」走向「能给出可信的健康结论」",
      "建立竞品对标，摸清我们的能力边界和优势在哪",
      "团队变大后，把算法工作做到可协同、可交接",
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
  s.addText("这需要数据、算法、业务定义三件事同步推进，也需要产品、兽医、后端、硬件持续配合，我会主动担起推动这条链路的责任。", {
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

// =====================================================================
// A1 / A2 — 放大页（隐藏，从演示页点击缩略图进入）
// =====================================================================
[
  ["media/platform.png", "皮肤评估每日跟踪"],
  ["media/tartar.png", "口腔牙齿检测"],
].forEach(([img, title]) => {
  const s = pres.addSlide();
  s.background = { color: PURPLE_DK };
  s.addText(title, {
    x: 0.5, y: 0.26, w: 7, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 15, bold: true, color: WHITE,
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 11.7, y: 0.26, w: 1.13, h: 0.36, rectRadius: 0.08,
    fill: { color: PURPLE }, hyperlink: { slide: 9, tooltip: "返回演示页" },
  });
  s.addText("← 返回", {
    x: 11.7, y: 0.26, w: 1.13, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: WHITE,
    align: "center", valign: "middle", hyperlink: { slide: 9, tooltip: "返回演示页" },
  });
  if (fs.existsSync(img)) fitImage(s, img, 0.4, 0.82, 12.53, 6.4);
});

pres.writeFile({ fileName: "转正述职报告-Toky.pptx" }).then(f => console.log("written:", f));
