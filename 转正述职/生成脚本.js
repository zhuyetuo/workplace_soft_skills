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
    ["02", "试用期关键成果", "行为识别 · 皮肤评估 · 算法平台"],
    ["03", "持续产出的验证", "进度曲线与迭代节奏，证明可持续交付"],
    ["04", "个人成长与沉淀", "方法论沉淀与认知提升"],
    ["05", "待提升点与下一步", "短板认知与后续规划"],
  ];

  items.forEach((it, i) => {
    const y = 1.72 + i * 1.03;
    card(s, { x: 0.6, y, w: 12.13, h: 0.86, fill: PURPLE_XLT });
    badge(s, 0.86, y + 0.15, 0.56, it[0], PURPLE_MD, WHITE, 15);
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
    ["入职时间", "2026 年 4 月中旬"],
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
// S4 — 成果总览：三大块的推进
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02  成果展示", "试用期成果：三大块");

  s.addText("入职时没有数据、没有模型、没有平台。下面是三块各自的推进", {
    x: 0.62, y: 1.44, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  const LX = 0.6, LW = 1.42, TX = 2.24, CW = 2.06, STEP = 2.1;
  ["5 月", "6 月", "7 月", "8 月", "9 月"].forEach((m, i) => {
    s.addText(m, {
      x: TX + i * STEP, y: 1.84, w: CW, h: 0.28, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true, color: PURPLE_MD, align: "center",
    });
  });

  const LH = 1.3;
  const lanes = [
    ["行为识别", [
      [0, "跑通流程", ""], [1, "拿到真数据", ""], [2, "采集定型", ""],
      [3, "抓挠达标", "8.31 目标 ≥ 85%"], [4, "四类可用", ""],
    ]],
    ["皮肤评估", [
      [2, "方案定下来", ""], [3, "规则版上线", ""],
      [4, "每天自动出分", "AI 版方案已验证"],
    ]],
    ["算法平台", [
      [1, "标注工具就绪", ""], [2, "采集工程标准化", ""], [3, "自研平台上线", ""],
      [4, "AI 先标 + 服务上线", "人只做核对"],
    ]],
  ];

  lanes.forEach(([label, items], li) => {
    const ly = 2.16 + li * (LH + 0.14);
    card(s, { x: LX, y: ly, w: LW, h: LH, fill: PURPLE_DK });
    s.addText(label, {
      x: LX, y: ly, w: LW, h: LH, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, bold: true, color: WHITE, align: "center", valign: "middle",
    });
    s.addShape(pres.ShapeType.rect, {
      x: TX, y: ly + LH / 2, w: 4 * STEP + CW, h: 0.03, fill: { color: PURPLE_LT },
    });
    items.forEach(([col, title, note]) => {
      const x = TX + col * STEP;
      const hot = note !== "";
      card(s, { x, y: ly, w: CW, h: LH, fill: hot ? PURPLE : PURPLE_XLT });
      s.addText(title, {
        x: x + 0.1, y: hot ? ly + 0.2 : ly, w: CW - 0.2, h: hot ? 0.44 : LH,
        isTextBox: true, margin: 0,
        fontFace: F, fontSize: 13, bold: true,
        color: hot ? WHITE : PURPLE_DK, align: "center", valign: "middle",
      });
      if (hot) {
        s.addText(note, {
          x: x + 0.1, y: ly + 0.7, w: CW - 0.2, h: 0.36, isTextBox: true, margin: 0,
          fontFace: F, fontSize: 9.5, color: ORANGE, align: "center", valign: "middle",
        });
      }
    });
  });

  card(s, { x: 0.6, y: 6.66, w: 12.13, h: 0.5, fill: PURPLE_XLT });
  s.addText("另有一块延伸：图像识别（皮肤照片辅助判断、口腔牙齿检测）", {
    x: 0.9, y: 6.66, w: 6.6, h: 0.5, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 10.5, color: INK,
  });
  s.addText("算法整体进度 45% → 87%，16 周无停滞", {
    x: 7.8, y: 6.66, w: 4.7, h: 0.5, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 10.5, bold: true, color: PURPLE_DK, align: "right",
  });
  pageNum(s, 4);
  s.addNotes("成果汇总成三块：行为识别、皮肤评估、算法平台。行为识别 5 月用模拟数据跑通链路，8 月抓挠识别达到公司定的 85% 目标；皮肤评估 7 月把评分口径谈定，9 月每只狗每天自动出分；算法平台 6 月标注工具就绪，9 月上了 AI 预标注，算法服务也交付后端上线。深色的是每块目前最关键的节点。");
}

// =====================================================================
// S4b — 成果展示 · 行为识别
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.1  成果展示", "行为识别：让项圈看懂狗在干什么");

  // the four classes
  const kinds = ["抓挠", "活动", "睡觉", "没戴项圈"];
  kinds.forEach((k, i) => {
    const x = 0.6 + i * 3.06;
    card(s, { x, y: 1.62, w: 2.88, h: 1.06, fill: i === 0 ? PURPLE : PURPLE_XLT });
    s.addText(k, {
      x, y: 1.62, w: 2.88, h: 1.06, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 19, bold: true,
      color: i === 0 ? WHITE : PURPLE_DK, align: "center", valign: "middle",
    });
  });

  // the milestone
  card(s, { x: 0.6, y: 2.92, w: 12.13, h: 0.6, fill: PURPLE_DK });
  s.addText("8.31 目标", {
    x: 0.88, y: 3.06, w: 1.4, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: ORANGE, valign: "middle",
  });
  s.addText("抓挠识别准确率 ≥ 85%", {
    x: 2.4, y: 3.06, w: 3.2, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, color: WHITE, valign: "middle",
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 5.72, y: 3.06, w: 0.86, h: 0.32, rectRadius: 0.07, fill: { color: ORANGE },
  });
  s.addText("已达成", {
    x: 5.72, y: 3.06, w: 0.86, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: PURPLE_DK, align: "center", valign: "middle",
  });
  s.addText("EVT → DVT 阶段评审节点，算法侧均按期完成对接", {
    x: 6.9, y: 3.06, w: 5.6, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, color: PURPLE_LT, valign: "middle",
  });

  // how it is verified
  s.addText("模型能不能用，看它在没见过的狗身上掉多少", {
    x: 0.62, y: 3.76, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: PURPLE_DK,
  });

  const steps = [
    ["第一版", "3 只狗训练", "这 3 只约 85%"],
    ["稳定版 v2", "补入第 4 只重训", "4 只都到 85%，误报变少"],
    ["怎么验证", "来一只单独测，来一批整批测", "换没见过的狗，看还准不准"],
  ];
  steps.forEach((st, i) => {
    const x = 0.6 + i * 4.11;
    const hot = i === 2;
    card(s, { x, y: 4.2, w: 3.91, h: 1.9, fill: hot ? PURPLE : PURPLE_XLT });
    s.addText(st[0], {
      x: x + 0.28, y: 4.42, w: 3.3, h: 0.32, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14, bold: true, color: hot ? WHITE : PURPLE_DK,
    });
    s.addText(st[1], {
      x: x + 0.29, y: 4.8, w: 3.3, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: hot ? PURPLE_LT : MUTED,
    });
    s.addText(st[2], {
      x: x + 0.28, y: 5.2, w: 3.36, h: 0.66, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 12, bold: true, color: hot ? ORANGE : PURPLE_DK,
      lineSpacingMultiple: 1.25,
    });
  });

  s.addText("下一步：狗场先加 6 只（1-2 周）批量验证，之后按批次继续扩。以上为算法自测结果，尚未经测试同事验收。", {
    x: 0.62, y: 6.3, w: 12.0, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 10.5, color: MUTED,
  });
  pageNum(s, 5);
  s.addNotes("行为识别能分辨四类：抓挠、活动、睡觉、没戴项圈。公司定的 8.31 目标是抓挠准确率不低于 85%，已达成。验证方法是拿没见过的狗测 —— 3 只狗训出来的模型在这 3 只上约 85%，换一只没见过的会掉，把它补进去重训，4 只又回到 85%。所以模型现在是靠见过才准，这也是下一步要加狗的原因。");
}

// =====================================================================
// S6 — 皮肤评估
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.2  成果展示", "皮肤评估：两条路一起走");

  // the two tracks
  const tracks = [
    ["规则版", "PM 与兽医主导", ["每一分怎么来的都能说清楚", "兽医好核对、好调整", "上限不高，靠人定规则"], PURPLE_XLT, false],
    ["AI 版", "把业务规则和兽医临床经验变成特征，让模型学", ["上限更高，能自己找规律", "已用合成数据验证方案可行", "要真实病例数据才能真正训好"], PURPLE_DK, true],
  ];
  tracks.forEach((t, i) => {
    const x = 0.6 + i * 6.19;
    const dark = t[4];
    card(s, { x, y: 1.66, w: 5.94, h: 2.16, fill: t[3] });
    s.addText(t[0], {
      x: x + 0.3, y: 1.84, w: 2.4, h: 0.34, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 16, bold: true, color: dark ? WHITE : PURPLE_DK,
    });
    s.addText(t[1], {
      x: x + 0.31, y: 2.2, w: 5.3, h: 0.28, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 10.5, color: dark ? ORANGE : MUTED,
    });
    s.addText(t[2].map((b, k) => ({
      text: b, options: { bullet: true, breakLine: k !== t[2].length - 1 },
    })), {
      x: x + 0.3, y: 2.56, w: 5.34, h: 1.1, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 11, color: dark ? WHITE : INK,
      paraSpaceAfter: 4, lineSpacingMultiple: 1.15,
    });
  });

  // the two-stage flow of the AI track
  s.addText("AI 版怎么判断", {
    x: 0.62, y: 4.02, w: 4, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 14, bold: true, color: PURPLE_DK,
  });

  const stages = [
    ["第一阶段", "行为数据", "→", "皮肤等级", PURPLE],
    ["第二阶段", "＋ 用户问答，若有照片再加图像识别", "→", "综合评估", ORANGE_DK],
  ];
  stages.forEach((st, i) => {
    const y = 4.46 + i * 0.9;
    card(s, { x: 0.6, y, w: 12.13, h: 0.78, fill: i === 0 ? PURPLE_XLT : PURPLE_LT });
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.84, y: y + 0.2, w: 1.16, h: 0.38, rectRadius: 0.08, fill: { color: st[4] },
    });
    s.addText(st[0], {
      x: 0.84, y: y + 0.2, w: 1.16, h: 0.38, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, bold: true,
      color: onColor(st[4]), align: "center", valign: "middle",
    });
    s.addText(st[1], {
      x: 2.3, y: y + 0.2, w: 5.2, h: 0.38, isTextBox: true, margin: 0, valign: "middle",
      fontFace: F, fontSize: 12.5, color: INK,
    });
    s.addText(st[2], {
      x: 7.6, y: y + 0.2, w: 0.5, h: 0.38, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 13, bold: true, color: PURPLE_MD, align: "center", valign: "middle",
    });
    s.addText(st[3], {
      x: 8.2, y: y + 0.2, w: 4.3, h: 0.38, isTextBox: true, margin: 0, valign: "middle",
      fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
    });
  });

  s.addText("图像识别是延伸出来的一块：用户上传照片时参与第二阶段判断；口腔牙齿检测是额外做的验证。\n用户问答和照片默认没有，所以多数情况只靠第一阶段的行为数据。", {
    x: 0.62, y: 6.28, w: 12.0, h: 0.62, isTextBox: true, margin: 0, valign: "top",
    fontFace: F, fontSize: 10.5, color: MUTED, lineSpacingMultiple: 1.3,
  });

  pageNum(s, 6);
  s.addNotes("两条路一起走的原因：规则版是 PM 和兽医定的，解释性好，兽医愿意核对，但上限受限于人能想到的规则；AI 版是把业务规则和兽医临床经验提炼成特征让模型学，上限更高，但要真实病例数据才能训好，现在只用合成数据验证了方案可行。AI 版分两阶段：先用行为数据出一个皮肤等级，如果用户回答了问题、上传了照片，第二阶段再综合判断。但用户互动默认没有，所以多数时候是第一阶段在起作用。");
}

// =====================================================================
// S9b — 视频演示：抓挠检出与人工确认（视频内嵌，放映时点击播放）
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "02.3  成果展示", "算法平台：AI 先干粗活，人只做把关");

  s.addText("一只狗一天录 24 段、每段 1 小时，早期只能靠人从头看到尾", {
    x: 0.62, y: 1.42, w: 11.9, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11.5, color: MUTED,
  });

  const stages = ["采集数据", "AI 先标", "人工核对", "训练模型", "测试效果", "上线使用"];
  stages.forEach((t, i) => {
    const sx = 0.6 + i * 2.03;
    const hot = i === 1 || i === 2;
    s.addShape(pres.ShapeType.roundRect, {
      x: sx, y: 1.82, w: 1.88, h: 0.5, rectRadius: 0.09,
      fill: { color: hot ? PURPLE : PURPLE_LT },
    });
    s.addText(t, {
      x: sx, y: 1.82, w: 1.88, h: 0.5, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11.5, bold: true,
      color: hot ? WHITE : PURPLE_DK, align: "center", valign: "middle",
    });
    if (i < stages.length - 1) {
      s.addText("→", {
        x: sx + 1.85, y: 1.82, w: 0.21, h: 0.5, isTextBox: true, margin: 0,
        fontFace: F, fontSize: 11, bold: true, color: PURPLE_MD, align: "center", valign: "middle",
      });
    }
  });

  const vw = 7.35, vh = 7.35 / 2.11;
  const vx = 0.6, vy = 2.58;
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
    x: vx, y: vy + vh + 0.08, w: vw, h: 0.24, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 9.5, color: MUTED,
  });

  s.addText("抓挠检出与人工核对  —  请看四处", {
    x: 8.28, y: 2.58, w: 4.5, h: 0.32, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 13, bold: true, color: PURPLE_DK,
  });

  const pts = [
    ["三路画面同步", "三路画面与项圈波形共用一条时间轴"],
    ["AI 标出的片段", "模型标出 11.9 秒抓挠，把握 84%"],
    ["拿不准的片段", "AI 没把握的地方单独挑出，等人判断"],
    ["人工核对", "认领修改 / 整份通过，多人可同时进行"],
  ];
  pts.forEach((p, i) => {
    const y = 2.98 + i * 0.79;
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
    card(s, { x, y: 6.22, w: 5.94, h: 1.0, fill: PURPLE_XLT });
    const tx = x + 0.18, ty = 6.3, tw = 1.5, th = 0.84;
    s.addShape(pres.ShapeType.roundRect, {
      x: tx, y: ty, w: tw, h: th, rectRadius: 0.06, fill: { color: PURPLE_DK },
    });
    if (fs.existsSync(m[0])) fitImage(s, m[0], tx, ty, tw, th, { slide: 13 + i, tooltip: "点击放大" });
    // transparent hit area so the whole thumbnail is clickable, not just the image
    s.addShape(pres.ShapeType.roundRect, {
      x: tx, y: ty, w: tw, h: th, rectRadius: 0.06,
      fill: { color: WHITE, transparency: 100 }, line: { type: "none" },
      hyperlink: { slide: 13 + i, tooltip: "点击放大" },
    });
    s.addText("点击放大", {
      x: x + 1.84, y: 6.96, w: 3.9, h: 0.24, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 9, color: PURPLE_MD,
    });
    s.addText(m[1], {
      x: x + 1.84, y: 6.36, w: 3.9, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
    });
    s.addText(m[2], {
      x: x + 1.84, y: 6.68, w: 3.9, h: 0.3, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 10, color: INK, lineSpacingMultiple: 1.2,
    });
  });

  pageNum(s, 7);
  s.addNotes("现场点开视频播放，另外两项在浏览器里现场演示。重点讲两个机制：AI 预标注让人从「从头标」变成「审核修正」；疑似片段召回把模型最不确定的样本主动推给人工，这两条是标注效率和数据质量的关键。播放控制在 1-2 分钟。");
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
  pageNum(s, 8);
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
  pageNum(s, 9);
  s.addNotes("这五条是我试用期最重要的方法论沉淀，也是我认为可以持续复用到后续工作中的东西。");
}

// =====================================================================
// S12 — 待提升点与改进计划
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "05  GAPS & PLAN", "待提升点：分清卡在什么上");

  const groups = [
    ["等数据", 0.6, 3.2, PURPLE, [
      ["用来训练的狗太少", "狗场先加 6 只，按批次继续扩"],
    ]],
    ["等配合", 4.06, 5.2, PURPLE_DK, [
      ["还没经过测试与线上验证", "推动测试同事介入"],
      ["皮肤评分标准要校准", "请兽医标注真实病例"],
      ["松动检测还没做", "先和产品、兽医对齐判定标准"],
    ]],
    ["靠自己补", 9.52, 3.21, PURPLE_MD, [
      ["业务理解还不够深", "系统补充宠物皮肤健康知识"],
    ]],
  ];

  groups.forEach(([label, gx, gw, color, items]) => {
    card(s, { x: gx, y: 1.66, w: gw, h: 4.14, fill: PURPLE_XLT });
    s.addShape(pres.ShapeType.roundRect, {
      x: gx + 0.24, y: 1.9, w: 1.5, h: 0.44, rectRadius: 0.08, fill: { color },
    });
    s.addText(label, {
      x: gx + 0.24, y: 1.9, w: 1.5, h: 0.44, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 13, bold: true, color: onColor(color),
      align: "center", valign: "middle",
    });
    s.addText(`${items.length} 项`, {
      x: gx + gw - 1.24, y: 1.9, w: 1.0, h: 0.44, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 12, bold: true, color: PURPLE_MD,
      align: "right", valign: "middle",
    });

    items.forEach((it, k) => {
      const y = 2.58 + k * 1.06;
      card(s, { x: gx + 0.24, y, w: gw - 0.48, h: 0.9, fill: WHITE });
      s.addText(it[0], {
        x: gx + 0.44, y: y + 0.1, w: gw - 0.88, h: 0.32, isTextBox: true, margin: 0,
        fontFace: F, fontSize: 12.5, bold: true, color: PURPLE_DK,
      });
      s.addText("→  " + it[1], {
        x: gx + 0.44, y: y + 0.46, w: gw - 0.88, h: 0.32, isTextBox: true, margin: 0,
        fontFace: F, fontSize: 10.5, color: MUTED,
      });
    });
  });

  card(s, { x: 0.6, y: 6.1, w: 12.13, h: 0.8, fill: PURPLE_LT });
  s.addText("五项里有三项卡在需要别人配合 —— 这也是我下一步要主动去推的地方。", {
    x: 0.92, y: 6.1, w: 11.5, h: 0.8, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 12.5, color: INK,
  });
  pageNum(s, 10);
  s.addNotes("待提升点我按「卡在什么上」分了三类：一类等数据，靠狗场加狗解决；一类等配合，需要测试同事、兽医、产品一起推，这类占了三项；还有一类是我自己要补的业务理解。分清性质才知道该找谁、该做什么。");
}

// =====================================================================
// S13 — 下一步规划
// =====================================================================
{
  const s = pres.addSlide();
  lightBg(s);
  pageTitle(s, "05  NEXT STEPS", "下一步：要做得更好，关键是这三件事");

  const items = [
    ["01", "样本要更多", "现在只有 4 只狗", "个体越多，模型换一只没见过的狗才稳",
      "狗场先加 6 只，之后按批次继续扩", PURPLE_DK, true],
    ["02", "标注要人工审核", "但会越来越省", "模型越准，AI 预标越准，人只需核对",
      "平台上线 AI 自动标注，把人力省下来", PURPLE_XLT, false],
    ["03", "换更好的模型", "现在的数据量还不够", "数据少时传统方法更稳、更省",
      "数据上来后上深度学习，上限更高", PURPLE_XLT, false],
  ];

  items.forEach((it, i) => {
    const x = 0.6 + i * 4.11;
    const dark = it[6];
    card(s, { x, y: 1.66, w: 3.91, h: 3.5, fill: it[5] });
    badge(s, x + 0.28, 1.9, 0.5, it[0], dark ? ORANGE : PURPLE_MD, WHITE, 12);
    s.addText(it[1], {
      x: x + 0.9, y: 1.94, w: 2.9, h: 0.32, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 14.5, bold: true, color: dark ? WHITE : PURPLE_DK,
    });
    s.addText(it[2], {
      x: x + 0.3, y: 2.54, w: 3.32, h: 0.3, isTextBox: true, margin: 0,
      fontFace: F, fontSize: 11, color: dark ? ORANGE : MUTED,
    });
    s.addText(it[3], {
      x: x + 0.3, y: 2.94, w: 3.32, h: 0.7, isTextBox: true, margin: 0, valign: "top",
      fontFace: F, fontSize: 11, color: dark ? WHITE : INK, lineSpacingMultiple: 1.25,
    });
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.3, y: 3.86, w: 3.32, h: 1.06, rectRadius: 0.08,
      fill: { color: dark ? PURPLE : WHITE },
    });
    s.addText(it[4], {
      x: x + 0.44, y: 3.86, w: 3.04, h: 1.06, isTextBox: true, margin: 0, valign: "middle",
      fontFace: F, fontSize: 11, bold: true, color: dark ? WHITE : PURPLE_DK,
      lineSpacingMultiple: 1.25,
    });
  });

  card(s, { x: 0.6, y: 5.44, w: 12.13, h: 1.2, fill: PURPLE_LT });
  s.addText("三件事是连着的", {
    x: 0.9, y: 5.62, w: 2.6, h: 0.3, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 12, bold: true, color: PURPLE_DK,
  });
  s.addText("狗多了数据才够，AI 预标准了人工才省，省下来的人力才能标更多数据、支撑更好的模型。\n目标：让算法从「能认出行为」走向「能给出可信的健康结论」。", {
    x: 3.6, y: 5.58, w: 8.9, h: 0.9, isTextBox: true, margin: 0, valign: "middle",
    fontFace: F, fontSize: 11.5, color: INK, lineSpacingMultiple: 1.3,
  });
  pageNum(s, 11);
  s.addNotes("下一步就是这三件事，而且是连着的：狗多了数据才够；模型越准 AI 预标越准，人工审核越来越省；省下来的人力能标更多数据，数据够了就能上深度学习，上限更高。现在数据量还不适合上深度学习，传统方法反而更稳。");
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
    fill: { color: PURPLE }, hyperlink: { slide: 7, tooltip: "返回演示页" },
  });
  s.addText("← 返回", {
    x: 11.7, y: 0.26, w: 1.13, h: 0.36, isTextBox: true, margin: 0,
    fontFace: F, fontSize: 11, bold: true, color: WHITE,
    align: "center", valign: "middle", hyperlink: { slide: 7, tooltip: "返回演示页" },
  });
  if (fs.existsSync(img)) fitImage(s, img, 0.4, 0.82, 12.53, 6.4);
});

pres.writeFile({ fileName: "转正述职报告-Toky.pptx" }).then(f => console.log("written:", f));
