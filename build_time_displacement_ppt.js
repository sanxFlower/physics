const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = '高中物理课堂';
pptx.subject = '人教版高中物理必修一 1.2 时间 位移';
pptx.title = '1.2 时间 位移';
pptx.company = '物理教学课件';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Microsoft YaHei',
  bodyFontFace: 'Microsoft YaHei',
  lang: 'zh-CN'
};
pptx.defineSlideMaster({
  title: 'MASTER',
  background: { color: 'F7F2E8' },
  objects: [
    { rect: { x: 0, y: 0, w: 13.333, h: 0.16, fill: { color: '2B7A78' }, line: { color: '2B7A78' } } },
    { line: { x: 0.45, y: 7.08, w: 12.42, h: 0, line: { color: 'C9D8D8', width: 1 } } },
    { text: { text: '人教版高中物理 必修一 · 第一章 运动的描述', options: { x: 0.48, y: 7.12, w: 7.4, h: 0.2, fontFace: 'Microsoft YaHei', fontSize: 8, color: '6B7C85', margin: 0 } } },
    { text: { text: '1.2', options: { x: 12.15, y: 7.1, w: 0.55, h: 0.22, fontFace: 'Microsoft YaHei', fontSize: 9, bold: true, color: '2B7A78', align: 'right', margin: 0 } } }
  ],
  slideNumber: { x: 12.76, y: 7.1, color: '6B7C85', fontFace: 'Microsoft YaHei', fontSize: 8 }
});

const C = {
  navy: '0C1B2A', teal: '2B7A78', cyan: '4FB3BF', cream: 'F7F2E8', orange: 'E98A5B', yellow: 'F4C95D',
  ink: '1C2B36', grey: '6B7C85', pale: 'E8F1F2', white: 'FFFFFF', red: 'C75C5C', green: '4E9F6E', blue: '3F76A8', line: 'C9D8D8'
};
const W = 13.333, H = 7.5;
const font = 'Microsoft YaHei';
const assetsDir = path.join(__dirname, 'ppt_assets');
if (!fs.existsSync(assetsDir)) fs.mkdirSync(assetsDir, { recursive: true });

function tx(slide, text, x, y, w, h, opt = {}) {
  slide.addText(text, {
    x, y, w, h, fontFace: font, fontSize: opt.fontSize || 18, color: opt.color || C.ink,
    bold: opt.bold || false, italic: opt.italic || false, align: opt.align || 'left',
    valign: opt.valign || 'mid', margin: opt.margin === undefined ? 0.06 : opt.margin,
    breakLine: false, fit: 'shrink', paraSpaceAfterPt: opt.paraSpaceAfterPt || 0,
    bullet: opt.bullet, indent: opt.indent, breakLineOnTextOverflow: false,
    transparency: opt.transparency, isTextBox: true
  });
}
function box(slide, x, y, w, h, fill, radius = 0.12, line = fill) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w, h, rectRadius: radius, fill: { color: fill }, line: { color: line, width: 1 } });
}
function line(slide, x, y, w, h, color = C.ink, width = 1.5, endArrow = 'none', beginArrow = 'none', dash = 'solid') {
  slide.addShape(pptx.ShapeType.line, { x, y, w, h, line: { color, width, beginArrowType: beginArrow, endArrowType: endArrow, dashType: dash } });
}
function dot(slide, x, y, r, fill, lineColor = fill) {
  slide.addShape(pptx.ShapeType.ellipse, { x: x - r, y: y - r, w: r * 2, h: r * 2, fill: { color: fill }, line: { color: lineColor, width: 1 } });
}
function title(slide, part, heading, kicker = '') {
  tx(slide, part.toUpperCase(), 0.55, 0.36, 2.0, 0.28, { fontSize: 11, bold: true, color: C.teal });
  tx(slide, heading, 0.55, 0.67, 9.8, 0.5, { fontSize: 26, bold: true, color: C.navy });
  if (kicker) tx(slide, kicker, 0.58, 1.22, 11.8, 0.28, { fontSize: 11, color: C.grey });
  line(slide, 0.55, 1.54, 12.2, 0, C.line, 1);
}
function pill(slide, text, x, y, w, fill, color = C.navy) {
  box(slide, x, y, w, 0.34, fill, 0.15, fill);
  tx(slide, text, x, y + 0.02, w, 0.28, { fontSize: 11, bold: true, color, align: 'center' });
}
function bulletText(slide, items, x, y, w, h, opt = {}) {
  const fs = opt.fontSize || 15;
  const gap = opt.gap || 0.33;
  items.forEach((it, i) => {
    dot(slide, x + 0.09, y + 0.16 + i * gap, 0.045, opt.dot || C.orange);
    tx(slide, it, x + 0.22, y + i * gap, w - 0.22, gap + 0.02, { fontSize: fs, color: opt.color || C.ink, bold: opt.boldFirst && i === 0 });
  });
}
function addImageSafe(slide, p, x, y, w, h, transparency = 0) {
  slide.addImage({ path: p, x, y, w, h, transparency });
}

function timeline(slide, x, y, w) {
  line(slide, x, y, w, 0, C.navy, 1.8, 'triangle');
  for (let i = 0; i <= 8; i++) {
    const xx = x + i * (w / 8);
    line(slide, xx, y - 0.08, 0, 0.16, C.navy, 1);
    tx(slide, String(i), xx - 0.08, y + 0.11, 0.16, 0.22, { fontSize: 10, align: 'center', color: C.grey });
  }
  tx(slide, 't / h', x + w + 0.08, y - 0.13, 0.45, 0.25, { fontSize: 11, color: C.navy, italic: true });
}
function coordAxis(slide, x, y, w, labels = true) {
  line(slide, x, y, w, 0, C.navy, 1.8, 'triangle');
  for (let i = 0; i <= 10; i++) {
    const xx = x + i * (w / 10);
    line(slide, xx, y - 0.07, 0, 0.14, C.navy, 1);
    if (labels && i % 2 === 0) tx(slide, String(i - 5), xx - 0.12, y + 0.1, 0.24, 0.2, { fontSize: 9, align: 'center', color: C.grey });
  }
  tx(slide, 'x / m', x + w + 0.06, y - 0.12, 0.55, 0.24, { fontSize: 11, italic: true, color: C.navy });
}
function xtAxes(slide, x, y, w, h) {
  line(slide, x, y + h, w, 0, C.navy, 1.8, 'triangle');
  line(slide, x, y + h, 0, -h, C.navy, 1.8, 'triangle');
  tx(slide, 't', x + w + 0.08, y + h - 0.11, 0.24, 0.22, { fontSize: 12, italic: true, color: C.navy });
  tx(slide, 'x', x - 0.03, y - 0.25, 0.24, 0.22, { fontSize: 12, italic: true, color: C.navy, align: 'center' });
}

async function cropImages() {
  const src = [
    'C:/Users/San/AppData/Local/Temp/codex-clipboard-8ec99169-c797-40b7-8003-247f1761b519.jpg',
    'C:/Users/San/AppData/Local/Temp/codex-clipboard-28a588da-9815-41c3-a120-350d8a4db695.jpg',
    'C:/Users/San/AppData/Local/Temp/codex-clipboard-cf04a845-4af4-44d3-abc2-6d14c2289167.jpg',
    'C:/Users/San/AppData/Local/Temp/codex-clipboard-14854746-6a2d-4754-bd58-b386554b67d6.jpg'
  ];
  // 原图是手机长截图（含上下黑边）。按教材内容区域分块裁切，
  // 每块都保持原始像素比例，避免在 PPT 中把整页强行拉成宽屏。
  const jobs = [
    ['introTop', src[0], { left: 90, top: 530, width: 1040, height: 520 }],
    ['introTime', src[0], { left: 90, top: 1040, width: 1040, height: 550 }],
    ['coordTop', src[1], { left: 90, top: 530, width: 1040, height: 650 }],
    ['route', src[1], { left: 90, top: 1120, width: 1040, height: 840 }],
    ['dx', src[2], { left: 90, top: 530, width: 1040, height: 760 }],
    ['xt', src[2], { left: 90, top: 1260, width: 1040, height: 760 }],
    ['timerTop', src[3], { left: 90, top: 530, width: 1040, height: 720 }],
    ['timerExp', src[3], { left: 90, top: 1270, width: 1040, height: 850 }]
  ];
  const out = {};
  for (const [name, input, extract] of jobs) {
    const target = path.join(assetsDir, `教材批注_${name}.jpg`);
    try {
      await sharp(input).extract(extract).jpeg({ quality: 96, chromaSubsampling: '4:4:4' }).toFile(target);
      out[name] = target;
    } catch (e) {
      out[name] = input;
    }
  }
  return out;
}

function slideCover() {
  const s = pptx.addSlide();
  s.background = { color: C.navy };
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: W, h: H, fill: { color: C.navy }, line: { color: C.navy } });
  s.addShape(pptx.ShapeType.arc, { x: -1.0, y: 4.65, w: 7.0, h: 3.4, adjustPoint: 0.35, rotate: 10, line: { color: C.teal, width: 3, transparency: 15 }, fill: { color: C.navy, transparency: 100 } });
  s.addShape(pptx.ShapeType.arc, { x: 5.4, y: -1.8, w: 9.0, h: 6.0, adjustPoint: 0.24, rotate: 190, line: { color: C.cyan, width: 2, transparency: 25 }, fill: { color: C.navy, transparency: 100 } });
  tx(s, '高中物理 · 必修一', 0.72, 0.78, 3.3, 0.3, { fontSize: 15, color: 'A6D7D5', bold: true });
  tx(s, '1.2', 0.72, 1.5, 2.1, 0.62, { fontSize: 38, color: C.yellow, bold: true });
  tx(s, '时间  位移', 0.72, 2.2, 7.8, 0.8, { fontSize: 34, color: C.white, bold: true });
  tx(s, '把“运动”说清楚：它在何时、何处、改变了多少？', 0.76, 3.14, 8.2, 0.34, { fontSize: 17, color: 'D6E4E6' });
  pill(s, '时刻 · 时间间隔', 0.78, 4.18, 1.82, C.teal, C.white);
  pill(s, '位置 · 路程 · 位移', 2.78, 4.18, 2.08, C.orange, C.white);
  pill(s, 'x–t 图像 · 测量', 5.02, 4.18, 1.72, C.cyan, C.navy);
  // timeline motif
  line(s, 0.82, 5.66, 7.0, 0, C.white, 2, 'triangle');
  for (let i = 0; i < 7; i++) {
    const xx = 0.82 + i * 1.05;
    line(s, xx, 5.53, 0, 0.26, 'B5D5D7', 1);
    if (i === 2 || i === 5) dot(s, xx, 5.66, 0.085, C.yellow, C.yellow);
  }
  tx(s, '点：时刻', 2.75, 5.96, 1.1, 0.25, { fontSize: 11, color: C.yellow, bold: true, align: 'center' });
  tx(s, '段：时间间隔', 5.12, 5.96, 1.45, 0.25, { fontSize: 11, color: 'A6D7D5', bold: true, align: 'center' });
  tx(s, '授课对象：高一 · 人教版教材第1章', 0.78, 6.76, 5.6, 0.25, { fontSize: 11, color: '9FB5BB' });
}

function slideOverview() {
  const s = pptx.addSlide('MASTER');
  title(s, '总括', '一节课的“描述工具箱”', '先分清概念，再用图像和实验把变化量化。');
  const cards = [
    ['01', '时刻和时间间隔', '时刻是一个“点”\n时间间隔是一段“过程”', C.teal],
    ['02', '位置和位移', '坐标系描述位置\n位移描述位置的变化', C.orange],
    ['03', '时间–位移图像', '用 x–t 图像观察\n位置随时间的变化', C.cyan],
    ['04', '位移和时间的测量', '打点计时器、频闪照相\n把“点”和“段”读出来', C.yellow]
  ];
  cards.forEach((c, i) => {
    const x = 0.64 + (i % 2) * 6.18;
    const y = 1.95 + Math.floor(i / 2) * 1.65;
    box(s, x, y, 5.55, 1.27, C.white, 0.15, 'D6E4E6');
    box(s, x + 0.18, y + 0.18, 0.68, 0.45, c[3], 0.1, c[3]);
    tx(s, c[0], x + 0.18, y + 0.23, 0.68, 0.22, { fontSize: 14, bold: true, color: c[3] === C.yellow ? C.navy : C.white, align: 'center' });
    tx(s, c[1], x + 1.08, y + 0.18, 4.2, 0.3, { fontSize: 18, bold: true, color: C.navy });
    tx(s, c[2], x + 1.08, y + 0.58, 4.2, 0.48, { fontSize: 14, color: C.grey, breakLine: true });
  });
  box(s, 0.64, 5.45, 12.05, 0.92, C.pale, 0.14, C.pale);
  tx(s, '贯穿主线', 0.88, 5.61, 1.1, 0.25, { fontSize: 13, bold: true, color: C.teal });
  tx(s, '同一辆车：', 2.02, 5.61, 1.05, 0.25, { fontSize: 14, color: C.ink, bold: true });
  pill(s, '何时？', 3.12, 5.58, 0.78, C.yellow, C.navy);
  line(s, 3.96, 5.74, 0.46, 0, C.grey, 1, 'triangle');
  pill(s, '在哪？', 4.52, 5.58, 0.78, C.cyan, C.navy);
  line(s, 5.36, 5.74, 0.46, 0, C.grey, 1, 'triangle');
  pill(s, '变了多少？', 5.94, 5.58, 1.15, C.orange, C.white);
  line(s, 7.18, 5.74, 0.46, 0, C.grey, 1, 'triangle');
  pill(s, '怎么测？', 7.74, 5.58, 0.95, C.teal, C.white);
  tx(s, '课堂目标：能用“时刻/时间间隔、路程/位移、标量/矢量”准确表达运动。', 0.89, 6.47, 11.6, 0.29, { fontSize: 13, color: C.navy, bold: true });
}

function slideIntro(imgs) {
  const s = pptx.addSlide('MASTER');
  title(s, '情境导入', '问题：如何描述汽车的位置变化？', '以“驶过北京长安街的汽车”为例，先从语言走向物理量。');
  box(s, 0.65, 1.86, 6.05, 4.72, C.white, 0.16, 'D6E4E6');
  tx(s, '要描述物体位置随时间的变化，至少要回答三个问题：', 0.96, 2.12, 5.35, 0.42, { fontSize: 16, bold: true, color: C.navy });
  const qs = [
    ['01', '什么时候？', '用“时刻”或“时间间隔”说清楚。', C.yellow],
    ['02', '在哪里？', '建立坐标系，用位置坐标表示。', C.cyan],
    ['03', '变化多少？', '用路程、位移描述变化。', C.orange]
  ];
  qs.forEach((q, i) => {
    const y = 2.94 + i * 0.92;
    dot(s, 1.07, y + 0.2, 0.18, q[3]);
    tx(s, q[0], 0.91, y + 0.07, 0.32, 0.22, { fontSize: 10, bold: true, color: q[3] === C.yellow ? C.navy : C.white, align: 'center' });
    tx(s, q[1], 1.43, y, 1.45, 0.3, { fontSize: 18, bold: true, color: C.navy });
    tx(s, q[2], 2.87, y + 0.01, 3.3, 0.34, { fontSize: 14, color: C.grey });
  });
  box(s, 0.95, 5.75, 5.35, 0.53, 'FFF5D9', 0.12, 'F6E4AE');
  tx(s, '本节主线：时刻 / 时间间隔 → 位置 / 位移 → 图像 → 测量', 1.11, 5.88, 5.0, 0.24, { fontSize: 12, color: C.navy, bold: true, align: 'center' });
  // Source page crops as visual reference. Each crop keeps the original aspect ratio.
  box(s, 7.05, 1.86, 5.6, 4.72, 'EDF4F4', 0.16, 'D6E4E6');
  tx(s, '教材示例 · 批注逻辑（局部截图）', 7.32, 2.06, 3.5, 0.28, { fontSize: 13, bold: true, color: C.teal });
  addImageSafe(s, imgs.introTop, 7.82, 2.34, 4.0, 2.0);
  addImageSafe(s, imgs.introTime, 7.82, 4.47, 4.0, 2.12);
  tx(s, '上：问题情境与本节主线　下：时刻 / 时间间隔时间轴', 7.38, 6.39, 4.9, 0.2, { fontSize: 10.5, color: C.grey, italic: true, align: 'center' });
}

function slideMoment() {
  const s = pptx.addSlide('MASTER');
  title(s, '第一部分 · 时刻和时间间隔', '时刻是点，时间间隔是段', '“上午8时上课，8时45分下课”中，两个“8时”是时刻，45 min 是时间间隔。');
  box(s, 0.65, 1.88, 5.32, 4.72, C.white, 0.16, 'D6E4E6');
  pill(s, '时刻 t', 0.94, 2.14, 1.14, C.yellow, C.navy);
  tx(s, '一个瞬间：状态量的标记', 2.25, 2.16, 3.2, 0.24, { fontSize: 16, bold: true, color: C.navy });
  bulletText(s, ['例如：8:00 开始上课', '时间轴上用“点”表示', '回答“什么时候”'], 1.02, 2.75, 4.45, 1.4, { fontSize: 15, dot: C.yellow });
  pill(s, '时间间隔 Δt', 0.94, 4.46, 1.65, C.orange, C.white);
  tx(s, '一段过程：两个时刻之间的长度', 2.76, 4.48, 2.65, 0.27, { fontSize: 15, bold: true, color: C.navy });
  bulletText(s, ['例如：8:00–8:45 = 45 min', '时间轴上用“线段”表示', '回答“经历了多久”'], 1.02, 5.07, 4.45, 1.27, { fontSize: 15, dot: C.orange });
  box(s, 6.32, 1.88, 6.33, 4.72, C.pale, 0.16, C.pale);
  tx(s, '时间轴：点与段的区别', 6.63, 2.15, 3.1, 0.32, { fontSize: 18, bold: true, color: C.navy });
  timeline(s, 6.87, 3.4, 4.88);
  // points and intervals
  dot(s, 8.09, 3.4, 0.1, C.yellow);
  dot(s, 10.83, 3.4, 0.1, C.yellow);
  tx(s, '8:00', 7.83, 3.72, 0.53, 0.22, { fontSize: 12, bold: true, color: C.navy, align: 'center' });
  tx(s, '8:45', 10.56, 3.72, 0.53, 0.22, { fontSize: 12, bold: true, color: C.navy, align: 'center' });
  line(s, 8.12, 4.28, 2.65, 0, C.orange, 5, 'none', 'none');
  tx(s, 'Δt = 45 min', 8.86, 4.44, 1.42, 0.24, { fontSize: 14, color: C.orange, bold: true, align: 'center' });
  tx(s, '点：时刻', 7.45, 5.13, 1.3, 0.25, { fontSize: 13, color: C.yellow, bold: true, align: 'center' });
  tx(s, '段：时间间隔', 10.08, 5.13, 1.7, 0.25, { fontSize: 13, color: C.orange, bold: true, align: 'center' });
  box(s, 6.78, 5.77, 5.38, 0.46, C.white, 0.1, 'D6E4E6');
  tx(s, '提醒：日常语言中的“时间”，要结合上下文判断含义。', 6.95, 5.88, 5.0, 0.24, { fontSize: 12, color: C.grey, align: 'center' });
}

function slideCoord(imgs) {
  const s = pptx.addSlide('MASTER');
  title(s, '第二部分 · 位置和位移', '坐标系：把“在哪里”变成一个数', '一维直线运动通常选运动直线为 x 轴；坐标系三要素：原点、正方向、单位长度。');
  box(s, 0.64, 1.88, 7.26, 4.82, C.white, 0.16, 'D6E4E6');
  tx(s, '坐标系三要素', 0.95, 2.13, 2.3, 0.3, { fontSize: 20, bold: true, color: C.navy });
  const three = [
    ['原点 O', '从哪里开始量？', C.orange],
    ['正方向', 'x 增大的方向', C.teal],
    ['单位长度', '每一格代表多少米？', C.cyan]
  ];
  three.forEach((d, i) => {
    const x = 0.95 + i * 2.15;
    box(s, x, 2.72, 1.86, 0.88, d[2], 0.12, d[2]);
    tx(s, d[0], x + 0.09, 2.84, 1.68, 0.26, { fontSize: 16, bold: true, color: d[2] === C.orange ? C.white : C.navy, align: 'center' });
    tx(s, d[1], x + 0.1, 3.17, 1.65, 0.22, { fontSize: 10, color: d[2] === C.orange ? 'FFF7F1' : C.navy, align: 'center' });
  });
  tx(s, '例：以高考场馆为 O，向东为正方向', 1.02, 4.06, 4.8, 0.26, { fontSize: 15, color: C.grey });
  coordAxis(s, 1.16, 5.19, 5.75, false);
  tx(s, '西', 1.04, 4.77, 0.3, 0.22, { fontSize: 11, color: C.grey, align: 'center' });
  tx(s, '东（正方向）', 6.03, 4.77, 1.18, 0.22, { fontSize: 11, color: C.teal, bold: true, align: 'center' });
  dot(s, 4.47, 5.19, 0.10, C.orange);
  tx(s, 'O', 4.33, 5.43, 0.28, 0.2, { fontSize: 12, bold: true, color: C.orange, align: 'center' });
  dot(s, 5.93, 5.19, 0.10, C.cyan);
  tx(s, 'x = +30 m', 5.58, 5.46, 0.86, 0.22, { fontSize: 11, color: C.navy, align: 'center' });
  dot(s, 3.02, 5.19, 0.10, C.teal);
  tx(s, 'x = −20 m', 2.64, 5.46, 0.9, 0.22, { fontSize: 11, color: C.navy, align: 'center' });
  box(s, 8.2, 1.88, 4.45, 4.82, 'EDF4F4', 0.16, 'D6E4E6');
  tx(s, '教材批注对照', 8.49, 2.13, 2.1, 0.28, { fontSize: 14, bold: true, color: C.teal });
  addImageSafe(s, imgs.coordTop, 8.48, 2.56, 3.88, 2.42);
  tx(s, '位置 = 物体在某时刻所在的空间点。', 8.49, 5.21, 3.82, 0.24, { fontSize: 11, color: C.grey, italic: true, align: 'center' });
  box(s, 8.48, 5.62, 3.88, 0.58, C.white, 0.1, 'D6E4E6');
  tx(s, '批注强调：坐标系 = 原点 + 正方向 + 单位长度', 8.62, 5.79, 3.6, 0.2, { fontSize: 10.5, color: C.teal, bold: true, align: 'center' });
}

function slidePathDisplacement(imgs) {
  const s = pptx.addSlide('MASTER');
  title(s, '第二部分 · 位置和位移', '路程与位移：路径不同，变化可相同', '位移只由初、末位置决定；路程取决于运动轨迹的长度。');
  box(s, 0.63, 1.88, 5.85, 4.75, C.white, 0.16, 'D6E4E6');
  tx(s, '同一位移，不同路程', 0.94, 2.15, 3.2, 0.3, { fontSize: 19, bold: true, color: C.navy });
  // map-like paths
  dot(s, 1.35, 3.12, 0.11, C.orange); tx(s, 'A（初）', 1.05, 3.38, 0.9, 0.22, { fontSize: 11, color: C.orange, bold: true });
  dot(s, 5.65, 5.26, 0.11, C.teal); tx(s, 'B（末）', 5.34, 5.52, 0.92, 0.22, { fontSize: 11, color: C.teal, bold: true });
  // direct displacement
  line(s, 1.42, 3.18, 4.18, 2.02, C.orange, 3, 'triangle');
  tx(s, '位移 Δx：A → B 的有向线段', 1.43, 4.18, 2.88, 0.25, { fontSize: 12, color: C.orange, bold: true, rotate: 24 });
  // winding path
  const pts = [[1.42,3.18],[2.05,2.58],[2.82,3.02],[2.65,4.02],[3.55,4.62],[4.2,4.18],[4.96,4.82],[5.58,5.18]];
  for (let i = 0; i < pts.length - 1; i++) line(s, pts[i][0], pts[i][1], pts[i+1][0]-pts[i][0], pts[i+1][1]-pts[i][1], C.cyan, 2.2, i === pts.length - 2 ? 'triangle' : 'none');
  tx(s, '路程 s：沿轨迹的总长度', 2.45, 2.46, 2.55, 0.25, { fontSize: 12, color: C.cyan, bold: true });
  box(s, 0.98, 5.96, 5.12, 0.42, 'FFF5D9', 0.1, 'F6E4AE');
  tx(s, '结论：s ≥ |Δx|；回到原点时 Δx = 0，但 s ≠ 0。', 1.16, 6.07, 4.8, 0.2, { fontSize: 11, color: C.navy, bold: true, align: 'center' });
  box(s, 6.78, 1.88, 5.87, 4.75, C.pale, 0.16, C.pale);
  tx(s, '教材路线图 · 批注对照', 7.08, 2.15, 2.9, 0.28, { fontSize: 16, bold: true, color: C.teal });
  addImageSafe(s, imgs.route, 7.08, 2.56, 5.28, 3.82);
  tx(s, '路径不同，但初、末位置相同 → 位移相同。', 7.18, 6.42, 5.08, 0.2, { fontSize: 10.5, color: C.grey, italic: true, align: 'center' });
  /* Concept comparison is kept as editable text on the following slide. */
  /*
  const rows = [
    ['定义', '运动轨迹的长度', '初位置指向末位置的有向线段'],
    ['符号 / 单位', 's，m（无方向）', 'Δx，m（有方向）'],
    ['决定因素', '运动路径', '初、末位置'],
    ['特殊情况', '回到原点仍可能很大', '回到原点时 Δx = 0']
  ];
  const colX = [7.06, 8.25, 10.36];
  rows.forEach((r, i) => {
    const yy = 2.73 + i * 0.73;
    if (i % 2 === 0) s.addShape(pptx.ShapeType.rect, { x: 7.03, y: yy - 0.04, w: 5.3, h: 0.55, fill: { color: 'FFFFFF', transparency: 20 }, line: { color: 'FFFFFF', transparency: 100 } });
    tx(s, r[0], colX[0], yy, 1.0, 0.3, { fontSize: 12, bold: true, color: C.teal });
    tx(s, r[1], colX[1], yy, 1.92, 0.38, { fontSize: 12, color: C.ink });
    tx(s, r[2], colX[2], yy, 2.02, 0.38, { fontSize: 12, color: C.ink });
  });
  tx(s, '练习提示：先问“初、末位置在哪里”，再判断路径。', 7.08, 5.94, 5.05, 0.24, { fontSize: 12, bold: true, color: C.orange });
  */
}

function slideScalarVector() {
  const s = pptx.addSlide('MASTER');
  title(s, '第二部分 · 位置和位移', '标量与矢量：有没有方向？', '批注强调：标量只有大小；矢量既有大小又有方向。');
  box(s, 0.68, 1.93, 5.76, 4.52, 'FFF5D9', 0.16, 'F6E4AE');
  box(s, 6.88, 1.93, 5.76, 4.52, 'EAF4F3', 0.16, 'C6E3E0');
  tx(s, '标量 scalar', 1.03, 2.26, 2.3, 0.35, { fontSize: 24, bold: true, color: C.orange });
  tx(s, '只有大小，没有方向', 1.03, 2.7, 2.8, 0.25, { fontSize: 15, color: C.ink });
  bulletText(s, ['路程 s', '时间 t、温度 T、质量 m', '比较大小：直接比较数值'], 1.09, 3.25, 4.35, 1.6, { fontSize: 16, dot: C.orange });
  tx(s, '例如：s = 10 m', 1.09, 5.57, 2.6, 0.26, { fontSize: 16, bold: true, color: C.orange });
  tx(s, '矢量 vector', 7.23, 2.26, 2.4, 0.35, { fontSize: 24, bold: true, color: C.teal });
  tx(s, '既有大小，又有方向', 7.23, 2.7, 2.9, 0.25, { fontSize: 15, color: C.ink });
  bulletText(s, ['位移 Δx', '速度 v、力 F（后续学习）', '比较大小：比较绝对值'], 7.29, 3.25, 4.35, 1.6, { fontSize: 16, dot: C.teal });
  // vector arrow
  line(s, 7.45, 5.72, 3.35, 0, C.teal, 4, 'triangle');
  tx(s, '方向', 8.7, 5.9, 0.65, 0.22, { fontSize: 11, color: C.teal, bold: true, align: 'center' });
  tx(s, '大小 |Δx|', 10.12, 5.9, 1.18, 0.22, { fontSize: 11, color: C.navy, bold: true, align: 'center' });
  box(s, 2.19, 6.24, 8.9, 0.32, C.white, 0.08, 'D6E4E6');
  tx(s, '判断口诀：题目问“走了多远”→路程；问“位置变了多少、朝哪边”→位移。', 2.35, 6.29, 8.55, 0.21, { fontSize: 11.5, color: C.navy, bold: true, align: 'center' });
}

function slideDx() {
  const s = pptx.addSlide('MASTER');
  title(s, '第二部分 · 位置和位移', '直线运动的位移：Δx = x₂ − x₁', '坐标差自带正负号：正负表示方向，绝对值表示大小。');
  box(s, 0.66, 1.88, 7.35, 4.78, C.white, 0.16, 'D6E4E6');
  coordAxis(s, 1.12, 3.18, 6.0, false);
  // positions
  line(s, 4.17, 2.63, 2.1, 0, C.orange, 3, 'triangle');
  dot(s, 4.17, 3.18, 0.1, C.orange); dot(s, 6.27, 3.18, 0.1, C.teal);
  tx(s, 'x₁', 4.03, 3.48, 0.28, 0.22, { fontSize: 13, bold: true, color: C.orange, align: 'center' });
  tx(s, 'x₂', 6.13, 3.48, 0.28, 0.22, { fontSize: 13, bold: true, color: C.teal, align: 'center' });
  tx(s, '初位置', 3.84, 2.26, 0.65, 0.22, { fontSize: 11, color: C.orange, align: 'center' });
  tx(s, '末位置', 5.96, 2.26, 0.65, 0.22, { fontSize: 11, color: C.teal, align: 'center' });
  box(s, 2.52, 4.25, 3.8, 0.75, 'EAF4F3', 0.12, 'C6E3E0');
  tx(s, 'Δx = x₂ − x₁', 2.78, 4.39, 3.28, 0.4, { fontSize: 27, bold: true, color: C.teal, align: 'center' });
  tx(s, '若 Δx > 0：向 x 轴正方向\n若 Δx < 0：向 x 轴负方向', 2.15, 5.38, 4.75, 0.6, { fontSize: 15, color: C.navy, align: 'center' });
  box(s, 8.35, 1.88, 4.3, 4.78, 'FFF5D9', 0.16, 'F6E4AE');
  tx(s, '思考与试论', 8.68, 2.17, 2.1, 0.3, { fontSize: 20, bold: true, color: C.orange });
  tx(s, '物体从 A 点运动到 B 点：\nxₐ = 5 m，xᵦ = 2 m。', 8.68, 2.78, 3.34, 0.72, { fontSize: 16, color: C.ink });
  tx(s, '① 代入：Δx = 2 − 5 = −3 m\n② 大小：|Δx| = 3 m\n③ 方向：沿 x 轴负方向', 8.68, 3.87, 3.34, 1.22, { fontSize: 15, color: C.navy, bold: true });
  box(s, 8.63, 5.58, 3.42, 0.5, C.white, 0.1, 'F6E4AE');
  tx(s, '计算位移时，坐标必须带符号。', 8.77, 5.72, 3.14, 0.2, { fontSize: 11.5, color: C.red, bold: true, align: 'center' });
}

function slideXt() {
  const s = pptx.addSlide('MASTER');
  title(s, '第三部分 · 时间–位移图像', 'x–t 图像：把“位置随时间变化”画出来', '横轴是时间 t，纵轴是位置 x；每一个点代表一个时刻的位置。');
  box(s, 0.67, 1.88, 7.16, 4.84, C.white, 0.16, 'D6E4E6');
  tx(s, '读图三步', 0.98, 2.15, 1.6, 0.3, { fontSize: 20, bold: true, color: C.navy });
  const steps = [
    ['1', '找点', '给定 t，向上作垂线，与曲线交点的纵坐标就是 x。', C.yellow],
    ['2', '看线', '斜率表示速度：越陡，速度大小越大。', C.orange],
    ['3', '比变化', '同一时刻纵坐标相同 → 位置相同；水平线 → 静止。', C.teal]
  ];
  steps.forEach((st, i) => {
    const yy = 2.73 + i * 0.86;
    dot(s, 1.16, yy + 0.2, 0.16, st[3]);
    tx(s, st[0], 1.05, yy + 0.1, 0.22, 0.2, { fontSize: 12, bold: true, color: st[3] === C.yellow ? C.navy : C.white, align: 'center' });
    tx(s, st[1], 1.48, yy + 0.03, 0.8, 0.25, { fontSize: 16, bold: true, color: C.navy });
    tx(s, st[2], 2.28, yy, 4.95, 0.47, { fontSize: 13.5, color: C.grey });
  });
  box(s, 0.96, 5.55, 6.3, 0.75, 'EAF4F3', 0.12, 'C6E3E0');
  tx(s, '注意：x–t 图像不是物体运动轨迹，它只描述位置和时间的关系。', 1.16, 5.75, 5.9, 0.28, { fontSize: 13, color: C.teal, bold: true, align: 'center' });
  // graph
  box(s, 8.14, 1.88, 4.51, 4.84, 'EDF4F4', 0.16, 'D6E4E6');
  tx(s, '示意图', 8.46, 2.15, 1.1, 0.28, { fontSize: 14, bold: true, color: C.teal });
  xtAxes(s, 8.82, 2.55, 3.15, 2.65);
  // curve x-t
  const curve = [[8.84,5.18],[9.26,5.12],[9.66,4.92],[10.08,4.56],[10.53,4.1],[10.92,3.48],[11.34,2.82]];
  for (let i = 0; i < curve.length - 1; i++) line(s, curve[i][0], curve[i][1], curve[i+1][0]-curve[i][0], curve[i+1][1]-curve[i][1], C.orange, 3);
  line(s, 8.84, 5.18, 2.54, 0, C.cyan, 2.5);
  tx(s, '静止', 9.88, 5.35, 0.55, 0.22, { fontSize: 11, color: C.cyan, bold: true, align: 'center' });
  tx(s, '加速（斜率变大）', 10.58, 3.18, 1.2, 0.32, { fontSize: 11, color: C.orange, bold: true, rotate: -48, align: 'center' });
  tx(s, 'x–t 图像', 9.92, 5.8, 0.88, 0.22, { fontSize: 12, color: C.navy, bold: true, align: 'center' });
}

function slideXtPractice() {
  const s = pptx.addSlide('MASTER');
  title(s, '第三部分 · 时间–位移图像', '从图像读运动：同一坐标系里的五条线', '先看“点”与“交点”，再看“斜率”。');
  box(s, 0.66, 1.88, 7.52, 4.87, C.white, 0.16, 'D6E4E6');
  xtAxes(s, 1.36, 2.3, 5.6, 3.46);
  // axes labels and ticks
  for (let i = 1; i <= 5; i++) { line(s, 1.36 + i * 0.92, 5.7, 0, 0.12, C.navy, 1); tx(s, String(i), 1.28 + i * 0.92, 5.83, 0.18, 0.18, { fontSize: 9, color: C.grey, align: 'center' }); }
  for (let i = 1; i <= 4; i++) { line(s, 1.24, 5.7 - i * 0.7, 0.12, 0, C.navy, 1); tx(s, String(i), 0.98, 5.59 - i * 0.7, 0.2, 0.18, { fontSize: 9, color: C.grey, align: 'right' }); }
  // A: steep from t=1
  line(s, 2.28, 5.7, 2.05, -2.8, C.orange, 2.7);
  tx(s, 'A', 4.42, 2.76, 0.24, 0.22, { fontSize: 12, bold: true, color: C.orange });
  // B: gentle line from x=1
  line(s, 1.36, 4.95, 4.62, -1.22, C.teal, 2.4);
  tx(s, 'B', 5.96, 3.55, 0.24, 0.22, { fontSize: 12, bold: true, color: C.teal });
  // C horizontal at x=1.2
  line(s, 1.36, 4.86, 4.62, 0, C.cyan, 2.4);
  tx(s, 'C', 6.02, 4.73, 0.24, 0.22, { fontSize: 12, bold: true, color: C.cyan });
  // D starts at x=0 and rises modestly
  line(s, 1.36, 5.7, 2.7, -1.0, C.blue, 2.4);
  tx(s, 'D', 4.15, 4.53, 0.24, 0.22, { fontSize: 12, bold: true, color: C.blue });
  // E piecewise (turn)
  line(s, 1.36, 5.7, 2.7, -2.0, C.red, 2.4);
  line(s, 4.06, 3.7, 1.6, 1.22, C.red, 2.4);
  tx(s, 'E', 4.0, 3.42, 0.24, 0.22, { fontSize: 12, bold: true, color: C.red });
  box(s, 8.53, 1.88, 4.1, 4.87, 'EDF4F4', 0.16, 'D6E4E6');
  tx(s, '读图任务', 8.85, 2.17, 1.55, 0.3, { fontSize: 20, bold: true, color: C.navy });
  bulletText(s, ['同一时刻 t=3，哪两物体位置相同？', '哪条线表示静止？', '谁的速度方向发生改变？', '哪条线斜率最大，速度最大？'], 8.88, 2.85, 3.3, 1.85, { fontSize: 15, dot: C.orange });
  box(s, 8.82, 5.23, 3.45, 0.92, 'FFF5D9', 0.12, 'F6E4AE');
  tx(s, '答题句式：\n“看 t=… 的竖线，与…线交点的 x 值…”', 9.04, 5.42, 3.0, 0.48, { fontSize: 12, color: C.navy, bold: true, align: 'center' });
}

function slideMeasureOverview() {
  const s = pptx.addSlide('MASTER');
  title(s, '第四部分 · 位移和时间的测量', '把连续运动“切成一格一格”', '实验的核心：用已知时间间隔标记位置，再从点迹读出位移。');
  const methods = [
    ['秒表 + 刻度尺', '适合较慢、易观察的运动', C.yellow],
    ['电磁 / 电火花打点', '固定频率，纸带上留下点迹', C.orange],
    ['频闪照相', '等时间间隔连续曝光', C.cyan]
  ];
  methods.forEach((m, i) => {
    const x = 0.8 + i * 4.15;
    box(s, x, 2.0, 3.55, 2.02, C.white, 0.16, 'D6E4E6');
    box(s, x + 0.22, 2.24, 0.52, 0.52, m[2], 0.1, m[2]);
    tx(s, String(i + 1), x + 0.22, 2.37, 0.52, 0.2, { fontSize: 14, bold: true, color: m[2] === C.yellow ? C.navy : C.white, align: 'center' });
    tx(s, m[0], x + 0.93, 2.28, 2.3, 0.32, { fontSize: 17, bold: true, color: C.navy });
    tx(s, m[1], x + 0.93, 2.75, 2.35, 0.48, { fontSize: 13, color: C.grey });
  });
  tx(s, '一条“数据链”', 0.82, 4.5, 2.0, 0.3, { fontSize: 19, bold: true, color: C.navy });
  const chain = [
    ['已知频率 f', C.teal], ['时间间隔 Δt = 1/f', C.yellow], ['点迹间距 Δx', C.orange], ['判断运动状态', C.cyan]
  ];
  chain.forEach((m, i) => {
    const x = 0.86 + i * 3.0;
    pill(s, m[0], x, 5.1, 2.25, m[1], m[1] === C.yellow ? C.navy : C.white);
    if (i < chain.length - 1) line(s, x + 2.33, 5.27, 0.46, 0, C.grey, 1.2, 'triangle');
  });
  box(s, 0.85, 6.15, 11.72, 0.42, C.pale, 0.1, C.pale);
  tx(s, '实验思维：先确定“时间单位”，再用刻度把“位移”量出来。', 1.08, 6.25, 11.25, 0.2, { fontSize: 12, color: C.teal, bold: true, align: 'center' });
}

function slideTimer(imgs) {
  const s = pptx.addSlide('MASTER');
  title(s, '第四部分 · 位移和时间的测量', '两种打点计时器：同频率，不同“落点方式”', '50 Hz 时，每隔 0.02 s 打一点：T = 1/f。');
  box(s, 0.65, 1.88, 6.05, 4.82, C.white, 0.16, 'D6E4E6');
  tx(s, '电磁打点计时器', 0.98, 2.13, 2.75, 0.3, { fontSize: 20, bold: true, color: C.orange });
  pill(s, '交流 8 V', 1.0, 2.62, 1.15, C.orange, C.white);
  pill(s, '50 Hz', 2.28, 2.62, 0.93, C.yellow, C.navy);
  pill(s, 'T = 0.02 s/点', 3.34, 2.62, 1.5, C.teal, C.white);
  // schematic
  box(s, 1.04, 3.33, 4.95, 1.47, 'EDF4F4', 0.12, 'D6E4E6');
  s.addShape(pptx.ShapeType.rect, { x: 1.42, y: 4.03, w: 4.16, h: 0.18, fill: { color: 'B7C8CC' }, line: { color: 'B7C8CC' } });
  line(s, 1.68, 3.72, 0, 0.46, C.orange, 2.6); dot(s, 1.68, 3.67, 0.08, C.orange);
  s.addShape(pptx.ShapeType.rect, { x: 2.05, y: 3.54, w: 1.18, h: 0.47, fill: { color: C.orange }, line: { color: C.orange } });
  tx(s, '振片', 2.18, 3.65, 0.92, 0.2, { fontSize: 12, color: C.white, bold: true, align: 'center' });
  s.addShape(pptx.ShapeType.ellipse, { x: 4.52, y: 3.54, w: 0.52, h: 0.52, fill: { color: C.teal }, line: { color: C.teal } });
  tx(s, '线圈 + 永久磁铁', 3.3, 4.44, 2.2, 0.2, { fontSize: 11, color: C.grey, align: 'center' });
  tx(s, '电火花打点计时器', 0.98, 5.17, 2.95, 0.3, { fontSize: 20, bold: true, color: C.teal });
  tx(s, '原理相同：以固定频率记录时刻；\n纸带上留下电火花和墨粉点。', 1.0, 5.61, 4.9, 0.52, { fontSize: 14, color: C.ink });
  box(s, 7.05, 1.88, 5.6, 4.82, 'EDF4F4', 0.16, 'D6E4E6');
  tx(s, '教材器材图 · 批注要点', 7.36, 2.13, 2.65, 0.28, { fontSize: 14, bold: true, color: C.teal });
  addImageSafe(s, imgs.timerTop, 7.35, 2.53, 4.95, 3.42);
  tx(s, '纸带上的每个点，都对应一个确定的时刻。', 7.39, 6.27, 4.9, 0.24, { fontSize: 11, color: C.grey, italic: true, align: 'center' });
}

function slideStrobe() {
  const s = pptx.addSlide('MASTER');
  title(s, '第四部分 · 位移和时间的测量', '频闪照相：等时间间隔的“连续定格”', '频闪灯按固定频率闪光；同一张照片上出现物体在不同时刻的位置。');
  box(s, 0.67, 1.88, 7.55, 4.8, C.white, 0.16, 'D6E4E6');
  tx(s, '位置序列', 0.98, 2.13, 1.55, 0.3, { fontSize: 20, bold: true, color: C.navy });
  // ground and dots
  line(s, 1.12, 4.62, 6.35, 0, C.grey, 1.4);
  const xs = [1.42, 2.12, 2.95, 3.95, 5.12, 6.28];
  xs.forEach((xx, i) => {
    dot(s, xx, 4.62, 0.13, i < 2 ? C.cyan : C.orange);
    tx(s, `t${i}`, xx - 0.17, 4.92, 0.34, 0.22, { fontSize: 11, color: C.navy, align: 'center' });
    if (i < xs.length - 1) line(s, xx + 0.15, 4.2, xs[i + 1] - xx - 0.3, 0, C.orange, 1.4, 'triangle');
  });
  tx(s, '点距越来越大 → 相同时间内位移变大 → 速度变大', 1.14, 5.54, 6.2, 0.3, { fontSize: 14, color: C.orange, bold: true, align: 'center' });
  // flashes
  tx(s, '闪光频率 f', 1.1, 2.9, 1.25, 0.24, { fontSize: 14, color: C.teal, bold: true });
  line(s, 2.45, 3.04, 3.4, 0, C.teal, 2, 'triangle');
  pill(s, '每隔 Δt = 1/f', 3.35, 2.84, 1.55, C.yellow, C.navy);
  tx(s, '曝光一次', 5.1, 2.9, 1.08, 0.24, { fontSize: 14, color: C.teal, bold: true });
  box(s, 8.57, 1.88, 4.05, 4.8, 'EAF4F3', 0.16, 'C6E3E0');
  tx(s, '如何读图？', 8.88, 2.15, 1.6, 0.3, { fontSize: 20, bold: true, color: C.teal });
  bulletText(s, ['先数点：点与点之间的时间相同', '再量距：用刻度尺测 Δx', '最后比较：Δx/Δt 反映快慢'], 8.9, 2.88, 3.1, 1.5, { fontSize: 15, dot: C.teal });
  box(s, 8.86, 5.23, 3.46, 0.83, C.white, 0.12, 'C6E3E0');
  tx(s, '照片能记录“在哪里”\n也能推出“运动多快”', 9.04, 5.42, 3.1, 0.43, { fontSize: 13, color: C.navy, bold: true, align: 'center' });
}

function slideExperiment() {
  const s = pptx.addSlide('MASTER');
  title(s, '第四部分 · 位移和时间的测量', '实验：从纸带点迹读出位移和时间', '建议让学生先设计记录表，再开始测量。');
  box(s, 0.67, 1.88, 7.3, 4.84, C.white, 0.16, 'D6E4E6');
  tx(s, '实验流程', 0.98, 2.15, 1.45, 0.3, { fontSize: 20, bold: true, color: C.navy });
  const flow = [
    ['1', '固定装置', '打点计时器与长木板稳定'],
    ['2', '安装纸带', '纸带穿过限位孔，接在小车上'],
    ['3', '打点', '先接通电源，再拉动纸带'],
    ['4', '数点与计时', 'n 个点 → t = (n−1)×0.02 s'],
    ['5', '量位移', '刻度尺量起始点到第 n 点的 xₙ']
  ];
  flow.forEach((f, i) => {
    const yy = 2.77 + i * 0.63;
    dot(s, 1.13, yy + 0.18, 0.15, i < 3 ? C.teal : C.orange);
    tx(s, f[0], 1.02, yy + 0.08, 0.22, 0.2, { fontSize: 11, bold: true, color: C.white, align: 'center' });
    tx(s, f[1], 1.48, yy + 0.02, 1.45, 0.25, { fontSize: 15, bold: true, color: C.navy });
    tx(s, f[2], 2.98, yy + 0.01, 4.45, 0.3, { fontSize: 13, color: C.grey });
    if (i < flow.length - 1) line(s, 1.13, yy + 0.35, 0, 0.26, C.line, 1.4);
  });
  box(s, 8.3, 1.88, 4.35, 4.84, 'FFF5D9', 0.16, 'F6E4AE');
  tx(s, '点迹示意', 8.62, 2.15, 1.45, 0.3, { fontSize: 20, bold: true, color: C.orange });
  line(s, 8.87, 3.52, 3.2, 0, C.grey, 1.2);
  const ds = [0.28, 0.45, 0.7, 1.02, 1.35, 1.73];
  let xx = 9.02;
  ds.forEach((d, i) => { dot(s, xx, 3.52, 0.065, C.orange); tx(s, String(i + 1), xx - 0.08, 3.74, 0.16, 0.16, { fontSize: 9, color: C.navy, align: 'center' }); xx += d; });
  tx(s, '相邻点距逐渐变大', 9.2, 4.16, 2.45, 0.25, { fontSize: 14, color: C.orange, bold: true, align: 'center' });
  tx(s, '→ 相同时间内位移变大\n→ 运动越来越快', 9.05, 4.56, 2.85, 0.56, { fontSize: 15, color: C.navy, bold: true, align: 'center' });
  box(s, 8.78, 5.52, 3.35, 0.57, C.white, 0.1, 'F6E4AE');
  tx(s, '误差来源：电压过低、振针过高、纸带松弛。', 8.95, 5.69, 3.0, 0.21, { fontSize: 10.5, color: C.red, bold: true, align: 'center' });
}

function slideSummary() {
  const s = pptx.addSlide('MASTER');
  title(s, '第五部分 · 小结', '把四句话带出教室', '用一句话、一张图、一个公式，完成知识闭环。');
  const cards = [
    ['①', '时刻 & 时间间隔', '时刻是点；时间间隔是段。', C.yellow],
    ['②', '位置 & 位移', '位置要靠坐标系；位移 Δx=x₂−x₁，带方向。', C.orange],
    ['③', 'x–t 图像', '横 t、纵 x；斜率表示速度，交点表示同一位置。', C.cyan],
    ['④', '测量', '50 Hz → 0.02 s/点；数点得时间，量距得位移。', C.teal]
  ];
  cards.forEach((c, i) => {
    const x = 0.75 + (i % 2) * 6.08;
    const y = 1.98 + Math.floor(i / 2) * 1.65;
    box(s, x, y, 5.42, 1.23, C.white, 0.15, 'D6E4E6');
    box(s, x + 0.2, y + 0.2, 0.58, 0.48, c[3], 0.1, c[3]);
    tx(s, c[0], x + 0.2, y + 0.33, 0.58, 0.2, { fontSize: 13, bold: true, color: c[3] === C.yellow ? C.navy : C.white, align: 'center' });
    tx(s, c[1], x + 0.98, y + 0.17, 3.8, 0.28, { fontSize: 17, bold: true, color: C.navy });
    tx(s, c[2], x + 0.98, y + 0.58, 4.1, 0.38, { fontSize: 13.5, color: C.grey });
  });
  box(s, 0.75, 5.54, 11.55, 0.82, C.navy, 0.14, C.navy);
  tx(s, '出口题', 1.05, 5.75, 1.05, 0.25, { fontSize: 15, bold: true, color: C.yellow });
  tx(s, '物体从 x₁=+8 m 运动到 x₂=−4 m：位移是多少？方向如何？', 2.2, 5.71, 6.65, 0.3, { fontSize: 16, color: C.white, bold: true });
  tx(s, 'Δx = −4 − (+8) = −12 m，方向沿 x 轴负方向', 2.2, 6.08, 6.65, 0.22, { fontSize: 12.5, color: 'BDE0DF' });
  pill(s, '下节课：速度', 9.84, 5.8, 1.48, C.orange, C.white);
}

async function main() {
  const imgs = await cropImages();
  slideCover();
  slideOverview();
  slideIntro(imgs);
  slideMoment();
  slideCoord(imgs);
  slidePathDisplacement(imgs);
  slideScalarVector();
  slideDx();
  slideXt();
  slideXtPractice();
  slideMeasureOverview();
  slideTimer(imgs);
  slideStrobe();
  slideExperiment();
  slideSummary();
  const out = path.join(__dirname, '人教版必修一_1.2时间位移_课堂课件.pptx');
  await pptx.writeFile({ fileName: out });
  console.log(out);
}

main().catch(err => { console.error(err); process.exit(1); });
