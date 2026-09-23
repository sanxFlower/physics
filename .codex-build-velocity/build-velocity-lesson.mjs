import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const skillDir = "C:\\Users\\San\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.905.11957\\skills\\presentations";
const workspaceDir = "C:\\Users\\San\\workspace\\physics";
const buildDir = path.join(workspaceDir, ".codex-build-velocity");
const outputDir = path.join(workspaceDir, "高中物理", "光谷二高", "26级-高一", "1.3 位置变化快慢的描述 速度");
const finalPath = path.join(outputDir, "1.3速度_课堂课件_v3.pptx");
const sourcePath = path.join(workspaceDir, "高中物理", "光谷二高", "26级-高一", "《考点同步解读 高中物理 必修 第一册 RJ（第4版）》图书配套PPT课件", "正文", "01-第一章 运动的描述", "03-第3讲 位置变化快慢的描述——速度.pptx");
const runtimePython = "C:\\Users\\San\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
const { finalizePresentation, resolvePresentationFont } = await import(pathToFileURL(path.join(skillDir, "container_tools", "artifact_tool_utils.mjs")).href);

await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });

const W = 1280, H = 720;
const C = {
  navy: "#122746", navy2: "#0E203C", cyan: "#35B8CF", cyan2: "#5DC8D9",
  orange: "#FF9A3D", yellow: "#FFD05A", blue: "#2C7BCB", green: "#42B883",
  red: "#E95454", white: "#FFFFFF", panel: "#F6F9FD", pale: "#EAF2FA",
  ink: "#152C4C", muted: "#BCCBE0", grid: "#6F90B6",
};
const font = resolvePresentationFont({ fontFamily: "Microsoft YaHei" });
const ppt = Presentation.create({ slideSize: { width: W, height: H } });

function box(slide, left, top, width, height, fill = "none", line = { fill: "none", width: 0 }, radius = 0) {
  return slide.shapes.add({
    geometry: radius ? "roundRect" : "rect",
    position: { left, top, width, height }, fill, line,
    ...(radius ? { borderRadius: radius } : {}),
  });
}
function text(slide, value, left, top, width, height, opts = {}) {
  const s = slide.shapes.add({ geometry: "textbox", position: { left, top, width, height }, fill: "none", line: { fill: "none", width: 0 } });
  s.text = value;
  s.text.style = {
    typeface: font, fontSize: opts.size ?? 22, color: opts.color ?? C.ink,
    bold: opts.bold ?? false, align: opts.align ?? "left", verticalAlignment: opts.valign ?? "middle",
    autoFit: "shrinkText",
  };
  return s;
}
function line(slide, left, top, width, height, color = C.grid, weight = 2, opts = {}) {
  return slide.shapes.add({ geometry: "line", position: { left, top, width, height, ...opts }, fill: "none", line: { style: "solid", fill: color, width: weight } });
}
function dot(slide, cx, cy, r, fill) { return box(slide, cx - r, cy - r, 2 * r, 2 * r, fill, { fill: fill, width: 0 }, r); }
function arrow(slide, left, top, width, height, color = C.cyan) {
  return slide.shapes.add({ geometry: "rightArrow", position: { left, top, width, height }, fill: color, line: { fill: color, width: 0 } });
}
function base(slide, n, section) {
  slide.background.fill = C.navy;
  box(slide, 56, 26, 58, 62, C.orange, { fill: C.orange, width: 0 }, 8);
  text(slide, String(n).padStart(2, "0"), 56, 26, 58, 62, { size: 17, color: C.white, bold: true, align: "center" });
  text(slide, section, 138, 32, 860, 28, { size: 14, color: C.yellow, bold: true });
  line(slide, 56, 136, 1166, 0, "#58769B", 1);
  line(slide, 60, 684, 1160, 0, "#58769B", 1);
  text(slide, "光谷二高  高一年级物理", 62, 692, 260, 16, { size: 11, color: C.muted });
  text(slide, "第一章 运动的描述 · 1.3 速度", 914, 692, 306, 16, { size: 11, color: C.muted, align: "right" });
}
function title(slide, value) { text(slide, value, 136, 68, 1050, 54, { size: 35, color: C.white, bold: true }); }
function panel(slide, left, top, width, height, accent = C.cyan) {
  box(slide, left, top, width, height, C.white, { fill: "#D7E2F0", width: 1 }, 12);
  box(slide, left, top, 8, height, accent, { fill: accent, width: 0 }, 0);
}
function note(slide, value, color = C.red) {
  box(slide, 136, 600, 1008, 46, "#FFF8F5", { fill: "#FFD5C8", width: 1 }, 6);
  text(slide, value, 158, 602, 962, 40, { size: 19, color, bold: true, align: "center" });
}
function addNotes(slide, body) {
  slide.speakerNotes.textFrame.setText(`${body}\n\n资料依据：${sourcePath}`);
}
function axis(slide, left, top, width, height, xLabel, yLabel) {
  line(slide, left, top + height, width, 0, C.cyan2, 2);
  line(slide, left, top, 0, height, C.cyan2, 2);
  arrow(slide, left + width - 14, top + height - 7, 18, 14, C.cyan2);
  const up = slide.shapes.add({ geometry: "upArrow", position: { left: left - 7, top: top - 4, width: 14, height: 18 }, fill: C.cyan2, line: { fill: C.cyan2, width: 0 } });
  text(slide, xLabel, left + width + 8, top + height - 12, 40, 22, { size: 18, color: C.white, bold: true });
  text(slide, yLabel, left - 28, top - 32, 44, 22, { size: 18, color: C.white, bold: true });
}

// 1 Cover
{
  const s = ppt.slides.add(); s.background.fill = C.navy2;
  box(s, 0, 0, W, 18, C.cyan, { fill: C.cyan, width: 0 });
  text(s, "高中物理 · 必修第一册", 78, 70, 360, 30, { size: 20, color: C.cyan, bold: true });
  text(s, "1.3 位置变化快慢的描述", 78, 150, 680, 60, { size: 42, color: C.white, bold: true });
  text(s, "速 度", 78, 214, 360, 72, { size: 54, color: C.white, bold: true });
  text(s, "把“快慢”说清楚，也把“方向”说清楚", 80, 320, 620, 44, { size: 25, color: "#D8E5F5" });
  box(s, 78, 388, 390, 5, C.orange, { fill: C.orange, width: 0 });
  text(s, "速度 · 平均速度 · 瞬时速度 · 测量 · v-t 图像", 80, 430, 640, 30, { size: 17, color: C.muted });
  line(s, 800, 566, 380, 0, C.cyan, 3); line(s, 800, 210, 0, 356, C.cyan, 3);
  [800, 925, 1070].forEach((x, i) => { dot(s, x, 566, 10, C.orange); text(s, ["O", "A", "B"][i], x - 18, 587, 36, 24, { size: 16, color: C.white, bold: true, align: "center" }); });
  arrow(s, 835, 532, 210, 20, C.cyan2); text(s, "位置变化", 865, 492, 150, 26, { size: 18, color: C.yellow, bold: true, align: "center" });
  line(s, 62, 684, 1158, 0, "#58769B", 1); text(s, "光谷二高  高一年级物理", 62, 692, 260, 16, { size: 11, color: C.muted });
  addNotes(s, "开场：同一段路，谁先到达？本节用速度定量描述位置变化的快慢。\n教学目标：建立速度、平均速度与瞬时速度的概念，并能读懂简单 v-t 图像。");
}

// 2 Learning map
{
  const s = ppt.slides.add(); base(s, 2, "01 · 学习目标"); title(s, "这一节课，我们要解决什么？");
  const items = [
    ["速度", "用位移与时间的比值描述快慢", C.cyan],
    ["平均速度", "对应一段过程，方向与位移相同", C.orange],
    ["瞬时速度", "对应某时刻或某位置", C.blue],
    ["速度测量与图像", "纸带、光电门、频闪照相和 v-t 图像", C.green],
  ];
  items.forEach((it, i) => { const x = i < 2 ? 82 : 672; const y = i % 2 === 0 ? 186 : 376; panel(s, x, y, 528, 148, it[2]); text(s, it[0], x + 30, y + 24, 180, 34, { size: 27, color: C.ink, bold: true }); text(s, it[1], x + 30, y + 72, 450, 48, { size: 18, color: C.ink }); });
  note(s, "核心线索：一段过程用平均速度，一个瞬间用瞬时速度");
  addNotes(s, "先展示本节的学习地图。学生能预判本节需要掌握的四组概念。建议让学生在学习单上标记最陌生的一项。");
}

// 3 Scenario
{
  const s = ppt.slides.add(); base(s, 3, "02 · 情境导入"); title(s, "同样跑 100 m，怎样比较谁更快？");
  panel(s, 80, 176, 520, 322, C.orange); panel(s, 678, 176, 520, 322, C.cyan);
  text(s, "甲同学", 110, 202, 180, 34, { size: 27, color: C.ink, bold: true }); text(s, "乙同学", 708, 202, 180, 34, { size: 27, color: C.ink, bold: true });
  line(s, 132, 382, 380, 0, C.orange, 5); line(s, 730, 382, 380, 0, C.cyan, 5);
  [132, 512, 730, 1110].forEach((x, i) => dot(s, x, 382, 11, i < 2 ? C.orange : C.cyan));
  text(s, "100 m", 260, 320, 120, 30, { size: 23, color: C.ink, bold: true, align: "center" }); text(s, "100 m", 858, 320, 120, 30, { size: 23, color: C.ink, bold: true, align: "center" });
  text(s, "用时 16 s", 110, 424, 270, 40, { size: 28, color: C.red, bold: true }); text(s, "用时 20 s", 708, 424, 270, 40, { size: 28, color: C.red, bold: true });
  note(s, "相同位移下，用时更短，位置变化更快。如何把这种比较变成一个物理量？");
  addNotes(s, "提问：甲、乙通过的位移相同，谁的位置变化更快？学生通常能直接回答甲。追问：若位移、时间都不同，怎样比较？引出比值定义。");
}

// 4 Velocity
{
  const s = ppt.slides.add(); base(s, 4, "03 · 速度"); title(s, "速度：位移与发生这段位移所用时间的比值");
  panel(s, 80, 182, 510, 340, C.cyan); text(s, "速度", 112, 210, 150, 36, { size: 31, color: C.ink, bold: true });
  text(s, "v = Δx / Δt", 115, 274, 410, 64, { size: 39, color: C.blue, bold: true, align: "center" });
  text(s, "国际单位：m/s 或 m·s⁻¹\n常用单位：km/h", 118, 365, 380, 76, { size: 21, color: C.ink });
  text(s, "速度既有大小，又有方向，是矢量", 112, 463, 420, 30, { size: 19, color: C.red, bold: true });
  line(s, 760, 392, 330, 0, C.cyan2, 4); arrow(s, 838, 375, 178, 32, C.cyan2); dot(s, 760, 392, 10, C.orange); dot(s, 1090, 392, 10, C.orange);
  text(s, "A", 744, 414, 32, 28, { size: 18, color: C.white, bold: true, align: "center" }); text(s, "B", 1074, 414, 32, 28, { size: 18, color: C.white, bold: true, align: "center" });
  text(s, "物体从 A 到 B\n速度方向：向右", 760, 246, 330, 80, { size: 25, color: C.white, bold: true, align: "center" });
  note(s, "写公式时先说清：Δx 是位移，不是路程；Δt 是与这段位移相对应的时间");
  addNotes(s, "强调速度方向就是物体的运动方向。说明 v=Δx/Δt 是定义式，不表示 v 与 Δx 成正比、与 Δt 成反比。资料对应原课件“速度 回归教材”“对定义式的理解”。");
}

// 5 Ratio and computation
{
  const s = ppt.slides.add(); base(s, 5, "04 · 定义式的理解"); title(s, "同一段过程：位移和时间必须对应");
  panel(s, 80, 176, 720, 350, C.orange); text(s, "例：质点从 x₁ = 2 m 运动到 x₂ = 14 m，用时 3 s", 112, 208, 620, 36, { size: 23, color: C.ink, bold: true });
  text(s, "Δx = x₂ − x₁ = 12 m", 140, 290, 520, 42, { size: 29, color: C.blue, bold: true });
  text(s, "v = Δx / Δt = 12 / 3 = 4 m/s", 140, 365, 570, 42, { size: 29, color: C.blue, bold: true });
  text(s, "方向沿 x 轴正方向", 140, 442, 400, 30, { size: 22, color: C.red, bold: true });
  panel(s, 852, 176, 346, 350, C.cyan); text(s, "比值定义", 884, 212, 240, 32, { size: 26, color: C.ink, bold: true });
  text(s, "Δx 更大\n不一定 v 更大", 892, 286, 250, 60, { size: 22, color: C.ink, bold: true, align: "center" });
  text(s, "比较速度时\n必须同时看 Δx 与 Δt", 882, 390, 270, 70, { size: 20, color: C.red, bold: true, align: "center" });
  note(s, "计算顺序：先定过程，再写位移和时间，最后代入公式并写明方向");
  addNotes(s, "让学生独立完成例题，再请一位学生说明“为什么不能只比较位移”。易错点：把位置坐标代替位移，或遗漏速度方向。");
}

// 6 Direction and sign
{
  const s = ppt.slides.add(); base(s, 6, "05 · 速度的方向"); title(s, "一维运动中，速度的正负号表示方向");
  panel(s, 80, 190, 540, 300, C.blue); panel(s, 660, 190, 540, 300, C.orange);
  text(s, "取向右为正方向", 108, 218, 320, 30, { size: 22, color: C.ink, bold: true }); text(s, "取向右为正方向", 688, 218, 320, 30, { size: 22, color: C.ink, bold: true });
  line(s, 142, 350, 380, 0, C.cyan2, 4); arrow(s, 256, 333, 160, 34, C.cyan2); text(s, "v = + 5 m/s", 170, 397, 300, 38, { size: 29, color: C.blue, bold: true, align: "center" }); text(s, "沿正方向运动", 175, 446, 290, 26, { size: 18, color: C.ink, align: "center" });
  line(s, 722, 350, 380, 0, C.cyan2, 4); const leftArrow = s.shapes.add({ geometry: "leftArrow", position: { left: 836, top: 333, width: 160, height: 34 }, fill: C.orange, line: { fill: C.orange, width: 0 } }); text(s, "v = − 5 m/s", 750, 397, 300, 38, { size: 29, color: C.orange, bold: true, align: "center" }); text(s, "沿负方向运动", 755, 446, 290, 26, { size: 18, color: C.ink, align: "center" });
  note(s, "“−5 m/s”不是速度大小为负，速度大小为 5 m/s，负号只表示方向");
  addNotes(s, "请学生判断两辆同样“5 m/s”的小车是否速度相同。答案：若方向相反，速度不相同。原课件中的变式题强调了正负号表示方向。");
}

// 7 Concept check
{
  const s = ppt.slides.add(); base(s, 7, "06 · 课堂辨析"); title(s, "判断：关于速度，哪些说法正确？");
  const options = [
    ["A", "速度的大小等于路程与时间的比值", C.red],
    ["B", "速度是描述物体运动快慢和方向的物理量", C.green],
    ["C", "v = Δx / Δt 适用于任何运动", C.green],
    ["D", "位移为 0 时，平均速度一定为 0", C.green],
  ];
  options.forEach((o, i) => { const x = i % 2 === 0 ? 82 : 660; const y = i < 2 ? 184 : 374; panel(s, x, y, 538, 138, o[2]); text(s, o[0], x + 30, y + 30, 52, 40, { size: 30, color: o[2], bold: true }); text(s, o[1], x + 96, y + 24, 398, 76, { size: 21, color: C.ink, bold: true }); });
  note(s, "答案：B、C、D。A 中应为“位移”，不是“路程”", C.red);
  addNotes(s, "让学生先独立判断，再同桌说明理由。A 是本节第一处关键辨析。C 是速度定义式的普遍适用性，区别于匀速直线运动中的 v 不变。");
}

// 8 Average velocity
{
  const s = ppt.slides.add(); base(s, 8, "07 · 平均速度"); title(s, "平均速度描述一段过程的平均快慢和方向");
  panel(s, 80, 182, 520, 332, C.orange); text(s, "平均速度", 112, 210, 210, 34, { size: 30, color: C.ink, bold: true }); text(s, "v̄ = Δx / Δt", 122, 290, 390, 58, { size: 38, color: C.orange, bold: true, align: "center" }); text(s, "对应一段位移和一段时间\n方向与该段位移方向相同", 122, 382, 390, 72, { size: 21, color: C.ink, align: "center" });
  panel(s, 676, 182, 522, 332, C.cyan); text(s, "平均速率", 708, 210, 210, 34, { size: 30, color: C.ink, bold: true }); text(s, "平均速率 = 路程 / 时间", 708, 290, 420, 42, { size: 27, color: C.blue, bold: true, align: "center" }); text(s, "只描述平均快慢\n没有方向", 740, 382, 354, 60, { size: 21, color: C.ink, align: "center" });
  note(s, "平均速度不是“各时刻速度的平均值”，而是“位移与时间的比值”");
  addNotes(s, "原课件将“平均速度”列为独立考点。此页同时引入平均速率，帮助学生在往返问题中区分位移和路程。");
}

// 9 Return trip
{
  const s = ppt.slides.add(); base(s, 9, "08 · 平均速度与平均速率"); title(s, "往返一次：平均速度和平均速率为什么不同？");
  line(s, 150, 332, 960, 0, C.cyan2, 5); [150, 630, 1110].forEach(x => dot(s, x, 332, 13, C.orange));
  text(s, "A", 133, 360, 34, 24, { size: 18, color: C.white, bold: true, align: "center" }); text(s, "B", 613, 360, 34, 24, { size: 18, color: C.white, bold: true, align: "center" }); text(s, "A", 1093, 360, 34, 24, { size: 18, color: C.white, bold: true, align: "center" });
  arrow(s, 260, 286, 260, 26, C.orange); const back = s.shapes.add({ geometry: "leftArrow", position: { left: 742, top: 356, width: 260, height: 26 }, fill: C.blue, line: { fill: C.blue, width: 0 } });
  text(s, "去程：100 m，用时 10 s", 214, 224, 350, 28, { size: 20, color: C.yellow, bold: true, align: "center" }); text(s, "返程：100 m，用时 15 s", 706, 410, 350, 28, { size: 20, color: C.cyan2, bold: true, align: "center" });
  panel(s, 120, 500, 470, 72, C.orange); text(s, "平均速度 = 0 / 25 = 0 m/s", 138, 516, 430, 40, { size: 22, color: C.ink, bold: true, align: "center" });
  panel(s, 690, 500, 470, 72, C.cyan); text(s, "平均速率 = 200 / 25 = 8 m/s", 708, 516, 430, 40, { size: 22, color: C.ink, bold: true, align: "center" });
  addNotes(s, "强调往返结束时初、末位置相同，因此总位移为零，平均速度为零；总路程为 200 m，因此平均速率为 8 m/s。让学生确认两者分母都是总时间。");
}

// 10 Instantaneous velocity
{
  const s = ppt.slides.add(); base(s, 10, "09 · 瞬时速度"); title(s, "瞬时速度描述某一时刻或某一位置的运动快慢和方向");
  panel(s, 80, 182, 480, 320, C.cyan); text(s, "平均速度", 112, 214, 220, 34, { size: 29, color: C.ink, bold: true }); text(s, "对应一段过程", 112, 280, 290, 30, { size: 23, color: C.ink }); text(s, "例如：前 5 s 内", 112, 354, 310, 30, { size: 22, color: C.blue, bold: true });
  panel(s, 720, 182, 480, 320, C.orange); text(s, "瞬时速度", 752, 214, 220, 34, { size: 29, color: C.ink, bold: true }); text(s, "对应一个时刻或位置", 752, 280, 360, 30, { size: 23, color: C.ink }); text(s, "例如：第 5 s 末", 752, 354, 310, 30, { size: 22, color: C.orange, bold: true });
  line(s, 560, 342, 160, 0, C.cyan2, 3); arrow(s, 612, 324, 54, 36, C.cyan2);
  note(s, "测量瞬时速度的思路：把取样时间缩短到足够小");
  addNotes(s, "追问：汽车仪表盘显示的速度，描述哪一个概念？引出瞬时速度。说明实际测量以很短时间内的平均速度近似瞬时速度。");
}

// 11 Paper tape
{
  const s = ppt.slides.add(); base(s, 11, "10 · 纸带测量"); title(s, "纸带：用一小段时间内的平均速度近似瞬时速度");
  panel(s, 80, 170, 1120, 158, C.green); box(s, 146, 220, 982, 60, "#D5E8B5", { fill: "#A7C782", width: 1 }, 6);
  const pts = [216, 302, 405, 527, 668, 828, 1007]; pts.forEach((x, i) => { dot(s, x, 250, 5, C.ink); text(s, String.fromCharCode(65 + i), x - 16, 282, 32, 22, { size: 16, color: C.ink, bold: true, align: "center" }); });
  line(s, 405, 204, 0, 50, C.ink, 2); line(s, 668, 204, 0, 50, C.ink, 2); arrow(s, 416, 178, 235, 18, C.orange); text(s, "Δx", 503, 146, 80, 26, { size: 21, color: C.yellow, bold: true, align: "center" });
  panel(s, 80, 390, 534, 150, C.orange); text(s, "平均速度", 110, 420, 160, 32, { size: 27, color: C.ink, bold: true }); text(s, "v̄ = Δx / Δt", 110, 470, 340, 34, { size: 27, color: C.orange, bold: true });
  panel(s, 666, 390, 534, 150, C.cyan); text(s, "瞬时速度近似", 696, 420, 220, 32, { size: 27, color: C.ink, bold: true }); text(s, "取靠近目标点的更短纸带区间", 696, 470, 430, 34, { size: 20, color: C.blue, bold: true });
  note(s, "纸带上相邻点间距越大，说明相等时间内位移越大，物体运动越快");
  addNotes(s, "原课件的实验部分强调：测某两点间平均速度，测量两点间位移和对应时间即可。进一步说明，取目标点附近的短时间区间可以近似该点瞬时速度。");
}

// 12 Experiment operation
{
  const s = ppt.slides.add(); base(s, 12, "11 · 实验操作"); title(s, "打点计时器：先通电，再拉动纸带");
  const steps = [
    ["1", "固定计时器，正确穿入纸带"], ["2", "接通交流电源"],
    ["3", "水平拉动纸带，留下清晰点迹"], ["4", "关闭电源，测量位移和时间"],
  ];
  steps.forEach((st, i) => { const x = 82 + i * 283; box(s, x, 234, 58, 58, i === 1 ? C.orange : C.cyan, { fill: "none", width: 0 }, 29); text(s, st[0], x, 234, 58, 58, { size: 22, color: C.white, bold: true, align: "center" }); text(s, st[1], x - 12, 320, 170, 84, { size: 20, color: C.white, bold: true, align: "center" }); if (i < 3) arrow(s, x + 180, 250, 66, 24, C.cyan2); });
  panel(s, 154, 470, 972, 80, C.orange); text(s, "若打点频率 f = 50 Hz，则相邻两点时间间隔 T = 1/f = 0.02 s", 178, 488, 920, 40, { size: 23, color: C.ink, bold: true, align: "center" });
  note(s, "数点时：n 个点对应 n − 1 个时间间隔", C.red);
  addNotes(s, "操作演示：严格强调先通电再拉纸带。可提问：从第 1 个点数到第 8 个点，共有几个时间间隔？答案为 7 个，时间 0.14 s。原课件给出电磁打点计时器和电火花计时器的基本使用。");
}

// 13 Measurement alternatives
{
  const s = ppt.slides.add(); base(s, 13, "12 · 测量方法"); title(s, "还有哪些方法可以测速度？");
  panel(s, 82, 184, 340, 330, C.cyan); text(s, "光电门", 112, 216, 220, 34, { size: 29, color: C.ink, bold: true }); box(s, 170, 300, 48, 110, C.cyan2, { fill: C.cyan2, width: 0 }, 5); box(s, 282, 300, 48, 110, C.cyan2, { fill: C.cyan2, width: 0 }, 5); box(s, 218, 332, 64, 48, C.orange, { fill: C.orange, width: 0 }, 0); text(s, "v ≈ d / Δt\nd 为遮光条宽度", 112, 436, 246, 54, { size: 18, color: C.ink, bold: true, align: "center" });
  panel(s, 470, 184, 340, 330, C.orange); text(s, "频闪照相", 500, 216, 230, 34, { size: 29, color: C.ink, bold: true }); line(s, 510, 368, 238, 0, C.grid, 3); [530, 564, 610, 678, 748].forEach((x, i) => dot(s, x, 368, 11, i < 2 ? C.cyan : C.orange)); text(s, "相等时间拍下多个位置", 500, 436, 250, 32, { size: 18, color: C.ink, bold: true, align: "center" });
  panel(s, 858, 184, 340, 330, C.green); text(s, "共同思路", 888, 216, 230, 34, { size: 29, color: C.ink, bold: true }); text(s, "测量位移\n记录对应时间\n用 v = Δx / Δt 计算", 890, 292, 248, 130, { size: 21, color: C.ink, bold: true, align: "center" });
  note(s, "无论用哪种仪器，核心都是“位移”与“时间”一一对应");
  addNotes(s, "原课件拓展了光电门和频闪照相法。光电门中遮光时间很短，遮光条在这段时间内的平均速度近似滑块经过光电门的瞬时速度。");
}

// 14 v-t chart intro
{
  const s = ppt.slides.add(); base(s, 14, "13 · 速度—时间图像"); title(s, "v-t 图像：速度随时间变化的规律");
  axis(s, 140, 188, 440, 310, "t", "v");
  line(s, 140, 260, 390, 166, C.cyan, 4, { horizontalFlip: true }); line(s, 140, 354, 390, 0, C.orange, 4); line(s, 140, 525, 390, 50, C.green, 4);
  text(s, "速度为正", 570, 246, 170, 26, { size: 19, color: C.cyan2, bold: true }); text(s, "匀速运动", 570, 338, 170, 26, { size: 19, color: C.yellow, bold: true }); text(s, "速度为负", 570, 440, 170, 26, { size: 19, color: C.green, bold: true });
  panel(s, 790, 194, 370, 300, C.cyan); text(s, "读图先看两件事", 824, 224, 300, 32, { size: 26, color: C.ink, bold: true }); text(s, "1. 点的纵坐标：某时刻的瞬时速度\n\n2. 图线是否在时间轴上方：速度方向", 824, 290, 300, 150, { size: 20, color: C.ink });
  note(s, "v-t 图像不是运动轨迹，而是速度随时间变化的图像");
  addNotes(s, "原课件的“速度—时间图像”部分：图像反映速度随时间的变化。时间轴上方表示沿正方向运动，下方表示沿负方向运动。平行于时间轴表示匀速运动。");
}

// 15 v-t chart application
{
  const s = ppt.slides.add(); base(s, 15, "14 · 图像应用"); title(s, "从一张 v-t 图像读出哪些信息？");
  axis(s, 130, 190, 390, 310, "t", "v");
  line(s, 130, 430, 120, 0, C.orange, 4); line(s, 250, 300, 150, 130, C.cyan, 4, { horizontalFlip: true }); line(s, 400, 300, 120, 130, C.green, 4);
  text(s, "I", 174, 398, 26, 20, { size: 18, color: C.yellow, bold: true, align: "center" }); text(s, "II", 300, 340, 34, 20, { size: 18, color: C.cyan2, bold: true, align: "center" }); text(s, "III", 445, 364, 42, 20, { size: 18, color: C.green, bold: true, align: "center" });
  const facts = [["I 段", "速度不变，做匀速直线运动", C.orange], ["II 段", "速度增大，沿正方向加速", C.cyan], ["III 段", "速度减小，仍沿正方向运动", C.green]];
  facts.forEach((f, i) => { panel(s, 660, 184 + i * 112, 500, 86, f[2]); text(s, f[0], 692, 200 + i * 112, 90, 28, { size: 21, color: f[2], bold: true }); text(s, f[1], 800, 200 + i * 112, 326, 34, { size: 19, color: C.ink, bold: true }); });
  note(s, "图线斜率越大，速度变化越快。斜率的物理意义将在下一节学习");
  addNotes(s, "先让学生读出 I、II、III 三段的速度方向和变化。提醒：速度变小不等于沿负方向运动，方向要看速度的正负。关于斜率与加速度仅作铺垫，不作本节重点。");
}

// 16 Class practice
{
  const s = ppt.slides.add(); base(s, 16, "15 · 当堂练习"); title(s, "练一练：将“快慢”和“方向”同时说准确");
  panel(s, 80, 176, 1120, 118, C.cyan); text(s, "1. 物体从 x₁ = −4 m 运动到 x₂ = +8 m，用时 3 s。求平均速度。", 112, 198, 1020, 38, { size: 23, color: C.ink, bold: true }); text(s, "提示：先求 Δx，再除以 Δt，最后说明方向", 112, 246, 950, 24, { size: 17, color: C.red, bold: true });
  panel(s, 80, 328, 1120, 118, C.orange); text(s, "2. 物体以 5 m/s 向东运动，若取向东为正方向，速度应写成什么？", 112, 350, 1020, 38, { size: 23, color: C.ink, bold: true }); text(s, "提示：速度的正负号表示方向", 112, 398, 950, 24, { size: 17, color: C.red, bold: true });
  panel(s, 80, 480, 1120, 86, C.green); text(s, "参考答案：1. v̄ = 4 m/s，沿 x 轴正方向；2. v = +5 m/s", 112, 501, 1020, 34, { size: 22, color: C.ink, bold: true, align: "center" });
  addNotes(s, "可先遮住参考答案，留 2 分钟独立完成。第 1 题检测学生是否能把带符号的位置坐标正确代入位移公式。第 2 题检测正负号的物理意义。");
}

// 17 Exit ticket
{
  const s = ppt.slides.add(); base(s, 17, "16 · 课堂检测"); title(s, "离开教室前，请完成这三题");
  const qs = [
    ["01", "v = Δx / Δt 中的 Δx 是路程还是位移？"],
    ["02", "一段往返运动的总位移为 0，平均速度是多少？"],
    ["03", "v-t 图像中，图线在时间轴下方说明什么？"],
  ];
  qs.forEach((q, i) => { const y = 174 + i * 126; box(s, 100, y, 76, 76, i === 1 ? C.orange : C.cyan, { fill: "none", width: 0 }, 10); text(s, q[0], 100, y, 76, 76, { size: 22, color: C.white, bold: true, align: "center" }); panel(s, 206, y, 930, 76, i === 1 ? C.orange : C.cyan); text(s, q[1], 236, y + 16, 850, 42, { size: 22, color: C.ink, bold: true }); });
  note(s, "答案：位移；0；速度为负，物体沿选定正方向的反方向运动");
  addNotes(s, "请学生在便签或学习单上作答，作为本节出口条。根据第 1、3 题判断学生对“方向”的理解是否牢固。");
}

// 18 Summary
{
  const s = ppt.slides.add(); s.background.fill = C.navy2;
  text(s, "本节小结", 78, 72, 350, 50, { size: 38, color: C.white, bold: true });
  const items = [
    ["速度", "v = Δx / Δt，矢量，方向与运动方向相同", C.cyan],
    ["平均速度", "描述一段过程，方向与该段位移方向相同", C.orange],
    ["瞬时速度", "描述某时刻或某位置，可由短时平均速度近似", C.blue],
    ["测量与图像", "纸带、光电门、频闪照相；v-t 图像描述速度变化", C.green],
  ];
  items.forEach((it, i) => { const y = 166 + i * 106; box(s, 82, y, 10, 74, it[2], { fill: it[2], width: 0 }); text(s, it[0], 122, y + 10, 190, 30, { size: 24, color: C.white, bold: true }); text(s, it[1], 330, y + 10, 760, 40, { size: 19, color: "#D8E5F5" }); });
  box(s, 80, 615, 740, 5, C.orange, { fill: C.orange, width: 0 }); text(s, "下一节：速度变化快慢的描述——加速度", 80, 640, 580, 30, { size: 20, color: C.yellow, bold: true });
  addNotes(s, "用四句话收束全课。布置预习：观察 v-t 图像上图线倾斜程度的不同，下节课研究速度变化快慢的物理量。\n资料参考：用户提供的速度原课件及 1.2 课堂课件的版式风格。");
}

const candidatePath = path.join(buildDir, "velocity-candidate.pptx");
await (await PresentationFile.exportPptx(ppt)).save(candidatePath);

const result = await finalizePresentation({
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: runtimePython,
  integrityValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
  explicitTotalSlideCount: 18,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
  verifyArtifactToolImport: true,
  receiptPath: path.join(buildDir, "velocity-v3-validation.json"),
});
console.log(JSON.stringify({ finalPath, result }, null, 2));
