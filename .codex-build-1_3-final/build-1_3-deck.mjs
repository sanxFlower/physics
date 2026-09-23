import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const sourcePath = "C:/Users/San/workspace/physics/高中物理/光谷二高/26级-高一/《考点同步解读 高中物理 必修 第一册 RJ（第4版）》图书配套PPT课件/正文/01-第一章 运动的描述/03-第3讲 位置变化快慢的描述——速度.pptx";
const workspaceDir = "C:/Users/San/workspace/physics";
const assetDir = path.join(workspaceDir, "调教ai用语");
const skillDir = "C:/Users/San/.codex/plugins/cache/openai-primary-runtime/presentations/26.905.11957/skills/presentations";
const stagingDir = path.join(workspaceDir, ".codex-finalizer-1_3-final");
const finalPath = path.join(workspaceDir, "高中物理/光谷二高/26级-高一/1.3 位置变化快慢的描述 速度/1.3速度_课堂课件_原稿补充版.pptx");
const runtimePython = "C:/Users/San/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe";

const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
const originalSlides = [...presentation.slides.items];
if (originalSlides.length !== 39) {
  throw new Error(`Expected 39 original slides, found ${originalSlides.length}`);
}

const blue = "#00A7DF";
const greenBand = "#E5F3EB";
const green = "#1AAE91";
const orange = "#F58220";
const textColor = "#1A1A1A";
const headerFont = "Microsoft YaHei";

function textBox(slide, text, position, style = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = {
    typeface: headerFont,
    fontSize: 24,
    color: textColor,
    autoFit: "shrinkText",
    ...style,
  };
  return shape;
}

function rect(slide, position, fill) {
  return slide.shapes.add({
    geometry: "rect",
    position,
    fill,
    line: { fill, width: 0 },
  });
}

function addShell(slide, tag, title) {
  rect(slide, { left: 0, top: 43, width: 1280, height: 2 }, blue);
  textBox(slide, "《考点同步解读 高中物理 必修第一册 RJ》", { left: 42, top: 11, width: 560, height: 30 }, {
    fontSize: 18, bold: true, color: "#007CC3",
  });
  textBox(slide, "第3讲 位置变化快慢的描述——速度", { left: 806, top: 11, width: 440, height: 30 }, {
    fontSize: 18, bold: true, color: "#007CC3", alignment: "right",
  });
  rect(slide, { left: 42, top: 54, width: 1190, height: 48 }, greenBand);
  rect(slide, { left: 42, top: 54, width: 145, height: 48 }, green);
  textBox(slide, tag, { left: 52, top: 63, width: 125, height: 28 }, {
    fontSize: 20, bold: true, color: "#FFFFFF", alignment: "center",
  });
  textBox(slide, title, { left: 202, top: 59, width: 1000, height: 35 }, {
    fontSize: 28, bold: true,
  });
}

async function addImage(slide, name, position, alt) {
  const blob = new Uint8Array(await fs.readFile(path.join(assetDir, name)));
  slide.images.add({
    blob,
    contentType: "image/png",
    alt,
    fit: "contain",
    position,
  });
}

function addPrompt(slide, prompt) {
  textBox(slide, prompt, { left: 88, top: 118, width: 1104, height: 47 }, {
    fontSize: 25, bold: true, color: "#2B5D71", alignment: "center",
  });
  rect(slide, { left: 88, top: 166, width: 1104, height: 3 }, orange);
}

function addLabel(slide, text, position) {
  textBox(slide, text, position, { fontSize: 18, bold: true, color: "#237B70" });
}

function insertAfter(anchor) {
  return presentation.slides.insert({ after: anchor }).slide;
}

let sensorTail;

// Before original slide 5: speed definition and a quick prompt.
{
  const slide = insertAfter(originalSlides[3]);
  addShell(slide, "概念导入", "速度");
  addPrompt(slide, "两名同学在相同时间内走过不同位移，谁更快？请用“位移 ÷ 时间”说明。");
  await addImage(slide, "image.png", { left: 355, top: 181, width: 570, height: 510 }, "教材补充：速度的定义、单位和方向");
}

// Before original slide 10: average velocity and instantaneous velocity.
{
  const average = insertAfter(originalSlides[8]);
  addShell(average, "概念辨析", "平均速度");
  addPrompt(average, "判断：描述一段时间内运动快慢时，使用哪一种速度？");
  await addImage(average, "image-1.png", { left: 132, top: 208, width: 1016, height: 394 }, "教材补充：平均速度的定义");

  const instant = insertAfter(average);
  addShell(instant, "概念辨析", "瞬时速度");
  addPrompt(instant, "汽车速度表显示的是平均速度还是瞬时速度？");
  await addImage(instant, "image-2.png", { left: 74, top: 199, width: 1132, height: 426 }, "教材补充：瞬时速度与速率");
}

// Before original slide 17: an interaction for the instantaneous-velocity section.
{
  const slide = insertAfter(originalSlides[15]);
  addShell(slide, "课堂互动", "读速度表");
  addPrompt(slide, "抢答：速度表从 40 km/h 变到 60 km/h，能否据此确定两分钟内的平均速度？");
  textBox(slide, "举牌选择：能 / 不能\n\n说出你的理由：平均速度需要知道这段时间内的位移和所用时间。", { left: 184, top: 285, width: 912, height: 190 }, {
    fontSize: 31, bold: true, alignment: "center", color: "#1E4962",
  });
  textBox(slide, "提示：速度表在某一时刻的读数反映瞬时速度。", { left: 210, top: 535, width: 860, height: 45 }, {
    fontSize: 24, color: "#D96B16", alignment: "center",
  });
}

// Before original slide 23: the supplied paper-strip measurement materials.
{
  const average = insertAfter(originalSlides[21]);
  addShell(average, "实验探究", "测量平均速度");
  addPrompt(average, "操作挑战：先确认哪两个量，再计算纸带上 D 到 G 的平均速度？");
  addLabel(average, "测量总思路", { left: 82, top: 184, width: 220, height: 26 });
  await addImage(average, "image-7.png", { left: 83, top: 211, width: 1114, height: 202 }, "测量平均速度的纸带示意图");
  addLabel(average, "计算关系", { left: 82, top: 430, width: 170, height: 26 });
  await addImage(average, "image-8.png", { left: 82, top: 463, width: 546, height: 167 }, "平均速度的计算关系");
  addLabel(average, "时间的测量", { left: 673, top: 430, width: 190, height: 26 });
  await addImage(average, "image-9.png", { left: 655, top: 471, width: 550, height: 76 }, "纸带计数点的选取");

  const instant = insertAfter(average);
  addShell(instant, "实验探究", "测量瞬时速度");
  addPrompt(instant, "比较 DF 和 DG 两个时间间隔，哪一个更适合近似 E 点瞬时速度？");
  addLabel(instant, "测量思路", { left: 84, top: 182, width: 180, height: 26 });
  await addImage(instant, "image-6.png", { left: 84, top: 207, width: 1112, height: 207 }, "利用较短时间间隔近似测量瞬时速度");
  addLabel(instant, "近似条件", { left: 84, top: 432, width: 180, height: 26 });
  await addImage(instant, "image-10.png", { left: 83, top: 463, width: 540, height: 195 }, "缩短时间间隔提高近似程度");
  addLabel(instant, "时间的测量", { left: 676, top: 432, width: 190, height: 26 });
  await addImage(instant, "image-11.png", { left: 653, top: 484, width: 548, height: 130 }, "纸带计数点的选取");
}

// After original slide 31: sensor measurement materials, split for readability.
{
  const equipment = insertAfter(originalSlides[30]);
  addShell(equipment, "拓展实验", "位移传感器测速度");
  addPrompt(equipment, "设备探秘：传感器怎样把小车的位置变化记录到计算机中？");
  await addImage(equipment, "image-13.png", { left: 298, top: 183, width: 684, height: 505 }, "实验室用位移传感器测速度的装置");

  const principle = insertAfter(equipment);
  addShell(principle, "拓展实验", "位移传感器的工作原理");
  addPrompt(principle, "观察 P 与 Q：小车移动时，哪一段对应位移 Δx？");
  addLabel(principle, "红外线位移传感器", { left: 90, top: 185, width: 280, height: 28 });
  addLabel(principle, "超声波位移传感器", { left: 685, top: 185, width: 300, height: 28 });
  await addImage(principle, "image-14.png", { left: 56, top: 220, width: 556, height: 293 }, "红外线位移传感器测速度原理");
  await addImage(principle, "image-15.png", { left: 660, top: 220, width: 556, height: 266 }, "超声波位移传感器测速度原理");
  sensorTail = principle;
}

// Before original slide 32: velocity-time graph construction.
{
  const slide = insertAfter(sensorTail);
  addShell(slide, "图像建模", "速度—时间图像");
  addPrompt(slide, "画图任务：横轴和纵轴各表示什么？先描点，再判断速度变化规律。");
  await addImage(slide, "image-12.png", { left: 277, top: 180, width: 726, height: 488 }, "教材补充：速度—时间图像的建立");
}

if (presentation.slides.items.length !== 48) {
  throw new Error(`Expected 48 slides after insertion, found ${presentation.slides.items.length}`);
}

// Keep the publisher's blank opening and closing slides while adding a non-visible
// drawable object so structural validators can retain them in the exported deck.
rect(originalSlides[0], { left: 0, top: 0, width: 1, height: 1 }, { color: "#FFFFFF", transparency: 100000 });
rect(originalSlides[38], { left: 0, top: 0, width: 1, height: 1 }, { color: "#FFFFFF", transparency: 100000 });

const { finalizePresentation } = await import(pathToFileURL(
  path.join(skillDir, "container_tools/artifact_tool_utils.mjs"),
).href);
await fs.mkdir(stagingDir, { recursive: true });
await fs.mkdir(path.dirname(finalPath), { recursive: true });
const candidatePath = path.join(stagingDir, "1.3速度_课堂课件_原稿补充版-candidate.pptx");
await (await PresentationFile.exportPptx(presentation)).save(candidatePath);

await finalizePresentation({
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: runtimePython,
  integrityValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
  requiredNativeTableOwnerSlides: [],
  verifyArtifactToolImport: true,
  receiptPath: path.join(stagingDir, "1.3速度_课堂课件_原稿补充版.validation.json"),
});
