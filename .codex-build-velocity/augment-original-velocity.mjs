import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "C:\\Users\\San\\workspace\\physics";
const skillDir = "C:\\Users\\San\\.codex\\plugins\\cache\\openai-primary-runtime\\presentations\\26.905.11957\\skills\\presentations";
const buildDir = path.join(workspaceDir, ".codex-build-original-augment");
const sourcePath = path.join(workspaceDir, "高中物理", "光谷二高", "26级-高一", "《考点同步解读 高中物理 必修 第一册 RJ（第4版）》图书配套PPT课件", "正文", "01-第一章 运动的描述", "03-第3讲 位置变化快慢的描述——速度.pptx");
const outputDir = path.join(workspaceDir, "高中物理", "光谷二高", "26级-高一", "1.3 位置变化快慢的描述 速度");
const finalPath = path.join(outputDir, "1.3速度_原版增补课件.pptx");
const assetDir = path.join(workspaceDir, "调教ai用语");
const runtimePython = "C:\\Users\\San\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\python\\python.exe";
const { finalizePresentation } = await import(pathToFileURL(path.join(skillDir, "container_tools", "artifact_tool_utils.mjs")).href);

await fs.mkdir(buildDir, { recursive: true });
await fs.mkdir(outputDir, { recursive: true });

const deck = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
const originals = [...deck.slides.items];
if (originals.length !== 39) throw new Error(`Expected 39 original slides, found ${originals.length}`);

const C = { blue: "#0070C0", teal: "#00A98F", pale: "#E4F2EB", orange: "#ED7D31", ink: "#111111", muted: "#3F6F98" };
const font = "Microsoft YaHei";

function addText(slide, value, left, top, width, height, opts = {}) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    position: { left, top, width, height },
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  shape.text = value;
  shape.text.style = {
    typeface: font,
    fontSize: opts.size ?? 24,
    color: opts.color ?? C.ink,
    bold: opts.bold ?? false,
    align: opts.align ?? "left",
    verticalAlignment: opts.valign ?? "middle",
    autoFit: "shrinkText",
  };
  return shape;
}

function addRect(slide, left, top, width, height, fill, line = { fill: "none", width: 0 }) {
  return slide.shapes.add({ geometry: "rect", position: { left, top, width, height }, fill, line });
}

function addHeader(slide, title) {
  slide.background.fill = "#FFFFFF";
  addText(slide, "《考点同步解读 高中物理 必修 第一册 RJ》", 14, 7, 510, 34, { size: 19, color: C.blue, bold: true });
  addText(slide, "第3讲 位置变化快慢的描述——速度", 774, 7, 492, 34, { size: 19, color: C.blue, bold: true, align: "right" });
  addRect(slide, 12, 46, 1256, 2, "#00A5D5");
  addRect(slide, 42, 60, 1190, 48, C.pale);
  addRect(slide, 42, 60, 164, 48, C.teal);
  addText(slide, "补充材料", 50, 62, 144, 42, { size: 25, color: "#FFFFFF", bold: true, align: "center" });
  addText(slide, title, 226, 62, 985, 42, { size: 27, color: C.ink, bold: true });
}

async function addAsset(slide, fileName, box) {
  const bytes = await fs.readFile(path.join(assetDir, fileName));
  slide.images.add({ blob: bytes, contentType: "image/png", alt: fileName, fit: "contain", position: box });
}

function addPrompt(slide, prompt) {
  addRect(slide, 54, 650, 1172, 42, "#FFF5EA", { fill: "#ED7D31", width: 1 });
  addText(slide, "想一想", 70, 653, 112, 34, { size: 21, color: C.orange, bold: true, align: "center" });
  addText(slide, prompt, 190, 653, 1004, 34, { size: 20, color: C.ink, bold: true });
}

async function insertMaterial(after, spec) {
  const slide = deck.slides.insert({ after });
  addHeader(slide, spec.title);
  for (const asset of spec.assets) await addAsset(slide, asset.file, asset.position);
  addPrompt(slide, spec.prompt);
  slide.speakerNotes.textFrame.setText(`补充材料：${spec.assets.map(asset => asset.file).join("、")}\n课堂互动：${spec.prompt}`);
  return slide;
}

// Insert after persistent references to the original slides, so page references remain those in the supplied 39-page deck.
await insertMaterial(originals[3], {
  title: "速度：用位移和时间描述快慢",
  assets: [{ file: "image.png", position: { left: 226, top: 122, width: 828, height: 500 } }],
  prompt: "两位同学跑相同路程，怎样比较谁运动得更快？",
});

let chain = await insertMaterial(originals[8], {
  title: "平均速度：一段过程的快慢",
  assets: [{ file: "image-1.png", position: { left: 80, top: 150, width: 1120, height: 435 } }],
  prompt: "为什么定义平均速度时，要强调“某一段时间内”？",
});
chain = await insertMaterial(chain, {
  title: "瞬时速度：一个时刻的快慢",
  assets: [{ file: "image-2.png", position: { left: 74, top: 130, width: 1132, height: 425 } }],
  prompt: "汽车速度表的示数，通常描述平均速度还是瞬时速度？",
});

chain = await insertMaterial(originals[21], {
  title: "测量平均速度：纸带上的总思路",
  assets: [
    { file: "image-7.png", position: { left: 80, top: 122, width: 1120, height: 323 } },
    { file: "image-8.png", position: { left: 160, top: 450, width: 960, height: 164 } },
  ],
  prompt: "计算 D、G 两点间的平均速度，需要测量哪两个物理量？",
});
chain = await insertMaterial(chain, {
  title: "测量平均速度：计数点与时间间隔",
  assets: [{ file: "image-9.png", position: { left: 80, top: 210, width: 1120, height: 155 } }],
  prompt: "每隔 5 个点选一个计数点时，相邻计数点之间对应多长时间？",
});
chain = await insertMaterial(chain, {
  title: "测量瞬时速度：用短时间平均速度逼近",
  assets: [{ file: "image-6.png", position: { left: 68, top: 178, width: 1144, height: 315 } }],
  prompt: "D、F 两点离 E 点越近，用平均速度估算 E 点瞬时速度会怎样？",
});
chain = await insertMaterial(chain, {
  title: "测量瞬时速度：计数点的选择",
  assets: [
    { file: "image-10.png", position: { left: 230, top: 120, width: 820, height: 396 } },
    { file: "image-11.png", position: { left: 100, top: 520, width: 1080, height: 105 } },
  ],
  prompt: "每隔 3 个点取一个计数点时，应该怎样确定对应的时间间隔？",
});

chain = await insertMaterial(originals[30], {
  title: "速度—时间图像：从数据到规律",
  assets: [{ file: "image-13.png", position: { left: 258, top: 118, width: 764, height: 564 } }],
  prompt: "观察 v-t 图像时，横轴和纵轴分别表示什么？",
});
chain = await insertMaterial(chain, {
  title: "位移传感器测速度：装置与原理",
  assets: [
    { file: "image-14.png", position: { left: 55, top: 178, width: 560, height: 295 } },
    { file: "image-15.png", position: { left: 655, top: 194, width: 560, height: 267 } },
  ],
  prompt: "传感器记录两个位置和对应时间后，怎样计算物体的速度？",
});
await insertMaterial(chain, {
  title: "位移传感器测速度：实验装置",
  assets: [{ file: "image-12.png", position: { left: 260, top: 118, width: 760, height: 510 } }],
  prompt: "改变小车的运动方向后，传感器记录的位移符号可能怎样变化？",
});

const candidatePath = path.join(buildDir, "1.3原版增补候选.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidatePath);

const requirements = {
  explicitTotalSlideCount: 49,
  requiredNativeTableOwnerSlides: [],
  requiredNativeChartOwnerSlides: [],
};
const result = await finalizePresentation({
  ...requirements,
  workspaceDir,
  candidatePath,
  finalPath,
  pythonExecutable: runtimePython,
  integrityValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(skillDir, "container_tools", "inspect_presentation_layout_geometry.py"),
  layoutArgs: ["--expected-slide-size-emu", "12192000,6858000", "--validate-bullet-geometry", "--validate-heading-fit"],
  verifyArtifactToolImport: true,
  receiptPath: path.join(buildDir, "1.3原版增补验证.json"),
});
console.log(JSON.stringify({ finalPath, slideCount: deck.slides.items.length, result }, null, 2));
