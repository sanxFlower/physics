import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";
const p = await PresentationFile.importPptx(await FileBlob.load("C:/Users/San/workspace/physics/高中物理/光谷二高/26级-高一/1.3 位置变化快慢的描述 速度/1.3速度_课堂课件_原稿补充版.pptx"));
for (const n of [5,11,12,20,27,28,38,39,40,41]) {
  const s = p.slides.getItem(n-1);
  const png = await s.export({format:"png",scale:1});
  await fs.writeFile(`C:/Users/San/workspace/physics/.codex-build-1_3-final/rendered-${String(n).padStart(2,"0")}.png`, new Uint8Array(await png.arrayBuffer()));
}
