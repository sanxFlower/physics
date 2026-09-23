import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const sourcePath = process.argv[2];
if (!sourcePath) throw new Error("Source path required");

const presentation = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
const snapshot = await presentation.inspect({
  kind: "deck,slide,textbox,shape,image,table,chart,notes,layout",
  include: "id,slide,title,textPreview,textChars,textLines,bbox,name,layoutId,placeholders",
  maxChars: 200000,
});
console.log(snapshot.ndjson);
console.log(JSON.stringify({
  slideCount: presentation.slides.items.length,
  slideSize: presentation.slides.items[0]?.frame,
  masters: presentation.masters.items.map(m => ({ id: m.id, name: m.name })),
  layouts: presentation.layouts.items.map(l => ({ id: l.id, name: l.name, placeholders: l.placeholders.summary() })),
}, null, 2));

for (let index = 0; index < presentation.slides.items.length; index += 1) {
  const slide = presentation.slides.getItem(index);
  const png = await slide.export({ format: "png", scale: 1 });
  await fs.writeFile(`C:/Users/San/workspace/physics/.codex-build-1_3-final/source-slide-${String(index + 1).padStart(2, "0")}.png`, new Uint8Array(await png.arrayBuffer()));
}
