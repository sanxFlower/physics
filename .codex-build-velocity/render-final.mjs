import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const deck = await PresentationFile.importPptx(await FileBlob.load(process.argv[2]));
await fs.mkdir(process.argv[3], { recursive: true });
for (let i = 0; i < deck.slides.items.length; i++) {
  const png = await deck.slides.getItem(i).export({ format: "png", scale: 1 });
  await fs.writeFile(`${process.argv[3]}/slide-${String(i + 1).padStart(2, "0")}.png`, new Uint8Array(await png.arrayBuffer()));
}
const snapshot = await deck.inspect({ kind: "slide,textbox,shape", include: "slide,title,textPreview,bbox", maxChars: 30000 });
await fs.writeFile(`${process.argv[3]}/inspect.ndjson`, snapshot.ndjson);
console.log(`Rendered ${deck.slides.items.length} slides`);
