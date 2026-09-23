import fs from "node:fs/promises";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";
const deck = await PresentationFile.importPptx(await FileBlob.load(process.argv[2]));
const snapshot = await deck.inspect({ kind: "slide", include: "id,slide,title", maxChars: 50000 });
console.log(snapshot.ndjson);
for (let i = 0; i < deck.slides.items.length; i++) {
  const preview = await deck.slides.getItem(i).export({ format: "png", scale: 1 });
  await fs.writeFile(`.codex-build-velocity/source-${String(i + 1).padStart(2, "0")}.png`, new Uint8Array(await preview.arrayBuffer()));
}
