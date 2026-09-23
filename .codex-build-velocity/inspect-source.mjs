import path from "node:path";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const sourcePath = process.argv[2];
const deck = await PresentationFile.importPptx(await FileBlob.load(sourcePath));
console.log(JSON.stringify({
  slideCount: deck.slides.items.length,
  masters: deck.masters.items.map(m => ({ id: m.id, name: m.name })),
  layouts: deck.layouts.items.map(l => ({ id: l.id, name: l.name, placeholders: l.placeholders.summary() })),
}, null, 2));
const snapshot = await deck.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes,layout",
  include: "id,slide,name,title,textPreview,textChars,textLines,bbox,bboxUnit,chartType,rows,cols,isPlaceholder,placeholders",
  maxChars: 100000,
});
console.log(snapshot.ndjson);
