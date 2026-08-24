import { writeFile, mkdir } from "node:fs/promises";
import { basename, join } from "node:path";
import {
  FileBlob,
  SpreadsheetFile,
} from "file:///C:/Users/Max/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const project = "C:/Users/Max/Downloads/OverallChancado_ENTREGA_3PM/OverallChancado";
const outputDir = join(project, "outputs", "previews");
await mkdir(outputDir, { recursive: true });

const workbooks = [
  ["CVB003_FAJA_DEMO.xlsx", "REPORTE DE INSPECCION CV0003"],
  ["CVB003_POLEAS_CAMPANA_P02.xlsx", "Hoja1"],
  ["CVB003_LIFE_SHAFT_CAMPANA_LS03.xlsx", "Hoja1"],
];

for (const [filename, sheetName] of workbooks) {
  const source = join(project, "outputs", filename);
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
  const rendered = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 0.22,
    format: "png",
  });
  const target = join(outputDir, `${basename(filename, ".xlsx")}_${sheetName}.png`);
  await writeFile(target, Buffer.from(await rendered.arrayBuffer()));
  process.stdout.write(`${target}\n`);
}
