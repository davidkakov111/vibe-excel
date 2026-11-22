import ExcelJS from "exceljs";
import fs from "node:fs";
import { parseExcelWorkbook, workbookToExcelBuffer } from "../src/lib/excel/workbook-io.ts";

const path =
  "/home/davidkov111/work/personal/projects/ai-test/Untitled spreadsheet 1 1 (3).xlsx";

const originalBuf = fs.readFileSync(path);

const wb0 = new ExcelJS.Workbook();
await wb0.xlsx.load(originalBuf);
const ws0 = wb0.worksheets[0];
console.log("orig media", ws0?.model.media?.length ?? 0);

const doc = await parseExcelWorkbook(originalBuf);
const outBuf = await workbookToExcelBuffer(doc);
const wb1 = new ExcelJS.Workbook();
await wb1.xlsx.load(outBuf);
const ws1 = wb1.worksheets[0];
console.log("rt media", ws1?.model.media?.length ?? 0);

const img0 = wb0.getImage(0);
const img1 = wb0.getImage(1);
console.log("images", [
  img0 && { ext: img0.extension, len: img0.buffer?.length },
  img1 && { ext: img1.extension, len: img1.buffer?.length },
]);
console.log(
  "embedded in doc",
  doc.sheets[0]?.embeddedImages?.length,
  doc.sheets[0]?.embeddedImages?.map((i) => i.dataBase64.length),
);