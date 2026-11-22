import PDFDocument from "pdfkit";
import { FormulaEngine } from "@/lib/formula/engine";
import {
  DEFAULT_COL_WIDTH,
  DEFAULT_ROW_HEIGHT,
  normalizeWorkbookDocument,
} from "@/lib/workbook/defaults";
import { cellKey, indexToColumnLabel } from "@/lib/spreadsheet-types";
import type { WorkbookDocument, Worksheet } from "@/lib/workbook/types";
import { getCellInput } from "@/lib/workbook/types";

function usedRange(sheet: Worksheet): {
  endRow: number;
  endCol: number;
} {
  let endRow = 0;
  let endCol = 0;
  for (const key of Object.keys(sheet.cells)) {
    const [row, col] = key.split(":").map(Number);
    endRow = Math.max(endRow, row);
    endCol = Math.max(endCol, col);
  }
  for (const chart of sheet.charts ?? []) {
    endRow = Math.max(endRow, chart.anchor.endRow);
    endCol = Math.max(endCol, chart.anchor.endCol);
  }
  return {
    endRow: Math.min(Math.max(endRow, 20), sheet.rowCount - 1),
    endCol: Math.min(Math.max(endCol, 8), sheet.colCount - 1),
  };
}

function cellDisplay(
  engine: FormulaEngine,
  sheet: Worksheet,
  sheetId: string,
  row: number,
  col: number,
): string {
  const input = getCellInput(sheet, row, col);
  if (input && !input.startsWith("=")) return input;
  return engine.getDisplayValue(sheetId, row, col) || input;
}

export function workbookToPdfBuffer(document: WorkbookDocument): Promise<Buffer> {
  const doc = normalizeWorkbookDocument(document);

  return new Promise((resolve, reject) => {
    const engine = FormulaEngine.fromWorkbook(doc);
    const pdf = new PDFDocument({ size: "A4", layout: "landscape", margin: 36 });
    const chunks: Buffer[] = [];
    pdf.on("data", (chunk: Buffer) => chunks.push(chunk));
    pdf.on("end", () => {
      engine.destroy();
      resolve(Buffer.concat(chunks));
    });
    pdf.on("error", (error) => {
      engine.destroy();
      reject(error);
    });

    const margin = 36;
    const pageWidth = pdf.page.width - margin * 2;

    for (let sheetIndex = 0; sheetIndex < doc.sheets.length; sheetIndex++) {
      const sheet = doc.sheets[sheetIndex]!;
      if (sheetIndex > 0) pdf.addPage({ size: "A4", layout: "landscape", margin: 36 });

      pdf.font("Helvetica-Bold").fontSize(14).text(sheet.name);
      pdf.moveDown(0.35);

      const { endRow, endCol } = usedRange(sheet);
      const colCount = endCol + 1;
      const defaultColPt = (DEFAULT_COL_WIDTH * 0.75) / 4;
      let colWidths = Array.from({ length: colCount }, (_, col) => {
        const px = sheet.colWidths[String(col)] ?? DEFAULT_COL_WIDTH;
        return Math.max(28, (px * 0.75) / 4);
      });
      const totalWidth = colWidths.reduce((a, b) => a + b, 0);
      if (totalWidth > pageWidth) {
        const scale = pageWidth / totalWidth;
        colWidths = colWidths.map((w) => w * scale);
      }

      const rowHeight = Math.max(14, (DEFAULT_ROW_HEIGHT * 0.75) / 3);
      const headerH = rowHeight;
      let y = pdf.y;

      const colX = (col: number) => {
        let x = margin;
        for (let i = 0; i < col; i++) x += colWidths[i]!;
        return x;
      };

      pdf.font("Helvetica-Bold").fontSize(8);
      pdf
        .rect(margin, y, colWidths[0] ?? defaultColPt, headerH)
        .fillAndStroke("#f3f3f3", "#999");
      pdf.fillColor("#000").text("", margin, y);
      for (let col = 0; col <= endCol; col++) {
        const w = colWidths[col]!;
        const x = colX(col);
        pdf.rect(x, y, w, headerH).fillAndStroke("#f3f3f3", "#999");
        pdf.fillColor("#333").text(indexToColumnLabel(col), x + 2, y + 3, {
          width: w - 4,
          align: "center",
        });
      }
      y += headerH;

      const maxY = pdf.page.height - margin;
      for (let row = 0; row <= endRow; row++) {
        if (y + rowHeight > maxY) {
          pdf.addPage({ size: "A4", layout: "landscape", margin: 36 });
          y = margin;
        }
        const rh =
          Math.max(14, ((sheet.rowHeights[String(row)] ?? DEFAULT_ROW_HEIGHT) * 0.75) / 3) *
          0.85;

        pdf.font("Helvetica").fontSize(7);
        pdf
          .rect(margin, y, colWidths[0]! * 0.6, rh)
          .stroke("#999");
        pdf.fillColor("#555").text(String(row + 1), margin + 2, y + 2, {
          width: colWidths[0]! * 0.55,
          align: "center",
        });

        for (let col = 0; col <= endCol; col++) {
          const w = colWidths[col]!;
          const x = colX(col);
          const data = sheet.cells[cellKey(row, col)];
          const text = cellDisplay(engine, sheet, sheet.id, row, col);
          const fill = data?.format?.fillColor;
          if (fill) {
            pdf.rect(x, y, w, rh).fillAndStroke(fill, "#bbb");
          } else {
            pdf.rect(x, y, w, rh).stroke("#bbb");
          }
          pdf.fillColor("#000").font(data?.format?.bold ? "Helvetica-Bold" : "Helvetica");
          pdf.text(text.slice(0, 120), x + 3, y + 2, {
            width: w - 6,
            height: rh - 2,
            ellipsis: true,
            align:
              data?.format?.horizontalAlign === "center"
                ? "center"
                : data?.format?.horizontalAlign === "right"
                  ? "right"
                  : "left",
          });
        }
        y += rh;
      }
    }

    pdf.end();
  });
}