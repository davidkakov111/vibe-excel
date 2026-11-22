import { expect, test } from "@playwright/test";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import * as XLSX from "xlsx";
import {
  createBlankSpreadsheet,
  signUpAndOpenDashboard,
} from "./auth-helpers";

const password = "test-password-123";
const unique = Date.now();

test.describe("spreadsheet workflows", () => {
  test("register, edit cells, persist, export, and isolate users", async ({
    page,
    browser,
  }) => {
    const email = `user-${unique}@example.com`;
    await signUpAndOpenDashboard(page, email, password);

    await createBlankSpreadsheet(page);

    await page.getByRole("textbox", { name: "Formula bar" }).click();
    await page.getByRole("textbox", { name: "Formula bar" }).fill("Hello Cell");
    await page.getByRole("textbox", { name: "Formula bar" }).press("Enter");
    await expect(page.getByText("All changes saved")).toBeVisible({ timeout: 15_000 });

    await page.reload();
    await expect(page.getByText("Hello Cell")).toBeVisible();

    const downloadPromise = page.waitForEvent("download");
    await page.getByRole("link", { name: "Download .xlsx" }).click();
    const download = await downloadPromise;
    const downloadPath = await download.path();
    expect(downloadPath).toBeTruthy();
    const fileBuffer = fs.readFileSync(downloadPath!);
    const workbook = XLSX.read(fileBuffer, { type: "buffer" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    expect(sheet.A1.v).toBe("Hello Cell");

    const tempXlsx = path.join(os.tmpdir(), `upload-${unique}.xlsx`);
    XLSX.writeFile(workbook, tempXlsx);

    await page.goto("/dashboard");
    await page.setInputFiles('input[type="file"]', tempXlsx);
    await page.getByRole("button", { name: "Upload and open" }).click();
    await expect(page).toHaveURL(/\/spreadsheet\//);
    await expect(page.getByText("Hello Cell")).toBeVisible();

    const otherContext = await browser.newContext({
      baseURL: process.env.PLAYWRIGHT_BASE_URL ?? "http://localhost:3000",
    });
    const otherPage = await otherContext.newPage();
    const otherEmail = `other-${unique}@example.com`;
    await signUpAndOpenDashboard(otherPage, otherEmail, password);
    await expect(otherPage.getByText("Untitled spreadsheet")).toHaveCount(0);

    const spreadsheetUrl = page.url();
    await otherPage.goto(spreadsheetUrl);
    await expect(otherPage.getByText(/not found|Unauthorized|404/i)).toBeVisible({
      timeout: 10_000,
    });

    await otherContext.close();
  });
});