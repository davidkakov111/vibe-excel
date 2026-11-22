import { expect, type Page } from "@playwright/test";

export async function signUp(page: Page, email: string, password: string) {
  const register = await page.request.post("/api/auth/register", {
    data: { email, password },
  });
  expect(register.status(), await register.text()).toBe(201);
}

export async function signIn(page: Page, email: string, password: string) {
  const csrfResponse = await page.request.get("/api/auth/csrf");
  expect(csrfResponse.ok()).toBeTruthy();
  const { csrfToken } = (await csrfResponse.json()) as { csrfToken: string };

  const callback = await page.request.post("/api/auth/callback/credentials", {
    form: {
      csrfToken,
      email,
      password,
      callbackUrl: "/dashboard",
    },
    maxRedirects: 0,
  });
  expect(
    callback.status(),
    await callback.text(),
  ).toBe(302);
}

export async function signUpAndOpenDashboard(
  page: Page,
  email: string,
  password: string,
) {
  await signUp(page, email, password);
  await signIn(page, email, password);
  await page.goto("/dashboard");
  await expect(page).toHaveURL(/\/dashboard/);
}

export async function createBlankSpreadsheet(page: Page) {
  const response = await page.request.post("/api/spreadsheets", {
    data: { name: "Untitled spreadsheet" },
  });
  const body = await response.text();
  expect(response.status(), body).toBe(201);
  const payload = JSON.parse(body) as { spreadsheet: { id: string } };
  await page.goto(`/spreadsheet/${payload.spreadsheet.id}`);
  await expect(page).toHaveURL(/\/spreadsheet\//);
}