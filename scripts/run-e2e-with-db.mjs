import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import EmbeddedPostgres from "embedded-postgres";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");
const dataDir = path.join(root, ".tmp", "embedded-postgres");
const port = 5433;

async function runCommand(command, args, env = process.env) {
  await new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: root,
      env,
      stdio: "inherit",
    });
    child.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${command} exited with ${code}`));
    });
  });
}

const postgres = new EmbeddedPostgres({
  databaseDir: dataDir,
  port,
  user: "postgres",
  password: "password",
  persistent: false,
  onLog: () => {},
  onError: (message) => console.error(message),
});

let serverProcess;

async function main() {
  fs.mkdirSync(dataDir, { recursive: true });
  await postgres.initialise();
  await postgres.start();
  await postgres.createDatabase("excel_editor");

  const databaseUrl = `postgresql://postgres:password@127.0.0.1:${port}/excel_editor`;
  const env = {
    ...process.env,
    DATABASE_URL: databaseUrl,
    AUTH_SECRET: "e2e-test-secret",
    AUTH_URL: "http://127.0.0.1:3000",
    NEXTAUTH_SECRET: "e2e-test-secret",
    NEXTAUTH_URL: "http://127.0.0.1:3000",
    PORT: "3000",
  };

  await runCommand("npx", ["prisma", "migrate", "deploy"], env);
  await runCommand("npm", ["run", "build"], env);

  // Use `next start` so API routes get the full traced dependency tree (standalone
  // misses some transitive packages such as `generator-function`).
  serverProcess = spawn("npx", ["next", "start", "-p", "3000"], {
    cwd: root,
    env,
    stdio: "inherit",
  });

  await waitForServer("http://127.0.0.1:3000/login");
  await runCommand("npx", ["playwright", "test"], {
    ...env,
    PLAYWRIGHT_BASE_URL: "http://127.0.0.1:3000",
  });
}

async function waitForServer(url) {
  for (let attempt = 0; attempt < 60; attempt++) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // retry
    }
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }
  throw new Error("Server did not become ready in time");
}

async function cleanup() {
  if (serverProcess && !serverProcess.killed) {
    serverProcess.kill("SIGTERM");
  }
  try {
    await postgres.stop();
  } catch {
    // ignore
  }
}

process.on("SIGINT", () => {
  void cleanup().finally(() => process.exit(1));
});

main()
  .then(async () => {
    await cleanup();
    process.exit(0);
  })
  .catch(async (error) => {
    console.error(error);
    await cleanup();
    process.exit(1);
  });