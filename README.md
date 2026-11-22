# Local-First Excel Editor

A Dockerized Next.js spreadsheet editor with user authentication, PostgreSQL persistence, custom grid editing, and Excel (`.xlsx` / `.xls`) import/export.

## Architecture

- **Frontend + API**: Next.js 15 (App Router) with TypeScript and Tailwind CSS
- **Auth**: Auth.js (NextAuth v5) credentials provider, JWT sessions, bcrypt password hashing
- **Database**: PostgreSQL 16 via Prisma ORM (`User`, `Spreadsheet` models)
- **Excel I/O**: SheetJS (`xlsx`) for parsing and generating workbooks
- **Runtime**: Docker Compose (`app` + `db` services)

Spreadsheet data is stored as a versioned **workbook JSON** document (v2) in the `Spreadsheet.cells` column: multiple worksheets, cell inputs/formulas, formatting, dimensions, merges, charts metadata, and view settings. Legacy v1 flat cell maps are migrated automatically on read.

Formula evaluation uses **HyperFormula** (client-side) with dependency recalculation. Excel I/O uses **ExcelJS** for multi-sheet workbooks, formulas, and basic styles.

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for persistence, save behavior, and Excel fidelity limits.

## Quick start (Docker Compose)

```bash
docker compose up --build
```

Open [http://localhost:3000](http://localhost:3000), register an account, and use the dashboard to create or upload spreadsheets.

## Local development (without Docker app container)

```bash
cp .env.example .env
docker compose up -d db
npm install --legacy-peer-deps
npx prisma migrate dev
npm run dev
```

## Scripts

| Command | Description |
| --- | --- |
| `npm run dev` | Start Next.js dev server |
| `npm run build` | Production build |
| `npm start` | Start production server (standalone output) |
| `npm test` | Unit tests (Vitest) |
| `npm run test:e2e` | Playwright tests (requires running app + DB) |
| `npm run test:e2e:full` | Boot embedded Postgres, migrate, build, run Playwright |
| `docker compose up --build` | Run full stack |

## Security notes

- Users can only access spreadsheets scoped to their authenticated user ID.
- API routes and editor pages are protected by middleware.
- Change `NEXTAUTH_SECRET` for non-local deployments.