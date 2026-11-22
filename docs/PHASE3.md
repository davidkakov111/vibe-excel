# Phase 3 — status report

## Implemented

### Architecture
- **Optimistic concurrency**: `Spreadsheet.revision` + granular `POST /api/spreadsheets/[id]/ops`.
- **Spreadsheet access layer** (`src/lib/spreadsheet-access.ts`): owner / VIEW / EDIT.
- **Real-time hub** (`src/lib/collaboration/hub.ts`): in-process SSE fan-out (single Node instance).
- **Granular ops** (`src/lib/workbook/ops.ts`): `setCell`, `setFormat`, `setComment`.

### Sharing
- `SpreadsheetShare` model + migration `20251122100000_phase3_sharing_revision`.
- APIs: `GET/POST/DELETE /api/spreadsheets/[id]/shares` (owner only).
- Dashboard: **Your spreadsheets** + **Shared with you**.
- Editor **Share** dialog (invite by email, VIEW/EDIT, revoke).

### Collaboration
- `GET /api/spreadsheets/[id]/stream` (SSE).
- `POST /api/spreadsheets/[id]/presence` (selection broadcast).
- Client `useCollaboration` + editor uses **ops** for cell/format/comment when `canEdit`.
- Conflict: **409** + full workbook resync (no full-sheet broadcast on every keystroke).
- Collaborator list + selection text in grid header.

### Comments
- `CellComment` on `CellData`; red marker on cells.
- **Shift+Alt+M** to add/edit; delete clears own comment.
- Excel **notes** import/export via ExcelJS.

### Excel fidelity (incremental)
- Format-only cells, merges, row heights, borders, number formats on export.
- **Excel tables** + **conditional formatting** import/export (`Worksheet.tables`, `conditionalFormats`).
- **Theme colors** resolved on import (`src/lib/excel/theme-colors.ts`).
- **Formatted clipboard**: `text/plain` TSV + `application/vnd.ai-test.cells+json` (values + formats).
- Runtime **conditional format** styling in the grid (`src/lib/workbook/conditional-format.ts`).
- See `src/lib/excel/workbook-io.test.ts`.

## Tests run (this session)

| Command | Result |
|---------|--------|
| `npm test` | **29/29** passed (after Phase 3 Excel features) |
| `npm run build` | **Success** (after lint fixes) |
| `npx prisma generate` | **Success** |
| DB migrate | Requires running Postgres (`npx prisma migrate deploy`) |
| `npm run test:e2e:full` | **Not re-run** in this session (needs DB + Playwright) |

## Limitations / next work

| Area | Limitation |
|------|------------|
| Real-time | In-memory hub only; **multiple app replicas** need Redis (or similar). |
| Collaboration | Bulk actions (sort, structure) still use full **PATCH**; should become ops. |
| Presence | Text indicator only, not colored range overlays on grid yet. |
| Excel (remaining) | Native chart objects, color scales / data bars / icon sets CF, table stripe themes in UI. |
| Comments | No threaded replies; Excel author metadata not preserved on import. |
| Auth sharing | Invite requires user already registered (email lookup). |

## Deploy note

After pull, run:

```bash
npx prisma migrate deploy
docker compose up -d --build
```