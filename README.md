# Vibe-coded Excel editor

This is a vibe-coded Excel editor built with Cursor, where you can create or import XLSX workbooks, edit them, collaborate with others in real time, and export them again — offering a somewhat smaller feature set than Excel while still supporting many commonly used spreadsheet features.

![Spreadsheet editor: grid, ribbon, comments, charts, and collaboration indicators](docs/images/editor-screenshot.png)

*The editor in use: multi-sheet grid, formatting ribbon, threaded comments, chart overlays, conditional formatting, version history, and real-time collaboration.*

---

## What this application is

A **Next.js** spreadsheet editor with **PostgreSQL** persistence: workbooks live as JSON in the database, with **ExcelJS** for `.xlsx` import/export. In practice it supports multi-sheet editing, formulas (**HyperFormula**), formatting, sort/filter, charts, comments, conditional formatting (partial), sharing, real-time collaboration over ops + SSE, and version history—with **Vitest** / light **Playwright** coverage.

It is **not** production-ready or full Excel parity (e.g. PDF is grid-only, no chart images; CF and sort interact oddly; single-instance collab). Treat unverified behavior as experimental. Stack: TypeScript, Prisma, Docker Compose (migrations on app startup). More detail: [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

---

## The purpose of this experiment

As a software engineer, I normally don't let AI take complete ownership of a codebase. Even when AI does most of the implementation, I review the code, verify its architecture and correctness, and inspect changes before committing them. Reliability is not something I want to compromise.

For this repository I wanted to explore a different approach: **pure vibe coding**.

I deliberately avoided reviewing the generated code. Instead, I described the features I wanted and let the Cursor coding agent implement them. I chose a relatively large and complex application — essentially recreating a smaller version of Excel from scratch — to see how far that approach could go.

The objective was to understand what AI can accomplish autonomously, where it starts struggling, and how much human involvement is actually necessary to produce a working application.

**Rough effort on my side:** The whole exercise took about **8 hours** of my time. Most of that was not dreaming up features—it was **bug-and-gap iterations**: using the editor, noticing what was wrong or missing, describing it to the agent, and repeating until that slice felt acceptable. The AI did the bulk of the typing; I supplied the acceptance pressure.

---

## What I learned from the experiment

These observations are from **this project and my workflow**. They are not universal claims about every tool or team.

### AI struggles with autonomous development of larger applications

In my experience, AI alone was not able to produce a reliable application of this complexity without substantial human feedback.

As the application grew, bugs and implementation gaps accumulated. Without manually testing the application and reporting issues, many features would have remained broken or only partially implemented.

### Human feedback makes AI significantly more effective

The most effective workflow I discovered was:

1. Ask the AI to implement a feature.  
2. Test the feature manually through the UI.  
3. Identify functional, visual, architectural, and integration problems.  
4. Describe the observed problems to the AI.  
5. Let the AI investigate and fix them.  
6. Repeat until the feature behaves as expected.

The AI was generally capable of fixing reported problems, often within one or a few iterations. This feedback loop allowed us to build a much larger and more functional application than autonomous implementation alone would have produced—and it accounts for most of those ~8 hours above.

### The reliability problem

Even though the AI wrote tests, ran builds, and used browser tooling, significant gaps remained.

Some problems I encountered included:

- Features that worked individually but failed when combined.  
- Data and metadata becoming disconnected during operations such as sorting.  
- Missing functionality in certain workflows (e.g. PDF without charts).  
- Incomplete edge-case handling.  
- Small details being overlooked.  
- Fixes for one feature introducing regressions in another.  
- Architectural inconsistencies that only became visible as the app grew (e.g. dual save paths, large editor hook orchestrating every feature).

Automated tests, linting, and browser checks are useful, but they did not guarantee complete functional correctness for me. AI-generated tests can share the same blind spots as the implementation.

The result was an application that could become **usable** through repeated human testing and AI corrections, but a codebase I would **not trust** the same way as one developed with continuous engineering ownership and code review.

### The reliability advantage of incremental engineering ownership

One of the most reliable approaches I know is for an engineer to maintain an independent understanding of the codebase: review changes in small increments, understand data flows, check critical paths, and ensure new features fit the existing system.

That builds a mental model over time, surfaces architectural problems earlier, and catches edge cases before they pile up. It is slower, because human review and judgment remain necessary; in that workflow, AI accelerates implementation rather than replacing ownership.

---

## Finding the right balance

There is no single optimal workflow for every project. Oversight should match complexity, reliability requirements, and the cost of failure.

### Pure vibe coding

- AI handles implementation, testing, and iteration.  
- Humans describe requirements and exercise the UI.  
- Fits prototypes, MVPs, experiments, and low-risk work.  
- Fast, but **reliability and maintainability should not be assumed**.

### AI-assisted development with selective engineering review

- AI implements features and tests.  
- An engineer reviews architecture, data flows, critical changes, and seams (persistence, auth, collaboration, export).  
- For less critical areas, summaries plus behavioral testing may suffice.  
- Users still validate features end-to-end before accepting them.  
- A practical middle ground when speed matters but some reliability is required.

### Engineering-led development with AI assistance

- AI helps with implementation, tests, debugging, and refactoring.  
- An engineer incrementally owns the codebase and reviews integration and edge cases.  
- Appropriate when reliability, security, maintainability, or failure cost demand stronger oversight.

AI still helps in strict reliability environments; the difference is **how much autonomy and verification** you allow, not whether you use AI at all.

---

## Final conclusion

Vibe coding is a powerful way to build software quickly, especially for MVPs, prototypes, and low-risk applications. It should not be confused with reliable software engineering.

AI can dramatically reduce implementation time, but it does not by itself guarantee architectural consistency, complete functionality, or long-term maintainability.

The amount of engineering oversight should be a **deliberate choice** based on complexity and reliability needs. The goal is not maximum AI autonomy at all costs, but an efficient balance between implementation speed and human judgment.

---

## Running the project

### Docker Compose (recommended)

```bash
docker compose up --build
```

Open [http://localhost:3000](http://localhost:3000), register, and use the dashboard to create or upload spreadsheets. Database migrations run when the **app** container starts (after the database health check).

### Local development (app on host, DB in Docker)

```bash
cp .env.example .env
docker compose up -d db
npm install --legacy-peer-deps
npx prisma migrate dev
npm run dev
```

### Useful commands

| Command | Description |
|--------|-------------|
| `npm run dev` | Next.js dev server |
| `npm run build` | Production build |
| `npm test` | Vitest unit tests |
| `npm run test:e2e` | Playwright (requires running app + DB) |
| `npm run test:e2e:full` | Embedded Postgres, migrate, build, Playwright |
| `docker compose up --build` | Full stack |

For non-local deployments, set strong `AUTH_SECRET` / `NEXTAUTH_SECRET` values and treat collaboration/export features as unvalidated until you run your own acceptance tests.