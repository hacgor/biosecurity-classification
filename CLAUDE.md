# CLAUDE.md — BioSecure Farm Classification System

## Project Overview

BioSecure is a Proof of Concept web application implementing a **3-Scale Selection Model** for classifying dairy cattle farms into biosecurity categories. The system is based on the Gent University biosecurity questionnaire (124 questions across 11 topics) and classifies farms as FC1 (Poor), FC2 (Medium), or FC3 (Good).

This is a **PoC-grade** application — prioritize working features and correctness of the classification algorithm over production polish.

## Tech Stack

- **Frontend**: React 18 (single-page app, no router)
- **Build**: Vite
- **Excel parsing**: SheetJS (`xlsx` package, imported as `import * as XLSX from "xlsx"`)
- **Persistence**: `window.storage` API in Claude artifacts; `localStorage` shim for standalone/GitHub Pages deployment
- **Hosting**: GitHub Pages via `gh-pages` package (`npm run deploy`)
- **No backend** — all computation client-side, all data in browser storage

## Architecture

### Single-File App (`src/App.jsx`)

The entire application lives in one file. This is intentional for PoC scope — do not split into multiple files unless explicitly asked. The file is structured in this order:

1. **Imports + storage shim** (localStorage polyfill for standalone mode)
2. **QUESTIONS array** — all 124 questionnaire items with `{ idx, num, text, topic }` where `idx` is the Excel column index (0-based)
3. **Index constants** — `SCORE_COLS`, `TOPIC_INDICES` (cols 138–148), `AGGREGATE_INDICES` (cols 149–151), `IBT_INDICES` (cols 138–148), `TOPICS` map (A–K)
4. **Theme object** — dark theme color tokens
5. **Classification algorithm** — `percentile()`, `classifyByPercentile()`, `runClassification()`
6. **UI components** — Badge, Button, Card, Input (generic)
7. **Feature components** — IBTConfig, SBMConfig, ExcelImport, ManualEntry, FarmList, ClassificationResults, ProcessDetail
8. **App root** — state management, persistent storage, routing via view state

### Data Model

```
Process {
  id: string              // generated via genId()
  name: string
  description: string
  created: ISO string
  farms: Farm[]
  sbmCriteria: SBMCriterion[]
  selectedIBT: number[]   // subset of IBT_INDICES (cols 138–148)
  results: ClassResult[]  // populated after running classification
  classified: boolean
}

Farm {
  id: string
  name: string
  answers: { [colIndex: number]: string }   // question answers keyed by Excel column
  scores: { [colIndex: number]: number }    // computed scores keyed by Excel column
  importedAt: ISO string
  updatedAt?: ISO string
}

SBMCriterion {
  questionIdx: number     // Excel column index
  questionNum: string     // display number like "8", "32", "77"
  questionText: string
  regexPattern: string    // validated against farm answers with /pattern/i
}
```

### Excel Column Mapping (Critical)

The Excel files use a flat row format. Column indices (0-based) are load-bearing:

- **Cols 0–137**: Questions and intermediate question scores
- **Cols 138–148**: Topic scores (A through K) — used in Scale 2
- **Col 149**: External biosecurity score — used in Scale 1
- **Col 150**: Internal biosecurity score — used in Scale 1
- **Col 151**: Total biosecurity score — used in Scale 1
- **Cols 152–157**: Legacy scoring columns (not used in our model)

Question score columns (non-question data interspersed among questions): 11, 36, 58, 73, 84, 101, 119, 126, 137. These are excluded from the QUESTIONS array.

Row 0 in Excel = question headers. Rows 1+ = farm data.

## Classification Algorithm — 3-Scale Selection Model

**This is the core of the application. Changes here must be verified carefully.**

### Scale 1 — Score Distribution
- Takes three scores per farm: Total (col 151), External (col 149), Internal (col 150)
- Computes each farm's **percentile within the set** of all farms in the classification process
- Classification = based on the **minimum** percentile across all three
- Thresholds: ≤25% → FC1, >25% and ≤75% → FC2, >75% → FC3

### Scale 2 — Important Biosecurity Topics (IBT)
- Uses **only the topics selected** by the expert (stored in `process.selectedIBT`)
- Available topics: A–K (cols 138–148) — **NOT External score (col 149)**
- Same percentile + threshold logic as Scale 1
- Classification = based on the **minimum** percentile across selected topics
- If no topics selected, Scale 2 is skipped (all farms pass)

### Scale 3 — Specific Biosecurity Measures (SBM)
- Expert defines criteria: question + **regex pattern** for valid answer
- Each farm's answer is tested with `new RegExp(pattern, "i").test(answer)`
- **Any single failed criterion** → farm fails Scale 3
- Failure → farm is downgraded to **FC1 regardless** of Scale 1 and 2 results

### Final Classification
1. Compute Scale 1 result and Scale 2 result independently
2. Combined = **worst** of Scale 1 and Scale 2 (FC1 < FC2 < FC3)
3. If Scale 3 fails → override to FC1

### Percentile Calculation
Uses midpoint method: values below get full credit, ties get half credit.
```javascript
function percentile(arr, val) {
  const sorted = [...arr].sort((a, b) => a - b);
  let count = 0;
  for (const v of sorted) {
    if (v < val) count++;
    else if (v === val) count += 0.5;
  }
  return (count / sorted.length) * 100;
}
```

## Coding Standards

### General
- All React components are **function components with hooks**
- State is managed with `useState`; no external state library
- No TypeScript — plain JSX
- No CSS files — all styling is **inline** via style objects
- Colors come from the `theme` object — never hardcode hex in components
- IDs generated with `genId()` (timestamp + random base36)

### Component Patterns
- **Inline editing pattern**: form state tracks `editIdx` (null = add mode, number = edit mode). One shared form UI, conditional save logic. See `SBMConfig` for reference.
- **Expand/collapse pattern**: `expandedId` state, click toggles, conditional render of detail section. See `FarmList` and `ClassificationResults`.
- **View routing**: `ProcessDetail` uses a `view` state string (`"overview"`, `"manual"`, `"editFarm"`, `"results"`) instead of a router. Never add react-router.
- **Edit vs. new**: `ManualEntry` accepts optional `existingFarm` prop. When present, pre-fills all fields and saves with the existing ID. The `isEdit` flag drives heading text and button labels.

### Data Flow
- All process data lives in the top-level `App` component's `processes` state
- `ProcessDetail` receives the current process + an `onUpdate` callback
- Child components call `updateProcess()` which triggers a state update that propagates to persistent storage
- Storage writes happen in a `useEffect` watching the `processes` array

### Excel Import
- Uses SheetJS `XLSX.read()` with `{ type: "array" }`
- Converts to array-of-arrays via `XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })`
- Row 0 = questions (ignored), rows 1+ = farms
- Columns ≥138 with numeric values go to `scores`, others go to `answers`

## Key Decisions and Rationale

| Decision | Rationale |
|----------|-----------|
| Single JSX file | PoC scope — minimal project structure, easy to iterate |
| No router | Only 3-4 "views" managed via state string — router is overkill |
| Dark theme | Domain users (agricultural biosecurity) often work in field conditions; dark theme reduces glare |
| Regex for Scale 3 | Exact string matching was too rigid for real-world answers that may have variable formatting (e.g., "cleaned: Always \| disinfected: Always") |
| Percentile-based classification | Relative ranking within a farm set, not absolute thresholds — a farm's class depends on the other farms in the same process |
| selectedIBT per process | Different classification rounds may focus on different biosecurity topics per expert judgment |
| localStorage shim | Allows the same codebase to run in Claude artifacts (window.storage) and standalone (localStorage) |
| GitHub Pages deployment | Free, simple, no backend needed for a PoC |

## Common Tasks

### Adding a new question
1. Add entry to `QUESTIONS` array with correct `idx` (Excel column), `num`, `text`, `topic`
2. If it has an associated score column, add the index to `SCORE_COLS`

### Changing classification thresholds
Edit `classifyByPercentile()` — currently ≤25 → FC1, ≤75 → FC2, >75 → FC3

### Adding a new topic
1. Add to `TOPICS` object
2. Add column index to `TOPIC_INDICES`
3. Add column index to `IBT_INDICES`
4. Update `SCORE_FIELDS` in the ManualEntry section

### Changing the base path for deployment
Edit `base` in `vite.config.js` to match your GitHub repo name

## Future Work (discussed but not yet implemented)
- jsonbin.io integration for shared cloud persistence
- Score computation logic (deriving topic/aggregate scores from raw question answers)
- Export classification results to Excel
- Python backend version for production deployment
- User authentication and role-based access
