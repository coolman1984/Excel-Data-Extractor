# xlextract — Authoritative Implementation Plan (V1)

> Phase: **architecture + planning**. No production implementation in this document.
> Audience: Opus-level reviewers, Sonnet 5 implementers, Codex Terra implementers.
> Status: architecture approved after final corrections; implementation begins with T-001 only. Supersedes nothing. Amendments require an Opus-level architecture review (see §23).

**Working name:** `xlextract` (package), `xlx` (CLI alias).
**Contract family:** `xlextract.contract/1`, `xlextract.manifest/1`.

---

## 0. Architecture thinking (evidence before decisions)

This section is the reasoning the rest of the plan rests on. It answers the twelve questions in the brief before any task decomposition.

### 0.1 The irreducible product

```
Excel file → understand physical structure → safely extract selected data
           → prove what was extracted → publish a stable machine-readable result
```

Everything in this plan must trace back to one of those four verbs. A component that serves none of them is out.

### 0.2 The three truths V1 must prove early

Most Excel tools die on one of these. V1 attacks all three in the first four phases rather than deferring them:

1. **Fidelity.** Getting *bytes out* is easy; getting *trustworthy values* out is not. Leading-zero IDs, 1904 date systems, percentage number formats, cached-vs-live formulas, Excel error sentinels, locale decimals. A tool that silently coerces `"007"` to `7` is worse than no tool, because it is wrong *confidently*.
2. **Memory reality.** The fastest Python reader available (calamine) materializes an entire sheet. Verified: ~400 MB resident for an 80 MB workbook; openpyxl's read-only iterator does the same file in ~40 MB at 5–6× the wall time. There is no library that is both fastest and bounded. Therefore adapter choice must be a **measured runtime decision against a memory budget**, not a project-wide library preference. This is the single most consequential finding of the research pass.
3. **Protected enterprise reality.** Direct parsing cannot read an IRM/DRM-protected workbook, by design and lawfully so. Only an Excel instance authenticated as the user can; the tool prefers a dedicated owned instance and treats access to an existing user session as explicit, read-only compatibility mode. This path is real, unavoidable for the target environment, and *cannot be proven by a synthetic fixture*.

### 0.3 Structure discovery vs business interpretation

The hard-won boundary:

| Plane | Owns | May assert | Example |
|---|---|---|---|
| **Structural** | physical facts read from the file | "this is true of the bytes" | sheet `Data` is very-hidden; A1:F980 is a defined Excel Table named `tbl_stock`; row 7 has 6 non-empty cells |
| **Interpretive** | decisions about what to extract | "this is a decision, here is its provenance" | row 7 is the header row; rows 981–983 are totals, excluded |

**Rule:** every interpretive decision is recorded in the manifest with `basis ∈ {contract, excel_table, named_range, configured, discovered, user_confirmed}`. Discovery is allowed to *act*, never to act *silently*, and is never allowed to *guess between equals* — genuine ambiguity is a terminal `NEEDS_INPUT`, not a coin flip. This is the design answer to "safe rather than magical".

### 0.4 Dependency decisions

Evaluated against: works offline, works as a standard user, no admin install, no Node/Docker, small and dependable, agent-maintainable.

| Candidate | Decision | Reason |
|---|---|---|
| **Python 3.12 only** | **ADOPT** | Only ecosystem with both a first-class Excel COM story (pywin32) and mature parsers. V1 intentionally does not support Python 3.11, reducing the compatibility and test surface. |
| **python-calamine** ≥0.8 | **ADOPT** (primary reader) | One dependency reads `.xlsx/.xlsm/.xlsb/.xls`. Rust-fast. Gives sheet metadata incl. hidden/very-hidden. **Known limit: materializes the sheet** — governed by the memory budget, never used blind. |
| **openpyxl** ≥3.1 | **ADOPT** (streaming + metadata) | The only bounded-memory path for huge `.xlsx/.xlsm` (`read_only=True, values_only=True`). Also the richest source of structural metadata: Tables, defined names, merged ranges, per-row hidden flags, `reset_dimensions()` to defeat phantom dimension records, and formula strings (`data_only=False`). |
| **pyarrow** ≥16 | **ADOPT** | Parquet writer with explicit row groups = our chunk boundary. Arrow types preserve fidelity (decimal, large_string, date32/timestamp) far better than CSV. Pure wheel, no admin. |
| **pydantic** v2 | **ADOPT** | Contracts are the product's public surface; we need validation *and* generated JSON Schema for non-Python consumers. |
| **typer** (+click) | **ADOPT** | Professional CLI ergonomics, real `--help`, subcommand groups, shell completion. |
| **rich** | **ADOPT** | Human rendering + progress. Already a typer dependency. Must be behind a renderer port so `--json` never touches it. |
| **pywin32** | **ADOPT** (Windows extra) | Required for the COM path. Direct `win32com.client` only. |
| **msoffcrypto-tool** | **ADOPT** (extra) | Password-protected OOXML/legacy → decrypt to an in-memory stream. Actively maintained (Jan 2026). Handles password protection **only**; not IRM. |
| **textual** | **ADOPT** (extra, Phase 7) | TUI over the same app services. Optional install; core never imports it. |
| **import-linter** | **ADOPT** (dev) | Makes §11 architecture rules mechanically enforced in CI, not aspirational prose. |
| **pandas** | **REJECT** (core) | Heavy, and its type coercion is actively hostile to fidelity (silent `object`→`float64`, NaN-ing of missing ints, date guessing). May appear in `tests/` and benchmarks only. |
| **polars** | **REJECT** (V1 core) | Excellent, but adds a large dependency for a job pyarrow already does at our boundary. Reconsider only if benchmarks show pyarrow write is the bottleneck. |
| **fastexcel** | **POSTPONE** | Arrow-native calamine binding; attractive for a future zero-copy path. Same materialization limit as calamine, so it buys speed, not safety. Re-evaluate at V1.1 with benchmark evidence. |
| **xlrd / pyxlsb** | **REJECT** | calamine covers `.xls` and `.xlsb`. Two fewer dependencies. |
| **xlwings** | **REJECT** | Convenience wrapper over the COM layer we must control precisely (session ownership, no-kill policy, retry semantics). Wrapping it would hide exactly the lifecycle we must own. |
| **SQLite / any DB** | **REJECT** (V1) | Run state lives in the filesystem run directory. No hidden global state, no migration burden, trivially inspectable by humans and agents. |
| **Any network/telemetry lib** | **FORBIDDEN** | See §17. Enforced by a test that blocks `socket` (§14). |
| **PyInstaller** | **ADOPT** (early smoke + release proof) | A portable Windows 10/11 one-dir build is mandatory. Smoke it immediately after the first working `doctor`/`inspect` slice, then repeat the proof on a clean standard-user machine for release. |

### 0.5 Where direct parsing beats COM, and where it does not

| Situation | Path | Why |
|---|---|---|
| Any readable `.xlsx/.xlsm/.xlsb/.xls` | **Direct parser** | 10–100× faster, no Excel required, no session risk, deterministic, testable in CI |
| Huge `.xlsx/.xlsm` over memory budget | **Direct streaming (openpyxl)** | Only bounded-memory option |
| Password-protected, password supplied | **Direct + msoffcrypto** | Decrypt in memory; source untouched |
| IRM/DRM/label-protected; Excel can open it, parsers cannot | **COM (explicit opt-in)** | Prefer a dedicated Excel instance owned by the tool; attached-user-session mode is a separately requested, effectively read-only compatibility mode. Neither is a bypass — both are authorized reads. |
| `.xlsb` over memory budget | **COM, else honest failure** | No streaming xlsb parser exists. We say so with a real error code rather than OOM-ing. |
| Values only obtainable by recalculation | **Out of scope V1** | We do not evaluate formulas. We report missing cached values. |

**COM is never a silent fallback.** A run that would need COM without `--allow-excel` terminates `NEEDS_INPUT` with the exact command to re-run.

### 0.6 What V1 postpones (explicitly)

Formula evaluation; cell formatting/styles/colors as extractable data; charts, images, pivot caches; `.ods`; writing Excel files; multi-process parallel sheet extraction; incremental/delta extraction; a daemon or server mode; any scheduler; resumable mid-file restart (see §13.6 for why this is deliberate).

---

## 1. Product Definition

**What it is.** `xlextract` is a Windows-first, offline, local command-line tool (plus an optional TUI on the same core) whose single job is to turn Excel workbooks into trustworthy structured data with proof. It reads a workbook, describes its physical structure, extracts a declared or discovered rectangular dataset, and publishes an immutable **run directory** containing the data, its schema, its lineage, an accounting ledger that reconciles every source row, and a manifest describing every decision made.

**What it is not.** Not a downloader, not a scraper, not a cleaner, not a validator of business rules, not a warehouse, not a dashboard, not an analytics engine, not a scheduler, not an orchestrator, not an AI interpreter of business meaning, not an Excel writer, not a DRM bypass. It does not know what a "material document" or a "KPI" is and must never learn.

**Who uses it.**
- *Primary:* a corporate data/automation engineer on Windows 11, standard user rights, driving the CLI or scripting it.
- *Secondary:* a non-technical employee running the TUI to extract a file someone asked them for.
- *Tertiary:* other tools and AI coding agents consuming `--json` output, exit codes, and the run manifest without knowing anything about Excel.

**Inputs.** One or more existing Excel files (`.xlsx`, `.xlsm`, `.xlsb`, `.xls`), optionally an extraction contract (YAML/JSON), optionally a password, optionally CLI overrides.

**Outputs.** A run directory: dataset(s) as Parquet (+ optional CSV/JSONL), a schema document, a manifest (identity, decisions, lineage anchors, timings, versions), a reconciliation ledger, an issues stream, a human preview, and a structured log. Plus a structured result on stdout and a stable exit code.

**Why it exists.** Because in enterprise reality, "just read the Excel file" fails in a hundred specific ways, and most tools fail *silently* — dropping rows, mangling identifiers, inventing dates. The value of this tool is not that it reads Excel. It is that **it can prove what it read, and refuses to be confidently wrong.**

---

## 2. Core Product Boundary

**CORE** (must exist for the mission; protected from growth)
Source identity & safety · format detection · workbook structure model · structure discovery · extraction planning · value typing/normalization · row ledger & reconciliation · lineage anchors · run state machine · atomic publication · error taxonomy · result/manifest contracts · CLI.

**OPTIONAL** (removable; the tool remains useful without any one of them)
COM/Excel adapter · encrypted-workbook adapter · `.xlsb`/`.xls` format support · CSV & JSONL sinks · TUI · `doctor` · benchmark harness.

**REQUIRED DISTRIBUTION** (not a product feature): portable Windows 10/11 PyInstaller one-dir build. It is smoke-proven early and release-proven on a clean standard-user machine offline.

**EXTERNAL** (integrates via published contracts; never imported, never embedded)
Downloaders (SAP/G-MES/web) · cleaners/transformers · business validators · history stores · databases · analytics · dashboards · orchestrators · AI agents.

**FUTURE** (designed *for*, not built)
Additional sinks (Arrow IPC, DuckDB file) · `.ods` · contract registry · delta extraction · parallel multi-sheet.

**FORBIDDEN** (rejecting these is a load-bearing architectural commitment)
Any file download · any business-meaning inference · any dashboard/report rendering · any scheduling · any network call · any telemetry · sending source data to an AI service · modifying, saving or deleting a source workbook · weakening Windows/Office/DRM/AV/execution-policy settings · killing Excel processes the tool does not own · a central application that "holds the tools together".

---

## 3. Architecture Decision

### 3.1 Chosen architecture — Hexagonal (ports & adapters) with a thin, explicit application layer

```
        ┌──────────────┐        ┌──────────────┐
        │     CLI      │        │     TUI      │     (presentation: no logic)
        └──────┬───────┘        └──────┬───────┘
               └──────────┬────────────┘
                     ┌────▼─────┐
                     │   app    │   use-cases: Inspect, Extract, Verify, Explain, Doctor
                     └────┬─────┘   owns run lifecycle; the ONLY caller of ports
                     ┌────▼─────┐
                     │   core   │   domain: workbook model, discovery, planner, typing,
                     │          │   ledger, run state, PORT INTERFACES
                     └────┬─────┘
             ┌────────────┼─────────────┐
        ┌────▼───┐   ┌────▼────┐   ┌────▼────┐
        │adapters│   │  sinks  │   │  fsys   │   (infrastructure: implement ports)
        └────────┘   └─────────┘   └─────────┘
             ▲
        ┌────┴───────────────────────────────┐
        │ contracts (pydantic + JSON Schema) │  depends on NOTHING in the project
        └────────────────────────────────────┘
```

Dependency direction is strictly inward: `cli|tui → app → core → contracts`; `adapters|sinks → core (ports) + contracts`. `core` never imports an adapter, a sink, `app`, `cli`, or `tui`. Composition happens once, in `app.wiring`.

**Why this and not something simpler.** The two riskiest parts of the system — the Excel COM lifecycle and the streaming reader — are exactly the parts we most need to be able to replace, disable, test in isolation, and keep from leaking. Ports make "Excel COM implementation may not leak across the project" a compile-time-checkable fact rather than a code-review hope. The cost (one indirection) is small; the benefit (optional capabilities that are genuinely removable, and adapters that agents can implement from a contract without reading the core) is exactly what the brief asks for.

**Why a thin app layer and not fat CLI commands.** CLI and TUI must share one core with no duplicated behavior. If the use-case lives in the Click callback, the TUI must either duplicate it or import the CLI. The app layer makes "no hidden TUI-only behavior" structural.

### 3.2 Rejected alternatives

| Alternative | Rejected because |
|---|---|
| **Layered monolith** (`readers/ transforms/ writers/` flat) | Nothing prevents the CLI from importing an adapter; the COM dependency spreads; optional features stop being optional. Cheapest to start, most expensive at exactly the moment we add COM. |
| **Plugin architecture with entry-point discovery** | Speculative extensibility. Adds dynamic loading, versioning, and trust problems for a V1 with four known adapters. Ports already give replaceability. *(POSTPONE, not reject forever.)* |
| **Everything through Excel COM** | Slower by 1–2 orders of magnitude; requires Excel installed; untestable in CI; fragile (modal dialogs, busy states); uses the user's own Excel session for work that needs no rights. |
| **Everything through direct parsing** | Cannot read IRM/DRM-protected enterprise workbooks — the exact files the target environment contains. |
| **pandas DataFrame as the internal data model** | Type coercion destroys the fidelity the product exists to protect; forces the whole dataset into memory; couples our public schema to pandas dtypes. |
| **Library-first (no CLI), CLI as afterthought** | The brief and the users demand a first-class CLI; the Python API is the byproduct, not the reverse. The core *is* importable, but the CLI is the supported contract. |
| **Event-sourced / message-bus internals** | Enormous ceremony for a single-process batch tool. Rejected as infrastructure "merely useful someday". |
| **Database-backed run store (SQLite)** | Filesystem run directories are inspectable, portable, atomic-publishable, and need no migrations. |

### 3.3 The three load-bearing invariants

Every design choice below is downstream of these. Violating one is a stop-the-line event.

- **I1 — Source is evidence.** The tool opens sources read-only and never writes to, renames, moves, or deletes a source path. Proven by a test fixture that hashes every corpus file before and after the *entire* suite.
- **I2 — No silent loss.** Every observed source row is accounted for in exactly one ledger bucket, and the reconciliation identity must hold or the run does not succeed.
- **I3 — No silent decision.** Every interpretive choice appears in the manifest with its basis. Ambiguity is a terminal state, not a guess.

---

## 4. System Map (Project Eye, five levels)

**LEVEL 1 — SYSTEM.** One offline Windows CLI/TUI process. Input: Excel files (+ optional contract). Output: an immutable run directory + structured stdout + exit code. No network, no daemon, no shared state between runs. Composes with external tools only through the run directory and exit codes.

**LEVEL 2 — DOMAINS.**
| Domain | Responsibility | Owner of |
|---|---|---|
| `contracts` | the public, versioned data language | schema definitions |
| `source` | identity, safety, format detection, staging | source bytes & hashes |
| `structure` | the physical workbook model | workbook facts |
| `discovery` | candidate datasets and header strategies | interpretive proposals |
| `extraction` | planning, reading, typing, chunking | the row stream |
| `evidence` | ledger, lineage, issues, reconciliation | proof |
| `publication` | run directory, atomic publish, sinks | outputs |
| `orchestration` | run lifecycle, state, cancellation | run state |
| `presentation` | CLI & TUI rendering | nothing but pixels/text |
| `diagnostics` | environment readiness | nothing |

**LEVEL 3 — COMPONENTS.** See §5.

**LEVEL 4 — CODE.** See §10 repository tree. Each `graph.yaml` node maps to exactly one module path.

**LEVEL 5 — RUNTIME.** Seven named journeys (§6). Each run manifest records the journey and the stage sequence actually executed; `eye check --runtime` compares that against the declared journey to detect runtime drift.

---

## 5. Domain / Component Map

Format per component: **purpose · owner · inputs · outputs · dependencies · public contract · state ownership · tests**.

### 5.1 `contracts` — `src/xlextract/contracts/`
Purpose: define every externally visible data structure (extraction contract, manifest, schema doc, ledger, issue, result envelope) and emit JSON Schema. · Owner: `contracts`. · In: nothing. · Out: pydantic models + `schemas/*.json`. · Deps: **pydantic only**. · Public contract: *itself* — this is the product's API. · State: none (pure). · Tests: round-trip, schema-snapshot (golden), backward-compat matrix, example-validation.

### 5.2 `source.identity` — `core/source/identity.py`
Purpose: resolve a path to a stable source identity. · In: path. · Out: `SourceRef{abs_path, long_path_normalized, size, mtime_ns, sha256, sniffed_format, extension, extension_matches_content}`. · Deps: hashlib, os. · Contract: `SourceRef` in `contracts`. · State: owns the source hash — **sole owner**; nothing else hashes sources. · Tests: hash stability, `\\?\` long-path handling, extension/content mismatch (a `.xls` that is really OOXML), file-in-use, file-still-being-copied (size/mtime settling probe), UNC paths.

### 5.3 `source.safety` — `core/source/safety.py`
Purpose: enforce **I1**. Provide read-only handles; decide when a staging copy is required (COM path, file on a volatile share, encrypted decrypt-to-memory); verify the source is unchanged at end of run (`E_SRC_CHANGED_DURING_RUN`). · State: owns the staging temp directory lifecycle. · Tests: no-write proof, staging cleanup on crash, mid-run mutation detection, permission-denied classification.

### 5.4 `source.probe` — `core/source/probe.py`
Purpose: cheap pre-read sizing so adapter selection is a *measurement*. For OOXML: read the zip central directory and sum uncompressed sheet XML sizes (no decompression). For xlsb/xls: file size heuristics + format capability table. · Out: `SizeEstimate{per_sheet_bytes, est_cells, confidence}`. · Tests: estimate-vs-actual error band on the corpus (must be within the band that keeps budget decisions correct).

### 5.5 `structure.model` — `core/structure/model.py`
Purpose: the format-neutral workbook fact model: `WorkbookFacts{date_system, sheets[], tables[], defined_names[], has_macros, has_external_links, is_encrypted, protection_flags}`; `SheetFacts{name, index, visibility, used_range, dimension_declared, dimension_trusted, merged_ranges, hidden_rows, hidden_cols, autofilter}`. Pure data. · Tests: model invariants only.

### 5.6 `structure.ports` — `core/structure/ports.py`
Purpose: the port interfaces every adapter implements — `WorkbookInspector`, `RowReader`, `CapabilityReport`. This is the boundary the whole hexagon rests on. · Contract: documented protocol + a shared **adapter conformance test suite** every adapter must pass. · Tests: the conformance suite itself.

### 5.7 `adapters.calamine` — `adapters/calamine_adapter.py`
Purpose: fast path for all four formats; primary inspector. · Constraint: **materializes a sheet** — may only be selected when `est_bytes ≤ memory_budget × safety_factor`. · Out: `WorkbookFacts`, bulk row batches. · Tests: conformance suite + format matrix + a memory-ceiling test that asserts the guard refuses oversize sheets.

### 5.8 `adapters.openpyxl_stream` — `adapters/openpyxl_adapter.py`
Purpose: bounded-memory streaming for `.xlsx/.xlsm`; structural metadata enrichment (Tables, defined names, merged ranges, per-row hidden flags, formula strings); `reset_dimensions()` to defeat phantom `UsedRange`. · Tests: conformance suite, peak-RSS ceiling test, phantom-dimension fixture, hidden-row flag fidelity, `data_only` cached-value-absent detection.

### 5.9 `adapters.excel_com` — `adapters/excel_com/`
Purpose: authorized read for workbooks direct parsers may not read. · **Session ownership is the entire design.** `OwnedExcelSession` is the default: it creates a dedicated hidden Excel instance, may set its own `DisplayAlerts=False`, `EnableEvents=False`, `AskToUpdateLinks=False`, and `AutomationSecurity=msoAutomationSecurityForceDisable`, opens read-only with `UpdateLinks=0` and `AddToMru=False`, restores settings, closes its own workbook, and quits only that owned instance. `AttachedExcelSession` is an explicitly selected compatibility mode: it never changes global `Application` settings (including alerts, events, or security), never closes or quits anything, and reads only an already-open/read-only workbook through rectangular `Range.Value2` blocks. Both modes record `ownership ∈ {owned, attached}` in the manifest. · Forbidden: `taskkill`, enumerating/terminating `EXCEL.EXE`, any `Save`, and any settings mutation in attached mode. · State: sole owner of the Excel session handle and the COM apartment. · Tests: conformance suite (Windows+Excel-gated), no-kill assertion, attached-instance survival and no-settings-mutation tests, busy/modal retry with bounded backoff, cleanup-after-crash, `Value2` block-size benchmark.

### 5.10 `adapters.encrypted` — `adapters/encrypted_adapter.py`
Purpose: password-protected workbooks via msoffcrypto → in-memory decrypted stream handed to a direct adapter. Never writes a decrypted file to disk. Password sourced from an env var or an interactive prompt, **never** a CLI argument in V1 (process-table leakage), never logged, never in the manifest. · Tests: encrypted fixtures, wrong-password error mapping, no-secret-in-logs assertion, no-decrypted-artifact-on-disk assertion.

### 5.11 `discovery.engine` — `core/discovery/engine.py`
Purpose: deterministic candidate generation and selection, in strict priority order:
```
1 explicit contract target      → basis=contract        authoritative
2 Excel Table (ListObject)      → basis=excel_table     authoritative
3 named range                   → basis=named_range     authoritative
4 configured sheet/range/header → basis=configured      authoritative
5 bounded structural discovery  → basis=discovered      reported, warns W_DISCOVERY_ASSUMED
6 ambiguous / none              → NEEDS_INPUT           never guesses
```
Bounded discovery scans a capped window (default 200 rows × 200 cols), scores rectangular density, detects header rows (first row where ≥N cells are non-empty strings and the following K rows are type-consistent), trims leading notes, trailing totals (leading label + numeric row after a blank), repeated in-body header rows, and blank separator columns. Deterministic: same bytes → same candidates → same scores → same choice. Ties within an epsilon → ambiguous, not arbitrary. · Out: `Candidate[]` with score, basis, evidence (which rows/cols justified it), and a `rejected_because` for each loser. · Tests: golden-file per corpus fixture (candidate set + chosen candidate are snapshotted), determinism test (100 repeats), ambiguity test, property test (never proposes a range outside the sheet).

### 5.12 `extraction.planner` — `core/extraction/planner.py`
Purpose: turn `(contract, WorkbookFacts, Candidate, SizeEstimate, budget, policy flags)` into an immutable `ExtractionPlan{adapter, sheet, range, header_rows, data_start, columns[], policies{hidden_rows, merged, empty_rows, formulas, errors}, chunk_rows, expected_row_count}`. **This is where adapter selection happens** — one place, auditable, recorded verbatim in the manifest. · Tests: decision-table tests over the full (format × size × protection × flags) matrix; the table is the spec.

### 5.13 `extraction.reader` — `core/extraction/reader.py`
Purpose: drive the chosen `RowReader` port in chunks; emit `RawRow` batches with source row numbers attached at the source. Applies structural policies (skip hidden rows, drop empty rows, forward-fill merged anchors, strip repeated headers) and **records every drop in the ledger as it happens** — a drop that is not ledgered is impossible because the same function does both. · Tests: policy matrix, chunk-boundary tests (a dropped row at a chunk edge), cancellation at a chunk boundary.

### 5.14 `extraction.typing` — `core/extraction/typing.py`
Purpose: the fidelity engine. Maps Excel cell types + number formats to the output schema. Rules:
- Declared contract types win; failure to coerce → row-level `W_TYPE_COERCION_FAILED` + configurable `on_type_error ∈ {warn_keep_text, null, fail}`.
- No declared type → conservative inference per column over the first N chunks: a column is numeric only if **every** non-empty value is numeric *and* none is text-typed in Excel. **Text cells stay text** — this is what protects leading-zero identifiers.
- Dates resolved against the workbook date system (1900 vs 1904, incl. the Lotus 1900-02-29 bug); out-of-range serials → `W_DATE_OUT_OF_RANGE` + preserved raw.
- Number formats classified: percentage (store the true fraction *and* record the format), currency, text-formatted numbers, accounting.
- Excel error sentinels (`#N/A`, `#REF!`, …) never become nulls or strings silently: value → null, flagged cell recorded, `W_CELL_ERROR_VALUE`.
- Formulas: cached values by default; **missing cached value** (common in non-Excel-written files) → `W_FORMULA_NO_CACHED_VALUE` per cell, never a silent null. Formula-text capture is an opt-in column mode.
- Unicode: preserve; strip only control characters that Arrow/Parquet cannot carry, and flag each.
· Tests: a dedicated fidelity corpus with per-cell expected values (golden), property test "text-typed Excel cells never become numbers", locale fixtures.

### 5.15 `evidence.ledger` — `core/evidence/ledger.py`
Purpose: enforce **I2**. Buckets: `observed`, `emitted`, `excluded_structural` (header/blank/repeated-header/merged-filler), `excluded_hidden`, `excluded_filtered`, `excluded_policy` (e.g. totals), `failed`. Invariant:
```
observed == emitted + excluded_structural + excluded_hidden
            + excluded_filtered + excluded_policy + failed
```
Checked in the verify stage; violation → `E_RECONCILE_MISMATCH`, run does not succeed. · State: sole owner of the counts. · Tests: property test over randomized synthetic sheets — the identity must hold for every generated policy combination.

### 5.16 `evidence.lineage` — `core/evidence/lineage.py`
Purpose: answer "where did this value come from?" without exploding output size. Three-tier design:
1. **Anchor (always, ~1 KB/dataset):** source path + sha256, sheet, range, header row(s), data start row, column letter per output column, run id.
2. **Row (default on, 4–8 bytes/row):** `__src_row` int column. Anchor + `__src_row` + column letter = an exact A1 address for any value. This is the whole point: full cell addressability at one integer per row.
3. **Cell (exceptions only):** `cells_flagged.parquet` holds one record per *anomalous* cell (coercion failure, Excel error, missing cached value, stripped char) with row/col/raw text/reason. Clean data costs nothing.
· Tests: `explain --cell` resolves a published value back to its A1 address on every corpus fixture.

### 5.17 `evidence.issues` — `core/evidence/issues.py`
Purpose: one append-only stream of `Issue{severity, code, message, action, where{sheet,row,col}, count}` with **aggregation** — 400k identical warnings become one record with `count: 400000` and up to 10 exemplars. Prevents the classic "10 GB of warnings" failure. · Tests: aggregation, cap behavior, no-PII rule (no cell values in messages unless `--include-values-in-issues`).

### 5.18 `publication.ports` and `publication.sinks` — `core/publication/ports.py`, `sinks/`
Purpose: `core.publication.ports` owns the output-only `Sink` protocol and its publication input contracts; it is entirely separate from the Excel-reader ports. `sinks/` implements that protocol only: `ParquetSink` (default; one row group per chunk; zstd), `CsvSink` (opt-in; UTF-8 BOM, RFC4180, lossy-by-definition and says so in the manifest), `JsonlSink` (opt-in / stdout streaming), `PreviewSink` (always: first 100 rows as CSV for humans). Sinks may import only public contracts and `core.publication.ports`; they must not import `core.structure.ports`, the extraction implementation, or any other core internals. · Tests: sink conformance, golden bytes for small fixtures, round-trip type fidelity, row-group boundary tests.

### 5.19 `publication.publisher` — `core/publication/publisher.py`
Purpose: atomic publication. Build everything in `runs/<run_id>.partial/`; write `manifest.json` last; `fsync` files and the directory; `os.replace` the directory to `runs/<run_id>/`. A `.partial` directory is **by definition** not a successful run and is never read by `verify`/`explain`. Crash at any point leaves either nothing or a complete run. · State: sole owner of output publication. · Tests: fault injection at each step (kill between write and rename), disk-full simulation, pre-flight free-space check, read-only output path.

### 5.20 `orchestration.run` — `app/run.py`
Purpose: the run lifecycle and the only owner of run state (§7.6) and the cancellation token. Executes stages, times each, records the journey, guarantees the terminal state. · State: sole owner of run state and cancellation. · Tests: state machine exhaustive transition test, cancellation at every stage boundary, guaranteed-terminal-state property test.

### 5.21 `app.services` — `app/services/`
`InspectService`, `ExtractService`, `VerifyService`, `ExplainService`, `DoctorService`. Each takes a typed request, returns a typed result from `contracts`, and never prints. **The only thing CLI and TUI may call.** · Tests: service-level contract tests that are reused unchanged as the CLI's and TUI's behavioral baseline.

### 5.22 `presentation.cli` — `cli/`
Typer app; two renderers (`human`, `json`) behind a `Renderer` port; exit-code mapping. Contains **no** extraction logic — enforced by import-linter and by a test asserting no `cli` module imports `adapters`, `sinks`, or `core.extraction`.

### 5.23 `presentation.tui` — `tui/`
Textual app calling the same services. No business logic, no direct file reads, no adapter imports. Optional install.

### 5.24 `diagnostics.doctor` — `diagnostics/`
Read-only environment report: Python/tool version, wheel integrity, temp & output writability, free disk, long-path support, Excel presence (registry read-only) and version, COM availability (create + immediately release an owned instance), format capability matrix, optional-extra status, config validity, offline check that asserts *no* network is attempted. **Never modifies anything.** · Tests: each probe's failure mode is simulated.

---

## 6. Runtime Journeys

Each journey is declared in `.project-eye/graph.yaml` and stamped into every run manifest as `journey` + the stage sequence actually executed.

**J1 — Inspect workbook.**
`identity → safety(read-only open) → probe(size) → select inspector → WorkbookFacts → discovery(candidates, no extraction) → InspectResult` → human table or `--json`. Never creates a run directory (read-only, no side effects) unless `--record`.

**J2 — Extract workbook (normal).**
`J1 facts → planner(plan) → run CREATE(run_id) → RUNNING → reader(chunks) → typing → ledger accrual → sink writes to .partial → verify(reconcile + re-hash source) → publish(atomic rename) → terminal state → result`.

**J3 — Extract huge workbook.**
As J2, but the planner selects the streaming adapter because `est_bytes > budget`; chunking, a memory governor sampling RSS at each chunk boundary, progress events every chunk, one Parquet row group per chunk, and a cancellation check at each boundary. Peak RSS must stay flat as file size grows — that flatness is the acceptance test.

**J4 — Protected / COM path.**
`identity → direct open attempt → classify failure (encrypted? IRM? unsupported?) → if password available: encrypted adapter → else if --allow-excel: OwnedExcelSession (default; record ownership) → read-only open → rectangular Value2 block reads → …J2 tail… → release owned session and restore its settings`. Attached mode is a separately explicit request and only reads an already-open/read-only target without mutating the user's Excel environment. Without `--allow-excel` the run terminates `NEEDS_INPUT` with `E_SRC_IRM_PROTECTED` and the exact re-run command.

**J5 — Failure path.**
Any stage raises a typed `XlxError` → orchestrator captures stage, evidence refs, and partial ledger → `.partial` directory is retained under `runs/_failed/<run_id>/` when `--keep-failed` (default on) with a manifest marked `FAILED` and never named as a success → error rendered as {code, message, detail, action, run ref} → exit code from the family map. Tracebacks only under `--debug`.

**J6 — Cancel path.**
Ctrl-C / TUI cancel → cancellation token set → current chunk completes → reader stops at the boundary → ledger closed with `cancelled_at_row` → `.partial` retained or removed per flag → terminal `CANCELLED`, exit 8. Never leaves a published-looking directory.

**J7 — Verify / explain output.**
`verify <run>`: re-reads the run directory, re-checks the reconciliation identity, re-hashes published data files against the manifest digests, optionally re-hashes the source to confirm it is unchanged since extraction. `explain <run> [--row N | --cell A:12 | --decisions]`: prints the decision list with bases, or resolves a published row/cell back to its source A1 address.

---

## 7. Data and Contract Design

All models live in `contracts/` and generate JSON Schema into `schemas/`. Every document carries `contract_version`.

### 7.1 Extraction contract (`xlextract.contract/1`) — declarative, YAML or JSON

```yaml
contract_version: "xlextract.contract/1"
name: stock-on-hand
source:
  allowed_extensions: [".xlsx", ".xlsb"]
  expect_sha256: null            # optional pin; mismatch -> E_SRC_IDENTITY_MISMATCH
targets:
  - id: main
    sheet: { by: name, value: "Report" }     # or by_index / by_pattern
    region:
      by: excel_table                        # excel_table | named_range | a1_range | discover
      value: "tbl_stock"
    header:
      strategy: single_row                   # single_row | multi_row | none | explicit
      rows: 1
      multi_row_join: " / "
      on_duplicate: suffix                   # suffix | fail | keep_first
    data_start: auto                         # auto | <int>
    columns:
      include: ["*"]                         # or explicit ordered list
      exclude: []
      aliases: { "Mat. No.": material_id }
      types:
        material_id: { type: string }        # declared: never inferred, never numeric-ified
        qty:         { type: decimal, scale: 3 }
        posting_date:{ type: date }
    policies:
      hidden_rows: include                   # include | exclude
      hidden_sheets: error                   # error | allow
      empty_rows: drop                       # drop | keep
      merged_cells: fill_from_anchor         # fill_from_anchor | null_non_anchor | fail
      formulas: cached_value                 # cached_value | cached_or_fail | formula_text
      cell_errors: null_and_flag             # null_and_flag | keep_text | fail
      trailing_totals: drop_flagged          # drop_flagged | keep | fail
      on_type_error: warn_keep_text
output:
  formats: [parquet]                         # + csv | jsonl
  include_src_row: true
  chunk_rows: 50000
execution:
  security_mode: direct_only                 # direct_only | allow_excel | require_excel
  memory_budget_mb: 1024
  require_explicit_structure: false          # true => basis 'discovered' is refused
```

**Rule:** no user- or project-specific assumption may live in engine code. If a behavior varies by workbook, it belongs here.

### 7.2 Run manifest (`xlextract.manifest/1`) — `manifest.json`

```
run_id, started_at, ended_at, duration_ms, state, journey, stages[{name, ms, peak_rss_mb}]
tool { name, version, contract_version, python, platform }
source { path, sha256, size, mtime, format_sniffed, extension_matches_content,
         has_macros, has_external_links, encrypted, protection }
contract { source: inline|file, path, sha256, name }
plan { adapter, adapter_reason, sheet, range, header_rows, data_start, chunk_rows,
       memory_budget_mb, policies{...} }
decisions[ { what, value, basis, confidence, evidence, alternatives_rejected[] } ]
datasets[ { id, rows, columns, files[{path, bytes, sha256, row_groups}],
            schema_ref, content_digest, lineage_anchor{...} } ]
ledger { observed, emitted, excluded_structural, excluded_hidden, excluded_filtered,
         excluded_policy, failed, reconciled: true }
issues { errors, warnings, top_codes[{code, count}] }
excel_session { used: bool, ownership: owned|attached|none, version }     # COM runs only
result { status, exit_code, message }
```

### 7.3 Schema document (`xlextract.schema/1`) — `dataset.schema.json`
Per column: `name, source_header_raw, source_column_letter, ordinal, type, nullable, type_basis (declared|inferred|fallback_text), inference_evidence{sampled_rows, distinct_types}, number_format_sample, warnings[]`. Plus the Arrow type string for machine consumers. Schema is **stable across runs of the same contract**; a change in inferred type between runs is itself reported by `verify --against <prior_run>`.

### 7.4 Lineage
Anchor in the manifest; `__src_row` column in the data; `cells_flagged.parquet` for anomalies. §5.16.

### 7.5 Result model (stdout with `--json`) — `xlextract.result/1`
```json
{"contract_version":"xlextract.result/1","status":"succeeded_with_warnings","exit_code":0,
 "run_id":"...","run_dir":"...","datasets":[{"id":"main","rows":812344,"path":"..."}],
 "warnings":[{"code":"W_DISCOVERY_ASSUMED","count":1,"message":"..."}],
 "errors":[],"duration_ms":91233}
```
Exactly one JSON document on stdout, nothing else, ever. Human text and progress go to **stderr** so `xlx extract f.xlsx --json | jq` always works.

### 7.6 Run state model (minimum correct)

```
            ┌──────────┐
            │ PENDING  │  (validated, nothing touched)
            └────┬─────┘
                 ▼
            ┌──────────┐   stages: probe→inspect→plan→read→write→verify→publish
            │ RUNNING  │
            └────┬─────┘
   ┌────────┬────┴────┬─────────────┬──────────┐
   ▼        ▼         ▼             ▼          ▼
SUCCEEDED  SUCCEEDED_ NEEDS_INPUT  FAILED   CANCELLED
           WITH_
           WARNINGS
```
Five terminal states, all unambiguous. `NEEDS_INPUT` is terminal *by design* — the tool never blocks waiting; the user re-runs with more information. Stages are observable progress within `RUNNING`, **not** states — this keeps the machine minimal while keeping observability rich. `SUCCEEDED_WITH_WARNINGS` exits 0 unless `--fail-on-warning`.

### 7.7 Error model

`XlxError{code, family, message (one human sentence), detail (technical), action (what to do), where{stage, sheet, row, col}, run_id, retryable: bool}`.

| Family | Codes | Exit |
|---|---|---|
| usage/contract | `E_CONTRACT_INVALID`, `E_CONTRACT_VERSION_UNSUPPORTED`, `E_ARG_INVALID` | 2 |
| source | `E_SRC_NOT_FOUND`, `E_SRC_PERMISSION_DENIED`, `E_SRC_LOCKED`, `E_SRC_IN_TRANSIT`, `E_SRC_CORRUPT`, `E_SRC_EXTENSION_MISMATCH`, `E_SRC_IDENTITY_MISMATCH`, `E_SRC_CHANGED_DURING_RUN`, `E_FMT_UNSUPPORTED` | 3 |
| needs input | `E_DISCOVERY_AMBIGUOUS`, `E_DISCOVERY_NONE`, `E_SRC_ENCRYPTED_NO_KEY`, `E_SRC_IRM_PROTECTED`, `E_EXCEL_REQUIRED_NOT_ALLOWED` | 4 |
| extraction | `E_TARGET_SHEET_MISSING`, `E_TARGET_RANGE_MISSING`, `E_MEM_BUDGET_EXCEEDED`, `E_TYPE_COERCION_FAILED`, `E_RECONCILE_MISMATCH` | 5 |
| output | `E_OUT_PATH_UNAVAILABLE`, `E_OUT_DISK_FULL`, `E_OUT_WRITE_FAILED`, `E_OUT_PUBLISH_FAILED` | 6 |
| environment | `E_EXCEL_UNAVAILABLE`, `E_EXCEL_BUSY`, `E_EXCEL_TIMEOUT`, `E_COM_UNAVAILABLE`, `E_DEP_MISSING` | 7 |
| cancelled | `E_CANCELLED` | 8 |
| internal | `E_INTERNAL` | 70 |

Warnings use the `W_` prefix (list in §5.14/§5.17). **Users never see a raw traceback as the normal experience**; an unexpected exception is wrapped as `E_INTERNAL` with a run reference and a "please report with `runs/<id>/run.log.jsonl`" action. Tracebacks appear only with `--debug`.

### 7.8 Versioning & compatibility
Contract versions are `<family>/<major>`. Within a major: additive fields only; consumers must ignore unknown fields; no field is ever repurposed. A breaking change means a new major, and the tool reads *N* and *N−1* for at least one minor release. Tool version is SemVer; `content_digest` (see §16/idempotence) may only change on a **major** tool version — a patch release that changes extracted values is a defect by definition. Golden schema snapshots in CI make an accidental break impossible to merge.

---

## 8. CLI Design

```
xlx inspect  <path>                     # J1: describe a workbook; no side effects
xlx extract  <path|dir> [options]       # J2/J3/J4: the main verb
xlx verify   <run-id|run-dir>           # J7: re-prove a published run
xlx explain  <run-id> [--row N|--cell REF|--decisions]
xlx contract init|validate|infer        # author/check/derive an extraction contract
xlx runs     list|show|prune
xlx doctor                              # environment readiness
xlx ui                                  # launch the TUI (optional extra)
```

Global: `--json` · `--quiet` · `--verbose/-v` (repeatable) · `--debug` · `--no-color` · `--config PATH` · `--version`.

`extract` options: `--contract PATH` · `--sheet NAME` · `--range A1:F900` · `--table NAME` · `--name RANGE` · `--header-rows N` · `--out DIR` · `--format parquet|csv|jsonl` (repeatable) · `--stdout jsonl` · `--allow-excel` / `--require-excel` · `--password-env VAR` · `--memory-budget-mb N` · `--chunk-rows N` · `--include-hidden-rows` · `--require-explicit` · `--fail-on-warning` · `--dry-run` · `--keep-failed/--no-keep-failed` · `--glob PATTERN` (folders) · `--continue-on-error` (folders).

Examples:
```
xlx inspect "Q3 stock.xlsx"
xlx inspect "Q3 stock.xlsx" --json | jq '.sheets[] | select(.visibility!="visible")'
xlx extract "Q3 stock.xlsx" --table tbl_stock --out .\out
xlx extract .\downloads\ --glob "*.xlsb" --contract stock.yaml --continue-on-error
xlx extract "protected.xlsx" --allow-excel --out .\out
xlx extract big.xlsx --memory-budget-mb 512 --chunk-rows 25000 -v
xlx extract small.xlsx --stdout jsonl | my-cleaner --in -
xlx verify 2026-09-12T101500Z-a3f9
xlx explain 2026-09-12T101500Z-a3f9 --row 4211
xlx doctor --json
```

**Behavioral rules.** Safe defaults (no COM, no overwrite of an existing run, source never touched). Human output and progress → stderr; the single JSON result → stdout. Progress starts within 2 s on any file. `--dry-run` performs J1 + planning and prints the plan without reading data. Shell completion is generated, not hand-written. `--help` for every command includes a worked example.

**Resumability** is deliberately *not* offered mid-file (§13.6). It is offered at the **folder** level: re-running a folder skips sources whose `(sha256, contract sha256, tool major)` already has a `SUCCEEDED` run, unless `--force`. That is where it is genuinely valuable and cheap to make correct.

---

## 9. TUI Design

One Textual app, `xlx ui`, calling **only** `app.services`. No adapter imports, no file reads, no logic that the CLI does not also have. Anything the TUI can do is expressible as a CLI command, and the TUI displays that command — so a user can graduate to automation.

Screens:
1. **Pick source** — file/folder browser, recent sources, drag-drop path paste.
2. **Inspect** — sheets table (name, visibility, dimensions, trusted?), Tables, named ranges, workbook flags (macros, external links, encryption).
3. **Choose target** — candidate list with basis and score, evidence pane, ambiguity clearly marked; selecting a candidate is the `user_confirmed` basis.
4. **Preview** — first 50 rows as the plan would extract them, with the detected header highlighted and flagged cells colored.
5. **Options** — formats, output dir, policies, `--allow-excel` toggle with a plain-language explanation of what it does.
6. **Run** — stage list, chunk progress, rows/sec, peak RSS, live warning counter, `c` to cancel.
7. **Result** — terminal state, ledger reconciliation table, top issues, output paths, **and the equivalent CLI command** for reuse.

Keyboard: `↑↓/jk` move · `Enter` select · `Tab/Shift-Tab` panes · `e` extract · `c` cancel · `?` help · `q` quit · `y` copy CLI command. Every destructive-looking action requires explicit confirmation; there are no destructive actions on sources by construction.

---

## 10. Repository Design

```
Excel-Data-Extractor/
├─ AGENTS.md                  # how agents work here (entry point for any AI agent)
├─ ARCHITECTURE.md            # intended architecture + rules rationale
├─ PROJECT_EYE.md             # the living 5-level map (concise)
├─ CURRENT_STATE.md           # what actually exists right now
├─ LESSONS.md                 # hard-won facts; append-only
├─ PLAN.md                    # this document
├─ README.md                  # user-facing quickstart
├─ pyproject.toml             # deps, extras [com,crypto,tui,dev], console scripts, tool config
├─ .project-eye/
│  ├─ graph.yaml              # nodes, edges, journeys, ownership (hand-maintained truth)
│  ├─ rules.yaml              # architecture rules, machine-checkable
│  └─ manifest.json           # GENERATED: file hashes, drift status, last verified
├─ schemas/                   # GENERATED JSON Schema for every public contract
├─ src/xlextract/
│  ├─ contracts/              # public models; depends on pydantic ONLY
│  ├─ core/
│  │  ├─ source/              # identity, safety, probe
│  │  ├─ structure/           # model, ports
│  │  ├─ discovery/           # engine, headers, scoring
│  │  ├─ extraction/          # planner, reader, typing, chunking
│  │  ├─ evidence/            # ledger, lineage, issues
│  │  ├─ publication/         # output Sink port, publisher, run directory layout
│  │  └─ errors/              # taxonomy, exit codes
│  ├─ adapters/               # calamine, openpyxl_stream, excel_com/, encrypted
│  ├─ sinks/                  # parquet, csv, jsonl, preview
│  ├─ app/                    # services/, run.py, wiring.py  (the ONLY CLI/TUI entry)
│  ├─ cli/                    # typer app, renderers, exit-code mapping
│  ├─ tui/                    # textual app  [optional extra]
│  └─ diagnostics/            # doctor probes
├─ tests/
│  ├─ unit/ component/ contract/ adapter/ golden/ property/ cli/ tui/ e2e/
│  ├─ perf/                   # benchmarks with thresholds
│  ├─ real_excel/             # marked; skipped without Excel  [gated]
│  └─ conftest.py             # source-integrity + no-network global fixtures
├─ fixtures/
│  ├─ generators/             # synthetic workbook builders (deterministic, seeded)
│  ├─ generated/              # GENERATED, git-ignored
│  └─ committed/              # small, hand-made edge cases (committed)
├─ benchmarks/                # scenario definitions + published results
├─ tools/                     # eye.py (map check/sync), gates.py, release.py
└─ docs/
   ├─ user-guide.md           # ≤2 pages, non-technical
   ├─ cli-reference.md        # GENERATED from typer
   ├─ contract-reference.md   # GENERATED from schemas
   └─ integration.md          # for downstream tools/agents
```

Every directory has one responsibility. No `utils/`, no `helpers/`, no `common/` — those are where architecture goes to die; shared code lives in the domain that owns it.

---

## 11. Architecture Rules (mechanically checked)

Declared in `.project-eye/rules.yaml`, enforced in CI by `import-linter` + `tools/gates.py`. A violation fails the build.

**Layering (import-linter `layers` contract):**
```
cli, tui  →  app  →  core  →  contracts
adapters    →  core.structure.ports + contracts
sinks       →  core.publication.ports + contracts
diagnostics →  core.errors + contracts
```

| # | Rule | Check |
|---|---|---|
| R1 | `contracts` imports nothing from `xlextract` | import-linter `independence` |
| R2 | `core` never imports `adapters`, `sinks`, `app`, `cli`, `tui` | forbidden contract |
| R3 | `cli` and `tui` never import `adapters`, `sinks`, or `core.extraction` | forbidden contract |
| R4 | `tui` never imports `cli` and vice versa | independence |
| R5 | `win32com`/`pythoncom` imported **only** inside `adapters/excel_com/` | grep gate + forbidden contract |
| R6 | `openpyxl`, `python_calamine`, `msoffcrypto` imported only in `adapters/` | forbidden contract |
| R7 | `pyarrow` imported only in `sinks/` | forbidden contract |
| R8 | `rich`/`textual` imported only in `cli`/`tui` | forbidden contract |
| R9 | No `print()` outside `cli/renderers` and `tui` | AST gate |
| R10 | No module-level mutable state outside `app.wiring` | AST gate (module-level assignments of mutable types) |
| R11 | No import cycles anywhere | import-linter |
| R12 | No module > 500 LOC / no package > 12 modules without an Eye entry (god-module guard) | `tools/gates.py` |
| R13 | Every public contract model appears in exactly one module (duplicate-schema guard) | `tools/gates.py` |
| R14 | Every `graph.yaml` node maps to an existing module, and every `src/` module ≥50 LOC maps to a node (orphan guard) | `tools/eye.py check` |
| R15 | Every declared journey's stage sequence is exercised by at least one test | `tools/eye.py check --journeys` |
| R16 | No `socket`/`urllib`/`requests`/`httpx` import anywhere in `src/` | grep gate + runtime test |
| R17 | No write-mode file open against a source path | AST gate + runtime `conftest` hash proof |
| R18 | Single-owner registry: run state, Excel session, publication, config, source identity, cancellation each have exactly one owning module, declared in `graph.yaml` | `tools/eye.py check --owners` |
| R19 | adapters may import only `contracts` and `core.structure.ports` from this project | AST import allowlist |
| R20 | sinks may import only `contracts` and `core.publication.ports` from this project | AST import allowlist |
| R21 | diagnostics may import only `contracts` and `core.errors` from this project | AST import allowlist |

**Detected drift classes** (reported by `tools/eye.py check`): cycles, god modules, duplicate ownership, duplicate schemas, hidden global state, cross-domain shortcuts, UI→persistence/extraction violations, unmapped runtime paths, orphan components. Three views are compared: **INTENDED** (`graph.yaml`) vs **STATIC ACTUAL** (AST import graph) vs **RUNTIME ACTUAL** (stage/adapter sequences harvested from run manifests produced by the test suite). Any meaningful disagreement is architecture drift and fails the map gate.

---

## 12. Excel Adapter Strategy

**Selection decision table** (owned solely by `extraction.planner`; recorded in the manifest as `adapter_reason`):

| Format | Encrypted? | Est. sheet bytes vs budget | `security_mode` | → Adapter |
|---|---|---|---|---|
| xlsx/xlsm | no | ≤ 0.6 × budget | any | **calamine** |
| xlsx/xlsm | no | > 0.6 × budget | any | **openpyxl_stream** |
| xlsx/xlsm | yes + password | any | any | **encrypted → calamine/openpyxl_stream** |
| xlsx/xlsm | yes, no password | — | `allow_excel`/`require_excel` | **excel_com** |
| xlsx/xlsm | yes, no password | — | `direct_only` | `E_SRC_ENCRYPTED_NO_KEY` (exit 4) |
| xlsb | no | ≤ 0.6 × budget | any | **calamine** |
| xlsb | no | > 0.6 × budget | `allow_excel` | **excel_com** |
| xlsb | no | > 0.6 × budget | `direct_only` | `E_MEM_BUDGET_EXCEEDED` (exit 5) with the exact flag to fix it |
| xls | no | (capped at 65k rows) | any | **calamine** |
| any | IRM/DRM | — | `allow_excel`/`require_excel` | **excel_com** |
| any | IRM/DRM | — | `direct_only` | `E_SRC_IRM_PROTECTED` (exit 4) |
| any | corrupt/unreadable | — | `allow_excel` | **excel_com** attempt, then `E_SRC_CORRUPT` |

**Fallback policy.**
- Direct → direct fallback (calamine → openpyxl) is **automatic** on `MemoryError`/budget breach, recorded as a decision with `W_ADAPTER_FALLBACK`.
- Direct → COM is **never automatic**. It requires `security_mode: allow_excel` (or `--allow-excel`). Without it: terminal `NEEDS_INPUT` naming the exact re-run command.
- COM → direct never happens; if COM was required, there is no direct answer.
- Every fallback is bounded (one step), recorded, and visible in `explain`.

**COM safety contract** (binding on `adapters/excel_com/`):
`owned` is the default and creates a dedicated hidden Excel instance. Only that owned instance may receive temporary application settings (`DisplayAlerts=False`, `EnableEvents=False`, `AskToUpdateLinks=False`, `AutomationSecurity=ForceDisable`); they are restored in `finally`, then only tool-opened workbooks are closed and only the owned instance is quit. `attached` is explicit compatibility mode and is effectively read-only to the user's Excel environment: it changes no global `Application` setting (including events, alerts, or security), opens nothing, closes nothing, quits nothing, and reads only an already-open/read-only target. Both modes use read-only data access, never save, never update links, use bounded busy/modal retry (default 3 attempts, 2/5/10 s), and rectangular `Value2` blocks sized to ~50k cells. Never kill or enumerate Excel processes. Session release is guaranteed on every exit path (context manager + `atexit` guard); staging is used only for an owned session when the source is on a volatile share.

**Honesty clause.** Real protected-file support is **unproven** until it is demonstrated on an actual protected workbook in the actual corporate environment. A synthetic password-protected fixture proves the *encrypted* path only. The COM gate (§23) cannot be signed off by CI.

---

## 13. Large-File Strategy

**13.1 Budget-first.** `memory_budget_mb` (default 1024, configurable) is a first-class input. The planner uses `probe` to estimate per-sheet uncompressed bytes from the zip central directory (no decompression) and selects the adapter by the table in §12. The budget is a **decision input**, not a hope.

**13.2 Chunking.** Fixed row chunks (default 50 000 rows, auto-reduced if a chunk's measured footprint exceeds `budget/4`). One chunk = one Parquet row group = one progress event = one cancellation point = one ledger accrual point. Aligning all five on the same boundary is what keeps the design small.

**13.3 Streaming.** `openpyxl` `read_only=True, values_only=True` with `reset_dimensions()` to defeat bloated/phantom dimension records; the workbook is closed explicitly to release the read-only cache. Never cell-by-cell for a rectangular region; never `ws[...]` random access in the streaming path.

**13.4 Memory governor.** RSS sampled at each chunk boundary. Soft breach (>80% budget) → halve `chunk_rows`, warn `W_MEMORY_PRESSURE`. Hard breach (>100%) → if a bounded adapter is available, fall back and restart the read (recorded); else fail `E_MEM_BUDGET_EXCEEDED` with the flag that would fix it. Never OOM-crash silently.

**13.5 Temporary data.** All intermediates in `runs/<id>.partial/`, on the same volume as the final output so the publish rename is atomic. A pre-flight free-space check requires `3 × estimated output size`. Cleanup is registered before the first byte is written.

**13.6 Publication & the resumability decision.** Atomic directory rename (§5.19). Mid-file resume is **deliberately not implemented in V1**: it would require a durable, verifiable position in a stream whose source may change under us, and a partial-output format that could be mistaken for a complete one — a direct threat to I1 and I2 for a benefit (restarting a ≤5-minute job) that does not justify the risk. Folder-level skip-if-done (§8) captures the real-world value at a fraction of the risk. Revisit only if benchmarks show single-file runs routinely exceeding 15 minutes.

**13.7 Progress.** First progress event ≤2 s (emitted after probe, before the first chunk). Per chunk: rows done, rows/sec, elapsed, ETA (only when total rows are known), peak RSS, warning count. Machine-readable progress on stderr as JSONL under `--progress json`.

**13.8 Benchmarking.** `benchmarks/` defines scenarios; `tests/perf/` asserts thresholds (§16). Adapter comparison is **empirical and published** — the chosen default must be justified by numbers in `benchmarks/results/`, not by popularity.

---

## 14. Test Strategy (tests as architecture evidence)

Risk-first mapping — every layer exists because a specific failure would otherwise reach a user.

| Layer | Guards which risk | Notes |
|---|---|---|
| **unit** | logic defects in typing, scoring, ledger math | fast, no I/O |
| **component** | a domain misbehaving as a whole (discovery, planner) | in-memory fakes for ports |
| **contract** | breaking a public schema | golden JSON-Schema snapshots + example validation + N−1 compatibility |
| **adapter conformance** | one adapter behaving differently from another | **one suite, run against every adapter**; COM-gated variants skipped without Excel |
| **golden-file** | silent change in extracted values | per-corpus expected datasets + expected manifests (with volatile fields masked) |
| **property/fuzz** | edge cases humans do not imagine | ledger identity over randomized sheets; "text-typed cells never become numeric"; discovery never returns out-of-bounds ranges; fuzz on truncated/corrupt zip bytes |
| **performance** | regressions in speed/memory | thresholds in §16; peak-RSS flatness across file sizes |
| **recovery / fault injection** | crash, disk full, source changed, Excel busy, kill during publish | `.partial` must never look successful |
| **CLI** | UX and machine-interface stability | exit codes, stdout-is-only-JSON, stderr separation, `--help` snapshots |
| **TUI** | the TUI diverging from core | Textual pilot driving the same services; assert **no** adapter imports |
| **e2e** | the whole journey | J1–J7 end to end on the corpus |
| **real Excel** | the COM path | `@pytest.mark.real_excel`, skipped without Excel; **required** for the COM gate |
| **protected-file conditional proof** | enterprise DRM reality | manual, environment-specific, evidence recorded in `LESSONS.md` — never claimed by CI |

**Two global fixtures in `tests/conftest.py` that carry disproportionate weight:**
- `source_integrity_guard` (session scope): SHA-256 every file under `fixtures/` before and after the entire suite; any change fails the run. This is the mechanical proof of **I1**.
- `no_network_guard` (autouse): patch `socket.socket` to raise. Any accidental network call fails the test that made it. This is the mechanical proof of the offline guarantee.

**Defect protocol (mandatory, no exceptions):** reproduce → capture evidence (the run directory) → write the smallest failing test → confirm it fails → fix → confirm it passes → run the affected journey e2e → update Eye/contracts if the fix touched either.

**Per-change verification checklist:** node · edges/contracts · affected subgraph · journey · architecture rules · map freshness. Enforced by `tools/gates.py` in CI and restated in `AGENTS.md`.

---

## 15. Test Corpus

`fixtures/generators/` produces every generated workbook deterministically from a seed (committed as code, not as binaries); `fixtures/committed/` holds small hand-made oddities that cannot be generated. `fixtures/generated/` is git-ignored and rebuilt by `python -m tools.fixtures build`.

| ID | Fixture | Proves |
|---|---|---|
| F01 | tiny clean (10×5, one sheet) | happy path, startup budget |
| F02 | multi-sheet, mixed visibility (visible/hidden/very-hidden) | visibility policy, `hidden_sheets: error` |
| F03 | Excel Table + named range + plain range in one workbook | discovery priority order |
| F04 | notes above, blanks, data, blank, totals below | trimming, `trailing_totals` |
| F05 | multi-row header (3 rows) + duplicate column names | header join, `on_duplicate` |
| F06 | repeated header rows inside the body | in-body header detection |
| F07 | merged cells (horizontal, vertical, block) | merge policies |
| F08 | blank columns/rows inside the region; sparse island layout | region bounding, candidate scoring |
| F09 | phantom/bloated `UsedRange` (declared 1 048 576 rows, 12 real) | `reset_dimensions`, trimmed-range warning |
| F10 | formulas with cached values | `cached_value` policy |
| F11 | formulas **without** cached values (non-Excel writer) | `W_FORMULA_NO_CACHED_VALUE` — never a silent null |
| F12 | every Excel error sentinel | error-cell policy |
| F13 | dates: 1900 system, 1904 system, pre-1900, 1900-02-29 bug, datetime, time-only, text-dates | date correctness |
| F14 | leading-zero IDs, long numeric strings (>15 digits), text-formatted numbers | **the fidelity fixture** |
| F15 | percentages, currency, accounting, locale decimal/thousand separators | number-format fidelity |
| F16 | very long text (>32k), emoji, RTL, control characters, mixed scripts | Unicode handling |
| F17 | mixed types within a column | inference conservatism |
| F18 | hidden rows + active autofilter | hidden/filtered accounting |
| F19 | wide (16 384 columns) | column-limit handling |
| F20 | large (1 000 000 rows × 20 cols, ~250 MB) | streaming, memory flatness, perf gate |
| F21 | huge-sparse (data at XFD1048576) | dimension sanity, not 10^10 cells |
| F22 | multiple independent datasets in one sheet | multi-candidate, ambiguity |
| F23 | legacy `.xls` | legacy adapter |
| F24 | `.xlsb` (small and over-budget) | xlsb path + honest over-budget failure |
| F25 | `.xlsm` with macros | macros never executed |
| F26 | external links | links never updated |
| F27 | password-protected `.xlsx` | encrypted adapter |
| F28 | truncated / corrupt zip | `E_SRC_CORRUPT`, no crash |
| F29 | wrong extension (`.xls` that is OOXML, and the reverse) | content sniffing |
| F30 | path > 260 chars, UNC path, spaces/Unicode in filename | Windows path reality |
| F31 | file locked by another process / mid-copy | `E_SRC_LOCKED`, `E_SRC_IN_TRANSIT` |
| F32 | **real protected enterprise workbook** | COM path — *environment-local, never committed, never in CI* |

**Honesty clause:** F27 proves password support. It does **not** prove enterprise DRM/IRM support. Only F32, run manually on the target environment, does — and that result is recorded in `LESSONS.md` with the date and Excel version.

---

## 16. Performance Acceptance

Reference machine: corporate Windows 11 laptop, 4 physical cores, 16 GB RAM, NVMe, antivirus active (the realistic case, not a clean benchmark box). All thresholds are asserted in `tests/perf/` and published to `benchmarks/results/`.

| Scenario | Budget | Rationale |
|---|---|---|
| `xlx --help` cold | ≤ 400 ms | CLI must feel like a tool; enforced by lazy adapter imports |
| `xlx inspect` on F01 | ≤ 800 ms | tiny must feel instant |
| `xlx inspect` on F20 (250 MB) | ≤ 6 s | inspection must not read data |
| time to first progress event, any file | ≤ 2 s | user confidence |
| calamine path throughput | ≥ 1.5 M cells/s | fast path must be worth its memory |
| streaming path throughput | ≥ 250 k cells/s | acceptance floor for the bounded path |
| COM path throughput | ≥ 150 k cells/s | with 50k-cell `Value2` blocks |
| F20 full extract → Parquet | ≤ 5 min | 20 M cells on the streaming path |
| **peak RSS, streaming path** | **≤ 700 MB, and flat (±15%) across 50 MB / 250 MB / 500 MB inputs** | the single most important number in this table |
| peak RSS, calamine path | ≤ 0.6 × budget, enforced by the guard | fast path stays inside its contract |
| publication (rename) | ≤ 200 ms | atomicity must not cost |
| cancellation acknowledged | ≤ 2 s (or one chunk) | responsiveness |
| `verify` on a published F20 run | ≤ 30 s | re-proof must be cheap enough to be routine |

**Method.** Each scenario runs 3× after one warm-up; the median is recorded; peak RSS via `psutil` sampled at chunk boundaries and at process exit. Regressions >20% fail the perf gate. Adapter choices must cite a `benchmarks/results/` entry — **no library is chosen for popularity**.

---

## 17. Security and Corporate-Safety Contract

**The tool MAY:** read files the invoking user can already read · write only inside the user-specified output directory and the system temp directory · read the Windows registry read-only to detect Excel · create and control an Excel instance **it created** · attach to an existing Excel instance for a read-only open, restoring every setting it changed · hold a password in memory for the lifetime of a decrypt operation.

**The tool MUST NOT:** make any network connection · emit telemetry · send source data, file names, or headers to any AI or remote service · modify, save, rename, move, or delete a source workbook · execute macros · update external links · write a decrypted copy of an encrypted workbook to disk · store secrets in code, logs, manifests, or run directories · accept a password as a command-line argument (process-table leakage) · terminate, enumerate for termination, or otherwise interfere with Excel processes it does not own · change Windows, Office, macro-security, DRM, execution-policy, or antivirus settings · require administrator rights at install or run time · install anything system-wide · log cell values by default.

**Enforcement is mechanical, not aspirational:** R16 (no network imports) + `no_network_guard` runtime test; R17 + `source_integrity_guard`; a log-scrubbing test that asserts no fixture password or cell value appears in any run log; the COM no-kill assertion test; a `doctor` probe that reports — and never repairs — the environment.

**Data handling.** Run directories contain business data by definition; they are written with default user ACLs inside the user-chosen output path and are never auto-shared, auto-uploaded, or auto-cleaned. Logs contain identities, counts, codes and addresses — **not values** — unless `--include-values-in-issues` is explicitly passed.

---

## 18. Project Eye Bootstrap

Five documents plus `.project-eye/`. Concise by mandate: **PROJECT_EYE.md must stay under ~250 lines**; if it grows, the architecture is too big, not the document.

**`AGENTS.md`** — the entry point for any AI agent. Contains: read-order (this file → PROJECT_EYE → the relevant node → the contract); the context-pack template (§ Agent Operating Contract); the stop conditions (§23.2); the scope rule (`RECORD_ONLY`); the per-change verification checklist; the anti-drift rule; the commands to run before claiming done (`ruff`, `mypy`, `pytest -m "not real_excel"`, `python -m tools.gates all`, `python -m tools.eye check`); and an explicit "never do this" list mirroring §17.

**`ARCHITECTURE.md`** — the intended architecture: the hexagon diagram, the three invariants I1–I3, the layering rules and *why* each exists, the rejected alternatives (§3.2), and the adapter decision table (§12). Changes here require an Opus-level review.

**`PROJECT_EYE.md`** — the living map, five levels, with one table per level. Level 3 lists each component with `purpose | owner | inputs | outputs | deps | contract | state | tests | journeys`. Level 5 lists the seven journeys with their stage sequences. Ends with a drift section reporting the last `eye check` result.

**`CURRENT_STATE.md`** — what actually exists *today*: implemented components, phase status, known gaps, unproven claims (the COM/DRM path lives here until F32 passes), open `RECORD_ONLY` items, and the next planned task IDs. Updated by every merged task — a task that does not update it is not done.

**`LESSONS.md`** — append-only, dated entries of hard-won facts: e.g. "2026-09-xx — calamine materializes sheets (~5× file size in RSS); this is why the planner probes before choosing." Each entry: date, context, finding, consequence for the design. This is the file that prevents re-learning.

**`.project-eye/graph.yaml`**
```yaml
version: 1
nodes:
  - id: core.discovery.engine
    level: 3
    domain: discovery
    path: src/xlextract/core/discovery/engine.py
    purpose: deterministic candidate generation and selection
    owner_of: []
    inputs: [WorkbookFacts, ExtractionContract]
    outputs: [Candidate]
    tests: [tests/component/test_discovery.py, tests/golden/discovery/]
    journeys: [J1, J2]
edges:
  - from: app.services.extract
    to: core.extraction.planner
    kind: calls
journeys:
  - id: J3
    name: extract huge workbook
    stages: [identity, safety, probe, inspect, plan, read, type, ledger, write, verify, publish]
owners:                      # the One Owner Rule, machine-checked (R18)
  run_state: app.run
  excel_session: adapters.excel_com.session
  output_publication: core.publication.publisher
  configuration: app.wiring
  source_identity: core.source.identity
  cancellation: app.run
```

**`.project-eye/rules.yaml`** — R1–R21 in machine-readable form (layer definitions, narrow infrastructure and diagnostics import allowlists, AST gate parameters, size limits, owner registry), consumed by both `import-linter` and `tools/gates.py`.

**`.project-eye/manifest.json`** — **generated**: per-node file hashes, last verification timestamp, tool version, and the three-view comparison result (INTENDED / STATIC ACTUAL / RUNTIME ACTUAL) with any drift findings. Never hand-edited.

---

## 19. Phased Build Plan

Eight vertical phases. **Every phase ends with a capability that can be demonstrated on a real file.** No phase builds a layer in isolation for later integration.

| Phase | Demonstrable capability at the end | Tasks |
|---|---|---|
| **P0 Foundation** | `xlx doctor` reports the environment; CI runs lint+types+tests+gates; the Eye exists and `eye check` passes on an almost-empty repo | T-001, T-002, T-003, S-001 |
| **P1 See** | `xlx inspect file.xlsx` truthfully describes any of the four formats, human and `--json` | S-002, S-003, T-004 |
| **P1.5 Package smoke** | portable one-dir `xlx doctor` and `xlx inspect F01` run locally with no runtime network or admin rights | T-011, T-014 |
| **P2 Extract (explicit)** | `xlx extract file.xlsx --table X` → published run dir with Parquet + manifest + ledger; `xlx verify` re-proves it | S-004, T-005, S-005, T-006 |
| **P3 Find** | `xlx extract file.xlsx` with no target: discovery finds it, reports the basis, refuses to guess when ambiguous; `xlx contract infer` writes a contract | S-006, T-007, T-008 |
| **P4 Scale** | F20 (1 M rows) extracts in ≤5 min with flat, bounded memory, live progress, and working cancellation | S-007, T-009 |
| **P5 Formats** | `.xlsb`, `.xls`, `.xlsm`, and password-protected workbooks all extract or fail honestly | T-010, S-008 |
| **P6 Protected** | An IRM-protected workbook extracts via `--allow-excel` on the real target machine without disturbing the user's Excel | S-009, T-011 |
| **P7 Friendly** | A non-technical user completes an extraction in the TUI and is shown the equivalent CLI command | T-012 |
| **P8 Harden** | Fault-injection suite green, perf gate green, docs complete, one-dir build runs on a clean standard-user machine | S-010, T-013 |

P0–P2 are the critical path; nothing else starts before P2 is green. P5 and P7 can overlap P4/P6 (see §22).

---

## 20. Sonnet 5 Execution Packs

Format is fixed. `Forbidden files` means: do not open, do not edit, do not "fix in passing".

---

### S-001 — Public contracts v1
- **Goal:** define and freeze the `xlextract.contract/1`, `manifest/1`, `schema/1`, `result/1`, and `ledger/1` models with generated JSON Schema and a compatibility policy.
- **Why Sonnet:** this is the product's public API; every later task depends on getting it right, and versioning/compat rules are judgement-heavy.
- **Prerequisites:** T-001.
- **Target node:** `contracts`. **Domain:** contracts. **Journey:** all.
- **Allowed:** `src/xlextract/contracts/**`, `schemas/**`, `tests/contract/**`, `docs/contract-reference.md`.
- **Forbidden:** everything under `core/`, `adapters/`, `sinks/`, `cli/`, `app/`.
- **Inputs:** §7 of this plan (authoritative field lists).
- **Requirements:** pydantic v2 models, frozen where semantically immutable; `contract_version` on every root model; JSON Schema emitted to `schemas/` by `python -m tools.schemas build`; a documented compat policy (additive-only within a major; unknown fields ignored on read); example documents for every model under `tests/contract/examples/`; no dependency other than pydantic.
- **Tests:** round-trip for every model; golden schema snapshots (byte-compared in CI); every example validates; a rejection test per invalid-example; N−1 read test scaffolding (even if N−1 == N today).
- **Acceptance evidence:** `schemas/*.json` committed; `pytest tests/contract` green; `import-linter` proves `contracts` imports nothing internal.
- **Stop if:** a field in §7 cannot be expressed without importing something outside `contracts`; a needed field is missing from §7 (propose, do not invent silently).
- **Handoff:** frozen schemas + a one-page "how to consume this output" note for `docs/integration.md`.

---

### S-002 — Source identity & safety layer (invariant I1)
- **Goal:** implement `SourceRef`, read-only access, format sniffing, staging policy, mid-run change detection, and the mechanical proof that sources are never written.
- **Why Sonnet:** this is a security/data-loss boundary with nasty Windows specifics (long paths, UNC, locks, in-transit files).
- **Prerequisites:** S-001.
- **Target node:** `core.source.*`. **Domain:** source. **Journey:** J1–J7.
- **Allowed:** `src/xlextract/core/source/**`, `tests/unit/source/**`, `tests/conftest.py`, `fixtures/generators/paths.py`.
- **Forbidden:** `adapters/**`, `cli/**`, `app/**`.
- **Requirements:** SHA-256 streaming hash (bounded memory); content sniffing (ZIP/OOXML vs OLE2/CFB vs xlsb inside OOXML) independent of extension, reporting mismatches; `\\?\` long-path normalization; lock and in-transit detection (size/mtime settling probe with bounded retries); staging copy API that preserves and records identity; end-of-run source re-hash → `E_SRC_CHANGED_DURING_RUN`; **all opens read-only** (`open(..., "rb")` only).
- **Tests:** hash stability; each Windows path shape (F30); F29 extension mismatch; F31 locked/in-transit; staging cleanup on simulated crash; plus author the session-scoped `source_integrity_guard` fixture.
- **Acceptance evidence:** the integrity guard is active for the whole suite and green; F29/F30/F31 tests pass on Windows.
- **Stop if:** any requirement seems to need a write to a source path; long-path handling requires a machine policy change (it must not).
- **Handoff:** `SourceRef` + safety API used by every adapter.

---

### S-003 — Workbook structure model, ports, and the calamine inspector
- **Goal:** define `WorkbookInspector`/`RowReader`/`CapabilityReport` ports, the format-neutral fact model, the adapter conformance suite, and the first inspector.
- **Why Sonnet:** the port design fixes the shape of every adapter, including COM. Getting it wrong is expensive later.
- **Prerequisites:** S-002.
- **Target node:** `core.structure.*`, `adapters.calamine`. **Domain:** structure. **Journey:** J1.
- **Allowed:** `src/xlextract/core/structure/**`, `src/xlextract/adapters/calamine_adapter.py`, `tests/adapter/**`, `fixtures/generators/**`.
- **Forbidden:** `cli/**`, `tui/**`, `core/extraction/**`, `core/discovery/**`.
- **Requirements:** ports as `typing.Protocol`, documented with pre/postconditions; `WorkbookFacts`/`SheetFacts` per §5.5 including date system, sheet visibility (visible/hidden/very_hidden), macros, external links, encryption; the conformance suite as a reusable parametrized base every adapter must pass; calamine adapter must **refuse** a sheet whose estimate exceeds its guard rather than attempting it; capability report declares what the adapter can and cannot do (e.g. "no formula text", "no per-row hidden flags").
- **Tests:** conformance suite; F01–F03, F09, F23–F26 inspect correctness; the memory-guard refusal test; a capability-matrix snapshot test.
- **Stop if:** the port shape cannot express what the COM adapter will need (design it now, with S-009's constraints in mind); calamine cannot report a fact the model requires (record the gap, do not fake the value).
- **Handoff:** ports + conformance suite that T-004 and S-007/S-009 build against.

---

### S-004 — Extraction planner, typing engine, and the row ledger
- **Goal:** the heart of the product: turn facts + contract into an immutable plan; convert raw cells into trustworthy typed values; account for every row.
- **Why Sonnet:** this is where fidelity and zero-silent-loss are won or lost; the rules interact subtly.
- **Prerequisites:** S-003.
- **Target node:** `core.extraction.planner|typing`, `core.evidence.ledger`. **Domain:** extraction, evidence. **Journey:** J2.
- **Allowed:** `src/xlextract/core/extraction/**`, `src/xlextract/core/evidence/**`, `tests/unit/extraction/**`, `tests/property/**`, `fixtures/generators/fidelity.py`.
- **Forbidden:** `adapters/**` (use fakes), `sinks/**`, `cli/**`.
- **Requirements:** the §12 selection table implemented as data, not branches, and emitted into the manifest as `adapter_reason`; typing rules exactly per §5.14 — **declared types win; text-typed Excel cells never become numeric; error sentinels and missing cached values are always flagged**; ledger buckets and the reconciliation identity per §5.15, accrued in the same code path that drops a row.
- **Tests:** decision-table test over the full (format × size × protection × flags) matrix; the F13/F14/F15/F16/F17 fidelity goldens with per-cell expectations; property test: the ledger identity holds for randomized sheets × randomized policies (≥1000 cases); property test: no text-typed cell becomes numeric.
- **Acceptance evidence:** F14 (leading zeros) and F11 (missing cached values) pass with zero silent coercions; the ledger property test is green.
- **Stop if:** a policy combination has no obviously correct behavior (propose, do not choose silently); the reconciliation identity cannot be made to hold for some policy (that is an architecture finding, escalate).
- **Handoff:** `ExtractionPlan` + `TypedRowBatch` + `Ledger` consumed by S-005.

---

### S-005 — Run orchestration, state machine, cancellation, and the verify stage
- **Goal:** the single owner of run lifecycle: stages, timings, terminal-state guarantee, cancellation, and the reconciliation/verify step that gates success.
- **Why Sonnet:** state machines, cancellation semantics, and crash-safety are exactly the class of work that must not be improvised.
- **Prerequisites:** S-004, T-005.
- **Target node:** `app.run`, `app.services.extract|verify`. **Domain:** orchestration. **Journey:** J2, J5, J6, J7.
- **Allowed:** `src/xlextract/app/**`, `tests/component/app/**`, `tests/e2e/**`.
- **Forbidden:** `cli/**`, `tui/**`, `adapters/**`, `core/**` (consume only).
- **Requirements:** the §7.6 machine, exactly five terminal states, exhaustively tested; cancellation token checked only at chunk boundaries; every exit path (success, error, cancel, crash) produces a terminal state and either a published run or a clearly-marked `.partial`/`_failed` directory; verify stage re-checks the ledger identity, re-hashes outputs, re-hashes the source; `app.wiring` is the **only** place adapters and sinks are constructed.
- **Tests:** exhaustive transition test; cancellation injected at every stage boundary; property test "every run reaches exactly one terminal state"; fault injection: kill between write and rename (§14 recovery layer).
- **Stop if:** two components appear to own run state or cancellation; verify cannot distinguish a published run from a partial one by structure alone.
- **Handoff:** services that CLI (T-006) and TUI (T-012) call unchanged.

---

### S-006 — Deterministic structure discovery engine
- **Goal:** candidate generation, scoring, header-strategy detection, trimming, and explicit ambiguity — the "safe rather than magical" component.
- **Why Sonnet:** heuristics that must be deterministic, explainable, bounded, and willing to refuse.
- **Prerequisites:** S-004.
- **Target node:** `core.discovery.*`. **Domain:** discovery. **Journey:** J1, J2.
- **Allowed:** `src/xlextract/core/discovery/**`, `tests/component/discovery/**`, `tests/golden/discovery/**`, `fixtures/generators/messy.py`.
- **Forbidden:** `adapters/**`, `app/**`, `cli/**`.
- **Requirements:** the §5.11 priority order; bounded scan window (configurable, default 200×200); every candidate carries score, basis, supporting evidence, and every rejected candidate carries `rejected_because`; determinism guaranteed (no set iteration order, no dict-order dependence, no time/random input); ties within epsilon → `E_DISCOVERY_AMBIGUOUS`; **never** returns a range outside the sheet; handles F04–F09, F22.
- **Tests:** golden candidate sets for F03–F09 and F22; determinism (100 repeats, identical output); ambiguity test on F22; property test for in-bounds ranges; a "no business meaning" test asserting no output field names a business concept.
- **Stop if:** a fixture requires business knowledge to resolve (that is by definition `NEEDS_INPUT`, not a better heuristic).
- **Handoff:** `Candidate[]` consumed by the planner, the TUI target screen, and `contract infer`.

---

### S-007 — Streaming adapter, memory governor, chunk pipeline
- **Goal:** bounded-memory extraction of arbitrarily large `.xlsx/.xlsm`, with progress and cancellation.
- **Why Sonnet:** performance/memory engineering plus adapter-fallback semantics.
- **Prerequisites:** S-005, T-005.
- **Target node:** `adapters.openpyxl_stream`, `core.extraction.chunking`. **Domain:** extraction. **Journey:** J3.
- **Allowed:** `src/xlextract/adapters/openpyxl_adapter.py`, `src/xlextract/core/extraction/chunking.py`, `tests/perf/**`, `tests/adapter/**`.
- **Forbidden:** `cli/**`, `core/discovery/**`, other adapters.
- **Requirements:** `read_only=True, values_only=True`; `reset_dimensions()` when the declared dimension is untrusted (F09); explicit workbook close; chunk = row group = progress = cancel point = ledger accrual; RSS governor with soft/hard thresholds per §13.4; one bounded fallback from calamine, recorded as `W_ADAPTER_FALLBACK`; per-row hidden flags surfaced for the hidden-row policy.
- **Tests:** adapter conformance suite; peak-RSS flatness across 50/250/500 MB inputs (the §16 headline number); F20 throughput; F09 phantom-dimension; cancellation mid-file; governor soft-breach chunk reduction.
- **Acceptance evidence:** a published `benchmarks/results/` entry comparing calamine vs streaming on F20 (time and peak RSS) that justifies the §12 threshold of 0.6 × budget.
- **Stop if:** the flat-memory target cannot be met with openpyxl (escalate with numbers before reaching for another dependency).
- **Handoff:** the bounded path + the empirical evidence behind the adapter decision table.

---

### S-008 — Encrypted-workbook path
- **Goal:** password-protected workbooks decrypted in memory, with airtight secret hygiene.
- **Why Sonnet:** a security boundary; secret handling errors are silent and serious.
- **Prerequisites:** S-003.
- **Target node:** `adapters.encrypted`. **Domain:** source/adapters. **Journey:** J4.
- **Allowed:** `src/xlextract/adapters/encrypted_adapter.py`, `tests/adapter/test_encrypted.py`, `fixtures/generators/encrypted.py`.
- **Forbidden:** `cli/**` beyond declaring `--password-env`, other adapters.
- **Requirements:** msoffcrypto → `io.BytesIO` → direct adapter; **never** write a decrypted file to disk; password from env var or interactive prompt only (never `argv`); password never logged, never in the manifest, never in an exception message; wrong password → `E_SRC_ENCRYPTED_BAD_KEY`; no password → `E_SRC_ENCRYPTED_NO_KEY` (exit 4) with the action text; memory budget still applies to the decrypted stream.
- **Tests:** F27 round-trip; wrong-password mapping; a log-scrubbing test asserting the fixture password appears nowhere in any produced artifact; a filesystem test asserting no decrypted artifact exists during or after the run.
- **Stop if:** a workbook decrypts but the format is then unreadable (classify honestly, do not retry blindly).
- **Handoff:** encrypted path + the secret-hygiene test pattern reused by S-009.

---

### S-009 — Excel COM adapter (the authorized protected path)
- **Goal:** read workbooks that only the user's authenticated Excel can open, without ever disturbing the user's Excel.
- **Why Sonnet:** the highest-risk component in the system — process lifecycle, ownership semantics, modal/busy states, and a hard no-kill rule.
- **Prerequisites:** S-005, S-003, T-011 (doctor probes available).
- **Target node:** `adapters.excel_com.*`. **Domain:** adapters. **Journey:** J4.
- **Allowed:** `src/xlextract/adapters/excel_com/**`, `tests/real_excel/**`, `docs/protected-workbooks.md`.
- **Forbidden:** every other adapter, `core/**`, `cli/**`, `app/**` (except registering in `wiring`).
- **Requirements:** the §12 COM safety contract in full; distinct `OwnedExcelSession` and explicitly requested `AttachedExcelSession` modes, recorded as `ownership ∈ {owned, attached}`. The owned mode is preferred and quits only its own instance after restoring its own temporary settings. Attached mode alters no global `Application` setting, opens/closes/quits nothing, and only reads an already-open/read-only target. Use `Value2` rectangular block reads (~50k cells, tuned by benchmark), bounded busy/modal retry, and release on every exit path; distinguish IRM from corruption and password protection; stage volatile shares only in owned mode; document no-kill and lawful authorized read.
- **Tests:** conformance suite (Windows+Excel-gated); **no-kill assertion** (no process-termination API is reachable from this package — AST + runtime); attached-instance survival plus no-settings-mutation tests; owned-instance settings-restored tests after success and forced failure; busy/modal retry; cleanup after an injected crash; throughput benchmark.
- **Acceptance evidence:** the attached-instance survival test, plus a **manual F32 run on the real target machine** recorded in `LESSONS.md` with date, Excel version, and outcome. CI cannot sign this off.
- **Stop if:** any requirement appears to need killing a process, saving a workbook, changing a machine security setting, or bypassing protection. **Escalate immediately — this is a hard boundary, not a trade-off.**
- **Handoff:** the COM path + an honest statement in `CURRENT_STATE.md` of exactly what has and has not been proven.

---

### S-010 — Adversarial review and fault-injection suite
- **Goal:** try to break the finished system the way reality will.
- **Why Sonnet:** requires adversarial thinking across component boundaries.
- **Prerequisites:** P0–P7 merged.
- **Target node:** system-wide (tests only). **Journey:** J5, J6.
- **Allowed:** `tests/recovery/**`, `tests/property/**`, `LESSONS.md`, `CURRENT_STATE.md`.
- **Forbidden:** `src/**` — findings are reported as issues with failing tests, **not** fixed in this task (scope control).
- **Requirements:** inject crash at each stage; disk full during write and during publish; output path removed mid-run; source modified/deleted mid-run; Excel busy/modal/closed by user mid-read; duplicate concurrent runs of the same source; kill during publication; truncated zip fuzzing; a 48-hour-equivalent repeat-run determinism check. For each: assert the terminal state, the exit code, that no `.partial` masquerades as success, and that the source is unchanged.
- **Acceptance evidence:** every scenario has a test and a documented expected behavior; all findings filed with reproducers.
- **Stop if:** a finding requires an architecture change (escalate, do not refactor).
- **Handoff:** a hardening report driving T-013's release gate.

---

## 21. Codex Terra Execution Packs

Terra packs contain **no architectural invention**. Every contract, file, and behavior is specified; a pack that requires a design decision is a defect in this plan and must be escalated.

---

### T-001 — Repository skeleton and toolchain
- **Goal:** a working, empty-but-correct project: `pyproject.toml`, src layout, extras `[com,crypto,tui,dev]`, console scripts `xlextract`/`xlx`, ruff + mypy (strict) + pytest config, a `tools/ci.ps1` running the full gate set.
- **Why Terra:** fully bounded, zero ambiguity.
- **Prerequisites:** none. **Target node:** repo root. **Journey:** none.
- **Allowed:** `pyproject.toml`, `tools/**`, `src/xlextract/__init__.py`, `tests/conftest.py` (skeleton), `.gitignore`, `README.md`.
- **Forbidden:** any component module.
- **Requirements:** exact directory tree of §10 with `__init__.py` where needed; pinned dependency floors from §0.4; require Python `>=3.12,<3.13`; PyInstaller available in the dev toolchain for T-014; mypy strict on `src/`; ruff with an explicit rule set; `xlx --version` works.
- **Tests:** three independent gates: source-tree `python -m xlextract --version`; installed CLI `xlx --version` after a standard-user offline install; and the packaged executable only in T-014/T-013 (`dist/.../xlx.exe`). Source-tree tests must never require the installed CLI. `tools/ci.ps1` exits 0 on the empty repo only when its required developer tooling is actually available.
- **Stop if:** a dependency cannot be installed as a standard user without admin rights.

---

### T-002 — Error taxonomy, exit codes, result envelope
- **Goal:** implement §7.7 as code: `XlxError` hierarchy, the code→family→exit-code map, and the human/JSON renderers for errors.
- **Prerequisites:** S-001. **Target node:** `core.errors`, `cli.renderers`. **Journey:** J5.
- **Allowed:** `src/xlextract/core/errors/**`, `src/xlextract/cli/renderers/**`, `tests/unit/errors/**`.
- **Forbidden:** everything else.
- **Requirements:** every code in §7.7 defined exactly once with a stable string, a one-sentence human message, and an `action`; a single mapping table code→exit code; unexpected exceptions wrapped as `E_INTERNAL` with a run reference; tracebacks only under `--debug`; renderers contain zero logic beyond formatting.
- **Tests:** exhaustive test that every code maps to exactly one family and exit code; snapshot tests of rendered output (human + JSON) for one code per family; a test that no code string is duplicated.
- **Stop if:** a situation in this plan has no matching code (report it; do not add one unilaterally).

---

### T-003 — Project Eye bootstrap and gates
- **Goal:** create the five documents and `.project-eye/`, plus `tools/eye.py` (check/sync) and `tools/gates.py` (R9, R10, R12, R13, R16, R17 AST/grep gates) and the import-linter config for R1–R8, R11.
- **Prerequisites:** T-001. **Target node:** `.project-eye`, `tools`. **Journey:** none.
- **Allowed:** `AGENTS.md`, `ARCHITECTURE.md`, `PROJECT_EYE.md`, `CURRENT_STATE.md`, `LESSONS.md`, `.project-eye/**`, `tools/eye.py`, `tools/gates.py`, `.importlinter`.
- **Forbidden:** `src/**`.
- **Requirements:** documents seeded from §18 of this plan (content is specified there; do not invent architecture); `eye check` compares INTENDED vs STATIC ACTUAL and reports orphans, unmapped modules, cycles, and owner conflicts; `eye sync` regenerates `manifest.json` only (never edits `graph.yaml`); every gate returns a non-zero exit and a precise message on violation. `tools/gates.py` must enforce the exact project-import allowlists of R19–R21, including diagnostics.
- **Tests:** a deliberately violating fixture package under `tests/fixtures_arch/` for each of R1–R21, asserting the gate catches it.
- **Stop if:** a rule in §11 cannot be checked mechanically (report it; it becomes a documented review-only rule).

---

### T-004 — `inspect` command and human renderer
- **Prerequisites:** S-003, T-002. **Target node:** `cli.commands.inspect`. **Journey:** J1.
- **Allowed:** `src/xlextract/cli/**`, `src/xlextract/app/services/inspect.py`, `tests/cli/test_inspect.py`.
- **Forbidden:** `adapters/**`, `core/**` (consume only).
- **Requirements:** `xlx inspect <path> [--json]`; human output = a sheets table (name, visibility, dimensions, trusted?), a Tables/named-ranges section, a workbook-flags section, and a candidates section; `--json` emits exactly one `InspectResult` document on stdout and nothing else; all human output and progress to stderr; no side effects, no run directory.
- **Tests:** snapshot tests for F01–F03, F09, F23–F26; a test asserting stdout is valid JSON and nothing else under `--json`; a test asserting no file is created anywhere.
- **Stop if:** a fact needed for the display is not in `WorkbookFacts`.

---

### T-005 — Sinks and the atomic publisher
- **Prerequisites:** S-001. **Target node:** `sinks.*`, `core.publication.publisher`. **Journey:** J2, J3.
- **Allowed:** `src/xlextract/sinks/**`, `src/xlextract/core/publication/**`, `tests/unit/sinks/**`, `tests/recovery/test_publish.py`.
- **Forbidden:** `adapters/**`, `core/extraction/**`, `cli/**`.
- **Requirements:** define the publication-owned `Sink` port in `core/publication/ports.py`; `sinks/` implements it without importing `core.structure.ports` or other core internals. Implement `ParquetSink` (one row group per chunk, zstd, explicit Arrow schema from the schema document — **never** inferred by the writer), `CsvSink` (UTF-8 BOM, RFC4180, records `lossy: true` in the manifest), `JsonlSink`, `PreviewSink` (first 100 rows); the run-directory layout exactly as §7.2; build in `<id>.partial/`, `manifest.json` written last, fsync files and directory, `os.replace` the directory; pre-flight free-space check of 3× estimated output; cleanup registered before the first write.
- **Tests:** golden bytes for F01; type-fidelity round-trip (write→read→compare) for F13–F17 values; row-group boundary test; kill-between-write-and-rename leaves no successful-looking directory; disk-full simulation → `E_OUT_DISK_FULL`.
- **Stop if:** an Arrow type cannot represent a value the typing engine produces (report the pair; do not silently downcast).

---

### T-006 — `extract`, `verify`, `explain`, `runs` commands
- **Prerequisites:** S-005. **Target node:** `cli.commands.*`. **Journey:** J2, J5, J6, J7.
- **Allowed:** `src/xlextract/cli/**`, `tests/cli/**`.
- **Forbidden:** `core/**`, `adapters/**`, `sinks/**`, `app/services/**` (call only).
- **Requirements:** every flag in §8 implemented and documented; folder mode with `--glob` and `--continue-on-error`; progress to stderr (`--progress json` = JSONL); exactly one JSON result document on stdout under `--json`; exit codes from the T-002 map; Ctrl-C → J6; `--dry-run` prints the plan without reading data; `--help` for each command includes a worked example; shell completion generated.
- **Tests:** exit code per error family; stdout-purity test under `--json`; Ctrl-C test asserting `CANCELLED` + exit 8 + no published directory; `--dry-run` asserts zero data reads; `--help` snapshots.
- **Stop if:** a flag would require logic not already in a service.

---

### T-007 — Fixture corpus generators
- **Prerequisites:** T-001. **Target node:** `fixtures.generators`. **Journey:** none.
- **Allowed:** `fixtures/generators/**`, `fixtures/committed/**`, `tools/fixtures.py`, `tests/unit/test_fixtures.py`.
- **Forbidden:** `src/**`.
- **Requirements:** deterministic, seeded builders for F01–F31 exactly as specified in §15 (F32 is environment-local and must **not** be generated or committed); `python -m tools.fixtures build` regenerates everything reproducibly; generated files are git-ignored; each fixture ships a sidecar `expected.json` describing its ground truth (row counts, header location, known anomalies) for use by golden tests; F20 (250 MB) is generated on demand, not in the default build.
- **Tests:** every fixture opens in at least one adapter; every `expected.json` validates against a schema; byte-identical regeneration from the same seed.
- **Stop if:** a fixture in §15 cannot be produced with the available writers (report which and why).

---

### T-008 — `contract init|validate|infer`
- **Prerequisites:** S-006, T-004. **Target node:** `cli.commands.contract`. **Journey:** J1.
- **Allowed:** `src/xlextract/cli/commands/contract.py`, `src/xlextract/app/services/contract.py`, `tests/cli/test_contract.py`, `docs/contract-reference.md`.
- **Forbidden:** `core/discovery/**`, `adapters/**`.
- **Requirements:** `init` writes a commented template; `validate` checks a file against the JSON Schema with precise, line-referenced errors; `infer <workbook>` runs inspect+discovery and writes a contract whose every field is explicit (no `auto`, no `discover`) with a header comment recording which values came from discovery and therefore need human review.
- **Tests:** `infer` output round-trips through `validate`; extracting with an inferred contract yields the same dataset as the discovery-driven extraction for F03–F08; invalid-contract error messages are snapshot-tested.
- **Stop if:** discovery output cannot be expressed as a contract field.

---

### T-009 — Benchmark harness and performance gate
- **Prerequisites:** S-007. **Target node:** `benchmarks`, `tests.perf`. **Journey:** J3.
- **Allowed:** `benchmarks/**`, `tests/perf/**`, `tools/bench.py`.
- **Forbidden:** `src/**`.
- **Requirements:** every scenario in §16 as a named, runnable benchmark; 3 runs after a warm-up, median reported; peak RSS via `psutil` sampled at chunk boundaries and process exit; results written to `benchmarks/results/<date>-<machine>.json`; a pytest perf gate that fails on a >20% regression against the committed baseline; a markdown comparison report (calamine vs streaming vs COM).
- **Tests:** the gate itself is tested with a synthetic regression.
- **Stop if:** a §16 threshold is unachievable (report the measured number; do not adjust the threshold unilaterally).

---

### T-010 — Format coverage and the capability matrix
- **Prerequisites:** S-003, S-007. **Target node:** `tests.adapter`, `docs`. **Journey:** J2.
- **Allowed:** `tests/adapter/test_formats.py`, `docs/formats.md`, `fixtures/generators/legacy.py`.
- **Forbidden:** `src/**` — if a format gap is found, file it; do not fix it here.
- **Requirements:** run the conformance suite over F23 (.xls), F24 (.xlsb small + over-budget), F25 (.xlsm), F28 (corrupt), F29 (wrong extension) for every applicable adapter; produce a generated capability matrix (format × adapter × capability) published to `docs/formats.md`; assert the over-budget `.xlsb` case fails with `E_MEM_BUDGET_EXCEEDED` and an actionable message rather than an OOM.
- **Tests:** the matrix is snapshot-tested so a silent capability change fails CI.
- **Stop if:** an adapter silently degrades rather than reporting a capability gap.

---

### T-011 — `doctor`
- **Prerequisites:** T-002. **Target node:** `diagnostics.*`, `cli.commands.doctor`. **Journey:** none.
- **Allowed:** `src/xlextract/diagnostics/**`, `src/xlextract/cli/commands/doctor.py`, `tests/cli/test_doctor.py`.
- **Forbidden:** `core/**`, `adapters/**` except a capability query.
- **Requirements:** probes per §5.24, each returning `{name, status ∈ ok|warn|fail|skipped, detail, action}`; **strictly read-only** — no registry writes, no setting changes, no installs; `--json`; exit 0 if no `fail`; the COM probe creates an instance it owns and releases it immediately, and is skipped cleanly when Excel is absent.
- **Tests:** each probe's failure mode simulated; an assertion that `doctor` performs no write anywhere (reuse the integrity guard, extended to the registry via a mock).
- **Stop if:** a useful check would require modifying the machine (drop it and document why).

---

### T-014 — Early portable packaging smoke
- **Goal:** prove early that a portable Windows 10/11 PyInstaller one-dir distribution can carry the first working `doctor` and `inspect` slice.
- **Prerequisites:** T-004, T-011. **Target node:** packaging. **Journey:** J1.
- **Allowed:** `tools/build_exe.ps1`, `tests/packaging/**`.
- **Forbidden:** `src/**`, `core/**`, `adapters/**`, `sinks/**`, `cli/**`, `docs/**`.
- **Requirements:** build one directory from the Python 3.12 environment with no runtime network access or administrator rights; include the installed direct-reader/output/COM dependencies needed by the portable distribution (`python-calamine`, `pyarrow`, and pywin32 when the COM extra is installed); run the packaged `xlx doctor` and packaged `xlx inspect` against F01. Record the exact build command and module-collection choices in the test evidence. This smoke is local; T-013 repeats it on a clean standard-user Windows 10/11 machine offline.
- **Tests:** a packaging smoke test fails if build output is missing, `doctor` fails, `inspect F01` fails, or either command attempts a network connection.
- **Stop if:** PyInstaller cannot bundle any required dependency or the output needs admin rights, a runtime download, or an installed Python interpreter. Report the exact module and failure; do not add a second packaging framework.

---

### T-012 — TUI
- **Prerequisites:** S-005, S-006, T-006. **Target node:** `tui.*`. **Journey:** J1, J2, J6.
- **Allowed:** `src/xlextract/tui/**`, `tests/tui/**`.
- **Forbidden:** `core/**`, `adapters/**`, `sinks/**`, `cli/**`.
- **Requirements:** the seven screens and key map of §9; calls **only** `app.services`; the result screen shows the equivalent CLI command and supports copying it; cancellation wired to the same token; no business logic, no direct file access.
- **Tests:** Textual pilot drives J1→J2 end to end on F01 and F04; a cancel test; an import test asserting `tui` imports nothing from `adapters`, `sinks`, `core.extraction`, or `cli`.
- **Stop if:** a screen needs data no service returns (request a service field; do not reach into core).

---

### T-013 — Documentation, packaging, release gate
- **Prerequisites:** P0–P7 merged, S-010 report, T-014. **Target node:** `docs`, `tools.release`. **Journey:** none.
- **Allowed:** `docs/**`, `README.md`, `tools/release.py`, `tools/build_exe.ps1`.
- **Forbidden:** `src/**`.
- **Requirements:** `docs/user-guide.md` ≤2 pages for a non-technical user (five tasks, screenshots of the TUI); `cli-reference.md` and `contract-reference.md` **generated** (typer + schemas) so they cannot drift; `integration.md` for downstream tools/agents (run-directory layout, exit codes, result envelope, stability promises); repeat and verify T-014's one-dir PyInstaller build on a clean standard-user Windows 10/11 machine with no admin rights and no internet; `tools/release.py` runs every gate in §23 and refuses to tag if any fails; **one authoritative location per rule** — the docs link to it, never restate it.
- **Tests:** a link/duplication check asserting no rule is stated in two documents; a smoke test running the built executable's `doctor` and `inspect` on F01.
- **Stop if:** packaging requires admin rights or a network fetch at runtime.

---

## 22. Dependency / Merge Order

```
T-001 ──┬─► S-001 ──┬─► T-002 ──► T-004 ──► T-006 ──► T-008 ──► T-012
        │           ├─► S-002 ──► S-003 ──┴─► S-004 ──► S-005 ──► S-007 ──► T-009
        ├─► T-003    └─► T-005 ─────────────────┘        │
        └─► T-007 ───────────────────────────────────────┤
                                                S-006 ───┘
                                      S-008 ──► T-010
                            T-004 ──┐
                                      ├─► T-014 ────────────────────► T-013
                            T-011 ──┴─► S-009
                                                  └────► S-010 ─────┘
```

| Wave | Parallel lanes (branch = `lane/<id>`) | Owned files | Integration point |
|---|---|---|---|
| **W0** | T-001 alone | repo root | everything blocks on it |
| **W1** | T-003 (`.project-eye/`, `tools/`) ‖ S-001 (`contracts/`, `schemas/`) ‖ T-007 (`fixtures/`) | disjoint by directory | CI green on an empty repo |
| **W2** | S-002 (`core/source/`) ‖ T-002 (`core/errors/`, `cli/renderers/`) | disjoint | error taxonomy available to source layer |
| **W3** | S-003 (`core/structure/`, `adapters/calamine`) ‖ T-005 (`sinks/`, `core/publication/`) | disjoint | ports + sinks both exist |
| **W4** | S-004 (`core/extraction/`, `core/evidence/`) — **alone**; it is the semantic center | — | plan+typing+ledger frozen |
| **W5** | S-005 (`app/`) ‖ T-004 (`cli/inspect`) | disjoint | first end-to-end extract |
| **W6** | S-006 (`core/discovery/`) ‖ T-006 (`cli/` commands) ‖ T-011 (`diagnostics/`) | disjoint | discovery reachable from CLI |
| **W7** | S-007 (`adapters/openpyxl`) ‖ S-008 (`adapters/encrypted`) ‖ T-008 (`cli/contract`) ‖ T-014 (`tools/build_exe.ps1`, `tests/packaging/`) | disjoint adapters/tooling | format coverage + early portable-build evidence |
| **W8** | S-009 (`adapters/excel_com/`) ‖ T-009 (`benchmarks/`) ‖ T-010 (`tests/adapter`) | disjoint | COM gate attempt |
| **W9** | T-012 (`tui/`) ‖ S-010 (`tests/recovery/`) | disjoint | feature complete |
| **W10** | T-013 alone | `docs/`, `tools/release.py` | release gate |

**Semantic-collision rules (beyond Git conflicts).**
1. Only one task may own a directory at a time; ownership is declared in the lane table and in `CURRENT_STATE.md` before work starts.
2. **Nobody edits `contracts/` after S-001 except through a contract-change review** — a contract change invalidates every in-flight lane.
3. S-004 runs alone because a change to the ledger or typing semantics silently invalidates the goldens of every parallel lane.
4. Any lane that discovers it needs a field in another lane's contract **stops and escalates**; it does not add the field.
5. Merge order within a wave: contracts → core → adapters/sinks → app → cli/tui → docs. After every merge: `tools/ci.ps1` + `eye check` must be green before the next merge.
6. **Review points:** end of W1 (contract review), end of W4 (semantic-core review), end of W8 (COM/security review), end of W10 (release review). All four are Opus-level.

---

## 23. Verification Gates

Each gate is a named, runnable check. A phase is not complete until its gates pass; a release requires all of them.

| Gate | Passes when | Runnable |
|---|---|---|
| **G1 Architecture** | import-linter R1–R8/R11 green; AST gates R9/R10/R12/R13/R16/R17/R19–R21 green | `python -m tools.gates arch` |
| **G0 Packaging smoke** | T-014 one-dir build runs `doctor` and `inspect F01` locally without admin, runtime network, or an installed Python interpreter | `powershell -File tools/build_exe.ps1 -Smoke` |
| **G2 Contract** | schema snapshots unchanged (or a major bump is declared); every example validates; N−1 read works | `pytest tests/contract` |
| **G3 Source safety** | `source_integrity_guard` green across the full suite; no write-mode open against a source; F31 handled | `pytest -m safety` |
| **G4 Zero-loss** | ledger identity holds in every e2e run and in ≥1000 property cases | `pytest tests/property/test_ledger.py tests/e2e` |
| **G5 Large-file** | F20 within time budget; peak RSS flat (±15%) across 50/250/500 MB | `pytest -m perf -k large` |
| **G6 CLI** | exit-code matrix correct; stdout is pure JSON under `--json`; `--help` snapshots current | `pytest tests/cli` |
| **G7 COM** | Windows+Excel suite green **incl. attached-instance survival and no-kill**; F32 manually verified on the target machine and recorded in `LESSONS.md` | `pytest -m real_excel` + manual record |
| **G8 Performance** | no §16 threshold regressed >20% vs the committed baseline | `pytest -m perf` |
| **G9 Recovery** | every fault-injection scenario ends in a correct terminal state with an unchanged source and no false success | `pytest tests/recovery` |
| **G10 Security** | no network import; `no_network_guard` green; no secret in any artifact; no machine setting modified by any test | `python -m tools.gates security && pytest -m security` |
| **G11 Map / anti-drift** | INTENDED == STATIC ACTUAL == RUNTIME ACTUAL; no orphans, no unmapped modules, no owner conflicts; `CURRENT_STATE.md` updated in the same change | `python -m tools.eye check --all` |
| **G12 Release** | G1–G11 green; docs generated and link-checked; one-dir build runs `doctor` + `inspect` on a clean standard-user machine offline | `python -m tools.release check` |

### 23.2 Stop conditions for implementation agents (restated, binding)

Sonnet or Terra must **stop and return to Opus-level review** when: ownership is unclear · a public contract must change · an architecture boundary must change · a cross-domain dependency is required · security behavior is unclear · a data-loss risk appears · a large refactor becomes necessary · runtime reality contradicts the map · a supposedly optional capability turns out to be required by core. They must not quietly redesign the system.

**Scope control:** an unrelated problem is `RECORD_ONLY` (appended to `CURRENT_STATE.md`) unless it blocks the current task or presents a critical security/data-loss risk.

**Anti-drift:** a material change is incomplete until code, tests, public contracts, architecture rules, `PROJECT_EYE.md`, `graph.yaml`, and `CURRENT_STATE.md` all agree. A code change with a stale map is incomplete; a map update unsupported by code is also incomplete.

---

## 24. Definition of Done

"Tests pass" is not done. V1 is done when the following **evidence** exists and is linked from `CURRENT_STATE.md`:

1. **Extracts correctly** — golden datasets for F01–F31 match per-cell expectations, including F13 (dates), F14 (leading zeros), F15 (number formats), F16 (Unicode). Evidence: golden test report.
2. **Does not modify sources** — `source_integrity_guard` green across the entire suite including the COM suite. Evidence: a before/after hash manifest for every fixture.
3. **Handles representative large files** — F20 extracted in ≤5 min with flat, bounded memory, on the reference machine, with antivirus active. Evidence: `benchmarks/results/`.
4. **Produces stable interoperable outputs** — a downstream consumer that has never seen Excel reads the run directory using only `schemas/` and `docs/integration.md`. Evidence: an independent reader script in `tests/e2e/test_downstream.py` that uses no `xlextract` import.
5. **Explains failures** — every error family produces a code, a human sentence, an action, and a run reference; no traceback in default mode. Evidence: the exit-code/message matrix test.
6. **Reconciles** — every run in the e2e suite satisfies the ledger identity, and `verify` re-proves a published run from disk alone. Evidence: G4.
7. **Works offline** — full suite green with sockets blocked. Evidence: G10.
8. **Works as a standard user** — install and full run on a clean Windows 11 account without admin rights, without internet, using the one-dir build. Evidence: a dated manual checklist in `LESSONS.md`.
9. **Clear UX** — a non-technical colleague completes an extraction in the TUI unaided, and the CLI `--help` is sufficient for an engineer to script it. Evidence: a dated usability note in `LESSONS.md`.
10. **Architecture intact** — G1 and G11 green; zero declared-vs-actual drift.
11. **Protected path honestly reported** — either F32 verified on the real target machine (date, Excel version, outcome in `LESSONS.md`), or `CURRENT_STATE.md` states plainly that enterprise DRM support is **unproven**. Claiming it without F32 is a defect.

---

## 25. Risks and Deferred Work

### BLOCKER (must be resolved before the dependent phase starts)
- **B1 — No real protected workbook available for F32.** Blocks G7 sign-off, not P6 development. *Mitigation:* build the COM path against synthetic + password fixtures, keep `CURRENT_STATE.md` honest, schedule a target-machine session before release. **Owner: the human, not an agent.**
- **B2 — Reference machine for §16 not fixed.** All thresholds are meaningless until one is named. Must be pinned at T-009 start.

### V1 REQUIRED (in scope, risk acknowledged)
- **R-A — calamine materializes sheets.** Mitigated by the probe + budget + guard (§12/§13). Residual risk: the probe's estimate is wrong for pathological files → the hard governor catches it, at the cost of one restarted read.
- **R-B — openpyxl streaming may miss the §16 throughput floor on some shapes.** Mitigated by benchmarking in S-007 *before* committing to the threshold; if it fails, escalate with numbers rather than adding a dependency.
- **R-C — COM fragility** (modal dialogs, add-ins, corporate Excel policies, Excel updates). Mitigated by bounded retries, the no-kill rule, and treating COM as optional. Residual risk is real and is disclosed to users in `docs/protected-workbooks.md`.
- **R-D — Discovery disappoints on genuinely messy files.** Mitigated by the explicit-basis model and `NEEDS_INPUT`: the tool is allowed to say "I don't know", which is the feature, not the failure.
- **R-E — Type inference disagreement between runs** when a later chunk contradicts an early inference. Mitigated by inferring over the first N chunks and, on contradiction, widening to text with `W_TYPE_WIDENED` rather than failing or lying.
- **R-F — Antivirus interference** on temp files and Excel automation. Mitigated by staging inside the user profile and by `doctor` reporting slow temp I/O. Never by asking the user to weaken AV.

### V1.1
Arrow IPC and DuckDB sinks · `fastexcel` zero-copy path (only with benchmark justification) · parallel multi-sheet extraction · folder-level parallelism · a contract registry · richer TUI preview (type coloring, flagged-cell navigation) · `verify --against <prior_run>` schema-drift reporting · shell-completion packaging.

### LATER
`.ods` support · formula evaluation · delta/incremental extraction · mid-file resumability (only if §13.6's premise changes) · plugin entry points for third-party adapters · a lightweight external orchestrator that *composes* tools without absorbing them.

### REJECTED (and why, so it is not re-proposed)
Downloading files · business KPI logic · dashboards · a scheduler · a server/daemon mode · a central "platform" application · pandas/polars in core · a database-backed run store · automatic COM fallback · writing Excel files · any telemetry · any AI call on source data · mid-file resume in V1 · killing Excel processes under any circumstance.

---

## Appendix A — Confirmed decisions

The following are approved for V1:
1. **Python 3.12 only.** The packaging target is a portable Windows 10/11 one-dir distribution; it is smoke-tested after the first `doctor`/`inspect` slice and release-tested on a clean standard-user machine.
2. **Tool name:** package `xlextract`, CLI `xlx`.
3. **Parquet is the default machine output;** CSV remains optional for people and integrations.
4. **`.ods` is out of scope** for V1 (calamine supports it; the cost is test-corpus breadth, not code).
5. **Single-file-at-a-time** extraction; folder mode is sequential in V1.
6. **The reference machine** in §16 is a plausible corporate laptop, not a measured one (see B2).

---

*End of plan. The correction pass is complete only when the six authoritative views agree; implementation then begins with T-001 and no later task.*
