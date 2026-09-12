# PROJECT_EYE.md — the living map

Concise by mandate. **Keep this file under ~250 lines.** If it outgrows that, the architecture is too big — not the document.

Machine-readable truth: `.project-eye/graph.yaml`. Rationale: `ARCHITECTURE.md`. Today's reality: `CURRENT_STATE.md`.

---

## LEVEL 1 — SYSTEM

One offline, single-process Windows CLI/TUI tool.

```
Excel file(s) [+ extraction contract]  →  xlextract  →  run directory
                                                        + structured stdout
                                                        + exit code
```

No network. No daemon. No database. No shared state between runs. Composes with other tools only through the run directory, the result envelope, and exit codes.

---

## LEVEL 2 — DOMAINS

| Domain | Responsibility | Owns |
|---|---|---|
| `contracts` | the public, versioned data language | schema definitions |
| `source` | identity, safety, format detection, staging | source bytes & hashes |
| `structure` | the physical workbook fact model + ports | workbook facts |
| `discovery` | candidates and header strategies | interpretive proposals |
| `extraction` | planning, reading, typing, chunking | the row stream |
| `evidence` | ledger, lineage, issues, reconciliation | proof |
| `publication` | run directory, atomic publish, sinks | outputs |
| `orchestration` | run lifecycle, state, cancellation | run state |
| `presentation` | CLI & TUI rendering | nothing but text |
| `diagnostics` | environment readiness | nothing |

---

## LEVEL 3 — COMPONENTS

| Node | Purpose | Owner of | Inputs → Outputs | Key deps | Tests |
|---|---|---|---|---|---|
| `contracts` | public models + JSON Schema | schema definitions | — → models, `schemas/*.json` | pydantic only | contract, golden-schema |
| `core.source.identity` | stable source identity | **source hash** | path → `SourceRef` | hashlib | unit, path-shapes |
| `core.source.safety` | enforce I1; staging; change detection | staging lifecycle | `SourceRef` → read-only handle | — | safety, recovery |
| `core.source.probe` | cheap size estimate (zip central dir) | — | `SourceRef` → `SizeEstimate` | zipfile | unit, estimate-band |
| `core.structure.model` | format-neutral workbook facts | — | — → `WorkbookFacts` | — | unit |
| `core.structure.ports` | `WorkbookInspector`/`RowReader`/`CapabilityReport` | the Excel-reader adapter boundary | — | — | conformance suite |
| `core.discovery.engine` | deterministic candidates + ambiguity | interpretive proposals | facts, contract → `Candidate[]` | — | component, golden, property |
| `core.extraction.planner` | the adapter/target decision | **adapter choice** | facts, candidate, budget → `ExtractionPlan` | — | decision-table |
| `core.extraction.reader` | chunked row streaming + policies | — | plan → `RawRow` batches | ports | component |
| `core.extraction.typing` | fidelity engine | output types | raw cells → typed values | — | fidelity goldens, property |
| `core.evidence.ledger` | enforce I2 | **row counts** | row events → `Ledger` | — | property (identity) |
| `core.evidence.lineage` | anchor + `__src_row` + flagged cells | lineage | plan, rows → lineage | — | e2e (`explain`) |
| `core.evidence.issues` | aggregated warning/error stream | issues | events → `Issue[]` | — | unit (aggregation) |
| `core.publication.ports` | `Sink` | the output-writer boundary | `PublicationBatch`, schema → output bytes | contracts | sink conformance |
| `core.publication.publisher` | atomic publish | **output publication** | datasets → run dir | os | recovery (fault injection) |
| `core.errors` | taxonomy + exit codes | error codes | — | — | exhaustive map |
| `adapters.calamine` | fast reader, all 4 formats | — | `SourceRef` → facts, rows | python-calamine | conformance, memory-guard |
| `adapters.openpyxl_stream` | bounded-memory xlsx/xlsm + metadata | — | `SourceRef` → facts, rows | openpyxl | conformance, RSS flatness |
| `adapters.encrypted` | password decrypt in memory | — | `SourceRef`+key → stream | msoffcrypto | secret hygiene |
| `adapters.excel_com` | authorized read through tool-owned Excel; attached read-only compatibility mode | **Excel session + COM apartment** | `SourceRef` → rows | pywin32 | real_excel (gated) |
| `sinks.*` | parquet / csv / jsonl / preview | output bytes | typed batches → files | publication `Sink` port, pyarrow | golden bytes, round-trip |
| `app.run` | run lifecycle | **run state, cancellation** | request → terminal state | core | state machine, cancel |
| `app.services.*` | Inspect/Extract/Verify/Explain/Doctor | — | typed request → typed result | core, ports | component, e2e |
| `app.wiring` | composition root | **configuration** | config → wired services | all infra | smoke |
| `cli.*` | typer app + renderers | — | argv → exit code | app | cli, exit-code matrix |
| `tui.*` | Textual app | — | keys → service calls | app | pilot, import guard |
| `diagnostics.*` | read-only environment report | — | — → probe results | — | per-probe failure sim |

---

## LEVEL 4 — CODE

See `PLAN.md` §10 for the full tree. Every node above maps to exactly one module path, declared in `.project-eye/graph.yaml`. No `utils/`, `helpers/`, or `common/`.

---

## LEVEL 5 — RUNTIME JOURNEYS

| ID | Journey | Stage sequence |
|---|---|---|
| **J1** | inspect workbook | identity → safety → probe → inspect → discover(candidates) → result |
| **J2** | extract workbook | J1 → plan → read → type → ledger → write → verify → publish |
| **J3** | extract huge workbook | J2 with streaming adapter, chunk governor, progress, cancel points |
| **J4** | protected / COM path | identity → direct attempt → classify → encrypted \| COM(acquire/ownership) → J2 tail → session release |
| **J5** | failure path | any stage → typed error → capture evidence → `_failed/` → render → exit code |
| **J6** | cancel path | token set → finish chunk → close ledger → terminal CANCELLED → exit 8 |
| **J7** | verify / explain | read run dir → re-check ledger identity → re-hash outputs → (optional) re-hash source |

Each run manifest stamps its journey and the stages actually executed; `tools/eye.py check --runtime` compares that against this table.

---

## RUN STATES

```
PENDING → RUNNING → { SUCCEEDED | SUCCEEDED_WITH_WARNINGS | NEEDS_INPUT | FAILED | CANCELLED }
```

Five terminal states, all unambiguous. Stages are observable progress *within* `RUNNING`, not states. `NEEDS_INPUT` is terminal by design — the tool never blocks waiting for a human.

---

## DRIFT STATUS

| View | Source | Last checked |
|---|---|---|
| INTENDED | `.project-eye/graph.yaml` | — |
| STATIC ACTUAL | AST import graph of `src/` | — |
| RUNTIME ACTUAL | stage sequences in test-produced run manifests | — |

**Current:** no code exists yet; the map describes the intended system only. Run `python -m tools.eye check --all` after T-003 lands and record the result here.
