# ARCHITECTURE.md — intended architecture and why

Authoritative for *structure and rationale*. Detailed field lists, task packs, and budgets live in `PLAN.md`; the live map lives in `PROJECT_EYE.md`. Each rule is stated in exactly one place — this file owns §Rules and §Rejected alternatives.

Changes to this document require an Opus-level architecture review.

## 1. Mission

```
Excel file → understand physical structure → safely extract selected data
           → prove what was extracted → publish a stable machine-readable result
```

Every component must justify itself against one of those four verbs.

## 2. Shape — hexagonal (ports & adapters) with a thin application layer

```
        ┌──────────────┐        ┌──────────────┐
        │     CLI      │        │     TUI      │   presentation: rendering only
        └──────┬───────┘        └──────┬───────┘
               └──────────┬────────────┘
                     ┌────▼─────┐
                     │   app    │  use-cases: Inspect / Extract / Verify / Explain / Doctor
                     └────┬─────┘  owns run lifecycle; the ONLY caller of ports
                     ┌────▼─────┐
                     │   core   │  domain: source, structure, discovery, extraction,
                     │          │  evidence, publication, errors + PORT INTERFACES
                     └────┬─────┘
             ┌────────────┼─────────────┐
        ┌────▼───┐   ┌────▼────┐   ┌────▼────┐
        │adapters│   │  sinks  │   │  fsys   │   infrastructure: implements ports
        └────────┘   └─────────┘   └─────────┘
             ▲
        ┌────┴───────────────────────────────┐
        │ contracts (pydantic + JSON Schema) │  depends on NOTHING in the project
        └────────────────────────────────────┘
```

Dependencies point inward only. Composition happens once, in `app.wiring`.

**Why hexagonal here specifically.** The two riskiest parts of this system — the Excel COM session lifecycle and the bounded-memory streaming reader — are exactly the parts that must be replaceable, disableable, and prevented from leaking. Ports turn "the COM implementation must not spread across the project" from a code-review hope into a mechanically checkable fact. That is the entire justification; the cost is one indirection.

**Why a thin app layer.** CLI and TUI must share one core with zero duplicated behavior. If a use-case lives in a CLI callback, the TUI must duplicate it or import it. The app layer makes "no hidden TUI-only behavior" structural rather than disciplinary.

## 3. The three invariants

- **I1 — Source is evidence.** Sources are opened read-only and never written, renamed, moved, or deleted. Proven mechanically by a session-scoped fixture that hashes every fixture file before and after the whole suite.
- **I2 — No silent loss.** `observed == emitted + excluded_structural + excluded_hidden + excluded_filtered + excluded_policy + failed`. Checked in the verify stage; a violation means the run does not succeed.
- **I3 — No silent decision.** Every interpretive choice carries `basis ∈ {contract, excel_table, named_range, configured, discovered, user_confirmed}` in the manifest. Ambiguity is a terminal state, not a coin flip.

## 4. The structural/interpretive split

| Plane | Owns | May assert |
|---|---|---|
| **Structural** | physical facts read from the file | "this is true of the bytes" |
| **Interpretive** | decisions about what to extract | "this is a decision, here is its basis and evidence" |

The extractor may discover that a block *looks* tabular. It must never invent what a business row *means*. Business interpretation is permanently outside this tool.

## 5. Architecture rules (mechanically enforced)

Declared in `.project-eye/rules.yaml`; enforced by `import-linter` (R1–R8, R11), `tools/gates.py` (R9, R10, R12, R13, R16, R17, R19–R21), and `tools/eye.py` (R14, R15, R18). A violation fails the build.

**Layering:** `cli, tui → app → core → contracts`. Infrastructure has no broad `core` privilege: `adapters → contracts + core.structure.ports`; `sinks → contracts + core.publication.ports`; `diagnostics → contracts + core.errors`. The three narrow import allowlists are machine-checked.

| # | Rule |
|---|---|
| R1 | `contracts` imports nothing from `xlextract` |
| R2 | `core` never imports `adapters`, `sinks`, `app`, `cli`, `tui` |
| R3 | `cli`/`tui` never import `adapters`, `sinks`, or `core.extraction` |
| R4 | `tui` and `cli` never import each other |
| R5 | `win32com`/`pythoncom` only inside `adapters/excel_com/` |
| R6 | `openpyxl`, `python_calamine`, `msoffcrypto` only inside `adapters/` |
| R7 | `pyarrow` only inside `sinks/` |
| R8 | `rich`/`textual` only inside `cli`/`tui` |
| R9 | no `print()` outside `cli/renderers` and `tui` |
| R10 | no module-level mutable state outside `app.wiring` |
| R11 | no import cycles |
| R12 | god-module guard: module ≤500 LOC, package ≤12 modules without an Eye entry |
| R13 | every public contract model is defined in exactly one module |
| R14 | orphan guard: every graph node maps to a module, every module ≥50 LOC maps to a node |
| R15 | every declared journey is exercised by ≥1 test |
| R16 | no `socket`/`urllib`/`requests`/`httpx` anywhere in `src/` |
| R17 | no write-mode open against a source path |
| R18 | One Owner Rule: run state, Excel session, publication, configuration, source identity, cancellation each have exactly one owning module, declared in `graph.yaml` |
| R19 | adapters may import only `contracts` and `core.structure.ports` from the project |
| R20 | sinks may import only `contracts` and `core.publication.ports` from the project |
| R21 | diagnostics may import only `contracts` and `core.errors` from the project |

**Drift classes detected:** cycles · god modules · duplicate ownership · duplicate schemas · hidden global state · cross-domain shortcuts · UI→persistence/extraction violations · unmapped runtime paths · orphan components.

**Three views compared:** INTENDED (`graph.yaml`) vs STATIC ACTUAL (AST import graph) vs RUNTIME ACTUAL (stage/adapter sequences harvested from run manifests produced by the test suite). Meaningful disagreement is architecture drift and fails the map gate.

## 6. The One Owner Rule

| State / object | Sole owner |
|---|---|
| run state, cancellation | `app.run` |
| Excel session and COM apartment | `adapters.excel_com.session` |
| output publication | `core.publication.publisher` |
| configuration | `app.wiring` |
| source identity (hashing) | `core.source.identity` |
| ledger counts | `core.evidence.ledger` |

## 7. Adapter strategy in one paragraph

Adapter choice is a **measurement, not a preference**. `core.source.probe` estimates uncompressed sheet bytes from the zip central directory; `core.extraction.planner` applies the decision table (PLAN.md §12) against the memory budget and the security mode, and records `adapter_reason` in the manifest. calamine is fast but materializes a sheet; openpyxl streams but is slower; COM is the only path for IRM/DRM-protected workbooks and is **never** an automatic fallback — it requires explicit opt-in. COM defaults to a dedicated instance owned by the tool. Attached mode is a separately recorded, read-only compatibility mode that must not mutate the user's Excel `Application` settings. Direct→direct fallback is automatic, bounded to one step, and recorded.

## 8. Rejected alternatives (do not re-propose without new evidence)

| Alternative | Rejected because |
|---|---|
| Layered monolith (`readers/ transforms/ writers/`) | nothing stops the CLI importing an adapter; the COM dependency spreads; "optional" stops being optional exactly when COM lands |
| Plugin architecture with entry-point discovery | speculative extensibility; dynamic loading and trust problems for four known adapters; ports already give replaceability *(POSTPONE, not permanent)* |
| Everything through Excel COM | 1–2 orders of magnitude slower; requires Excel; untestable in CI; fragile; uses the user's session for work needing no rights |
| Everything through direct parsing | cannot read the IRM/DRM workbooks the target environment actually contains |
| pandas as the internal data model | its type coercion destroys the fidelity this product exists to protect; forces full materialization; couples our schema to pandas dtypes |
| Library-first, CLI as afterthought | the CLI is the supported contract; the Python API is the byproduct |
| Event-sourced / message-bus internals | enormous ceremony for a single-process batch tool |
| SQLite-backed run store | filesystem run directories are inspectable, portable, atomically publishable, and need no migrations |

## 9. Boundary

**CORE:** source identity & safety · format detection · structure model · discovery · planning · typing · ledger · lineage · run state · atomic publication · errors · contracts · CLI.
**OPTIONAL (removable):** COM adapter · encrypted adapter · xlsb/xls support · CSV & JSONL sinks · TUI · doctor · benchmarks.
**REQUIRED DISTRIBUTION:** a portable Windows 10/11 PyInstaller one-dir build. Its early smoke build follows the first working `doctor`/`inspect` slice; release repeats that proof on a clean standard-user machine offline.
**EXTERNAL (contract-only, never imported):** downloaders · cleaners · validators · history stores · databases · analytics · dashboards · orchestrators · AI agents.
**FORBIDDEN:** file downloading · business-meaning inference · dashboards · scheduling · network · telemetry · sending source data to an AI · modifying sources · weakening machine security · killing Excel processes we do not own · a central application that holds the tools together.
