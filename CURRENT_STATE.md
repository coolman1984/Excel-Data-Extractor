# CURRENT_STATE.md — what actually exists right now

**Last updated:** 2026-09-12
**Phase:** P0 started. Architecture approved after corrections; T-001 is blocked in its standard-user toolchain-install check.

Every merged task must update this file. A task that does not update it is not done.

---

## Implemented components

No domain component is implemented. T-001 has created the package, source-entry bootstrap, project metadata, and toolchain scaffold only.

---

## Phase status

| Phase | Capability | Status | Gates |
|---|---|---|---|
| P0 Foundation | `xlx doctor` + CI + Eye | **not started** | G1, G11 |
| P1 See | `xlx inspect` on all 4 formats | not started | G2, G3 |
| P2 Extract (explicit) | run dir + manifest + ledger + `verify` | not started | G4, G6 |
| P3 Find | discovery + `contract infer` | not started | — |
| P4 Scale | F20 in ≤5 min, flat memory | not started | G5, G8 |
| P5 Formats | xlsb/xls/xlsm/encrypted | not started | G10 |
| P6 Protected | COM path on the real machine | not started | **G7** |
| P7 Friendly | TUI | not started | — |
| P8 Harden | fault injection, docs, build | not started | G9, G12 |

---

## Task board

| Task | Owner model | Wave | Status |
|---|---|---|---|
| T-001 repo skeleton & toolchain | Terra | W0 | **blocked — standard-user offline `pip` cannot create its build-tracker directory in this environment** |
| T-003 Eye tooling (`tools/eye.py`, `tools/gates.py`, `.importlinter`) | Terra | W1 | ready · **scope reduced**: the five documents and `.project-eye/*` are already seeded (see below); only the tooling remains |
| S-001 public contracts v1 | Sonnet | W1 | blocked by T-001 |
| T-007 fixture generators | Terra | W1 | blocked by T-001 |
| T-014 early portable packaging smoke | Terra | W7 | blocked by T-004 and T-011 |
| all others | — | W2+ | blocked |

Full packs: `PLAN.md` §20 (Sonnet) and §21 (Terra). Merge order and parallel lanes: `PLAN.md` §22.

---

## Seeded by the planning phase (2026-09-12)

`PLAN.md` · `AGENTS.md` · `ARCHITECTURE.md` · `PROJECT_EYE.md` · `CURRENT_STATE.md` · `LESSONS.md` · `.project-eye/graph.yaml` · `.project-eye/rules.yaml` · `.project-eye/manifest.json`

`graph.yaml` currently declares the **intended** system. No node maps to real code yet, so `eye check` will report every node as unimplemented until P1 — that is expected, not drift.

---

## Known gaps

- No code, no tests, no CI.
- `.project-eye/manifest.json` is a placeholder; it becomes generated output once `tools/eye.py` exists (T-003).
- **ENVIRONMENT_BLOCKER: local development dependencies unavailable.** No dependency has been installed or verified on the target machine yet. `pip cache list` reports no locally built wheels. A writable standard-user venv can be created under `%LOCALAPPDATA%`, but its offline editable install fails with `BackendUnavailable: Cannot import 'setuptools.build_meta'`; the required build backend is absent and no local wheel source exists. Workspace-local pip attempts also fail earlier with `Permission denied` when pip creates or accesses its build-tracker directory. This is not a product or architecture blocker.

---

## Unproven claims (must stay here until evidence exists)

- **Enterprise DRM/IRM support is UNPROVEN.** The COM path is designed but not built, and cannot be validated by CI or by a synthetic fixture. It is proven only by a manual F32 run on the real target machine, recorded in `LESSONS.md` with date and Excel version. Do not claim it before then.
- **Performance budgets (`PLAN.md` §16) are estimates**, not measurements. The reference machine is unpinned (risk B2). Every threshold is provisional until T-009 publishes a baseline.
- **The 0.6 × budget calamine/streaming threshold is a placeholder** awaiting the S-007 benchmark that justifies it.

---

## Open RECORD_ONLY items

*(none yet — append findings here rather than expanding scope)*

---

## Decisions awaiting the human

1. **A real protected workbook for F32** (risk B1) — needed before release, not before P6 development. Enterprise DRM/IRM support remains explicitly unproven until that manual target-machine run is recorded.
