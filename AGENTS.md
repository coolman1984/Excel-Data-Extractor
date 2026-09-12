# AGENTS.md — how AI agents work in this repository

You are working on **xlextract**: a Windows-first, offline CLI tool that extracts trustworthy structured data from Excel workbooks and proves what it extracted.

## Read order (do not read the whole repo)

1. **This file** — rules of engagement.
2. **`PROJECT_EYE.md`** — the 5-level map. Find your node.
3. **Your node's row in `.project-eye/graph.yaml`** — purpose, owner, inputs, outputs, deps, tests, journeys.
4. **`schemas/`** — any contract your node touches.
5. **`PLAN.md`** — only the sections your task pack names. It is long; do not read it end to end.
6. **`CURRENT_STATE.md`** — what exists today and what is in flight.

`ARCHITECTURE.md` is the *why*. Read it before proposing any structural change.

## Your task pack

Material work arrives as an execution pack (PLAN.md §20/§21) containing: Task ID · Goal · Prerequisites · Target node · Owning domain · Journey · **Allowed files** · **Forbidden files** · Inputs/contracts · Requirements · Required tests · Acceptance evidence · Stop conditions · Handoff.

**Allowed files is a whitelist.** Do not open, edit, or "fix in passing" anything outside it. If your task has no pack, ask for one instead of improvising scope.

## The three invariants (violating one is a stop-the-line event)

- **I1 — Source is evidence.** Never write to, save, rename, move, or delete a source workbook. Sources are opened read-only, always.
- **I2 — No silent loss.** Every observed source row lands in exactly one ledger bucket, and the reconciliation identity must hold or the run does not succeed.
- **I3 — No silent decision.** Every interpretive choice is recorded in the manifest with its basis. Genuine ambiguity is a terminal `NEEDS_INPUT`, never a guess.

## Stop and escalate to Opus-level review when

- ownership is unclear
- a public contract must change
- an architecture boundary must change
- a cross-domain dependency is required
- security behavior is unclear
- a data-loss risk appears
- a large refactor becomes necessary
- runtime reality contradicts the map
- a supposedly optional capability turns out to be required by core

Do not quietly redesign the system. Stopping with a clear question is a success; a silent redesign is a defect.

## Scope control

Found an unrelated problem? **`RECORD_ONLY`** — append it to `CURRENT_STATE.md` under *Open RECORD_ONLY items* and move on. Fix it only if it blocks your task or is a critical security/data-loss risk.

## Defect protocol (no exceptions)

reproduce → capture evidence (the run directory) → smallest failing test → confirm it fails → fix → confirm it passes → run the affected journey end to end → update map/contracts if the fix touched either.

## Before you claim done

```powershell
ruff check . ; ruff format --check .
mypy src
pytest -m "not real_excel"
python -m tools.gates all
python -m tools.eye check --all
```

Then verify, for your change: **node · edges/contracts · affected subgraph · journey · architecture rules · map freshness.**

**Anti-drift rule:** a material change is incomplete until code, tests, public contracts, architecture rules, `PROJECT_EYE.md`, `.project-eye/graph.yaml`, and `CURRENT_STATE.md` all agree. A code change with a stale map is incomplete. A map update unsupported by code is also incomplete.

## Never do this

- Make a network call, add a networking dependency, or emit telemetry.
- Send source data, file names, or headers to any AI or remote service.
- Modify, save, rename, or delete a source workbook.
- Execute macros or update external links.
- Write a decrypted copy of an encrypted workbook to disk.
- Accept a password as a command-line argument, or log/serialize one anywhere.
- Kill, enumerate-for-termination, or otherwise interfere with an Excel process the tool did not create.
- Change Windows, Office, macro-security, DRM, execution-policy, or antivirus settings.
- Require administrator rights.
- Import an adapter, a sink, or `core.extraction` from `cli/` or `tui/`.
- Import `win32com`, `openpyxl`, `python_calamine`, `msoffcrypto`, or `pyarrow` outside their owning package.
- Add a `utils/`, `helpers/`, or `common/` package. Shared code lives in the domain that owns it.
- Claim enterprise DRM support without a recorded F32 run on the real target machine.
