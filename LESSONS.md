# LESSONS.md — hard-won facts

Append-only. Newest last. One entry per finding.

Format: **date · context · finding · consequence for the design**. Write an entry whenever something surprised you, cost you time, or would cost the next agent time. This file exists so nothing is learned twice.

---

### 2026-09-12 · Planning research · calamine materializes an entire sheet

**Context:** choosing the primary reader during architecture planning.

**Finding:** `python-calamine` (0.8.2, Jul 2026) is the fastest option and covers `.xlsx/.xlsm/.xlsb/.xls` in one dependency, but `to_python()` loads a whole sheet into memory — roughly 400 MB resident for an 80 MB workbook. True lazy row iteration is still not available; the maintainer attributes this to pyo3 binding limits. `openpyxl` in `read_only=True, values_only=True` mode handles the same file in about 40 MB, at 5–6× the wall time.

**Consequence:** no single reader is both fastest and bounded. Adapter choice must be a **measured runtime decision against a memory budget** (`core.source.probe` → `core.extraction.planner`), never a project-wide library preference. This is the reason the probe stage exists at all and the reason `PLAN.md` §12 is a decision table rather than a default.

*Sources: python-calamine issue #42; PyPI python-calamine 0.8.2.*

---

### 2026-09-12 · Planning research · password protection ≠ enterprise DRM

**Context:** scoping the "protected workbook" capability.

**Finding:** `msoffcrypto-tool` (maintained, Jan 2026 release) decrypts password-protected OOXML and legacy binary files to an in-memory stream. It does **not** address IRM / rights-management / sensitivity-label protection, which is enforced by the user's authenticated Office session and cannot lawfully or technically be handled by a file parser.

**Consequence:** two distinct adapters, two distinct error codes (`E_SRC_ENCRYPTED_NO_KEY` vs `E_SRC_IRM_PROTECTED`), and a standing honesty rule: a green password-protected fixture never licenses a claim of enterprise DRM support. Only a recorded F32 run on the real target machine does. See the *Unproven claims* section of `CURRENT_STATE.md`.

---

### 2026-09-12 · Planning · why mid-file resume is deliberately absent from V1

**Context:** evaluating resumability for large-file extraction.

**Finding:** mid-file resume needs a durable, verifiable position in a stream whose source may change underneath it, plus a partial-output format that could be mistaken for a complete one.

**Consequence:** it threatens invariants I1 and I2 to save a restart of a job budgeted at ≤5 minutes. Rejected for V1 in favour of folder-level skip-if-already-extracted, keyed on `(source sha256, contract sha256, tool major version)`. Revisit only if measured single-file runs routinely exceed 15 minutes — and record the measurement here when you do.
