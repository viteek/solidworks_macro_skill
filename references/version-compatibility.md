# Compatibility evidence matrix

This matrix covers the 24 tasks in [the benchmark suite](../benchmarks/README.md). `NOT_TESTED` means no complete, task-bound live evidence has been accepted for that cell. A documented API member, installed type library, successful Python test, or successful macro invocation alone does not certify a SOLIDWORKS release.

| SOLIDWORKS / service pack | Static source review | Installed API signatures | Whole VBA project compile | Runtime | Geometric/document results | Safety |
| --- | --- | --- | --- | --- | --- | --- |
| 2018 / unspecified | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |
| 2020 / SP5 | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |
| 2021 / unspecified | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |
| 2022 / unspecified | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |
| 2023 / unspecified | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |
| 2024 / SP5 | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |
| 2025 / unspecified | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |
| 2026 / SP4 (34.4.0), Polish UI | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED | NOT_TESTED |

The earlier repository's SOLIDWORKS 2024 SP5 working-environment note is historical anecdotal context. It has no retained whole-project compilation, task fixtures, assertion reports or safety evidence and is not certification. Edgarev's upstream README separately states author testing on 2020 SP5; that third-party statement does not validate this repository or these tasks.

On 2026-11-01 the audit located an installed `SOLIDWORKS 3DEXPERIENCE R2026x` directory and an existing process. Installation/process discovery is capability evidence only. Any narrower installed-library audit belongs in its own report with actual library hashes and checked symbols; it must not promote all 24 tasks to API PASS. Update a matrix cell only with linked task IDs, exact version/service pack/API revision/architecture/UI language, source and fixture hashes, phase-specific artifacts and [verification reports](../benchmarks/README.md). A report for a different source/fixture or a changed library invalidates its applicability.

API introduction years are separately tracked in [the source index](source-index.md). They describe documentation availability, not tested target versions. The VBA project must select a documented API supported by the actual target library; switching to a newer replacement method requires new signature, compile and runtime evidence.

## Narrow local evidence, outside the full benchmark matrix

[The 2026-11-01 validation record](../validation/2026-11-01/README.md) retains actual installed-library
checks and CLI/COM runs on SOLIDWORKS 2026 SP4 (API 34.4.0, Windows x64, Polish UI).
These cover copied fixtures for nine modeling/selection scenarios, six assembly/drawing/inspection
recipes, a positive sheet-metal base flange and its persisted readback, source metrics, validated STEP,
COM Boolean behavior and native project invocation. Results are scoped to their recorded sources,
fixtures and assertions. They do not populate the 24-task matrix or establish compatibility with
other releases. Complete-project compile evidence remains NOT_TESTED: the saved experiment's
module inventory did not match the source collection, so its in-memory compile observation was rejected.
