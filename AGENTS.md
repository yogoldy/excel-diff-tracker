# Excel Scenario Analysis Tool — Agent Instructions

This repository owns the standalone local Windows application. It does not own any tracked
institutional workbook, financial model, or downstream visualization application.

## Start here

Read, in order:

1. `TASKS.md`
2. `HANDOFF.md`
3. `docs/PROJECT_STATUS.md`
4. `docs/PRODUCT_ROADMAP.md`
5. `docs/ARCHITECTURE.md`

Keep their roles distinct: `TASKS.md` contains one next action; `HANDOFF.md` is compact operating
context; Project Status records current maturity; Product Roadmap records future direction; and
Architecture describes implemented mechanics.

## Product and evidence boundaries

- Keep the product general-purpose. Model-specific semantic mappings are optional configuration,
  not assumptions embedded in the core application.
- Treat chronological captures as mechanical evidence. They do not establish user intent, business
  correctness, or semantic importance.
- Preserve the distinction between direct cell changes, formula-text changes, and calculated-result
  changes.
- Do not describe planned scenario recording, an Excel add-in, or downstream visualization
  integration as implemented.
- Do not represent a documentation commit as the source of an earlier frozen binary. Preserve the
  recorded build-source commit and artifact hashes.

## Confidentiality

Treat this repository as public. Never commit workbooks, workbook contents, raw reports, snapshots,
databases, local application data, attestations, financial values, personal information, absolute
local paths, screenshots containing private material, installers, or untracked test evidence.
Describe private evidence only through sanitized conclusions. If classification is uncertain, keep
the material local.

## Lean reconciliation

- Close one checkpoint for a substantial work session, not one per file edit.
- Update `HANDOFF.md`, leave exactly one action in `TASKS.md`, update Project Status only when current
  maturity changes, and append one compact entry to `.agents/SYNC_LOG.md`.
- Preserve prior sync-log entries; record corrections by appending a new entry.
- Scan for stale claims, dead links, duplicate authority, and planned features described as current.
- Run proportional checks, `git diff --check`, a complete diff review, and a confidentiality scan.
- Commit only scoped files on the authorized branch and report the resulting commit and final status.

## Release discipline

Public release approval is governed by `docs/ACCEPTANCE.md` and `docs/RELEASING.md`. A tag, local
installation, source test pass, or maintainer trial does not substitute for those gates.
