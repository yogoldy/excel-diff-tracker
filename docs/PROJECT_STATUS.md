# Project status

Updated 2026-09-03 UTC.

## Current phase: 0.2.0 local scenario-analysis candidate

Development on `excel-scenario-analysis-tool` reframes the desktop product around saved scenario evidence and selected-baseline comparisons. The current implementation keeps chronological save reports immutable while adding previous-save, original-baseline, and specific-scan comparison views; responsive compact/expanded workbook and history pages; path actions; new branding; native title-bar theme synchronization; and a fresh 0.2 local-data boundary. It is a local candidate only. The 30-test source suite, targeted minimum-size visual review, and two independent Sol-medium reviews passed on 2026-09-03. Installed replacement and upgrade testing still must complete before it can replace the everyday installation.

The 0.2 transition may delete only the old SQLite `history.db`, `history.db-shm`, and `history.db-wal` files after explicit confirmation. Existing Markdown reports are outside that cleanup boundary and must remain untouched.

## Installed baseline: everyday use of 0.1.2

Version 0.1.2 is installed on the maintainer's everyday Windows 11 ARM64 Parallels PC. Tag `v0.1.2` anchors the installed source line for maintainer use; it is not evidence that the full public-release qualification gates passed. The upgrade, history-preservation, tray, and real sign-out/sign-in checks passed.

The next step is to use the installed app during normal Excel work and report observed problems. Everyday use has been planned; sustained workflow reliability has not yet been established. Useful feedback includes what was edited, when it was saved, what history/report appeared, and any visible error. Screenshots and a small reproducible workbook are useful when available; private workbook contents need not be published.

The installed candidate was built from `3fb6304`. Later documentation updates do not change that installed binary or extend its test coverage. See [Validation](VALIDATION.md) for exact candidate hashes, completed checks, and remaining qualification.

## Working and review preferences

- Use one primary agent for routine fixes, documentation, and feedback triage. Do not automatically launch multiple independent reviews for each change or repeat broad audits without a concrete reason.
- The maintainer will assess visual details and workflow nuances during normal use. Focus follow-up work on their observations and targeted verification.
- Use additional agents when explicitly requested, or when an applicable formal release requirement calls for independent review. Keep such reviews bounded to the decision they need to support.
- This preference replaces the earlier blanket request for independent agents on every meaningful decision during day-to-day work. It does not mark formal release attestations as passed or remove the existing publication gates.
- Preserve the everyday PC, user workbooks, history, and customizations. Use recoverable backups before installed changes; do not run destructive clean-install/uninstall harnesses against this profile.

## Deferred qualification

The complete installed `.xlsx`/`.xlsm` semantic and recovery matrices, large-workbook benchmark, sustained-save soak, full visual/scaling review, clean lifecycle/uninstall runs, and aggregate release approval remain incomplete for this exact candidate. Earlier-version results and normal use do not substitute for these formal checks. Revisit [Releasing](RELEASING.md) before publishing a release.

Detailed local upgrade evidence and verified backups are retained under `TestResults/release-preflight/`; they are intentionally untracked because they include personal application data. They must not be added to this public repository.
