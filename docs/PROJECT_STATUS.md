# Project status

Updated 2026-09-03 UTC.

## Current phase: everyday use of 0.1.2

Version 0.1.2 is installed on the maintainer's everyday Windows 11 ARM64 Parallels PC. The source changes through `3fb6304e6d95f0b1b83180d065f3e8be8ce5cfc7` have been pushed to `main`. The upgrade, history-preservation, tray, and real sign-out/sign-in checks passed. No 0.1.2 release tag or public installer release was created during this work.

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
