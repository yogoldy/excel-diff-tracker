# Project status

Updated 2026-09-03 UTC.

## Current phase: installed 0.2.0 maintainer trial

Development on `excel-scenario-analysis-tool` reframes the desktop product around saved scenario evidence and selected-baseline comparisons. The current implementation keeps chronological save reports immutable while adding previous-save, original-baseline, and specific-scan comparison views; responsive compact/expanded workbook and history pages; path actions; new branding; native title-bar theme synchronization; and a fresh 0.2 local-data boundary. It remains a local candidate only. The 30-test source suite, targeted minimum-size visual review, and two independent Sol-medium reviews passed on 2026-09-03.

The maintainer installed and launched the 0.2 candidate. Initial feedback accepted the larger visual treatment, workbook three-dot menus, and compact/expanded views. During that launch, the prior and new applications were running simultaneously. Optional legacy cleanup could not remove a SQLite write-ahead-log file held by the older process. Cleanup failed closed, and no workbook or Markdown report was touched. Replacement, process handoff, and legacy cleanup therefore remain unresolved.

## Build and installation identity

The tested 0.2 source is frozen at `845f20283182c5b529c74370918cef3419ca7631`. Its installer SHA-256 is `3562301F1EA6375A1726981A263E443BB794627755F1746304396D33167A2144`, and its application SHA-256 is `9244736A10ED36F38F4088A5DF6A6369DCC664417AE8D4FF62F3CD04C3A182CE`. Later documentation commits do not change that binary or extend its test coverage.

Version 0.1.2 remains the prior everyday-use line and is tagged `v0.1.2`. The 0.2 maintainer trial does not establish that replacement succeeded, that 0.1.2 was removed, or that 0.2 is ready for sustained use.

## Immediate next step

Conduct the structured maintainer intake in `TASKS.md`, covering installer behavior, visual changes, workflow glitches, and missing behavior. Repair the installer/process-handoff issue and accepted intake findings together in one later stabilization pass. Do not begin Excel add-in implementation before that intake and stabilization scope are settled.

The fuller scenario-recording and product-integration direction is provisional and belongs in [Product roadmap](PRODUCT_ROADMAP.md). See [Validation](VALIDATION.md) for completed checks and their limits.

## Working and review preferences

- Use one primary agent for routine fixes, documentation, and feedback triage. Do not automatically launch multiple independent reviews for each change or repeat broad audits without a concrete reason.
- The maintainer will assess visual details and workflow nuances during normal use. Focus follow-up work on their observations and targeted verification.
- Use additional agents when explicitly requested, or when an applicable formal release requirement calls for independent review. Keep such reviews bounded to the decision they need to support.
- This preference replaces the earlier blanket request for independent agents on every meaningful decision during day-to-day work. It does not mark formal release attestations as passed or remove the existing publication gates.
- Preserve the everyday PC, user workbooks, history, and customizations. Use recoverable backups before installed changes; do not run destructive clean-install/uninstall harnesses against this profile.

## Deferred qualification

The complete installed `.xlsx`/`.xlsm` semantic and recovery matrices, large-workbook benchmark, sustained-save soak, full visual/scaling review, clean lifecycle/uninstall runs, and aggregate release approval remain incomplete for this exact candidate. Earlier-version results and normal use do not substitute for these formal checks. Revisit [Releasing](RELEASING.md) before publishing a release.

Detailed local evidence and backups remain intentionally untracked because they include personal application data. They must not be added to this public repository.
