# Handoff

## Current state

- Development is on `excel-scenario-analysis-tool`; `main` remains the simpler Diff Tracker line.
- Version 0.2.0 is an unqualified local candidate. Its frozen build source is
  `845f20283182c5b529c74370918cef3419ca7631`.
- The frozen installer SHA-256 is
  `3562301F1EA6375A1726981A263E443BB794627755F1746304396D33167A2144`; the candidate application payload
  SHA-256 is `9244736A10ED36F38F4088A5DF6A6369DCC664417AE8D4FF62F3CD04C3A182CE`.
- The source suite passed 30 tests, targeted minimum-size visual review passed, and two independent
  Sol-medium reviews accepted the local candidate within their stated scope.
- The maintainer installed and launched 0.2.0 and positively assessed the redesigned typography,
  workbook menus, and compact/expanded views.

## Known uncertainty

- The prior and new applications were running simultaneously during the first 0.2.0 launch.
- Optional legacy cleanup could not remove a locked SQLite write-ahead-log file. The cleanup failed
  closed: no workbook or Markdown report was touched.
- Installed replacement, process handoff, and legacy-cleanup behavior therefore require a targeted
  stabilization pass before 0.2.0 can replace the everyday baseline.
- Scenario recording sessions, an Excel add-in, and downstream visualization integration are planned
  directions, not implemented features.

## Working direction

The immediate task is a structured maintainer intake. Its accepted installer and usability findings
should be repaired together in one stabilization pass. Product direction beyond the implemented
comparison engine is provisional and is maintained in `docs/PRODUCT_ROADMAP.md`. The supporting
[product strategy corpus](docs/product-strategy/README.md) preserves the product thesis, tiered
scenario possibilities, read/operate/build trust boundary, market and competitor framing, GCBS
relationship, and the 2026-09-03 discovery synthesis without representing those ideas as
implemented.

## Evidence boundary

Private application data and test evidence remain local and untracked. This handoff records only
sanitized conclusions and frozen, non-sensitive build identity.
