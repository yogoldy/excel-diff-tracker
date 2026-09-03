# Public release checklist — Excel Diff Tracker 0.1.2

## Current decision

`v0.1.2` pins the source commit for the installer currently used in everyday work. It is **not** a claim that the candidate is fully qualified for public release, nor authorization to publish an installer asset, submit WinGet metadata, or represent the product as production-ready.

The existing installer is a real local candidate. Its documented SHA-256 is
`9536373FCC47119D2C10A75C8E6D8F1BFC2BF38FDF08639DD27790D95422133A`; its build-source commit is
`3fb6304e6d95f0b1b83180d065f3e8be8ce5cfc7`.

## Paused public-release work

Resume this checklist only when public distribution is again in scope. Freeze a clean candidate
source commit and installer before beginning; any code, packaging, or installer change requires a
fresh qualification run.

### 1. Establish reusable, lawful test workbooks

- Identify openly licensed or otherwise authorized Excel workbooks that have meaningful cross-sheet
  formulas, named ranges, formulas with cached results, substantial used ranges, and realistic save
  behavior.
- Keep fixture provenance, license, source URL, and any handling restrictions with each candidate.
- Add deliberately generated fixtures only for edge cases that open examples cannot safely cover.
- Include both `.xlsx` and macro-bearing `.xlsm` cases; ensure macro-preservation checks do not execute
  untrusted macros.
- Include a large-data fixture for the product benchmark. Do not use confidential institutional models
  as public test fixtures or commit their contents to this repository.

### 2. Run clean installed-product acceptance

- Use named, resettable Windows 11 ARM64 VM snapshots and a standard non-administrator account.
- Perform two independent clean critical-path runs using the actual installed product and visible
  Microsoft Excel—not source-level substitutes.
- Exercise onboarding, dashboard, tray, restart, sign-out/sign-in, in-place upgrade from the pinned
  prior public installer, and uninstall/data-retention behavior.

### 3. Prove Excel capture and recovery behavior

- Run the complete real-Excel `.xlsx` and `.xlsm` semantic matrix: values, formulas, cached-result
  changes, types, sheet operations, style-only saves, ordered repeated saves, and macro-stream
  preservation.
- Run watcher/recovery cases: exclusive locks, Save As/atomic replacement, AutoSave/write bursts,
  report-folder failure and recovery, missing-file restoration, encrypted/corrupt files, and app
  restart reconciliation.
- Run the large-workbook benchmark and the required multi-workbook sustained-save soak.

### 4. Complete visual and accessibility acceptance

- Capture the documented UI matrix across themes, high contrast, display scales, window sizes, tray,
  processing, warning, and error states.
- Verify keyboard navigation, screen-reader/UI-Automation names, focus visibility, contrast, clipping,
  and long-path accessibility.
- Obtain explicit independent human visual/accessibility approval.

### 5. Validate and publish only after approval

- Preserve the required checksummed acceptance-evidence bundle outside the public repository when it
  contains personal profile data.
- Obtain the functional, visual/accessibility, and release-custodian attestations for the exact same
  installer and evidence digest.
- Run aggregate validation. Resolve all P0-P2 findings.
- Address signing/SmartScreen readiness before public installer and WinGet distribution.
- Only then create or update the public GitHub release asset and publish the matching WinGet metadata.

## References

- `docs/ACCEPTANCE.md`
- `docs/VALIDATION.md`
- `docs/RELEASING.md`
- `docs/INSTALLED_LIFECYCLE_UPGRADE.md`
