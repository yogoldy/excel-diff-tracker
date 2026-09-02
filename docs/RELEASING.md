# Releasing

Release promotion is deliberately a build-once process. GitHub Actions can produce an unpromoted candidate for comparison, but it does not publish a different, untested installer.

1. On Windows 11 ARM64, run `scripts\build-release.ps1 -Version X.Y.Z`. This builds, tests, publishes the self-contained app, bundles exact third-party runtime notices, compiles the installer, writes checksums, and updates the matching WinGet manifest hash.
2. Commit the candidate source before acceptance. Record the 40-character commit, installer SHA-256, and the name/ID of the clean VM snapshot used for each run.
3. Test that exact `artifacts\release\ExcelDiffTracker-Setup-arm64.exe` on Windows 11 ARM64 using `scripts\acceptance\Invoke-BlackBoxAcceptance.ps1`. Pass the frozen commit/hash and snapshot identity with `-ExpectedSourceCommit`, `-ExpectedInstallerSha256`, `-VmSnapshotName`, `-VmSnapshotId`, and `-ConfirmCleanSnapshot`. Each run automatically includes the exclusive-lock recovery gate, installed 500,000-cell benchmark, and ten-minute/20-save real-Excel soak. Supporting smokes do not replace the two installed-product runs, visual matrix, and independent attestations.
4. Run `scripts\acceptance\Test-AcceptanceEvidence.ps1` against the complete evidence directory, exact installer, and exact published acceptance probe. It writes `approval.json` only when hashes, both runs, mandatory product gates, the installed semantic matrix, required visual configurations, and all attestations agree.
5. Make no repository changes after acceptance. The candidate commit from step 2 must already contain the final WinGet hash; if any source, packaging, metadata, or installer byte changes, discard the evidence and restart acceptance.
6. From an authenticated checkout, run `scripts\publish-tested-release.ps1 -AcceptanceDirectory <evidence-directory>`. It refuses to publish if the tree is dirty, the installer/checksums/WinGet manifest disagree, or `approval.json` does not approve the exact installer. It creates and pushes the version tag, then uploads the exact tested files.
7. Validate and install the published WinGet manifest locally. Submit it to `microsoft/winget-pkgs` only after the release URL is live and all required validation passes.

The first unsigned release should be promoted with `-Prerelease` until onboarding, tray interaction, themes, high contrast, keyboard navigation, and scaling have been manually accepted.
