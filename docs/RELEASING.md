# Releasing

Release promotion is deliberately a build-once process. GitHub Actions can produce an unpromoted candidate for comparison, but it does not publish a different, untested installer.

1. On Windows 11 ARM64, run `scripts\build-release.ps1 -Version X.Y.Z`. This builds, tests, publishes the self-contained app, bundles exact third-party runtime notices, compiles the installer, writes checksums, and updates the matching WinGet manifest hash.
2. Test that exact `artifacts\release\ExcelDiffTracker-Setup-arm64.exe` on Windows 11 ARM64. At minimum, run `scripts\test-installer.ps1 -RequireNoDotnet`, the real-Excel save smoke test, and the release checklist in `docs/TESTING.md`.
3. Commit the source and final WinGet hash. Do not rebuild the installer after validation.
4. From an authenticated checkout, run `scripts\publish-tested-release.ps1`. It refuses to publish if the tree is dirty or if the installer, checksum file, and WinGet manifest disagree. It creates and pushes the version tag, then uploads the exact tested files.
5. Validate and install the published WinGet manifest locally. Submit it to `microsoft/winget-pkgs` only after the release URL is live and all required validation passes.

The first unsigned release should be promoted with `-Prerelease` until onboarding, tray interaction, themes, high contrast, keyboard navigation, and scaling have been manually accepted.
