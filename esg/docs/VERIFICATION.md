# Release verification — ESG Banking Android 1.0.0

Executed on 27 September 2026 in the Windows build workspace.

| Check | Result |
|---|---|
| Financial, ESG, emissions and CSV regression checks | 38 passed |
| IndexedDB adapter checks | 38 passed, using fake-indexeddb |
| Web/native bridge contract checks | 22 passed, using fake window/timers |
| TypeScript | Passed |
| Production web bundle | Passed; all executable assets bundled |
| Android assembleRelease | Passed |
| Android lintRelease | Passed, 0 errors; 5 warnings (3 pinned dependency versions, 2 untranslated native button labels) |
| APK alignment/signature | Verified with Android zipalign/apksigner; RSA 3072-bit demonstration certificate, APK Signature Scheme v3 |
| Manifest | com.esgbanking.workbench; version 1.0.0 / 1; min API 28; target API 35; non-debuggable; no Internet or broad storage permission |
| Packaged web assets | Byte-for-byte equality with final compiled web assets |
| Browser workflow | Selected Physical shock, saved, reloaded and confirmed persistence; captured immutable run visible; mobile navigation to report works |
| Phone layout | 390 × 844 viewport tested; content/viewport width matched at 375 CSS px excluding scrollbar; screenshot visually reviewed; no console errors observed |

**Not performed:** physical Android installation, emulator execution, real Android file-picker/export round-trip, actual Android PDF output, process-kill/upgrade persistence, OS storage exhaustion. The browser and contract tests do not substitute for those device checks. The APK is suitable for a controlled demonstration and device acceptance testing; no production banking or regulatory certification is asserted.

The Android SDK metadata emitted a nonblocking schema-version warning because the installed SDK manager is newer than AGP. The compile platform is API 35, revision 2, and the build completed successfully. Java reports deprecated compatibility calls used for older supported Android versions.

APK bytes: 5887725

APK SHA-256: `6d42d1cc7db0da29322f090277477ea74c3d84e5b021b204f2682a5e324180e8`

Certificate SHA-256: `0cf28ef0a1efc2ca9b45cead9f03cf5f95f31da2494703d5e9f750c0c63bd6c9`

The source ZIP excludes private signing keys, account configuration, dependency/build caches and local user data. See APK_SIGNATURE.txt for the verifier output and the delivery SHA256SUMS.txt for all artifact checksums.
