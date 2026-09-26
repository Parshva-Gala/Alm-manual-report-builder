# Validation evidence and limits

Source review, automated calculation tests, browser tests, APK assembly and physical-device tests establish different things. The root release verification record identifies which were actually completed for the delivered APK. This document does **not** claim physical-device, emulator, process-death, native SAF or Android print testing.

## What the checks establish

| Check | Evidence it can provide | What it does not establish |
|---|---|---|
| Model parity and mutation tests | Reviewed synthetic outputs, recalculation and selected rejection/coverage rules. | Real-world calibration, regulatory acceptance or exhaustive method support. |
| Workspace/import tests | Schema guards, missing-versus-zero behavior, CSV parsing and shared-input propagation. | Correctness of customer-supplied evidence. |
| Local adapter tests | Save conflicts, atomic transactions, review-stamp handling and immutable captures in the tested IndexedDB environment. | Android lifecycle reliability on an untested device. |
| Type check and production build | Source compatibility and asset generation under the selected toolchain. | Successful installation or complete WebView rendering. |
| APK inspection/signature verification | Package identity, manifest/assets and signature/checksum properties. | Production signing-key custody or operational security certification. |
| Native device acceptance | Behavior on the recorded OS/WebView/device and tested flows. | Universal compatibility with all devices/providers. |

The reference-engine work recorded 11,952 financial comparisons, 5,404 emissions assertions across 34 mutation cases, and 63 ESG parity/mutation checks. These are prior model evidence, not additional Android device tests. The packaged `web/tests/verify-models.mjs` exercises integration and selected mutations.

The local adapter's `web/tests/verify-offline.mjs` passed 38 checks using `fake-indexeddb`, including concurrent/stale writes, immutable captures, explicit review gates, storage-write rollback injection, retention beyond the listing limits and module reload. This is an in-process IndexedDB test environment: injected quota failures and module reload are not an actual Android full-disk or process-kill test. Use the root release logs for final source/build check outcomes.

## Required native acceptance

- Install, start in airplane mode, navigate all nine views and verify dialogs/charts at phone and tablet sizes.
- Save an input change; restart, rotate and force-stop/reopen. Confirm saved state and revision survive without treating unsaved edits as saved.
- Trigger two writes from the same revision: one must win, the other must conflict without a partial history entry. Confirm failed storage or a rejected review leaves the old saved data intact.
- Capture an unsaved scenario and retrieve it after further edits/restart. Confirm capture does not advance the workspace revision or change the captured input/summary.
- Attempt ordinary Save/Capture with forged imported review fields; confirm that only trusted stored stamps are carried and changed evidence still requires review.
- Export JSON and CSV, cancel each picker, restore the exported JSON and compare inputs. Check 30 or more captures without silent deletion and test storage failure handling.
- Complete native print/Save as PDF and inspect charts, long tables, dates and boundary notes. Test Back from a dialog, from another view and from the app root.
- Test a compatible signed upgrade with saved data. Confirm older/incompatible model data is reported for migration rather than replaced by defaults.
- Verify that unknown schemes, off-origin web navigation and non-main-frame bridge calls cannot invoke privileged operations. External references must not load inside the privileged analytical WebView.

## Analytical boundaries

The default portfolio, factor registry, scenario calibrations and reviewer records are synthetic. An apparent positive headroom, green classification or complete inventory is conditional on those inputs and the implemented scope. It is not a bank's regulatory result.

The financial model uses explicit cash-flow transmission, simplified credit-risk sensitivities and a five-year run-off book. Reporting ECL, stressed provisions, realized default timing and prudential RWA are separate modeled concepts. Liquidity ratios are planning proxies, not full regulatory LCR/NSFR templates. Additional products, netting, collateral rules, legal entities and regulatory adjustments require specified methods and validation.

The emissions engine implements selected lending routes and operational GHG accounting gates. It does not implement all PCAF asset classes/options or every GHG estimation route. Missing/unsupported data reduces coverage; it must not become zero emissions. Externally sourced factors need documented geography, unit, boundary, version, validity and quality. Operational Scope 2 accounting alternatives are separate totals, and financed emissions are not priced-emissions inputs for financial stress.

Taxonomy eligibility, safeguards and approved financing allocation are separate. Internal demonstration criteria are not certification; applying a Jordan reference outside Jordan does not create local legal applicability. Readiness organizes requirements and evidence for review rather than attesting compliance. The bundled framework explains source authority, dependencies and version qualifications.

## Operational boundaries

This edition has one local workspace per app installation, no cloud synchronization, no authenticated maker/checker role system and no protected enterprise evidence repository. A review fingerprint is a change detector, not cryptographic assurance of identity or truth. Device time supplies history timestamps; the selected reporting date governs analytical validity.

Application validation reduces accidental malformed input; local software and data remain under the device holder's control. The app does not promise resistance to rooted-device modification, forensic extraction, arbitrary app-data manipulation or recovery after uninstall. No independent penetration test, accessibility certification or regulatory compliance assessment is claimed.
