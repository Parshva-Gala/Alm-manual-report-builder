# Installation, backup and release operations

The application ID is `com.esgbanking.workbench`, version name `1.0.0`, version code `1`. Use the root README for the delivered APK name, checksum and exact build commands. This document describes operation and release safeguards without asserting device tests that have not been performed.

## Install and start

The target is Android 9 / API 28 or newer with Android System WebView 111 or newer. Update the active WebView provider if the app reports an unsupported runtime. Android OS support alone does not guarantee the required browser capabilities.

Transfer the APK to the device and open it through the device's installer. If required, grant the selected installer permission to install from that source. Verify the delivered checksum before installation when the APK has passed through another transfer channel. The application can then start with its packaged demonstration data without signing in or connecting to a server.

The release is configured as nondebuggable and signed with a demonstration-only key. That private key is intentionally excluded from the source ZIP. Source recipients must use their own signing key; it cannot update the demonstration installation under the same application ID unless Android accepts the same signing identity. Use a distinct application ID for a parallel installation, or export important work before uninstalling the demonstration to install a differently signed build. Neither the signing choice nor APK delivery is a Play Store submission claim.

An update preserves app data only when Android accepts it as an update of the same application and compatible signing identity. Do not uninstall merely to resolve a signing mismatch: uninstalling can remove the local workspace and captures. Export important work first.

## Save, capture and export

Input edits are a working draft until Save succeeds. Android may terminate an application without a browser unload event. Save material changes before leaving the application; a warning on navigation is not a persistence guarantee.

Capturing a scenario stores the exact current inputs and a calculated financial summary for comparison. It does not save the workspace draft. Opening a capture restores its inputs as a draft; Save makes that draft the new workspace. Explicit evidence Review also saves the entire submitted workspace when its checks pass.

The standard **Download complete workspace** JSON contains model inputs for one workspace. It is not a backup of the IndexedDB database, captured-run collection or history. To retain a particular captured scenario independently, load it as a draft and export that workspace JSON with a distinctive name. Derived-results JSON, CSV exports and a report PDF support review but are not substitutes for the editable workspace JSON.

Use the system document picker to select an export destination. The native export operation must report completion after the file is written; cancelling leaves the workspace unchanged. Keep a copy outside app-private storage and verify it can be reopened. A user-selected cloud-backed document provider may synchronize the exported file under that provider's policy; the ESG application itself performs no synchronization.

To restore, choose a workspace JSON through Data → Import data. A compatible file is validated and loaded as an unsaved draft. Inspect its name, reporting date, controls and key outputs, then save deliberately. Model-version incompatibility requires a defined migration; it must not silently reset data. Imported review fields do not bypass the current stored-review rules.

## Data retention

Workspace data is stored locally for this application installation. The manifest disables automatic backup, and extraction rules exclude app data from cloud backup and device transfer. Clearing app storage, uninstalling, device loss or storage failure can remove it. No account recovery, cloud backup or remote administrator is provided by this edition. Local history is useful operational context, not an authenticated or tamper-proof audit log.

Review identities and initial evidence references are demonstration records. Avoid describing the app's local reviewer selection as user authentication or independent assurance. Device locks, export access and any subsequent organizational controls are outside the analytical model.

## Build and release checklist

The selected toolchain is AGP **8.9.2**, Gradle **8.11.1**, JDK **17**, compile/target SDK **35**, and Build Tools **35.0.0**. Native dependencies include AndroidX WebKit **1.12.1**, Activity **1.10.1** and Core **1.15.0**. The checked-in build files and root build record are authoritative. Web dependencies are lockfile-pinned; use the documented clean install/build procedure instead of silently upgrading model or UI dependencies.

For each release:

1. Build the web assets and copy them into the native `assets/www/` package through the documented build helper. Confirm no runtime assets depend on a CDN or cloud API.
2. Run the model, adapter and build checks recorded in the root verification report. Verify application/version identifiers, packaged files and manifest permissions.
3. Produce the APK with the documented signing mode and record its SHA-256. Preserve the signing identity securely for future updates; do not include a production private key in a public source bundle.
4. Perform the device acceptance checks in [validation limits](VALIDATION_AND_LIMITATIONS.md), including export/restore and process restart. Record device model, Android version, WebView provider/version, APK hash and results.
5. Package `android/`, `web/`, `docs/`, `scripts/`, the original `cloud-reference/` and `framework/`. Include relevant licenses and reproducible build configuration; exclude caches, installed dependencies, local credentials and unrelated workspace data.

Opening the Android print dialog is not a successful PDF test. Check the resulting document's charts, tables, units and footnotes. Pagination and available print destinations depend on the device's print implementation.
