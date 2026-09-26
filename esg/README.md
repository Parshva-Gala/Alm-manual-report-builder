# ESG Banking — Android offline edition

**Version 1.0.0 · 27 September 2026**

An installable Android application containing the ESG Banking workbench, its financial, emissions and ESG decision models, and a synthetic demonstration bank. The application runs offline and saves one workspace on the device. It does not synchronize with the separately supplied cloud application.

## Install the supplied APK

1. Copy `ESG-Banking-1.0.0.apk` to an Android phone or tablet and open it in Files.
2. If Android asks, allow that file manager to install this APK. Complete the installation, then turn the permission off again if desired.
3. Open **ESG Banking**. The demonstration bank is ready immediately; no sign-in is required.

Requires **Android 9 (API 28) or later** and **Android System WebView/Chrome 111 or later**. A tablet or landscape screen provides more room for the financial schedules. The APK is a non-debuggable release signed with a demonstration certificate. It is not a Play Store release or a certified production banking system.

## Demonstrate the application

Open **Scenario lab**, select a case, change severities or management actions, and compare the five-year capital paths. Choose **Capture run** to preserve that scenario. **Save workspace** commits current inputs; drafts are otherwise temporary. Explore **Capital & earnings**, **Emissions**, **Evidence & actions**, **Taxonomy**, and **Reports** using the same customer portfolio.

In **Data workspace**, export the current workspace as JSON, obtain editable CSV templates, or preview and apply an import. Android opens its Files picker for imports and exports. **Reports → Print / PDF** opens the Android print service; available destinations depend on the device.

**Before uninstalling or clearing app data, export a workspace JSON.** This is a backup of current inputs; it excludes saved scenario captures and event history. Those remain on this device. Imported files become a draft and still require Save. Ordinary import does not manufacture review approvals.

## Source package

| Folder | Contents |
|---|---|
| `web/` | React screens, shared analytical engines, reference data, IndexedDB adapter and automated checks |
| `android/` | Java host, secure bundled WebView, Android file/print integration, resources and Gradle wrapper |
| `scripts/` | Web asset bundling and Windows build automation |
| `docs/` | Architecture, contracts, technical limitations, build/install/release notes and verification record |
| `framework/` | Existing ESG framework and material explanations |
| `cloud-reference/` | Original cloud application's source, kept separate from the APK |

Build tools, dependency caches, account configuration, local workspaces and private signing keys are not included. Locked npm dependencies and pinned Android/Gradle versions are supplied. The packaged Android web assets are also included so Android Studio can build the supplied version immediately after resolving Android dependencies.

## Rebuild

Install Node.js 22.13+ with npm, JDK 17, and Android SDK platform 35 with Build Tools 35.0.0. Set `JAVA_HOME` and `ANDROID_HOME`, or configure Android Studio with this SDK. Gradle 8.11.1 is downloaded and checksum-verified by the wrapper. On Windows:

```powershell
powershell -File scripts/build.ps1
```

On macOS/Linux:

```sh
cd web
npm ci
npm run typecheck
npm test
npm run build
cd ..
node scripts/bundle-web.mjs
cd android
chmod +x gradlew
./gradlew assembleRelease lintRelease
```

The result is `android/app/build/outputs/apk/release/app-release-unsigned.apk`. Align and sign it with your own controlled certificate before installing/distributing it. A separately signed rebuild cannot update the demonstration-signed APK in place. Export the workspace first, then uninstall the demonstration APK, or use a different application ID for parallel installation. The private demonstration signing key is deliberately excluded.

## Technical documentation

- [Architecture and model map](docs/ARCHITECTURE_AND_MODELS.md)
- [Data and adapter contract](docs/DATA_AND_ADAPTER_CONTRACT.md)
- [Installation, backup and release](docs/INSTALLATION_BACKUP_AND_RELEASE.md)
- [Validation and limitations](docs/VALIDATION_AND_LIMITATIONS.md)
- [Cloud application reference](docs/CLOUD_REFERENCE.md)
- [Executed verification and release checksums](docs/VERIFICATION.md)
- [Dependency notices](docs/THIRD_PARTY_NOTICES.txt)

The country selector provides reporting context and supported rule comparisons. Prudential thresholds and bank calibrations remain explicit inputs. Synthetic data, local review identities and calculation fingerprints are for demonstration; they do not establish regulatory approval or authenticated independent sign-off.

## Downloads in this repository

- [Android APK](releases/ESG-Banking-1.0.0.apk)
- [Complete source ZIP](releases/ESG-Banking-Source-1.0.0.zip)
- [Technical handover notes](releases/ESG-Banking-Technical-Notes.md)
- [ESG framework](framework/ESG_Banking_Framework.md) and [HTML framework](framework/ESG_Banking_Framework.html)
- [Excel product](excel/ESG_Banking_Product.xlsx) and [Excel guide](excel/Product_Guide.md)
- [SHA-256 checksums](releases/SHA256SUMS.txt)
