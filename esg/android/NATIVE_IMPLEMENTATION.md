# Native Android host

The app package is `com.esgbanking.workbench`, displayed as **ESG Banking**, version 1.0.0 (code 1). The Java entry point is `app/src/main/java/com/esgbanking/workbench/MainActivity.java`.

The build pins AGP 8.9.2, compile/target SDK 35, Build Tools 35.0.0, Java 17, AndroidX Activity 1.10.1, Core 1.15.0 and WebKit 1.12.1. It requires Android 9/API 28 and an installed WebView provider version 111 or later. Runtime bridge and JavaScript feature checks show an update/error screen when requirements are unmet. Root delivery scripts provide the Gradle 8.11.1 wrapper, SDK path, web bundle and final signing. `assembleRelease` is intentionally unsigned.

## Implemented boundaries

- Only packaged `/assets/www/` resources at `https://appassets.androidplatform.net` load in the WebView. Resource requests outside this boundary receive a blocked response, including file/content/remote URLs; missing asset-handler results cannot fall through to the network. User-initiated external HTTPS links open in the system browser.
- No `INTERNET` or broad storage permission. Backup and device transfer are excluded. JavaScript/DOM storage are enabled for the bundled application; file access and universal file URL access are disabled. Content access is enabled specifically for HTML-selected files, while direct content navigation/resource requests remain blocked. Permissions requested by web content are denied. Debugging is controlled by `BuildConfig.DEBUG`.
- `ESGNative` is an origin-restricted `WebMessageListener`, registered before navigation. Messages are accepted only from the main frame of the local app origin. There is no `addJavascriptInterface` fallback.
- Requests use a string ID containing 1–128 letters, digits, `.`, `_`, `:`, or `-`. Export accepts `{id,action:'export',name,mime,text}` for exact MIME values `application/json`, `text/csv`, `text/plain`, with at most 5 MiB after UTF-8 encoding. Filenames are sanitized; the user chooses the destination with SAF. Success follows completed stream writing/closing, not merely opening the picker.
- Print accepts `{id,action:'print'}` and uses the current WebView's print adapter. The frontend must finish report rendering before sending it. It replies upon completed/cancelled/failed Android job status. At 290 seconds a nonterminal job returns an error explaining that it remains pending in the Android print queue; it is not cancelled. Other file operations are then available, while a second print remains blocked until that job ends.
- Replies are JSON `{id,ok,error?,cancelled?}` through the originating reply proxy. Unknown, duplicate and concurrent operations are rejected; native I/O uses a background executor.
- HTML file input uses `OpenDocument`, requested JSON/CSV/text filtering, a 5 MiB native byte limit and one pending callback. Selected data is copied into a narrowly exposed private cache directory through a nonexported `FileProvider`; its URI is returned to WebView. The web app retains stricter import limits and schema/preview validation. Cancellation returns `null` without changing web data.
- Back invokes synchronous `window.esgAndroidBack()`. If it returns true, the app stays open. Otherwise `window.esgHasUnsavedChanges` controls an exit confirmation. OS bar, display-cutout and keyboard insets are applied; orientation/screen-size changes preserve the Activity/WebView.

## Verification and limitations

The six source XML files parsed successfully and the source manifest declares zero permissions. A first Java compile exposed an incompatible error callback signature; it was corrected to `WebResourceErrorCompat`. Final Gradle/lint/signature results belong to the delivery owner's build record, not this source note.

Physical/emulated-device acceptance remains necessary: offline startup, WebView update behavior, bridge delivery, external references, HTML file URI reading, picker cancellation, cloud document providers, completed/cancelled/blocked printing, narrow/large screens, IME, back gesture and process recreation. Process death during a picker or print loses its in-memory callback; pending-operation recovery is not implemented. The frontend's five-minute timeout exceeds the native print deadline. Native printing does not provide custom headers/footers or guaranteed pagination. The minimum provider version is a compatibility floor, not a claim that an outdated version is secure. Imported temporary files remain in private app cache and may be removed by Android; they are not a durable source-record store.

The source ZIP must exclude SDK paths, caches and signing keys. The release APK must be signed and verified after all asset packaging is complete.
