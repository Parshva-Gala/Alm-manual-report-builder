package com.esgbanking.workbench;

import android.annotation.SuppressLint;
import android.app.AlertDialog;
import android.content.ActivityNotFoundException;
import android.content.Intent;
import android.content.pm.PackageInfo;
import android.database.Cursor;
import android.graphics.Color;
import android.graphics.Insets;
import android.net.Uri;
import android.net.http.SslError;
import android.os.Build;
import android.os.Bundle;
import android.os.Handler;
import android.os.Looper;
import android.os.SystemClock;
import android.print.PrintJob;
import android.print.PrintManager;
import android.provider.OpenableColumns;
import android.view.Gravity;
import android.view.View;
import android.view.WindowInsets;
import android.webkit.ConsoleMessage;
import android.webkit.PermissionRequest;
import android.webkit.RenderProcessGoneDetail;
import android.webkit.SslErrorHandler;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.widget.Button;
import android.widget.FrameLayout;
import android.widget.LinearLayout;
import android.widget.ProgressBar;
import android.widget.TextView;
import android.widget.Toast;
import androidx.activity.ComponentActivity;
import androidx.activity.OnBackPressedCallback;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.core.content.FileProvider;
import androidx.webkit.JavaScriptReplyProxy;
import androidx.webkit.WebMessageCompat;
import androidx.webkit.WebResourceErrorCompat;
import androidx.webkit.WebViewAssetLoader;
import androidx.webkit.WebViewClientCompat;
import androidx.webkit.WebViewCompat;
import androidx.webkit.WebViewFeature;
import org.json.JSONObject;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/** Offline, bundled-content host. Native capabilities are limited to files and printing. */
public final class MainActivity extends ComponentActivity {
    private static final String HOST = "appassets.androidplatform.net";
    private static final String ORIGIN = "https://" + HOST;
    private static final String START = ORIGIN + "/assets/www/index.html";
    private static final int MIN_WEBVIEW_MAJOR = 111;
    private static final int MAX_FILE_BYTES = 5 * 1024 * 1024;
    private static final int MAX_MESSAGE_CHARS = MAX_FILE_BYTES * 2 + 8192;
    private final Handler main = new Handler(Looper.getMainLooper());
    private final ExecutorService io = Executors.newSingleThreadExecutor();
    private final Set<String> requestIds = new HashSet<>();
    private WebView web;
    private LinearLayout root;
    private FrameLayout content;
    private ProgressBar progress;
    private LinearLayout errorPanel;
    private TextView errorText;
    private boolean ready, destroyed, bridgeBusy, backPending, pageFailed;
    private Set<String> importExtensions = Collections.emptySet();
    private ValueCallback<Uri[]> fileCallback;
    private ExportRequest pendingExport;
    private PrintJob printJob;

    private static final class ExportRequest {
        final String id, mime, name;
        final byte[] bytes;
        final JavaScriptReplyProxy reply;
        ExportRequest(String id, String mime, String name, byte[] bytes, JavaScriptReplyProxy reply) {
            this.id=id; this.mime=mime; this.name=name; this.bytes=bytes; this.reply=reply;
        }
    }

    private final ActivityResultLauncher<Intent> exportPicker = registerForActivityResult(
        new ActivityResultContracts.StartActivityForResult(), result -> {
            ExportRequest request = pendingExport;
            pendingExport = null;
            if (request == null) return;
            Uri uri = result.getData() == null ? null : result.getData().getData();
            if (result.getResultCode() != RESULT_OK || uri == null) {
                bridgeBusy=false; respond(request.reply,request.id,false,null,true); return;
            }
            if (!"content".equals(uri.getScheme())) {
                bridgeBusy=false; respond(request.reply,request.id,false,"The selected document location is unsupported.",false); return;
            }
            io.execute(() -> {
                String failure=null;
                try (OutputStream out=getContentResolver().openOutputStream(uri,"wt")) {
                    if (out==null) {
                        throw new java.io.IOException("No output stream");
                    }
                    out.write(request.bytes); out.flush();
                } catch (Exception ignored) { failure="The document could not be saved. Please choose another location."; }
                final String error=failure;
                main.post(() -> { bridgeBusy=false; respond(request.reply,request.id,error==null,error,false); });
            });
        });

    private final ActivityResultLauncher<String[]> importPicker = registerForActivityResult(
        new ActivityResultContracts.OpenDocument(), uri -> {
            if (fileCallback==null) return;
            if (uri==null) { finishFileChooser(null); return; }
            if (!"content".equals(uri.getScheme()) || (getPackageName()+".imports").equals(uri.getAuthority())) {
                toast("Choose a CSV, JSON or text document from Files."); finishFileChooser(null); return;
            }
            progress.setVisibility(View.VISIBLE);
            final Set<String> allowedExtensions=new HashSet<>(importExtensions);
            io.execute(() -> {
                Uri safeUri=null; String error=null;
                try {
                    String name="document";
                    try (Cursor c=getContentResolver().query(uri,new String[]{OpenableColumns.DISPLAY_NAME,OpenableColumns.SIZE},null,null,null)) {
                        if (c!=null && c.moveToFirst()) {
                            int n=c.getColumnIndex(OpenableColumns.DISPLAY_NAME), size=c.getColumnIndex(OpenableColumns.SIZE);
                            if (n>=0 && !c.isNull(n)) {
                                name=c.getString(n);
                            }
                            if (size>=0 && !c.isNull(size) && c.getLong(size)>MAX_FILE_BYTES) {
                                throw new java.io.IOException("large");
                            }
                        }
                    }
                    String ext=extensionForImport(name,getContentResolver().getType(uri));
                    if (ext==null || !allowedExtensions.contains(ext)) {
                        throw new java.io.IOException("type");
                    }
                    byte[] data;
                    try (InputStream in=getContentResolver().openInputStream(uri)) {
                        if (in==null) {
                            throw new java.io.IOException("read");
                        }
                        data=readBounded(in);
                    }
                    File dir=new File(getCacheDir(),"validated-imports");
                    if (!dir.isDirectory() && !dir.mkdirs()) {
                        throw new java.io.IOException("cache");
                    }
                    File copy=File.createTempFile("document-",ext,dir);
                    try (OutputStream out=new FileOutputStream(copy)) { out.write(data); }
                    safeUri=FileProvider.getUriForFile(this,getPackageName()+".imports",copy);
                } catch (Exception ignored) { error="Unable to open this document. Use a CSV, JSON or text file up to 5 MiB."; }
                final Uri selected=safeUri; final String failure=error;
                main.post(() -> { if (destroyed) return; progress.setVisibility(View.GONE); if(failure!=null)toast(failure); finishFileChooser(selected==null?null:new Uri[]{selected}); });
            });
        });

    @Override public void onCreate(Bundle state) {
        super.onCreate(state);
        createShell();
        getOnBackPressedDispatcher().addCallback(this,new OnBackPressedCallback(true) {
            @Override public void handleOnBackPressed() { handleAppBack(); }
        });
        try {
            PackageInfo provider=WebView.getCurrentWebViewPackage();
            int major=provider==null||provider.versionName==null?0:Integer.parseInt(provider.versionName.split("\\.")[0]);
            if (major<MIN_WEBVIEW_MAJOR) { showError("Update Android System WebView or Chrome to version 111 or later, then reopen ESG Banking.",true); return; }
            if (!WebViewFeature.isFeatureSupported(WebViewFeature.WEB_MESSAGE_LISTENER)) { showError("This WebView does not support the secure document bridge. Update Android System WebView or Chrome.",true); return; }
            createWebView();
        } catch (Exception ignored) { showError("Android System WebView is unavailable. Enable or update it, then reopen ESG Banking.",true); }
    }

    @SuppressLint("SetJavaScriptEnabled")
    private void createWebView() {
        WebView.setWebContentsDebuggingEnabled(BuildConfig.DEBUG);
        web=new WebView(this);
        content.addView(web,0,new FrameLayout.LayoutParams(-1,-1));
        WebSettings s=web.getSettings();
        s.setJavaScriptEnabled(true); s.setDomStorageEnabled(true);
        s.setAllowFileAccess(false); s.setAllowFileAccessFromFileURLs(false); s.setAllowUniversalAccessFromFileURLs(false);
        // Required for user-selected HTML file input; direct content: navigation is blocked below.
        s.setAllowContentAccess(true);
        s.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
        s.setJavaScriptCanOpenWindowsAutomatically(false); s.setSupportMultipleWindows(false);
        s.setMediaPlaybackRequiresUserGesture(true); s.setSafeBrowsingEnabled(true);
        WebViewAssetLoader loader=new WebViewAssetLoader.Builder()
            .addPathHandler("/assets/",new WebViewAssetLoader.AssetsPathHandler(this)).build();
        web.setWebViewClient(new WebViewClientCompat() {
            @Override public WebResourceResponse shouldInterceptRequest(WebView view,WebResourceRequest request) {
                if (!isLocal(request.getUrl()) || !"GET".equals(request.getMethod())) return blocked(403,"Blocked");
                WebResourceResponse response=loader.shouldInterceptRequest(request.getUrl());
                return response==null?blocked(404,"Not found"):response;
            }
            @Override public boolean shouldOverrideUrlLoading(WebView view,WebResourceRequest request) {
                if (isLocal(request.getUrl())) return false;
                if (request.isForMainFrame() && request.hasGesture()) openExternalHttps(request.getUrl());
                return true;
            }
            @Override public void onPageStarted(WebView view,String url,android.graphics.Bitmap icon) {
                ready=false; pageFailed=false; progress.setVisibility(View.VISIBLE);
                if (!isLocal(Uri.parse(url))) { view.stopLoading(); showError("A page outside the bundled application was blocked.",false); }
            }
            @Override public void onPageFinished(WebView view,String url) {
                if (!isLocal(Uri.parse(url)) || pageFailed || destroyed) return;
                view.evaluateJavascript("(()=>typeof structuredClone==='function'&&typeof BigInt==='function'&&typeof ''.replaceAll==='function')()",result -> {
                    if (destroyed || pageFailed) return;
                    if (!"true".equals(result)) { showError("This WebView lacks a required browser feature. Update Android System WebView or Chrome.",true); return; }
                    ready=true; progress.setVisibility(View.GONE);
                });
            }
            @Override public void onReceivedError(WebView view,WebResourceRequest request,WebResourceErrorCompat error) {
                if(request.isForMainFrame())showError("The bundled application could not load. Close and reopen the app; reinstall if the problem continues.",false);
            }
            @Override public void onReceivedHttpError(WebView view,WebResourceRequest request,WebResourceResponse response) {
                if (request.isForMainFrame()) showError("A required application page is missing from this build.",false);
            }
            @Override public void onReceivedSslError(WebView view,SslErrorHandler handler,SslError error) { handler.cancel(); }
            @Override public boolean onRenderProcessGone(WebView view,RenderProcessGoneDetail detail) {
                finishFileChooser(null);
                content.removeView(view); view.destroy(); web=null; ready=false;
                showError("The application renderer stopped. Close and reopen ESG Banking. Saved local work remains on this device.",false); return true;
            }
        });
        web.setWebChromeClient(new WebChromeClient() {
            @Override public void onProgressChanged(WebView view,int value) { progress.setProgress(value); }
            @Override public void onPermissionRequest(PermissionRequest request) { request.deny(); }
            @Override public boolean onConsoleMessage(ConsoleMessage message) { return !BuildConfig.DEBUG; }
            @Override public boolean onShowFileChooser(WebView view,ValueCallback<Uri[]> callback,FileChooserParams params) {
                if (!ready || !isLocal(Uri.parse(view.getUrl()==null?"":view.getUrl())) || bridgeBusy || fileCallback!=null || params.getMode()!=FileChooserParams.MODE_OPEN) { callback.onReceiveValue(null); return true; }
                importExtensions=allowedImportExtensions(params.getAcceptTypes());
                if(importExtensions.isEmpty()){callback.onReceiveValue(null);toast("Only CSV, JSON and plain-text imports are supported.");return true;}
                fileCallback=callback;
                Set<String> mimeTypes=new HashSet<>();mimeTypes.add("text/plain");mimeTypes.add("application/octet-stream");
                if(importExtensions.contains(".json"))mimeTypes.add("application/json");
                if(importExtensions.contains(".csv")){mimeTypes.add("text/csv");mimeTypes.add("text/comma-separated-values");mimeTypes.add("application/vnd.ms-excel");}
                try { importPicker.launch(mimeTypes.toArray(new String[0])); }
                catch (ActivityNotFoundException ignored) { finishFileChooser(null); toast("No document picker is available on this device."); }
                return true;
            }
        });
        if (WebViewFeature.isFeatureSupported(WebViewFeature.WEB_MESSAGE_LISTENER)) {
            WebViewCompat.addWebMessageListener(web,"ESGNative",Collections.singleton(ORIGIN),(view,message,origin,isMainFrame,reply) -> {
                if (!isMainFrame || !isOrigin(origin) || !isLocal(Uri.parse(view.getUrl()==null?"":view.getUrl())) || message.getType()!=WebMessageCompat.TYPE_STRING) return;
                String raw=message.getData();
                if (raw==null || raw.length()>MAX_MESSAGE_CHARS) { respond(reply,"",false,"The native request is too large.",false); return; }
                io.execute(() -> processMessage(raw,reply));
            });
        } else {
            showError("Update WebView to enable secure file handling.",true);
            return;
        }
        web.setDownloadListener((url,userAgent,disposition,mime,length) -> toast("Use the app's Export action to save this document."));
        web.loadUrl(START);
    }

    private void processMessage(String raw,JavaScriptReplyProxy reply) {
        String id="";
        try {
            JSONObject json=new JSONObject(raw);
            Object identifier=json.opt("id");
            if (!(identifier instanceof String) || !((String)identifier).matches("[A-Za-z0-9._:-]{1,128}")) throw new IllegalArgumentException("Invalid request ID.");
            id=(String)identifier;
            String action=json.optString("action","");
            final String requestId=id;
            if ("export".equals(action)) {
                if (!(json.opt("text") instanceof String)) throw new IllegalArgumentException("Export content must be text.");
                String mime=json.optString("mime","").toLowerCase(Locale.ROOT);
                if (!mime.equals("application/json")&&!mime.equals("text/csv")&&!mime.equals("text/plain")) throw new IllegalArgumentException("Unsupported export format.");
                byte[] bytes=json.getString("text").getBytes(StandardCharsets.UTF_8);
                if(bytes.length>MAX_FILE_BYTES)throw new IllegalArgumentException("The export exceeds 5 MiB.");
                String name=safeName(json.optString("name","esg-export"),mime);
                ExportRequest request=new ExportRequest(id,mime,name,bytes,reply);
                main.post(() -> startExport(request));
            } else if ("print".equals(action)) main.post(() -> startPrint(requestId,reply));
            else throw new IllegalArgumentException("Unknown native action.");
        } catch (Exception error) {
            final String requestId=id;
            final String explanation=error instanceof IllegalArgumentException?error.getMessage():"Invalid native request.";
            main.post(() -> respond(reply,requestId,false,explanation,false));
        }
    }

    private boolean acceptRequest(String id,JavaScriptReplyProxy reply) {
        if(destroyed||!ready||web==null){respond(reply,id,false,"The application is not ready.",false);return false;}
        if(bridgeBusy||fileCallback!=null){respond(reply,id,false,"Finish the current document operation first.",false);return false;}
        if(requestIds.contains(id)){respond(reply,id,false,"This request has already been handled.",false);return false;}
        if(requestIds.size()>1000)requestIds.clear();
        requestIds.add(id); bridgeBusy=true; return true;
    }

    private void startExport(ExportRequest request) {
        if(!acceptRequest(request.id,request.reply))return;
        pendingExport=request;
        Intent intent=new Intent(Intent.ACTION_CREATE_DOCUMENT).addCategory(Intent.CATEGORY_OPENABLE)
            .setType(request.mime).putExtra(Intent.EXTRA_TITLE,request.name);
        try { exportPicker.launch(intent); }
        catch(ActivityNotFoundException ignored){pendingExport=null;bridgeBusy=false;respond(request.reply,request.id,false,"No document picker is available.",false);}
    }

    private void startPrint(String id,JavaScriptReplyProxy reply) {
        // An active request owns its poller/reply; do not clear its completed job early.
        if(bridgeBusy){respond(reply,id,false,"Finish the current document operation first.",false);return;}
        if(printJob!=null){
            if(!printJob.isCompleted()&&!printJob.isCancelled()&&!printJob.isFailed()){
                respond(reply,id,false,"Print job is still pending; check Android print queue.",false);return;
            }
            printJob=null;
        }
        if(!acceptRequest(id,reply))return;
        try {
            PrintManager manager=(PrintManager)getSystemService(PRINT_SERVICE);
            if(manager==null)throw new IllegalStateException();
            printJob=manager.print("ESG Banking report",web.createPrintDocumentAdapter("ESG Banking report"),null);
            if(printJob==null)throw new IllegalStateException();
            pollPrint(id,reply,SystemClock.elapsedRealtime()+290000L);
        } catch(Exception ignored){printJob=null;bridgeBusy=false;respond(reply,id,false,"Printing is unavailable on this device.",false);}
    }

    private void pollPrint(String id,JavaScriptReplyProxy reply,long deadline) {
        if(destroyed||printJob==null)return;
        if(printJob.isCompleted()){printJob=null;bridgeBusy=false;respond(reply,id,true,null,false);}
        else if(printJob.isCancelled()){printJob=null;bridgeBusy=false;respond(reply,id,false,null,true);}
        else if(printJob.isFailed()){printJob=null;bridgeBusy=false;respond(reply,id,false,"The print job failed. Try Save as PDF or another printer.",false);}
        else if(SystemClock.elapsedRealtime()>=deadline){bridgeBusy=false;respond(reply,id,false,"Print job is still pending; check Android print queue.",false);}
        else main.postDelayed(() -> pollPrint(id,reply,deadline),750);
    }

    private void respond(JavaScriptReplyProxy reply,String id,boolean ok,String error,boolean cancelled) {
        if(destroyed)return;
        if (WebViewFeature.isFeatureSupported(WebViewFeature.WEB_MESSAGE_LISTENER)) {
            try {
                JSONObject response=new JSONObject().put("id",id).put("ok",ok);
                if(error!=null)response.put("error",error);
                if(cancelled)response.put("cancelled",true);
                reply.postMessage(response.toString());
            } catch(Exception ignored) { /* The requesting frame may have been destroyed. */ }
        }
    }

    private void finishFileChooser(Uri[] uris) {
        ValueCallback<Uri[]> callback=fileCallback;fileCallback=null;
        if(callback!=null)try{callback.onReceiveValue(uris);}catch(Exception ignored){ /* Renderer may already be gone. */ }
    }

    private static Set<String> allowedImportExtensions(String[] requested) {
        Set<String> result=new HashSet<>();
        if(requested==null||requested.length==0){Collections.addAll(result,".csv",".json",".txt");return result;}
        for(String entry:requested)for(String item:entry.toLowerCase(Locale.ROOT).split(",")){
            String type=item.trim();
            if(type.isEmpty()||type.equals("*/*")||type.equals("text/*"))Collections.addAll(result,".csv",".json",".txt");
            if(type.equals(".json")||type.equals("application/json"))result.add(".json");
            if(type.equals(".csv")||type.equals("text/csv")||type.equals("text/comma-separated-values")||type.equals("application/vnd.ms-excel"))result.add(".csv");
            if(type.equals(".txt")||type.equals("text/plain"))result.add(".txt");
        }
        return result;
    }

    private static byte[] readBounded(InputStream in) throws java.io.IOException {
        ByteArrayOutputStream out=new ByteArrayOutputStream();byte[] buffer=new byte[8192];int count,total=0;
        while((count=in.read(buffer))!=-1){total+=count;if(total>MAX_FILE_BYTES)throw new java.io.IOException("large");out.write(buffer,0,count);}return out.toByteArray();
    }

    private static String extensionForImport(String name,String mime) {
        String n=name==null?"":name.toLowerCase(Locale.ROOT);
        if(n.endsWith(".csv"))return ".csv";if(n.endsWith(".json"))return ".json";if(n.endsWith(".txt"))return ".txt";
        if("application/json".equals(mime))return ".json";if("text/csv".equals(mime))return ".csv";if("text/plain".equals(mime))return ".txt";return null;
    }

    private static String safeName(String supplied,String mime) {
        String name=supplied.replaceAll("[\\\\/:*?\"<>|\\p{Cntrl}]","_").trim();
        if(name.length()>100)name=name.substring(0,100);if(name.isEmpty()||name.equals(".")||name.equals(".."))name="esg-export";
        String ext=mime.equals("application/json")?".json":mime.equals("text/csv")?".csv":".txt";
        return name.toLowerCase(Locale.ROOT).endsWith(ext)?name:name+ext;
    }

    private static boolean isOrigin(Uri uri) {
        return uri!=null&&"https".equalsIgnoreCase(uri.getScheme())&&HOST.equalsIgnoreCase(uri.getHost())&&(uri.getPort()==-1||uri.getPort()==443)&&uri.getUserInfo()==null;
    }
    private static boolean isLocal(Uri uri) {
        if(!isOrigin(uri)||uri.getPath()==null||!uri.getPath().startsWith("/assets/www/")||uri.getPath().contains("\\"))return false;
        for(String segment:uri.getPathSegments())if(segment.equals("..")||segment.equals("."))return false;
        return true;
    }
    private static WebResourceResponse blocked(int status,String description) {
        Map<String,String> headers=new HashMap<>();headers.put("Cache-Control","no-store");headers.put("X-Content-Type-Options","nosniff");
        return new WebResourceResponse("text/plain","UTF-8",status,description,headers,new ByteArrayInputStream(description.getBytes(StandardCharsets.UTF_8)));
    }

    private void openExternalHttps(Uri uri) {
        if(uri==null||!"https".equalsIgnoreCase(uri.getScheme())||uri.getHost()==null||uri.getUserInfo()!=null)return;
        try { startActivity(new Intent(Intent.ACTION_VIEW,uri).addCategory(Intent.CATEGORY_BROWSABLE)); }
        catch(ActivityNotFoundException ignored){toast("No browser is available to open this reference.");}
    }

    private void handleAppBack() {
        if(backPending)return;
        if(web==null||!ready){finish();return;}
        backPending=true;
        web.evaluateJavascript("(()=>{try{return typeof window.esgAndroidBack==='function'&&window.esgAndroidBack()===true}catch(e){return false}})()",handled -> {
            if(destroyed||web==null){backPending=false;return;}
            if("true".equals(handled)){backPending=false;return;}
            web.evaluateJavascript("Boolean(window.esgHasUnsavedChanges)",dirty -> {
                if(destroyed){backPending=false;return;}
                if("true".equals(dirty))new AlertDialog.Builder(this).setTitle("Leave ESG Banking?").setMessage("There are unsaved changes. Return to the app to save them before leaving.").setNegativeButton("Keep working",null).setPositiveButton("Leave",(dialog,which)->finish()).setOnDismissListener(dialog->backPending=false).show();
                else {backPending=false;finish();}
            });
        });
    }

    @SuppressWarnings("deprecation")
    private void createShell() {
        getWindow().setStatusBarColor(Color.TRANSPARENT);getWindow().setNavigationBarColor(Color.TRANSPARENT);
        if(Build.VERSION.SDK_INT>=30)getWindow().setDecorFitsSystemWindows(false);
        getWindow().getDecorView().setSystemUiVisibility(View.SYSTEM_UI_FLAG_LAYOUT_STABLE|View.SYSTEM_UI_FLAG_LAYOUT_FULLSCREEN|View.SYSTEM_UI_FLAG_LAYOUT_HIDE_NAVIGATION|View.SYSTEM_UI_FLAG_LIGHT_STATUS_BAR|View.SYSTEM_UI_FLAG_LIGHT_NAVIGATION_BAR);
        root=new LinearLayout(this);root.setOrientation(LinearLayout.VERTICAL);root.setBackgroundColor(Color.rgb(243,247,248));
        progress=new ProgressBar(this,null,android.R.attr.progressBarStyleHorizontal);progress.setMax(100);root.addView(progress,new LinearLayout.LayoutParams(-1,dp(3)));
        content=new FrameLayout(this);root.addView(content,new LinearLayout.LayoutParams(-1,0,1));
        errorPanel=new LinearLayout(this);errorPanel.setOrientation(LinearLayout.VERTICAL);errorPanel.setGravity(Gravity.CENTER);errorPanel.setPadding(dp(28),dp(28),dp(28),dp(28));errorPanel.setBackgroundColor(Color.rgb(243,247,248));
        errorText=new TextView(this);errorText.setTextSize(18);errorText.setTextColor(Color.rgb(18,60,75));errorText.setGravity(Gravity.CENTER);errorPanel.addView(errorText,new LinearLayout.LayoutParams(-1,-2));
        Button update=new Button(this);update.setText("Update WebView / Chrome");update.setId(View.generateViewId());update.setTag("update");update.setOnClickListener(v->openExternalHttps(Uri.parse("https://play.google.com/store/apps/details?id=com.google.android.webview")));errorPanel.addView(update);
        Button close=new Button(this);close.setText("Close");close.setOnClickListener(v->finish());errorPanel.addView(close);errorPanel.setVisibility(View.GONE);content.addView(errorPanel,new FrameLayout.LayoutParams(-1,-1));setContentView(root);
        root.setOnApplyWindowInsetsListener((v,insets)->{
            if(Build.VERSION.SDK_INT>=30){Insets i=insets.getInsets(WindowInsets.Type.systemBars()|WindowInsets.Type.displayCutout()|WindowInsets.Type.ime());v.setPadding(i.left,i.top,i.right,i.bottom);return WindowInsets.CONSUMED;}
            v.setPadding(insets.getSystemWindowInsetLeft(),insets.getSystemWindowInsetTop(),insets.getSystemWindowInsetRight(),insets.getSystemWindowInsetBottom());return insets.consumeSystemWindowInsets();
        });root.requestApplyInsets();
    }

    private void showError(String message,boolean update) { ready=false;pageFailed=true;progress.setVisibility(View.GONE);errorText.setText(message);errorPanel.findViewWithTag("update").setVisibility(update?View.VISIBLE:View.GONE);errorPanel.setVisibility(View.VISIBLE);errorPanel.bringToFront(); }
    private int dp(int value){return Math.round(value*getResources().getDisplayMetrics().density);}
    private void toast(String message){if(!destroyed)Toast.makeText(this,message,Toast.LENGTH_LONG).show();}
    @Override protected void onDestroy(){destroyed=true;finishFileChooser(null);main.removeCallbacksAndMessages(null);io.shutdown();if(web!=null){web.stopLoading();content.removeView(web);web.destroy();web=null;}super.onDestroy();}
}
