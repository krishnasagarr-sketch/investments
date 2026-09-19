package com.fdmanager.app

import android.os.Bundle
import android.os.Handler
import android.os.Looper
import android.webkit.WebResourceError
import android.webkit.WebResourceRequest
import android.webkit.WebView
import android.webkit.WebViewClient
import androidx.appcompat.app.AppCompatActivity
import com.chaquo.python.Python
import com.chaquo.python.android.AndroidPlatform
import kotlin.concurrent.thread

private const val SERVER_URL = "http://127.0.0.1:5000"

class MainActivity : AppCompatActivity() {
    private lateinit var webView: WebView
    private var serverStarted = false

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        if (!Python.isStarted()) {
            Python.start(AndroidPlatform(this))
        }
        val python = Python.getInstance()

        // app.py reads this to know where it may write the database, session
        // key and ticker cache — Context.filesDir is this app's own private,
        // writable storage, created automatically by Android.
        python.getModule("os").get("environ")!!
            .callAttr("__setitem__", "FDMANAGER_DATA_DIR", filesDir.absolutePath)

        // app.run() blocks forever (it's a real Flask dev server), so it
        // needs its own thread — everything else here runs on the UI thread.
        thread(isDaemon = true) {
            python.getModule("app").callAttr("main")
        }

        webView = WebView(this)
        setContentView(webView)
        webView.settings.javaScriptEnabled = true
        webView.settings.domStorageEnabled = true
        webView.webViewClient = object : WebViewClient() {
            override fun onReceivedError(
                view: WebView?,
                request: WebResourceRequest?,
                error: WebResourceError?
            ) {
                // First launch: the Flask thread may not be listening on
                // :5000 yet by the time the WebView tries to load it. Retry
                // until it comes up rather than showing a dead error page.
                if (request?.isForMainFrame != false) {
                    Handler(Looper.getMainLooper()).postDelayed(
                        { view?.loadUrl(SERVER_URL) },
                        400,
                    )
                }
            }

            override fun onPageFinished(view: WebView?, url: String?) {
                serverStarted = true
            }
        }
        webView.loadUrl(SERVER_URL)
    }

    @Suppress("DEPRECATION")
    override fun onBackPressed() {
        if (::webView.isInitialized && webView.canGoBack()) {
            webView.goBack()
        } else {
            super.onBackPressed()
        }
    }
}
