// ─────────────────────────────────────────────────────────────────────────────
// Enhanced ProxyService.gs
// Advanced proxy server for bypassing iframe restrictions and enabling 
// secure internal web browsing within LuminaHQ
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Enhanced proxy service that handles iframe restrictions and content rewriting
 */
function serveEnhancedProxy(e) {
    try {
        const targetUrl = e.parameter.url;
        const mode = e.parameter.mode || 'iframe'; // 'iframe' or 'content'
        const baseProxyUrl = `${SCRIPT_URL}?page=proxy`;

        if (!targetUrl) {
            return createProxyErrorPage('Missing URL', 'No URL specified for proxy request.');
        }

        // Validate URL
        let normalizedUrl;
        try {
            normalizedUrl = normalizeUrl(targetUrl);
        } catch (error) {
            return createProxyErrorPage('Invalid URL', `Invalid URL format: ${targetUrl}`);
        }

        console.log(`Proxying: ${normalizedUrl}`);

        // Fetch the content with enhanced options
        const response = fetchWithRetry(normalizedUrl);
        
        if (!response.success) {
            return createProxyErrorPage('Fetch Error', response.error);
        }

        const contentType = response.headers['Content-Type'] || response.headers['content-type'] || '';
        const isHtml = contentType.toLowerCase().includes('text/html');

        if (isHtml) {
            // Process HTML content
            const processedHtml = processHtmlContent(
                response.content, 
                normalizedUrl, 
                baseProxyUrl
            );
            
            return HtmlService.createHtmlOutput(processedHtml)
                .setTitle('LuminaHQ Browser')
                .addMetaTag('viewport', 'width=device-width,initial-scale=1')
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } else {
            // Handle non-HTML content (images, CSS, JS, etc.)
            const processedContent = processNonHtmlContent(
                response.content, 
                contentType, 
                normalizedUrl, 
                baseProxyUrl
            );
            
            return ContentService.createTextOutput(processedContent)
                .setMimeType(getMimeType(contentType));
        }

    } catch (error) {
        console.error('Enhanced proxy error:', error);
        writeError('serveEnhancedProxy', error);
        return createProxyErrorPage('Proxy Error', error.message);
    }
}

/**
 * Fetch URL with retry logic and enhanced error handling
 */
function fetchWithRetry(url, maxRetries = 3) {
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            const response = UrlFetchApp.fetch(url, {
                method: 'GET',
                muteHttpExceptions: true,
                followRedirects: true,
                headers: {
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
                    'Accept-Language': 'en-US,en;q=0.5',
                    'Accept-Encoding': 'gzip, deflate',
                    'Cache-Control': 'no-cache',
                    'Pragma': 'no-cache'
                }
            });

            const responseCode = response.getResponseCode();
            const responseHeaders = response.getAllHeaders();
            const content = response.getContentText();

            if (responseCode >= 200 && responseCode < 400) {
                return {
                    success: true,
                    content: content,
                    headers: responseHeaders,
                    statusCode: responseCode
                };
            } else if (responseCode >= 300 && responseCode < 400) {
                // Handle redirects manually if needed
                const location = responseHeaders['Location'] || responseHeaders['location'];
                if (location && attempt < maxRetries) {
                    console.log(`Redirecting from ${url} to ${location}`);
                    url = resolveUrl(location, url);
                    continue;
                }
            }

            return {
                success: false,
                error: `HTTP ${responseCode}: ${content}`,
                statusCode: responseCode
            };

        } catch (error) {
            console.error(`Fetch attempt ${attempt} failed:`, error);
            if (attempt === maxRetries) {
                return {
                    success: false,
                    error: `Failed after ${maxRetries} attempts: ${error.message}`
                };
            }
            // Wait before retry
            Utilities.sleep(1000 * attempt);
        }
    }
}

/**
 * Process HTML content to rewrite URLs and remove frame restrictions
 */
function processHtmlContent(html, baseUrl, proxyUrl) {
    try {
        // Remove frame-busting JavaScript and X-Frame-Options
        html = html.replace(/<script[^>]*>[\s\S]*?if\s*\(\s*top\s*[!=]=?\s*self\s*\)[\s\S]*?<\/script>/gi, '');
        html = html.replace(/<script[^>]*>[\s\S]*?top\.location[\s\S]*?<\/script>/gi, '');
        html = html.replace(/<script[^>]*>[\s\S]*?parent\.location[\s\S]*?<\/script>/gi, '');
        html = html.replace(/<script[^>]*>[\s\S]*?window\.top[\s\S]*?<\/script>/gi, '');

        // Inject proxy CSS and JavaScript
        const proxyScript = createProxyScript(baseUrl, proxyUrl);
        const proxyStyles = createProxyStyles();

        // Find head tag or create one
        if (html.includes('</head>')) {
            html = html.replace('</head>', `${proxyStyles}${proxyScript}</head>`);
        } else if (html.includes('<html>')) {
            html = html.replace('<html>', `<html><head>${proxyStyles}${proxyScript}</head>`);
        } else {
            html = `<head>${proxyStyles}${proxyScript}</head>${html}`;
        }

        // Add proxy toolbar
        const toolbar = createProxyToolbar(baseUrl);
        html = html.replace('<body>', `<body>${toolbar}`);

        // Rewrite all URLs to go through proxy
        html = rewriteUrls(html, baseUrl, proxyUrl);

        return html;

    } catch (error) {
        console.error('Error processing HTML content:', error);
        return html; // Return original if processing fails
    }
}

/**
 * Create proxy toolbar for navigation
 */
function createProxyToolbar(currentUrl) {
    return `
        <div id="lumina-proxy-toolbar" style="
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            height: 50px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            display: flex;
            align-items: center;
            padding: 0 15px;
            font-family: Arial, sans-serif;
            font-size: 14px;
            z-index: 999999;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            border-bottom: 1px solid rgba(255,255,255,0.2);
        ">
            <div style="display: flex; align-items: center; gap: 10px; flex: 1;">
                <img src="data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMjQiIGhlaWdodD0iMjQiIHZpZXdCb3g9IjAgMCAyNCAyNCIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTEyIDJMMTMuMDkgOC4yNkwyMCA5TDEzLjA5IDE1Ljc0TDEyIDIyTDEwLjkxIDE1Ljc0TDQgOUwxMC45MSA4LjI2TDEyIDJaIiBmaWxsPSJ3aGl0ZSIvPgo8L3N2Zz4K" alt="LuminaHQ" style="width: 24px; height: 24px;">
                <span style="font-weight: 600; color: #fff;">LuminaHQ Browser</span>
                <div style="margin-left: 20px; display: flex; align-items: center; gap: 8px;">
                    <button onclick="luminaProxy.goBack()" style="
                        background: rgba(255,255,255,0.2);
                        border: none;
                        color: white;
                        padding: 5px 10px;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 12px;
                    ">← Back</button>
                    <button onclick="luminaProxy.goForward()" style="
                        background: rgba(255,255,255,0.2);
                        border: none;
                        color: white;
                        padding: 5px 10px;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 12px;
                    ">Forward →</button>
                    <button onclick="luminaProxy.refresh()" style="
                        background: rgba(255,255,255,0.2);
                        border: none;
                        color: white;
                        padding: 5px 10px;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 12px;
                    ">⟳ Refresh</button>
                </div>
            </div>
            <div style="
                background: rgba(255,255,255,0.15);
                padding: 8px 12px;
                border-radius: 20px;
                font-size: 13px;
                max-width: 300px;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
            ">
                📍 ${currentUrl}
            </div>
            <button onclick="luminaProxy.close()" style="
                background: rgba(255,255,255,0.2);
                border: none;
                color: white;
                padding: 5px 10px;
                border-radius: 4px;
                cursor: pointer;
                margin-left: 15px;
                font-size: 12px;
            ">✕ Close</button>
        </div>
        <div style="height: 50px;"></div>
    `;
}

/**
 * Create proxy JavaScript for enhanced functionality
 */
function createProxyScript(baseUrl, proxyUrl) {
    return `
        <script>
        (function() {
            // Create lumina proxy namespace
            window.luminaProxy = {
                currentUrl: '${baseUrl}',
                proxyUrl: '${proxyUrl}',
                history: [],
                historyIndex: -1,
                
                navigate: function(url) {
                    if (!url.startsWith('http')) {
                        url = this.resolveUrl(url, this.currentUrl);
                    }
                    window.location.href = this.proxyUrl + '&url=' + encodeURIComponent(url);
                },
                
                goBack: function() {
                    if (window.history.length > 1) {
                        window.history.back();
                    }
                },
                
                goForward: function() {
                    window.history.forward();
                },
                
                refresh: function() {
                    window.location.reload();
                },
                
                close: function() {
                    if (window.parent !== window) {
                        window.parent.postMessage({action: 'closeTab'}, '*');
                    } else {
                        window.close();
                    }
                },
                
                resolveUrl: function(url, base) {
                    if (url.startsWith('http')) return url;
                    if (url.startsWith('//')) return 'https:' + url;
                    if (url.startsWith('/')) {
                        const baseUrl = new URL(base);
                        return baseUrl.protocol + '//' + baseUrl.host + url;
                    }
                    const baseUrl = new URL(base);
                    return new URL(url, baseUrl).href;
                }
            };
            
            // Override window.open to use proxy
            const originalOpen = window.open;
            window.open = function(url, name, features) {
                if (url && typeof url === 'string') {
                    const resolvedUrl = luminaProxy.resolveUrl(url, luminaProxy.currentUrl);
                    const proxyUrl = luminaProxy.proxyUrl + '&url=' + encodeURIComponent(resolvedUrl);
                    return originalOpen.call(this, proxyUrl, name, features);
                }
                return originalOpen.apply(this, arguments);
            };
            
            // Prevent frame busting
            try {
                Object.defineProperty(window, 'top', { value: window, writable: false });
                Object.defineProperty(window, 'parent', { value: window, writable: false });
            } catch(e) {}
            
            // Handle forms to submit through proxy
            document.addEventListener('DOMContentLoaded', function() {
                const forms = document.querySelectorAll('form');
                forms.forEach(function(form) {
                    if (form.action && !form.action.includes('${proxyUrl}')) {
                        const originalAction = form.action;
                        form.addEventListener('submit', function(e) {
                            e.preventDefault();
                            const resolvedAction = luminaProxy.resolveUrl(originalAction, luminaProxy.currentUrl);
                            form.action = luminaProxy.proxyUrl + '&url=' + encodeURIComponent(resolvedAction);
                            form.submit();
                        });
                    }
                });
            });
            
        })();
        </script>
    `;
}

/**
 * Create proxy styles
 */
function createProxyStyles() {
    return `
        <style>
        /* Ensure proxy toolbar is always visible */
        #lumina-proxy-toolbar {
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            right: 0 !important;
            z-index: 999999 !important;
        }
        
        /* Adjust body margin to account for toolbar */
        body {
            margin-top: 50px !important;
            padding-top: 0 !important;
        }
        
        /* Fix any absolute positioned elements that might overlap */
        body > * {
            position: relative !important;
        }
        
        /* Prevent iframe breaking scripts */
        iframe[src*="frame-breaking"] {
            display: none !important;
        }
        </style>
    `;
}

/**
 * Rewrite URLs in HTML content to go through proxy
 */
function rewriteUrls(html, baseUrl, proxyUrl) {
    const baseObj = parseUrl(baseUrl);
    const baseHost = baseObj.protocol + '//' + baseObj.host;
    
    // Rewrite href attributes
    html = html.replace(/href\s*=\s*["']([^"']+)["']/gi, function(match, url) {
        const resolvedUrl = resolveUrl(url, baseUrl);
        if (resolvedUrl.startsWith('http') && !resolvedUrl.includes(proxyUrl)) {
            return `href="${proxyUrl}&url=${encodeURIComponent(resolvedUrl)}"`;
        }
        return match;
    });
    
    // Rewrite src attributes for images, scripts, etc.
    html = html.replace(/src\s*=\s*["']([^"']+)["']/gi, function(match, url) {
        const resolvedUrl = resolveUrl(url, baseUrl);
        if (resolvedUrl.startsWith('http') && !resolvedUrl.includes(proxyUrl)) {
            return `src="${proxyUrl}&url=${encodeURIComponent(resolvedUrl)}"`;
        }
        return match;
    });
    
    // Rewrite action attributes in forms
    html = html.replace(/action\s*=\s*["']([^"']+)["']/gi, function(match, url) {
        const resolvedUrl = resolveUrl(url, baseUrl);
        if (resolvedUrl.startsWith('http') && !resolvedUrl.includes(proxyUrl)) {
            return `action="${proxyUrl}&url=${encodeURIComponent(resolvedUrl)}"`;
        }
        return match;
    });
    
    // Rewrite CSS url() references
    html = html.replace(/url\s*\(\s*["']?([^"']+)["']?\s*\)/gi, function(match, url) {
        const resolvedUrl = resolveUrl(url, baseUrl);
        if (resolvedUrl.startsWith('http') && !resolvedUrl.includes(proxyUrl)) {
            return `url("${proxyUrl}&url=${encodeURIComponent(resolvedUrl)}")`;
        }
        return match;
    });
    
    return html;
}

/**
 * Process non-HTML content (CSS, JS, images, etc.)
 */
function processNonHtmlContent(content, contentType, baseUrl, proxyUrl) {
    try {
        if (contentType.includes('text/css')) {
            // Rewrite URLs in CSS
            return content.replace(/url\s*\(\s*["']?([^"']+)["']?\s*\)/gi, function(match, url) {
                const resolvedUrl = resolveUrl(url, baseUrl);
                if (resolvedUrl.startsWith('http') && !resolvedUrl.includes(proxyUrl)) {
                    return `url("${proxyUrl}&url=${encodeURIComponent(resolvedUrl)}")`;
                }
                return match;
            });
        }
        
        return content;
    } catch (error) {
        console.error('Error processing non-HTML content:', error);
        return content;
    }
}

/**
 * Utility functions
 */
function normalizeUrl(url) {
    if (!url) throw new Error('URL is required');
    
    // Handle relative URLs
    if (url.startsWith('//')) {
        url = 'https:' + url;
    } else if (!url.startsWith('http')) {
        url = 'https://' + url;
    }
    
    // Validate URL format
    new URL(url); // This will throw if invalid
    
    return url;
}

function resolveUrl(url, base) {
    try {
        if (!url) return base;
        if (url.startsWith('http')) return url;
        if (url.startsWith('//')) return 'https:' + url;
        if (url.startsWith('#')) return base + url;
        
        const baseUrl = new URL(base);
        if (url.startsWith('/')) {
            return baseUrl.protocol + '//' + baseUrl.host + url;
        }
        
        return new URL(url, baseUrl).href;
    } catch (error) {
        console.error('Error resolving URL:', error);
        return url;
    }
}

function parseUrl(url) {
    try {
        const parsed = new URL(url);
        return {
            protocol: parsed.protocol,
            host: parsed.host,
            hostname: parsed.hostname,
            pathname: parsed.pathname,
            search: parsed.search,
            hash: parsed.hash
        };
    } catch (error) {
        return { protocol: 'https:', host: '', hostname: '', pathname: '/', search: '', hash: '' };
    }
}

function getMimeType(contentType) {
    if (contentType.includes('text/css')) return ContentService.MimeType.CSS;
    if (contentType.includes('application/javascript') || contentType.includes('text/javascript')) {
        return ContentService.MimeType.JAVASCRIPT;
    }
    if (contentType.includes('image/')) {
        if (contentType.includes('png')) return ContentService.MimeType.PNG;
        if (contentType.includes('jpeg') || contentType.includes('jpg')) return ContentService.MimeType.JPEG;
        if (contentType.includes('gif')) return ContentService.MimeType.GIF;
        if (contentType.includes('svg')) return ContentService.MimeType.SVG;
    }
    return ContentService.MimeType.TEXT;
}

function createProxyErrorPage(title, message) {
    const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>${title} - LuminaHQ Browser</title>
            <meta name="viewport" content="width=device-width,initial-scale=1">
            <style>
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    color: white;
                    margin: 0;
                    padding: 50px 20px;
                    min-height: 100vh;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
                .error-container {
                    background: rgba(255,255,255,0.1);
                    padding: 40px;
                    border-radius: 15px;
                    text-align: center;
                    max-width: 500px;
                    backdrop-filter: blur(10px);
                    box-shadow: 0 8px 32px rgba(0,0,0,0.1);
                }
                .error-icon {
                    font-size: 64px;
                    margin-bottom: 20px;
                }
                .error-title {
                    font-size: 24px;
                    font-weight: 600;
                    margin-bottom: 15px;
                }
                .error-message {
                    font-size: 16px;
                    opacity: 0.9;
                    line-height: 1.5;
                    margin-bottom: 30px;
                }
                .error-actions {
                    display: flex;
                    gap: 15px;
                    justify-content: center;
                    flex-wrap: wrap;
                }
                .btn {
                    background: rgba(255,255,255,0.2);
                    border: 1px solid rgba(255,255,255,0.3);
                    color: white;
                    padding: 12px 24px;
                    border-radius: 8px;
                    text-decoration: none;
                    font-weight: 500;
                    transition: all 0.3s ease;
                    cursor: pointer;
                }
                .btn:hover {
                    background: rgba(255,255,255,0.3);
                    transform: translateY(-2px);
                }
            </style>
        </head>
        <body>
            <div class="error-container">
                <div class="error-icon">🚫</div>
                <div class="error-title">${title}</div>
                <div class="error-message">${message}</div>
                <div class="error-actions">
                    <button class="btn" onclick="history.back()">← Go Back</button>
                    <button class="btn" onclick="window.parent.postMessage({action: 'closeTab'}, '*')">✕ Close Tab</button>
                </div>
            </div>
        </body>
        </html>
    `;
    
    return HtmlService.createHtmlOutput(html)
        .setTitle(title)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

