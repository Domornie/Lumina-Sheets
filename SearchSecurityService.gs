// ───────────────────────────────────────────────────────────────────────────────
// ADVANCED CONTENT FILTERING AND SECURITY SERVICE
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Content filtering and security service for enhanced proxy
 */
const ContentFilterService = {
    
    // Blocked domains for security
    blockedDomains: [
        'malware-site.com',
        'phishing-site.com',
        // Add your company's blocked domains here
    ],
    
    // Allowed domains (whitelist mode - optional)
    allowedDomains: [
        // Leave empty for open browsing, or add specific domains for whitelist mode
    ],
    
    // Content type restrictions
    blockedContentTypes: [
        'application/x-executable',
        'application/x-msdownload',
        'application/x-msdos-program'
    ],
    
    /**
     * Check if domain is allowed
     */
    isDomainAllowed: function(url) {
        try {
            const domain = new URL(url).hostname.toLowerCase();
            
            // Check blocked list
            if (this.blockedDomains.some(blocked => domain.includes(blocked))) {
                return { allowed: false, reason: 'Domain is blocked by security policy' };
            }
            
            // Check whitelist if configured
            if (this.allowedDomains.length > 0) {
                const isWhitelisted = this.allowedDomains.some(allowed => 
                    domain.includes(allowed.toLowerCase())
                );
                if (!isWhitelisted) {
                    return { allowed: false, reason: 'Domain not in allowed list' };
                }
            }
            
            return { allowed: true };
            
        } catch (error) {
            return { allowed: false, reason: 'Invalid URL format' };
        }
    },
    
    /**
     * Filter and sanitize HTML content
     */
    filterContent: function(html, sourceUrl) {
        try {
            // Remove potentially dangerous scripts
            html = html.replace(/<script[^>]*src=[^>]*cryptocurrency[^>]*><\/script>/gi, '');
            html = html.replace(/<script[^>]*>[\s\S]*?crypto[\s\S]*?<\/script>/gi, '');
            html = html.replace(/<script[^>]*>[\s\S]*?bitcoin[\s\S]*?<\/script>/gi, '');
            
            // Remove download links for executables
            html = html.replace(/<a[^>]*href=[^>]*\.(exe|msi|dmg|pkg)[^>]*>[\s\S]*?<\/a>/gi, '');
            
            // Add security headers injection
            const securityScript = this.getSecurityScript(sourceUrl);
            html = html.replace('</head>', securityScript + '</head>');
            
            return html;
            
        } catch (error) {
            console.error('Error filtering content:', error);
            return html; // Return original if filtering fails
        }
    },
    
    /**
     * Get security enhancement script
     */
    getSecurityScript: function(sourceUrl) {
        return `
            <script>
            (function() {
                // Enhanced security measures
                
                // Disable right-click context menu on sensitive areas
                document.addEventListener('contextmenu', function(e) {
                    // Allow normal right-click for now, but could be customized
                });
                
                // Monitor for suspicious activity
                let suspiciousActivity = 0;
                
                // Track download attempts
                document.addEventListener('click', function(e) {
                    const target = e.target.closest('a');
                    if (target && target.href) {
                        const href = target.href.toLowerCase();
                        if (href.includes('.exe') || href.includes('.msi') || 
                            href.includes('.dmg') || href.includes('.pkg')) {
                            e.preventDefault();
                            console.warn('Download blocked by security policy');
                            alert('Downloads are not permitted through the proxy browser.');
                            return false;
                        }
                    }
                });
                
                // Report suspicious activity
                window.addEventListener('beforeunload', function() {
                    if (suspiciousActivity > 5) {
                        // Could report to admin - implement as needed
                        console.warn('Suspicious activity detected on: ${sourceUrl}');
                    }
                });
                
                // Prevent common attacks
                Object.defineProperty(window, 'eval', {
                    value: function() {
                        console.warn('eval() blocked by security policy');
                        return null;
                    },
                    writable: false,
                    configurable: false
                });
                
            })();
            </script>
        `;
    }
};

// ───────────────────────────────────────────────────────────────────────────────
// BROWSING ANALYTICS AND MONITORING SERVICE
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Analytics service for tracking browsing patterns and security
 */
const BrowsingAnalyticsService = {
    
    /**
     * Log browsing activity
     */
    logActivity: function(userId, url, action, userAgent = '', duration = 0) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let analyticsSheet = ss.getSheetByName('BrowsingAnalytics');
            
            if (!analyticsSheet) {
                analyticsSheet = ss.insertSheet('BrowsingAnalytics');
                const headers = [
                    'Timestamp', 'UserID', 'UserEmail', 'URL', 'Domain', 
                    'Action', 'UserAgent', 'Duration', 'Success', 'ErrorReason'
                ];
                analyticsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
                analyticsSheet.setFrozenRows(1);
            }
            
            const domain = this.extractDomain(url);
            const timestamp = new Date();
            const userEmail = this.getUserEmail(userId);
            
            const rowData = [
                timestamp, userId, userEmail, url, domain,
                action, userAgent, duration, true, ''
            ];
            
            analyticsSheet.appendRow(rowData);
            
        } catch (error) {
            console.error('Error logging browsing activity:', error);
            writeError('logBrowsingActivity', error);
        }
    },
    
    /**
     * Log failed attempts or security incidents
     */
    logSecurityIncident: function(userId, url, incidentType, details) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let securitySheet = ss.getSheetByName('SecurityIncidents');
            
            if (!securitySheet) {
                securitySheet = ss.insertSheet('SecurityIncidents');
                const headers = [
                    'Timestamp', 'UserID', 'UserEmail', 'URL', 'IncidentType', 
                    'Details', 'Severity', 'Resolved'
                ];
                securitySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
                securitySheet.setFrozenRows(1);
            }
            
            const timestamp = new Date();
            const userEmail = this.getUserEmail(userId);
            const severity = this.calculateSeverity(incidentType);
            
            const rowData = [
                timestamp, userId, userEmail, url, incidentType,
                details, severity, false
            ];
            
            securitySheet.appendRow(rowData);
            
            // Alert admins for high severity incidents
            if (severity === 'HIGH') {
                this.alertAdmins(userId, url, incidentType, details);
            }
            
        } catch (error) {
            console.error('Error logging security incident:', error);
        }
    },
    
    /**
     * Get browsing analytics for dashboard
     */
    getBrowsingAnalytics: function(days = 7) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const analyticsSheet = ss.getSheetByName('BrowsingAnalytics');
            
            if (!analyticsSheet) {
                return this.getEmptyAnalytics();
            }
            
            const data = analyticsSheet.getDataRange().getValues();
            const headers = data.shift();
            
            const cutoffDate = new Date();
            cutoffDate.setDate(cutoffDate.getDate() - days);
            
            const recentData = data.filter(row => {
                const timestamp = new Date(row[0]);
                return timestamp >= cutoffDate;
            });
            
            return {
                totalRequests: recentData.length,
                uniqueUsers: [...new Set(recentData.map(row => row[1]))].length,
                topDomains: this.getTopDomains(recentData),
                timeDistribution: this.getTimeDistribution(recentData),
                userActivity: this.getUserActivity(recentData),
                errorRate: this.getErrorRate(recentData),
                averageDuration: this.getAverageDuration(recentData)
            };
            
        } catch (error) {
            console.error('Error getting browsing analytics:', error);
            return this.getEmptyAnalytics();
        }
    },
    
    /**
     * Helper methods
     */
    extractDomain: function(url) {
        try {
            return new URL(url).hostname;
        } catch (error) {
            return 'invalid-url';
        }
    },
    
    getUserEmail: function(userId) {
        try {
            const users = getAllUsers();
            const user = users.find(u => (u.ID || u.id) === userId);
            return user ? (user.Email || user.email) : 'unknown@example.com';
        } catch (error) {
            return 'unknown@example.com';
        }
    },
    
    calculateSeverity: function(incidentType) {
        const highSeverity = ['MALWARE_DETECTED', 'PHISHING_ATTEMPT', 'UNAUTHORIZED_DOWNLOAD'];
        const mediumSeverity = ['BLOCKED_DOMAIN', 'SUSPICIOUS_ACTIVITY'];
        
        if (highSeverity.includes(incidentType)) return 'HIGH';
        if (mediumSeverity.includes(incidentType)) return 'MEDIUM';
        return 'LOW';
    },
    
    alertAdmins: function(userId, url, incidentType, details) {
        try {
            // Implement admin alerting based on your notification system
            console.warn(`SECURITY ALERT: ${incidentType} by user ${userId} on ${url}`);
            
            // Could send email to admins here
            // MailApp.sendEmail({
            //     to: 'admin@yourcompany.com',
            //     subject: `Security Alert: ${incidentType}`,
            //     body: `User ${userId} triggered ${incidentType} on ${url}. Details: ${details}`
            // });
            
        } catch (error) {
            console.error('Error alerting admins:', error);
        }
    },
    
    getEmptyAnalytics: function() {
        return {
            totalRequests: 0,
            uniqueUsers: 0,
            topDomains: [],
            timeDistribution: [],
            userActivity: [],
            errorRate: 0,
            averageDuration: 0
        };
    },
    
    getTopDomains: function(data) {
        const domainCounts = {};
        data.forEach(row => {
            const domain = row[4]; // Domain column
            domainCounts[domain] = (domainCounts[domain] || 0) + 1;
        });
        
        return Object.entries(domainCounts)
            .sort(([,a], [,b]) => b - a)
            .slice(0, 10)
            .map(([domain, count]) => ({ domain, count }));
    },
    
    getTimeDistribution: function(data) {
        const hourCounts = new Array(24).fill(0);
        data.forEach(row => {
            const hour = new Date(row[0]).getHours();
            hourCounts[hour]++;
        });
        
        return hourCounts.map((count, hour) => ({ hour, count }));
    },
    
    getUserActivity: function(data) {
        const userCounts = {};
        data.forEach(row => {
            const userId = row[1];
            userCounts[userId] = (userCounts[userId] || 0) + 1;
        });
        
        return Object.entries(userCounts)
            .sort(([,a], [,b]) => b - a)
            .slice(0, 10)
            .map(([userId, count]) => ({ userId, count }));
    },
    
    getErrorRate: function(data) {
        const totalRequests = data.length;
        const errors = data.filter(row => row[8] === false).length; // Success column
        return totalRequests > 0 ? (errors / totalRequests) * 100 : 0;
    },
    
    getAverageDuration: function(data) {
        const durations = data.map(row => parseFloat(row[7]) || 0).filter(d => d > 0);
        return durations.length > 0 ? durations.reduce((a, b) => a + b, 0) / durations.length : 0;
    }
};

// ───────────────────────────────────────────────────────────────────────────────
// ENHANCED PROXY WITH SECURITY AND ANALYTICS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Enhanced proxy service with security and analytics integration
 */
function serveSecureEnhancedProxy(e) {
    const startTime = new Date();
    const userId = getUserIdFromRequest(e);
    const targetUrl = e.parameter.url;
    
    try {
        if (!targetUrl) {
            return createProxyErrorPage('Missing URL', 'No URL specified for proxy request.');
        }

        // Security check
        const domainCheck = ContentFilterService.isDomainAllowed(targetUrl);
        if (!domainCheck.allowed) {
            // Log security incident
            BrowsingAnalyticsService.logSecurityIncident(
                userId, targetUrl, 'BLOCKED_DOMAIN', domainCheck.reason
            );
            
            return createSecurityBlockedPage(targetUrl, domainCheck.reason);
        }

        // Normalize URL
        const normalizedUrl = normalizeUrl(targetUrl);
        
        // Log access attempt
        BrowsingAnalyticsService.logActivity(
            userId, normalizedUrl, 'ACCESS_ATTEMPT', e.parameter.userAgent || ''
        );

        // Fetch content with enhanced options
        const response = fetchWithRetry(normalizedUrl);
        
        if (!response.success) {
            // Log failed attempt
            BrowsingAnalyticsService.logActivity(
                userId, normalizedUrl, 'ACCESS_FAILED', '', 
                (new Date() - startTime) / 1000
            );
            
            return createProxyErrorPage('Fetch Error', response.error);
        }

        const contentType = response.headers['Content-Type'] || response.headers['content-type'] || '';
        
        // Check content type restrictions
        if (ContentFilterService.blockedContentTypes.some(blocked => 
            contentType.toLowerCase().includes(blocked))) {
            
            BrowsingAnalyticsService.logSecurityIncident(
                userId, normalizedUrl, 'BLOCKED_CONTENT_TYPE', `Content type: ${contentType}`
            );
            
            return createSecurityBlockedPage(normalizedUrl, 'Content type not allowed');
        }

        const isHtml = contentType.toLowerCase().includes('text/html');

        if (isHtml) {
            // Process HTML content with security filtering
            let processedHtml = processHtmlContent(
                response.content, 
                normalizedUrl, 
                `${SCRIPT_URL}?page=proxy`
            );
            
            // Apply content filtering
            processedHtml = ContentFilterService.filterContent(processedHtml, normalizedUrl);
            
            // Log successful access
            const duration = (new Date() - startTime) / 1000;
            BrowsingAnalyticsService.logActivity(
                userId, normalizedUrl, 'ACCESS_SUCCESS', '', duration
            );
            
            return HtmlService.createHtmlOutput(processedHtml)
                .setTitle('LuminaHQ Secure Browser')
                .addMetaTag('viewport', 'width=device-width,initial-scale=1')
                .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } else {
            // Handle non-HTML content
            const processedContent = processNonHtmlContent(
                response.content, 
                contentType, 
                normalizedUrl, 
                `${SCRIPT_URL}?page=proxy`
            );
            
            // Log successful resource access
            BrowsingAnalyticsService.logActivity(
                userId, normalizedUrl, 'RESOURCE_ACCESS', '', 
                (new Date() - startTime) / 1000
            );
            
            return ContentService.createTextOutput(processedContent)
                .setMimeType(getMimeType(contentType));
        }

    } catch (error) {
        console.error('Secure proxy error:', error);
        writeError('serveSecureEnhancedProxy', error);
        
        // Log error
        BrowsingAnalyticsService.logActivity(
            userId, targetUrl, 'ACCESS_ERROR', '', 
            (new Date() - startTime) / 1000
        );
        
        return createProxyErrorPage('Proxy Error', error.message);
    }
}

/**
 * Create security blocked page
 */
function createSecurityBlockedPage(url, reason) {
    const html = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>Access Blocked - LuminaHQ Security</title>
            <meta name="viewport" content="width=device-width,initial-scale=1">
            <style>
                body {
                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                    background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);
                    color: white;
                    margin: 0;
                    padding: 50px 20px;
                    min-height: 100vh;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }
                .block-container {
                    background: rgba(255,255,255,0.1);
                    padding: 40px;
                    border-radius: 15px;
                    text-align: center;
                    max-width: 600px;
                    backdrop-filter: blur(10px);
                    box-shadow: 0 8px 32px rgba(0,0,0,0.1);
                }
                .block-icon {
                    font-size: 64px;
                    margin-bottom: 20px;
                }
                .block-title {
                    font-size: 28px;
                    font-weight: 600;
                    margin-bottom: 15px;
                }
                .block-message {
                    font-size: 16px;
                    opacity: 0.9;
                    line-height: 1.5;
                    margin-bottom: 20px;
                }
                .block-url {
                    background: rgba(0,0,0,0.2);
                    padding: 10px;
                    border-radius: 8px;
                    font-family: monospace;
                    margin: 20px 0;
                    word-break: break-all;
                }
                .block-reason {
                    background: rgba(255,255,255,0.1);
                    padding: 15px;
                    border-radius: 8px;
                    margin: 20px 0;
                    font-style: italic;
                }
                .block-actions {
                    display: flex;
                    gap: 15px;
                    justify-content: center;
                    flex-wrap: wrap;
                    margin-top: 30px;
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
            <div class="block-container">
                <div class="block-icon">🚫</div>
                <div class="block-title">Access Blocked</div>
                <div class="block-message">
                    This website has been blocked by your organization's security policy.
                </div>
                <div class="block-url">${url}</div>
                <div class="block-reason">
                    <strong>Reason:</strong> ${reason}
                </div>
                <div class="block-actions">
                    <button class="btn" onclick="history.back()">← Go Back</button>
                    <button class="btn" onclick="window.parent.postMessage({action: 'closeTab'}, '*')">✕ Close Tab</button>
                    <button class="btn" onclick="reportIssue()">📞 Report Issue</button>
                </div>
            </div>
            
            <script>
            function reportIssue() {
                // Implement your issue reporting system
                alert('Please contact your IT administrator to report this issue.');
            }
            </script>
        </body>
        </html>
    `;
    
    return HtmlService.createHtmlOutput(html)
        .setTitle('Access Blocked')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Get user ID from request
 */
function getUserIdFromRequest(e) {
    try {
        const token = e.parameter.token || '';
        if (token && typeof AuthenticationService !== 'undefined') {
            const user = AuthenticationService.validateToken(token);
            return user ? (user.ID || user.id || 'anonymous') : 'anonymous';
        }
        return 'anonymous';
    } catch (error) {
        return 'anonymous';
    }
}

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE ANALYTICS FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Client-accessible function to get browsing analytics
 */
function clientGetBrowsingAnalytics(days = 7) {
    try {
        return BrowsingAnalyticsService.getBrowsingAnalytics(days);
    } catch (error) {
        console.error('Error getting browsing analytics:', error);
        writeError('clientGetBrowsingAnalytics', error);
        return BrowsingAnalyticsService.getEmptyAnalytics();
    }
}

/**
 * Client-accessible function to update content filter settings
 */
function clientUpdateContentFilter(settings) {
    try {
        // Only admins can update filter settings
        // Implement admin check based on your auth system
        
        if (settings.blockedDomains) {
            ContentFilterService.blockedDomains = settings.blockedDomains;
        }
        
        if (settings.allowedDomains) {
            ContentFilterService.allowedDomains = settings.allowedDomains;
        }
        
        // Save settings to properties for persistence
        PropertiesService.getScriptProperties().setProperties({
            'contentFilter.blockedDomains': JSON.stringify(ContentFilterService.blockedDomains),
            'contentFilter.allowedDomains': JSON.stringify(ContentFilterService.allowedDomains)
        });
        
        return { success: true, message: 'Content filter updated successfully' };
        
    } catch (error) {
        console.error('Error updating content filter:', error);
        writeError('clientUpdateContentFilter', error);
        return { success: false, error: error.message };
    }
}

/**
 * Initialize analytics system on startup
 */
function initializeAnalyticsSystem() {
    try {
        // Load saved filter settings
        const properties = PropertiesService.getScriptProperties().getProperties();
        
        if (properties['contentFilter.blockedDomains']) {
            ContentFilterService.blockedDomains = JSON.parse(properties['contentFilter.blockedDomains']);
        }
        
        if (properties['contentFilter.allowedDomains']) {
            ContentFilterService.allowedDomains = JSON.parse(properties['contentFilter.allowedDomains']);
        }
        
        console.log('Analytics and security systems initialized');
        return true;
        
    } catch (error) {
        console.error('Error initializing analytics system:', error);
        writeError('initializeAnalyticsSystem', error);
        return false;
    }
}

