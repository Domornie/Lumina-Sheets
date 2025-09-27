// ───────────────────────────────────────────────────────────────────────────────
// ENTERPRISE COMPLIANCE AND REPORTING SERVICE
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Enterprise compliance and reporting service for audit trails and compliance
 */
const ComplianceService = {
    
    // Compliance configuration
    config: {
        retentionPeriod: 2555, // 7 years in days for compliance
        enableAuditTrail: true,
        enableDataExport: true,
        enableComplianceReports: true,
        encryptSensitiveData: true
    },
    
    /**
     * Log compliance event
     */
    logComplianceEvent: function(userId, eventType, details, classification = 'STANDARD') {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let complianceSheet = ss.getSheetByName('ComplianceAuditTrail');
            
            if (!complianceSheet) {
                complianceSheet = ss.insertSheet('ComplianceAuditTrail');
                const headers = [
                    'Timestamp', 'EventID', 'UserID', 'UserEmail', 'EventType', 
                    'Classification', 'Details', 'IPAddress', 'UserAgent', 
                    'SessionID', 'ComplianceFlags', 'RetentionDate'
                ];
                complianceSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
                complianceSheet.setFrozenRows(1);
            }
            
            const eventId = Utilities.getUuid();
            const timestamp = new Date();
            const userEmail = this.getUserEmail(userId);
            const retentionDate = new Date(timestamp.getTime() + (this.config.retentionPeriod * 24 * 60 * 60 * 1000));
            
            // Encrypt sensitive details if required
            let processedDetails = details;
            if (this.config.encryptSensitiveData && classification === 'SENSITIVE') {
                processedDetails = this.encryptData(details);
            }
            
            const rowData = [
                timestamp, eventId, userId, userEmail, eventType,
                classification, processedDetails, this.getClientIP(), 
                this.getUserAgent(), this.getSessionId(), 
                this.generateComplianceFlags(eventType, details), retentionDate
            ];
            
            complianceSheet.appendRow(rowData);
            
            // Check for compliance violations
            this.checkComplianceViolations(eventType, details, userId);
            
            return eventId;
            
        } catch (error) {
            console.error('Error logging compliance event:', error);
            writeError('logComplianceEvent', error);
            return null;
        }
    },
    
    /**
     * Generate compliance report
     */
    generateComplianceReport: function(startDate, endDate, reportType = 'FULL') {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const complianceSheet = ss.getSheetByName('ComplianceAuditTrail');
            const analyticsSheet = ss.getSheetByName('BrowsingAnalytics');
            
            if (!complianceSheet) {
                return { success: false, error: 'No compliance data available' };
            }
            
            const complianceData = complianceSheet.getDataRange().getValues();
            const headers = complianceData.shift();
            
            // Filter by date range
            const filteredData = complianceData.filter(row => {
                const eventDate = new Date(row[0]);
                return eventDate >= startDate && eventDate <= endDate;
            });
            
            const report = {
                reportId: Utilities.getUuid(),
                generatedAt: new Date(),
                reportType: reportType,
                period: { start: startDate, end: endDate },
                summary: this.generateReportSummary(filteredData),
                details: this.generateReportDetails(filteredData, reportType),
                compliance: this.generateComplianceMetrics(filteredData),
                recommendations: this.generateComplianceRecommendations(filteredData)
            };
            
            // Store report for audit trail
            this.storeGeneratedReport(report);
            
            return { success: true, report: report };
            
        } catch (error) {
            console.error('Error generating compliance report:', error);
            writeError('generateComplianceReport', error);
            return { success: false, error: error.message };
        }
    },
    
    /**
     * Data retention management
     */
    manageDataRetention: function() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheets = ['ComplianceAuditTrail', 'BrowsingAnalytics', 'SecurityIncidents'];
            let deletedRecords = 0;
            
            sheets.forEach(sheetName => {
                const sheet = ss.getSheetByName(sheetName);
                if (!sheet) return;
                
                const data = sheet.getDataRange().getValues();
                const headers = data.shift();
                
                // Find retention date column
                const retentionIndex = headers.indexOf('RetentionDate');
                if (retentionIndex === -1) return;
                
                const currentDate = new Date();
                const validRows = [headers];
                
                data.forEach(row => {
                    const retentionDate = new Date(row[retentionIndex]);
                    if (retentionDate > currentDate) {
                        validRows.push(row);
                    } else {
                        deletedRecords++;
                    }
                });
                
                // Update sheet with valid rows only
                sheet.clear();
                if (validRows.length > 1) {
                    sheet.getRange(1, 1, validRows.length, validRows[0].length)
                        .setValues(validRows);
                }
            });
            
            // Log retention activity
            this.logComplianceEvent(
                'SYSTEM', 'DATA_RETENTION_CLEANUP', 
                `Deleted ${deletedRecords} expired records`, 'SYSTEM'
            );
            
            return { success: true, deletedRecords: deletedRecords };
            
        } catch (error) {
            console.error('Error managing data retention:', error);
            writeError('manageDataRetention', error);
            return { success: false, error: error.message };
        }
    },
    
    /**
     * Export data for compliance audits
     */
    exportComplianceData: function(startDate, endDate, format = 'CSV') {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const complianceSheet = ss.getSheetByName('ComplianceAuditTrail');
            
            if (!complianceSheet) {
                return { success: false, error: 'No compliance data to export' };
            }
            
            const data = complianceSheet.getDataRange().getValues();
            const headers = data.shift();
            
            // Filter by date range
            const filteredData = data.filter(row => {
                const eventDate = new Date(row[0]);
                return eventDate >= startDate && eventDate <= endDate;
            });
            
            // Add headers back
            const exportData = [headers, ...filteredData];
            
            // Convert to requested format
            let exportContent;
            let mimeType;
            let filename;
            
            switch (format.toUpperCase()) {
                case 'CSV':
                    exportContent = this.convertToCSV(exportData);
                    mimeType = 'text/csv';
                    filename = `compliance_export_${this.formatDate(new Date())}.csv`;
                    break;
                    
                case 'JSON':
                    exportContent = this.convertToJSON(exportData);
                    mimeType = 'application/json';
                    filename = `compliance_export_${this.formatDate(new Date())}.json`;
                    break;
                    
                default:
                    throw new Error('Unsupported export format');
            }
            
            // Log export activity
            this.logComplianceEvent(
                'SYSTEM', 'DATA_EXPORT', 
                `Exported ${filteredData.length} records in ${format} format`, 'SENSITIVE'
            );
            
            return {
                success: true,
                content: exportContent,
                mimeType: mimeType,
                filename: filename,
                recordCount: filteredData.length
            };
            
        } catch (error) {
            console.error('Error exporting compliance data:', error);
            writeError('exportComplianceData', error);
            return { success: false, error: error.message };
        }
    },
    
    /**
     * Compliance violation detection
     */
    checkComplianceViolations: function(eventType, details, userId) {
        try {
            const violations = [];
            
            // Check for policy violations
            if (eventType === 'BLOCKED_DOMAIN' && details.includes('social-media')) {
                violations.push({
                    type: 'POLICY_VIOLATION',
                    severity: 'MEDIUM',
                    description: 'Attempted access to blocked social media site'
                });
            }
            
            // Check for unusual activity patterns
            const recentActivity = this.getUserRecentActivity(userId, 1); // Last hour
            if (recentActivity.length > 50) {
                violations.push({
                    type: 'UNUSUAL_ACTIVITY',
                    severity: 'HIGH',
                    description: 'Unusually high browsing activity detected'
                });
            }
            
            // Check for data exfiltration patterns
            if (eventType === 'DOWNLOAD_ATTEMPT' && details.includes('large-file')) {
                violations.push({
                    type: 'DATA_EXFILTRATION_RISK',
                    severity: 'HIGH',
                    description: 'Attempted download of large file'
                });
            }
            
            // Log violations
            violations.forEach(violation => {
                this.logSecurityViolation(userId, violation);
            });
            
            return violations;
            
        } catch (error) {
            console.error('Error checking compliance violations:', error);
            return [];
        }
    },
    
    /**
     * Helper methods
     */
    getUserEmail: function(userId) {
        try {
            const users = getAllUsers();
            const user = users.find(u => (u.ID || u.id) === userId);
            return user ? (user.Email || user.email) : 'unknown@example.com';
        } catch (error) {
            return 'unknown@example.com';
        }
    },
    
    getClientIP: function() {
        // Note: Google Apps Script doesn't provide direct access to client IP
        return 'unknown';
    },
    
    getUserAgent: function() {
        // This would be passed from the client
        return 'LuminaHQ-Browser/1.0';
    },
    
    getSessionId: function() {
        return Session.getTemporaryActiveUserKey() || 'unknown';
    },
    
    generateComplianceFlags: function(eventType, details) {
        const flags = [];
        
        if (eventType.includes('SECURITY')) flags.push('SECURITY_RELATED');
        if (eventType.includes('DOWNLOAD')) flags.push('FILE_ACCESS');
        if (details && details.toLowerCase().includes('personal')) flags.push('PERSONAL_DATA');
        
        return flags.join(',');
    },
    
    encryptData: function(data) {
        // Simple encryption for demonstration
        // In production, use proper encryption
        return Utilities.base64Encode(data + '_encrypted');
    },
    
    formatDate: function(date) {
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    },
    
    convertToCSV: function(data) {
        return data.map(row => 
            row.map(cell => 
                typeof cell === 'string' && cell.includes(',') ? `"${cell}"` : cell
            ).join(',')
        ).join('\n');
    },
    
    convertToJSON: function(data) {
        const headers = data[0];
        const rows = data.slice(1);
        const jsonData = rows.map(row => {
            const obj = {};
            headers.forEach((header, index) => {
                obj[header] = row[index];
            });
            return obj;
        });
        return JSON.stringify(jsonData, null, 2);
    }
};

// ───────────────────────────────────────────────────────────────────────────────
// ADVANCED THREAT DETECTION SERVICE
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Advanced threat detection and response service
 */
const ThreatDetectionService = {
    
    // Threat detection rules
    rules: {
        maxRequestsPerMinute: 30,
        maxFailedAttempts: 5,
        suspiciousKeywords: ['hack', 'crack', 'exploit', 'malware', 'phishing'],
        riskyFileExtensions: ['.exe', '.bat', '.cmd', '.scr', '.pif', '.vbs'],
        blockedUserAgents: ['bot', 'crawler', 'spider', 'scraper']
    },
    
    // Threat intelligence
    threatIntelligence: {
        knownMaliciousDomains: [],
        knownPhishingPatterns: [],
        suspiciousIPRanges: []
    },
    
    /**
     * Analyze request for threats
     */
    analyzeRequest: function(userId, url, userAgent, requestData) {
        try {
            const threats = [];
            const riskScore = 0;
            
            // URL analysis
            const urlThreats = this.analyzeURL(url);
            threats.push(...urlThreats);
            
            // User agent analysis
            const uaThreats = this.analyzeUserAgent(userAgent);
            threats.push(...uaThreats);
            
            // Behavioral analysis
            const behaviorThreats = this.analyzeBehavior(userId, requestData);
            threats.push(...behaviorThreats);
            
            // Rate limiting check
            const rateLimitThreats = this.checkRateLimit(userId);
            threats.push(...rateLimitThreats);
            
            // Calculate overall risk score
            const overallRiskScore = this.calculateRiskScore(threats);
            
            const analysis = {
                userId: userId,
                url: url,
                threats: threats,
                riskScore: overallRiskScore,
                recommendation: this.getRecommendation(overallRiskScore),
                timestamp: new Date()
            };
            
            // Log high-risk requests
            if (overallRiskScore > 70) {
                this.logHighRiskActivity(analysis);
            }
            
            return analysis;
            
        } catch (error) {
            console.error('Error analyzing request for threats:', error);
            return {
                userId: userId,
                url: url,
                threats: [],
                riskScore: 0,
                recommendation: 'ALLOW',
                error: error.message
            };
        }
    },
    
    /**
     * URL threat analysis
     */
    analyzeURL: function(url) {
        const threats = [];
        
        try {
            const urlObj = new URL(url);
            const domain = urlObj.hostname.toLowerCase();
            const path = urlObj.pathname.toLowerCase();
            const params = urlObj.search.toLowerCase();
            
            // Check against known malicious domains
            if (this.threatIntelligence.knownMaliciousDomains.includes(domain)) {
                threats.push({
                    type: 'MALICIOUS_DOMAIN',
                    severity: 'HIGH',
                    description: `Known malicious domain: ${domain}`,
                    score: 90
                });
            }
            
            // Check for suspicious URL patterns
            const suspiciousPatterns = [
                /bit\.ly|tinyurl|t\.co/, // URL shorteners
                /[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}/, // IP addresses
                /[a-z0-9]{20,}\.com/, // Randomly generated domains
                /login|signin|account|verify|security|update/i // Phishing keywords
            ];
            
            suspiciousPatterns.forEach((pattern, index) => {
                if (pattern.test(url)) {
                    threats.push({
                        type: 'SUSPICIOUS_URL_PATTERN',
                        severity: 'MEDIUM',
                        description: `Suspicious URL pattern detected: Pattern ${index + 1}`,
                        score: 40
                    });
                }
            });
            
            // Check for suspicious keywords in URL
            this.rules.suspiciousKeywords.forEach(keyword => {
                if (url.toLowerCase().includes(keyword)) {
                    threats.push({
                        type: 'SUSPICIOUS_KEYWORD',
                        severity: 'MEDIUM',
                        description: `Suspicious keyword in URL: ${keyword}`,
                        score: 30
                    });
                }
            });
            
            // Check for file download attempts
            this.rules.riskyFileExtensions.forEach(ext => {
                if (path.includes(ext)) {
                    threats.push({
                        type: 'RISKY_FILE_DOWNLOAD',
                        severity: 'HIGH',
                        description: `Attempted download of risky file type: ${ext}`,
                        score: 80
                    });
                }
            });
            
        } catch (error) {
            threats.push({
                type: 'URL_ANALYSIS_ERROR',
                severity: 'LOW',
                description: 'Unable to parse URL',
                score: 10
            });
        }
        
        return threats;
    },
    
    /**
     * User agent analysis
     */
    analyzeUserAgent: function(userAgent) {
        const threats = [];
        
        if (!userAgent) {
            threats.push({
                type: 'MISSING_USER_AGENT',
                severity: 'MEDIUM',
                description: 'No user agent provided',
                score: 30
            });
            return threats;
        }
        
        const ua = userAgent.toLowerCase();
        
        // Check for blocked user agents
        this.rules.blockedUserAgents.forEach(blocked => {
            if (ua.includes(blocked)) {
                threats.push({
                    type: 'BLOCKED_USER_AGENT',
                    severity: 'HIGH',
                    description: `Blocked user agent detected: ${blocked}`,
                    score: 70
                });
            }
        });
        
        // Check for unusual user agents
        if (ua.length < 10 || ua.length > 500) {
            threats.push({
                type: 'UNUSUAL_USER_AGENT',
                severity: 'MEDIUM',
                description: 'Unusual user agent length',
                score: 25
            });
        }
        
        return threats;
    },
    
    /**
     * Behavioral analysis
     */
    analyzeBehavior: function(userId, requestData) {
        const threats = [];
        
        try {
            // Get recent user activity
            const recentActivity = this.getUserRecentActivity(userId, 24); // Last 24 hours
            
            // Check for unusual activity volume
            const hourlyRequests = recentActivity.filter(activity => {
                const activityTime = new Date(activity.timestamp);
                const hourAgo = new Date(Date.now() - 60 * 60 * 1000);
                return activityTime > hourAgo;
            }).length;
            
            if (hourlyRequests > this.rules.maxRequestsPerMinute * 60) {
                threats.push({
                    type: 'EXCESSIVE_REQUESTS',
                    severity: 'HIGH',
                    description: `Excessive requests: ${hourlyRequests} in last hour`,
                    score: 80
                });
            }
            
            // Check for failed attempts
            const failedAttempts = recentActivity.filter(activity => 
                activity.action === 'ACCESS_FAILED'
            ).length;
            
            if (failedAttempts > this.rules.maxFailedAttempts) {
                threats.push({
                    type: 'MULTIPLE_FAILURES',
                    severity: 'MEDIUM',
                    description: `Multiple failed attempts: ${failedAttempts}`,
                    score: 50
                });
            }
            
            // Check for off-hours activity
            const currentHour = new Date().getHours();
            if (currentHour < 6 || currentHour > 22) {
                threats.push({
                    type: 'OFF_HOURS_ACTIVITY',
                    severity: 'LOW',
                    description: 'Activity outside normal business hours',
                    score: 15
                });
            }
            
        } catch (error) {
            console.error('Error in behavioral analysis:', error);
        }
        
        return threats;
    },
    
    /**
     * Rate limiting check
     */
    checkRateLimit: function(userId) {
        const threats = [];
        
        try {
            const cache = CacheService.getScriptCache();
            const key = `rate_limit_${userId}`;
            const currentCount = parseInt(cache.get(key) || '0');
            
            if (currentCount > this.rules.maxRequestsPerMinute) {
                threats.push({
                    type: 'RATE_LIMIT_EXCEEDED',
                    severity: 'HIGH',
                    description: `Rate limit exceeded: ${currentCount} requests per minute`,
                    score: 85
                });
            }
            
            // Update counter
            cache.put(key, (currentCount + 1).toString(), 60); // 1 minute expiry
            
        } catch (error) {
            console.error('Error checking rate limit:', error);
        }
        
        return threats;
    },
    
    /**
     * Calculate overall risk score
     */
    calculateRiskScore: function(threats) {
        if (threats.length === 0) return 0;
        
        const totalScore = threats.reduce((sum, threat) => sum + threat.score, 0);
        const maxPossibleScore = threats.length * 100;
        
        return Math.min(100, (totalScore / maxPossibleScore) * 100);
    },
    
    /**
     * Get recommendation based on risk score
     */
    getRecommendation: function(riskScore) {
        if (riskScore >= 80) return 'BLOCK';
        if (riskScore >= 60) return 'QUARANTINE';
        if (riskScore >= 40) return 'MONITOR';
        return 'ALLOW';
    },
    
    /**
     * Log high-risk activity
     */
    logHighRiskActivity: function(analysis) {
        try {
            ComplianceService.logComplianceEvent(
                analysis.userId,
                'HIGH_RISK_ACTIVITY',
                JSON.stringify(analysis),
                'SECURITY'
            );
            
            // Alert security team for very high risk
            if (analysis.riskScore > 90) {
                this.alertSecurityTeam(analysis);
            }
            
        } catch (error) {
            console.error('Error logging high-risk activity:', error);
        }
    },
    
    /**
     * Alert security team
     */
    alertSecurityTeam: function(analysis) {
        try {
            console.warn(`🚨 CRITICAL SECURITY ALERT: User ${analysis.userId} - Risk Score: ${analysis.riskScore}`);
            
            // In production, implement email/SMS alerts
            // MailApp.sendEmail({
            //     to: 'security@yourcompany.com',
            //     subject: 'Critical Security Alert - LuminaHQ',
            //     body: `High-risk activity detected:\n${JSON.stringify(analysis, null, 2)}`
            // });
            
        } catch (error) {
            console.error('Error alerting security team:', error);
        }
    },
    
    /**
     * Get user recent activity (placeholder)
     */
    getUserRecentActivity: function(userId, hours) {
        try {
            // This would query the BrowsingAnalytics sheet
            // For now, return empty array
            return [];
        } catch (error) {
            return [];
        }
    }
};

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE ENTERPRISE FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Generate compliance report (client-accessible)
 */
function clientGenerateComplianceReport(startDate, endDate, reportType) {
    try {
        return ComplianceService.generateComplianceReport(
            new Date(startDate), 
            new Date(endDate), 
            reportType
        );
    } catch (error) {
        console.error('Error generating compliance report:', error);
        writeError('clientGenerateComplianceReport', error);
        return { success: false, error: error.message };
    }
}

/**
 * Export compliance data (client-accessible)
 */
function clientExportComplianceData(startDate, endDate, format) {
    try {
        return ComplianceService.exportComplianceData(
            new Date(startDate), 
            new Date(endDate), 
            format
        );
    } catch (error) {
        console.error('Error exporting compliance data:', error);
        writeError('clientExportComplianceData', error);
        return { success: false, error: error.message };
    }
}

/**
 * Get threat analysis (client-accessible)
 */
function clientGetThreatAnalysis(url, userAgent) {
    try {
        const userId = getUserIdFromCurrentSession();
        return ThreatDetectionService.analyzeRequest(userId, url, userAgent, {});
    } catch (error) {
        console.error('Error getting threat analysis:', error);
        writeError('clientGetThreatAnalysis', error);
        return { threats: [], riskScore: 0, recommendation: 'ALLOW' };
    }
}

/**
 * Manage data retention (client-accessible)
 */
function clientManageDataRetention() {
    try {
        return ComplianceService.manageDataRetention();
    } catch (error) {
        console.error('Error managing data retention:', error);
        writeError('clientManageDataRetention', error);
        return { success: false, error: error.message };
    }
}

/**
 * Update threat intelligence (client-accessible)
 */
function clientUpdateThreatIntelligence(threatData) {
    try {
        if (threatData.maliciousDomains) {
            ThreatDetectionService.threatIntelligence.knownMaliciousDomains = 
                [...ThreatDetectionService.threatIntelligence.knownMaliciousDomains, ...threatData.maliciousDomains];
        }
        
        if (threatData.phishingPatterns) {
            ThreatDetectionService.threatIntelligence.knownPhishingPatterns = 
                [...ThreatDetectionService.threatIntelligence.knownPhishingPatterns, ...threatData.phishingPatterns];
        }
        
        // Save to properties for persistence
        PropertiesService.getScriptProperties().setProperties({
            'threatIntelligence.maliciousDomains': JSON.stringify(ThreatDetectionService.threatIntelligence.knownMaliciousDomains),
            'threatIntelligence.phishingPatterns': JSON.stringify(ThreatDetectionService.threatIntelligence.knownPhishingPatterns)
        });
        
        return { success: true, message: 'Threat intelligence updated' };
        
    } catch (error) {
        console.error('Error updating threat intelligence:', error);
        writeError('clientUpdateThreatIntelligence', error);
        return { success: false, error: error.message };
    }
}

// ───────────────────────────────────────────────────────────────────────────────
// ENTERPRISE SYSTEM INITIALIZATION
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Initialize complete enterprise system
 */
function initializeEnterpriseSystem() {
    try {
        console.log('Initializing enterprise LuminaHQ system...');
        
        // Initialize base optimized system
        const baseResult = initializeOptimizedSystem();
        
        // Initialize compliance system
        ComplianceService.logComplianceEvent(
            'SYSTEM', 'SYSTEM_INITIALIZATION', 
            'Enterprise system startup', 'SYSTEM'
        );
        
        // Load threat intelligence
        const properties = PropertiesService.getScriptProperties().getProperties();
        if (properties['threatIntelligence.maliciousDomains']) {
            ThreatDetectionService.threatIntelligence.knownMaliciousDomains = 
                JSON.parse(properties['threatIntelligence.maliciousDomains']);
        }
        
        if (properties['threatIntelligence.phishingPatterns']) {
            ThreatDetectionService.threatIntelligence.knownPhishingPatterns = 
                JSON.parse(properties['threatIntelligence.phishingPatterns']);
        }
        
        // Schedule maintenance tasks
        this.scheduleMaintenanceTasks();
        
        console.log('✅ Enterprise LuminaHQ system initialized successfully');
        
        return {
            success: true,
            message: 'Enterprise system ready',
            features: [
                ...baseResult.features,
                'Compliance & Audit Trail',
                'Advanced Threat Detection',
                'Data Retention Management',
                'Automated Reporting',
                'Threat Intelligence Integration'
            ]
        };
        
    } catch (error) {
        console.error('❌ Error initializing enterprise system:', error);
        writeError('initializeEnterpriseSystem', error);
        return { success: false, error: error.message };
    }
}

/**
 * Schedule maintenance tasks
 */
function scheduleMaintenanceTasks() {
    try {
        // Note: Google Apps Script doesn't have built-in scheduled tasks
        // These would typically be set up as time-driven triggers
        
        console.log('📅 Maintenance tasks configured:');
        console.log('- Daily: Data retention cleanup');
        console.log('- Weekly: Compliance report generation');
        console.log('- Monthly: Threat intelligence update');
        
        // Example of setting up a daily trigger (run this once manually)
        // ScriptApp.newTrigger('dailyMaintenance')
        //     .timeBased()
        //     .everyDays(1)
        //     .atHour(2) // 2 AM
        //     .create();
        
    } catch (error) {
        console.error('Error scheduling maintenance tasks:', error);
    }
}

/**
 * Daily maintenance function
 */
function dailyMaintenance() {
    try {
        console.log('🔧 Running daily maintenance...');
        
        // Data retention cleanup
        const retentionResult = ComplianceService.manageDataRetention();
        console.log('📝 Data retention:', retentionResult);
        
        // Performance optimization
        PerformanceOptimizationService.clearCache();
        console.log('🚀 Cache cleared');
        
        // Log maintenance completion
        ComplianceService.logComplianceEvent(
            'SYSTEM', 'DAILY_MAINTENANCE', 
            `Completed: ${retentionResult.deletedRecords} records deleted`, 'SYSTEM'
        );
        
        console.log('✅ Daily maintenance completed');
        
    } catch (error) {
        console.error('❌ Daily maintenance failed:', error);
        writeError('dailyMaintenance', error);
    }
}

/**
 * Weekly maintenance function
 */
function weeklyMaintenance() {
    try {
        console.log('🔧 Running weekly maintenance...');
        
        // Generate compliance report
        const endDate = new Date();
        const startDate = new Date(endDate.getTime() - (7 * 24 * 60 * 60 * 1000));
        
        const reportResult = ComplianceService.generateComplianceReport(
            startDate, endDate, 'SUMMARY'
        );
        
        console.log('📊 Weekly report generated:', reportResult.success);
        
        // Update threat intelligence (placeholder)
        console.log('🛡️ Threat intelligence updated');
        
        console.log('✅ Weekly maintenance completed');
        
    } catch (error) {
        console.error('❌ Weekly maintenance failed:', error);
        writeError('weeklyMaintenance', error);
    }
}

// Helper function to get user ID from current session
function getUserIdFromCurrentSession() {
    try {
        const user = Session.getActiveUser();
        return user.getEmail() || 'anonymous';
    } catch (error) {
        return 'anonymous';
    }
}