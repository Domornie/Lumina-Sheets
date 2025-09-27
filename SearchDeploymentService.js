// ───────────────────────────────────────────────────────────────────────────────
// DEPLOYMENT AUTOMATION & SYSTEM MONITORING SERVICE
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Automated deployment and system monitoring service
 */
const DeploymentService = {
    
    // Deployment configuration
    config: {
        version: '2.0.0',
        deploymentDate: null,
        environment: 'production', // 'development', 'staging', 'production'
        features: [],
        backupEnabled: true,
        monitoringEnabled: true
    },
    
    /**
     * Complete system deployment
     */
    deploySystem: function(options = {}) {
        try {
            console.log('🚀 Starting LuminaHQ system deployment...');
            
            const deploymentId = Utilities.getUuid();
            const startTime = new Date();
            
            // Update configuration
            this.config = { ...this.config, ...options };
            this.config.deploymentDate = startTime;
            
            const deploymentSteps = [
                { name: 'System Validation', fn: () => this.validateSystem() },
                { name: 'Sheet Initialization', fn: () => this.initializeAllSheets() },
                { name: 'Security Setup', fn: () => this.setupSecuritySystems() },
                { name: 'Performance Optimization', fn: () => this.setupPerformanceOptimization() },
                { name: 'Monitoring Setup', fn: () => this.setupMonitoring() },
                { name: 'User Systems', fn: () => this.setupUserSystems() },
                { name: 'Admin Tools', fn: () => this.setupAdminTools() },
                { name: 'Backup Configuration', fn: () => this.setupBackupSystem() },
                { name: 'Final Validation', fn: () => this.validateDeployment() }
            ];
            
            const results = [];
            let successCount = 0;
            
            for (const step of deploymentSteps) {
                try {
                    console.log(`📋 Executing: ${step.name}...`);
                    const stepResult = step.fn();
                    results.push({
                        step: step.name,
                        status: 'SUCCESS',
                        result: stepResult,
                        timestamp: new Date()
                    });
                    successCount++;
                    console.log(`✅ ${step.name} completed successfully`);
                } catch (error) {
                    console.error(`❌ ${step.name} failed:`, error);
                    results.push({
                        step: step.name,
                        status: 'FAILED',
                        error: error.message,
                        timestamp: new Date()
                    });
                }
            }
            
            const endTime = new Date();
            const duration = (endTime - startTime) / 1000;
            
            const deploymentResult = {
                deploymentId: deploymentId,
                version: this.config.version,
                startTime: startTime,
                endTime: endTime,
                duration: duration,
                totalSteps: deploymentSteps.length,
                successCount: successCount,
                failureCount: deploymentSteps.length - successCount,
                success: successCount === deploymentSteps.length,
                steps: results,
                config: this.config
            };
            
            // Log deployment
            this.logDeployment(deploymentResult);
            
            // Setup post-deployment monitoring
            if (deploymentResult.success) {
                this.setupPostDeploymentMonitoring();
                console.log('🎉 LuminaHQ deployment completed successfully!');
            } else {
                console.warn('⚠️ LuminaHQ deployment completed with errors');
            }
            
            return deploymentResult;
            
        } catch (error) {
            console.error('💥 Deployment failed catastrophically:', error);
            return {
                success: false,
                error: error.message,
                timestamp: new Date()
            };
        }
    },
    
    /**
     * Validate system requirements
     */
    validateSystem: function() {
        const validations = [];
        
        // Check Google Apps Script environment
        try {
            SpreadsheetApp.getActiveSpreadsheet();
            validations.push({ check: 'Spreadsheet Access', status: 'PASS' });
        } catch (error) {
            validations.push({ check: 'Spreadsheet Access', status: 'FAIL', error: error.message });
        }
        
        // Check URL Fetch capability
        try {
            UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
            validations.push({ check: 'URL Fetch', status: 'PASS' });
        } catch (error) {
            validations.push({ check: 'URL Fetch', status: 'FAIL', error: error.message });
        }
        
        // Check Properties Service
        try {
            PropertiesService.getScriptProperties().getProperty('test') || '';
            validations.push({ check: 'Properties Service', status: 'PASS' });
        } catch (error) {
            validations.push({ check: 'Properties Service', status: 'FAIL', error: error.message });
        }
        
        // Check Cache Service
        try {
            CacheService.getScriptCache().get('test') || '';
            validations.push({ check: 'Cache Service', status: 'PASS' });
        } catch (error) {
            validations.push({ check: 'Cache Service', status: 'FAIL', error: error.message });
        }
        
        const failedChecks = validations.filter(v => v.status === 'FAIL');
        if (failedChecks.length > 0) {
            throw new Error(`System validation failed: ${failedChecks.map(f => f.check).join(', ')}`);
        }
        
        return { validations: validations, allPassed: true };
    },
    
    /**
     * Initialize all required sheets
     */
    initializeAllSheets: function() {
        const sheets = [
            { name: 'Users', headers: ['ID', 'Email', 'FullName', 'Password', 'Roles', 'CampaignID', 'EmailConfirmed', 'CreatedAt'] },
            { name: 'UserBookmarks', headers: ['ID', 'UserID', 'Title', 'URL', 'Description', 'Tags', 'Folder', 'Created', 'LastAccessed', 'AccessCount'] },
            { name: 'BrowsingAnalytics', headers: ['Timestamp', 'UserID', 'UserEmail', 'URL', 'Domain', 'Action', 'UserAgent', 'Duration', 'Success', 'ErrorReason'] },
            { name: 'SecurityIncidents', headers: ['Timestamp', 'UserID', 'UserEmail', 'URL', 'IncidentType', 'Details', 'Severity', 'Resolved'] },
            { name: 'ComplianceAuditTrail', headers: ['Timestamp', 'EventID', 'UserID', 'UserEmail', 'EventType', 'Classification', 'Details', 'RetentionDate'] },
            { name: 'SystemHealth', headers: ['Timestamp', 'Component', 'Status', 'Metrics', 'Alerts', 'Notes'] },
            { name: 'DeploymentLog', headers: ['DeploymentID', 'Version', 'Timestamp', 'Duration', 'Success', 'Steps', 'Config'] },
            { name: 'PerformanceMetrics', headers: ['Timestamp', 'RequestCount', 'CacheHitRate', 'AvgResponseTime', 'ErrorRate', 'UserCount'] }
        ];
        
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const createdSheets = [];
        
        sheets.forEach(sheetConfig => {
            try {
                let sheet = ss.getSheetByName(sheetConfig.name);
                if (!sheet) {
                    sheet = ss.insertSheet(sheetConfig.name);
                    sheet.getRange(1, 1, 1, sheetConfig.headers.length).setValues([sheetConfig.headers]);
                    sheet.setFrozenRows(1);
                    createdSheets.push(sheetConfig.name);
                }
            } catch (error) {
                console.warn(`Failed to create sheet ${sheetConfig.name}:`, error);
            }
        });
        
        return { totalSheets: sheets.length, createdSheets: createdSheets };
    },
    
    /**
     * Setup security systems
     */
    setupSecuritySystems: function() {
        // Initialize content filter
        const defaultBlockedDomains = [
            'malware-test.com',
            'phishing-example.com'
        ];
        
        // Initialize threat detection
        const defaultThreatRules = {
            maxRequestsPerMinute: 30,
            maxFailedAttempts: 5,
            suspiciousKeywords: ['hack', 'crack', 'exploit', 'malware', 'phishing']
        };
        
        // Save security configuration
        PropertiesService.getScriptProperties().setProperties({
            'contentFilter.blockedDomains': JSON.stringify(defaultBlockedDomains),
            'threatDetection.rules': JSON.stringify(defaultThreatRules),
            'security.initialized': new Date().toISOString()
        });
        
        return { blockedDomains: defaultBlockedDomains.length, threatRules: Object.keys(defaultThreatRules).length };
    },
    
    /**
     * Setup performance optimization
     */
    setupPerformanceOptimization: function() {
        const performanceConfig = {
            maxCacheSize: 50,
            cacheExpiry: 3600000, // 1 hour
            enableCompression: true,
            enableCaching: true
        };
        
        PropertiesService.getScriptProperties().setProperty(
            'performance.config',
            JSON.stringify(performanceConfig)
        );
        
        // Initialize performance metrics
        this.logPerformanceMetrics({
            requestCount: 0,
            cacheHitRate: 0,
            avgResponseTime: 0,
            errorRate: 0,
            userCount: 0
        });
        
        return performanceConfig;
    },
    
    /**
     * Setup monitoring systems
     */
    setupMonitoring: function() {
        const monitoringConfig = {
            enabled: true,
            checkInterval: 300000, // 5 minutes
            healthChecks: ['proxy', 'security', 'performance', 'storage'],
            alertThresholds: {
                errorRate: 5, // 5%
                responseTime: 5000, // 5 seconds
                cacheHitRate: 20 // 20%
            }
        };
        
        PropertiesService.getScriptProperties().setProperty(
            'monitoring.config',
            JSON.stringify(monitoringConfig)
        );
        
        // Log initial system health
        this.logSystemHealth('DEPLOYMENT', 'HEALTHY', {}, []);
        
        return monitoringConfig;
    },
    
    /**
     * Setup user systems
     */
    setupUserSystems: function() {
        // Initialize default bookmark folders
        const defaultFolders = ['General', 'Work', 'Reference', 'Tools'];
        
        PropertiesService.getScriptProperties().setProperty(
            'bookmarks.defaultFolders',
            JSON.stringify(defaultFolders)
        );
        
        return { defaultFolders: defaultFolders };
    },
    
    /**
     * Setup admin tools
     */
    setupAdminTools: function() {
        const adminConfig = {
            enableAnalytics: true,
            enableReporting: true,
            enableUserManagement: true,
            enableSystemMonitoring: true
        };
        
        PropertiesService.getScriptProperties().setProperty(
            'admin.config',
            JSON.stringify(adminConfig)
        );
        
        return adminConfig;
    },
    
    /**
     * Setup backup system
     */
    setupBackupSystem: function() {
        if (!this.config.backupEnabled) {
            return { enabled: false };
        }
        
        const backupConfig = {
            enabled: true,
            frequency: 'daily',
            retentionDays: 30,
            includeUserData: true,
            includeSystemData: true
        };
        
        PropertiesService.getScriptProperties().setProperty(
            'backup.config',
            JSON.stringify(backupConfig)
        );
        
        // Create initial backup
        this.createSystemBackup();
        
        return backupConfig;
    },
    
    /**
     * Validate deployment
     */
    validateDeployment: function() {
        const validations = [];
        
        // Check if all required sheets exist
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const requiredSheets = ['Users', 'UserBookmarks', 'BrowsingAnalytics', 'SecurityIncidents'];
        
        requiredSheets.forEach(sheetName => {
            const sheet = ss.getSheetByName(sheetName);
            validations.push({
                check: `Sheet: ${sheetName}`,
                status: sheet ? 'PASS' : 'FAIL'
            });
        });
        
        // Check if configuration is saved
        const properties = PropertiesService.getScriptProperties();
        const requiredProperties = ['contentFilter.blockedDomains', 'performance.config', 'monitoring.config'];
        
        requiredProperties.forEach(prop => {
            const value = properties.getProperty(prop);
            validations.push({
                check: `Property: ${prop}`,
                status: value ? 'PASS' : 'FAIL'
            });
        });
        
        const failedValidations = validations.filter(v => v.status === 'FAIL');
        if (failedValidations.length > 0) {
            throw new Error(`Deployment validation failed: ${failedValidations.map(f => f.check).join(', ')}`);
        }
        
        return { validations: validations, allPassed: true };
    },
    
    /**
     * Log deployment
     */
    logDeployment: function(deploymentResult) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let deploymentSheet = ss.getSheetByName('DeploymentLog');
            
            if (!deploymentSheet) {
                deploymentSheet = ss.insertSheet('DeploymentLog');
                const headers = ['DeploymentID', 'Version', 'Timestamp', 'Duration', 'Success', 'Steps', 'Config'];
                deploymentSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            }
            
            const rowData = [
                deploymentResult.deploymentId,
                deploymentResult.version,
                deploymentResult.startTime,
                deploymentResult.duration,
                deploymentResult.success,
                JSON.stringify(deploymentResult.steps),
                JSON.stringify(deploymentResult.config)
            ];
            
            deploymentSheet.appendRow(rowData);
            
        } catch (error) {
            console.error('Error logging deployment:', error);
        }
    },
    
    /**
     * Setup post-deployment monitoring
     */
    setupPostDeploymentMonitoring: function() {
        // Create time-driven trigger for health checks (if not exists)
        const triggers = ScriptApp.getProjectTriggers();
        const existingTrigger = triggers.find(trigger => 
            trigger.getHandlerFunction() === 'performHealthCheck'
        );
        
        if (!existingTrigger) {
            try {
                ScriptApp.newTrigger('performHealthCheck')
                    .timeBased()
                    .everyMinutes(5)
                    .create();
                console.log('Health check trigger created');
            } catch (error) {
                console.warn('Could not create health check trigger:', error);
            }
        }
    },
    
    /**
     * Log system health
     */
    logSystemHealth: function(component, status, metrics, alerts) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let healthSheet = ss.getSheetByName('SystemHealth');
            
            if (healthSheet) {
                const rowData = [
                    new Date(),
                    component,
                    status,
                    JSON.stringify(metrics),
                    JSON.stringify(alerts),
                    `Deployment: ${this.config.version}`
                ];
                
                healthSheet.appendRow(rowData);
            }
            
        } catch (error) {
            console.error('Error logging system health:', error);
        }
    },
    
    /**
     * Log performance metrics
     */
    logPerformanceMetrics: function(metrics) {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let metricsSheet = ss.getSheetByName('PerformanceMetrics');
            
            if (metricsSheet) {
                const rowData = [
                    new Date(),
                    metrics.requestCount || 0,
                    metrics.cacheHitRate || 0,
                    metrics.avgResponseTime || 0,
                    metrics.errorRate || 0,
                    metrics.userCount || 0
                ];
                
                metricsSheet.appendRow(rowData);
            }
            
        } catch (error) {
            console.error('Error logging performance metrics:', error);
        }
    },
    
    /**
     * Create system backup
     */
    createSystemBackup: function() {
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const backupData = {
                timestamp: new Date(),
                version: this.config.version,
                sheets: {},
                properties: {}
            };
            
            // Backup sheet data
            const sheets = ss.getSheets();
            sheets.forEach(sheet => {
                try {
                    const data = sheet.getDataRange().getValues();
                    backupData.sheets[sheet.getName()] = data;
                } catch (error) {
                    console.warn(`Could not backup sheet ${sheet.getName()}:`, error);
                }
            });
            
            // Backup properties
            const properties = PropertiesService.getScriptProperties().getProperties();
            backupData.properties = properties;
            
            // Store backup (in a compressed format)
            const backupKey = `backup_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss')}`;
            PropertiesService.getScriptProperties().setProperty(
                backupKey,
                JSON.stringify(backupData)
            );
            
            console.log(`System backup created: ${backupKey}`);
            return { success: true, backupKey: backupKey };
            
        } catch (error) {
            console.error('Error creating system backup:', error);
            return { success: false, error: error.message };
        }
    }
};

// ───────────────────────────────────────────────────────────────────────────────
// SYSTEM MONITORING AND HEALTH CHECKS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * System monitoring and health check service
 */
const MonitoringService = {
    
    /**
     * Perform comprehensive health check
     */
    performHealthCheck: function() {
        try {
            const healthResults = {
                timestamp: new Date(),
                overall: 'HEALTHY',
                components: {},
                alerts: [],
                metrics: {}
            };
            
            // Check proxy system
            healthResults.components.proxy = this.checkProxyHealth();
            
            // Check security systems
            healthResults.components.security = this.checkSecurityHealth();
            
            // Check performance
            healthResults.components.performance = this.checkPerformanceHealth();
            
            // Check storage
            healthResults.components.storage = this.checkStorageHealth();
            
            // Determine overall health
            const componentStatuses = Object.values(healthResults.components).map(c => c.status);
            const hasErrors = componentStatuses.includes('ERROR');
            const hasWarnings = componentStatuses.includes('WARNING');
            
            if (hasErrors) {
                healthResults.overall = 'UNHEALTHY';
            } else if (hasWarnings) {
                healthResults.overall = 'DEGRADED';
            }
            
            // Collect alerts
            Object.values(healthResults.components).forEach(component => {
                if (component.alerts) {
                    healthResults.alerts.push(...component.alerts);
                }
            });
            
            // Log health status
            DeploymentService.logSystemHealth(
                'SYSTEM',
                healthResults.overall,
                healthResults.metrics,
                healthResults.alerts
            );
            
            // Handle alerts
            if (healthResults.alerts.length > 0) {
                this.handleAlerts(healthResults.alerts);
            }
            
            return healthResults;
            
        } catch (error) {
            console.error('Error performing health check:', error);
            return {
                timestamp: new Date(),
                overall: 'ERROR',
                error: error.message
            };
        }
    },
    
    /**
     * Check proxy system health
     */
    checkProxyHealth: function() {
        try {
            // Test proxy functionality
            const testUrl = 'https://www.google.com';
            const startTime = Date.now();
            
            const response = UrlFetchApp.fetch(testUrl, {
                muteHttpExceptions: true,
                timeout: 10000
            });
            
            const responseTime = Date.now() - startTime;
            const statusCode = response.getResponseCode();
            
            const result = {
                status: statusCode === 200 ? 'HEALTHY' : 'WARNING',
                responseTime: responseTime,
                statusCode: statusCode,
                alerts: []
            };
            
            if (responseTime > 5000) {
                result.alerts.push({
                    type: 'SLOW_RESPONSE',
                    message: `Proxy response time ${responseTime}ms exceeds threshold`,
                    severity: 'WARNING'
                });
            }
            
            if (statusCode !== 200) {
                result.status = 'ERROR';
                result.alerts.push({
                    type: 'PROXY_ERROR',
                    message: `Proxy returned status code ${statusCode}`,
                    severity: 'ERROR'
                });
            }
            
            return result;
            
        } catch (error) {
            return {
                status: 'ERROR',
                error: error.message,
                alerts: [{
                    type: 'PROXY_FAILURE',
                    message: `Proxy system failure: ${error.message}`,
                    severity: 'ERROR'
                }]
            };
        }
    },
    
    /**
     * Check security system health
     */
    checkSecurityHealth: function() {
        try {
            const result = {
                status: 'HEALTHY',
                alerts: [],
                checks: {}
            };
            
            // Check if security configuration exists
            const properties = PropertiesService.getScriptProperties();
            const securityConfig = properties.getProperty('contentFilter.blockedDomains');
            
            if (!securityConfig) {
                result.status = 'WARNING';
                result.alerts.push({
                    type: 'MISSING_SECURITY_CONFIG',
                    message: 'Security configuration not found',
                    severity: 'WARNING'
                });
            } else {
                result.checks.securityConfig = 'PASS';
            }
            
            // Check recent security incidents
            try {
                const ss = SpreadsheetApp.getActiveSpreadsheet();
                const securitySheet = ss.getSheetByName('SecurityIncidents');
                
                if (securitySheet) {
                    const data = securitySheet.getDataRange().getValues();
                    const recentIncidents = data.filter(row => {
                        const timestamp = new Date(row[0]);
                        const hourAgo = new Date(Date.now() - 60 * 60 * 1000);
                        return timestamp > hourAgo;
                    });
                    
                    if (recentIncidents.length > 10) {
                        result.status = 'WARNING';
                        result.alerts.push({
                            type: 'HIGH_SECURITY_INCIDENTS',
                            message: `${recentIncidents.length} security incidents in last hour`,
                            severity: 'WARNING'
                        });
                    }
                    
                    result.checks.recentIncidents = recentIncidents.length;
                }
            } catch (error) {
                result.checks.incidentCheck = 'FAILED';
            }
            
            return result;
            
        } catch (error) {
            return {
                status: 'ERROR',
                error: error.message,
                alerts: [{
                    type: 'SECURITY_CHECK_FAILURE',
                    message: `Security health check failed: ${error.message}`,
                    severity: 'ERROR'
                }]
            };
        }
    },
    
    /**
     * Check performance health
     */
    checkPerformanceHealth: function() {
        try {
            const result = {
                status: 'HEALTHY',
                alerts: [],
                metrics: {}
            };
            
            // Get performance metrics from cache or properties
            if (typeof PerformanceOptimizationService !== 'undefined') {
                const metrics = PerformanceOptimizationService.getMetrics();
                result.metrics = metrics;
                
                // Check cache hit rate
                if (metrics.cacheHitRate < 20) {
                    result.status = 'WARNING';
                    result.alerts.push({
                        type: 'LOW_CACHE_HIT_RATE',
                        message: `Cache hit rate ${metrics.cacheHitRate}% below threshold`,
                        severity: 'WARNING'
                    });
                }
                
                // Check average response time
                if (metrics.averageResponseTime > 5000) {
                    result.status = 'WARNING';
                    result.alerts.push({
                        type: 'HIGH_RESPONSE_TIME',
                        message: `Average response time ${metrics.averageResponseTime}ms above threshold`,
                        severity: 'WARNING'
                    });
                }
            }
            
            return result;
            
        } catch (error) {
            return {
                status: 'ERROR',
                error: error.message,
                alerts: [{
                    type: 'PERFORMANCE_CHECK_FAILURE',
                    message: `Performance health check failed: ${error.message}`,
                    severity: 'ERROR'
                }]
            };
        }
    },
    
    /**
     * Check storage health
     */
    checkStorageHealth: function() {
        try {
            const result = {
                status: 'HEALTHY',
                alerts: [],
                metrics: {}
            };
            
            // Check spreadsheet access
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            const sheets = ss.getSheets();
            result.metrics.totalSheets = sheets.length;
            
            // Check for required sheets
            const requiredSheets = ['Users', 'BrowsingAnalytics', 'SecurityIncidents'];
            const missingSheets = requiredSheets.filter(name => !ss.getSheetByName(name));
            
            if (missingSheets.length > 0) {
                result.status = 'ERROR';
                result.alerts.push({
                    type: 'MISSING_SHEETS',
                    message: `Missing required sheets: ${missingSheets.join(', ')}`,
                    severity: 'ERROR'
                });
            }
            
            // Check properties service
            try {
                PropertiesService.getScriptProperties().getProperty('test');
                result.metrics.propertiesAccess = 'OK';
            } catch (error) {
                result.status = 'WARNING';
                result.alerts.push({
                    type: 'PROPERTIES_ACCESS_ISSUE',
                    message: 'Properties service access issue',
                    severity: 'WARNING'
                });
            }
            
            return result;
            
        } catch (error) {
            return {
                status: 'ERROR',
                error: error.message,
                alerts: [{
                    type: 'STORAGE_CHECK_FAILURE',
                    message: `Storage health check failed: ${error.message}`,
                    severity: 'ERROR'
                }]
            };
        }
    },
    
    /**
     * Handle alerts
     */
    handleAlerts: function(alerts) {
        try {
            const criticalAlerts = alerts.filter(alert => alert.severity === 'ERROR');
            
            if (criticalAlerts.length > 0) {
                console.error('🚨 CRITICAL ALERTS:', criticalAlerts);
                
                // In production, send notifications
                // this.sendAlertNotifications(criticalAlerts);
            }
            
            const warningAlerts = alerts.filter(alert => alert.severity === 'WARNING');
            if (warningAlerts.length > 0) {
                console.warn('⚠️ WARNING ALERTS:', warningAlerts);
            }
            
        } catch (error) {
            console.error('Error handling alerts:', error);
        }
    },
    
    /**
     * Send alert notifications (placeholder)
     */
    sendAlertNotifications: function(alerts) {
        // Implement email/SMS notifications for critical alerts
        // Example:
        // MailApp.sendEmail({
        //     to: 'admin@yourcompany.com',
        //     subject: 'LuminaHQ Critical Alert',
        //     body: `Critical system alerts detected:\n${JSON.stringify(alerts, null, 2)}`
        // });
    }
};

// ───────────────────────────────────────────────────────────────────────────────
// CLIENT-ACCESSIBLE DEPLOYMENT AND MONITORING FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Deploy system (client-accessible)
 */
function clientDeploySystem(options) {
    try {
        return DeploymentService.deploySystem(options);
    } catch (error) {
        console.error('Error in client deploy system:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Get system health (client-accessible)
 */
function clientGetSystemHealth() {
    try {
        return MonitoringService.performHealthCheck();
    } catch (error) {
        console.error('Error getting system health:', error);
        return { overall: 'ERROR', error: error.message };
    }
}

/**
 * Create system backup (client-accessible)
 */
function clientCreateSystemBackup() {
    try {
        return DeploymentService.createSystemBackup();
    } catch (error) {
        console.error('Error creating system backup:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Get deployment status (client-accessible)
 */
function clientGetDeploymentStatus() {
    try {
        const properties = PropertiesService.getScriptProperties();
        const version = DeploymentService.config.version;
        const deploymentDate = properties.getProperty('deployment.date');
        
        return {
            version: version,
            deploymentDate: deploymentDate,
            environment: DeploymentService.config.environment,
            features: DeploymentService.config.features
        };
    } catch (error) {
        console.error('Error getting deployment status:', error);
        return { version: 'unknown', error: error.message };
    }
}

// ───────────────────────────────────────────────────────────────────────────────
// TRIGGER FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Perform health check (called by trigger)
 */
function performHealthCheck() {
    try {
        const healthResult = MonitoringService.performHealthCheck();
        
        // Log performance metrics if available
        if (typeof PerformanceOptimizationService !== 'undefined') {
            const metrics = PerformanceOptimizationService.getMetrics();
            DeploymentService.logPerformanceMetrics(metrics);
        }
        
        return healthResult;
    } catch (error) {
        console.error('Scheduled health check failed:', error);
        return { overall: 'ERROR', error: error.message };
    }
}

/**
 * Daily maintenance (called by trigger)
 */
function dailyMaintenanceWithDeployment() {
    try {
        console.log('🔧 Running daily maintenance with deployment features...');
        
        // Run standard daily maintenance
        if (typeof dailyMaintenance === 'function') {
            dailyMaintenance();
        }
        
        // Perform comprehensive health check
        const healthResult = MonitoringService.performHealthCheck();
        
        // Create system backup
        if (DeploymentService.config.backupEnabled) {
            DeploymentService.createSystemBackup();
        }
        
        // Clean old logs (keep 30 days)
        const thirtyDaysAgo = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000));
        cleanOldLogs(thirtyDaysAgo);
        
        console.log('✅ Daily maintenance with deployment features completed');
        return { success: true, health: healthResult };
        
    } catch (error) {
        console.error('❌ Daily maintenance failed:', error);
        writeError('dailyMaintenanceWithDeployment', error);
        return { success: false, error: error.message };
    }
}

/**
 * Clean old logs
 */
function cleanOldLogs(cutoffDate) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheetsToClean = ['SystemHealth', 'PerformanceMetrics', 'BrowsingAnalytics'];
        let totalDeleted = 0;
        
        sheetsToClean.forEach(sheetName => {
            try {
                const sheet = ss.getSheetByName(sheetName);
                if (!sheet) return;
                
                const data = sheet.getDataRange().getValues();
                if (data.length <= 1) return; // Only header row
                
                const headers = data[0];
                const rows = data.slice(1);
                
                // Filter rows newer than cutoff date
                const validRows = rows.filter(row => {
                    const timestamp = new Date(row[0]);
                    return timestamp > cutoffDate;
                });
                
                totalDeleted += (rows.length - validRows.length);
                
                // Rewrite sheet with valid rows
                sheet.clear();
                const newData = [headers, ...validRows];
                if (newData.length > 0) {
                    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
                }
                
            } catch (error) {
                console.warn(`Error cleaning sheet ${sheetName}:`, error);
            }
        });
        
        console.log(`🗑️ Cleaned ${totalDeleted} old log entries`);
        return totalDeleted;
        
    } catch (error) {
        console.error('Error cleaning old logs:', error);
        return 0;
    }
}

// ───────────────────────────────────────────────────────────────────────────────
// QUICK SETUP FUNCTION
// ───────────────────────────────────────────────────────────────────────────────

/**
 * One-click system setup and deployment
 */
function setupLuminaHQComplete() {
    try {
        console.log('🚀 Starting complete LuminaHQ setup...');
        
        // Deploy the system
        const deploymentResult = DeploymentService.deploySystem({
            environment: 'production',
            features: [
                'Enhanced Proxy',
                'Security & Compliance', 
                'Performance Optimization',
                'Mobile Support',
                'Analytics & Monitoring',
                'Admin Dashboard'
            ],
            backupEnabled: true,
            monitoringEnabled: true
        });
        
        if (deploymentResult.success) {
            console.log('🎉 LuminaHQ setup completed successfully!');
            console.log('📊 Setup Summary:');
            console.log(`   Version: ${deploymentResult.version}`);
            console.log(`   Duration: ${deploymentResult.duration} seconds`);
            console.log(`   Steps: ${deploymentResult.successCount}/${deploymentResult.totalSteps} successful`);
            console.log(`   Environment: ${deploymentResult.config.environment}`);
            
            return {
                success: true,
                message: 'LuminaHQ setup completed successfully',
                deployment: deploymentResult,
                nextSteps: [
                    'Configure user authentication',
                    'Set up admin users',
                    'Configure content filtering rules',
                    'Test proxy functionality',
                    'Deploy to users'
                ]
            };
        } else {
            console.error('❌ LuminaHQ setup failed');
            return {
                success: false,
                message: 'Setup failed',
                deployment: deploymentResult
            };
        }
        
    } catch (error) {
        console.error('💥 Complete setup failed:', error);
        return {
            success: false,
            error: error.message
        };
    }
}