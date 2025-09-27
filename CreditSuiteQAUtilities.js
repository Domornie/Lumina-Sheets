/**
 * COMPLETE CreditSuiteQAUtilities.gs - Campaign-specific QA utilities for Credit Suite
 * This includes ALL required functions for the Credit Suite QA system
 */

// ────────────────────────────────────────────────────────────────────────────
// REQUIRED CONFIGURATION CONSTANTS FOR CREDIT SUITE
// ────────────────────────────────────────────────────────────────────────────

// IMPORTANT: Replace these with your actual Google Sheets and Drive IDs
const CREDIT_SUITE_SHEET_ID = "1rw-K55uVkr6Pm3iPIM0zGUyOyyxrTs7EO54sB1ekKFo"; // Use same sheet as Independence for demo

// Drive folder configuration for audio uploads
const CREDIT_SUITE_DRIVE_FOLDER_ID = "1HUINOGK4TO9xkAjAv83A_no7yIoJywkq"; // Replace with your Drive folder ID

// Sheet names within your spreadsheet
const CREDIT_SUITE_QA_SHEET = "CreditSuiteQA";
const CREDIT_SUITE_QA_ANALYTICS_SHEET = "CreditSuiteQAAnalytics";
const CREDIT_SUITE_QA_CATEGORIES_SHEET = "CreditSuiteQACategories";
const CREDIT_SUITE_HEALTHCHECK_SHEET = "CreditSuiteHealthCheck";

// Company branding
const CREDIT_SUITE_COMPANY_LOGO = "https://via.placeholder.com/256x256/1e40af/ffffff?text=Credit+Suite";

// Complete Credit Suite QA Headers - ALL REQUIRED COLUMNS
const CREDIT_SUITE_QA_HEADERS = [
    // Basic Information
    "ID", "Timestamp", "ClientName", "AgentName", "AgentEmail", "ConsultationDate", "AuditorName", "AuditDate", "ConsultationType", "AudioURL",
    
    // Section 1: Initial Assessment & Client Intake (16 Points)
    "Q1_ProfessionalIntroduction", "Q1_ProfessionalIntroduction_Comments", "Q1_ProfessionalIntroduction_MaxPoints",
    "Q2_ClientGoalsAssessment", "Q2_ClientGoalsAssessment_Comments", "Q2_ClientGoalsAssessment_MaxPoints", 
    "Q3_CreditHistoryReview", "Q3_CreditHistoryReview_Comments", "Q3_CreditHistoryReview_MaxPoints",
    "Q4_DocumentationCollection", "Q4_DocumentationCollection_Comments", "Q4_DocumentationCollection_MaxPoints",
    "Q5_PrivacyDisclosure", "Q5_PrivacyDisclosure_Comments", "Q5_PrivacyDisclosure_MaxPoints",
    
    // Section 2: Credit Analysis & Strategy Development (18 Points)
    "Q6_AccurateCreditAnalysis", "Q6_AccurateCreditAnalysis_Comments", "Q6_AccurateCreditAnalysis_MaxPoints",
    "Q7_StrategicPlanDevelopment", "Q7_StrategicPlanDevelopment_Comments", "Q7_StrategicPlanDevelopment_MaxPoints",
    "Q8_PriorityIdentification", "Q8_PriorityIdentification_Comments", "Q8_PriorityIdentification_MaxPoints",
    "Q9_DisputeStrategyExplanation", "Q9_DisputeStrategyExplanation_Comments", "Q9_DisputeStrategyExplanation_MaxPoints",
    "Q10_CreditScoreEducation", "Q10_CreditScoreEducation_Comments", "Q10_CreditScoreEducation_MaxPoints",
    
    // Section 3: Compliance & Legal Requirements (12 Points) 
    "Q11_FCRACompliance", "Q11_FCRACompliance_Comments", "Q11_FCRACompliance_MaxPoints",
    "Q12_CRLComplianceEducation", "Q12_CRLComplianceEducation_Comments", "Q12_CRLComplianceEducation_MaxPoints",
    "Q13_TruthfulRepresentation", "Q13_TruthfulRepresentation_Comments", "Q13_TruthfulRepresentation_MaxPoints",
    "Q14_ProperDisclosures", "Q14_ProperDisclosures_Comments", "Q14_ProperDisclosures_MaxPoints",
    
    // Section 4: Client Education & Communication (14 Points)
    "Q15_ClearCommunication", "Q15_ClearCommunication_Comments", "Q15_ClearCommunication_MaxPoints",
    "Q16_EducationalContent", "Q16_EducationalContent_Comments", "Q16_EducationalContent_MaxPoints",
    "Q17_QuestionHandling", "Q17_QuestionHandling_Comments", "Q17_QuestionHandling_MaxPoints",
    "Q18_ActiveListening", "Q18_ActiveListening_Comments", "Q18_ActiveListening_MaxPoints",
    "Q19_ExpectationSetting", "Q19_ExpectationSetting_Comments", "Q19_ExpectationSetting_MaxPoints",
    
    // Section 5: Professional Standards & Soft Skills (8 Points)
    "Q20_ProfessionalDemeanor", "Q20_ProfessionalDemeanor_Comments", "Q20_ProfessionalDemeanor_MaxPoints",
    "Q21_KnowledgeExpertise", "Q21_KnowledgeExpertise_Comments", "Q21_KnowledgeExpertise_MaxPoints",
    "Q22_TimeManagement", "Q22_TimeManagement_Comments", "Q22_TimeManagement_MaxPoints",
    "Q23_ClientRapport", "Q23_ClientRapport_Comments", "Q23_ClientRapport_MaxPoints",
    
    // Overall Assessment
    "TotalPointsEarned", "TotalPointsPossible", "PercentageScore", "PassStatus",
    "OverallFeedback", "OverallFeedbackHtml", "AreasForImprovement", "Strengths", "ActionItems",
    "ComplianceNotes", "FeedbackShared", "AgentAcknowledgment", "CreatedAt", "UpdatedAt"
];

// ────────────────────────────────────────────────────────────────────────────
// Initialize Credit Suite QA System
// ────────────────────────────────────────────────────────────────────────────

/**
 * Initialize Credit Suite QA system - creates required sheets and folders
 */
function initializeCreditSuiteQASystem() {
    try {
        console.log('Initializing Credit Suite QA System...');
        
        // Validate headers first
        if (!CREDIT_SUITE_QA_HEADERS || CREDIT_SUITE_QA_HEADERS.length === 0) {
            throw new Error('CREDIT_SUITE_QA_HEADERS is not properly defined or is empty');
        }
        
        console.log('Headers validated:', CREDIT_SUITE_QA_HEADERS.length, 'columns');
        
        // Check if spreadsheet exists and is accessible
        let spreadsheet;
        try {
            spreadsheet = SpreadsheetApp.openById(CREDIT_SUITE_SHEET_ID);
            console.log('Spreadsheet accessed successfully:', spreadsheet.getName());
        } catch (error) {
            console.error('Error accessing spreadsheet:', error);
            throw new Error('Cannot access Credit Suite QA spreadsheet. Please check CREDIT_SUITE_SHEET_ID: ' + CREDIT_SUITE_SHEET_ID);
        }
        
        // Initialize QA sheet
        initializeCreditSuiteQASheet(spreadsheet);
        
        // Initialize analytics sheet
        initializeCreditSuiteAnalyticsSheet(spreadsheet);
        
        // Initialize Drive folder structure
        initializeCreditSuiteDriveFolders();
        
        console.log('Credit Suite QA System initialized successfully');
        return { success: true, message: 'Credit Suite system initialized' };
        
    } catch (error) {
        console.error('Error initializing Credit Suite QA system:', error);
        safeWriteError('initializeCreditSuiteQASystem', error);
        throw error;
    }
}

/**
 * Initialize main Credit Suite QA sheet with proper headers
 */
function initializeCreditSuiteQASheet(spreadsheet) {
    try {
        console.log('Initializing Credit Suite QA sheet...');
        
        // Validate headers again
        if (!CREDIT_SUITE_QA_HEADERS || CREDIT_SUITE_QA_HEADERS.length === 0) {
            throw new Error('CREDIT_SUITE_QA_HEADERS is not defined or empty');
        }
        
        console.log('Using headers array with', CREDIT_SUITE_QA_HEADERS.length, 'columns');
        
        let qaSheet = spreadsheet.getSheetByName(CREDIT_SUITE_QA_SHEET);
        
        if (!qaSheet) {
            console.log('Creating new Credit Suite QA sheet...');
            qaSheet = spreadsheet.insertSheet(CREDIT_SUITE_QA_SHEET);
        }
        
        // Check if headers are already set properly
        const maxColumns = qaSheet.getMaxColumns();
        console.log('Sheet max columns:', maxColumns);
        
        if (maxColumns > 0) {
            const existingHeaders = qaSheet.getRange(1, 1, 1, Math.min(maxColumns, CREDIT_SUITE_QA_HEADERS.length)).getValues()[0];
            if (existingHeaders.length > 0 && existingHeaders[0] !== '' && existingHeaders[0] === CREDIT_SUITE_QA_HEADERS[0]) {
                console.log('Credit Suite QA sheet headers already exist and match');
                return;
            }
        }
        
        // Ensure sheet has enough columns
        if (maxColumns < CREDIT_SUITE_QA_HEADERS.length) {
            console.log('Expanding sheet columns from', maxColumns, 'to', CREDIT_SUITE_QA_HEADERS.length);
            qaSheet.insertColumns(maxColumns, CREDIT_SUITE_QA_HEADERS.length - maxColumns);
        }
        
        // Set headers with validation
        console.log('Setting Credit Suite QA sheet headers...', CREDIT_SUITE_QA_HEADERS.length, 'columns');
        
        // Create the range safely
        const headerRange = qaSheet.getRange(1, 1, 1, CREDIT_SUITE_QA_HEADERS.length);
        console.log('Created header range: 1,1,1,' + CREDIT_SUITE_QA_HEADERS.length);
        
        // Set the values with Credit Suite branding
        headerRange.setValues([CREDIT_SUITE_QA_HEADERS])
                   .setFontWeight('bold')
                   .setBackground('#1e40af')
                   .setFontColor('white');
        
        // Freeze header row
        qaSheet.setFrozenRows(1);
        
        // Auto-resize columns (with limit to prevent timeout)
        const maxColumnsToResize = Math.min(CREDIT_SUITE_QA_HEADERS.length, 50);
        for (let i = 1; i <= maxColumnsToResize; i++) {
            try {
                qaSheet.autoResizeColumn(i);
            } catch (resizeError) {
                console.warn('Could not resize column', i, ':', resizeError.message);
            }
        }
        
        console.log('Credit Suite QA sheet initialized successfully with', CREDIT_SUITE_QA_HEADERS.length, 'columns');
        
    } catch (error) {
        console.error('Error initializing Credit Suite QA sheet:', error);
        console.error('Headers length:', CREDIT_SUITE_QA_HEADERS ? CREDIT_SUITE_QA_HEADERS.length : 'undefined');
        console.error('Headers sample:', CREDIT_SUITE_QA_HEADERS ? CREDIT_SUITE_QA_HEADERS.slice(0, 5) : 'undefined');
        throw error;
    }
}

/**
 * Initialize Credit Suite analytics sheet
 */
function initializeCreditSuiteAnalyticsSheet(spreadsheet) {
    try {
        console.log('Initializing Credit Suite analytics sheet...');
        
        let analyticsSheet = spreadsheet.getSheetByName(CREDIT_SUITE_QA_ANALYTICS_SHEET);
        
        if (!analyticsSheet) {
            console.log('Creating new Credit Suite analytics sheet...');
            analyticsSheet = spreadsheet.insertSheet(CREDIT_SUITE_QA_ANALYTICS_SHEET);
            
            const analyticsHeaders = [
                'Date', 'Period', 'Granularity', 'Consultant', 'ConsultationType',
                'TotalAssessments', 'AverageScore', 'PassRate', 'ExcellentRate',
                'CategoryScores', 'CriticalFailures', 'ComplianceViolations', 'TrendData', 'CreatedAt'
            ];
            
            // Validate analytics headers
            if (analyticsHeaders.length === 0) {
                throw new Error('Analytics headers array is empty');
            }
            
            console.log('Setting Credit Suite analytics headers:', analyticsHeaders.length, 'columns');
            
            analyticsSheet.getRange(1, 1, 1, analyticsHeaders.length)
                         .setValues([analyticsHeaders])
                         .setFontWeight('bold')
                         .setBackground('#1e40af')
                         .setFontColor('white');
            
            analyticsSheet.setFrozenRows(1);
            console.log('Credit Suite analytics sheet initialized successfully');
        } else {
            console.log('Credit Suite analytics sheet already exists');
        }
        
    } catch (error) {
        console.error('Error initializing Credit Suite analytics sheet:', error);
        // Don't throw error - analytics is not critical for basic functionality
        console.warn('Continuing without Credit Suite analytics sheet...');
    }
}

/**
 * Initialize Credit Suite Drive folder structure
 */
function initializeCreditSuiteDriveFolders() {
    try {
        console.log('Initializing Credit Suite Drive folder structure...');
        
        // Try to access configured folder
        let mainFolder;
        try {
            if (CREDIT_SUITE_DRIVE_FOLDER_ID && 
                CREDIT_SUITE_DRIVE_FOLDER_ID !== "1ABC123..." && 
                CREDIT_SUITE_DRIVE_FOLDER_ID.length > 10) {
                mainFolder = DriveApp.getFolderById(CREDIT_SUITE_DRIVE_FOLDER_ID);
                console.log('Using configured Credit Suite Drive folder:', mainFolder.getName());
            } else {
                throw new Error('No valid folder ID configured');
            }
        } catch (error) {
            console.log('Creating new Credit Suite main folder...');
            mainFolder = DriveApp.createFolder('Credit Suite QA');
            console.log('Created new Credit Suite main folder:', mainFolder.getId());
            console.log('⚠️ IMPORTANT: Update CREDIT_SUITE_DRIVE_FOLDER_ID to:', mainFolder.getId());
        }
        
        // Create subfolders if they don't exist
        const subfolders = ['Consultation Recordings', 'PDF Reports', 'Analytics', 'Compliance Documents'];
        subfolders.forEach(folderName => {
            try {
                const existing = mainFolder.getFoldersByName(folderName);
                if (!existing.hasNext()) {
                    mainFolder.createFolder(folderName);
                    console.log('Created Credit Suite subfolder:', folderName);
                } else {
                    console.log('Credit Suite subfolder already exists:', folderName);
                }
            } catch (subfolderError) {
                console.warn('Could not create Credit Suite subfolder', folderName, ':', subfolderError.message);
            }
        });
        
        console.log('Credit Suite Drive folder structure initialized');
        
    } catch (error) {
        console.error('Error initializing Credit Suite Drive folders:', error);
        console.warn('Continuing without Credit Suite Drive folder setup...');
        // Don't throw error - folder creation is not critical for basic functionality
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Generate Credit Suite QA ID
// ────────────────────────────────────────────────────────────────────────────

/**
 * Generate unique ID for Credit Suite QA assessment
 */
function generateCreditSuiteQAId() {
    const timestamp = new Date().getTime();
    const random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    return `CS_QA_${timestamp}_${random}`;
}

// ────────────────────────────────────────────────────────────────────────────
// Process Audio File From Form Data for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Process audio file from Credit Suite form data object
 */
function processCreditSuiteAudioFileFromFormData(audioFileData, assessmentId) {
    try {
        if (!audioFileData || !audioFileData.bytes) {
            console.log('No audio file data provided for Credit Suite');
            return '';
        }
        
        console.log('Processing Credit Suite audio file:', audioFileData.name);
        
        // Convert the array back to a Blob
        const blob = Utilities.newBlob(
            audioFileData.bytes,
            audioFileData.mimeType,
            audioFileData.name
        );
        
        // Upload to Drive
        const audioUrl = uploadCreditSuiteAudioFile(blob, assessmentId);
        
        console.log('Credit Suite audio file processed successfully:', audioUrl);
        return audioUrl;
        
    } catch (error) {
        console.error('Error processing Credit Suite audio file:', error);
        throw error;
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Upload Audio File to Drive for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Upload Credit Suite audio file to Google Drive
 */
function uploadCreditSuiteAudioFile(audioBlob, assessmentId) {
    try {
        console.log('Uploading Credit Suite audio file for assessment:', assessmentId);
        
        // Get or create audio uploads folder
        const audioFolder = getOrCreateCreditSuiteAudioFolder();
        
        // Generate filename with timestamp
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
        const originalName = audioBlob.getName() || 'consultation.mp3';
        const extension = originalName.split('.').pop() || 'mp3';
        const fileName = `${assessmentId}_${timestamp}.${extension}`;
        
        // Create file in Drive
        const audioFile = audioFolder.createFile(audioBlob.setName(fileName));
        
        // Set sharing permissions (optional - adjust based on your needs)
        audioFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        const audioUrl = audioFile.getUrl();
        console.log('Credit Suite audio file uploaded successfully:', audioUrl);
        
        return audioUrl;
        
    } catch (error) {
        console.error('Error uploading Credit Suite audio file:', error);
        safeWriteError('uploadCreditSuiteAudioFile', error);
        return '';
    }
}

/**
 * Get or create Credit Suite audio uploads folder
 */
function getOrCreateCreditSuiteAudioFolder() {
    try {
        // Try to get main Credit Suite folder first
        let mainFolder;
        try {
            mainFolder = DriveApp.getFolderById(CREDIT_SUITE_DRIVE_FOLDER_ID);
        } catch (e) {
            // If folder ID doesn't exist, create in root
            mainFolder = DriveApp.createFolder('Credit Suite QA');
        }
        
        // Look for audio subfolder
        const audioFolders = mainFolder.getFoldersByName('Consultation Recordings');
        if (audioFolders.hasNext()) {
            return audioFolders.next();
        } else {
            // Create audio subfolder
            return mainFolder.createFolder('Consultation Recordings');
        }
        
    } catch (error) {
        console.error('Error getting/creating Credit Suite audio folder:', error);
        // Fallback: create in root Drive
        return DriveApp.createFolder('Credit Suite QA Audio');
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Update Credit Suite Analytics
// ────────────────────────────────────────────────────────────────────────────

/**
 * Update Credit Suite QA analytics
 */
function updateCreditSuiteQAAnalytics(qaData, scoreResults) {
    try {
        const ss = SpreadsheetApp.openById(CREDIT_SUITE_SHEET_ID);
        let analyticsSheet = ss.getSheetByName(CREDIT_SUITE_QA_ANALYTICS_SHEET);
        
        if (!analyticsSheet) {
            // Create analytics sheet if it doesn't exist
            analyticsSheet = ss.insertSheet(CREDIT_SUITE_QA_ANALYTICS_SHEET);
            
            const headers = [
                'Date', 'Period', 'Granularity', 'Consultant', 'ConsultationType',
                'TotalAssessments', 'AverageScore', 'PassRate', 'ExcellentRate',
                'CategoryScores', 'CriticalFailures', 'ComplianceViolations', 'TrendData', 'CreatedAt'
            ];
            
            analyticsSheet.getRange(1, 1, 1, headers.length)
                         .setValues([headers])
                         .setFontWeight('bold')
                         .setBackground('#1e40af')
                         .setFontColor('white');
        }
        
        // Add analytics entry
        const now = new Date();
        const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        analyticsSheet.appendRow([
            today,
            'Daily',
            'Day',
            qaData.AgentName,
            qaData.ConsultationType,
            1, // This assessment
            scoreResults.overallPercentage,
            scoreResults.hasCriticalFailure ? 0 : (scoreResults.overallPercentage >= 85 ? 100 : 0),
            scoreResults.overallPercentage >= 95 ? 100 : 0,
            JSON.stringify(scoreResults.categoryScores),
            scoreResults.hasCriticalFailure ? scoreResults.criticalFailures.length : 0,
            scoreResults.hasCriticalFailure ? scoreResults.criticalFailures.length : 0,
            JSON.stringify({date: today, score: scoreResults.overallPercentage}),
            now
        ]);
        
        console.log('Credit Suite analytics updated successfully');
        
    } catch (error) {
        console.error('Error updating Credit Suite analytics:', error);
        // Don't throw error - analytics failure shouldn't stop submission
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Save Credit Suite QA to Sheet
// ────────────────────────────────────────────────────────────────────────────

/**
 * Save Credit Suite QA data to sheet
 */
function saveCreditSuiteQAToSheet(qaData) {
    try {
        console.log('Saving Credit Suite QA data to sheet...');
        
        // Validate constants are defined
        if (typeof CREDIT_SUITE_SHEET_ID === 'undefined') {
            throw new Error('CREDIT_SUITE_SHEET_ID is not defined');
        }
        if (typeof CREDIT_SUITE_QA_SHEET === 'undefined') {
            throw new Error('CREDIT_SUITE_QA_SHEET is not defined');
        }
        if (typeof CREDIT_SUITE_QA_HEADERS === 'undefined') {
            throw new Error('CREDIT_SUITE_QA_HEADERS is not defined');
        }
        
        console.log('Using Credit Suite spreadsheet ID:', CREDIT_SUITE_SHEET_ID);
        console.log('Using Credit Suite sheet name:', CREDIT_SUITE_QA_SHEET);
        
        const ss = SpreadsheetApp.openById(CREDIT_SUITE_SHEET_ID);
        let qaSheet = ss.getSheetByName(CREDIT_SUITE_QA_SHEET);
        
        if (!qaSheet) {
            console.log('Credit Suite QA sheet not found, creating it...');
            qaSheet = ss.insertSheet(CREDIT_SUITE_QA_SHEET);
            
            // Set headers if sheet is new
            qaSheet.getRange(1, 1, 1, CREDIT_SUITE_QA_HEADERS.length)
                   .setValues([CREDIT_SUITE_QA_HEADERS])
                   .setFontWeight('bold')
                   .setBackground('#1e40af')
                   .setFontColor('white');
            
            qaSheet.setFrozenRows(1);
        }
        
        // Prepare row data in the correct order
        const rowData = CREDIT_SUITE_QA_HEADERS.map(header => {
            const value = qaData[header];
            // Handle undefined values
            if (value === undefined || value === null) {
                return '';
            }
            // Handle dates
            if (value instanceof Date) {
                return value;
            }
            // Handle everything else
            return value;
        });
        
        console.log('Prepared Credit Suite row data with', rowData.length, 'columns');
        console.log('Sample data:', rowData.slice(0, 10));
        
        // Append the row
        qaSheet.appendRow(rowData);
        
        // Format the new row
        const lastRow = qaSheet.getLastRow();
        formatCreditSuiteQARow(qaSheet, lastRow, qaData);
        
        console.log('Credit Suite QA data saved to sheet successfully at row', lastRow);
        return { success: true, row: lastRow };
        
    } catch (error) {
        console.error('Error saving to Credit Suite QA sheet:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Format Credit Suite QA row in sheet with conditional formatting
 */
function formatCreditSuiteQARow(sheet, rowNumber, qaData) {
    try {
        if (!sheet || !rowNumber || !qaData) {
            console.warn('Invalid parameters for formatting Credit Suite row');
            return;
        }
        
        // Apply conditional formatting based on pass status
        const passStatus = qaData.PassStatus || 'Unknown';
        let backgroundColor;
        
        if (passStatus.includes('Critical Failure') || passStatus.includes('compliance violation')) {
            backgroundColor = '#fef2f2'; // Light red for compliance issues
        } else if (passStatus === 'Excellent Performance') {
            backgroundColor = '#f0fdf4'; // Light green
        } else if (passStatus === 'Meets Standards') {
            backgroundColor = '#fffbeb'; // Light yellow
        } else {
            backgroundColor = '#fafafa'; // Light gray
        }
        
        // Apply background color to the entire row
        const range = sheet.getRange(rowNumber, 1, 1, CREDIT_SUITE_QA_HEADERS.length);
        range.setBackground(backgroundColor);
        
        // Bold the assessment ID and score columns
        const idColumn = CREDIT_SUITE_QA_HEADERS.indexOf('ID') + 1;
        const scoreColumn = CREDIT_SUITE_QA_HEADERS.indexOf('PercentageScore') + 1;
        const statusColumn = CREDIT_SUITE_QA_HEADERS.indexOf('PassStatus') + 1;
        
        if (idColumn > 0) {
            sheet.getRange(rowNumber, idColumn).setFontWeight('bold');
        }
        if (scoreColumn > 0) {
            sheet.getRange(rowNumber, scoreColumn).setFontWeight('bold');
        }
        if (statusColumn > 0) {
            sheet.getRange(rowNumber, statusColumn).setFontWeight('bold');
        }
        
        console.log('Credit Suite row formatting applied successfully');
        
    } catch (error) {
        console.error('Error formatting Credit Suite QA row:', error);
        // Don't throw error - formatting is not critical
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Prepare Credit Suite QA Data
// ────────────────────────────────────────────────────────────────────────────

/**
 * Prepare Credit Suite QA data for sheet storage
 */
function prepareCreditSuiteQAData(formData, scoreResults, assessmentId, audioUrl) {
    try {
        console.log('Preparing Credit Suite QA data for storage...');
        
        const now = new Date();
        
        // Convert HTML feedback to plain text for the main field
        const overallFeedbackHtml = formData.overallFeedbackHtml || formData.overallFeedback || '';
        const overallFeedbackText = stripHtmlTags(overallFeedbackHtml);
        
        const qaData = {
            // Basic Information
            ID: assessmentId,
            Timestamp: now,
            ClientName: formData.clientName || '',
            AgentName: formData.agentName || '',
            AgentEmail: formData.agentEmail || '',
            ConsultationDate: formData.consultationDate || '',
            AuditorName: formData.auditorName || '',
            AuditDate: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            ConsultationType: formData.consultationType || '',
            AudioURL: audioUrl || '',
            
            // Scoring Results
            TotalPointsEarned: scoreResults.totalEarned || 0,
            TotalPointsPossible: scoreResults.totalPossible || 0,
            PercentageScore: scoreResults.overallPercentage || 0,
            PassStatus: scoreResults.passStatus || 'Unknown',
            
            // Overall Assessment
            OverallFeedback: overallFeedbackText, // Plain text for sheet
            OverallFeedbackHtml: overallFeedbackHtml, // HTML for web display
            AreasForImprovement: extractAreasForImprovement(formData),
            Strengths: extractStrengths(formData),
            ActionItems: extractActionItems(formData),
            ComplianceNotes: extractComplianceNotes(formData, scoreResults),
            
            // Metadata
            FeedbackShared: false,
            AgentAcknowledgment: false,
            CreatedAt: now,
            UpdatedAt: now
        };
        
        // Add individual question responses and comments
        if (CREDIT_SUITE_QA_CONFIG && CREDIT_SUITE_QA_CONFIG.categories) {
            Object.values(CREDIT_SUITE_QA_CONFIG.categories).forEach(category => {
                if (category.questions && Array.isArray(category.questions)) {
                    category.questions.forEach(question => {
                        const questionId = question.id;
                        qaData[questionId] = formData[questionId] || 'NA';
                        qaData[questionId + '_Comments'] = formData[questionId + '_Comments'] || '';
                        qaData[questionId + '_MaxPoints'] = question.maxPoints || 0;
                    });
                }
            });
        }
        
        console.log('Credit Suite QA data prepared successfully');
        console.log('Data keys:', Object.keys(qaData).length);
        
        return qaData;
        
    } catch (error) {
        console.error('Error preparing Credit Suite QA data:', error);
        throw error;
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Helper Functions for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Strip HTML tags from text (shared with Independence)
 */
function stripHtmlTags(html) {
    if (!html) return '';
    
    // Remove HTML tags and decode entities
    return html
        .replace(/<[^>]*>/g, '') // Remove HTML tags
        .replace(/&nbsp;/g, ' ') // Replace non-breaking spaces
        .replace(/&amp;/g, '&')  // Decode ampersands
        .replace(/&lt;/g, '<')   // Decode less than
        .replace(/&gt;/g, '>')   // Decode greater than
        .replace(/&quot;/g, '"') // Decode quotes
        .replace(/&#39;/g, "'")  // Decode apostrophes
        .replace(/\s+/g, ' ')    // Normalize whitespace
        .trim();
}

/**
 * Extract areas for improvement from Credit Suite comments
 */
function extractAreasForImprovement(formData) {
    const improvements = [];
    
    if (CREDIT_SUITE_QA_CONFIG && CREDIT_SUITE_QA_CONFIG.categories) {
        Object.values(CREDIT_SUITE_QA_CONFIG.categories).forEach(category => {
            if (category.questions && Array.isArray(category.questions)) {
                category.questions.forEach(question => {
                    const answer = formData[question.id];
                    const comment = formData[question.id + '_Comments'];
                    
                    if (answer === 'No' && comment) {
                        improvements.push(`${question.text}: ${comment}`);
                    }
                });
            }
        });
    }
    
    return improvements.join('; ');
}

/**
 * Extract strengths from Credit Suite comments
 */
function extractStrengths(formData) {
    const strengths = [];
    
    if (CREDIT_SUITE_QA_CONFIG && CREDIT_SUITE_QA_CONFIG.categories) {
        Object.values(CREDIT_SUITE_QA_CONFIG.categories).forEach(category => {
            if (category.questions && Array.isArray(category.questions)) {
                category.questions.forEach(question => {
                    const answer = formData[question.id];
                    const comment = formData[question.id + '_Comments'];
                    
                    if (answer === 'Yes' && comment) {
                        strengths.push(`${question.text}: ${comment}`);
                    }
                });
            }
        });
    }
    
    return strengths.join('; ');
}

/**
 * Extract action items from overall feedback
 */
function extractActionItems(formData) {
    const feedback = stripHtmlTags(formData.overallFeedbackHtml || formData.overallFeedback || '');
    
    // Simple extraction of action-oriented sentences
    const actionWords = ['should', 'must', 'need', 'recommend', 'suggest', 'improve', 'focus', 'ensure', 'comply'];
    const sentences = feedback.split(/[.!?]+/);
    
    const actionItems = sentences.filter(sentence => 
        actionWords.some(word => sentence.toLowerCase().includes(word))
    ).map(item => item.trim()).filter(item => item.length > 10);
    
    return actionItems.join('; ');
}

/**
 * Extract compliance notes from Credit Suite assessment
 */
function extractComplianceNotes(formData, scoreResults) {
    const complianceNotes = [];
    
    // Check for critical failures (compliance violations)
    if (scoreResults.hasCriticalFailure && scoreResults.criticalFailures) {
        scoreResults.criticalFailures.forEach(failure => {
            complianceNotes.push(`CRITICAL: ${failure.questionText}`);
        });
    }
    
    // Check specific compliance-related questions
    const complianceQuestions = ['Q11_FCRACompliance', 'Q12_CRLComplianceEducation', 'Q13_TruthfulRepresentation', 'Q14_ProperDisclosures'];
    
    complianceQuestions.forEach(questionId => {
        const answer = formData[questionId];
        const comment = formData[questionId + '_Comments'];
        
        if (answer === 'No' && comment) {
            complianceNotes.push(`Compliance Issue - ${questionId}: ${comment}`);
        }
    });
    
    return complianceNotes.join('; ');
}

// ────────────────────────────────────────────────────────────────────────────
// Configuration Validation for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Validate that all required Credit Suite constants are defined
 */
function validateCreditSuiteConfiguration() {
    console.log('=== Validating Credit Suite Configuration ===');
    
    const issues = [];
    const status = {
        CREDIT_SUITE_SHEET_ID: false,
        CREDIT_SUITE_QA_SHEET: false,
        CREDIT_SUITE_QA_HEADERS: false,
        CREDIT_SUITE_DRIVE_FOLDER_ID: false
    };
    
    // Check CREDIT_SUITE_SHEET_ID
    if (typeof CREDIT_SUITE_SHEET_ID === 'undefined') {
        issues.push('❌ CREDIT_SUITE_SHEET_ID is not defined');
    } else if (!CREDIT_SUITE_SHEET_ID || CREDIT_SUITE_SHEET_ID.length < 10) {
        issues.push('❌ CREDIT_SUITE_SHEET_ID appears to be invalid');
    } else {
        status.CREDIT_SUITE_SHEET_ID = true;
        console.log('✅ CREDIT_SUITE_SHEET_ID is defined');
    }
    
    // Check CREDIT_SUITE_QA_SHEET
    if (typeof CREDIT_SUITE_QA_SHEET === 'undefined') {
        issues.push('❌ CREDIT_SUITE_QA_SHEET is not defined');
    } else {
        status.CREDIT_SUITE_QA_SHEET = true;
        console.log('✅ CREDIT_SUITE_QA_SHEET is defined:', CREDIT_SUITE_QA_SHEET);
    }
    
    // Check CREDIT_SUITE_QA_HEADERS
    if (typeof CREDIT_SUITE_QA_HEADERS === 'undefined') {
        issues.push('❌ CREDIT_SUITE_QA_HEADERS is not defined');
    } else if (!Array.isArray(CREDIT_SUITE_QA_HEADERS)) {
        issues.push('❌ CREDIT_SUITE_QA_HEADERS is not an array');
    } else if (CREDIT_SUITE_QA_HEADERS.length === 0) {
        issues.push('❌ CREDIT_SUITE_QA_HEADERS array is empty');
    } else {
        status.CREDIT_SUITE_QA_HEADERS = true;
        console.log('✅ CREDIT_SUITE_QA_HEADERS is defined with', CREDIT_SUITE_QA_HEADERS.length, 'columns');
    }
    
    // Check CREDIT_SUITE_DRIVE_FOLDER_ID
    if (typeof CREDIT_SUITE_DRIVE_FOLDER_ID === 'undefined') {
        issues.push('⚠️ CREDIT_SUITE_DRIVE_FOLDER_ID is not defined (will auto-create)');
    } else if (!CREDIT_SUITE_DRIVE_FOLDER_ID || CREDIT_SUITE_DRIVE_FOLDER_ID.length < 10) {
        issues.push('⚠️ CREDIT_SUITE_DRIVE_FOLDER_ID appears to be empty (will auto-create)');
    } else {
        status.CREDIT_SUITE_DRIVE_FOLDER_ID = true;
        console.log('✅ CREDIT_SUITE_DRIVE_FOLDER_ID is defined');
    }
    
    const isValid = status.CREDIT_SUITE_SHEET_ID && 
                   status.CREDIT_SUITE_QA_SHEET && 
                   status.CREDIT_SUITE_QA_HEADERS;
    
    console.log('=== Credit Suite Configuration Status ===');
    console.log('Valid for basic operation:', isValid ? '✅ YES' : '❌ NO');
    
    if (issues.length > 0) {
        console.log('Issues found:');
        issues.forEach(issue => console.log(issue));
    }
    
    return {
        valid: isValid,
        issues: issues,
        status: status
    };
}

// ────────────────────────────────────────────────────────────────────────────
// Testing Functions for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Test function to verify all Credit Suite systems work
 */
function testCreditSuiteQASystem() {
    try {
        console.log('=== Testing Credit Suite QA System ===');
        
        // Check configuration
        const configCheck = validateCreditSuiteConfiguration();
        if (!configCheck.valid) {
            console.error('Credit Suite configuration issues:', configCheck.issues);
        }
        
        // Test system initialization
        initializeCreditSuiteQASystem();
        
        // Test scoring system (if available)
        if (typeof calculateCreditSuiteQAScores === 'function') {
            const testData = {
                Q1_ProfessionalIntroduction: 'Yes',
                Q2_ClientGoalsAssessment: 'Yes',
                Q3_CreditHistoryReview: 'No'
            };
            
            const testScores = calculateCreditSuiteQAScores(testData);
            console.log('Test Credit Suite scoring result:', testScores);
        }
        
        console.log('=== Credit Suite System Test Completed ===');
        return { 
            success: true, 
            message: 'All Credit Suite systems operational',
            configCheck: configCheck
        };
        
    } catch (error) {
        console.error('Credit Suite system test failed:', error);
        return { 
            success: false, 
            error: error.message,
            configCheck: configCheck 
        };
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Safe Error Writing (shared utility)
// ────────────────────────────────────────────────────────────────────────────

/**
 * Safely write errors to a log sheet or console
 */
function safeWriteError(functionName, error) {
    try {
        const errorMessage = `${new Date().toISOString()} - ${functionName}: ${error.message || error}`;
        console.error(errorMessage);
        
        // Try to write to error log sheet if it exists
        try {
            const ss = SpreadsheetApp.openById(CREDIT_SUITE_SHEET_ID);
            let errorSheet = ss.getSheetByName('ErrorLog');
            
            if (!errorSheet) {
                errorSheet = ss.insertSheet('ErrorLog');
                errorSheet.getRange(1, 1, 1, 4)
                         .setValues([['Timestamp', 'Function', 'Error', 'Stack']])
                         .setFontWeight('bold');
            }
            
            errorSheet.appendRow([
                new Date(),
                functionName,
                error.message || error.toString(),
                error.stack || 'No stack trace'
            ]);
            
        } catch (logError) {
            console.error('Could not write to Credit Suite error log:', logError);
        }
        
    } catch (safeError) {
        console.error('Error in safeWriteError:', safeError);
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Submission Function for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Client-accessible function to submit Credit Suite QA with audio
 */
function clientSubmitCreditSuiteQAWithAudio(formData) {
    try {
        console.log('=== Starting Credit Suite QA Submission ===');
        console.log('Form data keys:', Object.keys(formData || {}));
        
        if (!formData || typeof formData !== 'object') {
            throw new Error('Invalid form data provided');
        }
        
        // Generate unique assessment ID
        const assessmentId = generateCreditSuiteQAId();
        console.log('Generated Credit Suite assessment ID:', assessmentId);
        
        // Process audio file if present
        let audioUrl = '';
        if (formData.audioFile) {
            console.log('Processing Credit Suite audio file...');
            audioUrl = processCreditSuiteAudioFileFromFormData(formData.audioFile, assessmentId);
        }
        
        // Calculate final scores
        console.log('Calculating final Credit Suite scores...');
        const scoreResults = formData.finalScores || calculateCreditSuiteQAScores(formData);
        
        // Prepare data for sheet storage
        console.log('Preparing Credit Suite data for storage...');
        const qaData = prepareCreditSuiteQAData(formData, scoreResults, assessmentId, audioUrl);
        
        // Save to sheet
        console.log('Saving Credit Suite data to sheet...');
        const saveResult = saveCreditSuiteQAToSheet(qaData);
        
        if (!saveResult.success) {
            throw new Error('Failed to save to sheet: ' + saveResult.error);
        }
        
        // Update analytics
        console.log('Updating Credit Suite analytics...');
        updateCreditSuiteQAAnalytics(qaData, scoreResults);
        
        console.log('=== Credit Suite QA Submission Completed Successfully ===');
        
        return {
            success: true,
            assessmentId: assessmentId,
            audioUrl: audioUrl,
            scoreResults: scoreResults,
            sheetRow: saveResult.row,
            message: 'Credit Suite QA assessment submitted successfully'
        };
        
    } catch (error) {
        console.error('=== Credit Suite QA Submission Failed ===');
        console.error('Error:', error.message);
        console.error('Stack:', error.stack);
        
        safeWriteError('clientSubmitCreditSuiteQAWithAudio', error);
        
        return {
            success: false,
            error: error.message || 'Unknown error occurred',
            timestamp: new Date().toISOString()
        };
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Auto-validate on load
// ────────────────────────────────────────────────────────────────────────────

console.log('COMPLETE Credit Suite QA Configuration loaded');
console.log('Total headers defined:', CREDIT_SUITE_QA_HEADERS ? CREDIT_SUITE_QA_HEADERS.length : 'undefined');
console.log('Functions available: initializeCreditSuiteQASystem, generateCreditSuiteQAId, processCreditSuiteAudioFileFromFormData, clientSubmitCreditSuiteQAWithAudio');

// Run validation when this script loads
try {
    validateCreditSuiteConfiguration();
} catch (error) {
    console.error('Error during Credit Suite auto-validation:', error);
}