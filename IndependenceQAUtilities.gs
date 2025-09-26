/**
 * COMPLETE IndependenceQAUtilities.gs - Campaign-specific QA utilities for Independence Insurance
 * This includes ALL required functions for the Independence Insurance QA system
 * and constants for the IndependenceCoaching module (sheet names + headers).
 */

// ────────────────────────────────────────────────────────────────────────────
// REQUIRED CONFIG CONSTANTS
// ────────────────────────────────────────────────────────────────────────────

// IMPORTANT: Replace these with your actual Google Sheets and Drive IDs
const INDEPENDENCE_SHEET_ID = "1IWqKgxaHW04Dvt_3SpksSt48SJsg6JUaQsiPQA3nQ68";

// Drive folder configuration for audio uploads
const INDEPENDENCE_DRIVE_FOLDER_ID = "18_2ZHXPp65WNen-e8tEMJUP1BLUvOMHb"; // Replace with your Drive folder ID

// ── Sheet names within your spreadsheet (QA)
const INDEPENDENCE_QA_SHEET = "IndependenceQA";
const INDEPENDENCE_QA_ANALYTICS_SHEET = "IndependenceQAAnalytics";
const INDEPENDENCE_QA_CATEGORIES_SHEET = "IndependenceQACategories";
const INDEPENDENCE_HEALTHCHECK_SHEET = "HealthCheck";

// Company branding
const INDEPENDENCE_COMPANY_LOGO = "https://cdn.brandfetch.io/idX1p_nKjA/w/256/h/256/theme/dark/logo.png?c=1bxid64Mup7aczewSAYMX&t=1748807290801";

// ────────────────────────────────────────────────────────────────────────────
/**
 * INDEPENDENCE — COACHING CONSTANTS (sheet names + headers)
 * Mirrors the Independence QA naming/structure. No functions here; service code
 * can import/use these names directly.
 */
// ────────────────────────────────────────────────────────────────────────────

// Coaching lives in the same spreadsheet as QA (INDEPENDENCE_SHEET_ID)
const INDEPENDENCE_COACHING_SHEET = "IndependenceCoaching";
const INDEPENDENCE_COACHING_ANALYTICS_SHEET = "IndependenceCoachingAnalytics";

// Canonical Coaching headers for "IndependenceCoaching" sheet
const INDEPENDENCE_COACHING_HEADERS = [
  // Identifiers & linkage
  "ID",                 // e.g., IND_COACH_1700000000000_123
  "QAId",               // link to Independence QA row (optional)

  // Core session info
  "SessionDate",        // 'YYYY-MM-DD'
  "AgentName",          // coach name (kept as AgentName for compatibility)
  "CoacheeName",        // the agent/employee being coached
  "CoacheeEmail",       // used for notifications/ack

  // Topics & plan
  "TopicsPlanned",      // JSON or CSV of planned topics
  "CoveredTopics",      // JSON array string of topics actually covered
  "ActionPlan",         // agreed next steps

  // Follow-up & notes
  "FollowUpDate",       // 'YYYY-MM-DD'
  "Notes",              // free-form notes

  // Completion & acknowledgement
  "Completed",          // 'Yes' | '' when session marked complete
  "AcknowledgementText",// HTML/plain text acknowledgement from coachee
  "AcknowledgedOn",     // 'YYYY-MM-DD' or ISO string

  // Audit fields
  "CreatedAt",          // ISO timestamp or Date
  "UpdatedAt"           // ISO timestamp or Date
];

// Optional helper sets (handy for validation/formatting in services)
const INDEPENDENCE_COACHING_DATE_COLUMNS = [
  "SessionDate", "FollowUpDate", "AcknowledgedOn", "CreatedAt", "UpdatedAt"
];
const COACHING_DATE_COLUMNS = INDEPENDENCE_COACHING_DATE_COLUMNS;

const INDEPENDENCE_COACHING_REQUIRED_COLUMNS = [
  "ID", "SessionDate", "AgentName", "CoacheeName", "CoacheeEmail", "CreatedAt", "UpdatedAt"
];
const COACHING_REQUIRED_COLUMNS = INDEPENDENCE_COACHING_REQUIRED_COLUMNS;

const INDEPENDENCE_COACHING_DEFAULTS = {
  Completed: "",
  AcknowledgementText: "",
  AcknowledgedOn: "",
  TopicsPlanned: "",
  CoveredTopics: "[]",
  Notes: "",
  ActionPlan: ""
};
const COACHING_DEFAULTS = INDEPENDENCE_COACHING_DEFAULTS;

// Stable ID prefix for this campaign’s coaching records
const INDEPENDENCE_COACHING_ID_PREFIX = "IND_COACH";
const COACHING_ID_PREFIX = INDEPENDENCE_COACHING_ID_PREFIX;

// ────────────────────────────────────────────────────────────────────────────
// Complete Independence QA Headers - ALL REQUIRED COLUMNS
// ────────────────────────────────────────────────────────────────────────────
const INDEPENDENCE_QA_HEADERS = [
    // Basic Information
    "ID", "Timestamp", "CallerName", "AgentName", "AgentEmail", "CallDate", "AuditorName", "AuditDate", "CallType", "AudioURL",
    
    // Section 1: Call Opening (8 Points)
    "Q1_ProfessionalGreeting", "Q1_ProfessionalGreeting_Comments", "Q1_ProfessionalGreeting_MaxPoints",
    "Q2_ProperIntroduction", "Q2_ProperIntroduction_Comments", "Q2_ProperIntroduction_MaxPoints", 
    "Q3_ToneMatching", "Q3_ToneMatching_Comments", "Q3_ToneMatching_MaxPoints",
    "Q4_ConversationControl", "Q4_ConversationControl_Comments", "Q4_ConversationControl_MaxPoints",
    
    // Section 2: Needs Discovery & Qualification (8 Points)
    "Q5_TruckingCompanyConfirmation", "Q5_TruckingCompanyConfirmation_Comments", "Q5_TruckingCompanyConfirmation_MaxPoints",
    "Q6_BusinessOperationsVerification", "Q6_BusinessOperationsVerification_Comments", "Q6_BusinessOperationsVerification_MaxPoints",
    "Q7_ValueFocusedLanguage", "Q7_ValueFocusedLanguage_Comments", "Q7_ValueFocusedLanguage_MaxPoints",
    "Q8_ProperReinforcement", "Q8_ProperReinforcement_Comments", "Q8_ProperReinforcement_MaxPoints",
    
    // Section 3: Appointment Setting (20 Points)
    "Q9_LiveTransferAttempt", "Q9_LiveTransferAttempt_Comments", "Q9_LiveTransferAttempt_MaxPoints",
    "Q10_AppointmentOffer", "Q10_AppointmentOffer_Comments", "Q10_AppointmentOffer_MaxPoints",
    "Q11_EmailConfirmation", "Q11_EmailConfirmation_Comments", "Q11_EmailConfirmation_MaxPoints",
    "Q12_SchedulingLinkUsage", "Q12_SchedulingLinkUsage_Comments", "Q12_SchedulingLinkUsage_MaxPoints",
    "Q13_UrgencyAndConfidence", "Q13_UrgencyAndConfidence_Comments", "Q13_UrgencyAndConfidence_MaxPoints",
    "Q14_AppointmentConfirmation", "Q14_AppointmentConfirmation_Comments", "Q14_AppointmentConfirmation_MaxPoints",
    
    // Section 4: End of Call Procedure (5 Points)
    "Q15_AppointmentRecap", "Q15_AppointmentRecap_Comments", "Q15_AppointmentRecap_MaxPoints",
    "Q16_EmailSMSExplanation", "Q16_EmailSMSExplanation_Comments", "Q16_EmailSMSExplanation_MaxPoints",
    "Q17_ProfessionalClosing", "Q17_ProfessionalClosing_Comments", "Q17_ProfessionalClosing_MaxPoints",
    
    // Section 5: Soft Skills & Compliance (16 Points)
    "Q18_AttentiveListening", "Q18_AttentiveListening_Comments", "Q18_AttentiveListening_MaxPoints",
    "Q19_ComplianceAccuracy", "Q19_ComplianceAccuracy_Comments", "Q19_ComplianceAccuracy_MaxPoints",
    "Q20_ConversationalDelivery", "Q20_ConversationalDelivery_Comments", "Q20_ConversationalDelivery_MaxPoints",
    "Q21_ClearSpeech", "Q21_ClearSpeech_Comments", "Q21_ClearSpeech_MaxPoints",
    "Q22_QuietEnvironment", "Q22_QuietEnvironment_Comments", "Q22_QuietEnvironment_MaxPoints",
    "Q23_ObjectionHandling", "Q23_ObjectionHandling_Comments", "Q23_ObjectionHandling_MaxPoints",
    "Q24_OverallProfessionalism", "Q24_OverallProfessionalism_Comments", "Q24_OverallProfessionalism_MaxPoints",
    
    // Overall Assessment
    "TotalPointsEarned", "TotalPointsPossible", "PercentageScore", "PassStatus",
    "OverallFeedback", "OverallFeedbackHtml", "AreasForImprovement", "Strengths", "ActionItems",
    "FeedbackShared", "AgentAcknowledgment", "CreatedAt", "UpdatedAt"
];

// ────────────────────────────────────────────────────────────────────────────
// Initialize Independence QA System
// ────────────────────────────────────────────────────────────────────────────

/**
 * Initialize Independence QA system - creates required sheets and folders
 */
function initializeIndependenceQASystem() {
    try {
        console.log('Initializing Independence QA System...');
        
        // Validate headers first
        if (!INDEPENDENCE_QA_HEADERS || INDEPENDENCE_QA_HEADERS.length === 0) {
            throw new Error('INDEPENDENCE_QA_HEADERS is not properly defined or is empty');
        }
        
        console.log('Headers validated:', INDEPENDENCE_QA_HEADERS.length, 'columns');
        
        // Check if spreadsheet exists and is accessible
        let spreadsheet;
        try {
            spreadsheet = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
            console.log('Spreadsheet accessed successfully:', spreadsheet.getName());
        } catch (error) {
            console.error('Error accessing spreadsheet:', error);
            throw new Error('Cannot access Independence QA spreadsheet. Please check INDEPENDENCE_SHEET_ID: ' + INDEPENDENCE_SHEET_ID);
        }
        
        // Initialize QA sheet
        initializeQASheet(spreadsheet);
        
        // Initialize analytics sheet
        initializeAnalyticsSheet(spreadsheet);
        
        // Initialize Drive folder structure
        initializeDriveFolders();
        
        console.log('Independence QA System initialized successfully');
        return { success: true, message: 'System initialized' };
        
    } catch (error) {
        console.error('Error initializing Independence QA system:', error);
        safeWriteError('initializeIndependenceQASystem', error);
        throw error;
    }
}

/**
 * Initialize main QA sheet with proper headers
 */
function initializeQASheet(spreadsheet) {
    try {
        console.log('Initializing QA sheet...');
        
        // Validate headers again
        if (!INDEPENDENCE_QA_HEADERS || INDEPENDENCE_QA_HEADERS.length === 0) {
            throw new Error('INDEPENDENCE_QA_HEADERS is not defined or empty');
        }
        
        console.log('Using headers array with', INDEPENDENCE_QA_HEADERS.length, 'columns');
        
        let qaSheet = spreadsheet.getSheetByName(INDEPENDENCE_QA_SHEET);
        
        if (!qaSheet) {
            console.log('Creating new QA sheet...');
            qaSheet = spreadsheet.insertSheet(INDEPENDENCE_QA_SHEET);
        }
        
        // Check if headers are already set properly
        const maxColumns = qaSheet.getMaxColumns();
        console.log('Sheet max columns:', maxColumns);
        
        if (maxColumns > 0) {
            const existingHeaders = qaSheet.getRange(1, 1, 1, Math.min(maxColumns, INDEPENDENCE_QA_HEADERS.length)).getValues()[0];
            if (existingHeaders.length > 0 && existingHeaders[0] !== '' && existingHeaders[0] === INDEPENDENCE_QA_HEADERS[0]) {
                console.log('QA sheet headers already exist and match');
                return;
            }
        }
        
        // Ensure sheet has enough columns
        if (maxColumns < INDEPENDENCE_QA_HEADERS.length) {
            console.log('Expanding sheet columns from', maxColumns, 'to', INDEPENDENCE_QA_HEADERS.length);
            qaSheet.insertColumns(maxColumns, INDEPENDENCE_QA_HEADERS.length - maxColumns);
        }
        
        // Set headers with validation
        console.log('Setting QA sheet headers...', INDEPENDENCE_QA_HEADERS.length, 'columns');
        
        // Create the range safely
        const headerRange = qaSheet.getRange(1, 1, 1, INDEPENDENCE_QA_HEADERS.length);
        console.log('Created header range: 1,1,1,' + INDEPENDENCE_QA_HEADERS.length);
        
        // Set the values
        headerRange.setValues([INDEPENDENCE_QA_HEADERS])
                   .setFontWeight('bold')
                   .setBackground('#003177')
                   .setFontColor('white');
        
        // Freeze header row
        qaSheet.setFrozenRows(1);
        
        // Auto-resize columns (with limit to prevent timeout)
        const maxColumnsToResize = Math.min(INDEPENDENCE_QA_HEADERS.length, 50);
        for (let i = 1; i <= maxColumnsToResize; i++) {
            try {
                qaSheet.autoResizeColumn(i);
            } catch (resizeError) {
                console.warn('Could not resize column', i, ':', resizeError.message);
            }
        }
        
        console.log('QA sheet initialized successfully with', INDEPENDENCE_QA_HEADERS.length, 'columns');
        
    } catch (error) {
        console.error('Error initializing QA sheet:', error);
        console.error('Headers length:', INDEPENDENCE_QA_HEADERS ? INDEPENDENCE_QA_HEADERS.length : 'undefined');
        console.error('Headers sample:', INDEPENDENCE_QA_HEADERS ? INDEPENDENCE_QA_HEADERS.slice(0, 5) : 'undefined');
        throw error;
    }
}

/**
 * Initialize analytics sheet
 */
function initializeAnalyticsSheet(spreadsheet) {
    try {
        console.log('Initializing analytics sheet...');
        
        let analyticsSheet = spreadsheet.getSheetByName(INDEPENDENCE_QA_ANALYTICS_SHEET);
        
        if (!analyticsSheet) {
            console.log('Creating new analytics sheet...');
            analyticsSheet = spreadsheet.insertSheet(INDEPENDENCE_QA_ANALYTICS_SHEET);
            
            const analyticsHeaders = [
                'Date', 'Period', 'Granularity', 'Agent', 'CallType',
                'TotalAssessments', 'AverageScore', 'PassRate', 'ExcellentRate',
                'CategoryScores', 'CriticalFailures', 'TrendData', 'CreatedAt'
            ];
            
            // Validate analytics headers
            if (analyticsHeaders.length === 0) {
                throw new Error('Analytics headers array is empty');
            }
            
            console.log('Setting analytics headers:', analyticsHeaders.length, 'columns');
            
            analyticsSheet.getRange(1, 1, 1, analyticsHeaders.length)
                         .setValues([analyticsHeaders])
                         .setFontWeight('bold')
                         .setBackground('#003177')
                         .setFontColor('white');
            
            analyticsSheet.setFrozenRows(1);
            console.log('Analytics sheet initialized successfully');
        } else {
            console.log('Analytics sheet already exists');
        }
        
    } catch (error) {
        console.error('Error initializing analytics sheet:', error);
        // Don't throw error - analytics is not critical for basic functionality
        console.warn('Continuing without analytics sheet...');
    }
}

/**
 * Initialize Drive folder structure
 */
function initializeDriveFolders() {
    try {
        console.log('Initializing Drive folder structure...');
        
        // Try to access configured folder
        let mainFolder;
        try {
            if (INDEPENDENCE_DRIVE_FOLDER_ID && 
                INDEPENDENCE_DRIVE_FOLDER_ID !== "18_2ZHXPp65WNen-e8tEMJUP1BLUvOMHb" && 
                INDEPENDENCE_DRIVE_FOLDER_ID.length > 10) {
                mainFolder = DriveApp.getFolderById(INDEPENDENCE_DRIVE_FOLDER_ID);
                console.log('Using configured Drive folder:', mainFolder.getName());
            } else {
                throw new Error('No valid folder ID configured');
            }
        } catch (error) {
            console.log('Creating new main folder...');
            mainFolder = DriveApp.createFolder('Independence Insurance QA');
            console.log('Created new main folder:', mainFolder.getId());
            console.log('⚠️ IMPORTANT: Update INDEPENDENCE_DRIVE_FOLDER_ID to:', mainFolder.getId());
        }
        
        // Create subfolders if they don't exist
        const subfolders = ['Call Recordings', 'PDF Reports', 'Analytics'];
        subfolders.forEach(folderName => {
            try {
                const existing = mainFolder.getFoldersByName(folderName);
                if (!existing.hasNext()) {
                    mainFolder.createFolder(folderName);
                    console.log('Created subfolder:', folderName);
                } else {
                    console.log('Subfolder already exists:', folderName);
                }
            } catch (subfolderError) {
                console.warn('Could not create subfolder', folderName, ':', subfolderError.message);
            }
        });
        
        console.log('Drive folder structure initialized');
        
    } catch (error) {
        console.error('Error initializing Drive folders:', error);
        console.warn('Continuing without Drive folder setup...');
        // Don't throw error - folder creation is not critical for basic functionality
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** Safe Error Writing */
// ────────────────────────────────────────────────────────────────────────────

function safeWriteError(functionName, error) {
    try {
        const errorMessage = `${new Date().toISOString()} - ${functionName}: ${error.message || error}`;
        console.error(errorMessage);
        
        // Try to write to error log sheet if it exists
        try {
            const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
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
            console.error('Could not write to error log:', logError);
        }
        
    } catch (safeError) {
        console.error('Error in safeWriteError:', safeError);
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** Generate QA ID */
// ────────────────────────────────────────────────────────────────────────────

function generateIndependenceQAId() {
    const timestamp = new Date().getTime();
    const random = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    return `IND_QA_${timestamp}_${random}`;
}

// ────────────────────────────────────────────────────────────────────────────
/** Process Audio File From Form Data */
// ────────────────────────────────────────────────────────────────────────────

function processAudioFileFromFormData(audioFileData, assessmentId) {
    try {
        if (!audioFileData || !audioFileData.bytes) {
            console.log('No audio file data provided');
            return '';
        }
        
        console.log('Processing audio file:', audioFileData.name);
        
        // Convert the array back to a Blob
        const blob = Utilities.newBlob(
            audioFileData.bytes,
            audioFileData.mimeType,
            audioFileData.name
        );
        
        // Upload to Drive
        const audioUrl = uploadIndependenceAudioFile(blob, assessmentId);
        
        console.log('Audio file processed successfully:', audioUrl);
        return audioUrl;
        
    } catch (error) {
        console.error('Error processing audio file:', error);
        throw error;
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** Upload Audio File to Drive */
// ────────────────────────────────────────────────────────────────────────────

function uploadIndependenceAudioFile(audioBlob, assessmentId) {
    try {
        console.log('Uploading audio file for assessment:', assessmentId);
        
        // Get or create audio uploads folder
        const audioFolder = getOrCreateIndependenceAudioFolder();
        
        // Generate filename with timestamp
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
        const originalName = audioBlob.getName() || 'recording.mp3';
        const extension = originalName.split('.').pop() || 'mp3';
        const fileName = `${assessmentId}_${timestamp}.${extension}`;
        
        // Create file in Drive
        const audioFile = audioFolder.createFile(audioBlob.setName(fileName));
        
        // Set sharing permissions (optional - adjust based on your needs)
        audioFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        const audioUrl = audioFile.getUrl();
        console.log('Audio file uploaded successfully:', audioUrl);
        
        return audioUrl;
        
    } catch (error) {
        console.error('Error uploading audio file:', error);
        safeWriteError('uploadIndependenceAudioFile', error);
        return '';
    }
}

/**
 * Get or create audio uploads folder
 */
function getOrCreateIndependenceAudioFolder() {
    try {
        // Try to get main Independence folder first
        let mainFolder;
        try {
            mainFolder = DriveApp.getFolderById(INDEPENDENCE_DRIVE_FOLDER_ID);
        } catch (e) {
            // If folder ID doesn't exist, create in root
            mainFolder = DriveApp.createFolder('Independence Insurance QA');
        }
        
        // Look for audio subfolder
        const audioFolders = mainFolder.getFoldersByName('Call Recordings');
        if (audioFolders.hasNext()) {
            return audioFolders.next();
        } else {
            // Create audio subfolder
            return mainFolder.createFolder('Call Recordings');
        }
        
    } catch (error) {
        console.error('Error getting/creating audio folder:', error);
        // Fallback: create in root Drive
        return DriveApp.createFolder('Independence QA Audio');
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** Update Analytics */
// ────────────────────────────────────────────────────────────────────────────

function updateIndependenceQAAnalytics(qaData, scoreResults) {
    try {
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
        let analyticsSheet = ss.getSheetByName(INDEPENDENCE_QA_ANALYTICS_SHEET);
        
        if (!analyticsSheet) {
            // Create analytics sheet if it doesn't exist
            analyticsSheet = ss.insertSheet(INDEPENDENCE_QA_ANALYTICS_SHEET);
            
            const headers = [
                'Date', 'Period', 'Granularity', 'Agent', 'CallType',
                'TotalAssessments', 'AverageScore', 'PassRate', 'ExcellentRate',
                'CategoryScores', 'CriticalFailures', 'TrendData', 'CreatedAt'
            ];
            
            analyticsSheet.getRange(1, 1, 1, headers.length)
                         .setValues([headers])
                         .setFontWeight('bold')
                         .setBackground('#003177')
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
            qaData.CallType,
            1, // This assessment
            scoreResults.overallPercentage,
            scoreResults.hasCriticalFailure ? 0 : (scoreResults.overallPercentage >= 85 ? 100 : 0),
            scoreResults.overallPercentage >= 95 ? 100 : 0,
            JSON.stringify(scoreResults.categoryScores),
            scoreResults.hasCriticalFailure ? scoreResults.criticalFailures.length : 0,
            JSON.stringify({date: today, score: scoreResults.overallPercentage}),
            now
        ]);
        
        console.log('Analytics updated successfully');
        
    } catch (error) {
        console.error('Error updating analytics:', error);
        // Don't throw error - analytics failure shouldn't stop submission
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** Save Independence QA to Sheet */
// ────────────────────────────────────────────────────────────────────────────

function saveIndependenceQAToSheet(qaData) {
    try {
        console.log('Saving Independence QA data to sheet...');
        
        // Validate constants are defined
        if (typeof INDEPENDENCE_SHEET_ID === 'undefined') {
            throw new Error('INDEPENDENCE_SHEET_ID is not defined');
        }
        if (typeof INDEPENDENCE_QA_SHEET === 'undefined') {
            throw new Error('INDEPENDENCE_QA_SHEET is not defined');
        }
        if (typeof INDEPENDENCE_QA_HEADERS === 'undefined') {
            throw new Error('INDEPENDENCE_QA_HEADERS is not defined');
        }
        
        console.log('Using spreadsheet ID:', INDEPENDENCE_SHEET_ID);
        console.log('Using sheet name:', INDEPENDENCE_QA_SHEET);
        
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
        let qaSheet = ss.getSheetByName(INDEPENDENCE_QA_SHEET);
        
        if (!qaSheet) {
            console.log('QA sheet not found, creating it...');
            qaSheet = ss.insertSheet(INDEPENDENCE_QA_SHEET);
            
            // Set headers if sheet is new
            qaSheet.getRange(1, 1, 1, INDEPENDENCE_QA_HEADERS.length)
                   .setValues([INDEPENDENCE_QA_HEADERS])
                   .setFontWeight('bold')
                   .setBackground('#003177')
                   .setFontColor('white');
            
            qaSheet.setFrozenRows(1);
        }
        
        // Prepare row data in the correct order
        const rowData = INDEPENDENCE_QA_HEADERS.map(header => {
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
        
        console.log('Prepared row data with', rowData.length, 'columns');
        console.log('Sample data:', rowData.slice(0, 10));
        
        // Append the row
        qaSheet.appendRow(rowData);
        
        // Format the new row
        const lastRow = qaSheet.getLastRow();
        formatIndependenceQARow(qaSheet, lastRow, qaData);
        
        console.log('Independence QA data saved to sheet successfully at row', lastRow);
        return { success: true, row: lastRow };
        
    } catch (error) {
        console.error('Error saving to Independence QA sheet:', error);
        return { success: false, error: error.message };
    }
}

/**
 * Format QA row in sheet with conditional formatting
 */
function formatIndependenceQARow(sheet, rowNumber, qaData) {
    try {
        if (!sheet || !rowNumber || !qaData) {
            console.warn('Invalid parameters for formatting row');
            return;
        }
        
        // Apply conditional formatting based on pass status
        const passStatus = qaData.PassStatus || 'Unknown';
        let backgroundColor;
        
        if (passStatus.includes('Critical Failure')) {
            backgroundColor = '#ffebee'; // Light red
        } else if (passStatus === 'Excellent Performance') {
            backgroundColor = '#e8f5e8'; // Light green
        } else if (passStatus === 'Meets Standards') {
            backgroundColor = '#fff3e0'; // Light orange
        } else {
            backgroundColor = '#fafafa'; // Light gray
        }
        
        // Apply background color to the entire row
        const range = sheet.getRange(rowNumber, 1, 1, INDEPENDENCE_QA_HEADERS.length);
        range.setBackground(backgroundColor);
        
        // Bold the assessment ID and score columns
        const idColumn = INDEPENDENCE_QA_HEADERS.indexOf('ID') + 1;
        const scoreColumn = INDEPENDENCE_QA_HEADERS.indexOf('PercentageScore') + 1;
        const statusColumn = INDEPENDENCE_QA_HEADERS.indexOf('PassStatus') + 1;
        
        if (idColumn > 0) {
            sheet.getRange(rowNumber, idColumn).setFontWeight('bold');
        }
        if (scoreColumn > 0) {
            sheet.getRange(rowNumber, scoreColumn).setFontWeight('bold');
        }
        if (statusColumn > 0) {
            sheet.getRange(rowNumber, statusColumn).setFontWeight('bold');
        }
        
        console.log('Row formatting applied successfully');
        
    } catch (error) {
        console.error('Error formatting QA row:', error);
        // Don't throw error - formatting is not critical
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** Prepare Independence QA Data */
// ────────────────────────────────────────────────────────────────────────────

function prepareIndependenceQAData(formData, scoreResults, assessmentId, audioUrl) {
    try {
        console.log('Preparing Independence QA data for storage...');
        
        const now = new Date();
        
        // Convert HTML feedback to plain text for the main field
        const overallFeedbackHtml = formData.overallFeedbackHtml || formData.overallFeedback || '';
        const overallFeedbackText = stripHtmlTags(overallFeedbackHtml);
        
        const qaData = {
            // Basic Information
            ID: assessmentId,
            Timestamp: now,
            CallerName: formData.callerName || '',
            AgentName: formData.agentName || '',
            AgentEmail: formData.agentEmail || '',
            CallDate: formData.callDate || '',
            AuditorName: formData.auditorName || '',
            AuditDate: Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
            CallType: formData.callType || '',
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
            
            // Metadata
            FeedbackShared: false,
            AgentAcknowledgment: false,
            CreatedAt: now,
            UpdatedAt: now
        };
        
        // Add individual question responses and comments
        if (INDEPENDENCE_QA_CONFIG && INDEPENDENCE_QA_CONFIG.categories) {
            Object.values(INDEPENDENCE_QA_CONFIG.categories).forEach(category => {
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
        
        console.log('QA data prepared successfully');
        console.log('Data keys:', Object.keys(qaData).length);
        
        return qaData;
        
    } catch (error) {
        console.error('Error preparing QA data:', error);
        throw error;
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** Helper Functions */
// ────────────────────────────────────────────────────────────────────────────

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

function extractAreasForImprovement(formData) {
    const improvements = [];
    
    if (INDEPENDENCE_QA_CONFIG && INDEPENDENCE_QA_CONFIG.categories) {
        Object.values(INDEPENDENCE_QA_CONFIG.categories).forEach(category => {
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

function extractStrengths(formData) {
    const strengths = [];
    
    if (INDEPENDENCE_QA_CONFIG && INDEPENDENCE_QA_CONFIG.categories) {
        Object.values(INDEPENDENCE_QA_CONFIG.categories).forEach(category => {
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

function extractActionItems(formData) {
    const feedback = stripHtmlTags(formData.overallFeedbackHtml || formData.overallFeedback || '');
    
    // Simple extraction of action-oriented sentences
    const actionWords = ['should', 'must', 'need', 'recommend', 'suggest', 'improve', 'focus'];
    const sentences = feedback.split(/[.!?]+/);
    
    const actionItems = sentences.filter(sentence => 
        actionWords.some(word => sentence.toLowerCase().includes(word))
    ).map(item => item.trim()).filter(item => item.length > 10);
    
    return actionItems.join('; ');
}

// ────────────────────────────────────────────────────────────────────────────
/** Configuration Validation */
// ────────────────────────────────────────────────────────────────────────────

function validateIndependenceConfiguration() {
    console.log('=== Validating Independence Configuration ===');
    
    const issues = [];
    const status = {
        INDEPENDENCE_SHEET_ID: false,
        INDEPENDENCE_QA_SHEET: false,
        INDEPENDENCE_QA_HEADERS: false,
        INDEPENDENCE_DRIVE_FOLDER_ID: false
    };
    
    // Check INDEPENDENCE_SHEET_ID
    if (typeof INDEPENDENCE_SHEET_ID === 'undefined') {
        issues.push('❌ INDEPENDENCE_SHEET_ID is not defined');
    } else if (!INDEPENDENCE_SHEET_ID || INDEPENDENCE_SHEET_ID.length < 10) {
        issues.push('❌ INDEPENDENCE_SHEET_ID appears to be invalid');
    } else {
        status.INDEPENDENCE_SHEET_ID = true;
        console.log('✅ INDEPENDENCE_SHEET_ID is defined');
    }
    
    // Check INDEPENDENCE_QA_SHEET
    if (typeof INDEPENDENCE_QA_SHEET === 'undefined') {
        issues.push('❌ INDEPENDENCE_QA_SHEET is not defined');
    } else {
        status.INDEPENDENCE_QA_SHEET = true;
        console.log('✅ INDEPENDENCE_QA_SHEET is defined:', INDEPENDENCE_QA_SHEET);
    }
    
    // Check INDEPENDENCE_QA_HEADERS
    if (typeof INDEPENDENCE_QA_HEADERS === 'undefined') {
        issues.push('❌ INDEPENDENCE_QA_HEADERS is not defined');
    } else if (!Array.isArray(INDEPENDENCE_QA_HEADERS)) {
        issues.push('❌ INDEPENDENCE_QA_HEADERS is not an array');
    } else if (INDEPENDENCE_QA_HEADERS.length === 0) {
        issues.push('❌ INDEPENDENCE_QA_HEADERS array is empty');
    } else {
        status.INDEPENDENCE_QA_HEADERS = true;
        console.log('✅ INDEPENDENCE_QA_HEADERS is defined with', INDEPENDENCE_QA_HEADERS.length, 'columns');
    }
    
    // Check INDEPENDENCE_DRIVE_FOLDER_ID
    if (typeof INDEPENDENCE_DRIVE_FOLDER_ID === 'undefined') {
        issues.push('⚠️ INDEPENDENCE_DRIVE_FOLDER_ID is not defined (will auto-create)');
    } else if (!INDEPENDENCE_DRIVE_FOLDER_ID || INDEPENDENCE_DRIVE_FOLDER_ID.length < 10) {
        issues.push('⚠️ INDEPENDENCE_DRIVE_FOLDER_ID appears to be empty (will auto-create)');
    } else {
        status.INDEPENDENCE_DRIVE_FOLDER_ID = true;
        console.log('✅ INDEPENDENCE_DRIVE_FOLDER_ID is defined');
    }
    
    const isValid = status.INDEPENDENCE_SHEET_ID && 
                   status.INDEPENDENCE_QA_SHEET && 
                   status.INDEPENDENCE_QA_HEADERS;
    
    console.log('=== Configuration Status ===');
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
/** Testing */
// ────────────────────────────────────────────────────────────────────────────

function testIndependenceQASystem() {
    try {
        console.log('=== Testing Independence QA System ===');
        
        // Check configuration
        const configCheck = validateIndependenceConfiguration();
        if (!configCheck.valid) {
            console.error('Configuration issues:', configCheck.issues);
        }
        
        // Test system initialization
        initializeIndependenceQASystem();
        
        // Test scoring system (if available)
        if (typeof calculateIndependenceQAScores === 'function') {
            const testData = {
                Q1_ProfessionalGreeting: 'Yes',
                Q2_ProperIntroduction: 'Yes',
                Q3_ToneMatching: 'No'
            };
            
            const testScores = calculateIndependenceQAScores(testData);
            console.log('Test scoring result:', testScores);
        }
        
        console.log('=== System Test Completed ===');
        return { 
            success: true, 
            message: 'All systems operational',
            configCheck: configCheck
        };
        
    } catch (error) {
        console.error('System test failed:', error);
        return { 
            success: false, 
            error: error.message,
            configCheck: configCheck 
        };
    }
}

// ────────────────────────────────────────────────────────────────────────────
/** AUTO-VALIDATE ON LOAD */
// ────────────────────────────────────────────────────────────────────────────

console.log('COMPLETE Independence QA Configuration loaded');
console.log('Total headers defined:', INDEPENDENCE_QA_HEADERS ? INDEPENDENCE_QA_HEADERS.length : 'undefined');
console.log('Coaching headers defined:', INDEPENDENCE_COACHING_HEADERS ? INDEPENDENCE_COACHING_HEADERS.length : 'undefined');
console.log('Functions available: initializeIndependenceQASystem, safeWriteError, generateIndependenceQAId, processAudioFileFromFormData, uploadIndependenceAudioFile, updateIndependenceQAAnalytics');

try {
    validateIndependenceConfiguration();
} catch (error) {
    console.error('Error during auto-validation:', error);
}
