// TasksConfig.gs - Configuration and Setup for Tasks Module

// ───────────────────────────────────────────────────────────────────────────────
// GLOBAL CONSTANTS
// ───────────────────────────────────────────────────────────────────────────────

// Sheet Names
const TASKS_SHEET = 'Tasks';
const COMMENTS_SHEET = 'TaskComments';

// Drive Configuration
const TASKS_FOLDER_ID = PropertiesService.getScriptProperties().getProperty('TASKS_FOLDER_ID') || 'your_drive_folder_id_here';

// Task Headers (must match the order in addOrUpdateTask function)
const TASKS_HEADERS = [
  'ID',
  'Task', 
  'Owner',
  'Delegations',
  'StartDate',
  'StartTime',
  'EndDate', 
  'EndTime',
  'AllDay',
  'DueDate',
  'RecurrenceRule',
  'Dependencies',
  'Priority',
  'Calendar',
  'Category',
  'Status',
  'Notes',
  'CompletedDate',
  'NotifyOnComplete',
  'ApprovalStatus',
  'SharedLink',
  'Attachments',
  'CreatedAt',
  'UpdatedAt'
];

// Comments Headers
const COMMENTS_HEADERS = [
  'TaskID',
  'Timestamp', 
  'Author',
  'Text'
];

// Default Values
const DEFAULT_TASK_STATUS = 'Needs Action';
const DEFAULT_PRIORITY = 'None';
const DEFAULT_CALENDAR = 'Default';

// ───────────────────────────────────────────────────────────────────────────────
// SETUP AND INITIALIZATION
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Initialize the tasks module by creating necessary sheets and setting up triggers
 */
function setupTasksModule() {
  try {
    console.log('Setting up Tasks module...');
    
    // Create necessary sheets
    setupTasksSheets();
    
    // Set up triggers for recurring tasks and EOD reports
    setupTasksTriggers();
    
    // Initialize Drive folder
    setupTasksFolder();
    
    console.log('Tasks module setup completed successfully');
    
  } catch (error) {
    console.error('Error setting up Tasks module:', error);
    writeError('setupTasksModule', error);
    throw error;
  }
}

/**
 * Create and configure the Tasks and Comments sheets
 */
function setupTasksSheets() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Setup Tasks sheet
    let tasksSheet = ss.getSheetByName(TASKS_SHEET);
    if (!tasksSheet) {
      tasksSheet = ss.insertSheet(TASKS_SHEET);
      
      // Add headers
      const headerRange = tasksSheet.getRange(1, 1, 1, TASKS_HEADERS.length);
      headerRange.setValues([TASKS_HEADERS]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f1f3f7');
      
      // Format columns
      formatTasksSheet(tasksSheet);
      
      console.log('Tasks sheet created and formatted');
    }
    
    // Setup Comments sheet
    let commentsSheet = ss.getSheetByName(COMMENTS_SHEET);
    if (!commentsSheet) {
      commentsSheet = ss.insertSheet(COMMENTS_SHEET);
      
      // Add headers
      const headerRange = commentsSheet.getRange(1, 1, 1, COMMENTS_HEADERS.length);
      headerRange.setValues([COMMENTS_HEADERS]);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#f1f3f7');
      
      // Format columns
      formatCommentsSheet(commentsSheet);
      
      console.log('Comments sheet created and formatted');
    }
    
  } catch (error) {
    console.error('Error setting up tasks sheets:', error);
    writeError('setupTasksSheets', error);
    throw error;
  }
}

/**
 * Format the Tasks sheet columns
 */
function formatTasksSheet(sheet) {
  try {
    // Set column widths
    const columnWidths = [
      120, // ID
      300, // Task
      150, // Owner
      150, // Delegations
      100, // StartDate
      80,  // StartTime
      100, // EndDate
      80,  // EndTime
      60,  // AllDay
      100, // DueDate
      200, // RecurrenceRule
      150, // Dependencies
      80,  // Priority
      100, // Calendar
      100, // Category
      100, // Status
      300, // Notes
      100, // CompletedDate
      150, // NotifyOnComplete
      120, // ApprovalStatus
      150, // SharedLink
      200, // Attachments
      150, // CreatedAt
      150  // UpdatedAt
    ];
    
    columnWidths.forEach((width, index) => {
      sheet.setColumnWidth(index + 1, width);
    });
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set data validation for specific columns
    setTasksDataValidation(sheet);
    
  } catch (error) {
    console.error('Error formatting tasks sheet:', error);
    writeError('formatTasksSheet', error);
  }
}

/**
 * Set data validation rules for the Tasks sheet
 */
function setTasksDataValidation(sheet) {
  try {
    const lastRow = Math.max(sheet.getLastRow(), 100); // Set validation for at least 100 rows
    
    // Priority validation
    const priorityCol = TASKS_HEADERS.indexOf('Priority') + 1;
    const priorityRange = sheet.getRange(2, priorityCol, lastRow - 1, 1);
    const priorityRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['None', 'Low', 'Medium', 'High'])
      .setAllowInvalid(false)
      .build();
    priorityRange.setDataValidation(priorityRule);
    
    // Status validation
    const statusCol = TASKS_HEADERS.indexOf('Status') + 1;
    const statusRange = sheet.getRange(2, statusCol, lastRow - 1, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Needs Action', 'In Progress', 'Completed', 'On Hold', 'Overdue'])
      .setAllowInvalid(false)
      .build();
    statusRange.setDataValidation(statusRule);
    
    // AllDay validation
    const allDayCol = TASKS_HEADERS.indexOf('AllDay') + 1;
    const allDayRange = sheet.getRange(2, allDayCol, lastRow - 1, 1);
    const allDayRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Yes', 'No'])
      .setAllowInvalid(false)
      .build();
    allDayRange.setDataValidation(allDayRule);
    
    // Calendar validation
    const calendarCol = TASKS_HEADERS.indexOf('Calendar') + 1;
    const calendarRange = sheet.getRange(2, calendarCol, lastRow - 1, 1);
    const calendarRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Default', 'Work', 'Personal', 'Project'])
      .setAllowInvalid(false)
      .build();
    calendarRange.setDataValidation(calendarRule);
    
  } catch (error) {
    console.error('Error setting data validation:', error);
    writeError('setTasksDataValidation', error);
  }
}

/**
 * Format the Comments sheet columns
 */
function formatCommentsSheet(sheet) {
  try {
    // Set column widths
    sheet.setColumnWidth(1, 150); // TaskID
    sheet.setColumnWidth(2, 150); // Timestamp
    sheet.setColumnWidth(3, 150); // Author
    sheet.setColumnWidth(4, 400); // Text
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
  } catch (error) {
    console.error('Error formatting comments sheet:', error);
    writeError('formatCommentsSheet', error);
  }
}

/**
 * Setup triggers for the tasks module
 */
function setupTasksTriggers() {
  try {
    // Delete existing triggers for tasks functions
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      const functionName = trigger.getHandlerFunction();
      if (functionName === 'sendEODEmail' || functionName === 'generateRecurringTasks') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    // Create EOD email trigger (daily at 6 PM)
    ScriptApp.newTrigger('sendEODEmail')
      .timeBased()
      .everyDays(1)
      .atHour(18)
      .create();
    
    // Create recurring tasks trigger (daily at 12 AM)
    ScriptApp.newTrigger('generateRecurringTasks')
      .timeBased()
      .everyDays(1)
      .atHour(0)
      .create();
    
    console.log('Tasks triggers set up successfully');
    
  } catch (error) {
    console.error('Error setting up tasks triggers:', error);
    writeError('setupTasksTriggers', error);
  }
}

/**
 * Setup the Drive folder for task attachments
 */
function setupTasksFolder() {
  try {
    if (TASKS_FOLDER_ID === 'your_drive_folder_id_here') {
      // Create a new folder if none is specified
      const folder = DriveApp.createFolder('Task Attachments');
      PropertiesService.getScriptProperties().setProperty('TASKS_FOLDER_ID', folder.getId());
      console.log('Created new tasks folder:', folder.getId());
    } else {
      // Verify the existing folder is accessible
      try {
        const folder = DriveApp.getFolderById(TASKS_FOLDER_ID);
        console.log('Using existing tasks folder:', folder.getName());
      } catch (e) {
        console.warn('Specified folder not accessible, creating new one');
        const newFolder = DriveApp.createFolder('Task Attachments');
        PropertiesService.getScriptProperties().setProperty('TASKS_FOLDER_ID', newFolder.getId());
      }
    }
  } catch (error) {
    console.error('Error setting up tasks folder:', error);
    writeError('setupTasksFolder', error);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// UTILITY FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Enhanced readSheet function with error handling and caching
 */
if (typeof readSheet !== 'function') {
  function readSheet(sheetName, useCache = true) {
    try {
      if (typeof dbSelect === 'function') {
        return dbSelect(sheetName, { cache: useCache });
      }

      // Simple caching mechanism
      const cacheKey = `sheet_${sheetName}`;
      const cache = CacheService.getScriptCache();

      if (useCache) {
        const cached = cache.get(cacheKey);
        if (cached) {
          return JSON.parse(cached);
        }
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);

      if (!sheet) {
        console.warn(`Sheet '${sheetName}' not found`);
        return [];
      }

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return [];

      const headers = data[0];
      const rows = data.slice(1);

      const result = rows.map(row => {
        const obj = {};
        headers.forEach((header, index) => {
          obj[header] = row[index] || '';
        });
        return obj;
      });

      // Cache for 5 minutes
      if (useCache) {
        cache.put(cacheKey, JSON.stringify(result), 300);
      }

      return result;

    } catch (error) {
      console.error(`Error reading sheet ${sheetName}:`, error);
      writeError('readSheet', error);
      return [];
    }
  }
}

/**
 * Write error to console and optionally to error log sheet
 */
function writeError(functionName, error) {
  try {
    const errorMessage = `${new Date().toISOString()} - ${functionName}: ${error.message || error}`;
    console.error(errorMessage);
    
    // Optionally write to error log sheet
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let errorSheet = ss.getSheetByName('ErrorLog');
      
      if (!errorSheet) {
        errorSheet = ss.insertSheet('ErrorLog');
        errorSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Stack']]);
        errorSheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      }
      
      errorSheet.appendRow([
        new Date(),
        functionName,
        error.message || error.toString(),
        error.stack || ''
      ]);
      
    } catch (logError) {
      console.error('Failed to write to error log:', logError);
    }
    
  } catch (e) {
    console.error('Error in writeError function:', e);
  }
}

/**
 * Get current week string from date
 */
function weekStringFromDate(date) {
  try {
    const year = date.getFullYear();
    const onejan = new Date(year, 0, 1);
    const millisecsInDay = 86400000;
    const week = Math.ceil((((date - onejan) / millisecsInDay) + onejan.getDay() + 1) / 7);
    return `${year}-W${week.toString().padStart(2, '0')}`;
  } catch (error) {
    console.error('Error calculating week string:', error);
    return '2024-W01';
  }
}

/**
 * Format duration in minutes to human readable format
 */
function formatDuration(minutes) {
  try {
    if (isNaN(minutes) || minutes < 0) return '0:00';
    
    const hours = Math.floor(minutes / 60);
    const mins = Math.floor(minutes % 60);
    
    if (hours > 0) {
      return `${hours}:${mins.toString().padStart(2, '0')}`;
    } else {
      return `0:${mins.toString().padStart(2, '0')}`;
    }
  } catch (error) {
    console.error('Error formatting duration:', error);
    return '0:00';
  }
}

/**
 * Get user's timezone
 */
function getUserTimeZone() {
  try {
    return Session.getScriptTimeZone();
  } catch (error) {
    console.error('Error getting timezone:', error);
    return 'UTC';
  }
}

/**
 * Invalidate cache for a specific sheet
 */
function invalidateCache(sheetName) {
  try {
    const cache = CacheService.getScriptCache();
    cache.remove(`sheet_${sheetName}`);
  } catch (error) {
    console.error('Error invalidating cache:', error);
  }
}

/**
 * Get all users from the Users sheet
 */
function getAllUsers() {
  try {
    if (typeof getUsers === 'function') {
      const scoped = getUsers();
      if (Array.isArray(scoped)) {
        return scoped;
      }
    }

    console.warn('TasksConfig.getAllUsers: getUsers unavailable; returning empty list');
    return [];
  } catch (error) {
    console.error('Error getting all users:', error);
    writeError('getAllUsers', error);
    return [];
  }
}

/**
 * Get user list for dropdowns
 */
function getUserList() {
  try {
    const users = getAllUsers();
    return users.map(user => user.Email || user.UserName || user.FullName).filter(Boolean);
  } catch (error) {
    console.error('Error getting user list:', error);
    writeError('getUserList', error);
    return [];
  }
}

/**
 * Check if a user is an admin
 */
function isUserAdmin(user) {
  try {
    if (!user) return false;
    
    // Check if user has admin role
    if (user.roles && Array.isArray(user.roles)) {
      // Assuming admin role has ID 1 or name 'admin'
      return user.roles.includes(1) || user.roles.includes('admin');
    }
    
    // Check UserType field if it exists
    if (user.UserType === 'Admin' || user.UserType === 'admin') {
      return true;
    }
    
    // Check IsAdmin field if it exists
    if (user.IsAdmin === true || user.IsAdmin === 'true') {
      return true;
    }
    
    return false;
  } catch (error) {
    console.error('Error checking admin status:', error);
    return false;
  }
}

/**
 * Get tasks for EOD reporting by date
 */
function getEODTasksByDate(dateStr) {
  try {
    if (!dateStr) {
      dateStr = Utilities.formatDate(new Date(), getUserTimeZone(), 'yyyy-MM-dd');
    }
    
    return getAllTasks().filter(task => {
      return (task.Status === 'Completed' || task.Status === 'Done') &&
             task.CompletedDate === dateStr;
    });
  } catch (error) {
    console.error('Error getting EOD tasks:', error);
    writeError('getEODTasksByDate', error);
    return [];
  }
}

/**
 * Validate task data before saving
 */
function validateTaskData(taskData) {
  const errors = [];
  
  if (!taskData.task || typeof taskData.task !== 'string' || !taskData.task.trim()) {
    errors.push('Task name is required');
  }
  
  if (!taskData.owner || typeof taskData.owner !== 'string' || !taskData.owner.trim()) {
    errors.push('Task owner is required');
  }
  
  if (taskData.owner && !isValidEmail(taskData.owner)) {
    errors.push('Invalid email format for task owner');
  }
  
  if (taskData.priority && !['None', 'Low', 'Medium', 'High'].includes(taskData.priority)) {
    errors.push('Invalid priority value');
  }
  
  if (taskData.status && !['Needs Action', 'In Progress', 'Completed', 'On Hold', 'Overdue'].includes(taskData.status)) {
    errors.push('Invalid status value');
  }
  
  if (taskData.startDate && taskData.endDate) {
    const startDate = new Date(taskData.startDate);
    const endDate = new Date(taskData.endDate);
    
    if (startDate > endDate) {
      errors.push('Start date cannot be after end date');
    }
  }
  
  return errors;
}

/**
 * Check if email format is valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// ───────────────────────────────────────────────────────────────────────────────
// MIGRATION AND MAINTENANCE FUNCTIONS
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Migrate existing tasks data to new format (if needed)
 */
function migrateTasksData() {
  try {
    console.log('Starting tasks data migration...');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tasksSheet = ss.getSheetByName(TASKS_SHEET);
    
    if (!tasksSheet) {
      console.log('No tasks sheet found, skipping migration');
      return;
    }
    
    const data = tasksSheet.getDataRange().getValues();
    if (data.length < 2) {
      console.log('No tasks data found, skipping migration');
      return;
    }
    
    const currentHeaders = data[0];
    const missingHeaders = TASKS_HEADERS.filter(h => !currentHeaders.includes(h));
    
    if (missingHeaders.length > 0) {
      console.log('Adding missing headers:', missingHeaders);
      
      // Add missing headers
      const lastCol = currentHeaders.length;
      tasksSheet.getRange(1, lastCol + 1, 1, missingHeaders.length)
               .setValues([missingHeaders]);
      
      // Set default values for new columns
      const lastRow = tasksSheet.getLastRow();
      if (lastRow > 1) {
        const defaultValues = missingHeaders.map(header => {
          switch (header) {
            case 'AllDay': return 'No';
            case 'Priority': return 'None';
            case 'Status': return 'Needs Action';
            case 'Calendar': return 'Default';
            case 'Dependencies': return '[]';
            case 'NotifyOnComplete': return '[]';
            case 'Attachments': return '[]';
            default: return '';
          }
        });
        
        for (let row = 2; row <= lastRow; row++) {
          tasksSheet.getRange(row, lastCol + 1, 1, missingHeaders.length)
                   .setValues([defaultValues]);
        }
      }
    }
    
    console.log('Tasks data migration completed');
    
  } catch (error) {
    console.error('Error during tasks migration:', error);
    writeError('migrateTasksData', error);
  }
}

/**
 * Clean up old completed tasks (run monthly)
 */
function cleanupOldTasks() {
  try {
    console.log('Starting cleanup of old tasks...');
    
    const cutoffDate = new Date();
    cutoffDate.setMonth(cutoffDate.getMonth() - 6); // Keep tasks from last 6 months
    const cutoffStr = Utilities.formatDate(cutoffDate, getUserTimeZone(), 'yyyy-MM-dd');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tasksSheet = ss.getSheetByName(TASKS_SHEET);
    
    if (!tasksSheet) return;
    
    const data = tasksSheet.getDataRange().getValues();
    const headers = data[0];
    
    const statusCol = headers.indexOf('Status');
    const completedDateCol = headers.indexOf('CompletedDate');
    
    let deletedCount = 0;
    
    // Delete rows from bottom to top to avoid index shifting
    for (let i = data.length - 1; i >= 1; i--) {
      const row = data[i];
      const status = row[statusCol];
      const completedDate = row[completedDateCol];
      
      if ((status === 'Completed' || status === 'Done') && 
          completedDate && completedDate < cutoffStr) {
        tasksSheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    
    console.log(`Cleaned up ${deletedCount} old completed tasks`);
    
  } catch (error) {
    console.error('Error during task cleanup:', error);
    writeError('cleanupOldTasks', error);
  }
}

// ───────────────────────────────────────────────────────────────────────────────
// INITIALIZATION FUNCTION
// ───────────────────────────────────────────────────────────────────────────────

/**
 * Main initialization function - call this to set up the entire tasks module
 */
function initializeTasksModule() {
  try {
    console.log('Initializing Tasks Module...');
    
    // Run setup
    setupTasksModule();
    
    // Run migration (if needed)
    migrateTasksData();
    
    console.log('Tasks Module initialization completed successfully');
    
  } catch (error) {
    console.error('Failed to initialize Tasks Module:', error);
    writeError('initializeTasksModule', error);
    throw error;
  }
}

// Auto-run initialization when this script is loaded
// Uncomment the next line if you want automatic initialization
// initializeTasksModule();