/**
 * COMPLETE QA SERVICE - QAService.gs
 * Replace your existing QAService.gs with this complete implementation
 */

// ============================================================================
// CORE QA CONFIGURATION
// ============================================================================

const ROOT_PROP_KEY = '1GuTbUeWFdp7u6nVDdavc_rCXfUMKX2DH';
const FALLBACK_PATH = ['Lumina', 'QA Uploads'];

function qaWeights_() {
  return {
    q1: 3, q2: 5, q3: 7, q4: 10, q5: 5,
    q6: 8, q7: 8, q8: 15, q9: 9,
    q10: 8, q11: 6, q12: 6, q13: 7, q14: 3,
    q15: 10, q16: 5, q17: 4, q18: 5
  };
}

function qaQuestionText_() {
  return {
    q1: 'Did the agent say thank you for calling and brand the call with the client name?',
    q2: 'Did the agent offer further assistance before closing the call?',
    q3: 'Did the agent sound polite and courteous on the call?',
    q4: 'Did the agent empathize with callers issue?',
    q5: 'Did the agent modulate pitch and volume according to caller?',
    q6: 'Did the agent follow the call transfer protocol?',
    q7: 'Did the agent follow the correct HOLD procedure?',
    q8: 'Did the agent authenticate caller and confirm the issue?',
    q9: 'Did the agent do effective probing on the call?',
    q10: 'Did the agent provide accurate and complete resolution using tools?',
    q11: 'Did the agent provide clear understanding of the issue to the caller?',
    q12: 'Did the agent process the request as promised to the caller?',
    q13: 'Did the agent document the case/Wrapup up with Wrap-Up notes correctly?',
    q14: 'Did the agent escalate the case to the right department with all relevant details?',
    q15: 'Did the agent offer the survey at the end of the call?',
    q16: 'Did the agent create the call case correctly/Wrap-Up the call?',
    q17: 'Did the agent modify case fields to avoid survey going to the customer?',
    q18: 'Did the customer respond positively to the survey?'
  };
}

function getUserEmail(agentName) {
  try {
    console.log('getUserEmail called with:', agentName);
    
    if (!agentName || typeof agentName !== 'string') {
      return '';
    }
    
    // Clean the agent name for comparison
    const cleanName = agentName.trim();
    
    // Method 1: Try Users sheet directly
    try {
      const ss = getIBTRSpreadsheet();
      const userSheet = ss.getSheetByName('Users') || 
                       ss.getSheetByName('Agents') || 
                       ss.getSheetByName('Agent List');
      
      if (userSheet) {
        const data = userSheet.getDataRange().getValues();
        const headers = data[0].map(h => String(h).toLowerCase().trim());
        
        // Find column indices more flexibly
        const nameColIndices = headers.reduce((acc, h, i) => {
          if (h.includes('name') || h.includes('agent')) acc.push(i);
          return acc;
        }, []);
        
        const emailColIndices = headers.reduce((acc, h, i) => {
          if (h.includes('email') || h.includes('mail')) acc.push(i);
          return acc;
        }, []);
        
        // Search through rows
        for (let i = 1; i < data.length; i++) {
          for (let nameIdx of nameColIndices) {
            const cellName = String(data[i][nameIdx] || '').trim();
            if (cellName === cleanName || cellName.toLowerCase() === cleanName.toLowerCase()) {
              // Found the name, now get the email
              for (let emailIdx of emailColIndices) {
                const email = String(data[i][emailIdx] || '').trim();
                if (email && email.includes('@')) {
                  console.log('Email found in sheet:', email);
                  return email;
                }
              }
            }
          }
        }
      }
    } catch (sheetError) {
      console.warn('Sheet lookup failed:', sheetError);
    }
    
    // Method 2: Try getUsers function with improved matching
    try {
      const users = getUsers();
      if (Array.isArray(users) && users.length > 0) {
        // Try exact match first
        let user = users.find(u => {
          const uName = u.FullName || u.UserName || u.name || u.fullName || u.displayName || '';
          return uName === cleanName;
        });
        
        // Try case-insensitive match if exact match fails
        if (!user) {
          user = users.find(u => {
            const uName = (u.FullName || u.UserName || u.name || u.fullName || u.displayName || '').toLowerCase();
            return uName === cleanName.toLowerCase();
          });
        }
        
        if (user) {
          const email = user.Email || user.email || user.mail || 
                       user.emailAddress || user.EmailAddress || '';
          if (email) {
            console.log('Email found via getUsers:', email);
            return email;
          }
        }
      }
    } catch (error) {
      console.warn('getUsers failed:', error);
    }
    
    // Method 3: Try QA Records as last resort
    try {
      const qaSheet = getQaSheet_();
      const qaData = qaSheet.getDataRange().getValues();
      const qaHeaders = qaData[0];
      const agentNameCol = qaHeaders.findIndex(h => 
        String(h).toLowerCase().includes('agentname')
      );
      const agentEmailCol = qaHeaders.findIndex(h => 
        String(h).toLowerCase().includes('agentemail')
      );
      
      if (agentNameCol >= 0 && agentEmailCol >= 0) {
        for (let i = 1; i < qaData.length; i++) {
          if (String(qaData[i][agentNameCol]).trim() === cleanName) {
            const email = String(qaData[i][agentEmailCol] || '').trim();
            if (email && email.includes('@')) {
              console.log('Email found in QA records:', email);
              return email;
            }
          }
        }
      }
    } catch (qaError) {
      console.warn('QA lookup failed:', qaError);
    }
    
    console.log('No email found for:', cleanName);
    return '';
    
  } catch (error) {
    console.error('Error in getUserEmail:', error);
    return '';
  }
}

function getUsersWithEmails() {
  try {
    const allUsers = [];
    const processedNames = new Set();
    
    // Source 1: Users/Agents sheet
    try {
      const ss = getIBTRSpreadsheet();
      const sheets = ['Users', 'Agents', 'Agent List'];
      
      for (const sheetName of sheets) {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) continue;
        
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) continue;
        
        const headers = data[0].map(h => String(h).toLowerCase().trim());
        const nameIndices = headers.reduce((acc, h, i) => {
          if (h.includes('name') || h.includes('agent')) acc.push(i);
          return acc;
        }, []);
        
        const emailIndices = headers.reduce((acc, h, i) => {
          if (h.includes('email') || h.includes('mail')) acc.push(i);
          return acc;
        }, []);
        
        for (let i = 1; i < data.length; i++) {
          for (let nameIdx of nameIndices) {
            const name = String(data[i][nameIdx] || '').trim();
            if (!name || processedNames.has(name)) continue;
            
            for (let emailIdx of emailIndices) {
              const email = String(data[i][emailIdx] || '').trim();
              if (email && email.includes('@')) {
                allUsers.push({ name: name, email: email });
                processedNames.add(name);
                break;
              }
            }
          }
        }
      }
    } catch (e) {
      console.warn('Sheet processing error:', e);
    }
    
    // Source 2: getUsers function
    try {
      const users = getUsers();
      if (Array.isArray(users)) {
        users.forEach(u => {
          const name = u.FullName || u.UserName || u.name || u.fullName || '';
          const email = u.Email || u.email || u.mail || '';
          if (name && email && !processedNames.has(name)) {
            allUsers.push({ name: name, email: email });
            processedNames.add(name);
          }
        });
      }
    } catch (e) {
      console.warn('getUsers error:', e);
    }
    
    // Sort by name
    allUsers.sort((a, b) => a.name.localeCompare(b.name));
    
    console.log('Found', allUsers.length, 'users with emails');
    return allUsers;
    
  } catch (error) {
    console.error('Error in getUsersWithEmails:', error);
    return [];
  }
}

// ============================================================================
// MAIN QA SUBMISSION FUNCTION (SIMPLIFIED & ROBUST)
// ============================================================================

function clientUploadAudioAndSaveQA(formData) {
  try {
    console.log('=== QA SUBMISSION STARTED ===');
    console.log('Function called at:', new Date().toISOString());
    
    // Step 1: Validate input
    if (!formData) {
      throw new Error('No form data provided');
    }
    
    // Step 2: Extract data safely
    const qaData = extractFormData_(formData);
    console.log('Form data extracted successfully');
    
    // Step 3: Validate required fields
    validateRequiredFields_(qaData);
    console.log('Required fields validated');
    
    // Step 4: Process audio file/link
    const audioResult = processAudioFile_(qaData);
    console.log('Audio processing completed');
    
    // Step 5: Calculate QA score
    const scoreResult = calculateQAScore_(qaData);
    console.log('Score calculated:', scoreResult.finalScore);
    
    // Step 6: Save to sheet
    const saveResult = saveQARecord_(qaData, audioResult, scoreResult);
    console.log('Record saved with ID:', saveResult.qaId);
    
    // Step 7: Generate PDF (optional, don't fail if this errors)
    let pdfResult = null;
    try {
      pdfResult = generateQAPDF_(saveResult.record, scoreResult);
    } catch (pdfError) {
      console.warn('PDF generation failed (non-critical):', pdfError.message);
    }
    
    // Step 8: Return success response
    const response = {
      success: true,
      qaId: saveResult.qaId,
      audioUrl: audioResult.url,
      scoreResult: scoreResult,
      record: saveResult.record,
      timestamp: new Date().toISOString()
    };
    
    if (pdfResult && pdfResult.success) {
      response.qaPdfUrl = pdfResult.fileUrl;
      response.qaPdfId = pdfResult.fileId;
    }
    
    console.log('=== QA SUBMISSION COMPLETED SUCCESSFULLY ===');
    return response;
    
  } catch (error) {
    console.error('=== QA SUBMISSION FAILED ===');
    console.error('Error:', error.message);
    console.error('Stack:', error.stack);
    
    // Return a proper error response
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
      debug: {
        function: 'clientUploadAudioAndSaveQA',
        stack: error.stack
      }
    };
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function extractFormData_(input) {
  try {
    let data = {};
    
    // Handle different input types
    if (input && typeof input === 'object') {
      if (input.elements) {
        // HTML Form Element
        console.log('Processing HTML form element');
        for (let i = 0; i < input.elements.length; i++) {
          const element = input.elements[i];
          if (element.name) {
            if (element.type === 'file' && element.files && element.files.length > 0) {
              data[element.name] = element.files[0];
            } else if (element.type === 'checkbox') {
              data[element.name] = element.checked;
            } else if (element.type === 'radio' && element.checked) {
              data[element.name] = element.value;
            } else if (element.value !== undefined) {
              data[element.name] = element.value;
            }
          }
        }
      } else {
        // Direct object
        console.log('Processing direct data object');
        data = { ...input };
      }
    } else {
      throw new Error('Invalid form data format');
    }
    
    console.log('Extracted data keys:', Object.keys(data));
    return data;
    
  } catch (error) {
    console.error('Error extracting form data:', error);
    throw new Error('Failed to extract form data: ' + error.message);
  }
}

function validateRequiredFields_(data) {
  const required = ['agentName', 'callDate', 'auditorName'];
  const missing = required.filter(field => !data[field] || String(data[field]).trim() === '');
  
  if (missing.length > 0) {
    throw new Error('Missing required fields: ' + missing.join(', '));
  }
  
  // Validate at least some QA questions are answered
  const qaAnswers = [];
  for (let i = 1; i <= 18; i++) {
    const answer = data['q' + i];
    if (answer && answer !== 'na') {
      qaAnswers.push(answer);
    }
  }
  
  if (qaAnswers.length === 0) {
    throw new Error('At least one QA question must be answered');
  }
  
  console.log('Validation passed:', qaAnswers.length, 'questions answered');
}

function processAudioFile_(data) {
  try {
    console.log('Processing audio file...');
    
    // Check for base64 encoded file (new format)
    if (data.audioFileData && data.audioFileName) {
      console.log('Base64 audio file found:', data.audioFileName);
      console.log('File size:', (data.audioFileSize / 1024 / 1024).toFixed(2), 'MB');
      
      // Check file size (45MB limit)
      const fileSizeMB = data.audioFileSize / (1024 * 1024);
      if (fileSizeMB > 45) {
        throw new Error('Audio file too large. Maximum size is 45MB, your file is ' + fileSizeMB.toFixed(1) + 'MB');
      }
      
      // Convert base64 back to blob
      try {
        const base64Data = data.audioFileData;
        const binaryString = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(binaryString, data.audioFileType || 'audio/mpeg', data.audioFileName);
        
        // Upload to Drive
        const folder = ensureRootFolder_();
        const agentName = sanitizeName_(data.agentName || 'Unknown');
        const callDate = data.callDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        
        const agentFolder = getOrCreateFolder_(folder, agentName);
        const dateFolder = getOrCreateFolder_(agentFolder, callDate);
        
        const uploadedFile = dateFolder.createFile(blob);
        uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        console.log('Base64 audio file uploaded successfully');
        return {
          url: uploadedFile.getUrl(),
          id: uploadedFile.getId(),
          name: uploadedFile.getName(),
          size: fileSizeMB
        };
      } catch (conversionError) {
        console.error('Error converting base64 to blob:', conversionError);
        throw new Error('Failed to process audio file: ' + conversionError.message);
      }
    }
    
    // Check for traditional uploaded file (fallback)
    const audioFile = data.audioFile;
    if (audioFile && typeof audioFile.getName === 'function') {
      console.log('Traditional audio file found:', audioFile.getName());
      
      // Check file size (50MB limit)
      const fileSize = audioFile.getBlob().getBytes().length;
      const fileSizeMB = fileSize / (1024 * 1024);
      console.log('File size:', fileSizeMB.toFixed(2), 'MB');
      
      if (fileSizeMB > 45) { // Leave some buffer below 50MB
        throw new Error('Audio file too large. Maximum size is 45MB, your file is ' + fileSizeMB.toFixed(1) + 'MB');
      }
      
      // Upload to Drive
      const folder = ensureRootFolder_();
      const agentName = sanitizeName_(data.agentName || 'Unknown');
      const callDate = data.callDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      const agentFolder = getOrCreateFolder_(folder, agentName);
      const dateFolder = getOrCreateFolder_(agentFolder, callDate);
      
      const uploadedFile = dateFolder.createFile(audioFile.getBlob());
      uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      console.log('Traditional audio file uploaded successfully');
      return {
        url: uploadedFile.getUrl(),
        id: uploadedFile.getId(),
        name: uploadedFile.getName(),
        size: fileSizeMB
      };
    }
    
    // Check for call link
    if (data.callLink && String(data.callLink).trim() !== '') {
      console.log('Using provided call link');
      return {
        url: String(data.callLink).trim(),
        id: null,
        name: 'External Link',
        size: 0
      };
    }
    
    // No audio provided (this might be okay for existing records)
    console.log('No audio file or link provided');
    return {
      url: '',
      id: null,
      name: 'No Audio',
      size: 0
    };
    
  } catch (error) {
    console.error('Audio processing error:', error);
    throw new Error('Audio processing failed: ' + error.message);
  }
}

function calculateQAScore_(data) {
  try {
    console.log('Calculating QA score...');
    
    // Use existing scoring function if available
    if (typeof computeEnhancedQaScore === 'function') {
      const answers = {};
      for (let i = 1; i <= 18; i++) {
        answers['q' + i] = data['q' + i] || '';
      }
      return computeEnhancedQaScore(answers);
    }
    
    // Fallback scoring
    const weights = qaWeights_();
    let earned = 0;
    let applicable = 0;
    
    Object.keys(weights).forEach(q => {
      const answer = String(data[q] || '').toLowerCase();
      if (answer && answer !== 'na') {
        applicable += weights[q];
        if (answer === 'yes' || (q === 'q17' && answer === 'no')) {
          earned += weights[q];
        }
      }
    });
    
    let percentage = applicable ? earned / applicable : 0;
    
    // Apply penalties
    if (String(data.q8 || '').toLowerCase() === 'no') {
      percentage = 0; // Auto-fail
    } else {
      if (String(data.q15 || '').toLowerCase() === 'no') {
        percentage = Math.max(percentage - 0.5, 0);
      }
      if (String(data.q17 || '').toLowerCase() === 'yes') {
        percentage = Math.max(percentage - 0.5, 0);
      }
    }
    
    return {
      earned,
      applicable,
      percentage,
      finalScore: Math.round(percentage * 100),
      isPassing: percentage >= 0.8,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Score calculation error:', error);
    throw new Error('Score calculation failed: ' + error.message);
  }
}

function saveQARecord_(data, audioResult, scoreResult) {
  try {
    console.log('Saving QA record...');
    
    const qaId = Utilities.getUuid();
    const timestamp = new Date().toISOString();
    
    // Get sheet and headers
    const sheet = getQaSheet_();
    const headers = getQaHeaders_();
    
    // Build row data
    const rowData = headers.map(col => {
      switch (col) {
        case 'ID': return qaId;
        case 'Timestamp': return timestamp;
        case 'CallerName': return data.callerName || '';
        case 'AgentName': return data.agentName || '';
        case 'AgentEmail': return data.agentEmail || '';
        case 'ClientName': return data.clientName || '';
        case 'CallDate': return data.callDate || '';
        case 'CaseNumber': return data.caseNumber || '';
        case 'CallLink': return audioResult.url || '';
        case 'AuditorName': return data.auditorName || '';
        case 'AuditDate': return data.auditDate || '';
        case 'FeedbackShared': return data.feedbackShared ? 'Yes' : 'No';
        case 'TotalScore': return scoreResult.earned || 0;
        case 'Percentage': return scoreResult.percentage || 0;
        case 'OverallFeedback': return data.overallFeedback || '';
        case 'Notes': return data.notes || '';
        case 'AgentFeedback': return data.agentFeedback || '';
        default:
          // Handle Q1-Q18 and C1-C18
          if (/^Q\d+$/i.test(col)) {
            const qKey = col.toLowerCase();
            const val = data[qKey];
            if (!val || val === 'na') return 'N/A';
            return val === 'yes' ? 'Yes' : 'No';
          }
          if (/^C\d+$/i.test(col)) {
            return data[col.toLowerCase()] || '';
          }
          // Other fields
          return data[col.charAt(0).toLowerCase() + col.slice(1)] || '';
      }
    });
    
    // Append row
    sheet.appendRow(rowData);
    
    // Build record object
    const record = {};
    headers.forEach((header, index) => {
      record[header] = rowData[index];
    });
    
    console.log('QA record saved with ID:', qaId);
    return {
      qaId,
      record
    };
    
  } catch (error) {
    console.error('Save record error:', error);
    throw new Error('Failed to save QA record: ' + error.message);
  }
}

function generateQAPDF_(record, scoreResult) {
  try {
    if (typeof generateQaPdfReport === 'function') {
      return generateQaPdfReport(record, scoreResult, {
        template: 'standard',
        theme: 'professional'
      });
    }
    return { success: false, error: 'PDF generation not available' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

function sanitizeName_(name) {
  return String(name || 'Unknown')
    .replace(/[\\/:*?"<>|#\[\]]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 50);
}

function getOrCreateFolder_(parent, name) {
  const safeName = sanitizeName_(name);
  const existing = parent.getFoldersByName(safeName);
  return existing.hasNext() ? existing.next() : parent.createFolder(safeName);
}

function ensureRootFolder_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const savedId = (props.getProperty(ROOT_PROP_KEY) || '').trim();

    if (savedId) {
      try {
        const f = DriveApp.getFolderById(savedId);
        f.getFiles(); // Test access
        return f;
      } catch (_) { }
    }

    let parent = DriveApp.getRootFolder();
    for (const name of FALLBACK_PATH) {
      parent = getOrCreateFolder_(parent, name);
    }
    props.setProperty(ROOT_PROP_KEY, parent.getId());
    return parent;
  } catch (error) {
    console.error('Error ensuring root folder:', error);
    throw error;
  }
}

function getQaSheet_() {
  const ss = typeof getIBTRSpreadsheet === 'function' ? getIBTRSpreadsheet() : SpreadsheetApp.getActive();
  const sheetName = (typeof QA_RECORDS !== 'undefined' && QA_RECORDS) ? QA_RECORDS : 'QA Records';
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('QA sheet not found: ' + sheetName);
  return sh;
}

function getQaHeaders_() {
  try {
    if (typeof QA_HEADERS !== 'undefined' && Array.isArray(QA_HEADERS) && QA_HEADERS.length) {
      return QA_HEADERS.slice();
    }
    const sh = getQaSheet_();
    const vals = sh.getRange(1, 1, 1, sh.getLastColumn() || 1).getValues();
    const headers = (vals && vals[0]) ? vals[0] : [];
    if (!headers.length) throw new Error('Header row is empty');
    return headers;
  } catch (error) {
    console.error('Error getting QA headers:', error);
    throw error;
  }
}

function writeError(tag, e) {
  try {
    console.error('[' + tag + '] ' + (e && e.message ? e.message : String(e)));
    Logger.log('[' + tag + '] ' + (e && e.message ? e.message : String(e)));
  } catch (_) { }
}

// ============================================================================
// ENHANCED SCORING ENGINE (if computeEnhancedQaScore doesn't exist)
// ============================================================================

function computeEnhancedQaScore(answers) {
  try {
    const W = qaWeights_();
    let earned = 0, applicable = 0;
    const questionResults = {};

    Object.keys(W).forEach(k => {
      const ans = String(answers[k] || '').trim().toLowerCase();
      questionResults[k] = {
        answer: ans,
        weight: W[k],
        applicable: false,
        points: 0
      };

      if (!ans || ans === 'na') {
        questionResults[k].excluded = true;
        return;
      }

      if (k === 'q18' && ans === 'no') {
        questionResults[k].excluded = true;
        return;
      }

      applicable += W[k];
      questionResults[k].applicable = true;

      if (ans === 'yes' || (k === 'q17' && ans === 'no')) {
        earned += W[k];
        questionResults[k].points = W[k];
      }
    });

    let pct = applicable ? (earned / applicable) : 0;
    const penalties = [];

    // Auto-fail rule
    if (String(answers.q8 || '').toLowerCase() === 'no') {
      pct = 0;
      penalties.push({
        type: 'AUTO_FAIL',
        question: 'q8',
        description: 'Authentication and issue confirmation failed'
      });
    } else {
      // Apply penalties
      if (String(answers.q15 || '').toLowerCase() === 'no') {
        pct = Math.max(pct - 0.5, 0);
        penalties.push({
          type: 'PENALTY',
          question: 'q15',
          amount: -0.5,
          description: 'Survey not offered'
        });
      }
      if (String(answers.q17 || '').toLowerCase() === 'yes') {
        pct = Math.max(pct - 0.5, 0);
        penalties.push({
          type: 'PENALTY',
          question: 'q17',
          amount: -0.5,
          description: 'Case fields modified to avoid survey'
        });
      }
    }

    return {
      earned,
      applicable,
      percentage: pct,
      finalScore: Math.round(pct * 100),
      penalties,
      questionResults,
      isPassing: pct >= 0.8,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('Error computing QA score:', error);
    writeError('computeEnhancedQaScore', error);
    return {
      earned: 0,
      applicable: 0,
      percentage: 0,
      finalScore: 0,
      error: true,
      message: error.message
    };
  }
}

// ============================================================================
// DEBUG AND TESTING FUNCTIONS
// ============================================================================

function debugQASubmission() {
  try {
    console.log('=== QA DEBUG TEST ===');
    
    // Test basic sheet access
    const sheet = getQaSheet_();
    console.log('Sheet accessible:', sheet.getName());
    
    // Test scoring function
    const testAnswers = {
      q1: 'yes', q2: 'yes', q3: 'yes', q8: 'yes', q15: 'yes', q17: 'no'
    };
    const score = computeEnhancedQaScore(testAnswers);
    console.log('Scoring working:', score.finalScore);
    
    return {
      success: true,
      sheetAccess: true,
      scoringWorking: true,
      message: 'QA system is operational'
    };
    
  } catch (error) {
    console.error('QA Debug Error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================================================
// EXISTING FUNCTIONS (for backward compatibility)
// ============================================================================

function getAllQA() {
  try {
    const sh = getQaSheet_();
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data.shift();
    const tz = Session.getScriptTimeZone();

    return data.map(row => {
      const o = {};
      row.forEach((cell, i) => {
        const key = headers[i];
        if (cell instanceof Date) {
          o[key] = Utilities.formatDate(cell, tz, "yyyy-MM-dd'T'HH:mm:ss");
        } else {
          o[key] = cell;
        }
      });
      return o;
    });
  } catch (error) {
    console.error('Error getting all QA records:', error);
    return [];
  }
}

/**
 * Missing Helper Functions for QA PDF Service
 * Add these functions to your QAService.gs file
 */

// ============================================================================
// QA CATEGORIES DEFINITION (Missing from your QAService.gs)
// ============================================================================

function qaCategories_() {
  return {
    'Courtesy & Communication': ['q1', 'q2', 'q3', 'q4', 'q5'],
    'Resolution': ['q6', 'q7', 'q8', 'q9'],
    'Case Documentation': ['q10', 'q11', 'q12', 'q13', 'q14'],
    'Process Compliance': ['q15', 'q16', 'q17', 'q18']
  };
}

// ============================================================================
// ENHANCED PDF INTEGRATION UPDATE
// ============================================================================

// Update the generateQAPDF_ function in QAService.gs to pass form options
function generateQAPDF_(record, scoreResult, formData = {}) {
  try {
    if (typeof generateQaPdfReport === 'function') {
      // Extract PDF options from form data
      const pdfOptions = {
        template: formData.pdfTemplate || 'standard',
        theme: formData.pdfTheme || 'professional',
        includeCharts: formData.includeCharts !== false,
        includeRecommendations: formData.includeRecommendations !== false,
        includeFullSnapshot: formData.includeFullSnapshot === true
      };
      
      console.log('Generating PDF with options:', pdfOptions);
      return generateQaPdfReport(record, scoreResult, pdfOptions);
    }
    return { success: false, error: 'PDF generation not available' };
  } catch (error) {
    console.error('PDF generation error:', error);
    return { success: false, error: error.message };
  }
}

// ============================================================================
// UPDATE TO MAIN SUBMISSION FUNCTION
// ============================================================================

// Update the clientUploadAudioAndSaveQA function to pass form data to PDF generation
// Replace step 7 in your existing function with this:

/*
// Step 7: Generate PDF with form options
let pdfResult = null;
try {
  pdfResult = generateQAPDF_(saveResult.record, scoreResult, qaData);
  if (pdfResult && pdfResult.success) {
    console.log('PDF generated successfully:', pdfResult.fileName);
  }
} catch (pdfError) {
  console.warn('PDF generation failed (non-critical):', pdfError.message);
}
*/

// ============================================================================
// ADDITIONAL UTILITY FUNCTION FOR QA RECORD RETRIEVAL
// ============================================================================

function getQARecordById(qaId) {
  try {
    const sheet = getQaSheet_();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    
    const headers = data[0];
    const idColumnIndex = headers.findIndex(h => h.toLowerCase() === 'id');
    
    if (idColumnIndex === -1) {
      console.error('ID column not found in QA sheet');
      return null;
    }
    
    // Find the row with matching ID
    const rowIndex = data.findIndex((row, index) => 
      index > 0 && row[idColumnIndex] === qaId
    );
    
    if (rowIndex === -1) {
      console.warn('QA record not found for ID:', qaId);
      return null;
    }
    
    // Build record object
    const record = {};
    headers.forEach((header, index) => {
      record[header] = data[rowIndex][index];
    });
    
    return record;
    
  } catch (error) {
    console.error('Error retrieving QA record by ID:', error);
    return null;
  }
}

// ============================================================================
// STANDALONE PDF GENERATION FUNCTION
// ============================================================================

/**
 * Generate PDF for existing QA record by ID
 * Can be called independently from form submission
 */
function generatePdfForExistingQA(qaId, options = {}) {
  try {
    console.log('Generating PDF for existing QA record:', qaId);
    
    // Get the QA record
    const qaRecord = getQARecordById(qaId);
    if (!qaRecord) {
      return {
        success: false,
        error: 'QA record not found for ID: ' + qaId
      };
    }
    
    // Extract answers and compute score
    const answers = {};
    const weights = qaWeights_();
    Object.keys(weights).forEach(k => {
      const qNum = k.replace(/^q/i, '');
      answers[k] = qaRecord[`Q${qNum}`] || '';
    });
    
    // Calculate score using enhanced function
    const scoreResult = computeEnhancedQaScore(answers);
    if (scoreResult.error) {
      return {
        success: false,
        error: 'Score calculation failed: ' + scoreResult.message
      };
    }
    
    // Generate PDF
    const pdfResult = generateQaPdfReport(qaRecord, scoreResult, options);
    
    console.log('PDF generation completed for existing QA:', qaId);
    return pdfResult;
    
  } catch (error) {
    console.error('Error generating PDF for existing QA:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

// ============================================================================
// BATCH PDF GENERATION
// ============================================================================

/**
 * Generate PDFs for multiple QA records
 */
function generateBatchQAPdfs(qaIds, options = {}) {
  const results = [];
  
  qaIds.forEach(qaId => {
    try {
      const result = generatePdfForExistingQA(qaId, options);
      results.push({
        qaId: qaId,
        success: result.success,
        fileUrl: result.fileUrl,
        fileName: result.fileName,
        error: result.error
      });
    } catch (error) {
      results.push({
        qaId: qaId,
        success: false,
        error: error.message
      });
    }
  });
  
  return {
    success: true,
    results: results,
    successCount: results.filter(r => r.success).length,
    totalCount: results.length
  };
}

// ============================================================================
// TESTING FUNCTION
// ============================================================================

function testPdfGeneration() {
  try {
    console.log('=== PDF GENERATION TEST ===');
    
    // Create test data
    const testQaRecord = {
      ID: 'test-' + Utilities.getUuid(),
      AgentName: 'Test Agent',
      AgentEmail: 'test@example.com',
      ClientName: 'Test Client',
      CallDate: '2024-01-15',
      AuditDate: '2024-01-16',
      AuditorName: 'Test Auditor',
      CallerName: 'Test Caller',
      CaseNumber: 'TEST-001',
      CallLink: 'https://example.com/test-recording',
      Q1: 'Yes', Q2: 'Yes', Q3: 'Yes', Q4: 'Yes', Q5: 'Yes',
      Q6: 'Yes', Q7: 'Yes', Q8: 'Yes', Q9: 'Yes',
      Q10: 'Yes', Q11: 'Yes', Q12: 'No', Q13: 'Yes', Q14: 'N/A',
      Q15: 'Yes', Q16: 'Yes', Q17: 'No', Q18: 'Yes',
      C1: 'Great opening', C2: 'Good closing',
      OverallFeedback: '<p>Overall <strong>excellent</strong> performance with minor areas for improvement.</p>',
      TotalScore: 85,
      Percentage: 0.85
    };
    
    const testScoreResult = {
      earned: 85,
      applicable: 100,
      percentage: 0.85,
      finalScore: 85,
      isPassing: true,
      performanceBand: {
        label: 'Good',
        description: 'Meets expectations',
        color: '#3b82f6'
      }
    };
    
    // Test different templates
    const templates = ['standard', 'coaching', 'executive', 'simple'];
    const results = [];
    
    templates.forEach(template => {
      console.log('Testing template:', template);
      const result = generateQaPdfReport(testQaRecord, testScoreResult, {
        template: template,
        theme: 'professional',
        includeCharts: true,
        includeRecommendations: true
      });
      
      results.push({
        template: template,
        success: result.success,
        fileName: result.fileName,
        fileUrl: result.fileUrl,
        error: result.error
      });
    });
    
    console.log('=== TEST RESULTS ===');
    results.forEach(result => {
      console.log(`${result.template}: ${result.success ? 'SUCCESS' : 'FAILED'}`);
      if (result.success) {
        console.log(`  File: ${result.fileName}`);
        console.log(`  URL: ${result.fileUrl}`);
      } else {
        console.log(`  Error: ${result.error}`);
      }
    });
    
    return {
      success: true,
      message: 'PDF generation test completed',
      results: results
    };
    
  } catch (error) {
    console.error('PDF test error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
