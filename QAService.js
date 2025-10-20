/**
 * COMPLETE QA SERVICE - QAService.gs
 * Replace your existing QAService.gs with this complete implementation
 */

// ============================================================================
// CORE QA CONFIGURATION
// ============================================================================

const ROOT_PROP_KEY = '1GuTbUeWFdp7u6nVDdavc_rCXfUMKX2DH';
const FALLBACK_PATH = ['Lumina', 'QA Uploads'];

function clientGetQualitySummaries(context) {
  try {
    return getEntitySummaries('quality', context);
  } catch (error) {
    console.error('clientGetQualitySummaries failed:', error);
    throw error;
  }
}

function clientGetQualityDetail(id, context) {
  try {
    return getEntityDetail('quality', id, context);
  } catch (error) {
    console.error('clientGetQualityDetail failed:', error);
    throw error;
  }
}

function qaWeights_() {
  return {
    q1: 3, q2: 5, q3: 7, q4: 10, q5: 5, q19: 5,
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
    q18: 'Did the customer respond positively to the survey?',
    q19: 'Was the call free of unnecessary dead air?'
  };
}

function getUserEmail(agentName) {
  try {
    console.log('getUserEmail called with:', agentName);

    if (!agentName || typeof agentName !== 'string') {
      return '';
    }

    const cleanName = agentName.trim();

    // Method 1: Try getUsers function with improved matching
    try {
      const users = (typeof getUsers === 'function') ? getUsers() : [];
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

    // Method 2: Try QA Records as last resort
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

    try {
      const users = (typeof getUsers === 'function') ? getUsers() : [];
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
  for (let i = 1; i <= 19; i++) {
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
      for (let i = 1; i <= 19; i++) {
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
          // Handle Q1-Q19 answers and their note columns
          const questionMatch = col.match(/^Q(\d+)$/i);
          if (questionMatch) {
            const qKey = ('q' + questionMatch[1]).toLowerCase();
            const val = data[qKey];
            if (!val || val === 'na') return 'N/A';
            return val === 'yes' ? 'Yes' : 'No';
          }
          const noteMatch = col.match(/^Q(\d+)\s+Note$/i);
          if (noteMatch) {
            const number = noteMatch[1];
            const camelKey = 'q' + number + 'Note';
            const lowerKey = ('q' + number + 'note');
            const legacyKey = 'c' + number;
            return data[camelKey] || data[lowerKey] || data[legacyKey] || '';
          }
          if (/^C\d+$/i.test(col)) {
            const number = col.replace(/[^0-9]/g, '');
            const camelKey = number ? ('q' + number + 'Note') : '';
            const lowerKey = number ? ('q' + number + 'note') : '';
            const legacyKey = col.toLowerCase();
            return (camelKey && data[camelKey]) || (lowerKey && data[lowerKey]) || data[legacyKey] || '';
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

// ============================================================================
// AI INTELLIGENCE PIPELINE FOR QA DASHBOARD
// ============================================================================

const QA_INTEL_PASS_MARK = 0.95;
const QA_INTEL_PASS_SCORE_THRESHOLD = Math.round(QA_INTEL_PASS_MARK * 100);

function clientGetQAIntelligence(request = {}) {
  try {
    const rawRecords = getAllQA();
    const normalization = normalizeIntelligenceRequest_(request, rawRecords) || {};
    const context = normalization.context || {
      granularity: 'Week',
      period: '',
      timezone: Session.getScriptTimeZone(),
      filters: { agent: '', campaignId: '', program: '' },
      depth: 6,
      agentUniverse: null,
      passMark: QA_INTEL_PASS_MARK
    };
    const normalizedRecords = Array.isArray(normalization.records)
      ? normalization.records
      : [];

    const cache = getQAIntelligenceCache_();
    const cacheKey = cache ? getQAIntelligenceCacheKey_(context) : '';

    if (cache && cacheKey) {
      const cachedPayload = cache.get(cacheKey);
      if (cachedPayload) {
        try {
          const cached = JSON.parse(cachedPayload);
          if (cached && cached.intelligence && cached.intelligence.meta) {
            cached.intelligence.meta.cache = 'hit';
          }
          return cached;
        } catch (parseError) {
          console.warn('Unable to parse cached QA intelligence payload:', parseError);
        }
      }
    }

    const filtered = filterRecordsForIntelligence_(normalizedRecords, context);
    const previousContext = { ...context, period: getPreviousPeriod_(context.granularity, context.period) };
    const prevFiltered = previousContext.period
      ? filterRecordsForIntelligence_(normalizedRecords, previousContext)
      : [];

    const categoryMetrics = computeCategoryMetrics_(filtered);
    const prevCategoryMetrics = computeCategoryMetrics_(prevFiltered);

    const kpis = computeKpiSummary_(filtered, {
      previous: prevFiltered,
      agentUniverse: context.agentUniverse,
      allAgents: normalizedRecords.map(r => r.agent).filter(Boolean)
    });

    const trendSeries = buildTrendSeries_(context, normalizedRecords);
    const trendAnalysis = analyzeTrendSeries_(trendSeries, { granularity: context.granularity });

    const intelligence = buildAIIntelligenceAnalysis_({
      filtered,
      prevFiltered,
      categoryMetrics,
      prevCategoryMetrics,
      kpis,
      granularity: context.granularity
    });

    const generatedAt = new Date().toISOString();
    if (intelligence && intelligence.meta) {
      intelligence.meta.generatedAt = generatedAt;
    }

    const response = {
      generatedAt,
      context,
      kpis,
      intelligence,
      trend: {
        granularity: context.granularity,
        series: trendSeries,
        analysis: trendAnalysis
      }
    };

    if (cache && cacheKey) {
      try {
        if (response.intelligence && response.intelligence.meta) {
          response.intelligence.meta.cache = 'miss';
        }
        cache.put(cacheKey, JSON.stringify(response), 300);
      } catch (cacheError) {
        console.warn('Unable to cache QA intelligence payload:', cacheError);
      }
    }

    return response;
  } catch (error) {
    console.error('clientGetQAIntelligence failed:', error);
    writeError('clientGetQAIntelligence', error);
    throw error;
  }
}

function clientGetQADashboardSnapshot(request = {}) {
  try {
    const rawRecords = getAllQA();
    const normalization = normalizeIntelligenceRequest_(request, rawRecords) || {};
    const context = normalization.context || {
      granularity: 'Week',
      period: '',
      timezone: Session.getScriptTimeZone(),
      filters: { agent: '', campaignId: '', program: '' },
      depth: 6,
      agentUniverse: null,
      passMark: QA_INTEL_PASS_MARK
    };

    const records = Array.isArray(normalization.records)
      ? normalization.records
      : [];

    const filtered = filterRecordsForIntelligence_(records, context);
    const previousPeriod = getPreviousPeriod_(context.granularity, context.period);
    const previousContext = { ...context, period: previousPeriod };
    const prevFiltered = previousPeriod
      ? filterRecordsForIntelligence_(records, previousContext)
      : [];

    const universeOptions = {
      agentUniverse: context.agentUniverse,
      allAgents: records.map(record => record.agent).filter(Boolean)
    };

    const kpis = computeKpiSummary_(filtered, universeOptions);
    const prevKpis = previousPeriod
      ? computeKpiSummary_(prevFiltered, universeOptions)
      : null;

    const trendSeries = buildTrendSeries_(context, records);
    const trendAnalysis = analyzeTrendSeries_(trendSeries, { granularity: context.granularity });

    const categoryMetrics = computeCategoryMetrics_(filtered);
    const prevCategoryMetrics = computeCategoryMetrics_(prevFiltered);
    const categorySummary = summarizeCategoryChange_(categoryMetrics, prevCategoryMetrics);

    const questionMetrics = computeQuestionPerformance_(filtered.length ? filtered : records);
    const questionSignals = buildQuestionSignalHighlights_(questionMetrics);

    const programMetrics = computeProgramMetrics_(filtered.length ? filtered : records, {
      totalRecords: records.length
    });

    const agentDisplayLookup = buildAgentDisplayLookup_(records);

    const { profiles } = calculateAgentProfiles_(filtered, { displayLookup: agentDisplayLookup });
    const { profiles: prevProfiles } = calculateAgentProfiles_(prevFiltered, { displayLookup: agentDisplayLookup });
    const prevProfileLookup = {};
    prevProfiles.forEach(profile => {
      prevProfileLookup[profile.id || profile.name] = profile;
    });

    const timezone = context.timezone || Session.getScriptTimeZone();
    const agents = profiles.map(profile => {
      const key = profile.id || profile.name;
      const previous = prevProfileLookup[key] || null;
      return {
        id: profile.id || profile.name,
        name: profile.name,
        displayName: profile.displayName || profile.name,
        avgScore: profile.avgScore,
        passRate: profile.passRate,
        evaluations: profile.evaluations,
        evaluationShare: profile.evaluationShare,
        recentDate: profile.recentDate
          ? Utilities.formatDate(profile.recentDate, timezone, 'yyyy-MM-dd')
          : '',
        deltas: {
          avgScore: previous ? roundOneDecimal_(profile.avgScore - previous.avgScore) : null,
          passRate: previous ? roundOneDecimal_(profile.passRate - previous.passRate) : null,
          evaluations: previous ? (profile.evaluations - previous.evaluations) : null
        }
      };
    });

    const intelligence = buildAIIntelligenceAnalysis_({
      filtered,
      prevFiltered,
      categoryMetrics,
      prevCategoryMetrics,
      kpis,
      granularity: context.granularity,
      agentLookup: agentDisplayLookup,
      questionMetrics,
      programMetrics,
      trendAnalysis
    });

    const qualitySignals = buildQualitySignals_({
      questionSignals,
      programMetrics,
      kpis,
      trendAnalysis
    });

    const summary = Object.assign({}, kpis, {
      previous: prevKpis,
      delta: {
        avg: computeDelta_(kpis.avg, prevKpis ? prevKpis.avg : null),
        pass: computeDelta_(kpis.pass, prevKpis ? prevKpis.pass : null),
        coverage: computeDelta_(kpis.coverage, prevKpis ? prevKpis.coverage : null),
        completion: computeDelta_(kpis.completion, prevKpis ? prevKpis.completion : null),
        evaluations: computeDelta_(kpis.evaluations, prevKpis ? prevKpis.evaluations : null),
        agents: computeDelta_(kpis.agents, prevKpis ? prevKpis.agents : null)
      }
    });

    const periodOptions = trendSeries
      .slice()
      .reverse()
      .map(entry => ({ value: entry.period, label: entry.label }));

    const latestEvaluation = buildLatestEvaluationSummary_(filtered.length ? filtered : records);

    const availableAgents = Array.from(new Set(records.map(record => record.agent).filter(Boolean))).sort();
    const agentNameLookup = {};
    const agentOptions = availableAgents.map(identifier => {
      const label = resolveAgentDisplayNameFromLookup_(identifier, agentDisplayLookup);
      agentNameLookup[identifier] = label;
      return {
        value: identifier,
        label
      };
    });

    return {
      success: true,
      context,
      summary,
      trend: {
        series: trendSeries,
        analysis: trendAnalysis
      },
      categories: categorySummary,
      agents,
      insights: (intelligence && intelligence.insights) ? intelligence.insights : [],
      actions: (intelligence && intelligence.actions) ? intelligence.actions : [],
      nextBest: intelligence ? intelligence.nextBest : null,
      intelligenceSummary: intelligence ? intelligence.summary : '',
      qualitySignals,
      metadata: {
        generatedAt: new Date().toISOString(),
        totalRecords: records.length
      },
      periodOptions,
      latestEvaluation,
      availableAgents,
      agentOptions,
      agentNameLookup,
      questionSignals,
      programMetrics
    };
  } catch (error) {
    console.error('clientGetQADashboardSnapshot failed:', error);
    writeError('clientGetQADashboardSnapshot', error);
    return {
      success: false,
      error: error && error.message ? error.message : 'Unable to build QA dashboard snapshot.'
    };
  }
}

function normalizeIntelligenceRequest_(request, rawRecords) {
  const granularity = request && typeof request.granularity === 'string'
    ? request.granularity
    : 'Week';

  const timezone = typeof request.timezone === 'string' && request.timezone
    ? request.timezone
    : Session.getScriptTimeZone();

  const agentUniverse = Number(request.agentUniverse) > 0
    ? Number(request.agentUniverse)
    : null;

  const filters = {
    agent: (request.agent || '').toString().trim(),
    campaignId: (request.campaignId || request.campaign || '').toString().trim(),
    program: (request.program || '').toString().trim()
  };

  const depth = Number(request.depth) > 0 ? Math.min(Number(request.depth), 12) : 6;

  const passMark = typeof request.passMark === 'number' ? request.passMark : QA_INTEL_PASS_MARK;

  const normalizedRecords = (rawRecords || [])
    .map(record => normalizeQaRecord_(record, timezone, passMark))
    .filter(record => record.callDate instanceof Date);

  let period = (request && request.period) ? String(request.period) : '';
  if (!period) {
    period = determineLatestPeriod_(granularity, normalizedRecords);
  }

  return {
    context: {
      granularity,
      period,
      timezone,
      filters,
      depth,
      agentUniverse,
    passMark
  },
  records: normalizedRecords
  };
}

function getQAIntelligenceCache_() {
  try {
    return CacheService.getScriptCache();
  } catch (error) {
    console.warn('QA intelligence cache unavailable:', error);
    return null;
  }
}

function getQAIntelligenceCacheKey_(context) {
  if (!context || !context.period) {
    return '';
  }

  try {
    const filters = context.filters || {};
    const parts = [
      context.granularity || '',
      context.period || '',
      filters.agent || '',
      filters.campaignId || '',
      filters.program || '',
      context.agentUniverse || '',
      context.depth || '',
      context.passMark || '',
      context.timezone || ''
    ];

    const encoded = parts
      .map(part => encodeURIComponent(String(part || '')))
      .join('|');

    const key = `qa-intel:${encoded}`;
    return key.length > 230 ? key.substring(0, 230) : key;
  } catch (error) {
    console.warn('Unable to build QA intelligence cache key:', error);
    return '';
  }
}

function normalizeQaRecord_(record, timezone, passMarkOverride) {
  const entry = Object.assign({}, record);

  const agentValue = getRecordFieldValue_(entry, ['AgentName', 'Agent Name', 'Agent', 'AgentEmail', 'Agent Email', 'Associate']);
  const campaignValue = getRecordFieldValue_(entry, ['Campaign', 'Campaign Name', 'Program', 'Program Name', 'Line Of Business', 'LineOfBusiness', 'LOB']);
  const dateValue = getRecordFieldValue_(entry, ['CallDate', 'Call Date', 'CallTime', 'Call Time', 'EvaluationDate', 'Evaluation Date', 'QA Date', 'Date', 'Timestamp']);
  const percentageValue = getRecordFieldValue_(entry, ['Percentage', 'QA Score', 'QA%', 'QA %', 'Final Score', 'FinalScore', 'Score', 'Overall Score']);

  const agent = agentValue ? String(agentValue).trim() : 'Unassigned';
  const campaign = campaignValue ? String(campaignValue).trim() : '';
  const callDate = safeToDate_(dateValue);
  const percentage = parsePercentageValue_(percentageValue);
  const recordScore = Math.round(clamp01_(percentage) * 100);

  const passThreshold = typeof passMarkOverride === 'number' ? passMarkOverride : QA_INTEL_PASS_MARK;

  const tz = timezone || Session.getScriptTimeZone();
  const callDateIso = callDate instanceof Date
    ? Utilities.formatDate(callDate, tz, "yyyy-MM-dd'T'HH:mm:ssXXX")
    : '';

  return {
    raw: entry,
    agent,
    campaign,
    callDate,
    callDateIso,
    percentage,
    recordScore,
    pass: percentage >= passThreshold,
    week: callDate instanceof Date ? toISOWeek_(callDate) : '',
    month: callDate instanceof Date ? formatMonthKey_(callDate) : '',
    quarter: callDate instanceof Date ? `${getQuarter_(callDate)}-${callDate.getFullYear()}` : '',
    year: callDate instanceof Date ? String(callDate.getFullYear()) : ''
  };
}

function safeToDate_(value) {
  return coerceDateValue_(value);
}

function determineLatestPeriod_(granularity, records) {
  if (!records || !records.length) return '';
  const sorted = records.slice().sort((a, b) => b.callDate - a.callDate);
  const latest = sorted[0];
  switch (granularity) {
    case 'Week':
      return latest.week;
    case 'Month':
      return latest.month;
    case 'Quarter':
      return latest.quarter;
    case 'Year':
      return latest.year;
    default:
      return latest.week;
  }
}

function filterRecordsForIntelligence_(records, context) {
  const { filters, granularity, period } = context;
  return (records || []).filter(record => {
    if (filters.agent && record.agent !== filters.agent) return false;
    if (filters.campaignId && record.campaign !== filters.campaignId) return false;
    if (filters.program && record.raw && record.raw.Program !== filters.program) return false;

    if (!period) return true;

    switch (granularity) {
      case 'Week':
        return record.week === period;
      case 'Month':
        return record.month === period;
      case 'Quarter':
        return record.quarter === period;
      case 'Year':
        return record.year === period;
      default:
        return true;
    }
  });
}

function computeKpiSummary_(records, options) {
  const total = records.length;
  const averageScore = total
    ? Math.round((records.reduce((sum, record) => sum + record.percentage, 0) / total) * 100)
    : 0;
  const passCount = records.filter(record => record.pass).length;
  const passRate = total ? Math.round((passCount / total) * 100) : 0;

  const uniqueAgents = new Set(records.map(record => record.agent).filter(Boolean));
  const agentUniverse = options && options.agentUniverse
    ? Number(options.agentUniverse)
    : new Set((options && options.allAgents) || []).size;

  const coverage = agentUniverse
    ? Math.min(Math.round((uniqueAgents.size / agentUniverse) * 100), 100)
    : (uniqueAgents.size > 0 ? 100 : 0);

  const completion = uniqueAgents.size
    ? Math.min(Math.round((total / uniqueAgents.size) * 10), 100)
    : 0;

  return {
    avg: averageScore,
    pass: passRate,
    coverage,
    completion,
    evaluations: total,
    agents: uniqueAgents.size
  };
}

function computeDelta_(current, previous) {
  if (typeof current !== 'number' || typeof previous !== 'number' || Number.isNaN(current) || Number.isNaN(previous)) {
    return null;
  }

  const delta = current - previous;
  if (Number.isInteger(current) && Number.isInteger(previous)) {
    return delta;
  }

  return Math.round(delta * 10) / 10;
}

function roundOneDecimal_(value) {
  if (typeof value !== 'number' || Number.isNaN(value)) {
    return null;
  }

  return Math.round(value * 10) / 10;
}

function buildTrendSeries_(context, records) {
  const { granularity, period, depth } = context;
  const series = [];
  const visited = new Set();
  let cursor = period;
  let steps = 0;

  while (cursor && steps < depth && !visited.has(cursor)) {
    visited.add(cursor);
    const bucket = filterRecordsForIntelligence_(records, { ...context, period: cursor });
    const evalCount = bucket.length;
    const agentCount = new Set(bucket.map(r => r.agent).filter(Boolean)).size;
    const avgScore = evalCount
      ? Math.round((bucket.reduce((sum, r) => sum + r.percentage, 0) / evalCount) * 100)
      : 0;
    const passRate = evalCount
      ? Math.round((bucket.filter(r => r.pass).length / evalCount) * 100)
      : 0;
    const coverage = context.agentUniverse
      ? Math.min(Math.round((agentCount / context.agentUniverse) * 100), 100)
      : (agentCount > 0 ? 100 : 0);

    series.push({
      period: cursor,
      label: formatPeriodLabel_(granularity, cursor),
      avgScore,
      passRate,
      evalCount,
      agentCount,
      coverage
    });

    cursor = getPreviousPeriod_(granularity, cursor);
    steps += 1;
  }

  return series.reverse();
}

function linearRegression_(points) {
  if (!points || !points.length) {
    return { slope: 0, intercept: 0 };
  }

  if (points.length === 1) {
    return { slope: 0, intercept: points[0].y };
  }

  const n = points.length;
  let sumX = 0;
  let sumY = 0;
  let sumXY = 0;
  let sumXX = 0;

  points.forEach(point => {
    sumX += point.x;
    sumY += point.y;
    sumXY += point.x * point.y;
    sumXX += point.x * point.x;
  });

  const denominator = (n * sumXX) - (sumX * sumX);
  if (denominator === 0) {
    return { slope: 0, intercept: sumY / n };
  }

  const slope = ((n * sumXY) - (sumX * sumY)) / denominator;
  const intercept = (sumY - slope * sumX) / n;
  return { slope, intercept };
}

function analyzeTrendSeries_(series, context) {
  const granularity = context && context.granularity ? context.granularity : 'Period';
  const lowerGran = granularity.toLowerCase();

  if (!series || !series.length) {
    return {
      summary: `Lumina AI is waiting for enough history to analyze ${lowerGran} trends.`,
      points: [],
      health: 'monitoring',
      forecast: { avg: 0, pass: 0 },
      nextLabel: `next ${lowerGran}`
    };
  }

  const first = series[0];
  const last = series[series.length - 1];

  const avgPoints = series.map((point, index) => ({ x: index, y: point.avgScore }));
  const passPoints = series.map((point, index) => ({ x: index, y: point.passRate }));

  const avgReg = linearRegression_(avgPoints);
  const passReg = linearRegression_(passPoints);

  const avgDelta = last.avgScore - first.avgScore;
  const passDelta = last.passRate - first.passRate;
  const volumeDelta = last.evalCount - first.evalCount;

  const slopeAvg = avgReg.slope;
  const slopePass = passReg.slope;

  const improving = slopeAvg > 0.5 || slopePass > 0.5;
  const declining = slopeAvg < -0.5 || slopePass < -0.5;

  let health = 'stable';
  if (improving) health = 'improving';
  if (declining) health = 'risk';

  const summaryParts = [];
  summaryParts.push(`Average quality is ${avgDelta >= 0 ? 'up' : 'down'} ${Math.abs(avgDelta).toFixed(1)} pts`);
  summaryParts.push(`pass rate ${passDelta >= 0 ? 'gained' : 'slid'} ${Math.abs(passDelta).toFixed(1)} pts`);
  summaryParts.push(`${last.evalCount} evaluations this ${lowerGran}`);

  const points = [];

  points.push({
    icon: improving ? 'fa-arrow-up' : declining ? 'fa-arrow-down' : 'fa-arrows-alt-h',
    tone: improving ? 'positive' : declining ? 'negative' : '',
    title: `Average score ${improving ? 'rising' : declining ? 'dropping' : 'steady'}`,
    text: `${first.avgScore}% → ${last.avgScore}% across the last ${series.length} ${series.length === 1 ? lowerGran : lowerGran + 's'}.`
  });

  points.push({
    icon: passDelta >= 0 ? 'fa-shield-alt' : 'fa-exclamation-triangle',
    tone: passDelta >= 0 ? 'positive' : 'negative',
    title: `Pass rate ${passDelta >= 0 ? 'improving' : 'at risk'}`,
    text: `${first.passRate}% → ${last.passRate}% (${passDelta >= 0 ? '+' : ''}${passDelta.toFixed(1)} pts).`
  });

  if (Math.abs(volumeDelta) > 0) {
    points.push({
      icon: volumeDelta >= 0 ? 'fa-users' : 'fa-user-slash',
      tone: volumeDelta >= 0 ? 'positive' : 'negative',
      title: `Evaluation volume ${volumeDelta >= 0 ? 'growing' : 'contracting'}`,
      text: `${first.evalCount} → ${last.evalCount} evaluations (${volumeDelta >= 0 ? '+' : ''}${volumeDelta}).`
    });
  } else {
    points.push({
      icon: 'fa-stopwatch',
      tone: '',
      title: 'Volume steady',
      text: `Evaluation count steady at ${last.evalCount} per ${lowerGran}.`
    });
  }

  if (last.coverage < 80) {
    points.push({
      icon: 'fa-user-shield',
      tone: 'negative',
      title: 'Coverage gap detected',
      text: `Only ${last.coverage}% of agents covered in the latest ${lowerGran}.`
    });
  }

  const forecastAvg = clampPercent_(avgReg.intercept + avgReg.slope * avgPoints.length);
  const forecastPass = clampPercent_(passReg.intercept + passReg.slope * passPoints.length);

  return {
    summary: `${summaryParts.join(', ')}.`,
    points,
    health,
    forecast: { avg: forecastAvg, pass: forecastPass },
    nextLabel: `next ${lowerGran}`
  };
}

function clampPercent_(value) {
  if (Number.isNaN(value)) return 0;
  return Math.max(0, Math.min(100, Math.round(value)));
}

function buildAIIntelligenceAnalysis_(payload) {
  const {
    filtered = [],
    prevFiltered = [],
    categoryMetrics = {},
    prevCategoryMetrics = {},
    kpis = {},
    granularity,
    agentLookup = {},
    questionMetrics = [],
    programMetrics = [],
    trendAnalysis = null
  } = payload || {};

  const totalEvaluations = filtered.length;
  const periodLabel = granularity ? granularity.toLowerCase() : 'period';

  const confidenceScore = clampPercent_(
    Math.round(
      Math.max(5,
        ((kpis.coverage || 0) * 0.4) +
        ((kpis.pass || 0) * 0.3) +
        ((kpis.avg || 0) * 0.3)
      )
    )
  );

  const base = {
    summary: '',
    automationSummary: '',
    confidence: confidenceScore,
    automationState: 'Monitoring',
    insights: [],
    actions: [],
    nextBest: null,
    meta: {
      totalEvaluations,
      periodLabel,
      source: 'server',
      generatedAt: new Date().toISOString()
    }
  };

  if (!totalEvaluations) {
    return {
      ...base,
      summary: 'Lumina AI is monitoring for new evaluations. Adjust your filters or capture fresh QA reviews to generate insights.',
      automationSummary: 'No automation required yet. Log additional evaluations to unlock targeted recommendations.'
    };
  }

  const { profiles } = calculateAgentProfiles_(filtered, { displayLookup: agentLookup });
  const { profiles: prevProfiles } = calculateAgentProfiles_(prevFiltered, { displayLookup: agentLookup });
  const categorySummary = summarizeCategoryChange_(categoryMetrics, prevCategoryMetrics);

  const totalAgents = profiles.length;
  base.meta.totalAgents = totalAgents;

  base.summary = `AI reviewed ${totalEvaluations} ${totalEvaluations === 1 ? 'evaluation' : 'evaluations'} across ${totalAgents} ${totalAgents === 1 ? 'agent' : 'agents'} for this ${periodLabel}, spotlighting performance opportunities instantly.`;
  base.automationSummary = `Coverage at ${clampPercent_(kpis.coverage || 0)}% and completion at ${clampPercent_(kpis.completion || 0)}% give AI enough signal to trigger proactive workflows.`;

  if (profiles.length) {
    const topAgent = profiles[0];
    base.insights.push({
      icon: 'fa-star',
      tone: 'positive',
      title: `${topAgent.name} is leading`,
      text: `${topAgent.name} is averaging ${topAgent.avgScore}% quality with a ${topAgent.passRate}% pass rate.`
    });

    const bottomAgent = profiles[profiles.length - 1];
    if (bottomAgent && bottomAgent.avgScore < QA_INTEL_PASS_SCORE_THRESHOLD) {
      base.insights.push({
        icon: 'fa-life-ring',
        tone: 'negative',
        title: `${bottomAgent.name} needs attention`,
        text: `${bottomAgent.name} is trending at ${bottomAgent.avgScore}% with ${bottomAgent.passRate}% pass rate.`
      });
      base.actions.push({
        icon: 'fa-user-graduate',
        tone: 'urgent',
        title: `Launch coaching for ${bottomAgent.name}`,
        text: `Auto-create a coaching session to lift ${bottomAgent.name}'s quality score back above ${QA_INTEL_PASS_SCORE_THRESHOLD}%.`
      });
    }

    const prevProfileMap = {};
    prevProfiles.forEach(profile => {
      const key = profile.id || profile.name;
      prevProfileMap[key] = profile;
    });

    let strongestImprovement = null;
    let largestRegression = null;

    profiles.forEach(profile => {
      const key = profile.id || profile.name;
      const prev = prevProfileMap[key];
      if (!prev) return;
      const delta = profile.avgScore - prev.avgScore;
      if (strongestImprovement === null || delta > strongestImprovement.delta) {
        strongestImprovement = { ...profile, delta };
      }
      if (largestRegression === null || delta < largestRegression.delta) {
        largestRegression = { ...profile, delta };
      }
    });

    if (strongestImprovement && strongestImprovement.delta > 2) {
      base.insights.push({
        icon: 'fa-rocket',
        tone: 'positive',
        title: `${strongestImprovement.name} is improving`,
        text: `Up ${strongestImprovement.delta.toFixed(1)} pts vs last period.`
      });
    }

    if (largestRegression && largestRegression.delta < -2) {
      base.actions.push({
        icon: 'fa-reply',
        tone: 'urgent',
        title: `Check-in with ${largestRegression.name}`,
        text: `${largestRegression.name} dropped ${Math.abs(largestRegression.delta).toFixed(1)} pts period-over-period.`
      });
    }
  }

  if (categorySummary.length) {
    const bestCategory = categorySummary[0];
    base.insights.push({
      icon: 'fa-thumbs-up',
      tone: 'positive',
      title: `${bestCategory.category} excels`,
      text: `${bestCategory.category} is averaging ${bestCategory.avgScore}% quality.`
    });

    const weakestCategory = categorySummary[categorySummary.length - 1];
    if (weakestCategory && weakestCategory.avgScore < QA_INTEL_PASS_SCORE_THRESHOLD) {
      base.actions.push({
        icon: 'fa-sitemap',
        tone: 'urgent',
        title: `Reinforce ${weakestCategory.category}`,
        text: `Automate a calibration focused on ${weakestCategory.category} where scores average ${weakestCategory.avgScore}%.`
      });
    }

    const largestDelta = categorySummary.reduce((acc, entry) => {
      if (entry.delta === null) return acc;
      if (!acc || entry.delta < acc.delta) return entry;
      return acc;
    }, null);

    if (largestDelta && largestDelta.delta < -3) {
      base.actions.push({
        icon: 'fa-exclamation-circle',
        tone: 'urgent',
        title: `Reverse slide in ${largestDelta.category}`,
        text: `${largestDelta.category} fell ${Math.abs(largestDelta.delta).toFixed(1)} pts from the previous period.`
      });
    }
  }

  if ((kpis.pass || 0) < 90) {
    base.actions.push({
      icon: 'fa-headset',
      tone: 'urgent',
      title: 'Boost pass rate',
      text: `Configure an automated refresher for agents with pass rates below 90%. Current pass rate is ${clampPercent_(kpis.pass || 0)}%.`
    });
  }

  if ((kpis.coverage || 0) < 85) {
    base.actions.push({
      icon: 'fa-user-check',
      tone: 'urgent',
      title: 'Increase agent coverage',
      text: `Auto-assign additional evaluations to reach at least 90% agent coverage. Currently at ${clampPercent_(kpis.coverage || 0)}%.`
    });
  }

  const highImpactQuestion = Array.isArray(questionMetrics)
    ? questionMetrics.find(metric => metric.severity && metric.severity !== 'positive')
    : null;

  if (highImpactQuestion) {
    base.actions.push({
      icon: 'fa-triangle-exclamation',
      tone: 'urgent',
      title: `Stabilize ${highImpactQuestion.shortLabel}`,
      text: `${highImpactQuestion.passRate}% pass with ${highImpactQuestion.noCount} misses. Launch calibration or targeted coaching.`
    });

    if (highImpactQuestion.primaryNote) {
      base.insights.push({
        icon: 'fa-microphone-lines',
        tone: 'negative',
        title: `${highImpactQuestion.shortLabel} risk driver`,
        text: highImpactQuestion.primaryNote
      });
    }
  }

  const atRiskProgram = Array.isArray(programMetrics)
    ? programMetrics.find(program => program.severity && program.severity !== 'positive')
    : null;

  if (atRiskProgram) {
    base.actions.push({
      icon: 'fa-diagram-project',
      tone: 'urgent',
      title: `Stabilize ${atRiskProgram.name}`,
      text: `${atRiskProgram.passRate}% pass across ${atRiskProgram.evaluations} evaluations. Align QA and operations immediately.`
    });

    base.insights.push({
      icon: 'fa-network-wired',
      tone: 'negative',
      title: `${atRiskProgram.name} underperforming`,
      text: `Average score ${atRiskProgram.avgScore}% with ${atRiskProgram.agentCoverage} agents impacted.`
    });
  }

  if (trendAnalysis && trendAnalysis.forecast) {
    base.insights.push({
      icon: 'fa-chart-line',
      tone: trendAnalysis.health === 'risk' ? 'negative' : trendAnalysis.health === 'improving' ? 'positive' : '',
      title: `Forecast ${trendAnalysis.health === 'risk' ? 'signals risk' : trendAnalysis.health === 'improving' ? 'shows lift' : 'steady'}`,
      text: `Projected avg ${trendAnalysis.forecast.avg}% and pass ${trendAnalysis.forecast.pass}% next ${trendAnalysis.nextLabel}.`
    });
  }

  if (!base.insights.length) {
    base.insights.push({
      icon: 'fa-lightbulb',
      tone: 'positive',
      title: 'All clear',
      text: 'No critical anomalies detected. AI will notify if trends change.'
    });
  }

  base.automationState = base.actions.length ? 'Action Required' : 'Monitoring';
  base.nextBest = base.actions.length ? base.actions[0] : null;

  return base;
}

function buildQualitySignals_(payload = {}) {
  const {
    questionSignals = [],
    programMetrics = [],
    kpis = {},
    trendAnalysis = null
  } = payload;

  const signals = [];

  const primaryQuestion = questionSignals.find(signal => signal.severity !== 'positive');
  if (primaryQuestion) {
    signals.push({
      icon: 'fa-circle-exclamation',
      tone: primaryQuestion.severity === 'negative' ? 'negative' : 'warning',
      title: `${primaryQuestion.shortLabel} at ${primaryQuestion.passRate}%`,
      text: `${primaryQuestion.noCount} negative responses out of ${primaryQuestion.totalResponses}.`
    });
  }

  const riskProgram = programMetrics.find(program => program.severity === 'negative');
  if (riskProgram) {
    signals.push({
      icon: 'fa-sitemap',
      tone: 'negative',
      title: `${riskProgram.name} degradation`,
      text: `Average ${riskProgram.avgScore}% quality with ${riskProgram.passRate}% pass rate.`
    });
  }

  if (typeof kpis.coverage === 'number' && kpis.coverage < 80) {
    signals.push({
      icon: 'fa-user-shield',
      tone: 'warning',
      title: 'Coverage below 80%',
      text: `Only ${kpis.coverage}% of the agent population has been evaluated.`
    });
  }

  if (trendAnalysis && trendAnalysis.health === 'risk') {
    signals.push({
      icon: 'fa-arrow-trend-down',
      tone: 'negative',
      title: 'Declining trajectory',
      text: trendAnalysis.summary
    });
  }

  return signals;
}

function calculateAgentProfiles_(records, options = {}) {
  const totalEvaluations = records.length;
  const aggregates = {};
  const displayLookup = options.displayLookup || {};

  records.forEach(record => {
    const name = record.agent || 'Unassigned';
    if (!aggregates[name]) {
      aggregates[name] = {
        count: 0,
        scoreSum: 0,
        passCount: 0,
        recent: null,
        displayName: ''
      };
    }

    const bucket = aggregates[name];
    bucket.count += 1;
    bucket.scoreSum += record.recordScore;
    if (record.pass) {
      bucket.passCount += 1;
    }

    if (record.callDate instanceof Date) {
      if (!bucket.recent || record.callDate > bucket.recent) {
        bucket.recent = record.callDate;
      }
    }

    const candidateName = inferAgentDisplayNameFromRecord_(record, displayLookup);
    if (candidateName) {
      if (!bucket.displayName || bucket.displayName === name || bucket.displayName === prettifyAgentIdentifier_(name)) {
        bucket.displayName = candidateName;
      }
    }
  });

  const profiles = Object.keys(aggregates).map(name => {
    const stats = aggregates[name];
    const avgScore = stats.count ? Math.round(stats.scoreSum / stats.count) : 0;
    const passRate = stats.count ? Math.round((stats.passCount / stats.count) * 100) : 0;
    const evaluationShare = totalEvaluations ? Math.round((stats.count / totalEvaluations) * 100) : 0;
    const displayName = resolveAgentDisplayNameForIdentifier_(name, stats, displayLookup);

    return {
      id: name,
      name: displayName,
      displayName,
      rawName: name,
      evaluations: stats.count,
      avgScore,
      passRate,
      evaluationShare,
      recentDate: stats.recent || null
    };
  }).sort((a, b) => b.avgScore - a.avgScore);

  return { totalEvaluations, profiles };
}

function buildAgentDisplayLookup_(records) {
  const lookup = {};

  const assign = (key, display) => {
    const normalizedKey = normalizeAgentKey_(key);
    const cleanedDisplay = display ? String(display).trim() : '';
    if (!normalizedKey || !cleanedDisplay) {
      return;
    }
    const existing = lookup[normalizedKey];
    if (!existing || existing === prettifyAgentIdentifier_(key) || existing === String(key || '').trim()) {
      lookup[normalizedKey] = cleanedDisplay;
    }
  };

  try {
    const users = (typeof getUsers === 'function') ? getUsers() : [];
    if (Array.isArray(users)) {
      users.forEach(user => {
        const display = String(user.FullName || user.UserName || user.Email || '').trim();
        if (!display) {
          return;
        }
        [user.ID, user.UserName, user.Email].forEach(candidate => assign(candidate, display));
      });
    }
  } catch (directoryError) {
    console.warn('Unable to load user directory for QA dashboard:', directoryError);
  }

  (records || []).forEach(record => {
    if (!record || !record.raw) {
      return;
    }
    const raw = record.raw;
    const displayCandidate = getRecordFieldValue_(raw, [
      'AgentName',
      'Agent Name',
      'AgentFullName',
      'Agent Full Name',
      'Associate Name',
      'Associate'
    ]);
    const displayName = displayCandidate ? String(displayCandidate).trim() : '';
    if (!displayName) {
      return;
    }

    assign(record.agent, displayName);
    const emailCandidate = getRecordFieldValue_(raw, ['AgentEmail', 'Agent Email', 'Email']);
    assign(emailCandidate, displayName);
    const idCandidate = getRecordFieldValue_(raw, ['AgentID', 'Agent Id', 'UserID', 'User Id', 'AgentIdentifier', 'Agent Identifier']);
    assign(idCandidate, displayName);
  });

  return lookup;
}

function inferAgentDisplayNameFromRecord_(record, lookup) {
  if (!record) {
    return '';
  }
  const raw = record.raw || {};
  const direct = getRecordFieldValue_(raw, [
    'AgentName',
    'Agent Name',
    'AgentFullName',
    'Agent Full Name',
    'Associate Name',
    'Associate'
  ]);
  if (direct) {
    const name = String(direct).trim();
    if (name) {
      return name;
    }
  }

  const normalized = normalizeAgentKey_(record.agent);
  if (normalized && lookup && lookup[normalized]) {
    return lookup[normalized];
  }

  return '';
}

function resolveAgentDisplayNameForIdentifier_(identifier, stats, lookup) {
  const resolved = resolveAgentDisplayNameFromLookup_(identifier, lookup);
  if (resolved && resolved !== 'Unassigned') {
    return resolved;
  }
  if (stats && stats.displayName) {
    return stats.displayName;
  }
  if (resolved) {
    return resolved;
  }
  return 'Unassigned';
}

function resolveAgentDisplayNameFromLookup_(identifier, lookup) {
  const normalized = normalizeAgentKey_(identifier);
  if (normalized && lookup && lookup[normalized]) {
    return lookup[normalized];
  }
  const fallback = prettifyAgentIdentifier_(identifier);
  return fallback || 'Unassigned';
}

function normalizeAgentKey_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim().toLowerCase();
}

function prettifyAgentIdentifier_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const raw = String(value).trim();
  if (!raw) {
    return '';
  }
  if (raw.toLowerCase() === 'unassigned') {
    return 'Unassigned';
  }
  if (raw.includes('@')) {
    const local = raw.split('@')[0];
    return capitalizeAgentWords_(local.replace(/[._-]+/g, ' '));
  }
  if (raw.indexOf(' ') === -1 && /[._-]/.test(raw)) {
    return capitalizeAgentWords_(raw.replace(/[._-]+/g, ' '));
  }
  return capitalizeAgentWords_(raw);
}

function capitalizeAgentWords_(value) {
  return String(value || '')
    .split(/\s+/)
    .filter(Boolean)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');
}

function summarizeCategoryChange_(currentMetrics, previousMetrics) {
  const details = Object.keys(currentMetrics || {}).map(category => {
    const metrics = currentMetrics[category] || { avgScore: 0, passPct: 0 };
    const prev = previousMetrics ? previousMetrics[category] : null;
    const delta = prev ? Math.round((metrics.avgScore - prev.avgScore) * 10) / 10 : null;
    return {
      category,
      avgScore: metrics.avgScore,
      passPct: metrics.passPct,
      delta
    };
  });

  details.sort((a, b) => b.avgScore - a.avgScore);
  return details;
}

function computeQuestionPerformance_(records) {
  const questionText = qaQuestionText_();
  const weights = qaWeights_();
  const metrics = [];

  Object.keys(questionText).forEach(key => {
    let yesCount = 0;
    let noCount = 0;
    let naCount = 0;
    const notes = [];

    (records || []).forEach(record => {
      if (!record || !record.raw) {
        return;
      }

      const raw = record.raw;
      const answerRaw = getAnswerValue_(raw, key);
      if (answerRaw === undefined || answerRaw === null || answerRaw === '') {
        return;
      }

      const answer = String(answerRaw).trim().toLowerCase();
      if (answer === 'yes') {
        yesCount += 1;
      } else if (answer === 'no') {
        noCount += 1;
        const note = getQuestionNoteValue_(raw, key);
        if (note) {
          notes.push(String(note));
        }
      } else {
        naCount += 1;
      }
    });

    const total = yesCount + noCount + naCount;
    const passRate = total ? Math.round((yesCount / total) * 100) : 0;
    const failRate = total ? Math.round((noCount / total) * 100) : 0;
    const weight = weights[key] || weights[key.toLowerCase()] || 0;
    const impactScore = Math.round((failRate / 100) * Math.max(weight, 1) * total);
    const cleanedNotes = summarizeNotes_(notes);

    metrics.push({
      key,
      question: questionText[key],
      shortLabel: key.toUpperCase(),
      passRate,
      failRate,
      yesCount,
      noCount,
      naCount,
      totalResponses: total,
      weight,
      impactScore,
      notes: cleanedNotes,
      primaryNote: cleanedNotes.length ? cleanedNotes[0] : '',
      severity: determineSeverityFromRate_(passRate)
    });
  });

  return metrics.sort((a, b) => {
    if (a.severity === b.severity) {
      return b.impactScore - a.impactScore;
    }
    const order = { negative: 2, warning: 1, positive: 0 };
    return (order[b.severity] || 0) - (order[a.severity] || 0);
  });
}

function buildQuestionSignalHighlights_(metrics) {
  if (!Array.isArray(metrics)) {
    return [];
  }

  return metrics
    .filter(metric => metric.totalResponses > 0)
    .slice(0, 6);
}

function summarizeNotes_(notes) {
  if (!Array.isArray(notes) || !notes.length) {
    return [];
  }

  const seen = new Set();
  const result = [];

  notes.forEach(note => {
    const cleaned = String(note || '').trim();
    if (!cleaned) {
      return;
    }
    const normalized = cleaned.toLowerCase();
    if (seen.has(normalized)) {
      return;
    }
    seen.add(normalized);
    const truncated = cleaned.length > 160 ? `${cleaned.slice(0, 157)}…` : cleaned;
    result.push(truncated);
  });

  return result.slice(0, 3);
}

function determineSeverityFromRate_(passRate) {
  if (typeof passRate !== 'number') {
    return 'positive';
  }
  if (passRate >= 92) {
    return 'positive';
  }
  if (passRate >= 85) {
    return 'warning';
  }
  return 'negative';
}

function computeProgramMetrics_(records, options = {}) {
  const aggregates = {};
  const totalUniverse = Number(options.totalRecords) || 0;

  (records || []).forEach(record => {
    if (!record) {
      return;
    }

    const name = resolveProgramNameFromRecord_(record) || 'Unassigned';
    if (!aggregates[name]) {
      aggregates[name] = {
        evaluations: 0,
        scoreSum: 0,
        passCount: 0,
        agents: new Set()
      };
    }

    const bucket = aggregates[name];
    bucket.evaluations += 1;
    bucket.scoreSum += record.percentage || 0;
    if (record.pass) {
      bucket.passCount += 1;
    }
    if (record.agent) {
      bucket.agents.add(record.agent);
    }
  });

  return Object.keys(aggregates).map(name => {
    const bucket = aggregates[name];
    const avgScore = bucket.evaluations
      ? Math.round((bucket.scoreSum / bucket.evaluations) * 100)
      : 0;
    const passRate = bucket.evaluations
      ? Math.round((bucket.passCount / bucket.evaluations) * 100)
      : 0;
    const share = totalUniverse
      ? Math.round((bucket.evaluations / totalUniverse) * 100)
      : 0;

    return {
      name,
      evaluations: bucket.evaluations,
      avgScore,
      passRate,
      agentCoverage: bucket.agents.size,
      share,
      severity: determineSeverityFromRate_(passRate)
    };
  }).sort((a, b) => {
    if (a.severity === b.severity) {
      return b.evaluations - a.evaluations;
    }
    const order = { negative: 2, warning: 1, positive: 0 };
    return (order[b.severity] || 0) - (order[a.severity] || 0);
  }).slice(0, 8);
}

function resolveProgramNameFromRecord_(record) {
  if (!record) {
    return '';
  }

  const raw = record.raw || {};
  const programField = getRecordFieldValue_(raw, [
    'Program',
    'Program Name',
    'ProgramName',
    'Campaign',
    'Campaign Name',
    'Line Of Business',
    'LineOfBusiness',
    'LOB'
  ]);

  if (programField) {
    return String(programField).trim();
  }

  return record.campaign || '';
}

function computeCategoryMetrics_(records) {
  const categories = qaCategories_();
  const weights = qaWeights_();
  const metrics = {};

  Object.keys(categories).forEach(category => {
    const questionKeys = categories[category] || [];
    const scores = [];
    const passes = [];

    records.forEach(record => {
      const { raw } = record;
      const answers = questionKeys.map(key => getAnswerValue_(raw, key));
      const totalWeight = questionKeys.reduce((sum, key) => {
        const normalizedKey = key.toLowerCase();
        return sum + (weights[normalizedKey] || weights[key] || 0);
      }, 0);

      if (!totalWeight) {
        return;
      }

      const earned = questionKeys.reduce((sum, key, index) => {
        const normalizedKey = key.toLowerCase();
        const weight = weights[normalizedKey] || weights[key] || 0;
        const answer = String(answers[index] || '').toLowerCase();
        if (answer === 'yes' || (normalizedKey === 'q17' && answer === 'no')) {
          return sum + weight;
        }
        return sum;
      }, 0);

      const pct = totalWeight ? Math.round((earned / totalWeight) * 100) : 0;
      const pass = answers.every(answer => String(answer || '').toLowerCase() === 'yes');

      scores.push(pct);
      passes.push(pass ? 1 : 0);
    });

    const avgScore = scores.length
      ? Math.round(scores.reduce((sum, value) => sum + value, 0) / scores.length)
      : 0;
    const passPct = passes.length
      ? Math.round((passes.reduce((sum, value) => sum + value, 0) / passes.length) * 100)
      : 0;

    metrics[category] = { avgScore, passPct };
  });

  return metrics;
}

function getAnswerValue_(record, key) {
  if (!record) return '';
  if (key in record) return record[key];
  const upper = key.toUpperCase();
  if (upper in record) return record[upper];
  const lower = key.toLowerCase();
  if (lower in record) return record[lower];
  return '';
}

function getQuestionNoteValue_(record, questionNumber) {
  if (!record) {
    return '';
  }

  const suffix = String(questionNumber || '').replace(/^q/i, '').trim();
  if (!suffix) {
    return '';
  }

  const candidates = [
    `Q${suffix} Note`,
    `q${suffix} Note`,
    `Q${suffix} note`,
    `q${suffix} note`,
    `Q${suffix}Note`,
    `Q${suffix}Notes`,
    `q${suffix}Note`,
    `q${suffix}Notes`,
    `C${suffix}`,
    `c${suffix}`,
    `Note${suffix}`,
    `Notes${suffix}`
  ];

  for (let i = 0; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    if (candidate in record && record[candidate] !== undefined && record[candidate] !== null) {
      return record[candidate];
    }
  }

  return '';
}

function normalizeAnswerDisplay_(value) {
  if (value === null || value === undefined) {
    return 'N/A';
  }

  const text = String(value).trim();
  if (!text) {
    return 'N/A';
  }

  const normalized = text.toLowerCase();
  if (normalized === 'yes') return 'Yes';
  if (normalized === 'no') return 'No';
  if (normalized === 'na' || normalized === 'n/a') return 'N/A';
  return text;
}

function getComparableRecordTimestamp_(record) {
  if (!record) {
    return 0;
  }

  const dates = [];
  if (record.callDate instanceof Date) {
    dates.push(record.callDate.getTime());
  }

  const raw = record.raw || {};
  const dateFields = ['CallDate', 'AuditDate', 'Timestamp', 'CreatedDate', 'CreatedAt', 'UpdatedAt'];
  dateFields.forEach(field => {
    if (!raw[field]) {
      return;
    }
    const parsed = parseFlexibleDateString_(raw[field]);
    if (parsed instanceof Date && !isNaN(parsed.getTime())) {
      dates.push(parsed.getTime());
    } else {
      const direct = Date.parse(String(raw[field]));
      if (!Number.isNaN(direct)) {
        dates.push(direct);
      }
    }
  });

  if (!dates.length) {
    return 0;
  }

  return Math.max.apply(null, dates);
}

function buildLatestEvaluationSummary_(records) {
  if (!Array.isArray(records) || !records.length) {
    return null;
  }

  const withRaw = records.filter(record => record && record.raw && typeof record.raw === 'object');
  if (!withRaw.length) {
    return null;
  }

  const sorted = withRaw.slice().sort((a, b) => getComparableRecordTimestamp_(b) - getComparableRecordTimestamp_(a));
  const latest = sorted[0];
  const raw = latest.raw || {};

  const questionText = qaQuestionText_();
  const weights = qaWeights_();
  const tz = Session.getScriptTimeZone();

  const questions = Object.keys(questionText).map(key => {
    const number = key.replace(/^q/i, '');
    const weight = weights[key] || weights[key.toLowerCase()] || 0;
    const answerValue = getAnswerValue_(raw, key);
    const noteValue = getQuestionNoteValue_(raw, number);

    return {
      key,
      number: `Q${number}`,
      question: questionText[key],
      weight,
      answer: normalizeAnswerDisplay_(answerValue),
      note: noteValue || ''
    };
  });

  const callDate = latest.callDate instanceof Date
    ? Utilities.formatDate(latest.callDate, tz, 'yyyy-MM-dd')
    : (raw.CallDate || '');
  const percentage = raw.Percentage || raw.FinalScore;
  const normalizedScore = typeof latest.recordScore === 'number'
    ? Math.round(latest.recordScore)
    : (percentage !== undefined && percentage !== null && percentage !== ''
        ? Math.round(parsePercentageValue_(percentage) * 100)
        : '');

  return {
    id: raw.ID || raw.Id || raw.id || '',
    agent: raw.AgentName || latest.agent || '',
    auditor: raw.AuditorName || '',
    callDate,
    auditDate: raw.AuditDate || '',
    score: normalizedScore,
    questions
  };
}

function toISOWeek_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(week).padStart(2, '0')}`;
}

function formatMonthKey_(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function getQuarter_(date) {
  return 'Q' + (Math.floor(date.getMonth() / 3) + 1);
}

function normalizeFieldKey_(key) {
  return String(key || '')
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');
}

function getRecordFieldValue_(record, candidates) {
  if (!record || typeof record !== 'object') {
    return null;
  }

  const lookup = {};
  Object.keys(record).forEach(existingKey => {
    const normalized = normalizeFieldKey_(existingKey);
    if (!(normalized in lookup)) {
      lookup[normalized] = existingKey;
    }
  });

  for (let i = 0; i < candidates.length; i += 1) {
    const normalizedKey = normalizeFieldKey_(candidates[i]);
    const actualKey = lookup[normalizedKey];
    if (actualKey && record[actualKey] !== undefined && record[actualKey] !== null && record[actualKey] !== '') {
      return record[actualKey];
    }
  }

  return null;
}

function clamp01_(value) {
  if (!isFinite(value)) {
    return 0;
  }
  if (value < 0) return 0;
  if (value > 1) return 1;
  return value;
}

function parsePercentageValue_(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }

  if (typeof value === 'number' && isFinite(value)) {
    const normalized = value > 1.0001 ? value / 100 : value;
    return clamp01_(normalized);
  }

  const numeric = parseFloat(String(value).replace(/[^0-9.\-]/g, ''));
  if (!isFinite(numeric)) {
    return 0;
  }

  const normalized = numeric > 1.0001 ? numeric / 100 : numeric;
  return clamp01_(normalized);
}

function excelSerialToDate_(serial) {
  if (typeof serial !== 'number' || !isFinite(serial)) {
    return null;
  }

  if (serial <= 60) {
    return null;
  }

  const utcDays = Math.floor(serial - 25569);
  const utcMilliseconds = utcDays * 86400000;
  const remainder = serial - Math.floor(serial);
  const remainderMs = Math.round(remainder * 86400000);
  const date = new Date(utcMilliseconds + remainderMs);
  return isNaN(date.getTime()) ? null : date;
}

function parseFlexibleDateString_(raw) {
  if (!raw) {
    return null;
  }

  const value = String(raw).trim();
  if (!value) {
    return null;
  }

  if (/^\d+(\.\d+)?$/.test(value)) {
    const asNumber = parseFloat(value);
    const excelDate = excelSerialToDate_(asNumber);
    if (excelDate) {
      return excelDate;
    }
  }

  if (/^\d{8}$/.test(value)) {
    const year = Number(value.slice(0, 4));
    const month = Number(value.slice(4, 6)) - 1;
    const day = Number(value.slice(6, 8));
    const ymdDate = new Date(year, month, day);
    if (!isNaN(ymdDate.getTime())) {
      return ymdDate;
    }
  }

  let parsed = new Date(value);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }

  if (value.indexOf(' ') > -1 && value.indexOf('T') === -1) {
    parsed = new Date(value.replace(' ', 'T'));
    if (!isNaN(parsed.getTime())) {
      return parsed;
    }
  }

  const parts = value.split(/[\/\-]/).map(function(part) { return part.trim(); });
  if (parts.length === 3 && parts.every(function(part) { return /^\d+$/.test(part); })) {
    var p1 = Number(parts[0]);
    var p2 = Number(parts[1]);
    var p3 = Number(parts[2]);

    if (p3 < 100) {
      p3 = p3 < 50 ? 2000 + p3 : 1900 + p3;
    }

    var month;
    var day;
    var year;

    if (p1 > 12 && p2 <= 12) {
      day = p1;
      month = p2;
      year = p3;
    } else if (p2 > 12 && p1 <= 12) {
      month = p1;
      day = p2;
      year = p3;
    } else {
      month = p1;
      day = p2;
      year = p3;
    }

    const manualDate = new Date(year, month - 1, day);
    if (!isNaN(manualDate.getTime())) {
      return manualDate;
    }
  }

  return null;
}

function coerceDateValue_(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === 'number' && isFinite(value)) {
    const excelDate = excelSerialToDate_(value);
    if (excelDate) {
      return excelDate;
    }

    const numericDate = new Date(value);
    return isNaN(numericDate.getTime()) ? null : numericDate;
  }

  return parseFlexibleDateString_(value);
}

function getPreviousPeriod_(granularity, period) {
  if (!period) return '';
  switch (granularity) {
    case 'Week': {
      const parts = period.split('-W');
      if (parts.length !== 2) return '';
      const year = parseInt(parts[0], 10);
      const week = parseInt(parts[1], 10);
      if (week <= 1) {
        return `${year - 1}-W52`;
      }
      return `${year}-W${String(week - 1).padStart(2, '0')}`;
    }
    case 'Month': {
      const [y, m] = period.split('-').map(Number);
      if (!y || !m) return '';
      const date = new Date(y, m - 1, 1);
      date.setMonth(date.getMonth() - 1);
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    }
    case 'Quarter': {
      const [q, y] = period.split('-');
      if (!q || !y) return '';
      const n = parseInt(q.replace('Q', ''), 10);
      if (n <= 1) {
        return `Q4-${parseInt(y, 10) - 1}`;
      }
      return `Q${n - 1}-${y}`;
    }
    case 'Year':
      return String(parseInt(period, 10) - 1);
    default:
      return '';
  }
}

function formatPeriodLabel_(granularity, period) {
  if (!period) return 'Period';
  switch (granularity) {
    case 'Week':
      return period.replace(/^[0-9]{4}-/, '');
    case 'Month': {
      const [y, m] = period.split('-');
      if (!y || !m) return period;
      const date = new Date(Number(y), Number(m) - 1, 1);
      return date.toLocaleDateString('en-US', { month: 'short', year: '2-digit' });
    }
    case 'Quarter':
      return period.replace('-', ' ');
    case 'Year':
      return period;
    default:
      return period;
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
    'Courtesy & Communication': ['q1', 'q2', 'q3', 'q4', 'q5', 'q19'],
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
      Q15: 'Yes', Q16: 'Yes', Q17: 'No', Q18: 'Yes', Q19: 'Yes',
      'Q1 Note': 'Great opening', 'Q2 Note': 'Good closing',
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
