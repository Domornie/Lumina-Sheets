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
    const callbackAudioResult = processAudioFile_(qaData, { prefix: 'callback', label: 'Callback' });
    console.log('Audio processing completed');

    // Step 5: Calculate QA score
    const scoreResult = calculateQAScore_(qaData);
    console.log('Score calculated:', scoreResult.finalScore);

    // Step 6: Save to sheet
    const saveResult = saveQARecord_(qaData, { primary: audioResult, callback: callbackAudioResult }, scoreResult);
    console.log('Record saved with ID:', saveResult.qaId);

    // Step 7: Generate PDF (optional, don't fail if this errors)
    let pdfResult = null;
    try {
      pdfResult = generateQAPDF_(saveResult.record, scoreResult, qaData);
    } catch (pdfError) {
      console.warn('PDF generation failed (non-critical):', pdfError.message);
    }

    // Step 8: Prepare latest record details
    let finalRecord = saveResult.record || {};

    if (pdfResult && pdfResult.success) {
      const updatedRecord = updateQaRecordPdfInfo_(saveResult.qaId, pdfResult);
      if (updatedRecord) {
        finalRecord = updatedRecord;
      } else {
        const refreshed = getQARecordById(saveResult.qaId);
        if (refreshed) {
          finalRecord = refreshed;
        }
      }
    } else {
      const refreshed = getQARecordById(saveResult.qaId);
      if (refreshed) {
        finalRecord = refreshed;
      }
    }

    // Step 9: Return success response (sanitized for client)
    const response = {
      success: true,
      qaId: String(saveResult.qaId || ''),
      audioUrl: resolveAudioUrlFromRecord_(finalRecord, audioResult && audioResult.url),
      audioId: resolveAudioIdFromRecord_(finalRecord, audioResult && audioResult.id),
      audioName: resolveAudioNameFromRecord_(finalRecord, audioResult && audioResult.name),
      callbackAudioUrl: resolveCallbackAudioUrlFromRecord_(finalRecord, callbackAudioResult && callbackAudioResult.url),
      callbackAudioId: resolveCallbackAudioIdFromRecord_(finalRecord, callbackAudioResult && callbackAudioResult.id),
      callbackAudioName: resolveCallbackAudioNameFromRecord_(finalRecord, callbackAudioResult && callbackAudioResult.name),
      scoreResult: scoreResult,
      record: finalRecord,
      timestamp: new Date().toISOString()
    };

    if (pdfResult && pdfResult.success) {
      response.qaPdfUrl = pdfResult.fileUrl || '';
      response.qaPdfId = pdfResult.fileId || '';
      response.qaPdfName = pdfResult.fileName || '';
    }

    console.log('=== QA SUBMISSION COMPLETED SUCCESSFULLY ===');
    return sanitizeForClient_(response);

  } catch (error) {
    console.error('=== QA SUBMISSION FAILED ===');
    console.error('Error:', error.message);
    console.error('Stack:', error.stack);

    // Return a proper error response
    return sanitizeForClient_({
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
      debug: {
        function: 'clientUploadAudioAndSaveQA',
        stack: error.stack
      }
    });
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

function processAudioFile_(data, options = {}) {
  const prefixRaw = options.prefix ? String(options.prefix).trim() : '';
  const prefix = prefixRaw ? prefixRaw.replace(/\s+/g, '') : '';
  const optional = options.optional !== undefined ? options.optional : !!prefix;
  const labelBase = options.label || (prefix ? `${prefixRaw.charAt(0).toUpperCase()}${prefixRaw.slice(1)} Recording` : 'Call Recording');

  try {
    console.log(`Processing ${labelBase.toLowerCase()}...`);

    if (!data || typeof data !== 'object') {
      if (optional) {
        return null;
      }
      throw new Error('Invalid form data for audio processing');
    }

    const composeKey = (base) => {
      if (prefix) {
        return `${prefix}${base}`;
      }
      return base.charAt(0).toLowerCase() + base.slice(1);
    };

    const fileDataKey = composeKey('AudioFileData');
    const fileNameKey = composeKey('AudioFileName');
    const fileSizeKey = composeKey('AudioFileSize');
    const fileTypeKey = composeKey('AudioFileType');
    const fileKey = composeKey('AudioFile');
    const linkKey = composeKey('CallLink');

    const base64Name = data[fileNameKey];
    const base64Data = data[fileDataKey];
    const base64Size = data[fileSizeKey];
    const base64Type = data[fileTypeKey] || 'audio/mpeg';

    const ensureTargetFolder = () => {
      const folder = ensureRootFolder_();
      const agentName = sanitizeName_(data.agentName || 'Unknown');
      const callDate = data.callDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

      const agentFolder = getOrCreateFolder_(folder, agentName);
      const dateFolder = getOrCreateFolder_(agentFolder, callDate);
      if (prefix) {
        return getOrCreateFolder_(dateFolder, 'Callback Recordings');
      }
      return dateFolder;
    };

    if (base64Data && base64Name) {
      const fileSizeMB = base64Size ? base64Size / (1024 * 1024) : 0;
      if (fileSizeMB > 45) {
        throw new Error('Audio file too large. Maximum size is 45MB, your file is ' + fileSizeMB.toFixed(1) + 'MB');
      }

      try {
        const binaryString = Utilities.base64Decode(base64Data);
        const blob = Utilities.newBlob(binaryString, base64Type, base64Name);
        const targetFolder = ensureTargetFolder();
        const uploadedFile = targetFolder.createFile(blob);
        uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        console.log(`${labelBase} uploaded successfully from base64`);
        return {
          url: uploadedFile.getUrl(),
          id: uploadedFile.getId(),
          name: uploadedFile.getName(),
          size: fileSizeMB,
          category: prefix || 'primary'
        };
      } catch (conversionError) {
        console.error('Error converting base64 to blob:', conversionError);
        throw new Error('Failed to process audio file: ' + conversionError.message);
      }
    }

    const uploadedFileObj = data[fileKey];
    if (uploadedFileObj && typeof uploadedFileObj.getName === 'function') {
      const blob = uploadedFileObj.getBlob();
      const fileSize = blob.getBytes().length;
      const fileSizeMB = fileSize / (1024 * 1024);
      console.log(`${labelBase} (legacy upload) size:`, fileSizeMB.toFixed(2), 'MB');

      if (fileSizeMB > 45) {
        throw new Error('Audio file too large. Maximum size is 45MB, your file is ' + fileSizeMB.toFixed(1) + 'MB');
      }

      const targetFolder = ensureTargetFolder();
      const uploadedFile = targetFolder.createFile(blob);
      uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      console.log(`${labelBase} uploaded successfully (legacy file)`);
      return {
        url: uploadedFile.getUrl(),
        id: uploadedFile.getId(),
        name: uploadedFile.getName(),
        size: fileSizeMB,
        category: prefix || 'primary'
      };
    }

    const linkValue = data[linkKey];
    if (linkValue && String(linkValue).trim() !== '') {
      console.log(`Using provided ${labelBase.toLowerCase()} link`);
      return {
        url: String(linkValue).trim(),
        id: null,
        name: 'External Link',
        size: 0,
        category: prefix || 'primary'
      };
    }

    console.log(`No ${labelBase.toLowerCase()} provided`);
    if (optional) {
      return null;
    }

    return {
      url: '',
      id: null,
      name: 'No Audio',
      size: 0,
      category: prefix || 'primary'
    };

  } catch (error) {
    console.error('Audio processing error:', error);
    throw new Error(`${labelBase} processing failed: ` + error.message);
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

    const safeData = data || {};
    let primaryAudio = {};
    let callbackAudio = null;

    if (audioResult && typeof audioResult === 'object' && (audioResult.primary !== undefined || audioResult.callback !== undefined)) {
      primaryAudio = audioResult.primary || {};
      callbackAudio = audioResult.callback || null;
    } else {
      primaryAudio = audioResult || {};
    }
    const timestamp = new Date().toISOString();

    const providedIdRaw = safeData.recordId || safeData.qaId || safeData.id || '';
    const providedId = providedIdRaw ? String(providedIdRaw).trim() : '';
    const existingRowInfo = providedId ? findQaRecordRow_(providedId) : null;

    const qaId = existingRowInfo && existingRowInfo.record && existingRowInfo.record.ID
      ? existingRowInfo.record.ID
      : (providedId || Utilities.getUuid());

    const sheet = getQaSheet_();
    const headers = getQaHeaders_();
    const baseRowValues = existingRowInfo ? existingRowInfo.rowValues.slice() : new Array(headers.length).fill('');

    const lowerCaseDataMap = {};
    Object.keys(safeData).forEach(key => {
      lowerCaseDataMap[key.toLowerCase()] = safeData[key];
    });

    const hasOwn = (obj, key) => Object.prototype.hasOwnProperty.call(obj, key);

    const getDataValue = key => {
      if (!key) return undefined;
      if (hasOwn(safeData, key)) {
        return safeData[key];
      }
      const lowerKey = key.toLowerCase();
      if (hasOwn(lowerCaseDataMap, lowerKey)) {
        return lowerCaseDataMap[lowerKey];
      }
      return undefined;
    };

    const normalizeAnswer = (value, fallback) => {
      if (value === undefined || value === null) {
        return fallback;
      }
      const str = String(value).trim();
      if (!str) {
        return fallback;
      }
      const lower = str.toLowerCase();
      if (lower === 'na' || lower === 'n/a' || lower === 'n\\a') {
        return 'N/A';
      }
      if (lower === 'yes' || lower === 'y') {
        return 'Yes';
      }
      if (lower === 'no' || lower === 'n') {
        return 'No';
      }
      return str;
    };

    const resolveNoteValue = (number, existingValue) => {
      const candidates = [
        'q' + number + 'Note',
        'q' + number + 'note',
        'Q' + number + ' Note',
        'Q' + number + ' note',
        'Q' + number + 'Note',
        'c' + number,
        'C' + number
      ];

      for (let i = 0; i < candidates.length; i++) {
        const candidate = candidates[i];
        const candidateValue = getDataValue(candidate);
        if (candidateValue !== undefined) {
          return candidateValue;
        }
      }

      return existingValue || '';
    };

    const rowData = headers.map((col, index) => {
      const header = col || '';
      const normalized = header.toString().toLowerCase().replace(/\s+/g, '');
      const existingValue = baseRowValues[index] !== undefined ? baseRowValues[index] : '';

      if (/^q\d+$/.test(normalized)) {
        const questionNumber = normalized.replace('q', '');
        const answerKey = 'q' + questionNumber;
        const answerValue = getDataValue(answerKey);
        const fallback = existingValue || 'N/A';
        return normalizeAnswer(answerValue, fallback);
      }

      if (/^q\d+note$/.test(normalized)) {
        const number = normalized.replace('q', '').replace('note', '');
        return resolveNoteValue(number, existingValue);
      }

      if (/^c\d+$/.test(normalized)) {
        const number = normalized.replace(/[^0-9]/g, '');
        return resolveNoteValue(number, existingValue);
      }

      switch (normalized) {
        case 'id':
          return qaId;
        case 'timestamp':
          return timestamp;
        case 'callername': {
          const value = getDataValue('callerName');
          return value !== undefined ? value : existingValue || '';
        }
        case 'agentname': {
          const value = getDataValue('agentName');
          return value !== undefined ? value : existingValue || '';
        }
        case 'agentemail': {
          const value = getDataValue('agentEmail');
          return value !== undefined ? value : existingValue || '';
        }
        case 'clientname': {
          const value = getDataValue('clientName');
          return value !== undefined ? value : existingValue || '';
        }
        case 'calldate': {
          const value = getDataValue('callDate');
          return value !== undefined ? value : existingValue || '';
        }
        case 'casenumber': {
          const value = getDataValue('caseNumber');
          return value !== undefined ? value : existingValue || '';
        }
        case 'calllink': {
          const audioUrl = primaryAudio && primaryAudio.url ? String(primaryAudio.url).trim() : '';
          const linkCandidates = ['callLink', 'callRecordingUrl', 'callUrl', 'recordingLink'];
          let linkValue = '';
          for (let i = 0; i < linkCandidates.length; i++) {
            const candidate = getDataValue(linkCandidates[i]);
            if (candidate !== undefined) {
              linkValue = String(candidate || '').trim();
              break;
            }
          }
          return audioUrl || linkValue || existingValue || '';
        }
        case 'callrecordingid':
        case 'callrecordingfileid':
        case 'callrecordingdriveid':
        case 'callrecordinggid':
        case 'audiorecordingid':
        case 'audiofileid': {
          const audioId = primaryAudio && primaryAudio.id ? String(primaryAudio.id) : '';
          const idValue = getDataValue('callRecordingId');
          const fallbackId = idValue !== undefined ? idValue : getDataValue('audioFileId');
          return audioId || (fallbackId !== undefined ? fallbackId : existingValue || '');
        }
        case 'callrecordingname':
        case 'audiorecordingname':
        case 'audiofilename': {
          const audioName = primaryAudio && primaryAudio.name ? primaryAudio.name : '';
          const nameValue = getDataValue('callRecordingName');
          const fallbackName = nameValue !== undefined ? nameValue : getDataValue('audioFileName');
          return audioName || (fallbackName !== undefined ? fallbackName : existingValue || '');
        }
        case 'callbackcalllink':
        case 'callbackrecordinglink':
        case 'callbackrecordingurl':
        case 'callbackcallurl':
        case 'callbackaudiourl':
        case 'callbackrecording':
        case 'callbacklink': {
          const callbackUrl = callbackAudio && callbackAudio.url ? String(callbackAudio.url).trim() : '';
          if (callbackUrl) {
            return callbackUrl;
          }
          const callbackCandidates = ['callbackCallLink', 'callbackRecordingUrl', 'callbackCallUrl', 'callbackAudioUrl', 'callbackRecording', 'callbackLink'];
          for (let i = 0; i < callbackCandidates.length; i++) {
            const candidate = getDataValue(callbackCandidates[i]);
            if (candidate !== undefined) {
              const value = String(candidate || '').trim();
              if (value) {
                return value;
              }
            }
          }
          return existingValue || '';
        }
        case 'callbackcallrecordingid':
        case 'callbackrecordingid':
        case 'callbackaudiorecordingid':
        case 'callbackaudiofileid':
        case 'callbackaudioid': {
          const callbackId = callbackAudio && callbackAudio.id ? String(callbackAudio.id) : '';
          if (callbackId) {
            return callbackId;
          }
          const callbackIdCandidates = ['callbackCallRecordingId', 'callbackAudioFileId', 'callbackAudioId'];
          for (let i = 0; i < callbackIdCandidates.length; i++) {
            const candidate = getDataValue(callbackIdCandidates[i]);
            if (candidate !== undefined) {
              const value = String(candidate || '').trim();
              if (value) {
                return value;
              }
            }
          }
          return existingValue || '';
        }
        case 'callbackcallrecordingname':
        case 'callbackrecordingname':
        case 'callbackaudiorecordingname':
        case 'callbackaudiofilename': {
          const callbackName = callbackAudio && callbackAudio.name ? callbackAudio.name : '';
          if (callbackName) {
            return callbackName;
          }
          const callbackNameCandidates = ['callbackCallRecordingName', 'callbackAudioFileName'];
          for (let i = 0; i < callbackNameCandidates.length; i++) {
            const candidate = getDataValue(callbackNameCandidates[i]);
            if (candidate !== undefined) {
              const value = String(candidate || '').trim();
              if (value) {
                return value;
              }
            }
          }
          return existingValue || '';
        }
        case 'callbackrecordingsize':
        case 'callbackaudiosize': {
          if (callbackAudio && callbackAudio.size !== undefined) {
            return callbackAudio.size;
          }
          const callbackSize = getDataValue('callbackAudioSize');
          return callbackSize !== undefined ? callbackSize : (existingValue || '');
        }
        case 'auditorname': {
          const value = getDataValue('auditorName');
          return value !== undefined ? value : existingValue || '';
        }
        case 'auditdate': {
          const value = getDataValue('auditDate');
          return value !== undefined ? value : existingValue || '';
        }
        case 'feedbackshared': {
          const feedbackValue = getDataValue('feedbackShared');
          if (feedbackValue !== undefined) {
            if (typeof feedbackValue === 'string') {
              const normalizedFeedback = feedbackValue.trim().toLowerCase();
              if (['yes', 'true', '1'].indexOf(normalizedFeedback) !== -1) {
                return 'Yes';
              }
              if (['no', 'false', '0'].indexOf(normalizedFeedback) !== -1) {
                return 'No';
              }
            }
            return feedbackValue ? 'Yes' : 'No';
          }
          const flagValue = getDataValue('feedbackSharedFlag');
          if (flagValue !== undefined) {
            if (typeof flagValue === 'string') {
              const normalizedFlag = flagValue.trim().toLowerCase();
              if (['yes', 'true', '1'].indexOf(normalizedFlag) !== -1) {
                return 'Yes';
              }
              if (['no', 'false', '0'].indexOf(normalizedFlag) !== -1) {
                return 'No';
              }
            }
            return flagValue ? 'Yes' : 'No';
          }
          return existingValue || '';
        }
        case 'feedbacksharedat': {
          const value = getDataValue('feedbackSharedAt');
          return value !== undefined ? value : existingValue || '';
        }
        case 'totalscore':
          return (typeof scoreResult.earned === 'number') ? scoreResult.earned : (existingValue || 0);
        case 'percentage':
          return (typeof scoreResult.percentage === 'number') ? scoreResult.percentage : (existingValue || 0);
        case 'overallfeedback': {
          const value = getDataValue('overallFeedback');
          return value !== undefined ? value : existingValue || '';
        }
        case 'notes': {
          const value = getDataValue('notes');
          return value !== undefined ? value : existingValue || '';
        }
        case 'agentfeedback': {
          const value = getDataValue('agentFeedback');
          return value !== undefined ? value : existingValue || '';
        }
        case 'coachingprovided':
          return existingValue || 'No';
        default:
          if (normalized.indexOf('pdf') !== -1) {
            return existingValue || '';
          }

          const camelKey = header.charAt(0).toLowerCase() + header.slice(1);
          const value = getDataValue(camelKey);
          if (value !== undefined) {
            return value;
          }
          return existingValue || '';
      }
    });

    if (existingRowInfo) {
      sheet.getRange(existingRowInfo.rowIndex, 1, 1, headers.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

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

function normalizeBooleanOption_(value, defaultValue) {
  if (value === undefined || value === null || value === '') {
    return defaultValue;
  }

  if (typeof value === 'boolean') {
    return value;
  }

  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (['false', '0', 'no', 'off'].indexOf(normalized) !== -1) {
      return false;
    }
    if (['true', '1', 'yes', 'on'].indexOf(normalized) !== -1) {
      return true;
    }
  }

  return Boolean(value);
}

function generateQAPDF_(record, scoreResult, formData = {}) {
  try {
    if (typeof generateQaPdfReport === 'function') {
      const pdfOptions = {
        template: formData.pdfTemplate || 'standard',
        theme: formData.pdfTheme || 'professional',
        includeCharts: normalizeBooleanOption_(formData.includeCharts, true),
        includeRecommendations: normalizeBooleanOption_(formData.includeRecommendations, true),
        includeFullSnapshot: normalizeBooleanOption_(formData.includeFullSnapshot, true)
      };

      const existingPdfId = extractPdfIdFromRecord_(record);
      if (existingPdfId) {
        pdfOptions.existingFileId = existingPdfId;
      }

      return generateQaPdfReport(record, scoreResult, pdfOptions);
    }
    return { success: false, error: 'PDF generation not available' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function resolveAudioUrlFromRecord_(record, explicitUrl) {
  const direct = (explicitUrl || '').toString().trim();
  if (direct) {
    return direct;
  }

  if (!record || typeof record !== 'object') {
    return '';
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if ((normalized.indexOf('call') !== -1 || normalized.indexOf('audio') !== -1 || normalized.indexOf('recording') !== -1) &&
        (normalized.indexOf('url') !== -1 || normalized.indexOf('link') !== -1)) {
      const url = String(value).trim();
      if (url) {
        return url;
      }
    }
  }

  return '';
}

function resolveAudioIdFromRecord_(record, explicitId) {
  const direct = (explicitId || '').toString().trim();
  if (direct) {
    return direct;
  }

  if (!record || typeof record !== 'object') {
    return '';
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if ((normalized.indexOf('call') !== -1 || normalized.indexOf('audio') !== -1 || normalized.indexOf('recording') !== -1) &&
        (normalized.indexOf('id') !== -1 || normalized.indexOf('file') !== -1)) {
      const id = String(value).trim();
      if (id) {
        return id;
      }
    }
  }

  return '';
}

function resolveAudioNameFromRecord_(record, explicitName) {
  const direct = (explicitName || '').toString().trim();
  if (direct) {
    return direct;
  }

  if (!record || typeof record !== 'object') {
    return '';
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if ((normalized.indexOf('call') !== -1 || normalized.indexOf('audio') !== -1 || normalized.indexOf('recording') !== -1) &&
        (normalized.indexOf('name') !== -1 || normalized.indexOf('title') !== -1)) {
      const name = String(value).trim();
      if (name) {
        return name;
      }
    }
  }

  return '';
}

function resolveCallbackAudioUrlFromRecord_(record, explicitUrl) {
  const direct = (explicitUrl || '').toString().trim();
  if (direct) {
    return direct;
  }

  if (!record || typeof record !== 'object') {
    return '';
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if (normalized.indexOf('callback') !== -1 &&
        (normalized.indexOf('call') !== -1 || normalized.indexOf('audio') !== -1 || normalized.indexOf('recording') !== -1) &&
        (normalized.indexOf('url') !== -1 || normalized.indexOf('link') !== -1)) {
      const url = String(value).trim();
      if (url) {
        return url;
      }
    }
  }

  return '';
}

function resolveCallbackAudioIdFromRecord_(record, explicitId) {
  const direct = (explicitId || '').toString().trim();
  if (direct) {
    return direct;
  }

  if (!record || typeof record !== 'object') {
    return '';
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if (normalized.indexOf('callback') !== -1 &&
        (normalized.indexOf('call') !== -1 || normalized.indexOf('audio') !== -1 || normalized.indexOf('recording') !== -1) &&
        (normalized.indexOf('id') !== -1 || normalized.indexOf('file') !== -1)) {
      const id = String(value).trim();
      if (id) {
        return id;
      }
    }
  }

  return '';
}

function resolveCallbackAudioNameFromRecord_(record, explicitName) {
  const direct = (explicitName || '').toString().trim();
  if (direct) {
    return direct;
  }

  if (!record || typeof record !== 'object') {
    return '';
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if (normalized.indexOf('callback') !== -1 &&
        (normalized.indexOf('call') !== -1 || normalized.indexOf('audio') !== -1 || normalized.indexOf('recording') !== -1) &&
        (normalized.indexOf('name') !== -1 || normalized.indexOf('title') !== -1)) {
      const name = String(value).trim();
      if (name) {
        return name;
      }
    }
  }

  return '';
}

function sanitizeForClient_(value) {
  if (value === undefined) {
    return undefined;
  }

  if (value === null) {
    return null;
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (Array.isArray(value)) {
    return value.map(item => {
      const sanitized = sanitizeForClient_(item);
      return sanitized === undefined ? null : sanitized;
    });
  }

  if (typeof value === 'number') {
    return isFinite(value) ? value : null;
  }

  if (typeof value === 'object') {
    const sanitizedObject = {};
    Object.keys(value).forEach(key => {
      const sanitizedValue = sanitizeForClient_(value[key]);
      if (sanitizedValue !== undefined) {
        sanitizedObject[key] = sanitizedValue;
      }
    });
    return sanitizedObject;
  }

  return value;
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

function ensurePublicSharing_(driveItem) {
  try {
    if (!driveItem || typeof driveItem.setSharing !== 'function') {
      return;
    }

    const currentAccess = typeof driveItem.getSharingAccess === 'function'
      ? driveItem.getSharingAccess()
      : null;
    const currentPermission = typeof driveItem.getSharingPermission === 'function'
      ? driveItem.getSharingPermission()
      : null;

    if (currentAccess !== DriveApp.Access.ANYONE_WITH_LINK || currentPermission !== DriveApp.Permission.VIEW) {
      driveItem.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
  } catch (error) {
    console.warn('Unable to update sharing for item:', error);
  }
}

function ensureRootFolder_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const savedId = (props.getProperty(ROOT_PROP_KEY) || '').trim();

    if (savedId) {
      try {
        const f = DriveApp.getFolderById(savedId);
        f.getFiles(); // Test access
        ensurePublicSharing_(f);
        return f;
      } catch (_) { }
    }

    let parent = DriveApp.getRootFolder();
    for (const name of FALLBACK_PATH) {
      parent = getOrCreateFolder_(parent, name);
    }
    ensurePublicSharing_(parent);
    props.setProperty(ROOT_PROP_KEY, parent.getId());
    return parent;
  } catch (error) {
    console.error('Error ensuring root folder:', error);
    throw error;
  }
}

function normalizeExportSettings_(settings) {
  const safe = settings || {};
  return {
    includePerformance: safe.includePerformance !== false,
    includeRanking: safe.includeRanking !== false,
    includeLinks: safe.includeLinks !== false,
    includeContext: safe.includeContext !== false,
    period: typeof safe.period === 'string' ? safe.period : ''
  };
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

    const qualityRecognition = buildQualityRecognitionHighlights_(profiles, {
      timezone,
      passMark: context.passMark
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

    const filteredForMatrix = filterRecordsForIntelligence_(records, { ...context, period: '' });
    const agentMatrix = buildAgentGranularityMatrix_(context, filteredForMatrix, {
      displayLookup: agentDisplayLookup
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
      programMetrics,
      qualityRecognition,
      agentMatrix
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

function clientExportAgentMatrix(request = {}) {
  try {
    const settings = normalizeExportSettings_(request.settings);
    const targetPeriod = request.period
      || (settings.period && settings.period !== 'latest' ? settings.period : '');

    const normalization = normalizeIntelligenceRequest_({
      granularity: request.granularity || (request.context && request.context.granularity),
      period: targetPeriod,
      agent: request.agent || (request.context && request.context.filters ? request.context.filters.agent : ''),
      campaignId: request.campaignId || (request.context && request.context.filters ? request.context.filters.campaignId : ''),
      program: request.program || (request.context && request.context.filters ? request.context.filters.program : ''),
      depth: request.depth || (request.context && request.context.depth)
    }, getAllQA()) || {};

    const context = normalization.context || {};
    const records = normalization.records || [];
    const displayLookup = buildAgentDisplayLookup_(records);
    const filteredForMatrix = filterRecordsForIntelligence_(records, { ...context, period: '' });
    const matrix = buildAgentGranularityMatrix_(context, filteredForMatrix, { displayLookup });

    const exportResult = writeAgentMatrixToSheet_(matrix, settings, context);

    return {
      success: true,
      fileId: exportResult.id,
      url: exportResult.url,
      name: exportResult.name
    };
  } catch (error) {
    console.error('clientExportAgentMatrix failed:', error);
    writeError('clientExportAgentMatrix', error);
    return { success: false, error: error && error.message ? error.message : 'Unable to export agent matrix.' };
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
    biWeek: callDate instanceof Date ? formatBiWeekKey_(callDate) : '',
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
    case 'Bi-Week':
      return latest.biWeek;
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
      case 'Bi-Week':
        return record.biWeek === period;
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

function buildQualityRecognitionHighlights_(profiles = [], options = {}) {
  if (!Array.isArray(profiles) || !profiles.length) {
    return [];
  }

  const timezone = options.timezone || Session.getScriptTimeZone();
  const passMarkValue = typeof options.passMark === 'number' ? options.passMark : QA_INTEL_PASS_MARK;
  const passThreshold = passMarkValue > 1
    ? Math.round(passMarkValue)
    : Math.round(passMarkValue * 100);

  const eligible = profiles.filter(profile => {
    return typeof profile.avgScore === 'number'
      && Number.isFinite(profile.avgScore)
      && Number(profile.evaluations) > 0;
  });

  if (!eligible.length) {
    return [];
  }

  const sorted = eligible.slice().sort((a, b) => {
    const scoreDiff = (b.avgScore || 0) - (a.avgScore || 0);
    if (scoreDiff !== 0) {
      return scoreDiff;
    }

    const evalDiff = (Number(b.evaluations) || 0) - (Number(a.evaluations) || 0);
    if (evalDiff !== 0) {
      return evalDiff;
    }

    const passDiff = (b.passRate || 0) - (a.passRate || 0);
    if (passDiff !== 0) {
      return passDiff;
    }

    const recentA = a.recentDate instanceof Date ? a.recentDate.getTime() : 0;
    const recentB = b.recentDate instanceof Date ? b.recentDate.getTime() : 0;
    if (recentB !== recentA) {
      return recentB - recentA;
    }

    const nameA = (a.displayName || a.name || a.rawName || '').toString().toLowerCase();
    const nameB = (b.displayName || b.name || b.rawName || '').toString().toLowerCase();
    return nameA.localeCompare(nameB);
  });

  return sorted.slice(0, 3).map((profile, index) => {
    const avgScore = Math.round(profile.avgScore);
    const passRate = typeof profile.passRate === 'number' && Number.isFinite(profile.passRate)
      ? Math.round(profile.passRate)
      : null;

    const recognition = {
      rank: index + 1,
      agent: profile.displayName || profile.name || profile.rawName || 'Agent',
      avgScore,
      passRate,
      evaluations: Number(profile.evaluations) || 0
    };

    if (profile.recentDate instanceof Date) {
      recognition.lastEvaluation = Utilities.formatDate(profile.recentDate, timezone, 'MMM d, yyyy');
    } else if (profile.recentDate) {
      recognition.lastEvaluation = String(profile.recentDate);
    }

    if (Number.isFinite(avgScore) && Number.isFinite(passThreshold)) {
      recognition.deltaFromPass = avgScore - passThreshold;
    }

    return recognition;
  });
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

function resolveCallLinkFromRecord_(record) {
  const raw = record && record.raw ? record.raw : {};
  const link = getRecordFieldValue_(raw, [
    'callLink', 'Call Link', 'Call Recording Url', 'Call Recording URL',
    'callRecordingUrl', 'recordingLink', 'callUrl', 'Call Url',
    'AudioUrl', 'Audio Url', 'Audio URL', 'Recording URL'
  ]);
  return link ? String(link).trim() : '';
}

function resolvePdfLinkFromRecord_(record) {
  const raw = record && record.raw ? record.raw : {};
  const pdf = getRecordFieldValue_(raw, [
    'qaPdfUrl', 'QA PDF URL', 'qaPdfLink', 'PDF Url', 'PDF URL', 'Pdf Link',
    'QA PDF', 'QA Pdf', 'pdfLink'
  ]);
  return pdf ? String(pdf).trim() : '';
}

function resolveResultLinkFromRecord_(record) {
  const raw = record && record.raw ? record.raw : {};
  const result = getRecordFieldValue_(raw, [
    'Result Link', 'Result Url', 'QA Url', 'QA Link', 'Review Url',
    'Submission Link', 'Submission Url'
  ]);
  return result ? String(result).trim() : '';
}

function resolveResultLabelFromRecord_(record) {
  if (!record) {
    return '';
  }
  const raw = record.raw || {};
  const explicit = getRecordFieldValue_(raw, ['Result', 'QA Result', 'Outcome', 'Status']);
  if (explicit !== null && explicit !== undefined && explicit !== '') {
    const value = String(explicit).trim();
    if (value) {
      return value;
    }
  }

  if (typeof record.recordScore === 'number') {
    const score = Math.round(record.recordScore);
    const descriptor = record.pass ? 'Pass' : 'Needs Improvement';
    return `${score}% • ${descriptor}`;
  }

  return record.pass ? 'Pass' : '';
}

function buildAgentGranularityMatrix_(context, records, options = {}) {
  if (!context) {
    return { granularity: '', periods: [], agents: [] };
  }

  const granularity = context.granularity || 'Week';
  const depth = Number(context.depth) > 0 ? Number(context.depth) : 6;
  const startPeriod = context.period || '';
  if (!startPeriod) {
    return { granularity, periods: [], agents: [] };
  }

  const visited = new Set();
  const periods = [];
  let cursor = startPeriod;
  let steps = 0;
  const displayLookup = options.displayLookup || {};
  const universe = new Set();
  const safeRecords = Array.isArray(records) ? records : [];

  while (cursor && steps < depth && !visited.has(cursor)) {
    visited.add(cursor);
    const bucket = filterRecordsForIntelligence_(safeRecords, { ...context, period: cursor });
    const totalEvaluations = bucket.length;
    const aggregates = {};
    const agentDetails = {};

    bucket.forEach(record => {
      if (!record) {
        return;
      }
      const identifier = record.agent || 'Unassigned';
      if (!aggregates[identifier]) {
        aggregates[identifier] = {
          evaluations: 0,
          scoreSum: 0,
          passCount: 0
        };
      }
      if (!agentDetails[identifier]) {
        agentDetails[identifier] = { latest: null, bestScore: -Infinity };
      }

      const stats = aggregates[identifier];
      stats.evaluations += 1;
      stats.scoreSum += Number(record.percentage) || 0;
      if (record.pass) {
        stats.passCount += 1;
      }

      const resolvedScore = typeof record.recordScore === 'number'
        ? Math.round(record.recordScore)
        : Math.round((record.percentage || 0) * 100);
      const links = {
        callLink: resolveCallLinkFromRecord_(record),
        pdfLink: resolvePdfLinkFromRecord_(record),
        resultLink: resolveResultLinkFromRecord_(record)
      };

      const detail = {
        callDate: record.callDateIso || '',
        score: resolvedScore,
        pass: record.pass,
        result: resolveResultLabelFromRecord_(record),
        callLink: links.callLink,
        pdfLink: links.pdfLink,
        resultLink: links.resultLink
      };

      const currentDetail = agentDetails[identifier];
      if (!currentDetail.latest || (record.callDate && record.callDate > (currentDetail.callDate || new Date(0)))) {
        currentDetail.latest = detail;
        currentDetail.callDate = record.callDate || null;
      }
      if (resolvedScore > currentDetail.bestScore) {
        currentDetail.bestScore = resolvedScore;
        currentDetail.best = detail;
      }

      universe.add(identifier);
    });

    const metrics = {};
    Object.keys(aggregates).forEach(identifier => {
      const stats = aggregates[identifier];
      const evals = stats.evaluations;
      const avgScore = evals ? roundOneDecimal_((stats.scoreSum / evals) * 100) : null;
      const passRate = evals ? roundOneDecimal_((stats.passCount / evals) * 100) : null;
      const evaluationShare = totalEvaluations
        ? roundOneDecimal_((stats.evaluations / totalEvaluations) * 100)
        : null;

      metrics[identifier] = {
        avgScore,
        passRate,
        evaluations: evals,
        evaluationShare,
        latest: agentDetails[identifier] ? agentDetails[identifier].latest : null,
        best: agentDetails[identifier] ? agentDetails[identifier].best : null
      };
    });

    const ranking = Object.keys(metrics)
      .map(id => ({ id, score: typeof metrics[id].avgScore === 'number' ? metrics[id].avgScore : -Infinity }))
      .sort((a, b) => b.score - a.score);

    ranking.forEach((entry, index) => {
      const metric = metrics[entry.id];
      if (!metric) return;
      metric.rank = index + 1;
      metric.placement = index === 0
        ? 'Champion'
        : index === 1
          ? 'Runner-up'
          : index === 2
            ? 'Contender'
            : 'Needs Focus';
      metric.needsImprovement = typeof metric.avgScore === 'number' ? metric.avgScore < 80 : false;
    });

    periods.push({
      period: cursor,
      label: formatPeriodLabel_(granularity, cursor),
      totalEvaluations,
      agentCount: Object.keys(metrics).length,
      metrics
    });

    cursor = getPreviousPeriod_(granularity, cursor);
    steps += 1;
  }

  const ordered = periods.reverse();
  ordered.forEach((period, idx) => {
    if (idx === 0) return;
    const previous = ordered[idx - 1];
    Object.keys(period.metrics || {}).forEach(agentId => {
      const metric = period.metrics[agentId];
      const prevMetric = previous.metrics ? previous.metrics[agentId] : null;
      if (metric && prevMetric && typeof metric.avgScore === 'number' && typeof prevMetric.avgScore === 'number') {
        metric.momentum = roundOneDecimal_(metric.avgScore - prevMetric.avgScore);
      }
    });
  });

  const agents = Array.from(universe).map(identifier => ({
    id: identifier,
    label: resolveAgentDisplayNameFromLookup_(identifier, displayLookup)
  })).sort((a, b) => {
    const nameA = (a.label || a.id || '').toString().toLowerCase();
    const nameB = (b.label || b.id || '').toString().toLowerCase();
    return nameA.localeCompare(nameB);
  });

  return {
    granularity,
    periods: ordered,
    agents
  };
}

function determineExportPeriods_(matrix, settings) {
  if (!matrix || !Array.isArray(matrix.periods)) {
    return [];
  }

  const selected = settings && settings.period ? settings.period : '';
  if (!selected) {
    return matrix.periods;
  }

  if (selected === 'latest') {
    return matrix.periods.length ? [matrix.periods[matrix.periods.length - 1]] : [];
  }

  return matrix.periods.filter(period => period && period.period === selected);
}

function buildAgentMatrixSheetData_(matrix, settings = {}) {
  if (!matrix || !Array.isArray(matrix.periods) || !matrix.periods.length) {
    return { rows: [], headerRowIndex: 0, columns: [] };
  }

  const normalized = normalizeExportSettings_(settings);
  const includePerformance = normalized.includePerformance;
  const includeRanking = normalized.includeRanking;
  const includeLinks = normalized.includeLinks;
  const includeContext = normalized.includeContext;

  const exportPeriods = determineExportPeriods_(matrix, normalized);
  if (!exportPeriods.length) {
    return { rows: [], headerRowIndex: 0, columns: [] };
  }

  const agentDisplay = {};
  (matrix.agents || []).forEach(agent => {
    if (agent && agent.id) {
      agentDisplay[agent.id] = agent.label || agent.id;
    }
  });

  const rows = [];
  if (includeContext) {
    rows.push(['Report', 'QA Agent Performance Export']);
    rows.push(['Granularity', matrix.granularity || '']);
    rows.push(['Included Periods', exportPeriods.map(p => p && p.label ? p.label : p.period).filter(Boolean).join(' | ')]);
    rows.push(['Generated At', new Date().toISOString()]);
    rows.push([]);
  }

  const columns = [
    { key: 'granularity', label: 'Granularity' },
    { key: 'periodKey', label: 'Period Key' },
    { key: 'periodLabel', label: 'Period Label' },
    { key: 'agentId', label: 'Agent Identifier' },
    { key: 'agentName', label: 'Agent Name' }
  ];

  if (includeRanking) {
    columns.push(
      { key: 'rank', label: 'Rank' },
      { key: 'placement', label: 'Placement Badge' },
      { key: 'momentum', label: 'Momentum vs Prior (pts)' },
      { key: 'improvement', label: 'Needs Improvement?' }
    );
  }

  if (includePerformance) {
    columns.push(
      { key: 'avgScore', label: 'Average Score (%)' },
      { key: 'passRate', label: 'Pass Rate (%)' },
      { key: 'evaluations', label: 'Evaluations' },
      { key: 'evaluationShare', label: 'Evaluation Share (%)' }
    );
  }

  if (includeLinks) {
    columns.push(
      { key: 'result', label: 'Latest Result' },
      { key: 'callLink', label: 'Call Link' },
      { key: 'pdfLink', label: 'QA PDF Link' },
      { key: 'resultLink', label: 'Result / Playback Link' }
    );
  }

  const headerRowIndex = rows.length + 1;
  rows.push(columns.map(column => column.label));

  const allAgents = new Set(Object.keys(agentDisplay));
  exportPeriods.forEach(period => {
    if (period && period.metrics) {
      Object.keys(period.metrics).forEach(key => allAgents.add(key));
    }
  });

  const agentOrder = Array.from(allAgents);
  agentOrder.sort((a, b) => {
    const nameA = (agentDisplay[a] || a || '').toString().toLowerCase();
    const nameB = (agentDisplay[b] || b || '').toString().toLowerCase();
    return nameA.localeCompare(nameB);
  });

  exportPeriods.forEach(period => {
    if (!period) {
      return;
    }
    const metrics = period.metrics || {};

    agentOrder.forEach(agentId => {
      const detail = metrics[agentId] || {};
      const avgScore = typeof detail.avgScore === 'number' && !Number.isNaN(detail.avgScore)
        ? detail.avgScore
        : '';
      const passRate = typeof detail.passRate === 'number' && !Number.isNaN(detail.passRate)
        ? detail.passRate
        : '';
      const evaluations = typeof detail.evaluations === 'number' && !Number.isNaN(detail.evaluations)
        ? detail.evaluations
        : '';
      const share = typeof detail.evaluationShare === 'number' && !Number.isNaN(detail.evaluationShare)
        ? detail.evaluationShare
        : '';

      const latest = detail.latest || {};
      const momentum = typeof detail.momentum === 'number' && !Number.isNaN(detail.momentum)
        ? detail.momentum
        : '';

      const row = [];
      columns.forEach(column => {
        switch (column.key) {
          case 'granularity':
            row.push(matrix.granularity || '');
            break;
          case 'periodKey':
            row.push(period.period || '');
            break;
          case 'periodLabel':
            row.push(period.label || '');
            break;
          case 'agentId':
            row.push(agentId || '');
            break;
          case 'agentName':
            row.push(agentDisplay[agentId] || agentId || 'Unassigned');
            break;
          case 'rank':
            row.push(detail.rank || '');
            break;
          case 'placement':
            row.push(detail.placement || '');
            break;
          case 'momentum':
            row.push(momentum);
            break;
          case 'improvement':
            row.push(detail.needsImprovement ? 'Yes' : '');
            break;
          case 'avgScore':
            row.push(avgScore);
            break;
          case 'passRate':
            row.push(passRate);
            break;
          case 'evaluations':
            row.push(evaluations);
            break;
          case 'evaluationShare':
            row.push(share);
            break;
          case 'result':
            row.push(latest.result || '');
            break;
          case 'callLink':
            row.push(latest.callLink || '');
            break;
          case 'pdfLink':
            row.push(latest.pdfLink || '');
            break;
          case 'resultLink':
            row.push(latest.resultLink || '');
            break;
          default:
            row.push('');
            break;
        }
      });

      rows.push(row);
    });
  });

  return { rows, headerRowIndex, columns };
}

function writeAgentMatrixToSheet_(matrix, settings = {}, context = {}) {
  const sheetData = buildAgentMatrixSheetData_(matrix, settings);
  if (!sheetData.rows.length) {
    throw new Error('Agent matrix export has no rows to write.');
  }

  const columnCount = sheetData.rows.reduce((max, row) => Math.max(max, row.length), 0) || 1;
  const normalizedRows = sheetData.rows.map(row => {
    const padded = row.slice();
    while (padded.length < columnCount) {
      padded.push('');
    }
    return padded;
  });

  const granularity = matrix.granularity || context.granularity || 'Period';
  const periodLabel = settings.period || context.period || 'all-periods';
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const filename = `QA Agent Matrix - ${granularity} - ${periodLabel} - ${timestamp}`;

  const ss = SpreadsheetApp.create(filename);
  const sh = ss.getActiveSheet();
  sh.clear();
  sh.getRange(1, 1, normalizedRows.length, columnCount).setValues(normalizedRows);

  if (sheetData.headerRowIndex > 0) {
    sh.setFrozenRows(sheetData.headerRowIndex);
    sh.getRange(sheetData.headerRowIndex, 1, 1, columnCount)
      .setFontWeight('bold')
      .setBackground('#0d6efd')
      .setFontColor('#ffffff');
  }

  const dataStart = sheetData.headerRowIndex ? sheetData.headerRowIndex + 1 : 2;
  if (normalizedRows.length >= dataStart) {
    const dataRowCount = normalizedRows.length - dataStart + 1;
    if (dataRowCount > 0) {
      sh.getRange(dataStart, 1, dataRowCount, columnCount)
        .applyRowBanding(SpreadsheetApp.BandingTheme.TEAL, true, false);
    }
  }

  const numberColumns = sheetData.columns
    .map((col, idx) => ({ idx: idx + 1, key: col.key }))
    .filter(col => ['avgScore', 'passRate', 'evaluationShare', 'momentum'].includes(col.key));

  const numericRowCount = normalizedRows.length - sheetData.headerRowIndex;
  if (sheetData.headerRowIndex > 0 && numericRowCount > 0) {
    numberColumns.forEach(col => {
      const format = col.key === 'momentum' ? '0.0' : '0.0';
      sh.getRange(sheetData.headerRowIndex + 1, col.idx, numericRowCount, 1)
        .setNumberFormat(format);
    });
  }

  sh.autoResizeColumns(1, columnCount);

  const file = DriveApp.getFileById(ss.getId());
  const folder = ensureRootFolder_();
  if (folder) {
    folder.addFile(file);
    const parents = file.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      if (parent.getId() !== folder.getId()) {
        parent.removeFile(file);
      }
    }
  }

  ensurePublicSharing_(file);

  return {
    id: ss.getId(),
    url: ss.getUrl(),
    name: filename
  };
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

function getIsoWeeksInYear_(year) {
  const dec28 = new Date(Date.UTC(year, 11, 28));
  const weekKey = toISOWeek_(dec28);
  const parts = weekKey.split('-W');
  const week = parts.length === 2 ? parseInt(parts[1], 10) : 52;
  return Number.isFinite(week) ? week : 52;
}

function getIsoWeekStartDate_(year, week) {
  if (!Number.isFinite(year) || !Number.isFinite(week) || week < 1) {
    return null;
  }
  const simple = new Date(Date.UTC(year, 0, 1 + ((week - 1) * 7)));
  const dow = simple.getUTCDay();
  const isoStart = new Date(simple);
  if (dow <= 4 && dow !== 0) {
    isoStart.setUTCDate(simple.getUTCDate() - dow + 1);
  } else {
    isoStart.setUTCDate(simple.getUTCDate() + (8 - dow));
  }
  return isoStart;
}

function formatBiWeekKey_(date) {
  const weekKey = toISOWeek_(date);
  const parts = weekKey.split('-W');
  if (parts.length !== 2) {
    return '';
  }
  const year = parseInt(parts[0], 10);
  const week = parseInt(parts[1], 10);
  if (!year || !week) {
    return '';
  }
  const biWeekIndex = Math.floor((week - 1) / 2) + 1;
  return `${year}-BW${String(biWeekIndex).padStart(2, '0')}`;
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
    case 'Bi-Week': {
      const parts = period.split('-BW');
      if (parts.length !== 2) return '';
      const year = parseInt(parts[0], 10);
      const biWeek = parseInt(parts[1], 10);
      if (!year || !biWeek) return '';
      if (biWeek <= 1) {
        const previousYear = year - 1;
        const previousWeeks = getIsoWeeksInYear_(previousYear);
        const lastBiWeek = Math.ceil(previousWeeks / 2);
        return `${previousYear}-BW${String(lastBiWeek).padStart(2, '0')}`;
      }
      return `${year}-BW${String(biWeek - 1).padStart(2, '0')}`;
    }
    case 'Week': {
      const parts = period.split('-W');
      if (parts.length !== 2) return '';
      const year = parseInt(parts[0], 10);
      const week = parseInt(parts[1], 10);
      if (week <= 1) {
        const previousYear = year - 1;
        const priorYearWeeks = getIsoWeeksInYear_(previousYear);
        return `${previousYear}-W${String(priorYearWeeks).padStart(2, '0')}`;
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
    case 'Bi-Week': {
      const [yearPart, biWeekPart] = period.split('-BW');
      const year = parseInt(yearPart, 10);
      const biWeekNumber = parseInt(biWeekPart, 10);
      if (!year || !biWeekNumber) return period;
      const startWeek = ((biWeekNumber - 1) * 2) + 1;
      const startDate = getIsoWeekStartDate_(year, startWeek);
      if (!startDate) return period;
      const endDate = new Date(startDate.getTime());
      endDate.setUTCDate(endDate.getUTCDate() + 13);
      const dateOptions = { month: 'short', day: 'numeric' };
      const startLabel = startDate.toLocaleDateString('en-US', dateOptions);
      const endLabel = endDate.toLocaleDateString('en-US', dateOptions);
      const yearLabel = startDate.getUTCFullYear() === endDate.getUTCFullYear()
        ? startDate.getUTCFullYear()
        : `${startDate.getUTCFullYear()} / ${endDate.getUTCFullYear()}`;
      return `${startLabel} - ${endLabel} ${yearLabel}`;
    }
    case 'Week': {
      const parts = period.split('-W');
      if (parts.length !== 2) return period.replace(/^[0-9]{4}-/, '');
      const year = parseInt(parts[0], 10);
      const week = parseInt(parts[1], 10);
      const startDate = getIsoWeekStartDate_(year, week);
      if (!startDate) return period.replace(/^[0-9]{4}-/, '');
      const endDate = new Date(startDate.getTime());
      endDate.setUTCDate(endDate.getUTCDate() + 6);
      const options = { month: 'short', day: 'numeric' };
      const startLabel = startDate.toLocaleDateString('en-US', options);
      const endLabel = endDate.toLocaleDateString('en-US', options);
      const yearLabel = startDate.getUTCFullYear() === endDate.getUTCFullYear()
        ? startDate.getUTCFullYear()
        : `${startDate.getUTCFullYear()} / ${endDate.getUTCFullYear()}`;
      return `${startLabel} - ${endLabel} ${yearLabel}`;
    }
    case 'Month': {
      const [y, m] = period.split('-');
      if (!y || !m) return period;
      const date = new Date(Number(y), Number(m) - 1, 1);
      return date.toLocaleDateString('en-US', { month: 'long', year: 'numeric' });
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
// ADDITIONAL UTILITY FUNCTION FOR QA RECORD RETRIEVAL
// ==============================================================

function findQaRecordRow_(qaId) {
  try {
    if (!qaId && qaId !== 0) {
      return null;
    }

    const targetId = String(qaId).trim();
    if (!targetId) {
      return null;
    }

    const sheet = getQaSheet_();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return null;
    }

    const headers = data[0];
    const idColumnIndex = headers.findIndex(h => String(h).toLowerCase() === 'id');

    if (idColumnIndex === -1) {
      console.error('ID column not found in QA sheet');
      return null;
    }

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][idColumnIndex] || '').trim();
      if (rowId === targetId) {
        const rowValues = data[i].slice();
        const record = {};
        headers.forEach((header, index) => {
          record[header] = rowValues[index];
        });

        return {
          rowIndex: i + 1,
          headers,
          rowValues,
          record
        };
      }
    }

    console.warn('QA record not found for ID:', targetId);
    return null;

  } catch (error) {
    console.error('Error locating QA record row:', error);
    return null;
  }
}

function getQARecordById(qaId) {
  try {
    const info = findQaRecordRow_(qaId);
    return info ? info.record : null;
  } catch (error) {
    console.error('Error retrieving QA record by ID:', error);
    return null;
  }
}

function deleteQARecord(qaId) {
  try {
    if (qaId === undefined || qaId === null) {
      throw new Error('QA record ID is required.');
    }

    const targetId = String(qaId).trim();
    if (!targetId) {
      throw new Error('QA record ID is required.');
    }

    const info = findQaRecordRow_(targetId);
    if (!info || !info.rowIndex || info.rowIndex <= 1) {
      throw new Error('QA record not found.');
    }

    const sheet = getQaSheet_();
    sheet.deleteRow(info.rowIndex);

    return { success: true, deletedId: targetId };
  } catch (error) {
    console.error('Error deleting QA record:', error);
    throw error;
  }
}

function updateQaRecordPdfInfo_(qaId, pdfResult) {
  try {
    if (!pdfResult || !pdfResult.success) {
      return null;
    }

    const info = findQaRecordRow_(qaId);
    if (!info) {
      return null;
    }

    const sheet = getQaSheet_();
    const headers = info.headers;
    const updatedValues = info.rowValues.slice();

    const pdfUrl = pdfResult.fileUrl || '';
    const pdfId = pdfResult.fileId || '';
    const pdfName = pdfResult.fileName || '';
    let changed = false;

    headers.forEach((header, index) => {
      const normalized = String(header || '').toLowerCase();
      if (pdfUrl && normalized.indexOf('pdf') !== -1 && normalized.indexOf('url') !== -1) {
        if (updatedValues[index] !== pdfUrl) {
          updatedValues[index] = pdfUrl;
          changed = true;
        }
      }
      if (pdfId && normalized.indexOf('pdf') !== -1 && normalized.indexOf('id') !== -1) {
        if (updatedValues[index] !== pdfId) {
          updatedValues[index] = pdfId;
          changed = true;
        }
      }
      if (pdfName && normalized.indexOf('pdf') !== -1 && normalized.indexOf('name') !== -1) {
        if (updatedValues[index] !== pdfName) {
          updatedValues[index] = pdfName;
          changed = true;
        }
      }
    });

    if (changed) {
      sheet.getRange(info.rowIndex, 1, 1, headers.length).setValues([updatedValues]);
    }

    const record = {};
    headers.forEach((header, index) => {
      record[header] = updatedValues[index];
    });

    return record;

  } catch (error) {
    console.error('Error updating QA PDF info:', error);
    return null;
  }
}


function extractPdfIdFromRecord_(record) {
  if (!record || typeof record !== 'object') {
    return '';
  }

  const keys = Object.keys(record);
  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if (normalized.indexOf('pdf') !== -1 && normalized.indexOf('id') !== -1) {
      return String(value).trim();
    }
  }

  for (let i = 0; i < keys.length; i++) {
    const key = keys[i];
    const value = record[key];
    if (!value) {
      continue;
    }
    const normalized = String(key || '').toLowerCase();
    if (normalized.indexOf('pdf') !== -1 && normalized.indexOf('url') !== -1) {
      const url = String(value);
      let match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        return match[1];
      }
      match = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        return match[1];
      }
      match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) {
        return match[1];
      }
    }
  }

  return '';
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
