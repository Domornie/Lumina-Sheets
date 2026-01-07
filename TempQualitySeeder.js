/**
 * TempQualitySeeder.js
 * -----------------------------------------------------------------------------
 * Temporary QA seeding helpers for IBTR Quality sheet.
 * Use seedDecemberQualityDataIBTR({ force: true }) or runTempQualitySeeder()
 * to populate December QA rows for the provided agents.
 */

(function () {
  var G = (typeof globalThis === 'object')
    ? globalThis
    : (function () { return this; })();

  var DECEMBER_QA_AGENTS = [
    'Sonny Cuerdo',
    'Yonique Byfield',
    'Gabrielle Howell',
    'Dwayne Gordan',
    'Alexi Martin',
    'Romaun Grant',
    'Terry Ann Green',
    'Erzulye Parkes',
    'Ayeishia Rotti',
    'Sherika Williams',
    'Mishaunna Morrison',
    'Shanara Thompson',
    'Neako Powell',
    'Ayesha Reddie',
    'Sashana Dixon',
    'Sashagay Wanchope',
    'Briana Lindo-Dixon',
    'Vinnelle Fyine',
    'Shanique Buchanan',
    'Shaenelle Francis'
  ];

  var DECEMBER_QA_CLIENTS = [
    'CCBCU',
    'Protective',
    'Premier Transportation Resource Center',
    'CCBCU',
    'WellStreet Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU',
    'CCBCU',
    'CCBCU',
    'CCBCU',
    'Progress Rail Service Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'MSHS',
    'Progress Rail Service Center',
    'U.S. Silica',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'WellStreet Resource Center',
    'CCBCU',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Mayville Engineering Company',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'Progress Rail Service Center',
    'Premier Transportation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'Mayville Engineering Company',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'Premier Transportation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Mayville Engineering Company',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Premier Transportation',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Columbus School District',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Premier Transportation',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Parma City School District',
    'Columbus School District',
    'Premier Transportation',
    'CCBCU Resource Center',
    'Columbus City Schools',
    'Premier Transportation',
    'Columbus City Schools',
    'Columbus City Schools',
    'CCBCU Resource Center',
    'Premier Transportation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'Columbus City Schools',
    'Protective Life Corporation',
    'Protective Life Corporation',
    'Columbus School District',
    'Columbus City Schools',
    'Columbus City Schools',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'Columbus City Schools',
    'Wabash',
    'CCBCU Resource Center',
    'Parma City School District',
    'Protective Life Corporation',
    'Premier Transportation',
    'Wabash',
    'Columbus City Schools',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Columbus School District',
    'Columbus City Schools',
    'Columbus City Schools',
    'Wabash',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'CCBCU Resource Center',
    'Premier Transportation',
    'Parma City School District',
    'Premier Transportation',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Premier Transportation',
    'Premier Transportation',
    'Premier Transportation',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Premier Transportation',
    'Premier Transportation',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Wabash',
    'City of North Las Vegas',
    'Premier Transportation',
    'Premier Transportation',
    'Enstructure',
    'Columbus City Schools',
    'Premier Transportation',
    'CCBCU Resource Center',
    'Columbus City Schools',
    'Columbus City Schools',
    'Wabash',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Columbus City Schools',
    'Premier Transportation',
    'Columbus City Schools',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Premier Transportation',
    'CCBCU Resource Center',
    'Expro',
    'CCBCU Resource Center',
    'Columbus City Schools',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'Protective Life Corporation',
    'Premier Transportation',
    'Audacy',
    'CCBCU Resource Center',
    'Expro',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Audacy',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Wabash',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU',
    'Premier Transportation',
    'Premier Transportation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Enstructure',
    'Columbus City Schools',
    'Enstructure',
    'CCBCU',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Wabash',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Audacy',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'Wabash',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'City of North Las Vegas',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Wabash',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'Audacy',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Progress Rail Service Center',
    'Mayville Engineering Company',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Wabash',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Audacy',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Audacy',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Expro Americas LLC',
    'Wabash',
    'Expro Americas LLC',
    'Protective Life Corporation',
    'Protective Life Corporation',
    'Wabash',
    'Wabash',
    'Wabash',
    'Protective Life Corporation',
    'Areas',
    'Protective Life Corporation',
    'Wabash',
    'Wabash',
    'CCBCU Resource Center',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'Audacy',
    'Audacy',
    'Rogers Electrical',
    'Wabash',
    'Wabash',
    'Areas',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Expro Americas LLC',
    'CCBCU Resource Center',
    'Wabash',
    'CCBCU Resource Center',
    'Peachtree Orthopaedic Clinic',
    'Sisecam Chemical Resources',
    'Audacy',
    'CCBCU Resource Center',
    'Wabash',
    'Rogers Electrical',
    'Rogers Electrical',
    'Areas',
    'CCBCU Resource Center',
    'Kidde Global Solutions',
    'Wabash',
    'CCBCU Resource Center',
    'Kidde Global Solutions',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Areas',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Audacy',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Rogers Electrical',
    'CCBCU Resource Center',
    'Sisecam Chemical Resources LLC',
    'Wabash',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'Premier Transportation',
    'Wabash',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'Protective Life Corporation',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'Wabash',
    'Wabash',
    'Areas',
    'Areas',
    'Premier Transportation',
    'Audacy',
    'CCBCU Resource Center',
    'Protective Life Corporation',
    'CCBCU Resource Center',
    'Wabash'
  ];

  var CASE_LINKS = [
    'https://btr123.lightning.force.com/lightning/r/Task/00TVY00000TlJeV2AV/view',
    'https://btr123.lightning.force.com/lightning/r/Task/00TVY00000TkWQp2AN/view',
    'https://btr123.lightning.force.com/lightning/r/Task/00TVY00000TiCGM2A3/view'
  ];

  var CALL_LINKS = [
    'https://drive.google.com/file/d/1zngIwQ9Xv5TMIgrWaRoeGPbCrBs0lZdd/view',
    'https://drive.google.com/file/d/1JvYn0WWdDyNpY80jarRO0K8rfijCW_vu/view',
    'https://drive.google.com/file/d/1vE6TibH-OAC9nsfSmN4KLbazjimMTzQD/view'
  ];

  if (typeof G.seedDecemberQualityDataIBTR !== 'function') {
    G.seedDecemberQualityDataIBTR = function seedDecemberQualityDataIBTR(options) {
      var config = options || {};
      var year = (typeof config.year === 'number') ? config.year : new Date().getFullYear();
      var targetMonth = 11;
      var sheetName = G.QA_RECORDS || 'Quality';
      var headers = (G.QA_HEADERS && G.QA_HEADERS.length) ? G.QA_HEADERS.slice() : [];

      if (!headers.length) {
        throw new Error('QA_HEADERS is not available for IBTR QA seeding.');
      }

      if (typeof ensureCampaignSheetWithHeaders !== 'function') {
        throw new Error('ensureCampaignSheetWithHeaders is not available.');
      }

      var sheet = ensureCampaignSheetWithHeaders(sheetName, headers);
      var existingCount = countDecemberQaEntries_(sheet, headers, year, targetMonth);
      if (existingCount > 0 && !config.force) {
        return {
          inserted: 0,
          skipped: existingCount,
          message: 'December QA data already exists. Pass force=true to seed again.'
        };
      }

      var rows = buildDecemberQaRows_(headers, year, targetMonth);
      if (!rows.length) {
        return { inserted: 0, skipped: existingCount, message: 'No QA rows generated.' };
      }

      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);

      return { inserted: rows.length, skipped: existingCount, message: 'December QA seed data inserted.' };
    };
  }

  if (typeof G.runTempQualitySeeder !== 'function') {
    G.runTempQualitySeeder = function runTempQualitySeeder() {
      return G.seedDecemberQualityDataIBTR({ force: true });
    };
  }

  function countDecemberQaEntries_(sheet, headers, year, targetMonth) {
    var dateIndex = headers.indexOf('CallDate');
    if (dateIndex === -1 || sheet.getLastRow() < 2) {
      return 0;
    }

    var data = sheet.getRange(2, dateIndex + 1, sheet.getLastRow() - 1, 1).getValues();
    var count = 0;
    for (var i = 0; i < data.length; i += 1) {
      var value = data[i][0];
      if (Object.prototype.toString.call(value) === '[object Date]') {
        if (value.getFullYear() === year && value.getMonth() === targetMonth) {
          count += 1;
        }
      }
    }
    return count;
  }

  function buildDecemberQaRows_(headers, year, targetMonth) {
    var rows = [];
    var questionCount = 19;
    var auditors = ['Quality Auditor', 'QA Team Lead', 'Quality Coach'];
    var feedbackOptions = ['Yes', 'No'];
    var overallFeedback = [
      'Great call handling with clear documentation.',
      'Solid compliance with a few improvement areas.',
      'Needs additional attention to policy reminders.',
      'Strong empathy and problem resolution.',
      'Follow-up steps were well communicated.'
    ];
    var agentFeedback = [
      'Keep up the strong client rapport.',
      'Remember to confirm next steps before ending calls.',
      'Great job keeping the call concise.',
      'Focus on summarizing the resolution.',
      'Nice consistency across checks.'
    ];
    var customerServiceNotes = [
      'Seeded QA review focused on basic customer service experience.',
      'General customer service review with consistent greeting and closure.',
      'Customer service behaviors were noted for follow-up coaching.',
      'Baseline customer service checks logged for December QA seeding.'
    ];

    for (var day = 1; day <= 31; day += 1) {
      var callDate = new Date(year, targetMonth, day);
      for (var agentIndex = 0; agentIndex < DECEMBER_QA_AGENTS.length; agentIndex += 1) {
        var agentName = DECEMBER_QA_AGENTS[agentIndex];
        var randomSeed = (day * 1000) + agentIndex;
        var random = createSeededRandom_(randomSeed);
        var clientName = pickRandomFrom_(DECEMBER_QA_CLIENTS, random) || 'CCBCU Resource Center';
        var callerName = pickRandomFrom_(DECEMBER_QA_AGENTS, random);
        var auditorName = pickRandomFrom_(auditors, random);
        var feedbackShared = pickRandomFrom_(feedbackOptions, random);
        var questionResults = buildQuestionResults_(questionCount, random);
        var totalScore = questionResults.totalYes;
        var percentage = parseFloat((totalScore / questionCount).toFixed(10));
        var callTimestamp = new Date(year, targetMonth, day, 9 + Math.floor(random() * 7), Math.floor(random() * 60));
        var auditDate = new Date(year, targetMonth, day, 14 + Math.floor(random() * 4), Math.floor(random() * 60));

        var row = [];
        for (var h = 0; h < headers.length; h += 1) {
          row.push(buildQaCell_({
            header: headers[h],
            agentName: agentName,
            callerName: callerName,
            clientName: clientName,
            auditorName: auditorName,
            feedbackShared: feedbackShared,
            questionResults: questionResults,
            totalScore: totalScore,
            percentage: percentage,
            callDate: callDate,
            callTimestamp: callTimestamp,
            auditDate: auditDate,
            caseLink: pickRandomFrom_(CASE_LINKS, random),
            callLink: pickRandomFrom_(CALL_LINKS, random),
            random: random,
            overallFeedback: overallFeedback,
            agentFeedback: agentFeedback,
            customerServiceNotes: customerServiceNotes
          }));
        }

        rows.push(row);
      }
    }

    return rows;
  }

  function buildQuestionResults_(questionCount, random) {
    var results = [];
    var totalYes = 0;
    for (var i = 1; i <= questionCount; i += 1) {
      var roll = random();
      var answer = 'Yes';
      if (roll < 0.1) {
        answer = 'N/A';
      } else if (roll < 0.25) {
        answer = 'No';
      }
      var note = '';
      if (answer === 'No') {
        note = 'Needs coaching on Q' + i + ' for customer service clarity.';
      } else if (answer === 'N/A') {
        note = 'Not applicable for this call.';
      }
      results.push({
        label: 'Q' + i,
        value: answer,
        note: note
      });
      if (answer === 'Yes') {
        totalYes += 1;
      }
    }
    return { results: results, totalYes: totalYes };
  }

  function buildQaCell_(context) {
    var header = context.header;
    if (header === 'ID') return Utilities.getUuid();
    if (header === 'Timestamp') return context.callTimestamp;
    if (header === 'CallerName') return context.callerName;
    if (header === 'AgentName') return context.agentName;
    if (header === 'AgentEmail') return buildAgentEmail_(context.agentName);
    if (header === 'ClientName') return context.clientName;
    if (header === 'CallDate') return context.callDate;
    if (header === 'CaseNumber') return context.caseLink || 'CASE-' + String(Math.floor(context.random() * 900000) + 100000);
    if (header === 'CallLink') return context.callLink || 'https://call.example.com/' + Utilities.getUuid();
    if (header === 'AuditorName') return context.auditorName;
    if (header === 'AuditDate') return context.auditDate;
    if (header === 'FeedbackShared') return context.feedbackShared;
    if (header.indexOf('Q') === 0 && header.indexOf('Note') === -1) {
      return getQuestionValue_(header, context.questionResults);
    }
    if (header.indexOf('Q') === 0 && header.indexOf('Note') > -1) {
      return getQuestionNote_(header, context.questionResults);
    }
    if (header === 'OverallFeedback') return pickRandomFrom_(context.overallFeedback, context.random);
    if (header === 'TotalScore') return context.totalScore;
    if (header === 'Percentage') return context.percentage;
    if (header === 'Notes') return pickRandomFrom_(context.customerServiceNotes, context.random);
    if (header === 'AgentFeedback') return pickRandomFrom_(context.agentFeedback, context.random);
    if (header === 'CoachingProvided') return context.totalScore < 15 ? 'Yes' : 'No';
    return '';
  }

  function getQuestionValue_(header, questionResults) {
    var index = parseInt(header.replace('Q', ''), 10) - 1;
    var entry = questionResults.results[index];
    return entry ? entry.value : '';
  }

  function getQuestionNote_(header, questionResults) {
    var number = header.replace('Q', '').replace(' Note', '');
    var index = parseInt(number, 10) - 1;
    var entry = questionResults.results[index];
    return entry ? entry.note : '';
  }

  function pickRandomFrom_(list, random) {
    if (!list || !list.length) return '';
    var index = Math.floor(random() * list.length);
    return list[index];
  }

  function buildAgentEmail_(agentName) {
    var normalized = String(agentName || '')
      .toLowerCase()
      .replace(/[^a-z\s-]/g, '')
      .replace(/\s+/g, '.');
    return normalized ? normalized + '@vlbpo.com' : 'agent@vlbpo.com';
  }

  function createSeededRandom_(seedValue) {
    var seed = 0;
    if (typeof seedValue === 'number') {
      seed = seedValue;
    } else if (seedValue) {
      var text = String(seedValue);
      for (var i = 0; i < text.length; i += 1) {
        seed = (seed << 5) - seed + text.charCodeAt(i);
        seed |= 0;
      }
    } else {
      seed = Date.now();
    }

    return function seededRandom() {
      seed = (seed * 9301 + 49297) % 233280;
      return seed / 233280;
    };
  }
})();
