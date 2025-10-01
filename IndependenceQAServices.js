/*******************************************************
 * Independence Insurance QA — Consolidated Backend
 * Single-file version: initialization, scoring, PDF,
 * audio upload, sheet writes, analytics + client APIs.
 *******************************************************/

/* =========================
   CONFIG (scoring & schema)
   ========================= */

const INDEPENDENCE_QA_CONFIG = {
    campaignName: "Independence Insurance",
    campaignId: "independence-insurance",
    version: "2.1",
    lastUpdated: "2025-01-12",
    scoring: {
        passThreshold: 85,
        excellentThreshold: 95,
        totalPossiblePoints: 57,
        criticalFailureOverride: true,
        weightedScoring: true
    },
    callTypes: [
        {value: "New Venture", label: "New Venture", description: "First-time business prospect"},
        {value: "Renewal", label: "Renewal", description: "Existing client renewal"}
    ],
    categories: {
        "Call Opening": {
            icon: "fas fa-phone",
            color: "#00BFFF",
            totalPoints: 8,
            description: "Professional greeting and call introduction",
            weight: 1.2,
            questions: [
                {
                    id: "Q1_ProfessionalGreeting",
                    text: "Greets the customer professionally and confirms decision maker",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Call Opening",
                    critical: true,
                    tags: ["greeting", "contact-verification", "critical"]
                },
                {
                    id: "Q2_ProperIntroduction",
                    text: "Introduces self/company and purpose (new authority/renewal)",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Call Opening",
                    critical: true,
                    tags: ["introduction", "company", "purpose", "critical"]
                },
                {
                    id: "Q3_ToneMatching",
                    text: "Matches tone; friendly, confident, natural",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Call Opening",
                    critical: false,
                    tags: ["tone", "natural", "rapport"]
                },
                {
                    id: "Q4_ConversationControl",
                    text: "Maintains control while building rapport",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Call Opening",
                    critical: false,
                    tags: ["control", "rapport", "conversation-management"]
                }
            ]
        },
        "Needs Discovery & Qualification": {
            icon: "fas fa-search",
            color: "#FFB800",
            totalPoints: 8,
            description: "Understanding business and qualification",
            weight: 1.1,
            questions: [
                {
                    id: "Q5_TruckingCompanyConfirmation",
                    text: "Confirms prospect operates a trucking company",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Needs Discovery & Qualification",
                    critical: false,
                    tags: ["business-type", "trucking", "verification"]
                },
                {
                    id: "Q6_BusinessOperationsVerification",
                    text: "Confirms nature of operations / type of goods hauled",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Needs Discovery & Qualification",
                    critical: false,
                    tags: ["operations", "goods", "understanding"]
                },
                {
                    id: "Q7_ValueFocusedLanguage",
                    text: "Uses value-focused language (reduce costs / save thousands)",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Needs Discovery & Qualification",
                    critical: false,
                    tags: ["value", "savings", "benefits"]
                },
                {
                    id: "Q8_ProperReinforcement",
                    text: "Reinforces with 'Awesome', 'Great', etc.",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Needs Discovery & Qualification",
                    critical: false,
                    tags: ["reinforcement", "positive"]
                }
            ]
        },
        "Appointment Setting": {
            icon: "fas fa-calendar-check",
            color: "#10b981",
            totalPoints: 20,
            description: "Scheduling & confirmation",
            weight: 1.5,
            questions: [
                {
                    id: "Q9_LiveTransferAttempt",
                    text: "Attempts live transfer before scheduling",
                    maxPoints: 5,
                    weight: 1.0,
                    category: "Appointment Setting",
                    critical: true,
                    tags: ["live-transfer", "critical"]
                },
                {
                    id: "Q10_AppointmentOffer",
                    text: "Offers appointment with Program Director if no transfer",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Appointment Setting",
                    critical: false,
                    tags: ["appointment", "program-director"]
                },
                {
                    id: "Q11_EmailConfirmation",
                    text: "Confirms/updates customer email",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Appointment Setting",
                    critical: false,
                    tags: ["email", "verification"]
                },
                {
                    id: "Q12_SchedulingLinkUsage",
                    text: "Uses scheduling link (Calendly) and confirms use",
                    maxPoints: 1,
                    weight: 1.0,
                    category: "Appointment Setting",
                    critical: false,
                    tags: ["calendly", "scheduling"]
                },
                {
                    id: "Q13_UrgencyAndConfidence",
                    text: "Creates urgency; offers soonest times (today/tomorrow)",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Appointment Setting",
                    critical: false,
                    tags: ["urgency", "confidence"]
                },
                {
                    id: "Q14_AppointmentConfirmation",
                    text: "Confirms exact date & time",
                    maxPoints: 5,
                    weight: 1.0,
                    category: "Appointment Setting",
                    critical: false,
                    tags: ["confirmation", "details"]
                }
            ]
        },
        "End of Call Procedure": {
            icon: "fas fa-handshake",
            color: "#9333ea",
            totalPoints: 5,
            description: "Close & follow-up setup",
            weight: 1.0,
            questions: [
                {
                    id: "Q15_AppointmentRecap",
                    text: "Recaps appointment time & agenda",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "End of Call Procedure",
                    critical: false,
                    tags: ["recap", "agenda"]
                },
                {
                    id: "Q16_EmailSMSExplanation",
                    text: "Explains what to expect via email/SMS",
                    maxPoints: 1,
                    weight: 1.0,
                    category: "End of Call Procedure",
                    critical: false,
                    tags: ["email", "sms", "expectations"]
                },
                {
                    id: "Q17_ProfessionalClosing",
                    text: "Ends the call clearly & professionally",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "End of Call Procedure",
                    critical: false,
                    tags: ["closing", "professional"]
                }
            ]
        },
        "Soft Skills & Compliance": {
            icon: "fas fa-user-check",
            color: "#FF4000",
            totalPoints: 16,
            description: "Comm skills & compliance",
            weight: 1.0,
            questions: [
                {
                    id: "Q18_AttentiveListening",
                    text: "Listens attentively; avoids interrupting",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Soft Skills & Compliance",
                    critical: false,
                    tags: ["listening", "communication"]
                },
                {
                    id: "Q19_ComplianceAccuracy",
                    text: "Avoids false promises; accurate and honest",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Soft Skills & Compliance",
                    critical: false,
                    tags: ["compliance", "accuracy", "honesty"]
                },
                {
                    id: "Q20_ConversationalDelivery",
                    text: "Conversational, fluent, clear (no script reading)",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Soft Skills & Compliance",
                    critical: false,
                    tags: ["conversational", "natural"]
                },
                {
                    id: "Q21_ClearSpeech",
                    text: "Speech clear/concise; minimal filler words",
                    maxPoints: 1,
                    weight: 1.0,
                    category: "Soft Skills & Compliance",
                    critical: false,
                    tags: ["clarity", "speech"]
                },
                {
                    id: "Q22_QuietEnvironment",
                    text: "Quiet professional environment",
                    maxPoints: 1,
                    weight: 1.0,
                    category: "Soft Skills & Compliance",
                    critical: false,
                    tags: ["environment", "quiet"]
                },
                {
                    id: "Q23_ObjectionHandling",
                    text: "Handles objections confidently/effectively",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Soft Skills & Compliance",
                    critical: false,
                    tags: ["objections", "confidence"]
                },
                {
                    id: "Q24_OverallProfessionalism",
                    text: "Overall confidence & professionalism",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Soft Skills & Compliance",
                    critical: false,
                    tags: ["professionalism"]
                }
            ]
        }
    }
};

/* =========================
   INITIALIZATION & LOGGING
   ========================= */

function initializeIndependenceQASystem() {
    try {
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);

        // QA sheet
        let qa = ss.getSheetByName(INDEPENDENCE_QA_SHEET);
        if (!qa) qa = ss.insertSheet(INDEPENDENCE_QA_SHEET);
        if (qa.getLastRow() === 0) {
            qa.getRange(1, 1, 1, INDEPENDENCE_QA_HEADERS.length)
                .setValues([INDEPENDENCE_QA_HEADERS])
                .setFontWeight('bold').setBackground('#003177').setFontColor('white');
            qa.setFrozenRows(1);
        }

        // Analytics sheet (lightweight append-based)
        let an = ss.getSheetByName(INDEPENDENCE_QA_ANALYTICS_SHEET);
        if (!an) {
            an = ss.insertSheet(INDEPENDENCE_QA_ANALYTICS_SHEET);
            const analyticsHeaders = [
                'Date', 'Period', 'Granularity', 'Agent', 'CallType',
                'TotalAssessments', 'AverageScore', 'PassRate', 'ExcellentRate',
                'CategoryScores', 'CriticalFailures', 'TrendData', 'CreatedAt'
            ];
            an.getRange(1, 1, 1, analyticsHeaders.length)
                .setValues([analyticsHeaders])
                .setFontWeight('bold').setBackground('#003177').setFontColor('white');
            an.setFrozenRows(1);
        }

        // Folders
        initializeDriveFolders();

        return {success: true, message: 'System initialized'};
    } catch (err) {
        safeWriteError('initializeIndependenceQASystem', err);
        throw err;
    }
}

function initializeDriveFolders() {
    try {
        let mainFolder;
        try {
            if (INDEPENDENCE_DRIVE_FOLDER_ID && INDEPENDENCE_DRIVE_FOLDER_ID.length > 10) {
                mainFolder = DriveApp.getFolderById(INDEPENDENCE_DRIVE_FOLDER_ID);
            } else {
                throw new Error('No valid folder ID');
            }
        } catch (_) {
            mainFolder = DriveApp.createFolder('Independence Insurance QA');
            // You may wish to update the constant above with mainFolder.getId()
        }
        ['Call Recordings', 'PDF Reports', 'Analytics'].forEach(name => {
            const it = mainFolder.getFoldersByName(name);
            if (!it.hasNext()) mainFolder.createFolder(name);
        });
    } catch (err) {
        // non-fatal
        safeWriteError('initializeDriveFolders', err);
    }
}

function safeWriteError(functionName, error) {
    try {
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
        let sh = ss.getSheetByName('ErrorLog');
        if (!sh) {
            sh = ss.insertSheet('ErrorLog');
            sh.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Error', 'Stack']]).setFontWeight('bold');
        }
        sh.appendRow([new Date(), functionName, (error && error.message) || String(error), (error && error.stack) || 'n/a']);
    } catch (_) {
        // swallow; console only
        console.error(functionName, error);
    }
}

/* =========================
   CORE HELPERS
   ========================= */

function generateIndependenceQAId() {
    const ts = Date.now();
    const rnd = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
    return `IND_QA_${ts}_${rnd}`;
}

// Shared data-preparation helpers (e.g., stripHtmlTags, extractAreasForImprovement,
// prepareIndependenceQAData) are defined in IndependenceQAUtilities.js to avoid
// duplicating logic across service files.

/* =========================
   SCORING
   ========================= */

function calculateIndependenceQAScores(formData) {
    try {
        if (!formData || typeof formData !== 'object' || !INDEPENDENCE_QA_CONFIG?.categories) {
            return createEmptyScoreResult('Invalid form data or configuration missing');
        }

        let totalEarned = 0;
        let totalPossible = 0;
        const categoryScores = {};
        const questionScores = {};
        let hasCriticalFailure = false;
        const criticalFailures = [];

        Object.entries(INDEPENDENCE_QA_CONFIG.categories).forEach(([catName, cat]) => {
            let cEarned = 0;
            let cPossible = 0;

            (cat.questions || []).forEach(q => {
                const qid = q.id;
                const max = Number(q.maxPoints) || 0;
                const ans = formData[qid] || 'NA';
                const qWeight = Number(q.weight) || 1.0;
                const cWeight = Number(cat.weight) || 1.0;
                const critical = !!q.critical;

                let earned = 0;
                let counts = true;

                if (ans === 'Yes') {
                    earned = max;
                } else if (ans === 'No') {
                    earned = 0;
                    if (critical) {
                        hasCriticalFailure = true;
                        criticalFailures.push({
                            questionId: qid,
                            questionText: q.text || '',
                            category: catName,
                            maxPoints: max
                        });
                    }
                } else { // 'NA' or blank
                    counts = false;
                }

                // store per-question
                questionScores[qid] = {
                    answer: ans,
                    earned,
                    possible: max,
                    percentage: max > 0 ? Math.round((earned / max) * 100) : 0,
                    weight: qWeight,
                    categoryWeight: cWeight,
                    critical,
                    countsTowardTotal: counts,
                    category: catName,
                    tags: q.tags || [],
                    comment: formData[qid + '_Comments'] || ''
                };

                if (counts) {
                    const wEarn = earned * qWeight * cWeight;
                    const wPoss = max * qWeight * cWeight;
                    cEarned += wEarn;
                    cPossible += wPoss;
                    totalEarned += wEarn;
                    totalPossible += wPoss;
                }
            });

            categoryScores[catName] = {
                earned: Math.round(cEarned * 100) / 100,
                possible: Math.round(cPossible * 100) / 100,
                percentage: cPossible > 0 ? Math.round((cEarned / cPossible) * 100) : 0,
                color: cat.color || '#6b7280',
                icon: cat.icon || 'fas fa-question',
                weight: cat.weight || 1.0,
                description: cat.description || ''
            };
        });

        let overallPercentage = totalPossible > 0 ? Math.round((totalEarned / totalPossible) * 100) : 0;

        if (hasCriticalFailure && INDEPENDENCE_QA_CONFIG.scoring.criticalFailureOverride) {
            overallPercentage = 0;
            totalEarned = 0;
        }

        let passStatus = 'Calculating...';
        let statusColor = '#6b7280';

        const passTh = INDEPENDENCE_QA_CONFIG.scoring.passThreshold || 85;
        const excTh = INDEPENDENCE_QA_CONFIG.scoring.excellentThreshold || 95;

        if (hasCriticalFailure) {
            passStatus = `Critical Failure - ${criticalFailures.length} critical item(s) failed`;
            statusColor = '#dc2626';
        } else if (overallPercentage >= excTh) {
            passStatus = 'Excellent Performance';
            statusColor = '#10b981';
        } else if (overallPercentage >= passTh) {
            passStatus = 'Meets Standards';
            statusColor = '#f59e0b';
        } else if (overallPercentage >= 70) {
            passStatus = 'Needs Improvement';
            statusColor = '#ef4444';
        } else {
            passStatus = 'Unsatisfactory Performance';
            statusColor = '#dc2626';
        }

        const strengthAreas = getStrengthAreas(categoryScores);
        const improvementAreas = getImprovementAreas(categoryScores, questionScores);

        return {
            success: true,
            totalEarned: Math.round(totalEarned * 100) / 100,
            totalPossible: Math.round(totalPossible * 100) / 100,
            overallPercentage,
            passStatus,
            statusColor,
            categoryScores,
            questionScores,
            hasCriticalFailure,
            criticalFailures,
            criticalFailureCount: criticalFailures.length,
            strengthAreas,
            improvementAreas,
            calculatedAt: new Date().toISOString(),
            configVersion: INDEPENDENCE_QA_CONFIG.version,
            scoringMethod: 'weighted_yes_no_na'
        };
    } catch (err) {
        safeWriteError('calculateIndependenceQAScores', err);
        return createEmptyScoreResult(`Calculation error: ${err.message}`);
    }
}

function createEmptyScoreResult(errorMessage) {
    const t = new Date().toISOString();
    return {
        success: false,
        error: errorMessage || 'Unknown error',
        totalEarned: 0,
        totalPossible: 0,
        overallPercentage: 0,
        passStatus: 'Error',
        statusColor: '#dc2626',
        categoryScores: {},
        questionScores: {},
        hasCriticalFailure: false,
        criticalFailures: [],
        criticalFailureCount: 0,
        strengthAreas: [],
        improvementAreas: {categories: [], questions: []},
        calculatedAt: t,
        configVersion: INDEPENDENCE_QA_CONFIG ? INDEPENDENCE_QA_CONFIG.version : 'unknown',
        scoringMethod: 'error'
    };
}

function getStrengthAreas(categoryScores) {
    try {
        return Object.entries(categoryScores || {})
            .filter(([, s]) => s?.percentage >= 90)
            .map(([name, s]) => ({category: name, percentage: s.percentage || 0, description: s.description || ''}))
            .sort((a, b) => (b.percentage || 0) - (a.percentage || 0));
    } catch (e) {
        return [];
    }
}

function getImprovementAreas(categoryScores, questionScores) {
    try {
        const lowCategories = [];
        const failedQuestions = [];

        Object.entries(categoryScores || {}).forEach(([name, s]) => {
            if ((s?.percentage ?? 100) < 85) lowCategories.push({
                type: 'category', name, percentage: s?.percentage || 0, description: s?.description || ''
            });
        });

        Object.entries(questionScores || {}).forEach(([id, s]) => {
            if (s?.answer === 'No') failedQuestions.push({
                type: 'question',
                id,
                category: s.category || 'Unknown',
                critical: !!s.critical,
                comment: s.comment || ''
            });
        });

        return {categories: lowCategories, questions: failedQuestions};
    } catch (e) {
        return {categories: [], questions: []};
    }
}

/* =========================
   DATA PREP & SHEET WRITE
   ========================= */

/* =========================
   DRIVE (FOLDERS & AUDIO)
   ========================= */

function getOrCreateIndependenceMainFolder() {
    try {
        try {
            return DriveApp.getFolderById(INDEPENDENCE_DRIVE_FOLDER_ID);
        } catch (_) {
            // fall through
        }
        const it = DriveApp.getFoldersByName('Independence Insurance QA');
        if (it.hasNext()) return it.next();
        const created = DriveApp.createFolder('Independence Insurance QA');
        created.createFolder('Call Recordings');
        created.createFolder('PDF Reports');
        created.createFolder('Analytics');
        return created;
    } catch (err) {
        safeWriteError('getOrCreateIndependenceMainFolder', err);
        throw err;
    }
}

function uploadIndependenceAudioFileToAssessment(audioBlob, assessmentId, assessmentFolder) {
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const orig = audioBlob.getName() || 'recording.mp3';
    const ext = orig.split('.').pop() || 'mp3';
    const name = `${assessmentId}_${ts}.${ext}`;
    const file = assessmentFolder.createFile(audioBlob.setName(name));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
}

/* =========================
   PDF GENERATION (HTML->PDF)
   ========================= */

function generateIndependenceQAPDF(qaData, scoreResults) {
    const html = createIndependenceQAPDFContentWithRichText(qaData, scoreResults);
    const pdfBlob = Utilities.newBlob(html, 'text/html', 'temp.html')
        .getAs('application/pdf')
        .setName(`Independence_QA_${qaData.ID}.pdf`);
    return {success: true, pdfBlob, pdfUrl: null};
}

function createIndependenceQAPDFContentWithRichText(qaData, scoreResults) {
    const overallFeedbackHtml = qaData.OverallFeedbackHtml || qaData.OverallFeedback || '';
    const richTextFeedback = convertHtmlToRichTextForPDF(overallFeedbackHtml);
    return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>
@page { margin: 0.75in; size: letter; }
body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; line-height:1.4; color:#333; margin:0; padding:0; }
.header { background: linear-gradient(135deg,#003177 0%,#004ba0 100%); color:#fff; padding:20px; border-radius:8px; margin-bottom:20px; display:flex; align-items:center; justify-content:space-between; }
.logo { height:60px; width:auto; }
.header-content { flex:1; margin-left:20px; }
.header h1 { margin:0; font-size:24px; font-weight:700; }
.header p { margin:5px 0 0 0; opacity:.9; font-size:14px; }
.assessment-info { background:#f8fafc; border:2px solid #003177; border-radius:8px; padding:20px; margin-bottom:20px; display:grid; grid-template-columns:repeat(2,1fr); gap:15px; }
.info-item { display:flex; flex-direction:column; }
.info-label { font-weight:600; color:#003177; font-size:12px; text-transform:uppercase; letter-spacing:.5px; margin-bottom:2px; }
.info-value { font-size:14px; color:#333; font-weight:500; }
.score-summary { background:${getScoreColor(scoreResults.overallPercentage, scoreResults.hasCriticalFailure)}; color:#fff; padding:20px; border-radius:8px; text-align:center; margin-bottom:20px; }
.score-summary h2 { margin:0 0 10px 0; font-size:28px; font-weight:700; }
.score-summary p { margin:0; font-size:16px; opacity:.9; }
.pass-status { display:inline-block; padding:8px 16px; background:rgba(255,255,255,.2); border-radius:20px; font-weight:600; margin-top:10px; font-size:14px; }
.categories-section { margin-bottom:25px; }
.category { margin-bottom:20px; border:1px solid #e5e7eb; border-radius:8px; overflow:hidden; }
.category-header { background:#003177; color:#fff; padding:12px 20px; font-weight:600; font-size:16px; display:flex; justify-content:space-between; align-items:center; }
.category-score { background:rgba(255,255,255,.2); padding:4px 12px; border-radius:12px; font-size:14px; }
.questions { padding:15px 20px; }
.question { margin-bottom:15px; padding-bottom:15px; border-bottom:1px solid #f1f5f9; }
.question:last-child { border-bottom:none; margin-bottom:0; padding-bottom:0; }
.question-text { font-weight:500; color:#374151; margin-bottom:8px; line-height:1.4; }
.question-text.critical { color:#dc2626; font-weight:600; }
.critical-badge { background:#dc2626; color:#fff; padding:2px 8px; border-radius:10px; font-size:10px; font-weight:700; text-transform:uppercase; margin-left:8px; }
.question-details { display:grid; grid-template-columns:auto 1fr auto; gap:15px; align-items:center; margin-top:8px; }
.answer { display:inline-block; padding:4px 12px; border-radius:12px; font-weight:600; font-size:12px; text-transform:uppercase; }
.answer.yes { background:#10b981; color:#fff; }
.answer.no  { background:#ef4444; color:#fff; }
.answer.na  { background:#6b7280; color:#fff; }
.points { font-weight:600; color:#003177; font-size:12px; }
.comment { color:#6b7280; font-size:12px; font-style:italic; margin-top:5px; }
.feedback-section { background:#f8fafc; border:2px solid #003177; border-radius:8px; padding:20px; margin-top:25px; }
.feedback-section h3 { color:#003177; margin:0 0 15px 0; font-size:18px; font-weight:600; }
.feedback-content { color:#374151; line-height:1.6; }
.feedback-content p { margin:0 0 12px 0; }
.feedback-content strong, .feedback-content b { font-weight:700; color:#1f2937; }
.feedback-content em, .feedback-content i { font-style:italic; }
.feedback-content ul, .feedback-content ol { margin:8px 0 16px 0; padding-left:24px; }
.feedback-content ul li { list-style-type:disc; margin-bottom:4px; }
.feedback-content ol li { list-style-type:decimal; margin-bottom:4px; }
.feedback-content h1, .feedback-content h2, .feedback-content h3 { color:#003177; font-weight:600; margin:16px 0 8px 0; }
.feedback-content h1 { font-size:18px; } .feedback-content h2 { font-size:16px; } .feedback-content h3 { font-size:14px; }
.footer { margin-top:30px; padding-top:20px; border-top:2px solid #e5e7eb; text-align:center; color:#6b7280; font-size:12px; }
@media print { .score-summary, .category-header, .answer { -webkit-print-color-adjust:exact; color-adjust:exact; } }
</style></head>
<body>
  <div class="header">
    <img src="${INDEPENDENCE_COMPANY_LOGO}" class="logo" loading="lazy"/>
    <div class="header-content">
      <h1>Independence Insurance Quality Assessment Report</h1>
      <p>Call Quality Evaluation & Performance Analysis</p>
    </div>
  </div>
  <div class="assessment-info">
    <div class="info-item"><div class="info-label">Assessment ID</div><div class="info-value">${qaData.ID || 'N/A'}</div></div>
    <div class="info-item"><div class="info-label">Call Date</div><div class="info-value">${qaData.CallDate || 'N/A'}</div></div>
    <div class="info-item"><div class="info-label">Agent Name</div><div class="info-value">${qaData.AgentName || 'N/A'}</div></div>
    <div class="info-item"><div class="info-label">Caller Name</div><div class="info-value">${qaData.CallerName || 'N/A'}</div></div>
    <div class="info-item"><div class="info-label">Auditor</div><div class="info-value">${qaData.AuditorName || 'N/A'}</div></div>
    <div class="info-item"><div class="info-label">Audit Date</div><div class="info-value">${qaData.AuditDate || 'N/A'}</div></div>
    <div class="info-item"><div class="info-label">Call Type</div><div class="info-value">${qaData.CallType || 'N/A'}</div></div>
    <div class="info-item"><div class="info-label">Agent Email</div><div class="info-value">${qaData.AgentEmail || 'N/A'}</div></div>
  </div>
  <div class="score-summary">
    <h2>${scoreResults.overallPercentage || 0}%</h2>
    <p>Overall Score (${scoreResults.totalEarned || 0}/${scoreResults.totalPossible || 0} points)</p>
    <div class="pass-status">${scoreResults.passStatus || 'Calculating...'}</div>
  </div>
  <div class="categories-section">
    <h2 style="color:#003177;margin-bottom:20px;font-size:20px;">Category Breakdown</h2>
    ${generateCategoryBreakdown(scoreResults, qaData)}
  </div>
  ${richTextFeedback ? `
  <div class="feedback-section">
    <h3>Overall Feedback & Recommendations</h3>
    <div class="feedback-content">${richTextFeedback}</div>
  </div>` : ''}
  <div class="footer">
    <p>Generated by VLBPO LuminaHQ • Independence Insurance QA System</p>
    <p>Report generated on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy 'at' HH:mm:ss z")}</p>
  </div>
</body></html>`;
}

function convertHtmlToRichTextForPDF(htmlContent) {
    if (!htmlContent) return '';
    let rich = htmlContent
        .replace(/<div>/g, '<p>').replace(/<\/div>/g, '</p>')
        .replace(/<p><\/p>/g, '').replace(/&nbsp;/g, ' ')
        .replace(/<br\s*\/?>/g, '<br>').trim();
    if (rich && !rich.startsWith('<') && !rich.includes('<')) rich = `<p>${rich}</p>`;
    return rich;
}

function generateCategoryBreakdown(scoreResults, qaData) {
    if (!INDEPENDENCE_QA_CONFIG?.categories) return '<p>Category data not available.</p>';
    let html = '';
    Object.entries(INDEPENDENCE_QA_CONFIG.categories).forEach(([catName, cat]) => {
        const cs = scoreResults.categoryScores[catName] || {earned: 0, possible: 0, percentage: 0};
        html += `
    <div class="category">
      <div class="category-header">
        <span>${catName}</span>
        <div class="category-score">${Math.round(cs.earned)}/${Math.round(cs.possible)} pts</div>
      </div>
      <div class="questions">`;
        (cat.questions || []).forEach(q => {
            const ans = qaData[q.id] || 'NA';
            const cmt = qaData[q.id + '_Comments'] || '';
            const max = q.maxPoints || 0;
            const earned = ans === 'Yes' ? max : 0;
            html += `
      <div class="question">
        <div class="question-text ${q.critical ? 'critical' : ''}">
          ${q.text}${q.critical ? '<span class="critical-badge">Critical</span>' : ''}
        </div>
        <div class="question-details">
          <div class="answer ${ans.toLowerCase()}">${ans}</div>
          <div class="comment">${cmt || 'No comment provided'}</div>
          <div class="points">${earned}/${max} pts</div>
        </div>
      </div>`;
        });
        html += `</div></div>`;
    });
    return html;
}

function getScoreColor(percentage, hasCriticalFailure) {
    if (hasCriticalFailure) return '#dc2626';
    if ((percentage || 0) >= 95) return '#10b981';
    if ((percentage || 0) >= 85) return '#f59e0b';
    return '#ef4444';
}

/* =========================
   ANALYTICS (light append)
   ========================= */

function updateIndependenceQAAnalytics(qaData, scoreResults) {
    try {
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
        let an = ss.getSheetByName(INDEPENDENCE_QA_ANALYTICS_SHEET);
        if (!an) {
            an = ss.insertSheet(INDEPENDENCE_QA_ANALYTICS_SHEET);
            const headers = [
                'Date', 'Period', 'Granularity', 'Agent', 'CallType',
                'TotalAssessments', 'AverageScore', 'PassRate', 'ExcellentRate',
                'CategoryScores', 'CriticalFailures', 'TrendData', 'CreatedAt'
            ];
            an.getRange(1, 1, 1, headers.length)
                .setValues([headers])
                .setFontWeight('bold').setBackground('#003177').setFontColor('white');
            an.setFrozenRows(1);
        }
        const now = new Date();
        const today = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

        an.appendRow([
            today, 'Daily', 'Day', qaData.AgentName, qaData.CallType,
            1,
            scoreResults.overallPercentage,
            scoreResults.hasCriticalFailure ? 0 : (scoreResults.overallPercentage >= 85 ? 100 : 0),
            scoreResults.overallPercentage >= 95 ? 100 : 0,
            JSON.stringify(scoreResults.categoryScores || {}),
            scoreResults.hasCriticalFailure ? scoreResults.criticalFailures.length : 0,
            JSON.stringify({date: today, score: scoreResults.overallPercentage}),
            now
        ]);
    } catch (err) {
        // non-fatal
        safeWriteError('updateIndependenceQAAnalytics', err);
    }
}

/* =========================
   SUBMISSION PIPELINE (API)
   ========================= */

function submitIndependenceQAWithAudio(formData) {
    try {
        // ensure env
        initializeIndependenceQASystem();

        // ID + scores
        const assessmentId = generateIndependenceQAId();
        const scores = calculateIndependenceQAScores(formData);

        // create assessment folder inside main
        const mainFolder = getOrCreateIndependenceMainFolder();
        const agentSafe = (formData.agentName || 'Unknown').replace(/[^a-zA-Z0-9_-]/g, '_');
        const dateSafe = (formData.callDate || '').replace(/[^0-9-]/g, '');
        const folderName = `${assessmentId}_${agentSafe}_${dateSafe}`;
        const assessmentFolder = mainFolder.createFolder(folderName);

        // audio (optional)
        let audioUrl = '';
        if (formData.audioFile?.bytes) {
            const blob = Utilities.newBlob(formData.audioFile.bytes, formData.audioFile.mimeType || 'application/octet-stream', formData.audioFile.name || 'recording.mp3');
            audioUrl = uploadIndependenceAudioFileToAssessment(blob, assessmentId, assessmentFolder);
        }

        // prepare row + write
        const qaData = prepareIndependenceQAData(formData, scores, assessmentId, audioUrl);
        const writeRes = saveIndependenceQAToSheet(qaData);
        if (!writeRes.success) throw new Error(writeRes.error || 'Sheet save failed');

        // PDF
        let pdfUrl = '';
        try {
            const pdf = generateIndependenceQAPDF(qaData, scores);
            if (pdf?.pdfBlob) {
                const pdfFile = assessmentFolder.createFile(pdf.pdfBlob.setName('Independence_QA_Assessment.pdf'));
                pdfFile.setDescription(`Independence Insurance QA Report for ${qaData.AgentName} - ${qaData.CallDate}`);
                pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                pdfUrl = pdfFile.getUrl();
            }
        } catch (e) {
            safeWriteError('generateIndependenceQAPDF', e);
        }

        // analytics (non-blocking)
        updateIndependenceQAAnalytics(qaData, scores);

        return {
            success: true,
            message: 'Independence QA assessment submitted successfully!',
            assessmentId,
            folderName,
            folderUrl: assessmentFolder.getUrl(),
            audioUrl,
            pdfUrl,
            scoreResults: scores
        };
    } catch (err) {
        safeWriteError('submitIndependenceQAWithAudio', err);
        return {success: false, error: err.message || 'Failed to submit QA assessment'};
    }
}

/* =========================
   CLIENT-FACING WRAPPERS
   ========================= */

function clientSubmitIndependenceQAWithAudio(formData) {
    try {
        return submitIndependenceQAWithAudio(formData);
    } catch (e) {
        safeWriteError('clientSubmitIndependenceQAWithAudio', e);
        return {success: false, error: e.message};
    }
}

function clientPreviewIndependenceQAScore(formData) {
    try {
        if (!formData || typeof formData !== 'object') {
            return createEmptyScoreResult('Invalid form data provided to preview function');
        }
        const res = calculateIndependenceQAScores(formData);
        return res || createEmptyScoreResult('Score calculation returned null');
    } catch (e) {
        safeWriteError('clientPreviewIndependenceQAScore', e);
        return createEmptyScoreResult(`Preview error: ${e.message}`);
    }
}

function clientGetIndependenceQAConfig() {
    try {
        return JSON.parse(JSON.stringify(INDEPENDENCE_QA_CONFIG));
    } catch (e) {
        return createMinimalFallbackConfig();
    }
}

function clientGetIndependenceCallTypes() {
    try {
        if (!INDEPENDENCE_QA_CONFIG?.callTypes) throw new Error('no callTypes');
        return INDEPENDENCE_QA_CONFIG.callTypes.map(ct => ({
            ...ct,
            campaignSpecific: true,
            lastUpdated: INDEPENDENCE_QA_CONFIG.lastUpdated
        }));
    } catch (_) {
        return [
            {value: 'New Venture', label: 'New Venture', description: 'First-time business prospect'},
            {value: 'Renewal', label: 'Renewal', description: 'Existing client renewal'}
        ];
    }
}

/* ---------- Helpers (self-contained, no external dependencies) ---------- */

function readUsers_(ss) {
  const sh = ss.getSheetByName('Users');
  if (!sh) return {};
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {};

  const h = toIndex_(values[0]);
  const rows = values.slice(1);
  const map = {};

  rows.forEach(r => {
    const email = String(val(r, h, ['Email'])).trim().toLowerCase();
    if (!email) return;

    const activeCell = val(r, h, ['Active']);
    const isActive = activeCell === true ||
                     String(activeCell).toLowerCase() === 'true' ||
                     activeCell === 1 ||
                     String(activeCell).toLowerCase() === 'yes';

    const campaignsRaw = String(val(r, h, ['CampaignIds','CampaignId']) || '');
    const campaignIds = campaignsRaw
      ? campaignsRaw.split(/[,\s]+/).map(s => s.trim()).filter(Boolean)
      : [];

    map[email] = {
      id: String(val(r, h, ['UserId','ID']) || '').trim(),
      name: String(val(r, h, ['Name','FullName']) || '').trim() || email,
      email,
      department: String(val(r, h, ['Department']) || '').trim(),
      role: String(val(r, h, ['Role']) || '').trim(),
      managerEmail: String(val(r, h, ['ManagerEmail','Manager']) || '').trim().toLowerCase(),
      campaignIds,
      active: isActive
    };
  });

  return map;
}

function getAssignedEmails_(ss, managerEmail, campaignId, usersByEmail) {
  const out = new Set();
  const assign = ss.getSheetByName('ManagerAssignments');

  if (assign) {
    // Use ManagerAssignments if present
    const values = assign.getDataRange().getValues();
    if (values.length > 1) {
      const h = toIndex_(values[0]);
      values.slice(1).forEach(r => {
        const mgr = String(val(r, h, ['ManagerEmail','ManagerId']) || '').trim().toLowerCase();
        const user = String(val(r, h, ['UserEmail','UserId']) || '').trim().toLowerCase();
        const camp = String(val(r, h, ['CampaignId']) || '').trim();
        if (!mgr || !user) return;
        if (mgr !== managerEmail) return;
        if (campaignId && camp && camp !== String(campaignId)) return;
        if (usersByEmail[user] && usersByEmail[user].active) out.add(user);
      });
    }
  }

  if (out.size === 0) {
    // Fallback: infer by Users.ManagerEmail (and campaign filter if provided)
    Object.values(usersByEmail).forEach(u => {
      if (!u.active) return;
      const okMgr = u.managerEmail && u.managerEmail === managerEmail;
      const okCamp = !campaignId || (u.campaignIds || []).includes(String(campaignId));
      if (okMgr && okCamp) out.add(u.email);
    });
  }

  return out;
}

/* ---------- Small utilities ---------- */
function toIndex_(headerRow) {
  const idx = {};
  headerRow.forEach((col, i) => { idx[String(col).trim()] = i; idx[String(col).toLowerCase()] = i; });
  return idx;
}
function val(row, idx, keys) {
  for (const k of keys) {
    if (k in idx) return row[idx[k]];
    const lk = k.toLowerCase();
    if (lk in idx) return row[idx[lk]];
  }
  return '';
}

function clientPerformIndependenceHealthCheck() {
    try {
        const t0 = Date.now();
        const test = calculateIndependenceQAScores({
            Q1_ProfessionalGreeting: 'Yes',
            Q2_ProperIntroduction: 'Yes',
            Q3_ToneMatching: 'No'
        });
        const ok = !!(test && typeof test.overallPercentage === 'number');

        const res = {
            success: true,
            status: ok ? 'healthy' : 'warning',
            responseTime: (Date.now() - t0) + 'ms',
            components: {
                spreadsheet: {status: 'healthy', message: 'Spreadsheet connection available'},
                configuration: {
                    status: INDEPENDENCE_QA_CONFIG ? 'healthy' : 'warning',
                    message: INDEPENDENCE_QA_CONFIG ? 'Configuration loaded' : 'Using fallback',
                    categories: INDEPENDENCE_QA_CONFIG ? Object.keys(INDEPENDENCE_QA_CONFIG.categories).length : 0,
                    questions: INDEPENDENCE_QA_CONFIG ? Object.values(INDEPENDENCE_QA_CONFIG.categories).reduce((n, cat) => n + (cat.questions ? cat.questions.length : 0), 0) : 0
                },
                scoring: {
                    status: ok ? 'healthy' : 'error',
                    message: ok ? 'Scoring OK' : 'Scoring issue',
                    version: INDEPENDENCE_QA_CONFIG?.version || 'unknown',
                    testScore: test ? test.overallPercentage : 'failed'
                }
            },
            timestamp: new Date().toISOString()
        };
        return res;
    } catch (e) {
        safeWriteError('clientPerformIndependenceHealthCheck', e);
        return {success: false, status: 'error', error: e.message, timestamp: new Date().toISOString()};
    }
}

function clientGetIndependenceSystemStatus() {
    try {
        const cfg = INDEPENDENCE_QA_CONFIG;
        const totalQuestions = Object.values(cfg.categories).reduce((n, cat) => n + (cat.questions ? cat.questions.length : 0), 0);
        const criticalQuestions = Object.values(cfg.categories).reduce((n, cat) => n + (cat.questions ? cat.questions.filter(q => q.critical).length : 0), 0);
        return {
            configVersion: cfg.version,
            lastUpdated: cfg.lastUpdated,
            totalQuestions,
            criticalQuestions,
            categories: Object.keys(cfg.categories).length,
            callTypes: (cfg.callTypes || []).length,
            scoringMethod: 'weighted_yes_no_na',
            passThreshold: cfg.scoring.passThreshold,
            excellentThreshold: cfg.scoring.excellentThreshold,
            status: 'healthy'
        };
    } catch (e) {
        safeWriteError('clientGetIndependenceSystemStatus', e);
        return {error: e.message, status: 'error'};
    }
}

/* =========================
   FALLBACK CONFIG (rare)
   ========================= */

function createMinimalFallbackConfig() {
    return {
        categories: {
            "Basic Assessment": {
                icon: "fas fa-question",
                color: "#00BFFF",
                totalPoints: 5,
                description: "Basic evaluation criteria",
                weight: 1.0,
                questions: [{
                    id: "Q1_BasicQuestion",
                    text: "Agent performed adequately",
                    maxPoints: 5,
                    weight: 1.0,
                    critical: false,
                    category: "Basic Assessment",
                    tags: ["basic"]
                }]
            }
        },
        callTypes: [
            {value: "New Venture", label: "New Venture", description: "First-time business prospect"},
            {value: "Renewal", label: "Renewal", description: "Existing client renewal"}
        ],
        scoring: {
            passThreshold: 85,
            excellentThreshold: 95,
            totalPossiblePoints: 5,
            criticalFailureOverride: true,
            weightedScoring: false
        },
        version: "fallback-1.0", lastUpdated: new Date().toISOString().split('T')[0]
    };
}

/* =========================
   DEV: Quick validation
   ========================= */

(function autoValidateOnLoad() {
    try {
        // Touch the constants so issues surface early in logs
        Logger.log('Independence QA loaded. Headers: ' + (INDEPENDENCE_QA_HEADERS?.length || 0));
    } catch (e) {
        // ignore
    }
})();

function clientGetIndependenceQAAnalytics(granularity, periodKey, agentFilter, callTypeFilter) {
  const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
  const sh = ss.getSheetByName(INDEPENDENCE_QA_SHEET);
  if (!sh) throw new Error('Independence QA sheet not found');

  const values = sh.getDataRange().getValues();
  if (!values.length) return emptyAnalytics();

  const headers = values[0];
  const rows = values.slice(1);
  const idx = indexHeaders_(headers);

  // Period handling
  const key = periodKey || periodKeyFromToday_(granularity);
  const range = periodKeyToRange_(granularity, key);

  // Current-period rows
  const currentRows = filterRows_(rows, idx, range, agentFilter, callTypeFilter);

  // KPI metrics
  const kpi = computeKpis_(currentRows, idx);

  // Trends (last N periods ending at selected period)
  const n = granularity === 'Year' ? 5 : granularity === 'Quarter' ? 8 : granularity === 'Month' ? 12 : 8;
  const keys = getLastNPeriods_(n, key, granularity);
  const trendsVals = keys.map(k => {
    const r = periodKeyToRange_(granularity, k);
    const rRows = filterRows_(rows, idx, r, agentFilter, callTypeFilter);
    return averageScore_(rRows, idx);
  });

  // Category performance for the selected period
  const categories = computeCategories_(currentRows, idx);

  // Agent comparison for the selected period
  const agentAgg = aggregateByAgent_(currentRows, idx);
  const agentLabels = Object.keys(agentAgg);
  const agentValues = agentLabels.map(a => Math.round(agentAgg[a].sum / Math.max(1, agentAgg[a].count)));

  // Deltas vs previous period
  const prevKey = previousPeriodKey_(key, granularity);
  const prevRange = periodKeyToRange_(granularity, prevKey);
  const prevRows = filterRows_(rows, idx, prevRange, agentFilter, callTypeFilter);
  const prevKpi = computeKpis_(prevRows, idx);

  return {
    avgScore: kpi.avgScore,
    passRate: kpi.passRate,
    totalEvaluations: kpi.total,
    agentsEvaluated: kpi.agents,
    avgScoreChange: kpi.avgScore - prevKpi.avgScore,
    passRateChange: kpi.passRate - prevKpi.passRate,
    evaluationsChange: pctChange_(kpi.total, prevKpi.total),
    agentsChange: pctChange_(kpi.agents, prevKpi.agents),
    trends: { labels: keys, values: trendsVals },
    categories: { labels: categories.labels, values: categories.values },
    agents: { labels: agentLabels, values: agentValues }
  };
}

function exportIndependenceQAData(granularity, periodKey, agentFilter, callTypeFilter) {
  const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
  const sh = ss.getSheetByName(INDEPENDENCE_QA_SHEET);
  if (!sh) throw new Error('Independence QA sheet not found');

  const values = sh.getDataRange().getValues();
  if (!values.length) return [];

  const headers = values[0];
  const rows = values.slice(1);
  const idx = indexHeaders_(headers);

  const key = periodKey || periodKeyFromToday_(granularity);
  const range = periodKeyToRange_(granularity, key);
  const filtered = filterRows_(rows, idx, range, agentFilter, callTypeFilter);

  // Return a tidy array of objects (the front-end turns this into CSV)
  return filtered.map(r => ({
    ID: r[idx.ID] || '',
    AuditDate: fmtDate_(pickDate_(r[idx.AuditDate], r[idx.CallDate], r[idx.Timestamp])),
    AgentName: r[idx.AgentName] || '',
    AgentEmail: r[idx.AgentEmail] || '',
    CallType: r[idx.CallType] || '',
    PercentageScore: num_(r[idx.PercentageScore]),
    PassStatus: r[idx.PassStatus] || '',
    CallerName: r[idx.CallerName] || '',
    AudioURL: r[idx.AudioURL] || ''
  }));
}

function emptyAnalytics() {
  return {
    avgScore: 0, passRate: 0, totalEvaluations: 0, agentsEvaluated: 0,
    avgScoreChange: 0, passRateChange: 0, evaluationsChange: 0, agentsChange: 0,
    trends: { labels: ['No Data'], values: [0] },
    categories: { labels: ['No Data'], values: [0] },
    agents: { labels: ['No Data'], values: [0] }
  };
}

function indexHeaders_(headers) {
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  // required columns (sheet uses these exact names)
  ['ID','AuditDate','CallDate','Timestamp','AgentName','AgentEmail','CallType','PercentageScore','PassStatus','CallerName','AudioURL']
    .forEach(h => { if (!(h in idx)) idx[h] = -1; });
  // also index all question ids so categories can be computed
  if (typeof INDEPENDENCE_QA_CONFIG !== 'undefined') {
    Object.values(INDEPENDENCE_QA_CONFIG.categories || {}).forEach(cat => {
      (cat.questions || []).forEach(q => {
        if (!(q.id in idx)) idx[q.id] = headers.indexOf(q.id);
      });
    });
  }
  return idx;
}

function num_(v){ var n = Number(v); return isFinite(n) ? n : 0; }

function fmtDate_(d) {
  if (!d) return '';
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function pickDate_(/* AuditDate, CallDate, Timestamp */ a, b, c) {
  return coerceDate_(a) || coerceDate_(b) || coerceDate_(c);
}

function coerceDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return v;
  if (typeof v === 'number') return new Date(v);
  if (typeof v === 'string') {
    // try yyyy-MM-dd first
    const m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    const t = Date.parse(v);
    if (!isNaN(t)) return new Date(t);
  }
  return null;
}

function filterRows_(rows, idx, range, agentFilter, callTypeFilter) {
  return rows.filter(r => {
    const d = pickDate_(r[idx.AuditDate], r[idx.CallDate], r[idx.Timestamp]);
    if (!d || d < range.start || d >= range.end) return false;
    if (agentFilter && String(r[idx.AgentName] || '').trim() !== agentFilter) return false;
    if (callTypeFilter && String(r[idx.CallType] || '').trim() !== callTypeFilter) return false;
    return true;
  });
}

function averageScore_(rows, idx) {
  if (!rows.length) return 0;
  const s = rows.reduce((a,r)=> a + num_(r[idx.PercentageScore]), 0);
  return Math.round(s / rows.length);
}

function computeKpis_(rows, idx) {
  const total = rows.length;
  const avgScore = averageScore_(rows, idx);
  const passOk = r => {
    const s = String(r[idx.PassStatus] || '');
    return s === 'Excellent Performance' || s === 'Meets Standards';
  };
  const passRate = total ? Math.round(100 * rows.filter(passOk).length / total) : 0;
  const agents = new Set(rows.map(r => String(r[idx.AgentName] || '').trim()).filter(Boolean)).size;
  return { total, avgScore, passRate, agents };
}

function pctChange_(cur, prev) {
  if (!prev) return cur ? 100 : 0;
  return Math.round(((cur - prev) / prev) * 1000) / 10; // one decimal
}

// --- Period math (server-side mirror of the client helpers) ---
function periodKeyFromToday_(granularity) {
  const now = new Date();
  if (granularity === 'Year') return String(now.getFullYear());
  if (granularity === 'Month') return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}`;
  if (granularity === 'Quarter') {
    const q = Math.floor(now.getMonth()/3) + 1;
    return `Q${q}-${now.getFullYear()}`;
  }
  return isoWeekKey_(now);
}

function isoWeekKey_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2,'0')}`;
}

function isoWeekStart_(isoYear, isoWeek) {
  const simple = new Date(Date.UTC(isoYear, 0, 1 + (isoWeek - 1) * 7));
  const dow = simple.getUTCDay();
  const start = new Date(simple);
  if (dow <= 4) start.setUTCDate(simple.getUTCDate() - simple.getUTCDay() + 1);
  else start.setUTCDate(simple.getUTCDate() + 8 - simple.getUTCDay());
  return start; // Monday 00:00 UTC
}

function periodKeyToRange_(granularity, key) {
  const tz = Session.getScriptTimeZone();
  if (granularity === 'Year') {
    const y = Number(key);
    const start = new Date(y, 0, 1);
    const end = new Date(y + 1, 0, 1);
    return { start, end };
  }
  if (granularity === 'Month') {
    const [y,m] = key.split('-').map(Number);
    const start = new Date(y, m - 1, 1);
    const end = new Date(y, m, 1);
    return { start, end };
  }
  if (granularity === 'Quarter') {
    const m = key.match(/^Q([1-4])-(\d{4})$/);
    if (!m) throw new Error('Bad quarter key: ' + key);
    const q = Number(m[1]), y = Number(m[2]);
    const startMonth = (q - 1) * 3;
    const start = new Date(y, startMonth, 1);
    const end = new Date(y, startMonth + 3, 1);
    return { start, end };
  }
  // Week (YYYY-W##)
  const wm = key.match(/^(\d{4})-W(\d{2})$/);
  if (!wm) throw new Error('Bad week key: ' + key);
  const y = Number(wm[1]), w = Number(wm[2]);
  const start = isoWeekStart_(y, w);
  const end = new Date(start); end.setUTCDate(end.getUTCDate() + 7);
  // convert to local since sheet dates are local
  return { start: new Date(start.getFullYear(), start.getMonth(), start.getDate()),
           end:   new Date(end.getFullYear(),   end.getMonth(),   end.getDate()) };
}

function previousPeriodKey_(key, granularity) {
  if (granularity === 'Year') return String(Number(key) - 1);
  if (granularity === 'Month') {
    const [y,m] = key.split('-').map(Number);
    const d = new Date(y, m - 1, 1); d.setMonth(d.getMonth() - 1);
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;
  }
  if (granularity === 'Quarter') {
    const m = key.match(/^Q([1-4])-(\d{4})$/);
    let q = Number(m[1]), y = Number(m[2]); q -= 1; if (q < 1) { q = 4; y -= 1; }
    return `Q${q}-${y}`;
  }
  const wm = key.match(/^(\d{4})-W(\d{2})$/);
  const y = Number(wm[1]), w = Number(wm[2]);
  const start = isoWeekStart_(y, w); start.setUTCDate(start.getUTCDate() - 7);
  return isoWeekKey_(new Date(start.getFullYear(), start.getMonth(), start.getDate()));
}

function getLastNPeriods_(n, endKey, granularity) {
  const out = [];
  let k = endKey;
  for (let i=0;i<n;i++) { out.unshift(k); k = previousPeriodKey_(k, granularity); }
  return out;
}

// --- Categories from config ---
function computeCategories_(rows, idx) {
  const labels = [], values = [];
  if (!rows.length || !INDEPENDENCE_QA_CONFIG || !INDEPENDENCE_QA_CONFIG.categories) {
    return { labels: ['No Data'], values: [0] };
  }

  Object.entries(INDEPENDENCE_QA_CONFIG.categories).forEach(([catName, cat]) => {
    let earned = 0, possible = 0;
    (cat.questions || []).forEach(q => {
      const col = idx[q.id];
      if (col < 0) return;
      const max = Number(q.maxPoints) || 0;
      rows.forEach(r => {
        const ans = (r[col] || 'NA').toString();
        if (ans === 'Yes') { earned += max; possible += max; }
        else if (ans === 'No') { possible += max; }
      });
    });
    labels.push(catName);
    values.push(possible ? Math.round(100 * earned / possible) : 0);
  });
  return { labels, values };
}

function aggregateByAgent_(rows, idx) {
  const map = {};
  rows.forEach(r => {
    const a = (r[idx.AgentName] || '').toString().trim() || 'Unknown';
    const s = num_(r[idx.PercentageScore]);
    if (!map[a]) map[a] = { sum: 0, count: 0 };
    map[a].sum += s; map[a].count += 1;
  });
  return map;
}
