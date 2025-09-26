/***************************************************************
 * TGC_QAService.gs  (CHAT EDITION)
 * Chat QA logic: numeric scoring (0..max, N/A), Zero Tolerance,
 * CRUD, PDF summary, and dashboard stats.
 * Depends on: TGCUtilities.gs  (TGC.Util.*)
 ***************************************************************/
var TGC = this.TGC || {};
TGC.QA = (function () {
    /* --------------------- CHAT QA MODEL --------------------- */
    // Per-question maximums (MUST match client)
    var Q_MAX = {
        q1: 5,  // Thanked customer / response
        q2: 5,  // Understood concern
        q3: 5,  // Empathy
        q4: 10, // Attention
        q5: 50, // Appropriate answer + correct forms
        q6: 10, // Offers further assistance / awaiting response
        q7: 5,  // Customized email/macro
        q8: 5,  // Merge tickets when appropriate
        q9: 5   // Correct ticket status
    };

    // Category grouping (MUST match client)
    var CATS = {
        Professionalism: ['q1', 'q3', 'q4'],
        Comprehension: ['q2'],
        Resolution: ['q5'],
        FollowUp: ['q6'],
        Writing: ['q7'],
        Process: ['q8', 'q9']
    };

    // Question text (for PDFs/exports)
    var Q_TEXT = {
        q1: 'Did the agent thank the customer for reaching out or for their response?',
        q2: 'Did the agent effectively understand the user’s concern?',
        q3: 'Did the agent empathize with the customer?',
        q4: 'Did the agent display complete attention to the user’s situation/concern?',
        q5: 'Did the agent provide an appropriate answer to the customer’s issue and fill out the correct form(s)?',
        q6: 'The agent offers further assistance or advises the customer they are awaiting a response.',
        q7: 'Did the agent customize the email response and modify the macro to address the specific needs of the customer?',
        q8: 'Did the agent merge the tickets when appropriate?',
        q9: 'Did the agent place the ticket in the correct status after responding?'
    };

    var PASS_THRESHOLD = 85; // %
    var CHANNEL = 'Chat';

    /* --------------------- HELPERS --------------------- */
    function parseScore(val, key) {
        // Accept 'NA' (case-insensitive) or null/'' as N/A
        if (val == null) return null;
        var s = String(val).trim();
        if (!s || /^na$/i.test(s)) return null;

        var n = Number(s);
        if (!isFinite(n)) return null;
        var max = Q_MAX[key] || 0;
        if (n < 0) n = 0;
        if (n > max) n = max;
        return Math.round(n);
    }

    function computeTotals(scoresObj, zeroTolerance) {
        var earned = 0, possible = 0;
        Object.keys(Q_MAX).forEach(function (k) {
            var v = parseScore(scoresObj[k], k);
            if (v === null) return; // N/A excluded
            earned += v;
            possible += Q_MAX[k];
        });

        var pct = (possible > 0) ? Math.round((earned / possible) * 100) : 0;
        if (zeroTolerance === true) pct = 0;

        var passStatus = zeroTolerance ? 'Auto Fail' : (pct >= PASS_THRESHOLD ? 'Pass' : 'Fail');
        return {earned: earned, possible: possible, percentage: pct, passStatus: passStatus};
    }

    function computeCategoryBreakdown(scoresObj) {
        var out = {};
        Object.keys(CATS).forEach(function (cat) {
            var keys = CATS[cat];
            var earned = 0, possible = 0;
            keys.forEach(function (k) {
                var v = parseScore(scoresObj[k], k);
                if (v === null) return; // N/A exclude
                earned += v;
                possible += Q_MAX[k];
            });
            out[cat] = (possible > 0) ? Math.round((earned / possible) * 100) : 0;
        });
        return out;
    }

    function normalizeScoresToSheet(scoresObj) {
        // Return array of Q1..Q9 where each is either number or "N/A"
        var arr = [];
        for (var i = 1; i <= 9; i++) {
            var key = 'q' + i;
            var v = parseScore(scoresObj[key], key);
            arr.push(v === null ? 'N/A' : v);
        }
        return arr;
    }

    function createQAPdfSummaryChat(payload, attachFolder) {
        // payload: { title, meta:{...}, qScores:{q1..q9}, totals:{earned,possible,percentage,passStatus}, cats:{...}, notesHtml? }
        var doc = DocumentApp.create(payload.title || 'Chat QA');
        var body = doc.getBody();

        function spacer(h) {
            body.appendParagraph('').setSpacingAfter(h || 6);
        }

        // Title
        body.appendParagraph('TGC • Chat QA Summary')
            .setHeading(DocumentApp.ParagraphHeading.HEADING1);

        // Meta table
        var metaPairs = Object.entries(payload.meta || {});
        if (metaPairs.length) {
            var tbl = body.appendTable(metaPairs.map(function (p) {
                return [p[0] + ':', String(p[1] || '')];
            }));
            tbl.setBorderWidth(0);
            tbl.getRow(0).getCell(0).getChild(0).asText().setBold(true);
            metaPairs.forEach(function (_, i) {
                tbl.getCell(i, 0).getChild(0).asText().setBold(true);
            });
            spacer(10);
        }

        // Totals
        var t = payload.totals || {earned: 0, possible: 0, percentage: 0, passStatus: '—'};
        body.appendParagraph('Overall Results').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        body.appendParagraph('Score: ' + t.earned + ' / ' + t.possible);
        body.appendParagraph('Percentage: ' + t.percentage + '%');
        body.appendParagraph('Status: ' + t.passStatus);
        spacer(8);

        // Categories
        var cats = payload.cats || {};
        if (Object.keys(cats).length) {
            body.appendParagraph('Category Breakdown').setHeading(DocumentApp.ParagraphHeading.HEADING2);
            var ctbl = body.appendTable([['Category', '%']]);
            ctbl.getRow(0).getCell(0).getChild(0).asText().setBold(true);
            ctbl.getRow(0).getCell(1).getChild(0).asText().setBold(true);
            Object.keys(cats).forEach(function (name) {
                ctbl.appendTableRow().appendTableCell(name);
                ctbl.getRow(ctbl.getNumRows() - 1).appendTableCell(String(cats[name]) + '%');
            });
            spacer(8);
        }

        // Questions
        body.appendParagraph('Questions').setHeading(DocumentApp.ParagraphHeading.HEADING2);
        var qTbl = body.appendTable([['Question', 'Max', 'Score/NA']]);
        qTbl.getRow(0).getCell(0).getChild(0).asText().setBold(true);
        qTbl.getRow(0).getCell(1).getChild(0).asText().setBold(true);
        qTbl.getRow(0).getCell(2).getChild(0).asText().setBold(true);

        Object.keys(Q_MAX).forEach(function (k) {
            var r = qTbl.appendTableRow();
            r.appendTableCell(Q_TEXT[k]);
            r.appendTableCell(String(Q_MAX[k]));
            var val = payload.qScores && payload.qScores[k];
            r.appendTableCell(val == null ? 'N/A' : String(val));
        });
        spacer(10);

        // Notes
        if (payload.notesHtml) {
            body.appendParagraph('Internal Notes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
            // Strip HTML tags for DocApp (simple)
            var plain = payload.notesHtml.replace(/<[^>]+>/g, '').replace(/&nbsp;/g, ' ');
            body.appendParagraph(plain || '—');
        }

        doc.saveAndClose();

        var pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf')
            .setName((payload.title || 'Chat QA') + '.pdf');
        var file = DriveApp.createFile(pdfBlob)
            .setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        if (attachFolder) {
            attachFolder.addFile(file);
            DriveApp.getRootFolder().removeFile(file);
        }

        // Clean up temp doc
        DriveApp.getFileById(doc.getId()).setTrashed(true);
        return file.getUrl();
    }

    /* --------------------- PUBLIC OPS --------------------- */

    /**
     * Submit a CHAT QA record.
     * auditObj fields expected (from client):
     *  - customerName, agentName, agentEmail, conversationId, chatDate, auditorName
     *  - zeroTolerance: 'Yes' | 'No'
     *  - q1..q9: number (0..max) or 'NA'
     *  - notes: HTML string (internal notes)
     */
    function submitQA(auditObj) {
        try {
            var hc = TGC.Util.ensureQAInfra();
            if (!(hc.ok.spreadsheet && hc.ok.sheet)) {
                throw new Error('QA infra not healthy: ' + (hc.issues.join(' | ') || 'unknown issue'));
            }

            var zt = String(auditObj.zeroTolerance || 'No').toLowerCase() === 'yes';
            var scores = {};
            Object.keys(Q_MAX).forEach(function (k) {
                scores[k] = auditObj[k];
            });

            var totals = computeTotals(scores, zt);
            var cats = computeCategoryBreakdown(scores);

            var ss = TGC.Util.getSpreadsheet();
            var sh = TGC.Util.ensureSheetWithHeaders(ss, TGC.Util.getConfig().QA_SHEET_NAME, TGC.Util.getConfig().QA_HEADERS);

            var id = Utilities.getUuid();
            var ts = new Date().toISOString();
            var ymd = TGC.Util.safe(auditObj.chatDate || auditObj.callDate || TGC.Util.todayYMD());

            var rowObj = {
                'ID': id, 'Timestamp': ts,
                'CustomerName': TGC.Util.safe(auditObj.customerName),
                'AgentName': TGC.Util.safe(auditObj.agentName),
                'AgentEmail': TGC.Util.safe(auditObj.agentEmail),
                'Channel': CHANNEL,
                'ConversationId': TGC.Util.safe(auditObj.conversationId),
                'ChatDate': ymd,
                'AuditorName': TGC.Util.safe(auditObj.auditorName),
                'AuditDate': TGC.Util.todayYMD(),
                'ZeroTolerance': zt ? 'Yes' : 'No'
            };

            // Q1..Q9
            var qCells = normalizeScoresToSheet(scores);
            for (var i = 1; i <= 9; i++) {
                rowObj['Q' + i] = qCells[i - 1];
            }

            // totals + pass
            rowObj['TotalEarned'] = totals.earned;
            rowObj['TotalPossible'] = totals.possible;
            rowObj['Percentage'] = totals.percentage;
            rowObj['PassStatus'] = totals.passStatus;

            // categories (%)
            rowObj['Cat_Professionalism'] = cats.Professionalism || 0;
            rowObj['Cat_Comprehension'] = cats.Comprehension || 0;
            rowObj['Cat_Resolution'] = cats.Resolution || 0;
            rowObj['Cat_FollowUp'] = cats.FollowUp || 0;
            rowObj['Cat_Writing'] = cats.Writing || 0;
            rowObj['Cat_Process'] = cats.Process || 0;

            rowObj['NotesHtml'] = TGC.Util.safe(auditObj.notes);

            // write
            var headers = TGC.Util.getConfig().QA_HEADERS;
            var row = headers.map(function (h) {
                return (rowObj[h] != null ? rowObj[h] : '');
            });
            sh.appendRow(row);

            // OPTIONAL: create a PDF summary in root (or in a subfolder if you prefer)
            var pdfUrl;
            try {
                pdfUrl = createQAPdfSummaryChat({
                    title: rowObj.AgentName + ' - Chat QA - ' + rowObj.ChatDate,
                    meta: {
                        'Agent': rowObj.AgentName,
                        'Email': rowObj.AgentEmail,
                        'Conversation ID': rowObj.ConversationId,
                        'Channel': rowObj.Channel,
                        'Auditor': rowObj.AuditorName,
                        'Date': rowObj.ChatDate,
                        'Zero Tolerance': rowObj.ZeroTolerance
                    },
                    qScores: Object.keys(Q_MAX).reduce(function (o, k) {
                        o[k] = (scores[k] == null || /^na$/i.test(String(scores[k]))) ? null : parseScore(scores[k], k);
                        return o;
                    }, {}),
                    totals: totals,
                    cats: cats,
                    notesHtml: rowObj.NotesHtml
                });
            } catch (pdfErr) {
                TGC.Util.writeError('TGC.QA.createQAPdfSummaryChat', pdfErr);
            }

            return {id: id, percentage: totals.percentage, passStatus: totals.passStatus, pdfUrl: pdfUrl || ''};

        } catch (err) {
            TGC.Util.writeError('TGC.QA.submitQA', err);
            throw err;
        }
    }

    /* --------------------- CRUD --------------------- */
    function getAll() {
        var ss = TGC.Util.getSpreadsheet();
        var sh = TGC.Util.ensureSheetWithHeaders(ss, TGC.Util.getConfig().QA_SHEET_NAME, TGC.Util.getConfig().QA_HEADERS);
        var vals = sh.getDataRange().getValues();
        if (vals.length < 2) return [];
        var headers = vals.shift();
        return vals.map(function (r) {
            var o = {};
            r.forEach(function (c, i) {
                o[headers[i]] = c;
            });
            return o;
        });
    }

    function getById(id) {
        var ss = TGC.Util.getSpreadsheet();
        var sh = TGC.Util.ensureSheetWithHeaders(ss, TGC.Util.getConfig().QA_SHEET_NAME, TGC.Util.getConfig().QA_HEADERS);
        var vals = sh.getDataRange().getValues();
        var headers = vals.shift();
        var idCol = headers.indexOf('ID');
        for (var i = 0; i < vals.length; i++) {
            if (String(vals[i][idCol]) === String(id)) {
                var o = {};
                vals[i].forEach(function (c, j) {
                    o[headers[j]] = c;
                });
                return o;
            }
        }
        throw new Error('QA record not found: ' + id);
    }

    function updateById(id, data) {
        var ss = TGC.Util.getSpreadsheet();
        var sh = TGC.Util.ensureSheetWithHeaders(ss, TGC.Util.getConfig().QA_SHEET_NAME, TGC.Util.getConfig().QA_HEADERS);
        var vals = sh.getDataRange().getValues();
        var headers = vals[0];
        var idCol = headers.indexOf('ID');

        for (var i = 1; i < vals.length; i++) {
            if (String(vals[i][idCol]) === String(id)) {
                var rowOut = headers.map(function (h) {
                    return (h in data) ? data[h] : vals[i][headers.indexOf(h)];
                });
                sh.getRange(i + 1, 1, 1, rowOut.length).setValues([rowOut]);
                return true;
            }
        }
        throw new Error('QA record not found: ' + id);
    }

    function deleteById(id) {
        var ss = TGC.Util.getSpreadsheet();
        var sh = TGC.Util.ensureSheetWithHeaders(ss, TGC.Util.getConfig().QA_SHEET_NAME, TGC.Util.getConfig().QA_HEADERS);
        var vals = sh.getDataRange().getValues();
        var idCol = vals[0].indexOf('ID');
        for (var i = 1; i < vals.length; i++) {
            if (String(vals[i][idCol]) === String(id)) {
                sh.deleteRow(i + 1);
                return true;
            }
        }
        throw new Error('QA record not found: ' + id);
    }

    /* --------------------- DASHBOARD --------------------- */
    function getDashboardStats(opts) {
        opts = opts || {};
        var rows = getAll();

        // Optional filtering
        if (opts.agentName) {
            rows = rows.filter(function (r) {
                return String(r.AgentName) === String(opts.agentName);
            });
        }
        if (opts.fromDate || opts.toDate) {
            var from = opts.fromDate ? new Date(opts.fromDate) : null;
            var to = opts.toDate ? new Date(opts.toDate) : null;
            rows = rows.filter(function (r) {
                var d = new Date(r.ChatDate || r.AuditDate || r.Timestamp);
                if (from && d < from) return false;
                if (to && d > to) return false;
                return true;
            });
        }

        var total = rows.length;
        var avgPct = total ? Math.round(rows
            .map(function (r) {
                return +r.Percentage || 0;
            })
            .reduce(function (a, b) {
                return a + b;
            }, 0) / total) : 0;

        function avgCol(col) {
            return total ? Math.round(rows.map(function (r) {
                return +r[col] || 0;
            })
                .reduce(function (a, b) {
                    return a + b;
                }, 0) / total) : 0;
        }

        var category = {
            Professionalism: avgCol('Cat_Professionalism'),
            Comprehension: avgCol('Cat_Comprehension'),
            Resolution: avgCol('Cat_Resolution'),
            FollowUp: avgCol('Cat_FollowUp'),
            Writing: avgCol('Cat_Writing'),
            Process: avgCol('Cat_Process')
        };

        var autoFails = rows.filter(function (r) {
            return String(r.ZeroTolerance).toLowerCase() === 'yes';
        }).length;

        return {
            totalRecords: total,
            avgPercentage: avgPct,
            categoryAverages: category,
            autoFailCount: autoFails,
            passRate: total ? Math.round(rows.filter(function (r) {
                return String(r.PassStatus).toLowerCase() === 'pass';
            }).length * 100 / total) : 0
        };
    }

    /* --------------------- EXPORTS --------------------- */
    return {
        submitQA: submitQA,
        getAll: getAll,
        getById: getById,
        updateById: updateById,
        deleteById: deleteById,
        getDashboardStats: getDashboardStats,
        // expose for tests if needed
        _parseScore: parseScore,
        _computeTotals: computeTotals,
        _computeCategoryBreakdown: computeCategoryBreakdown
    };
})();

/* =======================================================================
   Thin wrappers for the UI (google.script.run.*)
   Configure once with TGC_bootstrap (or setConfig from a settings screen)
   ======================================================================= */
function TGC_bootstrap() {
    // Replace IDs/names here once and forget.
    TGC.Util.setConfig({
        ROOT_FOLDER_ID: 'YOUR_OPTIONAL_DRIVE_ROOT_ID',
        QA_SPREADSHEET_ID: 'YOUR_CHAT_QA_SPREADSHEET_ID',
        QA_SHEET_NAME: 'Chat_QA_Records',
        USERS_SHEET_NAME: 'Users'
    });
}

function TGC_QA_submitQA(obj) {
    TGC_bootstrap();
    return TGC.QA.submitQA(obj);
}

function TGC_QA_getAll() {
    TGC_bootstrap();
    return TGC.QA.getAll();
}

function TGC_QA_getById(id) {
    TGC_bootstrap();
    return TGC.QA.getById(id);
}

function TGC_QA_updateById(id, data) {
    TGC_bootstrap();
    return TGC.QA.updateById(id, data);
}

function TGC_QA_deleteById(id) {
    TGC_bootstrap();
    return TGC.QA.deleteById(id);
}

function TGC_QA_dashboard(opts) {
    TGC_bootstrap();
    return TGC.QA.getDashboardStats(opts);
}
