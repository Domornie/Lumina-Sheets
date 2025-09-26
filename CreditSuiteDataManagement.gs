/**
 * CreditSuiteQADataManagement.gs - Data Retrieval & Analytics for Credit Suite QA
 * Complete data management system for Credit Suite QA
 */

// ────────────────────────────────────────────────────────────────────────────
// Credit Suite QA Data Retrieval Functions
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get all Credit Suite QA records
 */
function getCreditSuiteQARecords() {
    try {
        const ss = SpreadsheetApp.openById(CREDIT_SUITE_SHEET_ID);
        const qaSheet = ss.getSheetByName(CREDIT_SUITE_QA_SHEET);
        
        if (!qaSheet) {
            console.log('Credit Suite QA sheet not found');
            return [];
        }
        
        const data = qaSheet.getDataRange().getValues();
        if (data.length <= 1) {
            return [];
        }
        
        const headers = data[0];
        const rows = data.slice(1);
        
        return rows.map(row => {
            const record = {};
            headers.forEach((header, index) => {
                record[header] = row[index];
            });
            return record;
        });
        
    } catch (error) {
        console.error('Error getting Credit Suite QA records:', error);
        safeWriteError('getCreditSuiteQARecords', error);
        return [];
    }
}

/**
 * Get Credit Suite QA record by ID
 */
function getCreditSuiteQAById(id) {
    try {
        const records = getCreditSuiteQARecords();
        return records.find(record => record.ID === id) || null;
        
    } catch (error) {
        console.error('Error getting Credit Suite QA by ID:', error);
        safeWriteError('getCreditSuiteQAById', error);
        return null;
    }
}

/**
 * Get Credit Suite QA analytics
 */
function getCreditSuiteQAAnalytics(granularity = 'Week', period = null, consultant = '', department = '') {
    try {
        const records = getCreditSuiteQARecords();
        
        if (!records.length) {
            return getEmptyCreditSuiteQAAnalytics();
        }
        
        // Filter records by period and consultant
        let filteredRecords = records;
        
        if (period) {
            filteredRecords = filterRecordsByPeriod(filteredRecords, granularity, period);
        }
        
        if (consultant) {
            filteredRecords = filteredRecords.filter(record => 
                (record.AgentName || '').toLowerCase().includes(consultant.toLowerCase())
            );
        }
        
        return calculateCreditSuiteAnalytics(filteredRecords, granularity);
        
    } catch (error) {
        console.error('Error getting Credit Suite QA analytics:', error);
        safeWriteError('getCreditSuiteQAAnalytics', error);
        return getEmptyCreditSuiteQAAnalytics();
    }
}

/**
 * Calculate analytics from filtered Credit Suite records
 */
function calculateCreditSuiteAnalytics(records, granularity) {
    if (!records.length) {
        return getEmptyCreditSuiteQAAnalytics();
    }
    
    // Basic metrics
    const totalEvaluations = records.length;
    const totalScore = records.reduce((sum, record) => sum + (parseFloat(record.PercentageScore) || 0), 0);
    const avgScore = totalEvaluations > 0 ? Math.round(totalScore / totalEvaluations) : 0;
    
    // Pass rate calculation
    const passedEvaluations = records.filter(record => {
        const passStatus = record.PassStatus || '';
        return passStatus.includes('Pass') || passStatus.includes('Excellent') || passStatus.includes('Meets Standards');
    }).length;
    const passRate = totalEvaluations > 0 ? Math.round((passedEvaluations / totalEvaluations) * 100) : 0;
    
    // Excellent rate calculation
    const excellentEvaluations = records.filter(record => {
        const passStatus = record.PassStatus || '';
        return passStatus.includes('Excellent');
    }).length;
    const excellentRate = totalEvaluations > 0 ? Math.round((excellentEvaluations / totalEvaluations) * 100) : 0;
    
    // Critical failures (compliance violations for Credit Suite)
    const criticalFailures = records.filter(record => {
        const passStatus = record.PassStatus || '';
        return passStatus.includes('Critical Failure') || passStatus.includes('compliance violation');
    }).length;
    
    // Compliance violations specifically
    const complianceViolations = records.filter(record => {
        const complianceNotes = record.ComplianceNotes || '';
        return complianceNotes.includes('CRITICAL') || complianceNotes.includes('Compliance Issue');
    }).length;
    
    // Consultant performance
    const consultantPerformance = calculateConsultantPerformance(records);
    
    // Category performance (Credit Suite specific)
    const categoryPerformance = calculateCreditSuiteCategoryPerformance(records);
    
    // Trends over time
    const trends = calculateTrends(records, granularity);
    
    // Consultation type breakdown
    const consultationTypeBreakdown = calculateConsultationTypeBreakdown(records);
    
    return {
        // Summary metrics
        avgScore,
        passRate,
        excellentRate,
        totalEvaluations,
        consultantsEvaluated: consultantPerformance.length,
        criticalFailures,
        complianceViolations,
        
        // Change indicators (would need historical data)
        avgScoreChange: 0,
        passRateChange: 0,
        evaluationsChange: 0,
        consultantsChange: 0,
        
        // Detailed breakdowns
        consultantPerformance,
        categoryPerformance,
        trends,
        consultationTypeBreakdown,
        
        // Chart data
        categories: {
            labels: categoryPerformance.map(cat => cat.name),
            values: categoryPerformance.map(cat => cat.avgScore)
        },
        
        consultants: {
            labels: consultantPerformance.slice(0, 10).map(consultant => consultant.name),
            values: consultantPerformance.slice(0, 10).map(consultant => consultant.avgScore)
        },
        
        trends: {
            labels: trends.map(trend => trend.period),
            values: trends.map(trend => trend.avgScore)
        }
    };
}

/**
 * Calculate consultant performance for Credit Suite
 */
function calculateConsultantPerformance(records) {
    const consultantGroups = {};
    
    records.forEach(record => {
        const consultantName = record.AgentName || 'Unknown';
        if (!consultantGroups[consultantName]) {
            consultantGroups[consultantName] = {
                name: consultantName,
                evaluations: [],
                totalScore: 0,
                passCount: 0,
                excellentCount: 0,
                criticalFailures: 0,
                complianceViolations: 0
            };
        }
        
        const score = parseFloat(record.PercentageScore) || 0;
        const passStatus = record.PassStatus || '';
        const complianceNotes = record.ComplianceNotes || '';
        
        consultantGroups[consultantName].evaluations.push(record);
        consultantGroups[consultantName].totalScore += score;
        
        if (passStatus.includes('Pass') || passStatus.includes('Excellent') || passStatus.includes('Meets Standards')) {
            consultantGroups[consultantName].passCount++;
        }
        
        if (passStatus.includes('Excellent')) {
            consultantGroups[consultantName].excellentCount++;
        }
        
        if (passStatus.includes('Critical Failure')) {
            consultantGroups[consultantName].criticalFailures++;
        }
        
        if (complianceNotes.includes('CRITICAL') || complianceNotes.includes('Compliance Issue')) {
            consultantGroups[consultantName].complianceViolations++;
        }
    });
    
    return Object.values(consultantGroups).map(consultant => {
        const evalCount = consultant.evaluations.length;
        return {
            name: consultant.name,
            evaluations: evalCount,
            avgScore: evalCount > 0 ? Math.round(consultant.totalScore / evalCount) : 0,
            passRate: evalCount > 0 ? Math.round((consultant.passCount / evalCount) * 100) : 0,
            excellentRate: evalCount > 0 ? Math.round((consultant.excellentCount / evalCount) * 100) : 0,
            criticalFailures: consultant.criticalFailures,
            complianceViolations: consultant.complianceViolations,
            latestScore: consultant.evaluations.length > 0 ? 
                parseFloat(consultant.evaluations[consultant.evaluations.length - 1].PercentageScore) || 0 : 0
        };
    }).sort((a, b) => b.avgScore - a.avgScore);
}

/**
 * Calculate category performance for Credit Suite
 */
function calculateCreditSuiteCategoryPerformance(records) {
    if (!CREDIT_SUITE_QA_CONFIG || !CREDIT_SUITE_QA_CONFIG.categories) {
        return [];
    }
    
    const categoryResults = {};
    
    // Initialize Credit Suite categories
    Object.keys(CREDIT_SUITE_QA_CONFIG.categories).forEach(categoryName => {
        categoryResults[categoryName] = {
            name: categoryName,
            totalQuestions: 0,
            totalAnswered: 0,
            yesCount: 0,
            noCount: 0,
            naCount: 0,
            complianceIssues: 0
        };
    });
    
    // Process each record
    records.forEach(record => {
        Object.entries(CREDIT_SUITE_QA_CONFIG.categories).forEach(([categoryName, categoryData]) => {
            categoryData.questions?.forEach(question => {
                const answer = record[question.id];
                if (answer) {
                    categoryResults[categoryName].totalQuestions++;
                    categoryResults[categoryName].totalAnswered++;
                    
                    if (answer === 'Yes') {
                        categoryResults[categoryName].yesCount++;
                    } else if (answer === 'No') {
                        categoryResults[categoryName].noCount++;
                        // Track compliance issues for critical questions
                        if (question.critical && categoryName === 'Compliance & Legal Requirements') {
                            categoryResults[categoryName].complianceIssues++;
                        }
                    } else if (answer === 'NA') {
                        categoryResults[categoryName].naCount++;
                    }
                }
            });
        });
    });
    
    // Calculate percentages
    return Object.values(categoryResults).map(category => {
        const answeredQuestions = category.yesCount + category.noCount; // Exclude N/A
        const avgScore = answeredQuestions > 0 ? 
            Math.round((category.yesCount / answeredQuestions) * 100) : 0;
        
        return {
            name: category.name,
            avgScore,
            yesCount: category.yesCount,
            noCount: category.noCount,
            naCount: category.naCount,
            complianceIssues: category.complianceIssues,
            totalAnswered: category.totalAnswered
        };
    });
}

/**
 * Calculate consultation type breakdown for Credit Suite
 */
function calculateConsultationTypeBreakdown(records) {
    const consultationTypeGroups = {};
    
    records.forEach(record => {
        const consultationType = record.ConsultationType || 'Unknown';
        if (!consultationTypeGroups[consultationType]) {
            consultationTypeGroups[consultationType] = {
                type: consultationType,
                count: 0,
                totalScore: 0,
                complianceIssues: 0
            };
        }
        
        consultationTypeGroups[consultationType].count++;
        consultationTypeGroups[consultationType].totalScore += parseFloat(record.PercentageScore) || 0;
        
        // Count compliance issues
        const complianceNotes = record.ComplianceNotes || '';
        if (complianceNotes.includes('CRITICAL') || complianceNotes.includes('Compliance Issue')) {
            consultationTypeGroups[consultationType].complianceIssues++;
        }
    });
    
    return Object.values(consultationTypeGroups).map(group => ({
        type: group.type,
        count: group.count,
        avgScore: group.count > 0 ? Math.round(group.totalScore / group.count) : 0,
        percentage: records.length > 0 ? Math.round((group.count / records.length) * 100) : 0,
        complianceIssues: group.complianceIssues
    }));
}

/**
 * Filter records by period (shared utility)
 */
function filterRecordsByPeriod(records, granularity, period) {
    return records.filter(record => {
        const date = record.Timestamp || record.ConsultationDate;
        if (!date) return false;
        
        const recordDate = new Date(date);
        let recordPeriod;
        
        switch (granularity) {
            case 'Week':
                recordPeriod = getISOWeek(recordDate);
                break;
            case 'Month':
                recordPeriod = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM');
                break;
            case 'Quarter':
                const quarter = Math.ceil((recordDate.getMonth() + 1) / 3);
                recordPeriod = `Q${quarter}-${recordDate.getFullYear()}`;
                break;
            case 'Year':
                recordPeriod = recordDate.getFullYear().toString();
                break;
            case 'Day':
            default:
                recordPeriod = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                break;
        }
        
        return recordPeriod === period;
    });
}

/**
 * Calculate trends over time (shared utility)
 */
function calculateTrends(records, granularity) {
    const trendGroups = {};
    
    records.forEach(record => {
        const date = record.Timestamp || record.ConsultationDate;
        if (!date) return;
        
        let period;
        const recordDate = new Date(date);
        
        switch (granularity) {
            case 'Week':
                period = getISOWeek(recordDate);
                break;
            case 'Month':
                period = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM');
                break;
            case 'Quarter':
                const quarter = Math.ceil((recordDate.getMonth() + 1) / 3);
                period = `Q${quarter}-${recordDate.getFullYear()}`;
                break;
            case 'Year':
                period = recordDate.getFullYear().toString();
                break;
            case 'Day':
            default:
                period = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                break;
        }
        
        if (!trendGroups[period]) {
            trendGroups[period] = {
                period,
                evaluations: [],
                totalScore: 0
            };
        }
        
        const score = parseFloat(record.PercentageScore) || 0;
        trendGroups[period].evaluations.push(record);
        trendGroups[period].totalScore += score;
    });
    
    return Object.values(trendGroups)
        .map(group => ({
            period: group.period,
            evaluations: group.evaluations.length,
            avgScore: group.evaluations.length > 0 ? 
                Math.round(group.totalScore / group.evaluations.length) : 0
        }))
        .sort((a, b) => a.period.localeCompare(b.period));
}

/**
 * Get week string from date (ISO week format)
 */
function getISOWeek(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return `${d.getUTCFullYear()}-W${Math.ceil((((d - yearStart) / 86400000) + 1) / 7).toString().padStart(2, '0')}`;
}

/**
 * Get empty analytics structure for Credit Suite
 */
function getEmptyCreditSuiteQAAnalytics() {
    return {
        avgScore: 0,
        passRate: 0,
        excellentRate: 0,
        totalEvaluations: 0,
        consultantsEvaluated: 0,
        criticalFailures: 0,
        complianceViolations: 0,
        avgScoreChange: 0,
        passRateChange: 0,
        evaluationsChange: 0,
        consultantsChange: 0,
        consultantPerformance: [],
        categoryPerformance: [],
        trends: [],
        consultationTypeBreakdown: [],
        categories: { labels: [], values: [] },
        consultants: { labels: [], values: [] },
        trends: { labels: [], values: [] }
    };
}

// ────────────────────────────────────────────────────────────────────────────
// Client-accessible Functions for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Client function to get all Credit Suite QA records
 */
function clientGetCreditSuiteQARecords() {
    try {
        return getCreditSuiteQARecords();
    } catch (error) {
        console.error('Error in clientGetCreditSuiteQARecords:', error);
        safeWriteError('clientGetCreditSuiteQARecords', error);
        return [];
    }
}

/**
 * Client function to get Credit Suite QA by ID
 */
function clientGetCreditSuiteQAById(id) {
    try {
        return getCreditSuiteQAById(id);
    } catch (error) {
        console.error('Error in clientGetCreditSuiteQAById:', error);
        safeWriteError('clientGetCreditSuiteQAById', error);
        return null;
    }
}

/**
 * Client function to get Credit Suite QA analytics
 */
function clientGetCreditSuiteQAAnalytics(granularity = 'Week', period = null, consultant = '', department = '') {
    try {
        return getCreditSuiteQAAnalytics(granularity, period, consultant, department);
    } catch (error) {
        console.error('Error in clientGetCreditSuiteQAAnalytics:', error);
        safeWriteError('clientGetCreditSuiteQAAnalytics', error);
        return getEmptyCreditSuiteQAAnalytics();
    }
}

/**
 * Delete Credit Suite QA record
 */
function deleteCreditSuiteQARecord(id) {
    try {
        const ss = SpreadsheetApp.openById(CREDIT_SUITE_SHEET_ID);
        const qaSheet = ss.getSheetByName(CREDIT_SUITE_QA_SHEET);
        
        if (!qaSheet) {
            throw new Error('Credit Suite QA sheet not found');
        }
        
        const data = qaSheet.getDataRange().getValues();
        const headers = data[0];
        const idColumn = headers.indexOf('ID');
        
        if (idColumn === -1) {
            throw new Error('ID column not found');
        }
        
        // Find the row to delete
        for (let i = 1; i < data.length; i++) {
            if (data[i][idColumn] === id) {
                qaSheet.deleteRow(i + 1);
                console.log('Credit Suite QA record deleted:', id);
                return { success: true, message: 'Record deleted successfully' };
            }
        }
        
        throw new Error('Record not found');
        
    } catch (error) {
        console.error('Error deleting Credit Suite QA record:', error);
        safeWriteError('deleteCreditSuiteQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Client function to delete Credit Suite QA record
 */
function clientDeleteCreditSuiteQARecord(id) {
    try {
        return deleteCreditSuiteQARecord(id);
    } catch (error) {
        console.error('Error in clientDeleteCreditSuiteQARecord:', error);
        safeWriteError('clientDeleteCreditSuiteQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Update Credit Suite QA record
 */
function updateCreditSuiteQARecord(id, updateData) {
    try {
        const ss = SpreadsheetApp.openById(CREDIT_SUITE_SHEET_ID);
        const qaSheet = ss.getSheetByName(CREDIT_SUITE_QA_SHEET);
        
        if (!qaSheet) {
            throw new Error('Credit Suite QA sheet not found');
        }
        
        const data = qaSheet.getDataRange().getValues();
        const headers = data[0];
        const idColumn = headers.indexOf('ID');
        
        if (idColumn === -1) {
            throw new Error('ID column not found');
        }
        
        // Find the row to update
        for (let i = 1; i < data.length; i++) {
            if (data[i][idColumn] === id) {
                // Update the row with new data
                Object.keys(updateData).forEach(key => {
                    const columnIndex = headers.indexOf(key);
                    if (columnIndex !== -1) {
                        qaSheet.getRange(i + 1, columnIndex + 1).setValue(updateData[key]);
                    }
                });
                
                // Update the UpdatedAt timestamp
                const updatedAtColumn = headers.indexOf('UpdatedAt');
                if (updatedAtColumn !== -1) {
                    qaSheet.getRange(i + 1, updatedAtColumn + 1).setValue(new Date());
                }
                
                console.log('Credit Suite QA record updated:', id);
                return { success: true, message: 'Record updated successfully' };
            }
        }
        
        throw new Error('Record not found');
        
    } catch (error) {
        console.error('Error updating Credit Suite QA record:', error);
        safeWriteError('updateCreditSuiteQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Client function to update Credit Suite QA record
 */
function clientUpdateCreditSuiteQARecord(id, updateData) {
    try {
        return updateCreditSuiteQARecord(id, updateData);
    } catch (error) {
        console.error('Error in clientUpdateCreditSuiteQARecord:', error);
        safeWriteError('clientUpdateCreditSuiteQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Export Credit Suite QA data to CSV
 */
function exportCreditSuiteQAToCSV(granularity = 'Week', period = null, consultant = '') {
    try {
        let records = getCreditSuiteQARecords();
        
        // Filter records if needed
        if (period) {
            records = filterRecordsByPeriod(records, granularity, period);
        }
        
        if (consultant) {
            records = records.filter(record => 
                (record.AgentName || '').toLowerCase().includes(consultant.toLowerCase())
            );
        }
        
        if (!records.length) {
            return 'No Credit Suite data available for the selected criteria.';
        }
        
        // Generate CSV
        const headers = CREDIT_SUITE_QA_HEADERS;
        const csvRows = [headers.join(',')];
        
        records.forEach(record => {
            const row = headers.map(header => {
                const value = record[header] || '';
                // Escape CSV values
                if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                    return '"' + value.replace(/"/g, '""') + '"';
                }
                return value;
            });
            csvRows.push(row.join(','));
        });
        
        return csvRows.join('\n');
        
    } catch (error) {
        console.error('Error exporting Credit Suite QA to CSV:', error);
        safeWriteError('exportCreditSuiteQAToCSV', error);
        return 'Error generating Credit Suite CSV export.';
    }
}

/**
 * Client function to export Credit Suite QA to CSV
 */
function clientExportCreditSuiteQAToCSV(granularity = 'Week', period = null, consultant = '') {
    try {
        return exportCreditSuiteQAToCSV(granularity, period, consultant);
    } catch (error) {
        console.error('Error in clientExportCreditSuiteQAToCSV:', error);
        safeWriteError('clientExportCreditSuiteQAToCSV', error);
        return 'Error generating Credit Suite CSV export.';
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Credit Suite Specific Analytics Functions
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get compliance analytics specifically for Credit Suite
 */
function getCreditSuiteComplianceAnalytics(granularity = 'Week', period = null) {
    try {
        const records = getCreditSuiteQARecords();
        
        if (!records.length) {
            return {
                totalAssessments: 0,
                complianceViolations: 0,
                complianceRate: 100,
                criticalIssues: 0,
                complianceByCategory: [],
                complianceByConsultant: [],
                complianceTrends: []
            };
        }
        
        // Filter records if needed
        let filteredRecords = records;
        if (period) {
            filteredRecords = filterRecordsByPeriod(filteredRecords, granularity, period);
        }
        
        const totalAssessments = filteredRecords.length;
        
        // Count compliance violations
        const complianceViolations = filteredRecords.filter(record => {
            const complianceNotes = record.ComplianceNotes || '';
            const passStatus = record.PassStatus || '';
            return complianceNotes.includes('CRITICAL') || 
                   complianceNotes.includes('Compliance Issue') ||
                   passStatus.includes('compliance violation');
        }).length;
        
        const complianceRate = totalAssessments > 0 ? 
            Math.round(((totalAssessments - complianceViolations) / totalAssessments) * 100) : 100;
        
        // Count critical issues specifically
        const criticalIssues = filteredRecords.filter(record => {
            const passStatus = record.PassStatus || '';
            return passStatus.includes('Critical Failure');
        }).length;
        
        // Compliance by category
        const complianceByCategory = calculateComplianceByCategory(filteredRecords);
        
        // Compliance by consultant
        const complianceByConsultant = calculateComplianceByConsultant(filteredRecords);
        
        // Compliance trends
        const complianceTrends = calculateComplianceTrends(filteredRecords, granularity);
        
        return {
            totalAssessments,
            complianceViolations,
            complianceRate,
            criticalIssues,
            complianceByCategory,
            complianceByConsultant,
            complianceTrends
        };
        
    } catch (error) {
        console.error('Error getting Credit Suite compliance analytics:', error);
        safeWriteError('getCreditSuiteComplianceAnalytics', error);
        return {
            totalAssessments: 0,
            complianceViolations: 0,
            complianceRate: 100,
            criticalIssues: 0,
            complianceByCategory: [],
            complianceByConsultant: [],
            complianceTrends: []
        };
    }
}

/**
 * Calculate compliance by category
 */
function calculateComplianceByCategory(records) {
    const complianceCategories = [
        'Q11_FCRACompliance',
        'Q12_CRLComplianceEducation', 
        'Q13_TruthfulRepresentation',
        'Q14_ProperDisclosures'
    ];
    
    return complianceCategories.map(questionId => {
        const violations = records.filter(record => record[questionId] === 'No').length;
        const evaluated = records.filter(record => record[questionId] && record[questionId] !== 'NA').length;
        const complianceRate = evaluated > 0 ? Math.round(((evaluated - violations) / evaluated) * 100) : 100;
        
        return {
            category: questionId.replace('Q', '').replace('_', ' '),
            violations,
            evaluated,
            complianceRate
        };
    });
}

/**
 * Calculate compliance by consultant
 */
function calculateComplianceByConsultant(records) {
    const consultantGroups = {};
    
    records.forEach(record => {
        const consultantName = record.AgentName || 'Unknown';
        if (!consultantGroups[consultantName]) {
            consultantGroups[consultantName] = {
                name: consultantName,
                totalAssessments: 0,
                complianceViolations: 0
            };
        }
        
        consultantGroups[consultantName].totalAssessments++;
        
        const complianceNotes = record.ComplianceNotes || '';
        const passStatus = record.PassStatus || '';
        
        if (complianceNotes.includes('CRITICAL') || 
            complianceNotes.includes('Compliance Issue') ||
            passStatus.includes('compliance violation')) {
            consultantGroups[consultantName].complianceViolations++;
        }
    });
    
    return Object.values(consultantGroups).map(consultant => ({
        name: consultant.name,
        totalAssessments: consultant.totalAssessments,
        complianceViolations: consultant.complianceViolations,
        complianceRate: consultant.totalAssessments > 0 ? 
            Math.round(((consultant.totalAssessments - consultant.complianceViolations) / consultant.totalAssessments) * 100) : 100
    })).sort((a, b) => b.complianceRate - a.complianceRate);
}

/**
 * Calculate compliance trends over time
 */
function calculateComplianceTrends(records, granularity) {
    const trendGroups = {};
    
    records.forEach(record => {
        const date = record.Timestamp || record.ConsultationDate;
        if (!date) return;
        
        let period;
        const recordDate = new Date(date);
        
        switch (granularity) {
            case 'Week':
                period = getISOWeek(recordDate);
                break;
            case 'Month':
                period = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM');
                break;
            case 'Quarter':
                const quarter = Math.ceil((recordDate.getMonth() + 1) / 3);
                period = `Q${quarter}-${recordDate.getFullYear()}`;
                break;
            case 'Year':
                period = recordDate.getFullYear().toString();
                break;
            case 'Day':
            default:
                period = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
                break;
        }
        
        if (!trendGroups[period]) {
            trendGroups[period] = {
                period,
                totalAssessments: 0,
                complianceViolations: 0
            };
        }
        
        trendGroups[period].totalAssessments++;
        
        const complianceNotes = record.ComplianceNotes || '';
        const passStatus = record.PassStatus || '';
        
        if (complianceNotes.includes('CRITICAL') || 
            complianceNotes.includes('Compliance Issue') ||
            passStatus.includes('compliance violation')) {
            trendGroups[period].complianceViolations++;
        }
    });
    
    return Object.values(trendGroups)
        .map(group => ({
            period: group.period,
            totalAssessments: group.totalAssessments,
            complianceViolations: group.complianceViolations,
            complianceRate: group.totalAssessments > 0 ? 
                Math.round(((group.totalAssessments - group.complianceViolations) / group.totalAssessments) * 100) : 100
        }))
        .sort((a, b) => a.period.localeCompare(b.period));
}

/**
 * Client function to get Credit Suite compliance analytics
 */
function clientGetCreditSuiteComplianceAnalytics(granularity = 'Week', period = null) {
    try {
        return getCreditSuiteComplianceAnalytics(granularity, period);
    } catch (error) {
        console.error('Error in clientGetCreditSuiteComplianceAnalytics:', error);
        safeWriteError('clientGetCreditSuiteComplianceAnalytics', error);
        return {
            totalAssessments: 0,
            complianceViolations: 0,
            complianceRate: 100,
            criticalIssues: 0,
            complianceByCategory: [],
            complianceByConsultant: [],
            complianceTrends: []
        };
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Utility Functions for Date/Time Handling (shared)
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get current week string
 */
function getCurrentWeekString() {
    return getISOWeek(new Date());
}

/**
 * Generate period options for dropdowns
 */
function getCreditSuitePeriodOptions(granularity = 'Week') {
    const options = [];
    const now = new Date();
    
    switch (granularity) {
        case 'Week':
            // Generate last 12 weeks
            for (let i = 0; i < 12; i++) {
                const date = new Date(now.getTime() - (i * 7 * 24 * 60 * 60 * 1000));
                options.push({
                    value: getISOWeek(date),
                    label: `Week of ${Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy')}`
                });
            }
            break;
            
        case 'Month':
            // Generate last 12 months
            for (let i = 0; i < 12; i++) {
                const date = new Date(now.getFullYear(), now.getMonth() - i, 1);
                options.push({
                    value: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM'),
                    label: Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM yyyy')
                });
            }
            break;
            
        case 'Quarter':
            // Generate last 8 quarters
            for (let i = 0; i < 8; i++) {
                const quarterDate = new Date(now.getFullYear(), now.getMonth() - (i * 3), 1);
                const quarter = Math.ceil((quarterDate.getMonth() + 1) / 3);
                const year = quarterDate.getFullYear();
                options.push({
                    value: `Q${quarter}-${year}`,
                    label: `Q${quarter} ${year}`
                });
            }
            break;
            
        case 'Year':
            // Generate last 5 years
            for (let i = 0; i < 5; i++) {
                const year = now.getFullYear() - i;
                options.push({
                    value: year.toString(),
                    label: year.toString()
                });
            }
            break;
            
        case 'Day':
            // Generate last 30 days
            for (let i = 0; i < 30; i++) {
                const date = new Date(now.getTime() - (i * 24 * 60 * 60 * 1000));
                options.push({
                    value: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
                    label: Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMM dd, yyyy')
                });
            }
            break;
    }
    
    return options;
}

/**
 * Client function to get Credit Suite period options
 */
function clientGetCreditSuitePeriodOptions(granularity = 'Week') {
    try {
        return getCreditSuitePeriodOptions(granularity);
    } catch (error) {
        console.error('Error in clientGetCreditSuitePeriodOptions:', error);
        safeWriteError('clientGetCreditSuitePeriodOptions', error);
        return [];
    }
}

console.log('Credit Suite QA Data Management functions loaded successfully');
console.log('Available functions: getCreditSuiteQARecords, getCreditSuiteQAAnalytics, getCreditSuiteComplianceAnalytics');

