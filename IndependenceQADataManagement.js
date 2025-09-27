/**
 * IndependenceQADataManagement.gs - Data Retrieval & Analytics for Independence Insurance QA
 * Complete data management system for Independence QA
 */

// ────────────────────────────────────────────────────────────────────────────
// Independence Insurance QA Data Retrieval Functions
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get all Independence QA records
 */
function getIndependenceQARecords() {
    try {
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
        const qaSheet = ss.getSheetByName(INDEPENDENCE_QA_SHEET);
        
        if (!qaSheet) {
            console.log('Independence QA sheet not found');
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
        console.error('Error getting Independence QA records:', error);
        safeWriteError('getIndependenceQARecords', error);
        return [];
    }
}

/**
 * Get Independence QA record by ID
 */
function getIndependenceQAById(id) {
    try {
        const records = getIndependenceQARecords();
        return records.find(record => record.ID === id) || null;
        
    } catch (error) {
        console.error('Error getting Independence QA by ID:', error);
        safeWriteError('getIndependenceQAById', error);
        return null;
    }
}

/**
 * Get Independence QA analytics
 */
function getIndependenceQAAnalytics(granularity = 'Week', period = null, agent = '', department = '') {
    try {
        const records = getIndependenceQARecords();
        
        if (!records.length) {
            return getEmptyIndependenceQAAnalytics();
        }
        
        // Filter records by period and agent
        let filteredRecords = records;
        
        if (period) {
            filteredRecords = filterRecordsByPeriod(filteredRecords, granularity, period);
        }
        
        if (agent) {
            filteredRecords = filteredRecords.filter(record => 
                (record.AgentName || '').toLowerCase().includes(agent.toLowerCase())
            );
        }
        
        return calculateIndependenceAnalytics(filteredRecords, granularity);
        
    } catch (error) {
        console.error('Error getting Independence QA analytics:', error);
        safeWriteError('getIndependenceQAAnalytics', error);
        return getEmptyIndependenceQAAnalytics();
    }
}

/**
 * Calculate analytics from filtered records
 */
function calculateIndependenceAnalytics(records, granularity) {
    if (!records.length) {
        return getEmptyIndependenceQAAnalytics();
    }
    
    // Basic metrics
    const totalEvaluations = records.length;
    const totalScore = records.reduce((sum, record) => sum + (parseFloat(record.PercentageScore) || 0), 0);
    const avgScore = totalEvaluations > 0 ? Math.round(totalScore / totalEvaluations) : 0;
    
    // Pass rate calculation
    const passedEvaluations = records.filter(record => {
        const passStatus = record.PassStatus || '';
        return passStatus.includes('Pass') || passStatus.includes('Excellent');
    }).length;
    const passRate = totalEvaluations > 0 ? Math.round((passedEvaluations / totalEvaluations) * 100) : 0;
    
    // Excellent rate calculation
    const excellentEvaluations = records.filter(record => {
        const passStatus = record.PassStatus || '';
        return passStatus.includes('Excellent');
    }).length;
    const excellentRate = totalEvaluations > 0 ? Math.round((excellentEvaluations / totalEvaluations) * 100) : 0;
    
    // Critical failures
    const criticalFailures = records.filter(record => {
        const passStatus = record.PassStatus || '';
        return passStatus.includes('Critical Failure');
    }).length;
    
    // Agent performance
    const agentPerformance = calculateAgentPerformance(records);
    
    // Category performance
    const categoryPerformance = calculateCategoryPerformance(records);
    
    // Trends over time
    const trends = calculateTrends(records, granularity);
    
    // Call type breakdown
    const callTypeBreakdown = calculateCallTypeBreakdown(records);
    
    return {
        // Summary metrics
        avgScore,
        passRate,
        excellentRate,
        totalEvaluations,
        agentsEvaluated: agentPerformance.length,
        criticalFailures,
        
        // Change indicators (would need historical data)
        avgScoreChange: 0,
        passRateChange: 0,
        evaluationsChange: 0,
        agentsChange: 0,
        
        // Detailed breakdowns
        agentPerformance,
        categoryPerformance,
        trends,
        callTypeBreakdown,
        
        // Chart data
        categories: {
            labels: categoryPerformance.map(cat => cat.name),
            values: categoryPerformance.map(cat => cat.avgScore)
        },
        
        agents: {
            labels: agentPerformance.slice(0, 10).map(agent => agent.name),
            values: agentPerformance.slice(0, 10).map(agent => agent.avgScore)
        },
        
        trends: {
            labels: trends.map(trend => trend.period),
            values: trends.map(trend => trend.avgScore)
        }
    };
}

/**
 * Calculate agent performance
 */
function calculateAgentPerformance(records) {
    const agentGroups = {};
    
    records.forEach(record => {
        const agentName = record.AgentName || 'Unknown';
        if (!agentGroups[agentName]) {
            agentGroups[agentName] = {
                name: agentName,
                evaluations: [],
                totalScore: 0,
                passCount: 0,
                excellentCount: 0,
                criticalFailures: 0
            };
        }
        
        const score = parseFloat(record.PercentageScore) || 0;
        const passStatus = record.PassStatus || '';
        
        agentGroups[agentName].evaluations.push(record);
        agentGroups[agentName].totalScore += score;
        
        if (passStatus.includes('Pass') || passStatus.includes('Excellent')) {
            agentGroups[agentName].passCount++;
        }
        
        if (passStatus.includes('Excellent')) {
            agentGroups[agentName].excellentCount++;
        }
        
        if (passStatus.includes('Critical Failure')) {
            agentGroups[agentName].criticalFailures++;
        }
    });
    
    return Object.values(agentGroups).map(agent => {
        const evalCount = agent.evaluations.length;
        return {
            name: agent.name,
            evaluations: evalCount,
            avgScore: evalCount > 0 ? Math.round(agent.totalScore / evalCount) : 0,
            passRate: evalCount > 0 ? Math.round((agent.passCount / evalCount) * 100) : 0,
            excellentRate: evalCount > 0 ? Math.round((agent.excellentCount / evalCount) * 100) : 0,
            criticalFailures: agent.criticalFailures,
            latestScore: agent.evaluations.length > 0 ? 
                parseFloat(agent.evaluations[agent.evaluations.length - 1].PercentageScore) || 0 : 0
        };
    }).sort((a, b) => b.avgScore - a.avgScore);
}

/**
 * Calculate category performance
 */
function calculateCategoryPerformance(records) {
    if (!INDEPENDENCE_QA_CONFIG || !INDEPENDENCE_QA_CONFIG.categories) {
        return [];
    }
    
    const categoryResults = {};
    
    // Initialize categories
    Object.keys(INDEPENDENCE_QA_CONFIG.categories).forEach(categoryName => {
        categoryResults[categoryName] = {
            name: categoryName,
            totalQuestions: 0,
            totalAnswered: 0,
            yesCount: 0,
            noCount: 0,
            naCount: 0
        };
    });
    
    // Process each record
    records.forEach(record => {
        Object.entries(INDEPENDENCE_QA_CONFIG.categories).forEach(([categoryName, categoryData]) => {
            categoryData.questions?.forEach(question => {
                const answer = record[question.id];
                if (answer) {
                    categoryResults[categoryName].totalQuestions++;
                    categoryResults[categoryName].totalAnswered++;
                    
                    if (answer === 'Yes') categoryResults[categoryName].yesCount++;
                    else if (answer === 'No') categoryResults[categoryName].noCount++;
                    else if (answer === 'NA') categoryResults[categoryName].naCount++;
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
            totalAnswered: category.totalAnswered
        };
    });
}

/**
 * Calculate trends over time
 */
function calculateTrends(records, granularity) {
    const trendGroups = {};
    
    records.forEach(record => {
        const date = record.Timestamp || record.CallDate;
        if (!date) return;
        
        let period;
        const recordDate = new Date(date);
        
        switch (granularity) {
            case 'Week':
                period = getWeekString(recordDate);
                break;
            case 'Month':
                period = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM');
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
 * Calculate call type breakdown
 */
function calculateCallTypeBreakdown(records) {
    const callTypeGroups = {};
    
    records.forEach(record => {
        const callType = record.CallType || 'Unknown';
        if (!callTypeGroups[callType]) {
            callTypeGroups[callType] = {
                type: callType,
                count: 0,
                totalScore: 0
            };
        }
        
        callTypeGroups[callType].count++;
        callTypeGroups[callType].totalScore += parseFloat(record.PercentageScore) || 0;
    });
    
    return Object.values(callTypeGroups).map(group => ({
        type: group.type,
        count: group.count,
        avgScore: group.count > 0 ? Math.round(group.totalScore / group.count) : 0,
        percentage: records.length > 0 ? Math.round((group.count / records.length) * 100) : 0
    }));
}

/**
 * Filter records by period
 */
function filterRecordsByPeriod(records, granularity, period) {
    return records.filter(record => {
        const date = record.Timestamp || record.CallDate;
        if (!date) return false;
        
        const recordDate = new Date(date);
        let recordPeriod;
        
        switch (granularity) {
            case 'Week':
                recordPeriod = getWeekString(recordDate);
                break;
            case 'Month':
                recordPeriod = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM');
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
 * Get week string from date (ISO week format)
 */
function getWeekString(date) {
    const year = date.getFullYear();
    const start = new Date(year, 0, 1);
    const diff = date - start;
    const oneWeek = 1000 * 60 * 60 * 24 * 7;
    const week = Math.ceil(diff / oneWeek);
    return `${year}-W${week.toString().padStart(2, '0')}`;
}

/**
 * Get empty analytics structure
 */
function getEmptyIndependenceQAAnalytics() {
    return {
        avgScore: 0,
        passRate: 0,
        excellentRate: 0,
        totalEvaluations: 0,
        agentsEvaluated: 0,
        criticalFailures: 0,
        avgScoreChange: 0,
        passRateChange: 0,
        evaluationsChange: 0,
        agentsChange: 0,
        agentPerformance: [],
        categoryPerformance: [],
        trends: [],
        callTypeBreakdown: [],
        categories: { labels: [], values: [] },
        agents: { labels: [], values: [] },
        trends: { labels: [], values: [] }
    };
}

// ────────────────────────────────────────────────────────────────────────────
// Client-accessible Functions
// ────────────────────────────────────────────────────────────────────────────

/**
 * Client function to get all Independence QA records
 */
function clientGetIndependenceQARecords() {
    try {
        return getIndependenceQARecords();
    } catch (error) {
        console.error('Error in clientGetIndependenceQARecords:', error);
        safeWriteError('clientGetIndependenceQARecords', error);
        return [];
    }
}

/**
 * Client function to get Independence QA by ID
 */
function clientGetIndependenceQAById(id) {
    try {
        return getIndependenceQAById(id);
    } catch (error) {
        console.error('Error in clientGetIndependenceQAById:', error);
        safeWriteError('clientGetIndependenceQAById', error);
        return null;
    }
}

/**
 * Client function to get Independence QA analytics
 */
function clientGetIndependenceQAAnalytics(granularity = 'Week', period = null, agent = '', department = '') {
    try {
        return getIndependenceQAAnalytics(granularity, period, agent, department);
    } catch (error) {
        console.error('Error in clientGetIndependenceQAAnalytics:', error);
        safeWriteError('clientGetIndependenceQAAnalytics', error);
        return getEmptyIndependenceQAAnalytics();
    }
}

/**
 * Delete Independence QA record
 */
function deleteIndependenceQARecord(id) {
    try {
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
        const qaSheet = ss.getSheetByName(INDEPENDENCE_QA_SHEET);
        
        if (!qaSheet) {
            throw new Error('Independence QA sheet not found');
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
                console.log('Independence QA record deleted:', id);
                return { success: true, message: 'Record deleted successfully' };
            }
        }
        
        throw new Error('Record not found');
        
    } catch (error) {
        console.error('Error deleting Independence QA record:', error);
        safeWriteError('deleteIndependenceQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Client function to delete Independence QA record
 */
function clientDeleteIndependenceQARecord(id) {
    try {
        return deleteIndependenceQARecord(id);
    } catch (error) {
        console.error('Error in clientDeleteIndependenceQARecord:', error);
        safeWriteError('clientDeleteIndependenceQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Update Independence QA record
 */
function updateIndependenceQARecord(id, updateData) {
    try {
        const ss = SpreadsheetApp.openById(INDEPENDENCE_SHEET_ID);
        const qaSheet = ss.getSheetByName(INDEPENDENCE_QA_SHEET);
        
        if (!qaSheet) {
            throw new Error('Independence QA sheet not found');
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
                
                console.log('Independence QA record updated:', id);
                return { success: true, message: 'Record updated successfully' };
            }
        }
        
        throw new Error('Record not found');
        
    } catch (error) {
        console.error('Error updating Independence QA record:', error);
        safeWriteError('updateIndependenceQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Client function to update Independence QA record
 */
function clientUpdateIndependenceQARecord(id, updateData) {
    try {
        return updateIndependenceQARecord(id, updateData);
    } catch (error) {
        console.error('Error in clientUpdateIndependenceQARecord:', error);
        safeWriteError('clientUpdateIndependenceQARecord', error);
        return { success: false, error: error.message };
    }
}

/**
 * Export Independence QA data to CSV
 */
function exportIndependenceQAToCSV(granularity = 'Week', period = null, agent = '') {
    try {
        let records = getIndependenceQARecords();
        
        // Filter records if needed
        if (period) {
            records = filterRecordsByPeriod(records, granularity, period);
        }
        
        if (agent) {
            records = records.filter(record => 
                (record.AgentName || '').toLowerCase().includes(agent.toLowerCase())
            );
        }
        
        if (!records.length) {
            return 'No data available for the selected criteria.';
        }
        
        // Generate CSV
        const headers = INDEPENDENCE_QA_HEADERS;
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
        console.error('Error exporting Independence QA to CSV:', error);
        safeWriteError('exportIndependenceQAToCSV', error);
        return 'Error generating CSV export.';
    }
}

/**
 * Client function to export Independence QA to CSV
 */
function clientExportIndependenceQAToCSV(granularity = 'Week', period = null, agent = '') {
    try {
        return exportIndependenceQAToCSV(granularity, period, agent);
    } catch (error) {
        console.error('Error in clientExportIndependenceQAToCSV:', error);
        safeWriteError('clientExportIndependenceQAToCSV', error);
        return 'Error generating CSV export.';
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Utility Functions for Date/Time Handling
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get current week string
 */
function getCurrentWeekString() {
    return getWeekString(new Date());
}

/**
 * Get week string from date (compatible with existing system)
 */
function weekStringFromDate(date) {
    return getWeekString(date);
}

/**
 * Generate period options for dropdowns
 */
function getIndependencePeriodOptions(granularity = 'Week') {
    const options = [];
    const now = new Date();
    
    switch (granularity) {
        case 'Week':
            // Generate last 12 weeks
            for (let i = 0; i < 12; i++) {
                const date = new Date(now.getTime() - (i * 7 * 24 * 60 * 60 * 1000));
                options.push({
                    value: getWeekString(date),
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
 * Client function to get period options
 */
function clientGetIndependencePeriodOptions(granularity = 'Week') {
    try {
        return getIndependencePeriodOptions(granularity);
    } catch (error) {
        console.error('Error in clientGetIndependencePeriodOptions:', error);
        safeWriteError('clientGetIndependencePeriodOptions', error);
        return [];
    }
}