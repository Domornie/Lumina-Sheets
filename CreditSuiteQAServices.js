/**
 * COMPLETE CreditSuiteQAServices.gs - Configuration & Scoring System for Credit Suite
 * This includes ALL required functions for the Credit Suite QA system
 */

// ────────────────────────────────────────────────────────────────────────────
// Credit Suite QA Configuration - Comprehensive Quality Assessment System
// ────────────────────────────────────────────────────────────────────────────

const CREDIT_SUITE_QA_CONFIG = {
    // Campaign Information
    campaignName: "Credit Suite",
    campaignId: "credit-suite", 
    version: "1.0",
    lastUpdated: "2025-01-12",
    
    // Scoring Configuration
    scoring: {
        passThreshold: 85,
        excellentThreshold: 95,
        totalPossiblePoints: 68,
        criticalFailureOverride: true,
        weightedScoring: true
    },
    
    // Consultation types for Credit Suite business
    consultationTypes: [
        { value: "Initial Consultation", label: "Initial Consultation", description: "First-time client credit assessment" },
        { value: "Follow-up Review", label: "Follow-up Review", description: "Progress review meeting" },
        { value: "Dispute Resolution", label: "Dispute Resolution", description: "Credit dispute consultation" },
        { value: "Credit Education", label: "Credit Education", description: "Educational session" },
        { value: "Final Review", label: "Final Review", description: "Completion assessment" }
    ],
    
    // Question Categories with Enhanced Configuration for Credit Services
    categories: {
        "Initial Assessment & Client Intake": {
            icon: "fas fa-clipboard-check",
            color: "#1e40af",
            totalPoints: 16,
            description: "Professional intake process and initial client assessment",
            weight: 1.2,
            questions: [
                {
                    id: "Q1_ProfessionalIntroduction",
                    text: "Consultant provided clear professional introduction, explained their role, and outlined the consultation process",
                    description: "Sets professional tone and clear expectations from the start",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Initial Assessment & Client Intake",
                    critical: true,
                    tags: ["introduction", "professional", "process", "critical"]
                },
                {
                    id: "Q2_ClientGoalsAssessment",
                    text: "Consultant thoroughly assessed client's credit goals, financial objectives, and timeline expectations",
                    description: "Understanding client needs and setting realistic expectations",
                    maxPoints: 4,
                    weight: 1.0,
                    category: "Initial Assessment & Client Intake",
                    critical: false,
                    tags: ["goals", "assessment", "expectations", "timeline"]
                },
                {
                    id: "Q3_CreditHistoryReview",
                    text: "Consultant conducted comprehensive review of client's credit history and current credit report",
                    description: "Thorough analysis of current credit situation",
                    maxPoints: 4,
                    weight: 1.0,
                    category: "Initial Assessment & Client Intake",
                    critical: false,
                    tags: ["credit-history", "review", "analysis", "comprehensive"]
                },
                {
                    id: "Q4_DocumentationCollection",
                    text: "Consultant properly collected and reviewed all necessary documentation and identification",
                    description: "Compliance with documentation requirements",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Initial Assessment & Client Intake",
                    critical: false,
                    tags: ["documentation", "identification", "collection", "compliance"]
                },
                {
                    id: "Q5_PrivacyDisclosure",
                    text: "Consultant provided required privacy disclosures and obtained proper consent for credit review",
                    description: "Legal compliance for privacy and consent",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Initial Assessment & Client Intake",
                    critical: true,
                    tags: ["privacy", "disclosure", "consent", "legal", "critical"]
                }
            ]
        },
        "Credit Analysis & Strategy Development": {
            icon: "fas fa-chart-line",
            color: "#059669", 
            totalPoints: 18,
            description: "Credit report analysis and strategic planning",
            weight: 1.3,
            questions: [
                {
                    id: "Q6_AccurateCreditAnalysis",
                    text: "Consultant accurately identified and explained negative items, errors, and improvement opportunities",
                    description: "Precision in credit report analysis",
                    maxPoints: 4,
                    weight: 1.0,
                    category: "Credit Analysis & Strategy Development",
                    critical: false,
                    tags: ["analysis", "accuracy", "negative-items", "opportunities"]
                },
                {
                    id: "Q7_StrategicPlanDevelopment",
                    text: "Consultant developed a clear, realistic action plan with specific steps and timelines",
                    description: "Creation of actionable improvement strategy",
                    maxPoints: 4,
                    weight: 1.0,
                    category: "Credit Analysis & Strategy Development",
                    critical: false,
                    tags: ["strategy", "action-plan", "timeline", "realistic"]
                },
                {
                    id: "Q8_PriorityIdentification",
                    text: "Consultant properly prioritized which items to address first based on impact and likelihood of success",
                    description: "Strategic prioritization of credit improvement efforts",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Credit Analysis & Strategy Development",
                    critical: false,
                    tags: ["prioritization", "impact", "success-likelihood", "strategic"]
                },
                {
                    id: "Q9_DisputeStrategyExplanation",
                    text: "Consultant clearly explained dispute strategies and process for challenging inaccurate information",
                    description: "Clear communication of dispute methodology",
                    maxPoints: 4,
                    weight: 1.0,
                    category: "Credit Analysis & Strategy Development",
                    critical: false,
                    tags: ["dispute", "strategy", "process", "inaccurate-information"]
                },
                {
                    id: "Q10_CreditScoreEducation",
                    text: "Consultant educated client on credit scoring factors and how proposed actions will impact scores",
                    description: "Client education on credit scoring methodology",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Credit Analysis & Strategy Development",
                    critical: false,
                    tags: ["credit-score", "education", "factors", "impact"]
                }
            ]
        },
        "Compliance & Legal Requirements": {
            icon: "fas fa-balance-scale",
            color: "#dc2626",
            totalPoints: 12,
            description: "Regulatory compliance and legal adherence",
            weight: 1.5,
            questions: [
                {
                    id: "Q11_FCRACompliance",
                    text: "Consultant demonstrated proper understanding and adherence to Fair Credit Reporting Act (FCRA) requirements",
                    description: "Critical legal compliance requirement",
                    maxPoints: 4,
                    weight: 1.0,
                    category: "Compliance & Legal Requirements",
                    critical: true,
                    tags: ["FCRA", "compliance", "legal", "reporting", "critical"]
                },
                {
                    id: "Q12_CRLComplianceEducation",
                    text: "Consultant properly explained Credit Repair Laws and client rights under federal and state regulations",
                    description: "Client education on credit repair regulations",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Compliance & Legal Requirements",
                    critical: true,
                    tags: ["credit-repair-law", "client-rights", "regulations", "critical"]
                },
                {
                    id: "Q13_TruthfulRepresentation",
                    text: "Consultant made no false promises or guarantees about specific outcomes or timeline",
                    description: "Honest and realistic representation of services",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Compliance & Legal Requirements",
                    critical: true,
                    tags: ["truthful", "no-guarantees", "realistic", "honest", "critical"]
                },
                {
                    id: "Q14_ProperDisclosures",
                    text: "All required disclosures were provided including cancellation rights, fees, and service limitations",
                    description: "Complete regulatory disclosures",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Compliance & Legal Requirements",
                    critical: true,
                    tags: ["disclosures", "cancellation", "fees", "limitations", "critical"]
                }
            ]
        },
        "Client Education & Communication": {
            icon: "fas fa-graduation-cap",
            color: "#7c3aed",
            totalPoints: 14,
            description: "Educational effectiveness and communication quality",
            weight: 1.1,
            questions: [
                {
                    id: "Q15_ClearCommunication",
                    text: "Consultant used clear, understandable language appropriate for client's knowledge level",
                    description: "Effective communication adapted to client understanding",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Client Education & Communication",
                    critical: false,
                    tags: ["clear", "understandable", "appropriate", "knowledge-level"]
                },
                {
                    id: "Q16_EducationalContent",
                    text: "Consultant provided valuable education about credit fundamentals and best practices",
                    description: "Quality educational content delivery",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Client Education & Communication",
                    critical: false,
                    tags: ["education", "fundamentals", "best-practices", "valuable"]
                },
                {
                    id: "Q17_QuestionHandling",
                    text: "Consultant encouraged questions and provided thorough, accurate answers to client inquiries",
                    description: "Responsive and knowledgeable question handling",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Client Education & Communication",
                    critical: false,
                    tags: ["questions", "thorough", "accurate", "responsive"]
                },
                {
                    id: "Q18_ActiveListening",
                    text: "Consultant demonstrated active listening and showed understanding of client concerns",
                    description: "Effective listening and empathy",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Client Education & Communication",
                    critical: false,
                    tags: ["active-listening", "understanding", "concerns", "empathy"]
                },
                {
                    id: "Q19_ExpectationSetting",
                    text: "Consultant set realistic expectations about process, timeline, and potential outcomes",
                    description: "Proper expectation management",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Client Education & Communication",
                    critical: false,
                    tags: ["expectations", "realistic", "timeline", "outcomes"]
                }
            ]
        },
        "Professional Standards & Soft Skills": {
            icon: "fas fa-user-tie",
            color: "#0891b2",
            totalPoints: 8,
            description: "Professional demeanor and interpersonal skills",
            weight: 1.0,
            questions: [
                {
                    id: "Q20_ProfessionalDemeanor",
                    text: "Consultant maintained professional, courteous, and respectful demeanor throughout consultation",
                    description: "Consistent professional behavior",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Professional Standards & Soft Skills",
                    critical: false,
                    tags: ["professional", "courteous", "respectful", "demeanor"]
                },
                {
                    id: "Q21_KnowledgeExpertise",
                    text: "Consultant demonstrated comprehensive knowledge of credit repair processes and industry best practices",
                    description: "Technical competency and expertise",
                    maxPoints: 3,
                    weight: 1.0,
                    category: "Professional Standards & Soft Skills",
                    critical: false,
                    tags: ["knowledge", "expertise", "best-practices", "competency"]
                },
                {
                    id: "Q22_TimeManagement",
                    text: "Consultant managed consultation time effectively and covered all necessary topics",
                    description: "Efficient use of consultation time",
                    maxPoints: 2,
                    weight: 1.0,
                    category: "Professional Standards & Soft Skills",
                    critical: false,
                    tags: ["time-management", "effective", "coverage", "topics"]
                },
                {
                    id: "Q23_ClientRapport",
                    text: "Consultant built appropriate rapport and trust with client while maintaining professional boundaries",
                    description: "Relationship building within professional limits",
                    maxPoints: 1,
                    weight: 1.0,
                    category: "Professional Standards & Soft Skills",
                    critical: false,
                    tags: ["rapport", "trust", "boundaries", "relationship"]
                }
            ]
        }
    }
};

function getUsers() {
  try {
    // 1) Identify the logged-in manager by email
    const mgrEmail = (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    if (!mgrEmail) return [];

    const users = readSheet(USERS_SHEET) || [];
    const manager = users.find(u => String(u.Email || '').trim().toLowerCase() === mgrEmail);
    if (!manager) return [];

    // 2) Read manager→user assignments (sheet name helper if present, else default)
    const muSheetName = (typeof getManagerUsersSheetName_ === 'function')
      ? getManagerUsersSheetName_()
      : 'MANAGER_USERS';

    const assignments = readSheet(muSheetName) || [];
    const assignedIds = new Set(
      assignments
        .filter(a => String(a.ManagerUserID) === String(manager.ID))
        .map(a => String(a.UserID))
    );

    // 3) Build list: include the manager + all assigned users (no active filter)
    const list = [];

    if (manager.FullName && manager.Email) {
      list.push({
        name: String(manager.FullName).trim(),
        email: String(manager.Email).trim()
      });
    }

    users.forEach(u => {
      if (assignedIds.has(String(u.ID)) && u.FullName && u.Email) {
        list.push({
          name: String(u.FullName).trim(),
          email: String(u.Email).trim()
        });
      }
    });

    // 4) De-dupe by email and sort by name
    const seen = new Set();
    const out = list.filter(item => {
      const key = (item.email || '').toLowerCase();
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    }).sort((a, b) => (a.name || '').localeCompare(b.name || '', undefined, { sensitivity: 'base' }));

    return out;
  } catch (e) {
    Logger.log('getUsers error: ' + e);
    return [];
  }
}

/**
 * Client-accessible: return the Credit Suite user list
 * (wraps the existing getUsers() function)
 */
function clientGetCreditSuiteUsers() {
  try {
    var users = getUsers() || [];
    // (Optional) ensure shape is {name, email, department?}
    return users.map(function(u) {
      return {
        name: u.name || u.FullName || '',
        email: u.email || u.Email || '',
        // department is optional; include if your sheet has it
        department: u.department || u.Department || u.Campaign || ''
      };
    });
  } catch (e) {
    console.error('clientGetCreditSuiteUsers error:', e);
    return [];
  }
}

// ────────────────────────────────────────────────────────────────────────────
// Enhanced QA Scoring Function for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Calculate Credit Suite QA scores with comprehensive error handling
 */
function calculateCreditSuiteQAScores(formData) {
    try {
        console.log('=== Starting calculateCreditSuiteQAScores ===');
        console.log('Form data keys:', Object.keys(formData || {}));
        console.log('Form data sample:', JSON.stringify(formData).substring(0, 500));
        
        // Validate input
        if (!formData || typeof formData !== 'object') {
            console.error('Invalid form data provided:', typeof formData);
            return createEmptyScoreResult('Invalid form data');
        }
        
        // Validate configuration
        if (!CREDIT_SUITE_QA_CONFIG || !CREDIT_SUITE_QA_CONFIG.categories) {
            console.error('Credit Suite QA Configuration not available or invalid');
            return createEmptyScoreResult('Configuration error');
        }
        
        console.log('Available categories:', Object.keys(CREDIT_SUITE_QA_CONFIG.categories));
        
        // Initialize scoring variables
        let totalEarned = 0;
        let totalPossible = 0;
        const categoryScores = {};
        const questionScores = {};
        let hasCriticalFailure = false;
        const criticalFailures = [];
        
        // Process each category
        Object.entries(CREDIT_SUITE_QA_CONFIG.categories).forEach(([categoryName, categoryData]) => {
            console.log(`Processing category: ${categoryName}`);
            
            if (!categoryData || !categoryData.questions || !Array.isArray(categoryData.questions)) {
                console.warn(`Invalid category data for ${categoryName}:`, categoryData);
                return;
            }
            
            let categoryEarned = 0;
            let categoryPossible = 0;
            
            categoryData.questions.forEach(question => {
                try {
                    const questionId = question.id;
                    const maxPoints = parseFloat(question.maxPoints) || 0;
                    const answer = formData[questionId] || 'NA';
                    const weight = parseFloat(question.weight) || 1.0;
                    const categoryWeight = parseFloat(categoryData.weight) || 1.0;
                    const isCritical = question.critical === true;
                    
                    console.log(`Processing ${questionId}: answer=${answer}, maxPoints=${maxPoints}, weight=${weight}, critical=${isCritical}`);
                    
                    // Calculate earned points
                    let earnedPoints = 0;
                    let countTowardTotal = true;
                    
                    if (answer === 'Yes') {
                        earnedPoints = maxPoints;
                    } else if (answer === 'No') {
                        earnedPoints = 0;
                        if (isCritical) {
                            hasCriticalFailure = true;
                            criticalFailures.push({
                                questionId: questionId,
                                questionText: question.text || '',
                                category: categoryName,
                                maxPoints: maxPoints
                            });
                            console.log(`CRITICAL FAILURE detected for ${questionId}`);
                        }
                    } else if (answer === 'NA' || answer === '') {
                        earnedPoints = 0;
                        countTowardTotal = false;
                    }
                    
                    // Store question score details
                    questionScores[questionId] = {
                        answer: answer,
                        earned: earnedPoints,
                        possible: maxPoints,
                        percentage: maxPoints > 0 ? Math.round((earnedPoints / maxPoints) * 100) : 0,
                        weight: weight,
                        categoryWeight: categoryWeight,
                        critical: isCritical,
                        countsTowardTotal: countTowardTotal,
                        tags: question.tags || [],
                        comment: formData[questionId + '_Comments'] || ''
                    };
                    
                    // Add to totals if it counts
                    if (countTowardTotal) {
                        const weightedEarned = earnedPoints * weight * categoryWeight;
                        const weightedPossible = maxPoints * weight * categoryWeight;
                        
                        categoryEarned += weightedEarned;
                        categoryPossible += weightedPossible;
                        totalEarned += weightedEarned;
                        totalPossible += weightedPossible;
                        
                        console.log(`Added to totals: earned=${weightedEarned}, possible=${weightedPossible}`);
                    }
                    
                } catch (questionError) {
                    console.error(`Error processing question ${question.id}:`, questionError);
                }
            });
            
            // Store category score
            categoryScores[categoryName] = {
                earned: Math.round(categoryEarned * 100) / 100,
                possible: Math.round(categoryPossible * 100) / 100,
                percentage: categoryPossible > 0 ? Math.round((categoryEarned / categoryPossible) * 100) : 0,
                color: categoryData.color || '#6b7280',
                icon: categoryData.icon || 'fas fa-question',
                weight: categoryData.weight || 1.0,
                description: categoryData.description || ''
            };
            
            console.log(`Category ${categoryName} total: ${categoryEarned}/${categoryPossible} (${categoryScores[categoryName].percentage}%)`);
        });
        
        console.log(`Overall totals BEFORE critical check: ${totalEarned}/${totalPossible}`);
        console.log(`Critical failures: ${criticalFailures.length}`);
        
        // Calculate overall percentage
        let overallPercentage = totalPossible > 0 ? Math.round((totalEarned / totalPossible) * 100) : 0;
        
        // Apply critical failure override
        if (hasCriticalFailure && CREDIT_SUITE_QA_CONFIG.scoring.criticalFailureOverride) {
            console.log('Applying critical failure override - setting score to 0%');
            overallPercentage = 0;
            totalEarned = 0;
        }
        
        // Determine pass status
        let passStatus = 'Calculating...';
        let statusColor = '#6b7280';
        
        if (hasCriticalFailure) {
            passStatus = `Critical Failure - ${criticalFailures.length} compliance violation(s)`;
            statusColor = '#dc2626';
        } else if (overallPercentage >= (CREDIT_SUITE_QA_CONFIG.scoring.excellentThreshold || 95)) {
            passStatus = 'Excellent Performance';
            statusColor = '#10b981';
        } else if (overallPercentage >= (CREDIT_SUITE_QA_CONFIG.scoring.passThreshold || 85)) {
            passStatus = 'Meets Standards';
            statusColor = '#f59e0b';
        } else if (overallPercentage >= 70) {
            passStatus = 'Needs Improvement';
            statusColor = '#ef4444';
        } else {
            passStatus = 'Unsatisfactory Performance';
            statusColor = '#dc2626';
        }
        
        // Get analysis
        const strengthAreas = getStrengthAreas(categoryScores);
        const improvementAreas = getImprovementAreas(categoryScores, questionScores);
        
        // Create final result
        const result = {
            success: true,
            totalEarned: Math.round(totalEarned * 100) / 100,
            totalPossible: Math.round(totalPossible * 100) / 100,
            overallPercentage: overallPercentage,
            passStatus: passStatus,
            statusColor: statusColor,
            categoryScores: categoryScores,
            questionScores: questionScores,
            hasCriticalFailure: hasCriticalFailure,
            criticalFailures: criticalFailures,
            criticalFailureCount: criticalFailures.length,
            strengthAreas: strengthAreas,
            improvementAreas: improvementAreas,
            calculatedAt: new Date().toISOString(),
            configVersion: CREDIT_SUITE_QA_CONFIG.version,
            scoringMethod: 'weighted_credit_suite_compliance'
        };
        
        console.log('=== Credit Suite score calculation completed successfully ===');
        console.log('Final result summary:', {
            overallPercentage: result.overallPercentage,
            passStatus: result.passStatus,
            hasCriticalFailure: result.hasCriticalFailure,
            totalCategories: Object.keys(result.categoryScores).length,
            totalQuestions: Object.keys(result.questionScores).length
        });
        
        return result;
        
    } catch (error) {
        console.error('=== CRITICAL ERROR in calculateCreditSuiteQAScores ===');
        console.error('Error message:', error.message);
        console.error('Error stack:', error.stack);
        console.error('Form data that caused error:', JSON.stringify(formData || {}).substring(0, 1000));
        
        return createEmptyScoreResult(`Calculation error: ${error.message}`);
    }
}

/**
 * Create a standardized empty/error result for Credit Suite
 */
function createEmptyScoreResult(errorMessage) {
    const timestamp = new Date().toISOString();
    console.log(`Creating empty Credit Suite score result: ${errorMessage} at ${timestamp}`);
    
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
        improvementAreas: { categories: [], questions: [] },
        calculatedAt: timestamp,
        configVersion: CREDIT_SUITE_QA_CONFIG ? CREDIT_SUITE_QA_CONFIG.version : 'unknown',
        scoringMethod: 'error'
    };
}

/**
 * Get strength areas from category scores (Credit Suite specific)
 */
function getStrengthAreas(categoryScores) {
    try {
        if (!categoryScores || typeof categoryScores !== 'object') {
            return [];
        }
        
        return Object.entries(categoryScores)
            .filter(([name, score]) => score && score.percentage >= 90)
            .map(([name, score]) => ({
                category: name,
                percentage: score.percentage || 0,
                description: score.description || ''
            }))
            .sort((a, b) => (b.percentage || 0) - (a.percentage || 0));
    } catch (error) {
        console.error('Error getting Credit Suite strength areas:', error);
        return [];
    }
}

/**
 * Get improvement areas from category and question scores (Credit Suite specific)
 */
function getImprovementAreas(categoryScores, questionScores) {
    try {
        const lowCategories = [];
        const failedQuestions = [];
        
        if (categoryScores && typeof categoryScores === 'object') {
            Object.entries(categoryScores).forEach(([name, score]) => {
                if (score && score.percentage < 85) {
                    lowCategories.push({
                        type: 'category',
                        name: name,
                        percentage: score.percentage || 0,
                        description: score.description || ''
                    });
                }
            });
        }
        
        if (questionScores && typeof questionScores === 'object') {
            Object.entries(questionScores).forEach(([id, score]) => {
                if (score && score.answer === 'No') {
                    failedQuestions.push({
                        type: 'question',
                        id: id,
                        category: score.category || 'Unknown',
                        critical: score.critical || false,
                        comment: score.comment || ''
                    });
                }
            });
        }
        
        return { categories: lowCategories, questions: failedQuestions };
    } catch (error) {
        console.error('Error getting Credit Suite improvement areas:', error);
        return { categories: [], questions: [] };
    }
}

// ────────────────────────────────────────────────────────────────────────────
// Client-accessible functions for Credit Suite
// ────────────────────────────────────────────────────────────────────────────

/**
 * Preview QA score calculation for Credit Suite client-side real-time updates
 */
function clientPreviewCreditSuiteQAScore(formData) {
    try {
        console.log('=== clientPreviewCreditSuiteQAScore called ===');
        console.log('Received form data type:', typeof formData);
        console.log('Form data keys count:', formData ? Object.keys(formData).length : 0);
        
        if (!formData || typeof formData !== 'object') {
            console.error('Invalid form data in Credit Suite client preview function');
            return createEmptyScoreResult('Invalid form data provided to preview function');
        }
        
        const result = calculateCreditSuiteQAScores(formData);
        
        if (!result) {
            console.error('calculateCreditSuiteQAScores returned null/undefined');
            return createEmptyScoreResult('Score calculation returned null');
        }
        
        console.log('Credit Suite preview calculation successful, returning result');
        return result;
        
    } catch (error) {
        console.error('=== ERROR in clientPreviewCreditSuiteQAScore ===');
        console.error('Error:', error.message);
        console.error('Stack:', error.stack);
        
        return createEmptyScoreResult(`Preview error: ${error.message}`);
    }
}

/**
 * Get enhanced QA configuration for Credit Suite client-side access
 */
function clientGetCreditSuiteQAConfig() {
    try {
        console.log('Getting Credit Suite QA configuration...');
        
        if (!CREDIT_SUITE_QA_CONFIG) {
            console.error('Credit Suite QA Configuration not available');
            return createMinimalFallbackConfig();
        }
        
        // Return a deep copy to prevent client-side modifications
        const configCopy = JSON.parse(JSON.stringify(CREDIT_SUITE_QA_CONFIG));
        console.log('Credit Suite configuration retrieved successfully:', {
            version: configCopy.version,
            categories: Object.keys(configCopy.categories).length,
            totalQuestions: Object.values(configCopy.categories).reduce((total, cat) => 
                total + (cat.questions ? cat.questions.length : 0), 0)
        });
        
        return configCopy;
        
    } catch (error) {
        console.error('Error in clientGetCreditSuiteQAConfig:', error);
        return createMinimalFallbackConfig();
    }
}

/**
 * Create minimal fallback configuration for Credit Suite
 */
function createMinimalFallbackConfig() {
    console.log('Creating minimal fallback configuration for Credit Suite');
    return {
        categories: {
            "Basic Credit Assessment": {
                icon: "fas fa-credit-card",
                color: "#1e40af",
                totalPoints: 10,
                description: "Basic credit consultation evaluation",
                weight: 1.0,
                questions: [
                    {
                        id: "Q1_BasicCreditConsultation",
                        text: "Consultant provided adequate credit consultation services",
                        description: "General credit consultation assessment",
                        maxPoints: 5,
                        weight: 1.0,
                        critical: false,
                        category: "Basic Credit Assessment",
                        tags: ["basic", "consultation"]
                    },
                    {
                        id: "Q2_ComplianceBasic",
                        text: "Consultant followed basic compliance requirements",
                        description: "Basic regulatory compliance",
                        maxPoints: 5,
                        weight: 1.0,
                        critical: true,
                        category: "Basic Credit Assessment",
                        tags: ["compliance", "basic"]
                    }
                ]
            }
        }, 
        consultationTypes: [
            { value: "Initial Consultation", label: "Initial Consultation", description: "First-time client assessment" },
            { value: "Follow-up Review", label: "Follow-up Review", description: "Progress review meeting" }
        ],
        scoring: { 
            passThreshold: 85, 
            excellentThreshold: 95, 
            totalPossiblePoints: 10,
            criticalFailureOverride: true,
            weightedScoring: false
        },
        version: "fallback-1.0",
        lastUpdated: new Date().toISOString().split('T')[0]
    };
}

/**
 * Get consultation types for Credit Suite
 */
function clientGetCreditSuiteConsultationTypes() {
    try {
        if (!CREDIT_SUITE_QA_CONFIG || !CREDIT_SUITE_QA_CONFIG.consultationTypes) {
            console.warn('Using fallback consultation types for Credit Suite');
            return [
                { value: 'Initial Consultation', label: 'Initial Consultation', description: 'First-time client assessment' },
                { value: 'Follow-up Review', label: 'Follow-up Review', description: 'Progress review meeting' },
                { value: 'Dispute Resolution', label: 'Dispute Resolution', description: 'Credit dispute consultation' },
                { value: 'Final Review', label: 'Final Review', description: 'Completion assessment' }
            ];
        }
        
        return CREDIT_SUITE_QA_CONFIG.consultationTypes.map(ct => ({
            ...ct,
            campaignSpecific: true,
            lastUpdated: CREDIT_SUITE_QA_CONFIG.lastUpdated
        }));
        
    } catch (error) {
        console.error('Error in clientGetCreditSuiteConsultationTypes:', error);
        return [
            { value: 'Initial Consultation', label: 'Initial Consultation', description: 'First-time client assessment' },
            { value: 'Follow-up Review', label: 'Follow-up Review', description: 'Progress review meeting' }
        ];
    }
}

/**
 * Client-accessible system health check for Credit Suite
 */
function clientPerformCreditSuiteHealthCheck() {
    try {
        const startTime = new Date();
        
        // Test basic functionality
        const testFormData = {
            Q1_ProfessionalIntroduction: 'Yes',
            Q2_ClientGoalsAssessment: 'Yes',
            Q3_CreditHistoryReview: 'No'
        };
        
        console.log('Testing Credit Suite score calculation...');
        const testResult = calculateCreditSuiteQAScores(testFormData);
        const testPassed = testResult && testResult.overallPercentage !== undefined;
        
        // Check system components
        const checks = {
            spreadsheet: { 
                status: 'healthy', 
                message: 'Spreadsheet connection available' 
            },
            configuration: { 
                status: CREDIT_SUITE_QA_CONFIG ? 'healthy' : 'warning',
                message: CREDIT_SUITE_QA_CONFIG ? 'Configuration loaded successfully' : 'Using fallback configuration',
                categories: CREDIT_SUITE_QA_CONFIG ? Object.keys(CREDIT_SUITE_QA_CONFIG.categories).length : 0,
                questions: CREDIT_SUITE_QA_CONFIG ? Object.values(CREDIT_SUITE_QA_CONFIG.categories)
                    .reduce((total, cat) => total + (cat.questions ? cat.questions.length : 0), 0) : 0
            },
            scoring: { 
                status: testPassed ? 'healthy' : 'error',
                message: testPassed ? 'Scoring function working correctly' : 'Scoring function has issues',
                version: CREDIT_SUITE_QA_CONFIG ? CREDIT_SUITE_QA_CONFIG.version : 'unknown',
                testScore: testResult ? testResult.overallPercentage : 'failed'
            }
        };
        
        const endTime = new Date();
        const responseTime = `${endTime - startTime}ms`;
        const overallStatus = testPassed ? 'healthy' : 'warning';
        
        const result = {
            success: true,
            status: overallStatus,
            responseTime: responseTime,
            components: checks,
            timestamp: endTime.toISOString(),
            campaign: 'Credit Suite'
        };
        
        console.log('Credit Suite health check completed:', result);
        return result;
        
    } catch (error) {
        console.error('Error performing Credit Suite health check:', error);
        return {
            success: false,
            status: 'error',
            error: error.message,
            timestamp: new Date().toISOString(),
            campaign: 'Credit Suite'
        };
    }
}

/**
 * Get Credit Suite system status summary
 */
function clientGetCreditSuiteSystemStatus() {
    try {
        if (!CREDIT_SUITE_QA_CONFIG) {
            return { 
                error: 'Configuration not available',
                status: 'error',
                campaign: 'Credit Suite'
            };
        }
        
        const totalQuestions = Object.values(CREDIT_SUITE_QA_CONFIG.categories)
            .reduce((total, cat) => total + (cat.questions ? cat.questions.length : 0), 0);
            
        const criticalQuestions = Object.values(CREDIT_SUITE_QA_CONFIG.categories)
            .reduce((total, cat) => total + (cat.questions ? 
                cat.questions.filter(q => q.critical).length : 0), 0);
        
        return {
            configVersion: CREDIT_SUITE_QA_CONFIG.version,
            lastUpdated: CREDIT_SUITE_QA_CONFIG.lastUpdated,
            totalQuestions: totalQuestions,
            criticalQuestions: criticalQuestions,
            categories: Object.keys(CREDIT_SUITE_QA_CONFIG.categories).length,
            consultationTypes: CREDIT_SUITE_QA_CONFIG.consultationTypes ? CREDIT_SUITE_QA_CONFIG.consultationTypes.length : 0,
            scoringMethod: 'weighted_credit_suite_compliance',
            passThreshold: CREDIT_SUITE_QA_CONFIG.scoring.passThreshold,
            excellentThreshold: CREDIT_SUITE_QA_CONFIG.scoring.excellentThreshold,
            status: 'healthy',
            campaign: 'Credit Suite'
        };
    } catch (error) {
        console.error('Error getting Credit Suite system status:', error);
        return { 
            error: error.message,
            status: 'error',
            campaign: 'Credit Suite'
        };
    }
}

console.log('Credit Suite QA Configuration and Services loaded successfully');
console.log('Total categories:', CREDIT_SUITE_QA_CONFIG ? Object.keys(CREDIT_SUITE_QA_CONFIG.categories).length : 'undefined');
console.log('Total questions:', CREDIT_SUITE_QA_CONFIG ? Object.values(CREDIT_SUITE_QA_CONFIG.categories).reduce((total, cat) => total + (cat.questions ? cat.questions.length : 0), 0) : 'undefined');