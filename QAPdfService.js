/**
 * QAPdfService.gs — Dedicated PDF Generation Service
 * 
 * This service handles all PDF generation for QA records with multiple templates
 * and enhanced HTML rendering capabilities.
 */
function safeGetPerformanceBand_(scoreResult) {
  if (scoreResult && scoreResult.performanceBand && scoreResult.performanceBand.label) {
    return scoreResult.performanceBand;
  }

  // Fallback: calculate performance band from finalScore
  const score = scoreResult ? (scoreResult.finalScore || 0) : 0;

  if (score >= 100) {
    return { label: 'Excellent', description: 'Exceeds expectations', color: '#0E67E3' }; // brand primary
  } else if (score >= 95) {
    return { label: 'Good', description: 'Meets expectations', color: '#1D499B' }; // brand primary-dark
  } else if (score >= 90) {
    return { label: 'Satisfactory', description: 'Meets minimum standards', color: '#9AAED0' }; // brand muted
  } else if (score >= 80) {
    return { label: 'Needs Improvement', description: 'Below expectations', color: '#1A314C' }; // brand slate
  } else {
    return { label: 'Unsatisfactory', description: 'Significant improvement required', color: '#DC2626' }; // danger red (kept for clarity)
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// PDF GENERATION CORE SERVICE
// ═══════════════════════════════════════════════════════════════════════════════

class QAPdfService {
  constructor() {
    this.defaultOptions = {
      template: 'standard',
      includeCharts: true,
      includeRecommendations: true,
      includeFullSnapshot: true,
      theme: 'professional'
    };
  }

  /**
   * Main PDF generation function
   */
  generateQaPdf(qaRecord, scoreResult, options = {}) {
    try {
      const config = { ...this.defaultOptions, ...options };

      // Generate HTML content based on template
      const htmlContent = this.generateHtmlTemplate(qaRecord, scoreResult, config);

      // Create PDF blob
      const blob = Utilities.newBlob(htmlContent, 'text/html')
        .getAs('application/pdf')
        .setName(this.generatePdfFilename(qaRecord));

      // Store in appropriate folder structure
      const folder = this.getOrCreateQaFolder(qaRecord.AgentName, qaRecord.CallDate);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      return {
        success: true,
        fileId: file.getId(),
        fileUrl: file.getUrl(),
        fileName: file.getName(),
        template: config.template,
        generatedAt: new Date().toISOString()
      };

    } catch (error) {
      console.error('Error generating QA PDF:', error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  /**
   * Generate HTML template based on type
   */
  generateHtmlTemplate(qaRecord, scoreResult, config) {
    switch (config.template) {
      case 'executive':
        return this.generateExecutiveTemplate(qaRecord, scoreResult, config);
      case 'detailed':
        return this.generateDetailedTemplate(qaRecord, scoreResult, config);
      case 'coaching':
        return this.generateCoachingTemplate(qaRecord, scoreResult, config);
      case 'simple':
        return this.generateSimpleTemplate(qaRecord, scoreResult, config);
      default:
        return this.generateStandardTemplate(qaRecord, scoreResult, config);
    }
  }

  /**
   * Generate standard template HTML
   */
  generateStandardTemplate(qaRecord, scoreResult, config) {
    const theme = this.getThemeStyles(config.theme);
    const performanceBadge = this.generatePerformanceBadge(scoreResult);
    const questionSections = this.generateQuestionSections(qaRecord, scoreResult);
    const charts = config.includeCharts ? this.generateChartsSection(scoreResult) : '';
    const recommendations = config.includeRecommendations ? this.generateRecommendations(scoreResult) : '';
    const snapshot = config.includeFullSnapshot ? this.generateDataSnapshot(qaRecord) : '';

    return `
      <!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>QA Audit Report - ${this.escapeHtml(qaRecord.AgentName)}</title>
          <style>${theme}</style>
      </head>
      <body>
          <div class="container">
              <!-- Header Section -->
                <div class="header">
                  <div class="brand-lockup">
                    <div class="brand-title">
                      <h1>Lumina</h1>
                      <p>high visibility command center</p>
                    </div>
                  </div>
                  <div class="brand-meta">
                    <p><strong>Generated:</strong> ${new Date().toLocaleDateString()}</p>
                    <p><strong>Report ID:</strong> ${qaRecord.ID}</p>
                  </div>
                </div>
              <!-- Title Banner -->
              <div class="title-banner">
                  <h2>Quality Audit Report</h2>
                  <h3>${this.escapeHtml(qaRecord.AgentName)} • ${qaRecord.CallDate}</h3>
              </div>

              <!-- Key Metrics Section -->
              <div class="metrics-grid">
                  <div class="metric-card score">
                      <div class="metric-label">Final Score</div>
                      <div class="metric-value">${scoreResult.finalScore}%</div>
                      ${performanceBadge}
                  </div>
                  <div class="metric-card">
                      <div class="metric-label">Points Earned</div>
                      <div class="metric-value">${scoreResult.earned}/${scoreResult.applicable}</div>
                  </div>
                  <div class="metric-card">
                      <div class="metric-label">Status</div>
                      <div class="metric-value ${scoreResult.isPassing ? 'passing' : 'failing'}">
                          ${scoreResult.isPassing ? 'PASS' : 'FAIL'}
                      </div>
                  </div>
                  <div class="metric-card">
                      <div class="metric-label">Auditor</div>
                      <div class="metric-value">${this.escapeHtml(qaRecord.AuditorName)}</div>
                  </div>
              </div>

              <!-- Agent Information -->
              <div class="info-section">
                  <h3>Call Information</h3>
                  <div class="info-grid">
                      <div class="info-item">
                          <strong>Agent:</strong> ${this.escapeHtml(qaRecord.AgentName)}
                      </div>
                      <div class="info-item">
                          <strong>Email:</strong> ${this.escapeHtml(qaRecord.AgentEmail)}
                      </div>
                      <div class="info-item">
                          <strong>Client:</strong> ${this.escapeHtml(qaRecord.ClientName)}
                      </div>
                      <div class="info-item">
                          <strong>Case Number:</strong> ${this.escapeHtml(qaRecord.CaseNumber)}
                      </div>
                      <div class="info-item">
                          <strong>Call Date:</strong> ${qaRecord.CallDate}
                      </div>
                      <div class="info-item">
                          <strong>Audit Date:</strong> ${qaRecord.AuditDate}
                      </div>
                      ${qaRecord.CallLink ? `
                      <div class="info-item full-width">
                          <strong>Recording:</strong> 
                          <a href="${qaRecord.CallLink}" target="_blank">Listen to Call</a>
                      </div>
                      ` : ''}
                  </div>
              </div>

              <!-- Question Sections -->
              ${questionSections}

              <!-- Charts Section -->
              ${charts}

              <!-- Overall Feedback -->
              <div class="feedback-section">
                  <h3>Overall Feedback</h3>
                  <div class="feedback-content">
                      ${this.renderRichContent(qaRecord.OverallFeedback)}
                  </div>
              </div>

              <!-- Notes and Agent Feedback -->
              <div class="notes-grid">
                  <div class="notes-section">
                      <h3>Auditor Notes</h3>
                      <div class="notes-content">
                          ${this.renderRichContent(qaRecord.Notes)}
                      </div>
                  </div>
                  <div class="notes-section">
                      <h3>Agent Response</h3>
                      <div class="notes-content">
                          ${this.renderRichContent(qaRecord.AgentFeedback)}
                      </div>
                  </div>
              </div>

              <!-- Recommendations -->
              ${recommendations}

              <!-- Data Snapshot -->
              ${snapshot}

              <!-- Footer -->
              <div class="footer">
                  <p>This report was generated automatically in Lumina QA System</p>
                  <p>Generated on ${new Date().toLocaleString()}</p>
              </div>
          </div>
      </body>
      </html>`;
  }

  /**
   * Generate coaching-focused template
   */
  generateCoachingTemplate(qaRecord, scoreResult, config) {
    const theme = this.getThemeStyles(config.theme);
    const actionPlan = this.generateActionPlan(scoreResult);
    const strengths = this.identifyStrengths(scoreResult);
    const improvements = this.identifyImprovements(scoreResult);

    // Use any of these fields if present, otherwise fall back to an optional config.signatureUrl
    const ackUrl = qaRecord.CoachingAckUrl
                || qaRecord.AcknowledgmentUrl
                || qaRecord.SignatureUrl
                || qaRecord.AgentSignatureUrl
                || config.signatureUrl
                || '';

    return `
      <!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Coaching Report - ${this.escapeHtml(qaRecord.AgentName)}</title>
          <style>${theme}</style>
      </head>
      <body>
      <div class="container coaching-theme">
        <!-- Header -->
        <div class="header">
          <div class="brand-lockup">
            <div class="brand-title">
              <h1>Quality Coaching Report</h1>
              <p>high visibility command center</p>
            </div>
          </div>
          <div class="brand-meta">
            <p><strong>Agent:</strong> ${this.escapeHtml(qaRecord.AgentName)}</p>
            <p><strong>Reviewed by:</strong> ${this.escapeHtml(qaRecord.AuditorName)}</p>
          </div>
        </div>

        <!-- Performance Overview -->
        <div class="performance-overview">
          <div class="score-display">
            <span class="score-number">${scoreResult.finalScore}%</span>
            <span class="score-label">${scoreResult.performanceBand && scoreResult.performanceBand.label ? scoreResult.performanceBand.label : 'Performance'}</span>
          </div>
        </div>

        <!-- Strengths Section -->
        <div class="coaching-section strengths">
          <h3>🌟 Strengths Demonstrated</h3>
          ${strengths}
        </div>

        <!-- Areas for Improvement -->
        <div class="coaching-section improvements">
          <h3>📈 Growth Opportunities</h3>
          ${improvements}
        </div>

        <!-- Action Plan -->
        <div class="coaching-section action-plan">
          <h3>🎯 Development Action Plan</h3>
          ${actionPlan}
        </div>

        <!-- Follow-up Section -->
        <div class="follow-up">
          <h3>Next Steps</h3>
          <div class="follow-up-grid">
            <div class="follow-up-item">
              <strong>Next Review:</strong> ${this.calculateNextReviewDate()}
            </div>
            <div class="follow-up-item">
              <strong>Focus Areas:</strong> ${this.getPriorityAreas(scoreResult)}
            </div>
          </div>
        </div>

        <!-- Digital Acknowledgment (replaces signature lines) -->
        <div class="acknowledgment" style="margin-top:24px;">
          <h3>Digital Acknowledgment</h3>
          ${
            ackUrl
              ? `<p>Complete the acknowledgment digitally here:</p>
                <a href="${this.escapeHtml(ackUrl)}" target="_blank" rel="noopener"
                    style="display:inline-block;padding:10px 14px;border-radius:999px;text-decoration:none;
                          background:var(--brand-primary);color:#fff;font-weight:700;">
                    Open Digital Acknowledgment
                </a>`
              : `<p>This coaching uses the digital acknowledgment template in Lumina.</p>`
          }
        </div>
      </div>
      </body>
      </html>`;
  }

  /**
   * Generate executive summary template
   */
  generateExecutiveTemplate(qaRecord, scoreResult, config) {
    const theme = this.getThemeStyles('executive');
    const kpiSummary = this.generateKpiSummary(scoreResult);
    const riskAssessment = this.generateRiskAssessment(scoreResult);

    return `
      <!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <title>Executive QA Summary - ${this.escapeHtml(qaRecord.AgentName)}</title>
          <style>${theme}</style>
      </head>
      <body>
          <div class="container executive-theme">
              <!-- Executive Header -->
              <div class="header">
                <div class="brand-lockup">
                  <div class="brand-title">
                    <h1>QA Executive Summary</h1>
                    <p>high visibility command center</p>
                  </div>
                </div>
                <div class="brand-meta">
                  <p><strong>Agent:</strong> ${this.escapeHtml(qaRecord.AgentName)}</p>
                  <p><strong>Date:</strong> ${qaRecord.CallDate}</p>
                  <p><strong>Score:</strong> ${scoreResult.finalScore}%</p>
                </div>
              </div>


              <!-- KPI Dashboard -->
              <div class="kpi-dashboard">
                  ${kpiSummary}
              </div>

              <!-- Key Insights -->
              <div class="insights-section">
                  <h2>Key Insights</h2>
                  <div class="insights-grid">
                      ${this.generateKeyInsights(scoreResult)}
                  </div>
              </div>

              <!-- Risk Assessment -->
              <div class="risk-section">
                  <h2>Risk Assessment</h2>
                  ${riskAssessment}
              </div>

              <!-- Recommendations -->
              <div class="exec-recommendations">
                  <h2>Strategic Recommendations</h2>
                  ${this.generateExecutiveRecommendations(scoreResult)}
              </div>
          </div>
      </body>
      </html>`;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // THEME AND STYLING
  // ═══════════════════════════════════════════════════════════════════════════

getThemeStyles(theme = 'professional') {
  const baseStyles = `
  @page { size: A4; margin: 15mm; }

  :root{
    --brand-bg: #F2F4F5;
    --brand-navy: #011836;
    --brand-slate: #1A314C;
    --brand-primary: #0E67E3;
    --brand-primary-dark: #1D499B;
    --brand-muted: #9AAED0;
    --brand-shadow: rgba(1,24,54,0.15);
    --danger: #DC2626;
  }

  *{ box-sizing:border-box; margin:0; padding:0; }

  body{
    font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    line-height:1.6;
    color:var(--brand-slate);
    background:#fff;
  }

  .container{ max-width:100%; margin:0 auto; padding:20px; }

  /* Header with brand */
  .header{
    display:flex; justify-content:space-between; align-items:center;
    border-bottom: 3px solid var(--brand-primary-dark);
    padding-bottom:18px; margin-bottom:28px;
  }

  .brand-meta{
    text-align:right; font-size:12px; color:#6B7280;
  }

  .brand-lockup{
    display:flex; align-items:center; gap:12px;
  }

  .brand-logo{
    width:140px; height:auto; display:block; 
    filter: drop-shadow(0 6px 16px var(--brand-shadow));
  }

  .brand-title{
    display:flex; flex-direction:column;
  }
  .brand-title h1{
    color:var(--brand-navy); font-size:24px; letter-spacing:0.2px;
  }
  .brand-title p{
    color:var(--brand-muted); font-size:13px; margin-top:2px;
  }

  /* Title Banner */
  .title-banner{
    background: linear-gradient(135deg, var(--brand-primary) 0%, var(--brand-primary-dark) 100%);
    color:#fff; padding:24px; border-radius:12px; margin-bottom:26px; text-align:center;
    box-shadow: 0 10px 24px var(--brand-shadow);
  }
  .title-banner h2{ font-size:28px; margin-bottom:6px; }
  .title-banner h3{ font-size:16px; opacity:0.95; font-weight:500; }

  /* Metrics */
  .metrics-grid{ display:grid; grid-template-columns:repeat(4,1fr); gap:18px; margin-bottom:26px; }
  .metric-card{
    background: var(--brand-bg);
    border:1px solid #E2E8F0; border-radius:10px; padding:18px; text-align:center;
    box-shadow: 0 6px 14px var(--brand-shadow);
  }
  .metric-card.score{
    background: linear-gradient(135deg, var(--brand-primary) 0%, var(--brand-primary-dark) 100%);
    color:#fff; border:none;
  }
  .metric-label{ font-size:11px; text-transform:uppercase; letter-spacing:.5px; margin-bottom:6px; opacity:.9; }
  .metric-value{ font-size:22px; font-weight:800; }
  .metric-value.passing{ color:#0EA5E9; } /* subtle cyan accent on light cards */
  .metric-value.failing{ color:var(--danger); }

  /* Performance badge mapped to brand palette */
  .performance-badge{
    display:inline-block; padding:4px 12px; border-radius:999px; font-size:11px;
    font-weight:800; margin-top:8px; text-transform:uppercase; letter-spacing:.5px;
    box-shadow: 0 4px 10px var(--brand-shadow);
  }
  .badge-excellent{ background: rgba(14,103,227,.12); color: var(--brand-primary); }
  .badge-good{ background: rgba(29,73,155,.12); color: var(--brand-primary-dark); }
  .badge-satisfactory{ background: rgba(154,174,208,.18); color: var(--brand-slate); }
  .badge-needs-improvement{ background: rgba(26,49,76,.14); color: var(--brand-slate); }
  .badge-unsatisfactory{ background: rgba(220,38,38,.12); color: var(--danger); }

  /* Info blocks */
  .info-section{ background:var(--brand-bg); border-radius:10px; padding:22px; margin-bottom:26px; }
  .info-section h3{ color:var(--brand-primary-dark); margin-bottom:16px; font-size:18px; }
  .info-grid{ display:grid; grid-template-columns:repeat(2,1fr); gap:14px; }
  .info-item{ padding:8px 0; }
  .info-item.full-width{ grid-column:1 / -1; }

  /* Questions */
  .questions-section{ margin-bottom:26px; }
  .category-header{
    background: var(--brand-primary-dark); color:#fff; padding:14px 18px; margin:18px 0 0 0; border-radius:10px 10px 0 0;
  }
  .questions-table{
    width:100%; border-collapse:collapse; border:1px solid #E2E8F0; border-radius:0 0 10px 10px; overflow:hidden;
  }
  .questions-table th{
    background:#F1F5F9; padding:12px; text-align:left; font-weight:700; font-size:12px; text-transform:uppercase; letter-spacing:.5px; border-bottom:1px solid #E2E8F0;
  }
  .questions-table td{ padding:12px; border-bottom:1px solid #E2E8F0; vertical-align:top; }
  .questions-table tr:hover{ background:#F8FAFC; }

  .answer-chip{
    display:inline-block; padding:4px 10px; border-radius:16px; font-size:11px; font-weight:800; text-transform:uppercase;
  }
  .answer-yes{ background: rgba(14,103,227,.12); color: var(--brand-primary); }
  .answer-no{ background: rgba(220,38,38,.12); color: var(--danger); }
  .answer-na{ background: rgba(154,174,208,.18); color: var(--brand-slate); }

  /* Rich areas */
  .feedback-section, .notes-section{ background:var(--brand-bg); border-radius:10px; padding:22px; margin-bottom:18px; }
  .notes-grid{ display:grid; grid-template-columns:1fr 1fr; gap:18px; margin-bottom:26px; }
  .feedback-content, .notes-content{
    background:#fff; padding:18px; border-radius:8px; border:1px solid #E2E8F0; min-height:100px;
  }

  /* Footer */
  .footer{
    border-top:1px solid #E2E8F0; padding-top:18px; margin-top:34px; text-align:center; font-size:12px; color:#64748B;
  }

  /* Rich content normalizer */
  .rich-content p{ margin-bottom:10px; }
  .rich-content ul, .rich-content ol{ margin:10px 0 10px 20px; }
  .rich-content li{ margin-bottom:5px; }
  .rich-content strong{ font-weight:600; }
  .rich-content em{ font-style:italic; }

  /* Print */
  @media print{
    .container{ padding:0; }
    .metric-card, .questions-section{ break-inside:avoid; }
  }
  `;

  const themeStyles = {
    professional: baseStyles,
    coaching: baseStyles + `
      .coaching-theme{ color:var(--brand-slate); }
      .coaching-header{
        background: linear-gradient(135deg, var(--brand-primary) 0%, var(--brand-primary-dark) 100%);
        color:#fff; padding:28px; border-radius:12px; text-align:center; margin-bottom:26px;
        box-shadow: 0 10px 24px var(--brand-shadow);
      }
      .performance-overview{ text-align:center; margin-bottom:34px; }
      .score-display{
        display:inline-block; background:#fff; padding:28px; border-radius:50%;
        box-shadow:0 8px 22px var(--brand-shadow);
      }
      .score-number{ display:block; font-size:46px; font-weight:900; color:var(--brand-primary); }
      .score-label{ font-size:12px; color:#6B7280; text-transform:uppercase; letter-spacing:1px; }
      .coaching-section{ margin-bottom:26px; padding:22px; border-radius:12px; }
      .coaching-section.strengths{ background:rgba(14,103,227,.08); border-left:4px solid var(--brand-primary); }
      .coaching-section.improvements{ background:rgba(29,73,155,.10); border-left:4px solid var(--brand-primary-dark); }
      .coaching-section.action-plan{ background:rgba(154,174,208,.18); border-left:4px solid var(--brand-muted); }
      .signature-line{ display:flex; justify-content:space-between; margin-top:26px; }
      .signature-field{ flex:1; margin:0 18px; }
    `,
    executive: baseStyles + `
      .executive-theme{ font-size:14px; }
      .executive-header{
        background: var(--brand-navy); color:#fff; padding:24px; border-radius:10px; margin-bottom:26px;
        box-shadow:0 10px 24px var(--brand-shadow);
      }
      .executive-meta{ display:flex; gap:28px; margin-top:12px; font-size:12px; opacity:.95; }
      .kpi-dashboard{ display:grid; grid-template-columns:repeat(3,1fr); gap:18px; margin-bottom:26px; }
      .kpi-card{
        background:#fff; border:1px solid #E5E7EB; border-radius:10px; padding:18px; text-align:center;
        box-shadow:0 6px 14px var(--brand-shadow);
      }
      .kpi-card h4{ margin-bottom:6px; color:var(--brand-slate); }
      .kpi-value{ font-size:22px; font-weight:900; color:var(--brand-primary-dark); }
      .kpi-label{ font-size:12px; color:#6B7280; }
    `
  };

  return themeStyles[theme] || themeStyles.professional;
}


  // ═══════════════════════════════════════════════════════════════════════════
  // CONTENT GENERATION HELPERS
  // ═══════════════════════════════════════════════════════════════════════════

  generatePerformanceBadge(scoreResult) {
    const band = safeGetPerformanceBand_(scoreResult);  // Use safe function
    const badgeClass = `badge-${band.label.toLowerCase().replace(' ', '-')}`;

    return `<div class="performance-badge ${badgeClass}">${band.label}</div>`;
  }

  generateQuestionSections(qaRecord, scoreResult) {
    const categories = qaCategories_();
    const questionText = qaQuestionText_();
    const weights = qaWeights_();

    let html = '<div class="questions-section">';

    Object.entries(categories).forEach(([categoryName, questions]) => {
      html += `
        <div class="category-section">
          <div class="category-header">
            <h3>${this.escapeHtml(categoryName)}</h3>
          </div>
          <table class="questions-table">
            <thead>
              <tr>
                <th style="width: 8%">#</th>
                <th style="width: 50%">Question</th>
                <th style="width: 10%">Weight</th>
                <th style="width: 12%">Answer</th>
                <th style="width: 20%">Notes</th>
              </tr>
            </thead>
            <tbody>
      `;

      questions.forEach(questionKey => {
        const qNum = questionKey.replace(/^q/i, '');
        const answer = qaRecord[`Q${qNum}`] || 'N/A';
        const notes = qaRecord[`Q${qNum} Note`] || qaRecord[`Q${qNum} note`] || qaRecord[`C${qNum}`] || '';
        const weight = weights[questionKey] || 0;
        const questionDesc = questionText[questionKey] || questionKey;

        html += `
          <tr>
            <td><strong>Q${qNum}</strong></td>
            <td>${this.escapeHtml(questionDesc)}</td>
            <td>${weight}</td>
            <td>${this.generateAnswerChip(answer)}</td>
            <td class="rich-content">${this.renderRichContent(notes)}</td>
          </tr>
        `;
      });

      html += `
            </tbody>
          </table>
        </div>
      `;
    });

    html += '</div>';
    return html;
  }

  generateAnswerChip(answer) {
    const normalizedAnswer = String(answer || '').toLowerCase();
    let chipClass, chipText;

    switch (normalizedAnswer) {
      case 'yes':
        chipClass = 'answer-yes';
        chipText = 'YES';
        break;
      case 'no':
        chipClass = 'answer-no';
        chipText = 'NO';
        break;
      default:
        chipClass = 'answer-na';
        chipText = 'N/A';
    }

    return `<span class="answer-chip ${chipClass}">${chipText}</span>`;
  }

  generateChartsSection(scoreResult) {
    // Placeholder for charts - could integrate with Chart.js or similar
    return `
      <div class="charts-section">
        <h3>Performance Analysis</h3>
        <div class="charts-grid">
          <div class="chart-placeholder">
            <p>Score Breakdown Chart</p>
            <p>Final Score: ${scoreResult.finalScore}%</p>
          </div>
        </div>
      </div>
    `;
  }

  generateRecommendations(scoreResult) {
    const recommendations = this.generateRecommendationList(scoreResult);

    return `
      <div class="recommendations-section">
        <h3>Recommendations</h3>
        <ul class="recommendations-list">
          ${recommendations.map(rec => `<li>${this.escapeHtml(rec)}</li>`).join('')}
        </ul>
      </div>
    `;
  }

  generateDataSnapshot(qaRecord) {
    const headers = getQaHeaders_();

    let html = `
      <div class="data-snapshot">
        <h3>Complete Data Record</h3>
        <table class="snapshot-table">
          <thead>
            <tr>
              <th>Field</th>
              <th>Value</th>
            </tr>
          </thead>
          <tbody>
    `;

    headers.forEach(header => {
      const value = qaRecord[header] || '';
      html += `
        <tr>
          <td><strong>${this.escapeHtml(header)}</strong></td>
          <td>${this.escapeHtml(String(value))}</td>
        </tr>
      `;
    });

    html += `
          </tbody>
        </table>
      </div>
    `;

    return html;
  }

  // ═══════════════════════════════════════════════════════════════════════════
  // UTILITY FUNCTIONS
  // ═══════════════════════════════════════════════════════════════════════════

  renderRichContent(content) {
    if (!content || content.trim() === '') {
      return '<p class="empty-content">No content provided</p>';
    }

    // Enhanced rich content rendering with HTML support
    let processed = String(content);

    // Convert newlines to <br> tags
    processed = processed.replace(/\r\n|\r|\n/g, '<br>');

    // Basic safety - escape dangerous tags but keep formatting
    processed = processed.replace(/<script[^>]*>.*?<\/script>/gi, '');
    processed = processed.replace(/<iframe[^>]*>.*?<\/iframe>/gi, '');

    return `<div class="rich-content">${processed}</div>`;
  }

  escapeHtml(text) {
    if (!text) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  generatePdfFilename(qaRecord) {
    const agentName = String(qaRecord.AgentName || 'Agent').replace(/[^a-zA-Z0-9]/g, '_');
    const callDate = String(qaRecord.CallDate || '').replace(/[^0-9-]/g, '');
    const timestamp = new Date().toISOString().split('T')[0];

    return `QA_Report_${agentName}_${callDate}_${timestamp}.pdf`;
  }

  getOrCreateQaFolder(agentName, callDate) {
    try {
      const rootFolder = ensureRootFolder_();
      const agentFolder = getOrCreateFolder_(rootFolder, sanitizeName_(agentName || 'Unknown'));
      const dateFolder = getOrCreateFolder_(agentFolder, callDate || new Date().toISOString().split('T')[0]);
      return dateFolder;
    } catch (error) {
      console.error('Error creating QA folder:', error);
      // Fallback to root folder
      return DriveApp.getRootFolder();
    }
  }

  generateRecommendationList(scoreResult) {
    const recommendations = [];

    if (scoreResult.finalScore < 70) {
      recommendations.push('Schedule immediate coaching session to address performance gaps');
      recommendations.push('Review call handling procedures and best practices');
    }

    if (scoreResult.penalties && scoreResult.penalties.length > 0) {
      recommendations.push('Focus on compliance areas that resulted in penalties');
    }

    if (scoreResult.finalScore >= 90) {
      recommendations.push('Excellent performance - consider for mentoring opportunities');
    }

    if (recommendations.length === 0) {
      recommendations.push('Continue current performance standards');
      recommendations.push('Regular follow-up monitoring recommended');
    }

    return recommendations;
  }

  // Additional helper methods for different templates would go here...
  generateActionPlan(scoreResult) {
    // Implementation for coaching template
    return '<p>Customized action plan based on performance analysis</p>';
  }

  identifyStrengths(scoreResult) {
    // Implementation for coaching template
    return '<p>Areas where the agent excelled</p>';
  }

  identifyImprovements(scoreResult) {
    // Implementation for coaching template  
    return '<p>Specific areas for development</p>';
  }

  calculateNextReviewDate() {
    const nextReview = new Date();
    nextReview.setDate(nextReview.getDate() + 30);
    return nextReview.toLocaleDateString();
  }

  getPriorityAreas(scoreResult) {
    return 'Communication, Documentation';
  }

  /**
 * Missing Methods for QAPdfService Class
 * Add these methods to your QAPdfService class in QAPdfService.gs
 */

// Add these methods inside the QAPdfService class:

/**
 * Generate simple template HTML
 */
generateSimpleTemplate(qaRecord, scoreResult, config) {
  const theme = this.getThemeStyles('professional');
  
  return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>QA Summary - ${this.escapeHtml(qaRecord.AgentName)}</title>
    <style>${theme}</style>
</head>
<body>
    <div class="container simple-theme">
        <!-- Simple Header -->
        <div class="header">
          <div class="brand-lockup">
            <div class="brand-title">
              <h1>Quality Assessment Summary</h1>
              <p>high visibility command center</p>
            </div>
          </div>
          <div class="brand-meta">
            <p><strong>Agent:</strong> ${this.escapeHtml(qaRecord.AgentName)}</p>
            <p><strong>Date:</strong> ${qaRecord.CallDate}</p>
            <p><strong>Score:</strong> ${scoreResult.finalScore}%</p>
          </div>
        </div>

        <!-- Key Results -->
        <div class="simple-results">
            <div class="result-item">
                <span class="label">Final Score:</span>
                <span class="value score-${scoreResult.finalScore >= 85 ? 'good' : scoreResult.finalScore >= 70 ? 'fair' : 'poor'}">${scoreResult.finalScore}%</span>
            </div>
            <div class="result-item">
                <span class="label">Status:</span>
                <span class="value ${scoreResult.isPassing ? 'pass' : 'fail'}">${scoreResult.isPassing ? 'PASS' : 'FAIL'}</span>
            </div>
            <div class="result-item">
                <span class="label">Points:</span>
                <span class="value">${scoreResult.earned} / ${scoreResult.applicable}</span>
            </div>
        </div>

        <!-- Key Feedback -->
        ${qaRecord.OverallFeedback ? `
        <div class="simple-feedback">
            <h3>Key Feedback</h3>
            ${this.renderRichContent(qaRecord.OverallFeedback)}
        </div>
        ` : ''}

        <!-- Footer -->
        <div class="simple-footer">
            <p>Audited by: ${this.escapeHtml(qaRecord.AuditorName)} | Generated: ${new Date().toLocaleDateString()}</p>
        </div>
    </div>

    <style>
        .simple-theme { font-family: Arial, sans-serif; }
        .simple-header { 
            text-align: center; 
            padding: 20px; 
            border-bottom: 2px solid #333; 
            margin-bottom: 30px; 
        }
        .simple-results { 
            display: flex; 
            justify-content: space-around; 
            margin-bottom: 30px; 
            padding: 20px; 
            background: #f5f5f5; 
        }
        .result-item { text-align: center; }
        .result-item .label { display: block; font-weight: bold; margin-bottom: 5px; }
        .result-item .value { font-size: 18px; font-weight: bold; }
        .score-good { color: #059669; }
        .score-fair { color: #f59e0b; }
        .score-poor { color: #dc2626; }
        .pass { color: #059669; }
        .fail { color: #dc2626; }
        .simple-feedback { 
            padding: 20px; 
            background: white; 
            border: 1px solid #ddd; 
            margin-bottom: 20px; 
        }
        .simple-footer { 
            text-align: center; 
            padding: 20px; 
            border-top: 1px solid #ccc; 
            color: #656; 
            font-size: 12px; 
        }
    </style>
</body>
</html>`;
}

/**
 * Generate KPI summary for executive template
 */
generateKpiSummary(scoreResult) {
  const performanceBand = safeGetPerformanceBand_(scoreResult);
  
  return `
    <div class="kpi-card">
        <h4>Overall Score</h4>
        <div class="kpi-value">${scoreResult.finalScore}%</div>
        <div class="kpi-label">${performanceBand.label}</div>
    </div>
    <div class="kpi-card">
        <h4>Points Earned</h4>
        <div class="kpi-value">${scoreResult.earned}</div>
        <div class="kpi-label">out of ${scoreResult.applicable}</div>
    </div>
    <div class="kpi-card">
        <h4>Pass Status</h4>
        <div class="kpi-value ${scoreResult.isPassing ? 'pass' : 'fail'}">
            ${scoreResult.isPassing ? 'PASS' : 'FAIL'}
        </div>
        <div class="kpi-label">${scoreResult.isPassing ? 'Meets Standards' : 'Below Standards'}</div>
    </div>
  `;
}

/**
 * Generate key insights for executive template
 */
generateKeyInsights(scoreResult) {
  const insights = [];
  
  if (scoreResult.finalScore >= 95) {
    insights.push({
      type: 'success',
      title: 'Exceptional Performance',
      description: 'Agent consistently exceeds quality standards'
    });
  } else if (scoreResult.finalScore >= 85) {
    insights.push({
      type: 'success',
      title: 'Strong Performance',
      description: 'Agent meets or exceeds most quality criteria'
    });
  } else if (scoreResult.finalScore >= 70) {
    insights.push({
      type: 'warning',
      title: 'Moderate Performance',
      description: 'Agent meets basic standards with room for improvement'
    });
  } else {
    insights.push({
      type: 'critical',
      title: 'Performance Gap',
      description: 'Immediate intervention required to meet standards'
    });
  }

  // Check for penalties
  if (scoreResult.penalties && scoreResult.penalties.length > 0) {
    scoreResult.penalties.forEach(penalty => {
      insights.push({
        type: 'critical',
        title: 'Compliance Issue',
        description: penalty.description
      });
    });
  }

  return insights.map(insight => `
    <div class="insight-card insight-${insight.type}">
        <h4>${insight.title}</h4>
        <p>${insight.description}</p>
    </div>
  `).join('');
}

/**
 * Generate risk assessment for executive template
 */
generateRiskAssessment(scoreResult) {
  let riskLevel = 'Low';
  let riskDescription = 'Performance within acceptable parameters';
  let riskColor = '#059669';

  if (scoreResult.finalScore < 60) {
    riskLevel = 'High';
    riskDescription = 'Significant performance deficiencies identified';
    riskColor = '#dc2626';
  } else if (scoreResult.finalScore < 80) {
    riskLevel = 'Medium';
    riskDescription = 'Some areas require attention and improvement';
    riskColor = '#f59e0b';
  }

  // Check for critical failures
  if (scoreResult.penalties && scoreResult.penalties.some(p => p.type === 'AUTO_FAIL')) {
    riskLevel = 'Critical';
    riskDescription = 'Critical failures detected - immediate action required';
    riskColor = '#991b1b';
  }

  return `
    <div class="risk-assessment" style="border-left: 4px solid ${riskColor};">
        <div class="risk-header">
            <span class="risk-level" style="color: ${riskColor};">${riskLevel} Risk</span>
            <span class="risk-score">${scoreResult.finalScore}%</span>
        </div>
        <p class="risk-description">${riskDescription}</p>
        ${riskLevel !== 'Low' ? `
        <div class="risk-actions">
            <h5>Recommended Actions:</h5>
            <ul>
                ${riskLevel === 'Critical' ? '<li>Immediate supervisor intervention</li>' : ''}
                ${riskLevel === 'High' ? '<li>Formal coaching program enrollment</li>' : ''}
                ${riskLevel === 'Medium' ? '<li>Additional training and monitoring</li>' : ''}
                <li>Follow-up assessment in 30 days</li>
            </ul>
        </div>
        ` : ''}
    </div>
  `;
}

/**
 * Generate executive recommendations
 */
generateExecutiveRecommendations(scoreResult) {
  const recommendations = [];

  if (scoreResult.finalScore >= 90) {
    recommendations.push('Consider agent for advanced training or mentoring roles');
    recommendations.push('Recognize exceptional performance in team meetings');
  } else if (scoreResult.finalScore >= 80) {
    recommendations.push('Continue current performance standards');
    recommendations.push('Consider for specialized project assignments');
  } else if (scoreResult.finalScore >= 70) {
    recommendations.push('Implement targeted skill development program');
    recommendations.push('Increase monitoring frequency to monthly reviews');
  } else {
    recommendations.push('Immediate placement in performance improvement program');
    recommendations.push('Daily monitoring and coaching sessions');
    recommendations.push('Consider additional training resources');
  }

  // Add penalty-specific recommendations
  if (scoreResult.penalties && scoreResult.penalties.length > 0) {
    recommendations.push('Address compliance gaps identified in assessment');
    recommendations.push('Review and reinforce procedural training');
  }

  return `
    <div class="exec-recommendations-list">
        ${recommendations.map((rec, index) => `
        <div class="recommendation-item">
            <span class="rec-number">${index + 1}</span>
            <span class="rec-text">${rec}</span>
        </div>
        `).join('')}
    </div>
  `;
}

/**
 * Generate detailed template HTML
 */
generateDetailedTemplate(qaRecord, scoreResult, config) {
  const theme = this.getThemeStyles(config.theme);
  const performanceBadge = this.generatePerformanceBadge(scoreResult);
  const questionSections = this.generateQuestionSections(qaRecord, scoreResult);
  const detailedAnalysis = this.generateDetailedAnalysis(scoreResult);

  return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Detailed QA Report - ${this.escapeHtml(qaRecord.AgentName)}</title>
    <style>${theme}</style>
</head>
<body>
    <div class="container detailed-theme">
        <!-- Detailed Header -->
        <div class="header">
          <div class="brand-lockup">
            <div class="brand-title">
              <h1>Comprehensive Quality Assessment</h1>
              <p>high visibility command center</p>
            </div>
          </div>
          <div class="brand-meta">
            <p><strong>Call:</strong> ${qaRecord.CallDate}</p>
            <p><strong>Audit:</strong> ${qaRecord.AuditDate}</p>
            <p><strong>Auditor:</strong> ${this.escapeHtml(qaRecord.AuditorName)}</p>
          </div>
        </div>


        <!-- Executive Summary -->
        <div class="executive-summary">
            <h3>Executive Summary</h3>
            <div class="summary-grid">
                <div class="summary-item">
                    <strong>Final Score:</strong> ${scoreResult.finalScore}%
                    ${performanceBadge}
                </div>
                <div class="summary-item">
                    <strong>Performance:</strong> ${scoreResult.isPassing ? 'MEETS STANDARDS' : 'BELOW STANDARDS'}
                </div>
            </div>
        </div>

        <!-- Detailed Analysis -->
        ${detailedAnalysis}

        <!-- Complete Question Breakdown -->
        <div class="detailed-questions">
            <h3>Complete Assessment Breakdown</h3>
            ${questionSections}
        </div>

        <!-- Comprehensive Feedback -->
        <div class="comprehensive-feedback">
            <h3>Comprehensive Feedback Analysis</h3>
            ${this.renderRichContent(qaRecord.OverallFeedback)}
        </div>

        <!-- Action Items -->
        <div class="action-items">
            <h3>Recommended Action Items</h3>
            ${this.generateActionItems(scoreResult)}
        </div>
    </div>
</body>
</html>`;
}

/**
 * Generate detailed analysis section
 */
generateDetailedAnalysis(scoreResult) {
  return `
    <div class="detailed-analysis">
        <h3>Performance Analysis</h3>
        <div class="analysis-grid">
            <div class="analysis-section">
                <h4>Score Breakdown</h4>
                <p>Points Earned: ${scoreResult.earned} out of ${scoreResult.applicable}</p>
                <p>Success Rate: ${Math.round((scoreResult.earned / scoreResult.applicable) * 100)}%</p>
            </div>
            ${scoreResult.penalties && scoreResult.penalties.length > 0 ? `
            <div class="analysis-section penalties">
                <h4>Penalties Applied</h4>
                ${scoreResult.penalties.map(penalty => `
                <p><strong>${penalty.type}:</strong> ${penalty.description}</p>
                `).join('')}
            </div>
            ` : ''}
        </div>
    </div>
  `;
}

/**
 * Generate action items for detailed template
 */
generateActionItems(scoreResult) {
  const items = this.generateRecommendationList(scoreResult);
  
  return `
    <ul class="action-items-list">
      ${items.map((item, index) => `
      <li class="action-item">
        <span class="item-number">${index + 1}</span>
        <span class="item-text">${this.escapeHtml(item)}</span>
      </li>
      `).join('')}
    </ul>
  `;
}

}

// ═══════════════════════════════════════════════════════════════════════════════
// GLOBAL INSTANCE AND FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

// Create global instance
const qaPdfService = new QAPdfService();

// Export functions for use in main QA service
function generateQaPdfReport(qaRecord, scoreResult, options = {}) {
  return qaPdfService.generateQaPdf(qaRecord, scoreResult, options);
}

// Also update the generateQaPdfById function if it exists:
function generateQaPdfById(qaRecordId, options = {}) {
  try {
    const qaRecord = getQARecordById(qaRecordId);
    if (!qaRecord) {
      return {
        success: false,
        error: 'QA record not found'
      };
    }

    // Extract answers and compute score using enhanced function
    const answers = {};
    const weights = qaWeights_();
    Object.keys(weights).forEach(k => {
      const qNum = k.replace(/^q/i, '');
      answers[k] = qaRecord[`Q${qNum}`];
    });

    const scoreResult = computeEnhancedQaScore(answers);  // <-- Use enhanced function

    if (scoreResult.error) {
      return {
        success: false,
        error: 'Score calculation failed: ' + scoreResult.message
      };
    }

    return generateQaPdfReport(qaRecord, scoreResult, options);

  } catch (error) {
    console.error('Error generating PDF by ID:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
