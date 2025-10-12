/**
 * EmailService.gs - Enhanced Email Service for User Management
 * Handles password setup emails, password resets, and other notifications.
 * Version: 2025-08-26
 *
 * What’s new:
 * - Shared renderer for consistent header/footer + preheader
 * - URL-safe tokens (encodeURIComponent)
 * - Centralized sendEmail_ with replyTo + optional audit BCC
 * - Graceful writeError fallback (sheet or logs) — no external dependency required
 * - Lightweight HTML escaping for user-provided data
 * - Plain-text fallbacks automatically generated when not provided
 * - Safer defaults (strict subject prefixes, brand/org names)
 * - NEW: Auto-label sent mail under Gmail label "Lumina" (+ subcategories)
 */

// ────────────────────────────────────────────────────────────────────────────
// Email Configuration (adjust as needed)
// ────────────────────────────────────────────────────────────────────────────
const EMAIL_CONFIG = {
  brandName: 'Lumina HQ',
  orgName: 'VLBPO',
  fromName: 'Lumina HQ',
  fromEmail: 'lumina@vlbpo.com',
  supportEmail: 'it@vlbpo.com',
  baseUrl: 'https://script.google.com/a/macros/vlbpo.com/s/AKfycbxeQ0AnupBHM71M6co3LVc5NPrxTblRXLd6AuTOpxMs2rMehF9dBSkGykIcLGHROywQ/exec',
  logoUrl: 'https://res.cloudinary.com/dr8qd3xfc/image/upload/v1754763514/vlbpo/lumina/2_eb1h4a.png',
  datalogLogoUrl: 'https://res.cloudinary.com/dr8qd3xfc/image/upload/v1754763514/vlbpo/lumina/2_eb1h4a.png',

  // Optional auditing for compliance — set to '' to disable
  auditBcc: '', // e.g., 'security-audit@vlbpo.com'
  subjectPrefix: '[Lumina HQ] ',
  dryRun: false // when true, logs the email instead of sending
};

// ────────────────────────────────────────────────────────────────────────────
// NEW: Gmail label configuration (for your Sent mailbox)
// ────────────────────────────────────────────────────────────────────────────
const EMAIL_LABEL_ROOT = 'Lumina';
const EMAIL_LABEL_DEFAULT_CATEGORY = 'Notifications';

// ────────────────────────────────────────────────────────────────────────────
const ENHANCED_STYLES = `
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    body {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
        line-height: 1.6;
        color: #1e293b;
        margin: 0;
        padding: 0;
        background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%);
        min-height: 100vh
    }

    .email-container {
        max-width: 650px;
        margin: 40px auto;
        padding: 0;
        background: #fff;
        border-radius: 16px;
        box-shadow: 0 20px 40px rgba(0, 0, 0, .08), 0 8px 16px rgba(0, 0, 0, .04);
        overflow: hidden;
        border: 1px solid rgba(255, 255, 255, .2)
    }

    .header {
        background: linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%);
        color: #fff;
        padding: 40px 30px;
        text-align: center;
        position: relative;
        overflow: hidden
    }

    .header::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle, rgba(255, 255, 255, .15) 0%, transparent 70%);
        transform: rotate(45deg);
        pointer-events: none
    }

    .logo {
        max-width: 220px;
        height: auto;
        margin-bottom: 20px;
        filter: brightness(1.1) contrast(1.1);
        position: relative;
        z-index: 2
    }

    .header-title {
        margin: 0;
        font-size: 28px;
        font-weight: 600;
        letter-spacing: -.5px;
        position: relative;
        z-index: 2
    }

    .content {
        padding: 50px 40px;
        background: #fff
    }

    .welcome-badge, .status-badge {
        background: linear-gradient(135deg, #10b981, #059669);
        color: #fff;
        padding: 12px 24px;
        border-radius: 25px;
        font-size: 14px;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 30px;
        box-shadow: 0 4px 12px rgba(16, 185, 129, .3);
        border: 2px solid rgba(255, 255, 255, .2)
    }

    .security-badge {
        background: linear-gradient(135deg, #f59e0b, #d97706);
        box-shadow: 0 4px 12px rgba(245, 158, 11, .3)
    }

    .success-badge {
        background: linear-gradient(135deg, #10b981, #059669);
        box-shadow: 0 4px 12px rgba(16, 185, 129, .3)
    }

    .cta-wrap {
        margin: 40px 0;
        text-align: center
    }

    .cta-button {
        display: inline-block;
        padding: 18px 40px;
        background: linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%);
        color: #fff;
        text-decoration: none;
        border-radius: 12px;
        font-weight: 600;
        font-size: 16px;
        transition: all .3s ease;
        box-shadow: 0 8px 20px rgba(14, 165, 233, .3);
        border: 2px solid rgba(255, 255, 255, .1);
        position: relative;
        overflow: hidden
    }

    .cta-button::before {
        content: '';
        position: absolute;
        top: 0;
        left: -100%;
        width: 100%;
        height: 100%;
        background: linear-gradient(90deg, transparent, rgba(255, 255, 255, .2), transparent);
        transition: left .5s
    }

    .cta-button:hover {
        transform: translateY(-2px);
        box-shadow: 0 12px 30px rgba(14, 165, 233, .4)
    }

    .cta-button:hover::before {
        left: 100%
    }

    .reset-button {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        box-shadow: 0 8px 20px rgba(245, 158, 11, .3)
    }

    .reset-button:hover {
        box-shadow: 0 12px 30px rgba(245, 158, 11, .4)
    }

    .footer {
        background: linear-gradient(135deg, #f8fafc, #e2e8f0);
        padding: 30px;
        text-align: center;
        font-size: 14px;
        color: #64748b;
        border-top: 1px solid #e2e8f0
    }

    .divider {
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent, #e2e8f0, transparent);
        margin: 40px 0
    }

    .info-card {
        background: linear-gradient(135deg, #f0f9ff, #e0f2fe);
        border-left: 5px solid #0ea5e9;
        padding: 25px;
        margin: 30px 0;
        border-radius: 12px;
        position: relative;
        box-shadow: 0 4px 12px rgba(14, 165, 233, .1)
    }

    .warning-card {
        background: linear-gradient(135deg, #fffbeb, #fef3c7);
        border-left: 5px solid #f59e0b;
        box-shadow: 0 4px 12px rgba(245, 158, 11, .1);
        padding: 25px;
        border-radius: 12px;
        margin: 30px 0
    }

    .success-card {
        background: linear-gradient(135deg, #ecfdf5, #d1fae5);
        border-left: 5px solid #10b981;
        box-shadow: 0 4px 12px rgba(16, 185, 129, .1);
        padding: 25px;
        border-radius: 12px;
        margin: 30px 0
    }

    .icon-large {
        font-size: 64px;
        text-align: center;
        margin: 30px 0;
        filter: drop-shadow(0 4px 8px rgba(0, 0, 0, .1))
    }

    .security-icon {
        background: linear-gradient(135deg, #f59e0b, #d97706);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text
    }

    .success-icon {
        background: linear-gradient(135deg, #10b981, #059669);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text
    }

    .link-display {
        word-break: break-all;
        color: #0ea5e9;
        background: linear-gradient(135deg, #f8fafc, #f1f5f9);
        padding: 20px;
        border-radius: 12px;
        font-family: 'Monaco', 'Menlo', 'Ubuntu Mono', monospace;
        font-size: 14px;
        border: 2px solid #e2e8f0;
        position: relative;
        overflow: hidden
    }

    .feature-list {
        list-style: none;
        padding: 0;
        margin: 20px 0
    }

    .feature-list li {
        background: #f8fafc;
        padding: 15px 20px;
        margin: 8px 0;
        border-radius: 8px;
        border-left: 4px solid #0ea5e9;
        transition: all .3s ease;
        position: relative
    }

    .feature-list li::before {
        content: '✓';
        color: #10b981;
        font-weight: 700;
        margin-right: 10px;
        font-size: 16px
    }

    .account-details {
        background: linear-gradient(135deg, #fff, #f8fafc);
        border: 2px solid #e2e8f0;
        border-radius: 12px;
        padding: 25px;
        margin: 25px 0;
        box-shadow: 0 4px 12px rgba(0, 0, 0, .05)
    }

    .account-details strong {
        color: #334155;
        display: inline-block;
        min-width: 120px
    }

    .futuristic-card {
        background: linear-gradient(135deg, rgba(0, 63, 135, .06), rgba(0, 174, 239, .08));
        border: 1px solid rgba(14, 165, 233, .18);
        border-radius: 18px;
        padding: 28px;
        box-shadow: 0 18px 40px rgba(15, 23, 42, .08);
    }

    .detail-grid {
        display: flex;
        flex-wrap: wrap;
        gap: 18px;
    }

    .detail-item {
        flex: 1 1 220px;
        background: rgba(255, 255, 255, .9);
        border: 1px solid rgba(148, 163, 184, .25);
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 10px 25px rgba(15, 23, 42, .08);
    }

    .muted-label {
        display: block;
        font-size: 11px;
        letter-spacing: .12em;
        text-transform: uppercase;
        color: #64748b;
        margin-bottom: 6px;
        font-weight: 700;
    }

    .role-chip, .page-chip {
        display: inline-block;
        margin: 6px 8px 0 0;
        padding: 8px 16px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        letter-spacing: .01em;
    }

    .role-chip {
        background: linear-gradient(135deg, #003f87 0%, #00aeef 100%);
        color: #fff;
        box-shadow: 0 10px 22px rgba(0, 63, 135, .28);
    }

    .role-chip--ghost {
        background: rgba(148, 163, 184, .15);
        color: #475569;
        border: 1px dashed rgba(148, 163, 184, .7);
        box-shadow: none;
    }

    .page-chip {
        background: rgba(14, 165, 233, .15);
        color: #0f172a;
        border: 1px solid rgba(14, 165, 233, .35);
        box-shadow: none;
    }

    .page-chip--ghost {
        background: rgba(226, 232, 240, .6);
        color: #475569;
        border: 1px dashed rgba(148, 163, 184, .6);
    }

    .status-indicator {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        padding: 8px 14px;
        border-radius: 999px;
        font-size: 12px;
        font-weight: 600;
        background: rgba(16, 185, 129, .14);
        color: #0f766e;
        box-shadow: 0 8px 18px rgba(16, 185, 129, .18);
    }

    .status-indicator--warning {
        background: rgba(248, 113, 113, .18);
        color: #b91c1c;
        box-shadow: 0 8px 18px rgba(248, 113, 113, .2);
    }

    .status-indicator--info {
        background: rgba(59, 130, 246, .18);
        color: #1d4ed8;
    }

    .status-indicator--muted {
        background: rgba(148, 163, 184, .18);
        color: #475569;
        box-shadow: none;
    }

    .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: #10b981;
        box-shadow: 0 0 0 4px rgba(16, 185, 129, .25);
        display: inline-block;
    }

    .status-indicator--warning .status-dot {
        background: #ef4444;
        box-shadow: 0 0 0 4px rgba(248, 113, 113, .25);
    }

    .status-indicator--info .status-dot {
        background: #38bdf8;
        box-shadow: 0 0 0 4px rgba(56, 189, 248, .25);
    }

    .status-indicator--muted .status-dot {
        background: #94a3b8;
        box-shadow: none;
    }

    .info-matrix {
        margin-top: 28px;
        border: 1px solid rgba(14, 165, 233, .2);
        border-radius: 16px;
        overflow: hidden;
    }

    .info-matrix-row {
        display: flex;
        flex-wrap: wrap;
        border-bottom: 1px solid rgba(226, 232, 240, .6);
    }

    .info-matrix-row:last-child {
        border-bottom: none;
    }

    .info-matrix-cell {
        flex: 1 1 260px;
        padding: 18px 22px;
        background: rgba(248, 250, 252, .9);
        border-right: 1px solid rgba(226, 232, 240, .6);
    }

    .info-matrix-cell:nth-child(2n) {
        background: rgba(255, 255, 255, .95);
    }

    .info-matrix-cell:last-child {
        border-right: none;
    }

    .subtitle {
        color: #64748b;
        font-size: 16px;
        margin-bottom: 30px;
        font-weight: 400;
        line-height: 1.5
    }

    .emphasis {
        color: #334155;
        font-weight: 600
    }

    .danger-text {
        color: #dc2626;
        font-weight: 600
    }

    .support-link {
        color: #0ea5e9;
        text-decoration: none;
        font-weight: 600;
        border-bottom: 2px solid transparent;
        transition: border-color .3s ease
    }

    .support-link:hover {
        border-bottom-color: #0ea5e9
    }

    .company-footer {
        font-weight: 600;
        color: #334155;
        margin-bottom: 10px
    }

    .disclaimer {
        color: #94a3b8;
        font-size: 12px;
        font-style: italic
    }

    .employee-welcome {
        background: linear-gradient(135deg, #0ea5e9, #06b6d4);
        color: #fff;
        padding: 20px;
        border-radius: 12px;
        margin: 30px 0;
        text-align: center
    }

    .workplace-notice {
        background: linear-gradient(135deg, #f0fdfa, #ccfbf1);
        border-left: 5px solid #14b8a6;
        padding: 20px;
        border-radius: 8px;
        margin: 20px 0
    }

    .security-emphasis {
        background: linear-gradient(135deg, #fef7ff, #fae8ff);
        border: 2px solid #d8b4fe;
        border-radius: 12px;
        padding: 20px;
        margin: 25px 0
    }

    .preheader {
        display: none !important;
        visibility: hidden;
        opacity: 0;
        color: transparent;
        height: 0;
        width: 0;
        overflow: hidden;
        mso-hide: all
    }
</style>
`;

// ────────────────────────────────────────────────────────────────────────────
// Utilities
// ────────────────────────────────────────────────────────────────────────────
/** @param {any} v */
function _isNil(v) { return v === null || v === undefined; }
/** @param {string} s */
function escapeHtml_(s) {
  if (_isNil(s)) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
/** naive html->text for fallback */
function htmlToText_(html) {
  if (_isNil(html)) return '';
  return String(html)
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<head[\s\S]*?<\/head>/gi, '')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}
/** log + optional sheet persist */
function writeError(where, err) {
  try {
    const msg = (err && err.stack) ? err.stack : (err && err.message) ? err.message : String(err);
    console.error(`[EmailService] ${where}: ${msg}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet && SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss ? (ss.getSheetByName('SystemLogs') || ss.insertSheet('SystemLogs')) : null;
    if (sheet) {
      sheet.appendRow([new Date(), 'EmailService', where, Session.getActiveUser().getEmail ? Session.getActiveUser().getEmail() : '', msg]);
    }
  } catch (e) {
    console.error('[EmailService] writeError fallback: ' + e);
  }
}

// ────────────────────────────────────────────────────────────────────────────
// NEW: Labeling helpers (Sent mailbox labeling)
// ────────────────────────────────────────────────────────────────────────────
function ensureLabel_(name) {
  try {
    return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
  } catch (e) {
    writeError('ensureLabel_', e);
    return null;
  }
}

function injectTrackingUid_(html, uid) {
  try {
    const tag = `
<div style="display:none;opacity:0;visibility:hidden;height:0;width:0">LuminaUID:${uid}</div>`;
    if (!html) return tag;
    const replaced = html.replace(/<\/body>\s*<\/html>\s*$/i, `${tag}</body></html>`);
    return replaced === html ? (html + tag) : replaced;
  } catch (e) {
    writeError('injectTrackingUid_', e);
    return html || '';
  }
}

function labelRecentlySent_(to, subject, uid, category) {
  try {
    const root = ensureLabel_(EMAIL_LABEL_ROOT);
    const cat = ensureLabel_(`${EMAIL_LABEL_ROOT}/${category || EMAIL_LABEL_DEFAULT_CATEGORY}`);
    // Quote-safe subject for search
    const safeSubject = String(subject || '').replace(/"/g, '\\"');
    const query = [
      'in:sent',
      `to:"${to}"`,
      `subject:"${safeSubject}"`,
      'newer_than:7d'
    ].join(' ');
    const threads = GmailApp.search(query, 0, 10);
    for (var i = 0; i < threads.length; i++) {
      const msgs = threads[i].getMessages();
      for (var j = msgs.length - 1; j >= 0; j--) {
        const m = msgs[j];
        let body = '', pbody = '';
        try {
          body = m.getBody() || '';
        } catch (e) {
        }
        try {
          pbody = m.getPlainBody() || '';
        } catch (e) {
        }
        if (body.indexOf(uid) !== -1 || pbody.indexOf(uid) !== -1) {
          if (root) threads[i].addLabel(root);
          if (cat) threads[i].addLabel(cat);
          return true;
        }
      }
    }
  } catch (e) {
    writeError('labelRecentlySent_', e);
  }
  return false;
}

/** centralized sender with options */
function sendEmail_(opts) {
  const { to, subject, htmlBody, textBody, replyTo, bcc, category } = opts;

  // DRY RUN path
  if (EMAIL_CONFIG.dryRun) {
    const uidPreview = Utilities.getUuid();
    console.log('[EmailService] DRY RUN - would send:', {
      to,
      subject,
      replyTo: replyTo || EMAIL_CONFIG.supportEmail,
      bcc,
      category,
      uidPreview
    });
    return;
  }

  // Generate UID and inject tracking for labeling
  const uid = Utilities.getUuid();
  const htmlWithTag = injectTrackingUid_(htmlBody || '', uid);
  const textWithTag = (textBody || htmlToText_(htmlBody)).concat(`\n\n[LuminaUID:${uid}]`);

  const payload = {
    to: to,
    subject: subject,
    htmlBody: htmlWithTag,
    body: textWithTag,
    name: EMAIL_CONFIG.fromName,
    replyTo: replyTo || EMAIL_CONFIG.supportEmail
  };
  if (EMAIL_CONFIG.auditBcc || bcc) {
    payload.bcc = [EMAIL_CONFIG.auditBcc, bcc].filter(Boolean).join(',');
  }

  MailApp.sendEmail(payload);

  // Attempt to label the just-sent thread in Sent
  try {
    labelRecentlySent_(to, subject, uid, category);
  } catch (e) {
    writeError('sendEmail_ labeling', e);
  }
}

/** shared shell render */
function renderEmail_({ headerTitle, headerGradient, logoUrl, preheader, contentHtml }) {
  const year = new Date().getFullYear();
  const safePreheader = escapeHtml_(preheader || '');
  const hdrGrad = headerGradient || 'linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%)';
  const logo = logoUrl || EMAIL_CONFIG.logoUrl;

  return `
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    ${ENHANCED_STYLES}
</head>
<body>
<span class="preheader">${safePreheader}</span>
<div class="email-container">
    <div class="header" style="background:${hdrGrad};">
        <img src="${logo}" alt="${escapeHtml_(EMAIL_CONFIG.orgName)} Logo" class="logo" loading="lazy">
        <h1 class="header-title">${escapeHtml_(headerTitle || EMAIL_CONFIG.brandName)}</h1>
    </div>
    <div class="content">
        ${contentHtml || ''}
    </div>
    <div class="footer">
        <p class="company-footer">© ${year} ${escapeHtml_(EMAIL_CONFIG.orgName)} —
            ${escapeHtml_(EMAIL_CONFIG.brandName)}. All rights reserved.</p>
        <p class="disclaimer">System notification. Please do not reply. For help contact <a class="support-link"
                                                                                            href="mailto:${EMAIL_CONFIG.supportEmail}">${EMAIL_CONFIG.supportEmail}</a>.
        </p>
    </div>
</div>
</body>
</html>`;
}

/** make absolute, tokenized link safely */
function buildPasswordUrl_(token) {
  const t = encodeURIComponent(String(token || ''));
  const sep = EMAIL_CONFIG.baseUrl.indexOf('?') >= 0 ? '&' : '?';
  return `${EMAIL_CONFIG.baseUrl}${sep}page=setpassword&token=${t}&utm_source=email&utm_medium=auth&utm_campaign=${encodeURIComponent(EMAIL_CONFIG.brandName)}`;
}

// ────────────────────────────────────────────────────────────────────────────
// Main Email Functions
// ────────────────────────────────────────────────────────────────────────────

/**
 * Send password setup email for new users (NO EMAIL VERIFICATION)
 * @param {string} email - Recipient email
 * @param {Object} data - { userName, fullName, passwordSetupToken }
 * @return {boolean}
 */
function sendPasswordSetupEmail(email, data) {
  try {
    const safeName = escapeHtml_(data.fullName || data.userName || 'there');
    const safeUser = escapeHtml_(data.userName || '');
    const setupUrl = buildPasswordUrl_(data.passwordSetupToken);

    const content = `
<div class="welcome-badge">🎉 Employee Account Created</div>
<p class="subtitle">Hi <span class="emphasis">${safeName}</span>,</p>
<p>Welcome to ${escapeHtml_(EMAIL_CONFIG.brandName)}! Your employee account has been created and is ready for setup.
    Create your secure password using the button below:</p>

<div class="cta-wrap">
    <a href="${setupUrl}" class="cta-button">Set Up Your Account</a>
</div>

<div class="account-details">
    <h3 style="margin-top:0;color:#334155;">👤 Your Account Details</h3>
    <p><strong>Employee Username:</strong> ${safeUser}</p>
    <p><strong>Email Address:</strong> ${escapeHtml_(email)}</p>
    <p><strong>Access Level:</strong> <span style="color:#10b981;font-weight:600;">Employee Portal ✅</span></p>
</div>

<div class="info-card">
    <p><strong>🔗 Manual Setup Link:</strong></p>
    <p>If the button above doesn't work, copy and paste this secure link into your browser:</p>
    <div class="link-display">${escapeHtml_(setupUrl)}</div>
</div>

<hr class="divider">

<div class="workplace-notice">
    <h3 style="color:#14b8a6;margin-top:0;">🏢 Next Steps</h3>
    <ul class="feature-list">
        <li style="border-left-color:#14b8a6;">Create your secure password</li>
        <li style="border-left-color:#14b8a6;">Access your employee dashboard</li>
        <li style="border-left-color:#14b8a6;">Explore workplace tools and resources</li>
        <li style="border-left-color:#14b8a6;">Set your preferences</li>
        <li style="border-left-color:#14b8a6;">Contact IT for any assistance</li>
    </ul>
</div>
`;

    const htmlBody = renderEmail_({
      headerTitle: 'Welcome to the Team',
      headerGradient: 'linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'Your employee account is ready — set your password to get started.',
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + 'Complete Your Account Setup';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Onboarding'
    });

    console.log(`Password setup email sent to ${email}`);
    return true;
  } catch (error) {
    writeError('sendPasswordSetupEmail', error);
    return false;
  }
}

/**
 * Send password reset email
 * @param {string} email
 * @param {string} resetToken
 * @return {boolean}
 */
function sendPasswordResetEmail(email, resetToken) {
  try {
    const resetUrl = buildPasswordUrl_(resetToken);

    const content = `
<div class="security-badge">🔐 Security Request</div>
<div class="icon-large"><span class="security-icon">🛡️</span></div>

<p class="subtitle">We received a request to reset your ${escapeHtml_(EMAIL_CONFIG.brandName)} password.</p>
<p>For your security, use the button below to create a new password. This link is time-limited and can be used once:</p>

<div class="cta-wrap">
    <a href="${resetUrl}" class="cta-button reset-button">Reset Your Password</a>
</div>

<div class="warning-card">
    <h3 style="margin-top:0;color:#d97706;">⚠️ Important</h3>
    <ul class="feature-list">
        <li style="border-left-color:#f59e0b;">Link expires in 24 hours</li>
        <li style="border-left-color:#f59e0b;">One-time use</li>
        <li style="border-left-color:#f59e0b;">Ignore if you didn’t request this</li>
        <li style="border-left-color:#f59e0b;">Current password remains active until reset</li>
    </ul>
</div>

<div class="info-card">
    <p><strong>🔗 Alternative Access:</strong></p>
    <p>If the button doesn't work, copy this link:</p>
    <div class="link-display" style="color:#f59e0b;">${escapeHtml_(resetUrl)}</div>
</div>

<hr class="divider">

<h3 style="color:#334155;margin-bottom:20px;">🛡️ Security Tips</h3>
<ul class="feature-list">
    <li>Use a unique, strong password (12+ characters)</li>
    <li>Mix upper/lowercase, numbers, and symbols</li>
    <li>Never share your password</li>
    <li>Log out on shared devices</li>
    <li>Report suspicious activity immediately</li>
</ul>

<div style="margin-top:40px;padding:25px;background:#fef2f2;border-radius:12px;border-left:4px solid #dc2626;">
    <p><strong class="danger-text">Didn’t request this?</strong> Ignore this email or contact IT at
        <a href="mailto:${EMAIL_CONFIG.supportEmail}" class="support-link" style="color:#dc2626;">${EMAIL_CONFIG.supportEmail}</a>
    </p>
</div>
`;

    const htmlBody = renderEmail_({
      headerTitle: 'Password Reset Request',
      headerGradient: 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'Reset your password securely — link expires in 24 hours.',
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + 'Secure Password Reset';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Security'
    });

    console.log(`Password reset email sent to ${email}`);
    return true;
  } catch (error) {
    writeError('sendPasswordResetEmail', error);
    return false;
  }
}

/**
 * Send password change confirmation email
 * @param {string} email
 * @param {Object} data - { timestamp }
 * @return {boolean}
 */
function sendPasswordChangeConfirmation(email, data) {
  try {
    const changeTime = new Date(data && data.timestamp ? data.timestamp : Date.now()).toLocaleString();

    const content = `
<div class="success-badge">✅ Security Update Complete</div>
<div class="icon-large"><span class="success-icon">🎯</span></div>

<p class="subtitle"><strong>Your ${escapeHtml_(EMAIL_CONFIG.brandName)} password was successfully updated.</strong></p>
<div class="success-card">
    <h3 style="margin-top:0;color:#059669;">📋 Details</h3>
    <div class="account-details" style="background:#ffffff;border-color:#10b981;">
        <p><strong>Updated On:</strong> ${escapeHtml_(changeTime)}</p>
        <p><strong>Employee Account:</strong> ${escapeHtml_(email)}</p>
        <p><strong>Security Status:</strong> <span style="color:#10b981;font-weight:600;">Secure & Protected ✅</span>
        </p>
        <p><strong>Action Completed:</strong> Password Changed</p>
    </div>
</div>

<div style="margin:40px 0;padding:25px;background:#fef2f2;border-radius:12px;border-left:4px solid #dc2626;">
    <h3 style="margin-top:0;color:#dc2626;">🚨 Security Alert</h3>
    <p><strong class="danger-text">This wasn’t you?</strong> Contact IT immediately at
        <a href="mailto:${EMAIL_CONFIG.supportEmail}" class="support-link" style="color:#dc2626;font-weight:700;">${EMAIL_CONFIG.supportEmail}</a>
    </p>
    <p style="font-size:14px;margin-top:15px;">Report unauthorized access within 24 hours for maximum protection.</p>
</div>

<hr class="divider">

<h3 style="color:#334155;margin-bottom:20px;">🛡️ Ongoing Security</h3>
<ul class="feature-list">
    <li>Never share your password</li>
    <li>Use unique passwords per system</li>
    <li>Log out on shared/public workstations</li>
    <li>Monitor for unusual activity</li>
    <li>Report issues immediately to IT support</li>
    <li>Keep your HR contact info updated</li>
</ul>
`;

    const htmlBody = renderEmail_({
      headerTitle: 'Password Successfully Updated',
      headerGradient: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'Your password was changed successfully.',
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + 'Password Successfully Updated';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Security'
    });

    console.log(`Password change confirmation sent to ${email}`);
    return true;
  } catch (error) {
    writeError('sendPasswordChangeConfirmation', error);
    return false;
  }
}

function sendDeviceVerificationEmail(email, data) {
  try {
    const safeName = escapeHtml_(data && data.fullName ? data.fullName : 'there');
    const code = escapeHtml_(String((data && data.verificationCode) || ''));
    const expiresAt = data && data.expiresAt ? new Date(data.expiresAt) : null;
    const expiresText = expiresAt && !isNaN(expiresAt.getTime())
      ? Utilities.formatDate(expiresAt, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a')
      : '15 minutes';
    const ipAddress = escapeHtml_((data && data.ipAddress) || 'Unknown');
    const userAgent = escapeHtml_((data && data.userAgent) || 'Unknown');
    const platform = escapeHtml_((data && data.platform) || 'Unknown device');
    const languages = Array.isArray(data && data.languages) && data.languages.length
      ? escapeHtml_(data.languages.join(', '))
      : '';

    const content = `
<div class="security-badge">🛡️ New Device Sign-In Verification</div>
<p class="subtitle">Hi <span class="emphasis">${safeName}</span>,</p>
<p>We noticed a sign-in from a device or location we haven't seen before. To keep your account secure, please confirm it was you by entering the verification code below:</p>

<div class="cta-wrap">
  <div class="cta-button" style="font-size:26px;letter-spacing:4px;">${code}</div>
</div>

<div class="info-card">
  <h3 style="color:#0ea5e9;margin-top:0;">🔍 Attempt details</h3>
  <ul class="feature-list">
    <li style="border-left-color:#0ea5e9;">Observed IP: <strong>${ipAddress}</strong></li>
    <li style="border-left-color:#0ea5e9;">Device: <strong>${platform}</strong></li>
    <li style="border-left-color:#0ea5e9;">Browser: <strong>${userAgent}</strong></li>
    ${languages ? `<li style="border-left-color:#0ea5e9;">Languages: <strong>${languages}</strong></li>` : ''}
    <li style="border-left-color:#0ea5e9;">Expires: <strong>${escapeHtml_(expiresText)}</strong></li>
  </ul>
</div>

<p>If this was you, enter the code in the sign-in screen within the next few minutes. If it wasn’t you, deny the attempt immediately and our security team will be alerted.</p>
`;

    const htmlBody = renderEmail_({
      headerTitle: 'Confirm New Sign-In',
      headerGradient: 'linear-gradient(135deg, #0ea5e9 0%, #2563eb 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'Confirm whether this new sign-in was you to keep your account secure.',
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + 'Confirm New Sign-In';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Security'
    });

    console.log('Device verification email sent to', email);
    return { success: true };
  } catch (error) {
    writeError('sendDeviceVerificationEmail', error);
    return { success: false, error: error.message };
  }
}

function sendDeniedDeviceAlertEmail(details) {
  try {
    const recipient = (EMAIL_CONFIG.supportEmail || '').trim();
    if (!recipient) {
      console.warn('sendDeniedDeviceAlertEmail: support email not configured');
      return false;
    }

    const userName = escapeHtml_((details && details.userName) || 'Unknown user');
    const userEmail = escapeHtml_((details && details.userEmail) || '');
    const ipAddress = escapeHtml_((details && details.ipAddress) || 'Unknown');
    const clientIp = escapeHtml_((details && details.clientIp) || '');
    const userAgent = escapeHtml_((details && details.userAgent) || 'Unknown');
    const platform = escapeHtml_((details && details.platform) || 'Unknown');
    const verificationId = escapeHtml_((details && details.verificationId) || '');
    const fingerprint = escapeHtml_((details && details.fingerprint) || '');
    const occurredAt = escapeHtml_((details && details.occurredAt) || new Date().toISOString());

    const htmlBody = renderEmail_({
      headerTitle: 'Device Verification Denied',
      headerGradient: 'linear-gradient(135deg, #ef4444 0%, #b91c1c 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'A user denied a new device sign-in attempt.',
      contentHtml: `
<div class="security-badge" style="background:linear-gradient(135deg,#ef4444,#b91c1c)">🚨 User Denied Sign-In Attempt</div>
<p><strong>${userName}</strong> (${userEmail || 'no email on file'}) denied a new device/IP attempting to access the system.</p>

<div class="info-card">
  <h3 style="margin-top:0;color:#b91c1c;">Attempt details</h3>
  <ul class="feature-list">
    <li style="border-left-color:#b91c1c;">Server-observed IP: <strong>${ipAddress}</strong></li>
    ${clientIp ? `<li style="border-left-color:#b91c1c;">Client-reported IP: <strong>${clientIp}</strong></li>` : ''}
    <li style="border-left-color:#b91c1c;">Platform: <strong>${platform}</strong></li>
    <li style="border-left-color:#b91c1c;">User agent: <strong>${userAgent}</strong></li>
    ${verificationId ? `<li style="border-left-color:#b91c1c;">Verification ID: <strong>${verificationId}</strong></li>` : ''}
    ${fingerprint ? `<li style="border-left-color:#b91c1c;">Fingerprint: <strong>${fingerprint}</strong></li>` : ''}
    <li style="border-left-color:#b91c1c;">Denied at: <strong>${occurredAt}</strong></li>
  </ul>
</div>

<p>Please investigate this activity. If suspicious, lock the account and follow the incident response process.</p>
`
    });

    const subject = EMAIL_CONFIG.subjectPrefix + 'Denied Device Sign-In Attempt';
    sendEmail_({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Security',
      bcc: EMAIL_CONFIG.auditBcc || ''
    });

    console.log('Denied device alert sent for', userEmail || userName);
    return true;
  } catch (error) {
    writeError('sendDeniedDeviceAlertEmail', error);
    return false;
  }
}

function sendMfaCodeEmail(email, data) {
  try {
    const recipient = String(email || '').trim();
    if (!recipient) {
      throw new Error('Recipient email is required for MFA delivery');
    }

    const payload = data || {};
    const code = String(payload.code || '').trim();
    if (!code) {
      throw new Error('Verification code is required');
    }

    let expiresDisplay = '5 minutes';
    if (payload.expiresAt) {
      const expiresAt = new Date(payload.expiresAt);
      if (!isNaN(expiresAt.getTime())) {
        expiresDisplay = Utilities.formatDate(expiresAt, Session.getScriptTimeZone(), 'MMM d, yyyy h:mm a');
      }
    }

    const friendlyName = (payload.fullName || '').trim();
    const fullName = escapeHtml_(friendlyName);
    const subject = (EMAIL_CONFIG.subjectPrefix || '') + 'Your LuminaHQ verification code';
    const greetingName = fullName || 'Lumina teammate';

    const htmlBody = `
      <div class="email-container">
        <div class="header">
          <img src="${EMAIL_CONFIG.logoUrl}" alt="${EMAIL_CONFIG.brandName}" class="logo">
          <h1 class="header-title">Multi-factor verification</h1>
        </div>
        <div class="content">
          <div class="security-badge">🔐 Secure Sign-In</div>
          <p>Hi ${greetingName},</p>
          <p>Use the verification code below to finish signing in to <strong>${EMAIL_CONFIG.brandName}</strong>:</p>
          <div style="margin: 32px 0; text-align: center;">
            <div style="display: inline-block; padding: 18px 36px; font-size: 28px; font-weight: 700; letter-spacing: 6px; background: #0ea5e9; color: #fff; border-radius: 14px;">${escapeHtml_(code)}</div>
          </div>
          <p>This code expires at <strong>${expiresDisplay}</strong>. If you didn't request this code, please ignore this email.</p>
          <hr class="divider">
          <p style="font-size: 13px; color: #64748b;">Need help? Contact <a href="mailto:${EMAIL_CONFIG.supportEmail}">${EMAIL_CONFIG.supportEmail}</a>.</p>
        </div>
        <div class="footer">
          &copy; ${new Date().getFullYear()} ${EMAIL_CONFIG.orgName}. All rights reserved.
        </div>
      </div>
    `;

    const textBody = [
      `Hi ${friendlyName || 'there'},`,
      '',
      'Use the verification code below to finish signing in to ' + EMAIL_CONFIG.brandName + ':',
      '',
      code,
      '',
      'The code expires at ' + expiresDisplay + '.',
      '',
      'If you did not request this code, please ignore this email.',
      '',
      'Need help? Contact ' + EMAIL_CONFIG.supportEmail + '.',
      '',
      `— ${EMAIL_CONFIG.brandName} Team`
    ].join('\n');

    sendEmail_({
      to: recipient,
      subject: subject,
      htmlBody: htmlBody,
      textBody: textBody,
      category: 'Security'
    });

    return { success: true, message: 'Verification code sent via email.' };
  } catch (error) {
    console.error('sendMfaCodeEmail failed:', error);
    if (typeof writeError === 'function') {
      writeError('sendMfaCodeEmail', error);
    }
    return { success: false, error: error.message || 'Failed to send MFA email.' };
  }
}

/**
 * Send test email to verify email configuration
 * @param {string} email
 * @return {{success:boolean, message?:string, error?:string}}
 */
function sendTestEmail(email) {
  try {
    const content = `
<div class="success-badge">🎯 System Operational</div>
<div class="icon-large"><span style="font-size:64px;">📧</span></div>

<div class="success-card">
    <h3 style="margin-top:0;color:#059669;">🎉 Email System Status: Operational</h3>
    <p><strong>Excellent!</strong> This confirms the ${escapeHtml_(EMAIL_CONFIG.brandName)} email service is
        functioning.</p>
</div>

<div class="account-details">
    <h3 style="margin-top:0;color:#334155;">📊 Test Results</h3>
    <p><strong>Timestamp:</strong> ${escapeHtml_(new Date().toLocaleString())}</p>
    <p><strong>Recipient:</strong> ${escapeHtml_(email)}</p>
    <p><strong>System:</strong> ${escapeHtml_(EMAIL_CONFIG.brandName)} Email Service</p>
    <p><strong>Status:</strong> <span style="color:#10b981;font-weight:600;">All Systems Operational ✅</span></p>
</div>

<div class="info-card">
    <h3 style="margin-top:0;color:#0891b2;">🔧 Capabilities Verified</h3>
    <ul class="feature-list">
        <li>Onboarding emails</li>
        <li>Password resets</li>
        <li>Security notifications</li>
        <li>Professional formatting</li>
    </ul>
</div>

<p style="margin-top:30px;font-weight:500;">If you received this, account setup & security notifications are working
    properly.</p>
`;

    const htmlBody = renderEmail_({
      headerTitle: '✅ Email System Test',
      headerGradient: 'linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'Email delivery test for Lumina HQ.',
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + 'System Test Email';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'System'
    });

    return { success: true, message: 'Modern workplace email sent successfully' };
  } catch (error) {
    writeError('sendTestEmail', error);
    return { success: false, error: error.message };
  }
}

/**
 * Admin-initiated password reset email (distinct copy & header)
 * @param {string} email
 * @param {Object} data - { resetToken }
 * @return {boolean}
 */
function sendAdminPasswordResetEmail(email, data) {
  try {
    var token = (data && (data.resetToken || data.token || data.passwordSetupToken)) || '';
    var resetUrl = buildPasswordUrl_(token);

    var content = `
<div class="security-badge">🔐 Admin Security Action</div>
<div class="icon-large"><span class="security-icon">🛡️</span></div>

<p class="subtitle"><strong>An administrator initiated a password reset for your ${escapeHtml_(EMAIL_CONFIG.brandName)}
    account.</strong></p>
<p>Use the secure button below to set a new password. This link is one-time use and time-limited:</p>

<div class="cta-wrap">
    <a href="${resetUrl}" class="cta-button reset-button">Reset Your Password</a>
</div>

<div class="warning-card">
    <h3 style="margin-top:0;color:#d97706;">⚠️ Important</h3>
    <ul class="feature-list">
        <li style="border-left-color:#f59e0b;">Link may expire in 24 hours</li>
        <li style="border-left-color:#f59e0b;">One-time use</li>
        <li style="border-left-color:#f59e0b;">If you didn’t expect this, contact IT support</li>
    </ul>
</div>

<div class="info-card">
    <p><strong>🔗 Alternative Access:</strong></p>
    <p>If the button doesn't work, copy this link:</p>
    <div class="link-display" style="color:#f59e0b;">${escapeHtml_(resetUrl)}</div>
</div>

<hr class="divider">

<h3 style="color:#334155;margin-bottom:20px;">🛡️ Security Tips</h3>
<ul class="feature-list">
    <li>Use a strong, unique password (12+ chars)</li>
    <li>Never share your password</li>
    <li>Log out on shared devices</li>
    <li>Report suspicious activity immediately</li>
</ul>
`;

    var htmlBody = renderEmail_({
      headerTitle: 'Admin-Initiated Password Reset',
      headerGradient: 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'An administrator triggered a password reset for your account.',
      contentHtml: content
    });

    var subject = EMAIL_CONFIG.subjectPrefix + 'Password Reset (Admin Initiated)';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Security'
    });

    console.log('[EmailService] Admin password reset email sent to ' + email);
    return true;
  } catch (error) {
    writeError('sendAdminPasswordResetEmail', error);
    return false;
  }
}

/**
 * Resend first-login / setup email with adjusted copy
 * @param {string} email
 * @param {Object} data - { userName, fullName, passwordSetupToken }
 * @return {boolean}
 */
function sendFirstLoginResendEmail(email, data) {
  try {
    var safeName = escapeHtml_((data && (data.fullName || data.userName)) || 'there');
    var safeUser = escapeHtml_((data && data.userName) || '');
    var setupUrl = buildPasswordUrl_(data && (data.passwordSetupToken || data.token || ''));

    var content = `
<div class="welcome-badge">📨 Setup Link Resent</div>
<p class="subtitle">Hi <span class="emphasis">${safeName}</span>,</p>
<p>Here’s your ${escapeHtml_(EMAIL_CONFIG.brandName)} first-login link again. Use it to create your password and access
    your account:</p>

<div class="cta-wrap">
    <a href="${setupUrl}" class="cta-button">Set Up Your Account</a>
</div>

<div class="account-details">
    <h3 style="margin-top:0;color:#334155;">👤 Your Account</h3>
    <p><strong>Username:</strong> ${safeUser || '—'}</p>
    <p><strong>Email:</strong> ${escapeHtml_(email)}</p>
</div>

<div class="info-card">
    <p><strong>🔗 Manual Setup Link:</strong></p>
    <p>If the button above doesn't work, copy and paste this link:</p>
    <div class="link-display">${escapeHtml_(setupUrl)}</div>
</div>

<hr class="divider">

<div class="workplace-notice">
    <h3 style="margin-top:0;color:#14b8a6;">Need help?</h3>
    <p>If you didn’t request this, or the link has expired, contact IT at
        <a href="mailto:${EMAIL_CONFIG.supportEmail}" class="support-link">${EMAIL_CONFIG.supportEmail}</a>.</p>
</div>
`;

    var htmlBody = renderEmail_({
      headerTitle: 'Your Account Setup Link (Resent)',
      headerGradient: 'linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'Here is your first-login setup link again.',
      contentHtml: content
    });

    var subject = EMAIL_CONFIG.subjectPrefix + 'Your Account Setup Link (Resent)';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Onboarding'
    });

    console.log('[EmailService] First-login setup email resent to ' + email);
    return true;
  } catch (error) {
    writeError('sendFirstLoginResendEmail', error);
    return false;
  }
}

// Add these functions to your EmailService.gs file

/**
 * Send QA results email with PDF attachment when no coaching is required
 * @param {Object} emailData - Contains QA info and recipient details
 * @returns {Object} Success/error response
 */
function sendQAResultsEmail(emailData) {
  try {
    const {
      qaId,
      agentName,
      agentEmail,
      additionalEmails,
      emailNote,
      scoreResult,
      qaPdfUrl
    } = emailData;

    // Parse email addresses
    const emailList = additionalEmails.split(',').map(email => email.trim()).filter(email => email);
    
    if (emailList.length === 0) {
      return { success: false, error: 'No valid email addresses provided' };
    }

    // Get the PDF file if available
    let pdfBlob = null;
    if (qaPdfUrl) {
      try {
        // Extract file ID from Drive URL
        const fileIdMatch = qaPdfUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
        if (fileIdMatch) {
          const fileId = fileIdMatch[1];
          const file = DriveApp.getFileById(fileId);
          pdfBlob = file.getBlob();
        }
      } catch (pdfError) {
        console.warn('Could not retrieve PDF:', pdfError);
      }
    }

    // Get QA record details
    const qaRecord = getQARecordById(qaId);
    
    // Create email content
    const subject = `QA Assessment Results - ${agentName} - ${scoreResult.finalScore}% (${scoreResult.isPassing ? 'PASS' : 'FAIL'})`;
    
    const htmlBody = createQAResultsEmailTemplate({
      agentName,
      scoreResult,
      qaRecord,
      emailNote,
      qaPdfUrl
    });

    // Send to each recipient
    const emailOptions = {
      subject: subject,
      htmlBody: htmlBody
    };

    // Add PDF attachment if available
    if (pdfBlob) {
      emailOptions.attachments = [pdfBlob];
    }

    // Send emails
    let successCount = 0;
    const errors = [];

    emailList.forEach(email => {
      try {
        MailApp.sendEmail({
          to: email,
          ...emailOptions
        });
        successCount++;
      } catch (sendError) {
        errors.push(`Failed to send to ${email}: ${sendError.message}`);
      }
    });

    // Update QA record to mark feedback as shared
    updateQAFeedbackShared(qaId);

    return {
      success: true,
      emailsSent: successCount,
      totalEmails: emailList.length,
      errors: errors.length > 0 ? errors : null
    };

  } catch (error) {
    console.error('Error sending QA results email:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Generate category breakdown HTML
 */
function generateCategoryBreakdown(scoreResult) {
  const categories = {
    'Courtesy & Communication': { questions: ['q1', 'q2', 'q3', 'q4', 'q5'], maxPoints: 30 },
    'Resolution': { questions: ['q6', 'q7', 'q8', 'q9'], maxPoints: 40 },
    'Case Documentation': { questions: ['q10', 'q11', 'q12', 'q13', 'q14'], maxPoints: 30 },
    'Process Compliance': { questions: ['q15', 'q16', 'q17', 'q18'], maxPoints: 20 }
  };

  let html = '';
  
  Object.entries(categories).forEach(([categoryName, categoryInfo]) => {
    let earned = 0;
    let applicable = 0;
    
    categoryInfo.questions.forEach(q => {
      const result = scoreResult.questionResults?.[q];
      if (result && result.applicable) {
        applicable += result.weight;
        earned += result.points;
      }
    });
    
    const percentage = applicable > 0 ? Math.round((earned / applicable) * 100) : 0;
    const barColor = percentage >= 80 ? '#10b981' : percentage >= 60 ? '#f59e0b' : '#ef4444';
    
    html += `
      <div style="margin-bottom: 15px;">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 5px;">
          <span style="font-weight: 600; color: #1e293b; font-size: 14px;">${categoryName}</span>
          <span style="font-weight: 600; color: #64748b; font-size: 14px;">${earned}/${applicable} (${percentage}%)</span>
        </div>
        <div style="background: #e2e8f0; height: 8px; border-radius: 4px; overflow: hidden;">
          <div style="background: ${barColor}; height: 100%; width: ${percentage}%; transition: width 0.3s ease;"></div>
        </div>
      </div>
    `;
  });
  
  return html;
}

/**
 * Create HTML email template for QA results
 */
function createQAResultsEmailTemplate(data) {
  const { agentName, scoreResult, qaRecord, emailNote, qaPdfUrl } = data;
  
  // Determine status styling
  const statusColor = scoreResult.isPassing ? '#10b981' : '#ef4444';
  const statusIcon = scoreResult.isPassing ? '✅' : '❌';
  
  // Performance band info
  const performanceBand = scoreResult.performanceBand || { label: 'N/A', description: '' };
  
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>QA Assessment Results</title>
    </head>
    <body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f5f5f5;">
      <table width="100%" cellpadding="0" cellspacing="0" style="background-color: #f5f5f5; padding: 20px;">
        <tr>
          <td align="center">
            <table width="600" cellpadding="0" cellspacing="0" style="background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); overflow: hidden;">
              
              <!-- Header -->
              <tr>
                <td style="background: linear-gradient(135deg, #003177 0%, #004ba0 100%); padding: 30px 40px; text-align: center;">
                  <h1 style="color: #ffffff; font-size: 28px; margin: 0; font-weight: 700;">
                    QA Assessment Results
                  </h1>
                  <p style="color: rgba(255,255,255,0.9); font-size: 16px; margin: 10px 0 0 0;">
                    Quality Evaluation Report
                  </p>
                </td>
              </tr>
              
              <!-- Main Content -->
              <tr>
                <td style="padding: 40px;">
                  
                  <!-- Agent Info -->
                  <div style="background: #f8fafc; padding: 20px; border-radius: 8px; margin-bottom: 30px;">
                    <h2 style="color: #1e293b; font-size: 20px; margin: 0 0 15px 0;">Assessment Summary</h2>
                    <table width="100%" cellpadding="8" cellspacing="0">
                      <tr>
                        <td style="font-weight: 600; color: #64748b; width: 30%;">Agent:</td>
                        <td style="color: #1e293b;">${agentName}</td>
                      </tr>
                      <tr>
                        <td style="font-weight: 600; color: #64748b;">Call Date:</td>
                        <td style="color: #1e293b;">${qaRecord.CallDate || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="font-weight: 600; color: #64748b;">Case Number:</td>
                        <td style="color: #1e293b;">${qaRecord.CaseNumber || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="font-weight: 600; color: #64748b;">Audit Date:</td>
                        <td style="color: #1e293b;">${qaRecord.AuditDate || 'N/A'}</td>
                      </tr>
                      <tr>
                        <td style="font-weight: 600; color: #64748b;">Auditor:</td>
                        <td style="color: #1e293b;">${qaRecord.AuditorName || 'N/A'}</td>
                      </tr>
                    </table>
                  </div>
                  
                  <!-- Score Section -->
                  <div style="text-align: center; margin-bottom: 30px;">
                    <div style="background: ${statusColor}; color: white; padding: 20px; border-radius: 12px; margin-bottom: 20px;">
                      <h2 style="margin: 0; font-size: 48px; font-weight: 800;">${scoreResult.finalScore}%</h2>
                      <p style="margin: 10px 0 0 0; font-size: 18px; font-weight: 600;">
                        ${statusIcon} ${scoreResult.isPassing ? 'PASSED' : 'FAILED'}
                      </p>
                    </div>
                    <div style="background: #f1f5f9; padding: 15px; border-radius: 8px;">
                      <p style="margin: 0; color: #64748b; font-size: 14px; font-weight: 600;">PERFORMANCE LEVEL</p>
                      <p style="margin: 5px 0 0 0; color: #1e293b; font-size: 18px; font-weight: 700;">${performanceBand.label}</p>
                      ${performanceBand.description ? `<p style="margin: 5px 0 0 0; color: #64748b; font-size: 14px;">${performanceBand.description}</p>` : ''}
                    </div>
                  </div>
                  
                  <!-- Category Breakdown -->
                  <div style="margin-bottom: 30px;">
                    <h3 style="color: #1e293b; font-size: 18px; margin: 0 0 20px 0;">Category Breakdown</h3>
                    ${generateCategoryBreakdown(scoreResult)}
                  </div>
                  
                  <!-- Overall Feedback -->
                  ${qaRecord.OverallFeedback ? `
                    <div style="margin-bottom: 30px;">
                      <h3 style="color: #1e293b; font-size: 18px; margin: 0 0 15px 0;">Overall Feedback</h3>
                      <div style="background: #f8fafc; padding: 20px; border-radius: 8px; border-left: 4px solid #00bfff;">
                        ${qaRecord.OverallFeedback}
                      </div>
                    </div>
                  ` : ''}
                  
                  <!-- Custom Note -->
                  ${emailNote ? `
                    <div style="margin-bottom: 30px;">
                      <h3 style="color: #1e293b; font-size: 18px; margin: 0 0 15px 0;">Additional Notes</h3>
                      <div style="background: #fef3c7; padding: 20px; border-radius: 8px; border-left: 4px solid #f59e0b;">
                        ${emailNote}
                      </div>
                    </div>
                  ` : ''}
                  
                  <!-- PDF Link -->
                  ${qaPdfUrl ? `
                    <div style="text-align: center; margin-bottom: 20px;">
                      <a href="${qaPdfUrl}" target="_blank" style="
                        display: inline-block;
                        background: linear-gradient(135deg, #00bfff 0%, #0099cc 100%);
                        color: white;
                        padding: 15px 30px;
                        text-decoration: none;
                        border-radius: 8px;
                        font-weight: 600;
                        font-size: 16px;">
                        📄 View Detailed PDF Report
                      </a>
                    </div>
                  ` : ''}
                  
                </td>
              </tr>
              
              <!-- Footer -->
              <tr>
                <td style="background: #f8fafc; padding: 20px; text-align: center; border-top: 1px solid #e2e8f0;">
                  <p style="margin: 0; color: #64748b; font-size: 14px;">
                    This QA assessment was completed on ${new Date().toLocaleDateString()} as part of our continuous quality improvement process.
                  </p>
                  <p style="margin: 10px 0 0 0; color: #64748b; font-size: 12px;">
                    For questions about this assessment, please contact your supervisor or QA team.
                  </p>
                </td>
              </tr>
              
            </table>
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;
}

/**
 * Send Coaching Session Completion Email
 * @param {string} email - Coachee email
 * @param {Object} data - Coaching session data
 * @return {boolean}
 */
function sendCoachingCompletionEmail(email, data) {
  try {
    const { coacheeName, coachName, sessionDate, ackUrl, topicsCovered, actionPlan } = data;
    
    const safeName = escapeHtml_(coacheeName || 'there');
    const safeCoach = escapeHtml_(coachName || 'Your Coach');
    const safeDate = escapeHtml_(sessionDate || new Date().toLocaleDateString());

    const content = `
<div class="success-badge">🎯 Coaching Session Complete</div>

<p class="subtitle">Hi <span class="emphasis">${safeName}</span>,</p>
<p>Your coaching session with <strong>${safeCoach}</strong> on <strong>${safeDate}</strong> has been completed and documented.</p>

<div class="account-details">
    <h3 style="margin-top:0;color:#334155;">📅 Session Summary</h3>
    <p><strong>Coach:</strong> ${safeCoach}</p>
    <p><strong>Date:</strong> ${safeDate}</p>
    <p><strong>Status:</strong> <span style="color:#10b981;font-weight:600;">Completed ✅</span></p>
</div>

${topicsCovered ? `
<div class="success-card">
    <h3 style="margin-top:0;color:#059669;">📚 Topics Discussed</h3>
    <div style="background: white; padding: 15px; border-radius: 8px; margin-top: 15px;">
        ${escapeHtml_(topicsCovered)}
    </div>
</div>
` : ''}

${actionPlan ? `
<div class="info-card">
    <h3 style="margin-top:0;color:#0891b2;">🎯 Action Items</h3>
    <div style="background: white; padding: 15px; border-radius: 8px; margin-top: 15px;">
        ${escapeHtml_(actionPlan)}
    </div>
</div>
` : ''}

<div class="cta-wrap">
    <a href="${ackUrl}" target="_blank" class="cta-button">✅ Acknowledge Coaching Session</a>
</div>

<div class="warning-card">
    <h3 style="margin-top:0;color:#d97706;">📋 Required Action</h3>
    <p><strong>Please acknowledge this coaching session</strong> by clicking the button above. This confirms you've reviewed the session notes and understand the discussed points.</p>
</div>

<hr class="divider">

<div class="workplace-notice">
    <h3 style="color:#14b8a6;margin-top:0;">💡 Remember</h3>
    <ul class="feature-list">
        <li style="border-left-color:#14b8a6;">Review action items and implement improvements</li>
        <li style="border-left-color:#14b8a6;">Ask your coach questions if anything is unclear</li>
        <li style="border-left-color:#14b8a6;">Apply learned concepts in your daily work</li>
        <li style="border-left-color:#14b8a6;">Prepare for follow-up sessions as scheduled</li>
    </ul>
</div>

<div style="margin-top:30px;padding:20px;background:#f0f9ff;border-radius:12px;border-left:4px solid #0ea5e9;">
    <p><strong>Questions about this session?</strong> Reply to this email or contact your coach directly.</p>
</div>
`;

    const htmlBody = renderEmail_({
      headerTitle: 'Coaching Session Completed',
      headerGradient: 'linear-gradient(135deg, #0ea5e9 0%, #0891b2 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: `Your coaching session with ${safeCoach} is complete - acknowledgment required`,
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + `Coaching Session Complete - Acknowledgment Required`;
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Coaching'
    });

    console.log(`Coaching completion email sent to ${email}`);
    return true;
  } catch (error) {
    writeError('sendCoachingCompletionEmail', error);
    return false;
  }
}

/**
 * Send Coaching Acknowledgment Confirmation
 * @param {string} email - Coachee email
 * @param {Object} data - Acknowledgment data
 * @return {boolean}
 */
function sendCoachingAcknowledgmentConfirmation(email, data) {
  try {
    const { coacheeName, sessionDate, acknowledgedDate } = data;
    
    const safeName = escapeHtml_(coacheeName || 'there');
    const safeSessionDate = escapeHtml_(sessionDate || 'Recent');
    const safeAckDate = escapeHtml_(acknowledgedDate || new Date().toLocaleDateString());

    const content = `
<div class="success-badge">✅ Acknowledgment Received</div>
<div class="icon-large"><span class="success-icon">🎯</span></div>

<p class="subtitle">Hi <span class="emphasis">${safeName}</span>,</p>
<p><strong>Thank you for acknowledging your coaching session.</strong> This confirms you've reviewed the session notes and action items.</p>

<div class="success-card">
    <h3 style="margin-top:0;color:#059669;">📋 Confirmation Details</h3>
    <div class="account-details" style="background:#ffffff;border-color:#10b981;">
        <p><strong>Session Date:</strong> ${safeSessionDate}</p>
        <p><strong>Acknowledged On:</strong> ${safeAckDate}</p>
        <p><strong>Status:</strong> <span style="color:#10b981;font-weight:600;">Complete & Documented ✅</span></p>
    </div>
</div>

<div class="info-card">
    <h3 style="margin-top:0;color:#0891b2;">🎯 Next Steps</h3>
    <ul class="feature-list">
        <li>Continue applying the discussed improvements in your work</li>
        <li>Track your progress on action items</li>
        <li>Reach out to your coach with any questions</li>
        <li>Prepare for any scheduled follow-up sessions</li>
    </ul>
</div>

<hr class="divider">

<div class="workplace-notice">
    <h3 style="color:#14b8a6;margin-top:0;">💼 Professional Development</h3>
    <p>Coaching is an investment in your professional growth. Continue to embrace learning opportunities and apply new skills in your daily work.</p>
</div>
`;

    const htmlBody = renderEmail_({
      headerTitle: 'Coaching Acknowledgment Confirmed',
      headerGradient: 'linear-gradient(135deg, #10b981 0%, #059669 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: 'Thank you for acknowledging your coaching session',
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + 'Coaching Session Acknowledgment Confirmed';
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Coaching'
    });

    console.log(`Coaching acknowledgment confirmation sent to ${email}`);
    return true;
  } catch (error) {
    writeError('sendCoachingAcknowledgmentConfirmation', error);
    return false;
  }
}

/**
 * Update QA record to mark feedback as shared
 */
function updateQAFeedbackShared(qaId) {
  try {
    const ss = getIBTRSpreadsheet();
    const sh = ss.getSheetByName(QA_RECORDS);
    const data = sh.getDataRange().getValues();
    const headers = data.shift();
    
    const idCol = headers.indexOf('ID');
    const feedbackSharedCol = headers.indexOf('FeedbackShared');
    
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][idCol]) === String(qaId)) {
        const rowIdx = i + 2;
        sh.getRange(rowIdx, feedbackSharedCol + 1).setValue('Yes');
        break;
      }
    }
  } catch (error) {
    console.warn('Could not update feedback shared status:', error);
  }
}

/**
 * Send Coaching Follow-up Reminder
 * @param {string} email - Coachee email
 * @param {Object} data - Follow-up data
 * @return {boolean}
 */
function sendCoachingFollowUpReminder(email, data) {
  try {
    const { coacheeName, coachName, followUpDate, actionItems, originalSessionDate } = data;
    
    const safeName = escapeHtml_(coacheeName || 'there');
    const safeCoach = escapeHtml_(coachName || 'Your Coach');
    const safeFollowUpDate = escapeHtml_(followUpDate || 'Soon');
    const safeOriginalDate = escapeHtml_(originalSessionDate || 'Recent');

    const content = `
<div class="warning-card" style="background: linear-gradient(135deg, #fef3c7, #fed7aa); border-left-color: #f59e0b;">
    <div style="text-align: center; margin-bottom: 20px;">
        <span style="background: #f59e0b; color: white; padding: 8px 20px; border-radius: 25px; font-size: 14px; font-weight: 700;">📅 Follow-up Reminder</span>
    </div>
</div>

<p class="subtitle">Hi <span class="emphasis">${safeName}</span>,</p>
<p>This is a friendly reminder about your upcoming coaching follow-up session with <strong>${safeCoach}</strong>.</p>

<div class="account-details">
    <h3 style="margin-top:0;color:#334155;">📅 Session Details</h3>
    <p><strong>Original Session:</strong> ${safeOriginalDate}</p>
    <p><strong>Follow-up Date:</strong> ${safeFollowUpDate}</p>
    <p><strong>Coach:</strong> ${safeCoach}</p>
</div>

${actionItems ? `
<div class="info-card">
    <h3 style="margin-top:0;color:#0891b2;">✅ Review Your Action Items</h3>
    <p>Please review these action items before your follow-up session:</p>
    <div style="background: white; padding: 15px; border-radius: 8px; margin-top: 15px;">
        ${escapeHtml_(actionItems)}
    </div>
</div>
` : ''}

<div class="workplace-notice">
    <h3 style="color:#14b8a6;margin-top:0;">🎯 Prepare for Success</h3>
    <ul class="feature-list">
        <li style="border-left-color:#14b8a6;">Review progress on your action items</li>
        <li style="border-left-color:#14b8a6;">Prepare questions or challenges you've encountered</li>
        <li style="border-left-color:#14b8a6;">Gather examples of improvements you've made</li>
        <li style="border-left-color:#14b8a6;">Think about additional support you might need</li>
    </ul>
</div>

<div style="margin-top:30px;padding:20px;background:#f0f9ff;border-radius:12px;border-left:4px solid #0ea5e9;">
    <p><strong>Need to reschedule?</strong> Contact ${safeCoach} or your supervisor as soon as possible.</p>
</div>
`;

    const htmlBody = renderEmail_({
      headerTitle: 'Coaching Follow-up Reminder',
      headerGradient: 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)',
      logoUrl: EMAIL_CONFIG.logoUrl,
      preheader: `Follow-up coaching session with ${safeCoach} scheduled for ${safeFollowUpDate}`,
      contentHtml: content
    });

    const subject = EMAIL_CONFIG.subjectPrefix + `Coaching Follow-up Reminder - ${safeFollowUpDate}`;
    sendEmail_({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      category: 'Coaching'
    });

    console.log(`Coaching follow-up reminder sent to ${email}`);
    return true;
  } catch (error) {
    writeError('sendCoachingFollowUpReminder', error);
    return false;
  }
}

/**
 * Helper function to get performance band text
 * @param {number} score
 * @return {string}
 */
function getPerformanceBandText(score) {
  if (score >= 95) return 'Excellent';
  if (score >= 90) return 'Good';
  if (score >= 80) return 'Satisfactory';
  if (score >= 60) return 'Needs Improvement';
  return 'Unsatisfactory';
}

// ────────────────────────────────────────────────────────────────────────────
// Client-accessible wrapper functions
// ────────────────────────────────────────────────────────────────────────────

/**
 * Client-accessible QA results email sender
 */
function clientSendQAResultsEmail(email, data) {
  try {
    return sendQAResultsEmail(email, data);
  } catch (error) {
    writeError('clientSendQAResultsEmail', error);
    return false;
  }
}

/**
 * Client-accessible coaching completion email sender
 */
function clientSendCoachingCompletionEmail(email, data) {
  try {
    return sendCoachingCompletionEmail(email, data);
  } catch (error) {
    writeError('clientSendCoachingCompletionEmail', error);
    return false;
  }
}

/**
 * Client-accessible coaching acknowledgment confirmation sender
 */
function clientSendCoachingAcknowledgmentConfirmation(email, data) {
  try {
    return sendCoachingAcknowledgmentConfirmation(email, data);
  } catch (error) {
    writeError('clientSendCoachingAcknowledgmentConfirmation', error);
    return false;
  }
}

/**
 * Client-accessible coaching follow-up reminder sender
 */
function clientSendCoachingFollowUpReminder(email, data) {
  try {
    return sendCoachingFollowUpReminder(email, data);
  } catch (error) {
    writeError('clientSendCoachingFollowUpReminder', error);
    return false;
  }
}
/**
 * Hook for AuthenticationService
 */
function onPasswordChanged(email, data) {
  sendPasswordChangeConfirmation(email, data || {});
}
