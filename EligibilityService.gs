/**
 * EligibilityService.gs
 * -----------------------------------------------------------------------------
 * Evaluates eligibility hints and lifecycle suggestions using the Eligibility
 * rules defined in Sheets.
 */
(function bootstrapEligibilityService(global) {
  if (!global) return;
  if (global.EligibilityService && typeof global.EligibilityService === 'object') {
    return;
  }

  var IdentityRepository = global.IdentityRepository;
  var PolicyService = global.PolicyService;

  if (!IdentityRepository || !PolicyService) {
    throw new Error('EligibilityService requires IdentityRepository and PolicyService');
  }

  function parseJson(value) {
    if (!value) return {};
    if (typeof value === 'object') return value;
    try {
      return JSON.parse(value);
    } catch (err) {
      return {};
    }
  }

  function evaluateEligibility(userId, campaignId) {
    var rules = PolicyService.listEligibilityRules(campaignId).filter(function(rule) {
      return rule.Active === 'Y' || rule.Active === true;
    });
    var employmentHistory = IdentityRepository.list('EmploymentStatus').filter(function(row) {
      return row.UserId === userId && row.CampaignId === campaignId;
    });
    var hints = [];
    rules.forEach(function(rule) {
      var params = parseJson(rule.ParamsJSON);
      switch (rule.RuleType) {
        case 'Insurance':
          hints.push(evaluateInsuranceRule(employmentHistory, params));
          break;
        case 'Promo':
        case 'Promotion':
          hints.push(evaluatePromotionRule(userId, campaignId, params));
          break;
        case 'Terminate':
          hints.push(evaluateTerminationRule(userId, campaignId, params));
          break;
        case 'Watch':
          hints.push(evaluateWatchRule(userId, campaignId, params));
          break;
        default:
          break;
      }
    });
    return hints.filter(Boolean);
  }

  function evaluateInsuranceRule(history, params) {
    params = params || {};
    var minTenureDays = params.minTenureDays || 90;
    if (!history || !history.length) {
      return {
        category: 'Insurance',
        status: 'Ineligible',
        reason: 'No employment records found.'
      };
    }
    var hired = history.find(function(record) {
      return record.State === 'Hired' || record.State === 'Active';
    });
    if (!hired) {
      return {
        category: 'Insurance',
        status: 'Ineligible',
        reason: 'No active employment status.'
      };
    }
    var hiredDate = new Date(hired.EffectiveDate);
    var today = new Date();
    var tenureDays = Math.floor((today.getTime() - hiredDate.getTime()) / (1000 * 60 * 60 * 24));
    if (tenureDays >= minTenureDays) {
      return {
        category: 'Insurance',
        status: 'Eligible',
        reason: 'Minimum tenure satisfied (' + tenureDays + ' days).'
      };
    }
    return {
      category: 'Insurance',
      status: 'Pending',
      reason: 'Tenure ' + tenureDays + '/' + minTenureDays + ' days. Next review on ' + formatFutureDate(minTenureDays - tenureDays) + '.'
    };
  }

  function evaluatePromotionRule(userId, campaignId, params) {
    params = params || {};
    var qaScores = IdentityRepository.list('QualityScores').filter(function(score) {
      return score.UserId === userId && score.CampaignId === campaignId;
    });
    var attendance = IdentityRepository.list('Attendance').filter(function(row) {
      return row.UserId === userId && row.CampaignId === campaignId;
    });
    var watchlist = IdentityRepository.list('UserCampaigns').some(function(row) {
      return row.UserId === userId && row.CampaignId === campaignId && row.Watchlist === 'Y';
    });
    var qaTarget = params.qaTarget || 90;
    var attendanceTarget = params.attendanceTarget || 95;
    var windowDays = params.windowDays || 60;
    var qaOk = averageRecentScores(qaScores, windowDays) >= qaTarget;
    var attendanceOk = averageRecentScores(attendance, windowDays, 'Attendance') >= attendanceTarget;
    if (qaOk && attendanceOk && !watchlist) {
      return {
        category: 'Promotion',
        status: 'Eligible',
        reason: 'QA and attendance targets met over last ' + windowDays + ' days.'
      };
    }
    return {
      category: 'Promotion',
      status: 'Not Ready',
      reason: 'Targets unmet or watchlist flag present.'
    };
  }

  function evaluateTerminationRule(userId, campaignId, params) {
    params = params || {};
    var attendance = IdentityRepository.list('Attendance').filter(function(row) {
      return row.UserId === userId && row.CampaignId === campaignId;
    });
    var threshold = params.absenceThreshold || 3;
    var lookback = params.lookbackDays || 30;
    var consecutiveMisses = countConsecutiveAbsences(attendance, lookback);
    if (consecutiveMisses >= threshold) {
      return {
        category: 'Termination',
        status: 'Review',
        reason: 'Repeated absences detected (' + consecutiveMisses + ' in last ' + lookback + ' days).'
      };
    }
    return null;
  }

  function evaluateWatchRule(userId, campaignId, params) {
    params = params || {};
    if (!params.kpiThreshold) {
      return null;
    }
    var kpi = IdentityRepository.list('Performance').filter(function(row) {
      return row.UserId === userId && row.CampaignId === campaignId;
    });
    var belowThreshold = kpi.some(function(row) {
      return Number(row.Score || 0) < params.kpiThreshold;
    });
    if (belowThreshold) {
      return {
        category: 'Watch',
        status: 'Flagged',
        reason: 'Performance metrics below threshold of ' + params.kpiThreshold
      };
    }
    return null;
  }

  function averageRecentScores(rows, windowDays, valueKey) {
    if (!rows || !rows.length) {
      return 0;
    }
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - (windowDays || 30));
    var total = 0;
    var count = 0;
    rows.forEach(function(row) {
      var date = new Date(row.Date || row.EffectiveDate || row.Timestamp);
      if (date >= cutoff) {
        total += Number(row.Score || row[valueKey || 'Score'] || 0);
        count++;
      }
    });
    return count ? Math.round((total / count) * 100) / 100 : 0;
  }

  function countConsecutiveAbsences(rows, lookbackDays) {
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - (lookbackDays || 30));
    var misses = 0;
    rows.forEach(function(row) {
      var date = new Date(row.Date || row.EffectiveDate || row.Timestamp);
      if (date >= cutoff && row.Status === 'Absent') {
        misses++;
      }
    });
    return misses;
  }

  function formatFutureDate(daysAhead) {
    var date = new Date();
    date.setDate(date.getDate() + Math.max(daysAhead, 0));
    return date.toISOString().split('T')[0];
  }

  global.EligibilityService = {
    evaluateEligibility: evaluateEligibility
  };
})(GLOBAL_SCOPE);
