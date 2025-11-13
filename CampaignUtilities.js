/**
 * CampaignUtilities.js
 * -----------------------------------------------------------------------------
 * Registry-driven utilities that bootstrap and report on campaign-specific
 * systems. Each campaign can register a set of initialization steps that run
 * whenever utilities are ensured. The module exposes helpers to ensure the
 * utilities for a single campaign or for every campaign in the tenant.
 */
(function () {
  const global = typeof globalThis === 'object' ? globalThis : (function () { return this; })();

  if (global.CampaignUtilities) {
    if (typeof global.ensureCampaignUtilities !== 'function') {
      global.ensureCampaignUtilities = function ensureCampaignUtilities() {
        return global.CampaignUtilities;
      };
    }
    if (typeof global.ensureUtilitiesForCampaign !== 'function') {
      global.ensureUtilitiesForCampaign = function ensureUtilitiesForCampaign(identifier) {
        return global.CampaignUtilities.ensureForCampaign(identifier);
      };
    }
    if (typeof global.ensureUtilitiesForAllCampaigns !== 'function') {
      global.ensureUtilitiesForAllCampaigns = function ensureUtilitiesForAllCampaigns() {
        return global.CampaignUtilities.ensureForAllCampaigns();
      };
    }
    if (typeof global.getCampaignUtilitiesStatus !== 'function') {
      global.getCampaignUtilitiesStatus = function getCampaignUtilitiesStatus(identifier) {
        if (identifier || identifier === 0) {
          return global.CampaignUtilities.getStatusForCampaign(identifier);
        }
        return global.CampaignUtilities.getStatusForAllCampaigns();
      };
    }
    return;
  }

  const DEFAULT_ICON = 'fas fa-bullhorn';
  const definitions = [];
  let defaultDefinition = null;

  function normalizeKey(value) {
    if (value === null || value === undefined) return '';
    return String(value).trim().toLowerCase();
  }

  function logError(where, error) {
    try { console.error('[CampaignUtilities] ' + where + ':', error); } catch (e) {}
    if (typeof safeWriteError === 'function') {
      try { safeWriteError(where, error); } catch (e) {}
    }
  }

  function logWarning(where, message) {
    try { console.warn('[CampaignUtilities] ' + where + ':', message); } catch (e) {}
  }

  function isActiveValue(value) {
    if (value === true) return true;
    if (value === false || value === null || value === undefined) return false;
    const normalized = normalizeKey(value);
    if (!normalized) return false;
    return normalized === 'true' || normalized === '1' || normalized === 'yes' || normalized === 'y' || normalized === 'active';
  }

  function safeReadSheetRecords(sheetName) {
    if (!sheetName) return [];
    try {
      if (typeof readSheet === 'function') {
        const data = readSheet(sheetName);
        if (Array.isArray(data)) return data;
      }
    } catch (error) {
      logError('safeReadSheetRecords(' + sheetName + ')', error);
    }
    return [];
  }

  function buildCampaignIdentity(campaign) {
    if (!campaign || typeof campaign !== 'object') {
      return {
        campaignId: '',
        campaignName: '',
        campaignDescription: ''
      };
    }
    const id = campaign.id || campaign.ID || '';
    const name = campaign.name || campaign.Name || '';
    const description = campaign.description || campaign.Description || '';
    return {
      campaignId: id ? String(id) : '',
      campaignName: name ? String(name) : '',
      campaignDescription: description ? String(description) : ''
    };
  }

  function buildSummary(campaign) {
    const identity = buildCampaignIdentity(campaign);
    const summary = {
      campaignId: identity.campaignId,
      campaignName: identity.campaignName,
      description: identity.campaignDescription,
      pageCount: 0,
      activePageCount: 0,
      categoryCount: 0,
      activeCategoryCount: 0,
      systemPageCount: null,
      coveragePercent: null,
      timestamp: new Date().toISOString()
    };

    if (!summary.campaignId) return summary;

    const idKey = normalizeKey(summary.campaignId);
    const campaignPages = safeReadSheetRecords(typeof CAMPAIGN_PAGES_SHEET !== 'undefined' ? CAMPAIGN_PAGES_SHEET : 'CampaignPages');
    const pageCategories = safeReadSheetRecords(typeof PAGE_CATEGORIES_SHEET !== 'undefined' ? PAGE_CATEGORIES_SHEET : 'PageCategories');
    const systemPages = safeReadSheetRecords(typeof PAGES_SHEET !== 'undefined' ? PAGES_SHEET : 'Pages');

    for (let i = 0; i < campaignPages.length; i++) {
      const row = campaignPages[i] || {};
      const rowCampaignId = normalizeKey(row.CampaignID || row.CampaignId || row.campaignId);
      if (rowCampaignId !== idKey) continue;
      summary.pageCount++;
      if (isActiveValue(row.IsActive)) summary.activePageCount++;
    }

    for (let j = 0; j < pageCategories.length; j++) {
      const row = pageCategories[j] || {};
      const rowCampaignId = normalizeKey(row.CampaignID || row.CampaignId || row.campaignId);
      if (rowCampaignId !== idKey) continue;
      summary.categoryCount++;
      if (isActiveValue(row.IsActive)) summary.activeCategoryCount++;
    }

    if (Array.isArray(systemPages)) {
      summary.systemPageCount = systemPages.length;
      if (systemPages.length > 0) {
        summary.coveragePercent = Math.min(100, Math.round((summary.activePageCount / systemPages.length) * 100));
      }
    }

    return summary;
  }

  function fetchAllCampaigns() {
    try {
      if (typeof csGetAllCampaigns === 'function') {
        const campaigns = csGetAllCampaigns();
        if (Array.isArray(campaigns) && campaigns.length) {
          return campaigns.map(function (c) {
            return {
              id: c.id || c.ID || '',
              ID: c.id || c.ID || '',
              name: c.name || c.Name || '',
              Name: c.name || c.Name || '',
              description: c.description || c.Description || ''
            };
          });
        }
      }
    } catch (error) {
      logError('fetchAllCampaigns.csGetAllCampaigns', error);
    }

    try {
      const rows = safeReadSheetRecords(typeof CAMPAIGNS_SHEET !== 'undefined' ? CAMPAIGNS_SHEET : 'Campaigns');
      if (Array.isArray(rows) && rows.length) {
        return rows.map(function (row) {
          return {
            id: row.ID || row.Id || row.id || '',
            ID: row.ID || row.Id || row.id || '',
            name: row.Name || row.name || '',
            Name: row.Name || row.name || '',
            description: row.Description || row.description || ''
          };
        });
      }
    } catch (error) {
      logError('fetchAllCampaigns.readSheet', error);
    }

    return [];
  }

  function resolveCampaign(identifier) {
    if (identifier && typeof identifier === 'object') {
      const idKey = normalizeKey(identifier.id || identifier.ID);
      const nameKey = normalizeKey(identifier.name || identifier.Name);
      const campaigns = fetchAllCampaigns();

      if (campaigns.length) {
        if (idKey) {
          for (let i = 0; i < campaigns.length; i++) {
            const campaign = campaigns[i];
            if (normalizeKey(campaign.id || campaign.ID) === idKey) {
              return campaign;
            }
          }
        }
        if (nameKey) {
          for (let j = 0; j < campaigns.length; j++) {
            const campaignByName = campaigns[j];
            if (normalizeKey(campaignByName.name || campaignByName.Name) === nameKey) {
              return campaignByName;
            }
          }
        }
      }

      if (idKey || nameKey) {
        return {
          id: identifier.id || identifier.ID || '',
          ID: identifier.id || identifier.ID || '',
          name: identifier.name || identifier.Name || '',
          Name: identifier.name || identifier.Name || '',
          description: identifier.description || identifier.Description || ''
        };
      }
    }

    const key = normalizeKey(identifier);
    if (!key) return null;

    const campaigns = fetchAllCampaigns();
    for (let i = 0; i < campaigns.length; i++) {
      const campaign = campaigns[i];
      if (normalizeKey(campaign.id || campaign.ID) === key) {
        return campaign;
      }
      if (normalizeKey(campaign.name || campaign.Name) === key) {
        return campaign;
      }
    }

    const definition = resolveDefinitionByKey(key);
    if (definition && definition.matchers && definition.matchers.length) {
      for (let j = 0; j < campaigns.length; j++) {
        const match = campaigns[j];
        const nameKey = normalizeKey(match.name || match.Name);
        if (definition.matchers.indexOf(nameKey) !== -1) {
          return match;
        }
      }
    }

    return null;
  }

  function resolveDefinitionByKey(key) {
    const normalized = normalizeKey(key);
    if (!normalized) return null;
    for (let i = 0; i < definitions.length; i++) {
      if (definitions[i].key === normalized) return definitions[i];
    }
    return null;
  }

  function resolveDefinitionForCampaign(campaign) {
    const nameKey = normalizeKey(campaign && (campaign.name || campaign.Name));
    const idKey = normalizeKey(campaign && (campaign.id || campaign.ID));

    for (let i = 0; i < definitions.length; i++) {
      const def = definitions[i];
      if (def.matchers.indexOf(nameKey) !== -1) return def;
      if (idKey && def.matchers.indexOf(idKey) !== -1) return def;
    }

    return defaultDefinition;
  }

  function resolveArguments(argsTemplate, context) {
    if (!Array.isArray(argsTemplate) || !argsTemplate.length) {
      return { args: [] };
    }

    const args = [];
    let missingReason = null;

    for (let i = 0; i < argsTemplate.length; i++) {
      const value = argsTemplate[i];
      if (value === ':campaignId') {
        if (context && context.campaignId) {
          args.push(context.campaignId);
        } else {
          missingReason = missingReason || 'campaignId unavailable';
        }
      } else if (value === ':campaignName') {
        if (context && context.campaignName) {
          args.push(context.campaignName);
        } else {
          missingReason = missingReason || 'campaignName unavailable';
        }
      } else if (value === ':campaign') {
        if (context && context.campaign) {
          args.push(context.campaign);
        } else {
          missingReason = missingReason || 'campaign unavailable';
        }
      } else if (value === ':context') {
        args.push(context);
      } else if (value === ':enhancedCategories') {
        if (context && context.enhancedCategories) {
          args.push(context.enhancedCategories);
        } else {
          missingReason = missingReason || 'enhanced categories unavailable';
        }
      } else {
        args.push(value);
      }
    }

    return { args: args, missingReason: missingReason };
  }

  function runFunctionByName(stepName, functionName, argsTemplate, context) {
    const fn = (function () {
      if (!functionName) return null;
      const parts = String(functionName).split('.');
      let target = global;
      for (let i = 0; i < parts.length; i++) {
        if (!target) return null;
        target = target[parts[i]];
      }
      return typeof target === 'function' ? target : null;
    })();

    if (!fn) {
      return { success: false, skipped: true, message: 'Function ' + functionName + ' not available' };
    }

    const resolved = resolveArguments(argsTemplate || [], context);
    if (resolved.missingReason) {
      return { success: false, skipped: true, message: resolved.missingReason };
    }

    try {
      const output = fn.apply(null, resolved.args);
      if (output && typeof output === 'object') {
        return output;
      }
      return { success: true, result: output };
    } catch (error) {
      logError('runFunctionByName(' + stepName + ')', error);
      return { success: false, error: error.message || String(error) };
    }
  }

  function runInlineFunction(stepName, fn, context) {
    try {
      const output = fn(context);
      if (output && typeof output === 'object') {
        return output;
      }
      return { success: true, result: output };
    } catch (error) {
      logError('runInlineFunction(' + stepName + ')', error);
      return { success: false, error: error.message || String(error) };
    }
  }

  function normalizeStep(stepConfig) {
    if (!stepConfig) return null;

    if (typeof stepConfig === 'function') {
      const name = stepConfig.name || 'step';
      return {
        name: name,
        invoke: function (context) {
          return runInlineFunction(name, stepConfig, context);
        }
      };
    }

    if (typeof stepConfig === 'string') {
      const name = stepConfig;
      return {
        name: name,
        invoke: function (context) {
          return runFunctionByName(name, stepConfig, [], context);
        }
      };
    }

    if (typeof stepConfig === 'object') {
      const name = stepConfig.name || stepConfig.functionName || 'step';
      if (typeof stepConfig.run === 'function') {
        const fn = stepConfig.run;
        return {
          name: name,
          invoke: function (context) {
            return runInlineFunction(name, fn, context);
          }
        };
      }
      if (typeof stepConfig.functionName === 'string') {
        const functionName = stepConfig.functionName;
        const argsTemplate = Array.isArray(stepConfig.args) ? stepConfig.args.slice() : [];
        return {
          name: name,
          invoke: function (context) {
            return runFunctionByName(name, functionName, argsTemplate, context);
          }
        };
      }
    }

    return null;
  }

  function exportDefinition(definition) {
    if (!definition) return null;
    return {
      key: definition.key,
      name: definition.name,
      icon: definition.icon,
      description: definition.description,
      tags: definition.tags.slice(),
      matchers: definition.matchers.slice(),
      stepCount: definition.steps.length
    };
  }

  function normalizeStepResult(stepName, rawResult) {
    const normalized = {
      name: stepName,
      success: true,
      skipped: false
    };

    if (rawResult === null || rawResult === undefined) {
      return normalized;
    }

    if (typeof rawResult === 'boolean') {
      normalized.success = rawResult;
      return normalized;
    }

    if (typeof rawResult === 'string') {
      normalized.message = rawResult;
      return normalized;
    }

    if (typeof rawResult === 'object') {
      if (Object.prototype.hasOwnProperty.call(rawResult, 'success')) {
        normalized.success = rawResult.success !== false;
      }
      if (Object.prototype.hasOwnProperty.call(rawResult, 'skipped')) {
        normalized.skipped = rawResult.skipped === true;
      }
      if (Object.prototype.hasOwnProperty.call(rawResult, 'message')) {
        normalized.message = rawResult.message;
      }
      if (Object.prototype.hasOwnProperty.call(rawResult, 'error')) {
        normalized.error = rawResult.error;
        if (typeof rawResult.success === 'undefined') {
          normalized.success = false;
        }
      }
      if (Object.prototype.hasOwnProperty.call(rawResult, 'result')) {
        normalized.result = rawResult.result;
      } else if (Object.prototype.hasOwnProperty.call(rawResult, 'raw')) {
        normalized.result = rawResult.raw;
      }
      normalized.raw = rawResult;
      return normalized;
    }

    normalized.result = rawResult;
    return normalized;
  }

  function registerDefinition(config) {
    if (!config) return null;
    const key = normalizeKey(config.key || config.name);
    if (!key) {
      logWarning('registerDefinition', 'Skipping definition without key or name');
      return null;
    }

    for (let i = 0; i < definitions.length; i++) {
      if (definitions[i].key === key) {
        return definitions[i];
      }
    }

    const includeDefaultSteps = typeof config.includeDefaultSteps === 'boolean'
      ? config.includeDefaultSteps
      : (key !== 'default');

    const tags = Array.isArray(config.tags) ? config.tags.slice() : [];
    const matcherSeeds = []
      .concat(config.matchers || [])
      .concat(config.aliases || [])
      .concat(config.key ? [config.key] : [])
      .concat(config.name ? [config.name] : []);
    const matchers = [];
    for (let i = 0; i < matcherSeeds.length; i++) {
      const normalized = normalizeKey(matcherSeeds[i]);
      if (normalized && matchers.indexOf(normalized) === -1) {
        matchers.push(normalized);
      }
    }

    const stepConfigs = [];
    if (includeDefaultSteps && defaultDefinition && key !== defaultDefinition.key) {
      for (let i = 0; i < defaultDefinition.steps.length; i++) {
        stepConfigs.push(defaultDefinition.steps[i]);
      }
    }

    const customSteps = []
      .concat(Array.isArray(config.steps) ? config.steps : [])
      .concat(Array.isArray(config.initializers) ? config.initializers : [])
      .concat(typeof config.initializer === 'function' ? [config.initializer] : []);

    for (let j = 0; j < customSteps.length; j++) {
      const step = normalizeStep(customSteps[j]);
      if (step) stepConfigs.push(step);
    }

    const definition = {
      key: key,
      name: config.name || config.key || 'Campaign',
      icon: config.icon || DEFAULT_ICON,
      description: config.description || '',
      tags: tags,
      matchers: matchers,
      steps: stepConfigs
    };

    definitions.push(definition);
    if (!defaultDefinition || key === 'default') {
      defaultDefinition = definition;
    }

    return definition;
  }

  function ensureForCampaign(identifier) {
    const campaign = resolveCampaign(identifier);
    if (!campaign) {
      return { success: false, error: 'Campaign not found', steps: [], summary: null };
    }

    const definition = resolveDefinitionForCampaign(campaign);
    const identity = buildCampaignIdentity(campaign);
    const context = {
      campaign: campaign,
      campaignId: identity.campaignId,
      campaignName: identity.campaignName,
      campaignDescription: identity.campaignDescription,
      definition: definition,
      enhancedCategories: null,
      runAt: new Date().toISOString()
    };

    if (typeof getEnhancedPageCategories === 'function') {
      try {
        context.enhancedCategories = getEnhancedPageCategories();
      } catch (error) {
        logError('getEnhancedPageCategories', error);
      }
    }

    const steps = [];
    let overallSuccess = true;

    if (definition && definition.steps.length) {
      for (let i = 0; i < definition.steps.length; i++) {
        const step = definition.steps[i];
        const rawResult = step.invoke(context);
        const normalized = normalizeStepResult(step.name, rawResult);
        steps.push(normalized);
        if (!normalized.skipped && normalized.success === false) {
          overallSuccess = false;
        }
      }
    } else {
      steps.push({
        name: 'no-ops',
        success: true,
        skipped: true,
        message: 'No utility steps registered for this campaign.'
      });
    }

    const summary = buildSummary(campaign);

    return {
      success: overallSuccess,
      campaign: identity,
      definition: exportDefinition(definition),
      steps: steps,
      summary: summary
    };
  }

  function ensureForAllCampaigns() {
    const campaigns = fetchAllCampaigns();
    const details = [];
    let succeeded = 0;
    let failed = 0;

    for (let i = 0; i < campaigns.length; i++) {
      const result = ensureForCampaign(campaigns[i]);
      details.push(result);
      if (result && result.success) succeeded++;
      else failed++;
    }

    return {
      success: failed === 0,
      totalCampaigns: campaigns.length,
      succeeded: succeeded,
      failed: failed,
      details: details
    };
  }

  function getStatusForCampaign(identifier) {
    const campaign = resolveCampaign(identifier);
    if (!campaign) {
      return { success: false, error: 'Campaign not found', summary: null };
    }

    const definition = resolveDefinitionForCampaign(campaign);
    const summary = buildSummary(campaign);

    return {
      success: true,
      campaign: buildCampaignIdentity(campaign),
      definition: exportDefinition(definition),
      summary: summary
    };
  }

  function getStatusForAllCampaigns() {
    const campaigns = fetchAllCampaigns();
    const summaries = [];
    for (let i = 0; i < campaigns.length; i++) {
      summaries.push(getStatusForCampaign(campaigns[i]));
    }
    return {
      success: true,
      totalCampaigns: campaigns.length,
      summaries: summaries
    };
  }

  function listDefinitions() {
    return definitions.map(function (definition) { return exportDefinition(definition); });
  }

  const module = {
    register: registerDefinition,
    ensureForCampaign: ensureForCampaign,
    ensureForAllCampaigns: ensureForAllCampaigns,
    getStatusForCampaign: getStatusForCampaign,
    getStatusForAllCampaigns: getStatusForAllCampaigns,
    listDefinitions: listDefinitions
  };

  global.CampaignUtilities = module;
  global.ensureCampaignUtilities = function ensureCampaignUtilities() { return module; };
  if (typeof global.ensureUtilitiesForCampaign !== 'function') {
    global.ensureUtilitiesForCampaign = function ensureUtilitiesForCampaign(identifier) {
      return module.ensureForCampaign(identifier);
    };
  }
  if (typeof global.ensureUtilitiesForAllCampaigns !== 'function') {
    global.ensureUtilitiesForAllCampaigns = function ensureUtilitiesForAllCampaigns() {
      return module.ensureForAllCampaigns();
    };
  }
  if (typeof global.getCampaignUtilitiesStatus !== 'function') {
    global.getCampaignUtilitiesStatus = function getCampaignUtilitiesStatus(identifier) {
      if (identifier || identifier === 0) {
        return module.getStatusForCampaign(identifier);
      }
      return module.getStatusForAllCampaigns();
    };
  }

  registerDefinition({
    key: 'default',
    name: 'Standard Campaign Utilities',
    icon: DEFAULT_ICON,
    description: 'Baseline utilities applied to every campaign.',
    includeDefaultSteps: false,
    steps: [
      { name: 'Ensure campaign pages', functionName: 'createCampaignPagesFromSystem', args: [':campaignId'] },
      { name: 'Ensure enhanced categories', functionName: 'createEnhancedCategoriesForCampaign', args: [':campaignId', ':enhancedCategories'] },
      { name: 'Assign pages to categories', functionName: 'assignPagesToEnhancedCategories', args: [':campaignId'] }
    ]
  });

  registerDefinition({
    key: 'lumina-hq',
    name: 'Lumina HQ',
    icon: 'fas fa-building',
    description: 'Internal Lumina HQ workspace utilities.',
    matchers: ['lumina hq', 'lumina'],
    tags: ['internal', 'operations'],
    steps: [
      { name: 'Initialize enhanced system pages', functionName: 'initializeEnhancedSystemPages' }
    ]
  });

  registerDefinition({
    key: 'credit-suite',
    name: 'Credit Suite',
    icon: 'fas fa-credit-card',
    description: 'Credit Suite QA utilities and reporting.',
    matchers: ['credit suite', 'creditsuite'],
    tags: ['qa', 'credit'],
    steps: [
      { name: 'Initialize Credit Suite QA system', functionName: 'initializeCreditSuiteQASystem' }
    ]
  });

  registerDefinition({
    key: 'independence-insurance',
    name: 'Independence Insurance',
    icon: 'fas fa-shield-alt',
    description: 'Independence Insurance QA and coaching utilities.',
    matchers: ['independence insurance', 'independence'],
    tags: ['qa', 'insurance'],
    steps: [
      { name: 'Initialize Independence QA system', functionName: 'initializeIndependenceQASystem' }
    ]
  });

})();
