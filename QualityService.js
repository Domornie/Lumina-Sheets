/**
 * QualityService.js
 * Background helpers for warming caches related to the quality domain.
 */

var QualityService = typeof QualityService !== 'undefined' ? QualityService : {};

QualityService.queueBackgroundInitialization = function (context, options) {
  var safeConsole = (typeof console !== 'undefined' && console) ? console : {
    log: function () { },
    warn: function () { },
    error: function () { }
  };

  try {
    var shouldWarm = true;
    if (options && Array.isArray(options.entities)) {
      shouldWarm = options.entities.some(function (name) {
        if (!name) return false;
        var normalized = String(name).toLowerCase();
        return normalized === 'quality' || normalized === 'qa' || normalized === 'qualityrecords';
      });
    }

    if (!shouldWarm) {
      return;
    }

    if (typeof DatabaseManager === 'undefined' || !DatabaseManager || typeof DatabaseManager.table !== 'function') {
      return;
    }

    if (typeof resolveLuminaEntityDefinition !== 'function') {
      return;
    }

    var definition;
    try {
      definition = resolveLuminaEntityDefinition('quality');
    } catch (resolveError) {
      if (safeConsole && typeof safeConsole.warn === 'function') {
        safeConsole.warn('QualityService.queueBackgroundInitialization: unable to resolve entity', resolveError);
      }
      return;
    }

    if (!definition || !definition.tableName) {
      return;
    }

    var readOptions = { cache: true, limit: 25 };
    if (Array.isArray(definition.summaryColumns) && definition.summaryColumns.length) {
      readOptions.columns = definition.summaryColumns.slice();
    }

    DatabaseManager.table(definition.tableName, context).read(readOptions);
  } catch (error) {
    if (safeConsole && typeof safeConsole.error === 'function') {
      safeConsole.error('QualityService.queueBackgroundInitialization failed:', error);
    }
    if (typeof writeError === 'function') {
      try {
        writeError('QualityService.queueBackgroundInitialization', error);
      } catch (loggingError) {
        if (safeConsole && typeof safeConsole.error === 'function') {
          safeConsole.error('QualityService.queueBackgroundInitialization logging error:', loggingError);
        }
      }
    }
  }
};
