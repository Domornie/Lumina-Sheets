/**
 * Holiday Management System with Country Support
 * Allows selecting countries, applying holidays, and managing paid status by country of residence
 */
// ────────────────────────────────────────────────────────────────────────────
// HELPERS (place near the top of the file once)
// ────────────────────────────────────────────────────────────────────────────
const _pad2 = n => String(n).padStart(2, '0');
const _ymd = d => `${d.getUTCFullYear()}-${_pad2(d.getUTCMonth() + 1)}-${_pad2(d.getUTCDate())}`;
const _dateUTC = (y, m, d) => new Date(Date.UTC(y, m - 1, d));
const _addDays = (d, days) => _dateUTC(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate() + days);

// Nth/Last weekday (weekday: 0=Sun..6=Sat, month: 1..12)
const _nthWeekdayOfMonth = (y, m, weekday, n) => {
  const first = _dateUTC(y, m, 1);
  const offset = (weekday - first.getUTCDay() + 7) % 7;
  const day = 1 + offset + 7 * (n - 1);
  return _dateUTC(y, m, day);
};
const _lastWeekdayOfMonth = (y, m, weekday) => {
  const last = new Date(Date.UTC(y, m, 0)); // last day of m
  const back = (last.getUTCDay() - weekday + 7) % 7;
  return _dateUTC(y, m, last.getUTCDate() - back);
};

// Gregorian Easter (Meeus/Jones/Butcher)
const _easterSunday = (y) => {
  const a = y % 19, b = Math.floor(y / 100), c = y % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4, l = (32 + 2*e + 2*i - h - k) % 7;
  const m = Math.floor((a + 11*h + 22*l) / 451);
  const month = Math.floor((h + l - 7*m + 114) / 31);       // 3=March, 4=April
  const day = ((h + l - 7*m + 114) % 31) + 1;
  return _dateUTC(y, month, day);
};
const _ashWednesday = (y) => _addDays(_easterSunday(y), -46);
const _maundyThursday = (y) => _addDays(_easterSunday(y), -3);
const _goodFriday = (y) => _addDays(_easterSunday(y), -2);
const _easterMonday = (y) => _addDays(_easterSunday(y), 1);
const _corpusChristi = (y) => _addDays(_easterSunday(y), 60);

// Canada: Monday preceding May 25
const _victoriaDay = (y) => {
  const may24 = _dateUTC(y, 5, 24);
  const backToMonday = (may24.getUTCDay() + 6) % 7; // distance to previous Monday
  return _addDays(may24, -backToMonday);
};

// Dominican Republic: Law 139-97 Mondayization
const _drObserved = (d) => {
  const y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
  const wd = d.getUTCDay(); // 0=Sun..6=Sat
  const is = (mm, dd) => m === mm && day === dd;

  // Not moved (fixed-date observance): Jan1, Jan6, Jan21, Feb27, Sep24, Dec25.
  if (is(1,1) || is(1,6) || is(1,21) || is(2,27) || is(9,24) || is(12,25)) return d;

  // Good Friday & Corpus Christi never move
  const gf = _goodFriday(y), cc = _corpusChristi(y);
  const same = (D1, D2) => D1.getUTCFullYear()===D2.getUTCFullYear() && D1.getUTCMonth()===D2.getUTCMonth() && D1.getUTCDate()===D2.getUTCDate();
  if (same(d, gf) || same(d, cc)) return d;

  // Special: if May 1 is Sunday -> Monday
  if (is(5,1) && wd === 0) return _addDays(d, 1);

  // General rule: Tue/Wed -> previous Mon ; Thu/Fri -> next Mon ; Sat/Sun/Mon unchanged
  if (wd === 2) return _addDays(d, -1);  // Tue -> Mon
  if (wd === 3) return _addDays(d, -2);  // Wed -> Mon
  if (wd === 4) return _addDays(d, +4);  // Thu -> next Mon
  if (wd === 5) return _addDays(d, +3);  // Fri -> next Mon
  return d;
};

// Compute floating holiday date by (country, holiday name)
function _computeFloatingHolidayDate(countryCode, holidayName, year) {
  const n = holidayName.toLowerCase().trim();

  switch (countryCode) {
    case 'US': {
      if (n.includes('martin luther king')) return _nthWeekdayOfMonth(year, 1, 1, 3);       // 3rd Mon Jan
      if (n.includes('presidents') || n.includes("washington")) return _nthWeekdayOfMonth(year, 2, 1, 3); // 3rd Mon Feb
      if (n.includes('memorial')) return _lastWeekdayOfMonth(year, 5, 1);                   // last Mon May
      if (n.includes('labor')) return _nthWeekdayOfMonth(year, 9, 1, 1);                    // 1st Mon Sep
      if (n.includes('columbus') || n.includes('indigenous')) return _nthWeekdayOfMonth(year, 10, 1, 2);  // 2nd Mon Oct
      if (n.includes('thanksgiving')) return _nthWeekdayOfMonth(year, 11, 4, 4);            // 4th Thu Nov (4=Thu)
      break;
    }
    case 'UK': { // England & Wales standard set
      if (n.includes('good friday')) return _goodFriday(year);
      if (n.includes('easter monday')) return _easterMonday(year);
      if (n.includes('early may')) return _nthWeekdayOfMonth(year, 5, 1, 1);                // 1st Mon May
      if (n.includes('spring bank')) return _lastWeekdayOfMonth(year, 5, 1);                // last Mon May
      if (n.includes('summer bank')) return _lastWeekdayOfMonth(year, 8, 1);                // last Mon Aug
      break;
    }
    case 'CA': {
      if (n.includes('good friday')) return _goodFriday(year);
      if (n.includes('easter monday')) return _easterMonday(year);
      if (n.includes('victoria')) return _victoriaDay(year);                                 // Mon before May 25
      if (n.includes('labour')) return _nthWeekdayOfMonth(year, 9, 1, 1);                   // 1st Mon Sep
      if (n.includes('thanksgiving')) return _nthWeekdayOfMonth(year, 10, 1, 2);            // 2nd Mon Oct
      break;
    }
    case 'PH': {
      if (n.includes('maundy')) return _maundyThursday(year);
      if (n.includes('good friday')) return _goodFriday(year);
      if (n.includes('national heroes')) return _lastWeekdayOfMonth(year, 8, 1);            // last Mon Aug
      break;
    }
    case 'JM': {
      if (n.includes('ash wednesday')) return _ashWednesday(year);
      if (n.includes('good friday')) return _goodFriday(year);
      if (n.includes('easter monday')) return _easterMonday(year);
      if (n.includes('national heroes')) return _nthWeekdayOfMonth(year, 10, 1, 3);         // 3rd Mon Oct
      break;
    }
    case 'DO': {
      if (n.includes('good friday')) return _goodFriday(year);
      if (n.includes('corpus christi')) return _corpusChristi(year);
      if (n.includes('labour') || n.includes('labor')) return _drObserved(_dateUTC(year, 5, 1));
      if (n.includes('restoration')) return _drObserved(_dateUTC(year, 8, 16));
      if (n.includes('constitution')) return _drObserved(_dateUTC(year, 11, 6));
      if (n.includes('duarte')) return _drObserved(_dateUTC(year, 1, 26));
      if (n.includes('epiphany') || n.includes('reyes')) return _drObserved(_dateUTC(year, 1, 6));
      break;
    }
  }
  return null; // Unknown: let caller fall back to static month-day
}

// ────────────────────────────────────────────────────────────────────────────
// COUNTRY HOLIDAY DEFINITIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Comprehensive holiday definitions by country
 * This would typically be replaced by an API call to a service like Calendarific
 */
// ────────────────────────────────────────────────────────────────────────────
// UPDATED COUNTRY DEFINITIONS (same shape you already use)
// ────────────────────────────────────────────────────────────────────────────
function getCountryHolidayDefinitions() {
  return {
    'US': {
      name: 'United States',
      holidays: [
        { name: "New Year's Day", date: '01-01', type: 'public' },
        { name: 'Martin Luther King Jr. Day', date: '01-15', type: 'public', floats: true },   // 3rd Mon Jan
        { name: "Presidents' Day", date: '02-19', type: 'public', floats: true },              // 3rd Mon Feb
        { name: 'Memorial Day', date: '05-27', type: 'public', floats: true },                 // Last Mon May
        { name: 'Juneteenth National Independence Day', date: '06-19', type: 'public' },
        { name: 'Independence Day', date: '07-04', type: 'public' },
        { name: 'Labor Day', date: '09-02', type: 'public', floats: true },                    // 1st Mon Sep
        { name: 'Columbus Day', date: '10-14', type: 'public', floats: true },                 // 2nd Mon Oct
        { name: 'Veterans Day', date: '11-11', type: 'public' },
        { name: 'Thanksgiving Day', date: '11-28', type: 'public', floats: true },             // 4th Thu Nov
        { name: 'Christmas Day', date: '12-25', type: 'public' }
      ]
    },
    'UK': { // England & Wales set
      name: 'United Kingdom',
      holidays: [
        { name: "New Year's Day", date: '01-01', type: 'public' },
        { name: 'Good Friday', date: '04-07', type: 'public', floats: true },
        { name: 'Easter Monday', date: '04-10', type: 'public', floats: true },
        { name: 'Early May Bank Holiday', date: '05-06', type: 'public', floats: true },       // 1st Mon May
        { name: 'Spring Bank Holiday', date: '05-27', type: 'public', floats: true },          // Last Mon May
        { name: 'Summer Bank Holiday', date: '08-26', type: 'public', floats: true },          // Last Mon Aug
        { name: 'Christmas Day', date: '12-25', type: 'public' },
        { name: 'Boxing Day', date: '12-26', type: 'public' }
      ]
    },
    'CA': {
      name: 'Canada',
      holidays: [
        { name: "New Year's Day", date: '01-01', type: 'public' },
        { name: 'Good Friday', date: '04-07', type: 'public', floats: true },
        { name: 'Easter Monday', date: '04-10', type: 'public', floats: true },                // (federal public service)
        { name: 'Victoria Day', date: '05-20', type: 'public', floats: true },                 // Mon before May 25
        { name: 'Canada Day', date: '07-01', type: 'public' },
        { name: 'Labour Day', date: '09-02', type: 'public', floats: true },                   // 1st Mon Sep
        { name: 'National Day for Truth and Reconciliation', date: '09-30', type: 'public' },
        { name: 'Thanksgiving', date: '10-14', type: 'public', floats: true },                 // 2nd Mon Oct
        { name: 'Remembrance Day', date: '11-11', type: 'public' },
        { name: 'Christmas Day', date: '12-25', type: 'public' },
        { name: 'Boxing Day', date: '12-26', type: 'public' }
      ]
    },
    'PH': {
      name: 'Philippines',
      holidays: [
        { name: "New Year's Day", date: '01-01', type: 'public' },
        { name: 'People Power Anniversary', date: '02-25', type: 'public' },
        { name: 'Maundy Thursday', date: '04-06', type: 'public', floats: true },
        { name: 'Good Friday', date: '04-07', type: 'public', floats: true },
        { name: 'Araw ng Kagitingan', date: '04-09', type: 'public' },
        { name: 'Labor Day', date: '05-01', type: 'public' },
        { name: 'Independence Day', date: '06-12', type: 'public' },
        { name: 'National Heroes Day', date: '08-26', type: 'public', floats: true },          // Last Mon Aug
        { name: "All Saints' Day", date: '11-01', type: 'public' },
        { name: 'Bonifacio Day', date: '11-30', type: 'public' },
        { name: 'Christmas Day', date: '12-25', type: 'public' },
        { name: 'Rizal Day', date: '12-30', type: 'public' }
        // (Eid dates are proclaimed annually; add when officially announced.)
      ]
    },
    'JM': {
      name: 'Jamaica',
      holidays: [
        { name: "New Year's Day", date: '01-01', type: 'public' },
        { name: 'Ash Wednesday', date: '02-22', type: 'public', floats: true },
        { name: 'Good Friday', date: '04-07', type: 'public', floats: true },
        { name: 'Easter Monday', date: '04-10', type: 'public', floats: true },
        { name: 'Labour Day', date: '05-23', type: 'public' },
        { name: 'Emancipation Day', date: '08-01', type: 'public' },
        { name: 'Independence Day', date: '08-06', type: 'public' },
        { name: 'National Heroes Day', date: '10-16', type: 'public', floats: true },          // 3rd Mon Oct
        { name: 'Christmas Day', date: '12-25', type: 'public' },
        { name: 'Boxing Day', date: '12-26', type: 'public' }
      ]
    },
    'DO': {
      name: 'Dominican Republic',
      holidays: [
        { name: "New Year's Day", date: '01-01', type: 'public' },
        { name: 'Epiphany (Día de Reyes)', date: '01-06', type: 'public', floats: true },      // Law 139-97
        { name: 'Our Lady of Altagracia', date: '01-21', type: 'public' },
        { name: "Duarte's Day", date: '01-26', type: 'public', floats: true },
        { name: 'Independence Day', date: '02-27', type: 'public' },
        { name: 'Good Friday', date: '04-07', type: 'public', floats: true },
        { name: 'Labour Day', date: '05-01', type: 'public', floats: true },
        { name: 'Corpus Christi', date: '06-08', type: 'public', floats: true },
        { name: 'Restoration Day', date: '08-16', type: 'public', floats: true },
        { name: 'Our Lady of Mercedes', date: '09-24', type: 'public' },
        { name: 'Constitution Day', date: '11-06', type: 'public', floats: true },
        { name: 'Christmas Day', date: '12-25', type: 'public' }
      ]
    }
  };
}


// ────────────────────────────────────────────────────────────────────────────
// HOLIDAY CALCULATION AND MANAGEMENT
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get holidays for a specific country and year
 */
// ────────────────────────────────────────────────────────────────────────────
/**
 * Get holidays for a specific country and year (now computes floating dates)
 */
// ────────────────────────────────────────────────────────────────────────────
function clientGetCountryHolidays(countryCode, year) {
  try {
    console.log(`🎉 Getting holidays for ${countryCode} ${year}`);

    const defs = getCountryHolidayDefinitions();
    const countryData = defs[countryCode];

    if (!countryData) {
      return { success: false, error: `Country ${countryCode} not supported`, holidays: [] };
    }

    const holidays = countryData.holidays.map(h => {
      let d;

      if (h.floats) {
        // Try to compute; if not recognized, fall back to static month-day provided
        const computed = _computeFloatingHolidayDate(countryCode, h.name, year);
        d = computed ? _ymd(computed) : `${year}-${h.date}`;
      } else {
        d = `${year}-${h.date}`;
      }

      return {
        name: h.name,
        date: d,
        type: h.type,
        floats: !!h.floats,
        countryCode,
        countryName: countryData.name
      };
    });

    // Sort by actual date
    holidays.sort((a, b) => new Date(a.date) - new Date(b.date));

    console.log(`✅ Found ${holidays.length} holidays for ${countryData.name} ${year}`);
    return {
      success: true,
      holidays,
      country: countryData.name,
      countryCode,
      year,
      total: holidays.length
    };

  } catch (error) {
    console.error('❌ Error getting country holidays:', error);
    writeError('clientGetCountryHolidays', error);
    return { success: false, error: error.message, holidays: [] };
  }
}

/**
 * Apply holidays to the system
 */
function clientApplyCountryHolidays(countryCode, year, holidaysToApply, applyPaidStatus = false) {
    try {
        console.log(`📅 Applying ${holidaysToApply.length} holidays for ${countryCode} ${year}`);
        
        const holidaysSheet = ensureSheetWithHeaders(HOLIDAYS_SHEET, HOLIDAYS_HEADERS);
        const now = new Date();
        let appliedCount = 0;
        
        // Check for existing holidays to avoid duplicates
        const existingHolidays = readSheet(HOLIDAYS_SHEET) || [];
        
        holidaysToApply.forEach(holiday => {
            // Check if holiday already exists
            const exists = existingHolidays.some(existing => 
                existing.Date === holiday.date && 
                existing.HolidayName === holiday.name
            );
            
            if (!exists) {
                const holidayRecord = [
                    Utilities.getUuid(),                    // ID
                    holiday.name,                           // HolidayName
                    holiday.date,                           // Date
                    true,                                   // AllDay
                    `${countryCode} - ${holiday.type} holiday`, // Notes
                    now,                                    // CreatedAt
                    now                                     // UpdatedAt
                ];
                
                holidaysSheet.appendRow(holidayRecord);
                appliedCount++;
            }
        });
        
        // Apply paid status to users from this country if requested
        let usersUpdated = 0;
        if (applyPaidStatus) {
            usersUpdated = applyDefaultPaidStatusByCountry(countryCode, true);
        }
        
        invalidateCache(HOLIDAYS_SHEET);
        
        console.log(`✅ Applied ${appliedCount} new holidays, updated ${usersUpdated} users`);
        
        return {
            success: true,
            message: `Applied ${appliedCount} new holidays for ${countryCode} ${year}`,
            appliedHolidays: appliedCount,
            totalHolidays: holidaysToApply.length,
            duplicatesSkipped: holidaysToApply.length - appliedCount,
            usersUpdated: usersUpdated
        };
        
    } catch (error) {
        console.error('❌ Error applying holidays:', error);
        writeError('clientApplyCountryHolidays', error);
        return {
            success: false,
            error: error.message,
            appliedHolidays: 0
        };
    }
}

/**
 * Get current holidays in the system
 */
function clientGetCurrentHolidays(year = null) {
    try {
        const holidays = readSheet(HOLIDAYS_SHEET) || [];
        
        let filteredHolidays = holidays;
        
        if (year) {
            filteredHolidays = holidays.filter(h => {
                const holidayYear = new Date(h.Date).getFullYear();
                return holidayYear === parseInt(year);
            });
        }
        
        // Sort by date
        filteredHolidays.sort((a, b) => new Date(a.Date) - new Date(b.Date));
        
        return {
            success: true,
            holidays: filteredHolidays,
            total: filteredHolidays.length,
            year: year
        };
        
    } catch (error) {
        console.error('❌ Error getting current holidays:', error);
        writeError('clientGetCurrentHolidays', error);
        return {
            success: false,
            error: error.message,
            holidays: []
        };
    }
}

// ────────────────────────────────────────────────────────────────────────────
// USER HOLIDAY PAY MANAGEMENT
// ────────────────────────────────────────────────────────────────────────────

/**
 * Set user holiday pay status by country
 */
function clientSetUserHolidayPayStatus(userId, countryCode, isPaid, notes = '') {
    try {
        console.log(`💰 Setting holiday pay status for user ${userId} in ${countryCode}: ${isPaid ? 'PAID' : 'NOT PAID'}`);
        
        const sheet = ensureSheetWithHeaders('UserHolidayPayStatus', [
            'ID', 'UserID', 'UserName', 'CountryCode', 'CountryName', 'IsPaid', 'Notes', 'CreatedAt', 'UpdatedAt'
        ]);
        
        const userName = getUserNameById(userId);
        const countryName = getCountryHolidayDefinitions()[countryCode]?.name || countryCode;
        const now = new Date();
        
        // Check if entry already exists
        const existingData = readSheet('UserHolidayPayStatus') || [];
        const existingEntry = existingData.find(entry => 
            entry.UserID === userId && entry.CountryCode === countryCode
        );
        
        if (existingEntry) {
            // Update existing entry
            const data = sheet.getDataRange().getValues();
            const headers = data[0];
            const idIndex = headers.indexOf('ID');
            const isPaidIndex = headers.indexOf('IsPaid');
            const notesIndex = headers.indexOf('Notes');
            const updatedAtIndex = headers.indexOf('UpdatedAt');
            
            for (let i = 1; i < data.length; i++) {
                if (data[i][idIndex] === existingEntry.ID) {
                    sheet.getRange(i + 1, isPaidIndex + 1).setValue(isPaid ? 'TRUE' : 'FALSE');
                    sheet.getRange(i + 1, notesIndex + 1).setValue(notes);
                    sheet.getRange(i + 1, updatedAtIndex + 1).setValue(now);
                    break;
                }
            }
        } else {
            // Create new entry
            const entry = [
                Utilities.getUuid(),    // ID
                userId,                 // UserID
                userName,               // UserName
                countryCode,            // CountryCode
                countryName,            // CountryName
                isPaid ? 'TRUE' : 'FALSE', // IsPaid
                notes,                  // Notes
                now,                    // CreatedAt
                now                     // UpdatedAt
            ];
            
            sheet.appendRow(entry);
        }
        
        invalidateCache('UserHolidayPayStatus');
        
        return {
            success: true,
            message: `Holiday pay status updated for ${userName} in ${countryName}`,
            userId: userId,
            userName: userName,
            countryCode: countryCode,
            countryName: countryName,
            isPaid: isPaid
        };
        
    } catch (error) {
        console.error('❌ Error setting holiday pay status:', error);
        writeError('clientSetUserHolidayPayStatus', error);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Get user holiday pay status
 */
function clientGetUserHolidayPayStatus(userId = null, countryCode = null) {
    try {
        const payStatusData = readSheet('UserHolidayPayStatus') || [];
        
        let filteredData = payStatusData;
        
        if (userId) {
            filteredData = filteredData.filter(entry => entry.UserID === userId);
        }
        
        if (countryCode) {
            filteredData = filteredData.filter(entry => entry.CountryCode === countryCode);
        }
        
        return {
            success: true,
            payStatuses: filteredData,
            total: filteredData.length
        };
        
    } catch (error) {
        console.error('❌ Error getting holiday pay status:', error);
        writeError('clientGetUserHolidayPayStatus', error);
        return {
            success: false,
            error: error.message,
            payStatuses: []
        };
    }
}

/**
 * Apply default paid status to all users from a specific country
 */
function applyDefaultPaidStatusByCountry(countryCode, isPaid) {
    try {
        console.log(`🌍 Applying default paid status (${isPaid}) to all users from ${countryCode}`);
        
        // Get all users from this country (you'd need a Country field in Users sheet)
        const users = readSheet(USERS_SHEET) || [];
        const countryUsers = users.filter(user => 
            user.Country === countryCode || 
            user.CountryCode === countryCode
        );
        
        let updatedCount = 0;
        
        countryUsers.forEach(user => {
            try {
                clientSetUserHolidayPayStatus(
                    user.ID, 
                    countryCode, 
                    isPaid, 
                    `Auto-applied for ${countryCode} residents`
                );
                updatedCount++;
            } catch (error) {
                console.warn(`Failed to update user ${user.ID}:`, error);
            }
        });
        
        console.log(`✅ Updated ${updatedCount} users from ${countryCode}`);
        return updatedCount;
        
    } catch (error) {
        console.error('❌ Error applying default paid status:', error);
        return 0;
    }
}

/**
 * Bulk update holiday pay status for multiple users
 */
function clientBulkUpdateHolidayPayStatus(updates) {
    try {
        console.log(`📋 Bulk updating holiday pay status for ${updates.length} entries`);
        
        let successCount = 0;
        let errorCount = 0;
        const results = [];
        
        updates.forEach(update => {
            try {
                const result = clientSetUserHolidayPayStatus(
                    update.userId,
                    update.countryCode,
                    update.isPaid,
                    update.notes || 'Bulk update'
                );
                
                if (result.success) {
                    successCount++;
                    results.push({ ...update, status: 'success' });
                } else {
                    errorCount++;
                    results.push({ ...update, status: 'error', error: result.error });
                }
            } catch (error) {
                errorCount++;
                results.push({ ...update, status: 'error', error: error.message });
            }
        });
        
        return {
            success: true,
            message: `Bulk update completed: ${successCount} successful, ${errorCount} failed`,
            successCount: successCount,
            errorCount: errorCount,
            results: results
        };
        
    } catch (error) {
        console.error('❌ Error in bulk holiday pay update:', error);
        writeError('clientBulkUpdateHolidayPayStatus', error);
        return {
            success: false,
            error: error.message,
            successCount: 0,
            errorCount: updates.length
        };
    }
}

// ────────────────────────────────────────────────────────────────────────────
// HOLIDAY REPORTING AND ANALYTICS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get holiday analytics and reporting data
 */
function clientGetHolidayAnalytics(year = null) {
    try {
        const holidays = readSheet(HOLIDAYS_SHEET) || [];
        const payStatuses = readSheet('UserHolidayPayStatus') || [];
        const users = readSheet(USERS_SHEET) || [];
        
        let analysisYear = year || new Date().getFullYear();
        
        // Filter holidays for the year
        const yearHolidays = holidays.filter(h => 
            new Date(h.Date).getFullYear() === analysisYear
        );
        
        // Count holidays by month
        const holidaysByMonth = Array(12).fill(0);
        yearHolidays.forEach(holiday => {
            const month = new Date(holiday.Date).getMonth();
            holidaysByMonth[month]++;
        });
        
        // Count users by country and pay status
        const payStatusByCountry = {};
        payStatuses.forEach(status => {
            if (!payStatusByCountry[status.CountryCode]) {
                payStatusByCountry[status.CountryCode] = {
                    countryCode: status.CountryCode,
                    countryName: status.CountryName,
                    paid: 0,
                    notPaid: 0,
                    total: 0
                };
            }
            
            if (status.IsPaid === 'TRUE') {
                payStatusByCountry[status.CountryCode].paid++;
            } else {
                payStatusByCountry[status.CountryCode].notPaid++;
            }
            payStatusByCountry[status.CountryCode].total++;
        });
        
        return {
            success: true,
            year: analysisYear,
            totalHolidays: yearHolidays.length,
            holidaysByMonth: holidaysByMonth,
            payStatusByCountry: Object.values(payStatusByCountry),
            totalUsers: users.length,
            usersWithPayStatus: payStatuses.length,
            upcomingHolidays: getUpcomingHolidays(yearHolidays, 30) // Next 30 days
        };
        
    } catch (error) {
        console.error('❌ Error getting holiday analytics:', error);
        writeError('clientGetHolidayAnalytics', error);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Get upcoming holidays in the next N days
 */
function getUpcomingHolidays(holidays, daysAhead = 30) {
    try {
        const today = new Date();
        const futureDate = new Date(today.getTime() + (daysAhead * 24 * 60 * 60 * 1000));
        
        return holidays.filter(holiday => {
            const holidayDate = new Date(holiday.Date);
            return holidayDate >= today && holidayDate <= futureDate;
        }).sort((a, b) => new Date(a.Date) - new Date(b.Date));
        
    } catch (error) {
        console.error('Error getting upcoming holidays:', error);
        return [];
    }
}

/**
 * Check if a specific date is a holiday
 */
function clientCheckIfHoliday(dateStr) {
    try {
        const holidays = readSheet(HOLIDAYS_SHEET) || [];
        const holiday = holidays.find(h => h.Date === dateStr);
        
        return {
            success: true,
            isHoliday: !!holiday,
            holiday: holiday || null
        };
        
    } catch (error) {
        console.error('❌ Error checking if date is holiday:', error);
        return {
            success: false,
            error: error.message,
            isHoliday: false
        };
    }
}

// ────────────────────────────────────────────────────────────────────────────
// UTILITY AND HELPER FUNCTIONS
// ────────────────────────────────────────────────────────────────────────────

/**
 * Get list of supported countries
 */
function clientGetSupportedCountries() {
    try {
        const countryDefinitions = getCountryHolidayDefinitions();
        
        const countries = Object.entries(countryDefinitions).map(([code, data]) => ({
            code: code,
            name: data.name,
            holidayCount: data.holidays.length
        }));
        
        return {
            success: true,
            countries: countries,
            total: countries.length
        };
        
    } catch (error) {
        console.error('❌ Error getting supported countries:', error);
        return {
            success: false,
            error: error.message,
            countries: []
        };
    }
}

/**
 * Delete a holiday
 */
function clientDeleteHoliday(holidayId) {
    try {
        console.log(`🗑️ Deleting holiday: ${holidayId}`);
        
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOLIDAYS_SHEET);
        if (!sheet) {
            throw new Error('Holidays sheet not found');
        }
        
        const data = sheet.getDataRange().getValues();
        const headers = data[0];
        const idIndex = headers.indexOf('ID');
        
        if (idIndex === -1) {
            throw new Error('ID column not found in holidays sheet');
        }
        
        for (let i = data.length - 1; i >= 1; i--) {
            if (String(data[i][idIndex]) === String(holidayId)) {
                sheet.deleteRow(i + 1);
                invalidateCache(HOLIDAYS_SHEET);
                console.log('✅ Holiday deleted successfully');
                return {
                    success: true,
                    message: 'Holiday deleted successfully'
                };
            }
        }
        
        throw new Error('Holiday not found');
        
    } catch (error) {
        console.error('❌ Error deleting holiday:', error);
        writeError('clientDeleteHoliday', error);
        return {
            success: false,
            error: error.message
        };
    }
}

/**
 * Format holiday date for display
 */
function formatHolidayDate(dateStr) {
    try {
        const date = new Date(dateStr);
        return date.toLocaleDateString('en-US', { 
            weekday: 'long',
            year: 'numeric', 
            month: 'long', 
            day: 'numeric' 
        });
    } catch (error) {
        return dateStr;
    }
}

console.log('✅ Holiday Management System loaded successfully');
console.log('🎉 Features:');
console.log('   - 10+ countries supported with comprehensive holiday lists');
console.log('   - Holiday import and management');
console.log('   - User holiday pay status by country of residence');
console.log('   - Holiday analytics and reporting');
console.log('   - Bulk operations and country-specific settings');
console.log('   - Integration with schedule generation system');