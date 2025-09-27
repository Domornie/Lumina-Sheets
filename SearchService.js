// ─────────────────────────────────────────────────────────────────────────────
// SearchService.gs
// Implements server-side search against Google’s Custom Search JSON API.
// ─────────────────────────────────────────────────────────────────────────────

// Replace with your actual CSE engine ID (the "cx=" value) and your API key.
// Use var with a guard so the identifiers aren't redeclared if this file is
// loaded alongside other scripts that define the same configuration (e.g.
// Code.js). Google Apps Script evaluates all top-level statements, so
// re-declaring a const would raise "Identifier has already been declared".
// eslint-disable-next-line no-var
var CSE_ID = typeof CSE_ID !== 'undefined' ? CSE_ID : '130aba31c8a2d439c';
// eslint-disable-next-line no-var
var API_KEY = typeof API_KEY !== 'undefined' ? API_KEY : 'AIzaSyAg-puM5l9iQpjz_NplMJaKbUNRH7ld7sY';

/**
 * Performs a web search via Google Custom Search JSON API.
 *
 * @param {string} query      The search query string.
 * @param {number=} startIndex Optional: 1-based index of first result to return (for paging).
 * @return {Object} Parsed JSON response from Google.
 * @throws If the HTTP response code is not 200.
 */
function searchWeb(query, startIndex) {
  if (!query || typeof query !== 'string') {
    throw new Error('Invalid search query.');
  }

  const baseUrl = 'https://www.googleapis.com/customsearch/v1';
  const params  = [
    `key=${API_KEY}`,
    `cx=${CSE_ID}`,
    `q=${encodeURIComponent(query)}`,
    startIndex ? `start=${startIndex}` : ''
  ].filter(Boolean).join('&');
  const url = `${baseUrl}?${params}`;

  // Fetch the JSON
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code !== 200) {
    throw new Error(`Search API error [${code}]: ${text}`);
  }

  return JSON.parse(text);
}
