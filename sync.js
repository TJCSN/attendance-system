/* Shared sync helper for all Seahawk pages.
 * Talks to the Google Apps Script web app whose URL is stored in roster.json.
 * Usage: <script src="sync.js"></script> then SeahawkSync.fetchAll() / .post() / .poll()
 */
(function () {
  var cachedUrl = null;

  async function getUrl() {
    if (cachedUrl) return cachedUrl;
    try {
      var r = await fetch('roster.json', { cache: 'no-store' });
      if (r.ok) {
        var d = await r.json();
        if (d.googleSheetsApiUrl) cachedUrl = d.googleSheetsApiUrl;
      }
    } catch (e) { /* fall through */ }
    if (!cachedUrl) {
      try { cachedUrl = localStorage.getItem('googleSheetsApiUrl') || null; } catch (e) {}
    }
    return cachedUrl;
  }

  // Returns { availability: [], attendance: [], config: {}, roster: [] }
  async function fetchAll() {
    var url = await getUrl();
    if (!url) throw new Error('Google Sheets URL is not configured');
    var res = await fetch(url + (url.indexOf('?') > -1 ? '&' : '?') + 't=' + Date.now());
    if (!res.ok) throw new Error('HTTP ' + res.status);
    var j = await res.json();
    if (!j.success) throw new Error(j.error || 'Server error');
    return { availability: j.data || [], attendance: j.attendance || [], config: j.config || {}, roster: j.roster || [] };
  }

  // payload must include "action": "upsert" | "attendance" | "config"
  async function post(payload) {
    var url = await getUrl();
    if (!url) throw new Error('Google Sheets URL is not configured');
    var res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // avoids CORS preflight
      body: JSON.stringify(payload)
    });
    var j = await res.json();
    if (!j.success) throw new Error(j.error || 'Server error');
    return j;
  }

  // Runs fn now, then every `ms` while the tab is visible, and again when the tab regains focus.
  function poll(fn, ms) {
    fn();
    setInterval(function () { if (!document.hidden) fn(); }, ms || 30000);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) fn(); });
  }

  window.SeahawkSync = { getUrl: getUrl, fetchAll: fetchAll, post: post, poll: poll };
})();
