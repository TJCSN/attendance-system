/* Shared sync helper for all Seahawk pages.
 * Talks to the Google Apps Script web app whose URL is stored in roster.json.
 *   SeahawkSync.fetchAll({force, all})  -> { availability, attendance, config, roster }
 *   SeahawkSync.post(payload)           -> server reply (payload needs "action")
 *   SeahawkSync.poll(fn, ms)            -> run fn now and every ms while the tab is visible
 *   SeahawkSync.loadXLSX()              -> loads the Excel library only when needed
 * Speed: identical requests made at the same moment share one network call, and results are
 * reused for 8 seconds. force:true skips that. all:true also returns older weeks.
 */
(function () {
  var urlPromise = null;
  var inflight = {};
  var cache = {};
  var TTL = 8000;

  function getUrl() {
    if (!urlPromise) {
      urlPromise = (async function () {
        var u = null;
        try {
          var r = await fetch('roster.json', { cache: 'no-store' });
          if (r.ok) {
            var d = await r.json();
            if (d.googleSheetsApiUrl) u = d.googleSheetsApiUrl;
          }
        } catch (e) { /* fall through */ }
        if (!u) {
          try { u = localStorage.getItem('googleSheetsApiUrl') || null; } catch (e) {}
        }
        return u;
      })();
    }
    return urlPromise;
  }

  async function fetchAll(opts) {
    opts = opts || {};
    var key = opts.all ? 'all' : 'recent';
    if (!opts.force && cache[key] && Date.now() - cache[key].time < TTL) return cache[key].data;
    if (!opts.force && inflight[key]) return inflight[key];

    var p = (async function () {
      var url = await getUrl();
      if (!url) throw new Error('Google Sheets URL is not configured');
      var q = 't=' + Date.now() + (opts.all ? '&all=1' : '');
      var res = await fetch(url + (url.indexOf('?') > -1 ? '&' : '?') + q);
      if (!res.ok) throw new Error('HTTP ' + res.status);
      var j = await res.json();
      if (!j.success) throw new Error(j.error || 'Server error');
      var data = { availability: j.data || [], attendance: j.attendance || [], config: j.config || {}, roster: j.roster || [] };
      cache[key] = { data: data, time: Date.now() };
      return data;
    })();
    inflight[key] = p;
    try { return await p; } finally { delete inflight[key]; }
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
    cache = {}; // data changed, so never reuse older results
    return j;
  }

  function poll(fn, ms) {
    fn();
    setInterval(function () { if (!document.hidden) fn(); }, ms || 30000);
    document.addEventListener('visibilitychange', function () { if (!document.hidden) fn(); });
  }

  var xlsxPromise = null;
  function loadXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxPromise) {
      xlsxPromise = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
        s.onload = function () { resolve(window.XLSX); };
        s.onerror = function () { xlsxPromise = null; reject(new Error('Could not load the Excel library')); };
        document.head.appendChild(s);
      });
    }
    return xlsxPromise;
  }

  window.SeahawkSync = { getUrl: getUrl, fetchAll: fetchAll, post: post, poll: poll, loadXLSX: loadXLSX };
})();
