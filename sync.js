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

  // Monday-Sunday of the current week as local YYYY-MM-DD strings (the same week the member form uses)
  function currentWeekRange() {
    function pad(n) { return String(n).padStart(2, '0'); }
    function ymd(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
    var now = new Date();
    var dow = now.getDay();
    var mon = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (dow === 0 ? 6 : dow - 1));
    var sun = new Date(mon.getFullYear(), mon.getMonth(), mon.getDate() + 6);
    return { start: ymd(mon), end: ymd(sun), today: ymd(now) };
  }

  // ---- Adding / removing a member on one training day on their behalf (coach + attendance pages) ----
  // The day is saved as a normal sign-up flagged addedByCoach, so every page shows it.

  // The Tuesday date the member form saves as "weekStart" (calculated exactly like the form does)
  function currentWeekStart() {
    var today = new Date();
    var cd = today.getDay();
    var thisMonday = new Date(today);
    thisMonday.setDate(today.getDate() + (cd === 0 ? -6 : -(cd - 1)));
    var tue = new Date(thisMonday);
    tue.setDate(thisMonday.getDate() + 1);
    return tue.getFullYear() + '-' + String(tue.getMonth() + 1).padStart(2, '0') + '-' + String(tue.getDate()).padStart(2, '0');
  }

  function isoWeekNumber(date) {
    var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    var dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  }

  // This week's sign-up for a member from a list of responses (or undefined)
  function findThisWeek(list, name) {
    var ws = currentWeekStart();
    var n = String(name || '').trim().toLowerCase();
    return (list || []).find(function (r) {
      return String(r.name || '').trim().toLowerCase() === n && String(r.weekStart) === ws;
    });
  }

  function saveMemberDays(person, existing, days) {
    var weekStart = currentWeekStart();
    return post({
      action: 'upsert',
      name: person.name,
      category: person.category || '',
      paddlingSide: person.paddlingSide || '',
      gender: person.gender || '',
      days: days,
      weekStart: weekStart,
      weekNumber: existing && existing.weekNumber ? existing.weekNumber : isoWeekNumber(new Date(weekStart))
    });
  }

  // person: {name, category, paddlingSide, gender}; existing: their sign-up this week (or undefined);
  // dayInfo: {date, name, display}. Keeps any days they already signed up for.
  function addCoachDay(person, existing, dayInfo) {
    var days = (existing ? existing.days : []).filter(function (d) { return d.date !== dayInfo.date; });
    days.push({ date: dayInfo.date, name: dayInfo.name, display: dayInfo.display, addedByCoach: true });
    return saveMemberDays(person, existing, days);
  }

  // Removes only a day that the coach added; days the member signed up for themselves are kept.
  function removeCoachDay(person, existing, date) {
    var days = (existing ? existing.days : []).filter(function (d) { return !(d.date === date && d.addedByCoach); });
    return saveMemberDays(person, existing, days);
  }

  window.SeahawkSync = {
    getUrl: getUrl, fetchAll: fetchAll, post: post, poll: poll, loadXLSX: loadXLSX,
    currentWeekStart: currentWeekStart, currentWeekRange: currentWeekRange, findThisWeek: findThisWeek, addCoachDay: addCoachDay, removeCoachDay: removeCoachDay
  };
})();
