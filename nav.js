/* Navigation bar for the staff pages (coach view, attendance, report, setup). */
(function () {
  var links = [
    ['coach-view-enhanced.html', 'Coach view'],
    ['mark-attendance.html', 'Mark attendance'],
    ['attendance-report.html', 'Report'],
    ['setup-hub.html', 'Setup']
  ];
  var here = (location.pathname.split('/').pop() || '').toLowerCase();
  var nav = document.createElement('nav');
  nav.className = 'sh-nav';
  nav.setAttribute('aria-label', 'Staff pages');
  nav.innerHTML = '<span class="brand">Seahawk training</span>' +
    links.map(function (l) {
      return '<a href="' + l[0] + '"' + (l[0] === here ? ' aria-current="page"' : '') + '>' + l[1] + '</a>';
    }).join('') +
    '<span class="spacer"></span><a href="weekly-availability.html">Member form</a>';
  document.body.insertBefore(nav, document.body.firstChild);
})();
