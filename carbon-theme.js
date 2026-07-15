/* Shared light/dark theme controller for the MaaS360 Sales Guide.
   Light is the default (per IBM brand); the choice persists in localStorage.
   Include this in <head> so the theme is applied before first paint. */
(function () {
  var KEY = 'maas-theme';
  var root = document.documentElement;

  function stored() {
    try { return localStorage.getItem(KEY); } catch (e) { return null; }
  }
  function apply(theme) {
    root.setAttribute('data-theme', theme === 'dark' ? 'dark' : 'light');
  }

  // Apply as early as possible (default: light).
  apply(stored() || 'light');

  function label(btn, theme) {
    var t = btn.querySelector('.tt-label');
    if (t) t.textContent = theme === 'dark' ? 'Light' : 'Dark';
    btn.setAttribute('aria-pressed', theme === 'dark' ? 'true' : 'false');
    btn.setAttribute('title', theme === 'dark' ? 'Switch to light theme' : 'Switch to dark theme');
  }

  function wire() {
    var current = root.getAttribute('data-theme') || 'light';
    var btns = document.querySelectorAll('.theme-toggle');
    btns.forEach(function (btn) {
      label(btn, current);
      btn.addEventListener('click', function () {
        current = (root.getAttribute('data-theme') === 'dark') ? 'light' : 'dark';
        apply(current);
        try { localStorage.setItem(KEY, current); } catch (e) {}
        document.querySelectorAll('.theme-toggle').forEach(function (b) { label(b, current); });
      });
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wire);
  } else {
    wire();
  }
})();
