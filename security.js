// --- SECURITY SCRIPT ---
// 1. Right Click Block
document.addEventListener('contextmenu', function(e) {
    e.preventDefault();
});

// 2. Keyboard Shortcuts Block (Ctrl+U, F12, Ctrl+Shift+I, etc.)
document.addEventListener('keydown', function(e) {
    // Ctrl+U (View Source)
    if (e.ctrlKey && (e.key === 'u' || e.key === 'U')) {
        e.preventDefault();
        return false;
    }
    // F12 (Developer Tools)
    if (e.key === 'F12') {
        e.preventDefault();
        return false;
    }
    // Ctrl+Shift+I (Developer Tools)
    if (e.ctrlKey && e.shiftKey && (e.key === 'I' || e.key === 'i')) {
        e.preventDefault();
        return false;
    }
    // Ctrl+Shift+J (Console)
    if (e.ctrlKey && e.shiftKey && (e.key === 'J' || e.key === 'j')) {
        e.preventDefault();
        return false;
    }
    // Ctrl+Shift+C (Inspect Element)
    if (e.ctrlKey && e.shiftKey && (e.key === 'C' || e.key === 'c')) {
        e.preventDefault();
        return false;
    }
    // Ctrl+S (Save Page)
    if (e.ctrlKey && (e.key === 's' || e.key === 'S')) {
        e.preventDefault();
        return false;
    }
});

// --- DARK / LIGHT MODE TOGGLE ---
(function () {
    var KEY = 'appTheme';
    function getSaved() { try { return localStorage.getItem(KEY); } catch (e) { return null; } }
    function save(t) { try { localStorage.setItem(KEY, t); } catch (e) {} }

    // Default: dark (change 'dark' to 'light' for default bright)
    var theme = getSaved() || 'dark';
    document.documentElement.setAttribute('data-theme', theme);

    function label(btn) {
        var dark = document.documentElement.getAttribute('data-theme') === 'dark';
        btn.innerHTML = dark
            ? '<i class="fa-solid fa-sun"></i> Bright Mode'
            : '<i class="fa-solid fa-moon"></i> Dark Mode';
        btn.title = dark ? 'Switch to Bright mode' : 'Switch to Dark mode';
    }

    function toggle(btn) {
        var next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        document.documentElement.setAttribute('data-theme', next);
        save(next);
        label(btn);
    }

    function init() {
        if (document.getElementById('themeToggleBtn')) return;
        var btn = document.createElement('button');
        btn.id = 'themeToggleBtn';
        btn.type = 'button';
        btn.className = 'theme-toggle';
        var host = document.querySelector('.header-actions');
        if (host) {
            host.insertBefore(btn, host.firstChild);
        } else {
            btn.classList.add('floating');
            document.body.appendChild(btn);
        }
        label(btn);
        btn.addEventListener('click', function () { toggle(btn); });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
