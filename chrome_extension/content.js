(function () {
  const KEYWORDS = ['sponsorship', 'citizen', 'citizenship'];
  let lastUrl = location.href;
  let activeBanner = null;
  let lastFoundKey = '';
  let debounceTimer = null;

  function showBanner(found) {
    const banner = document.createElement('div');
    banner.style.cssText = [
      'position:fixed', 'top:20px', 'right:20px', 'z-index:2147483647',
      'background:#1e1e2e', 'color:#fab387',
      'border:1px solid rgba(250,179,135,0.45)',
      'border-radius:8px', 'padding:12px 36px 12px 14px',
      'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
      'font-size:13px', 'line-height:1.5', 'max-width:280px',
      'box-shadow:0 4px 16px rgba(0,0,0,0.5)',
    ].join(';');

    banner.innerHTML = `<strong style="display:block;margin-bottom:3px;color:#cdd6f4">⚠ Keyword Alert</strong>${found.join(', ')}`;

    const closeBtn = document.createElement('span');
    closeBtn.textContent = '×';
    closeBtn.style.cssText = 'position:absolute;top:7px;right:11px;cursor:pointer;font-size:17px;color:#6c7086;line-height:1';
    closeBtn.onmouseenter = () => { closeBtn.style.color = '#cdd6f4'; };
    closeBtn.onmouseleave = () => { closeBtn.style.color = '#6c7086'; };
    closeBtn.onclick = () => { banner.remove(); activeBanner = null; };
    banner.appendChild(closeBtn);

    document.body.appendChild(banner);
    setTimeout(() => { banner.remove(); if (activeBanner === banner) activeBanner = null; }, 10000);
    return banner;
  }

  function check() {
    const currentUrl = location.href;

    if (currentUrl !== lastUrl) {
      lastUrl = currentUrl;
      activeBanner?.remove();
      activeBanner = null;
      lastFoundKey = '';
    }

    const text = document.body.innerText.toLowerCase();
    const found = KEYWORDS.filter(kw => text.includes(kw));
    if (found.length === 0) return;

    const key = currentUrl + '|' + found.join(',');
    if (key === lastFoundKey) return;
    lastFoundKey = key;

    activeBanner?.remove();
    activeBanner = showBanner(found);
  }

  check();

  new MutationObserver(() => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(check, 800);
  }).observe(document.body, { childList: true, subtree: true });
})();
