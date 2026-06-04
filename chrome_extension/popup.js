const ALERT_KEYWORDS = ['sponsorship', 'citizen', 'citizenship'];

async function scanForKeywords() {
  const warningsDiv = document.getElementById('warnings');
  warningsDiv.innerHTML = '';

  const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
  if (tabs.length === 0) return;

  for (const tab of tabs) {
    let matched = [];
    try {
      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: (keywords) => {
          const text = document.body.innerText.toLowerCase();
          return keywords.filter(kw => text.includes(kw));
        },
        args: [ALERT_KEYWORDS],
      });
      matched = results[0]?.result || [];
    } catch (_) {
      // Tab not yet loaded or inaccessible — skip
    }

    if (matched.length > 0) {
      const name = (tab.title || tab.url).replace(/\s*[|\-–].*$/, '').trim();
      const div = document.createElement('div');
      div.className = 'warning';
      div.textContent = `⚠ ${name}: ${matched.join(', ')}`;
      warningsDiv.appendChild(div);
    }
  }
}

document.getElementById('copyBtn').addEventListener('click', async () => {
  const btn = document.getElementById('copyBtn');
  const status = document.getElementById('status');

  btn.disabled = true;
  status.textContent = '';
  status.className = '';

  const tabs = await chrome.tabs.query({});
  const urls = tabs
    .map(t => t.url || '')
    .filter(url => /linkedin\.com\/in\//.test(url))
    .map(url => {
      const match = url.match(/(https?:\/\/(?:www\.)?linkedin\.com\/in\/[^/?#]+)/);
      return match ? match[1] : null;
    })
    .filter(Boolean);

  if (urls.length === 0) {
    status.textContent = 'No LinkedIn profile tabs found.';
    status.className = 'error';
    btn.disabled = false;
    return;
  }

  await navigator.clipboard.writeText(urls.join('\n'));
  status.textContent = `Copied ${urls.length} URL${urls.length !== 1 ? 's' : ''}.`;
  btn.disabled = false;
});

scanForKeywords();
