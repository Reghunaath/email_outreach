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
  const profileTabs = tabs
    .map(t => {
      const url = t.url || '';
      if (!/linkedin\.com\/in\//.test(url)) return null;
      const match = url.match(/(https?:\/\/(?:www\.)?linkedin\.com\/in\/[^/?#]+)/);
      return match ? { id: t.id, url: match[1] } : null;
    })
    .filter(Boolean);

  const urls = profileTabs.map(t => t.url);

  if (urls.length === 0) {
    status.textContent = 'No LinkedIn profile tabs found.';
    status.className = 'error';
    btn.disabled = false;
    return;
  }

  await navigator.clipboard.writeText(urls.join('\n'));

  const tabIds = profileTabs.map(t => t.id).filter(id => id !== undefined);
  if (tabIds.length > 0) {
    await chrome.tabs.remove(tabIds);
  }

  status.textContent = `Copied ${urls.length} URL${urls.length !== 1 ? 's' : ''} and closed ${tabIds.length} tab${tabIds.length !== 1 ? 's' : ''}.`;
  btn.disabled = false;
});

// --- Manual AI answer generator (for sites AI fill can't reach) -------------

document.getElementById('generateBtn').addEventListener('click', async () => {
  const genBtn = document.getElementById('generateBtn');
  const aiStatus = document.getElementById('aiStatus');
  const resultField = document.getElementById('aiResultField');
  const resultBox = document.getElementById('aiResult');

  const question = document.getElementById('aiQuestion').value.trim();
  const jd = document.getElementById('aiJd').value.trim();
  const context = document.getElementById('aiContext').value.trim();

  const setStatus = (text, isError) => {
    aiStatus.textContent = text;
    aiStatus.className = isError ? 'error' : '';
  };

  if (!question) {
    setStatus('Enter a question first.', true);
    return;
  }
  if (!chrome.runtime || !chrome.runtime.id) {
    setStatus('Extension was reloaded. Reopen this popup.', true);
    return;
  }

  genBtn.disabled = true;
  genBtn.textContent = 'Generating…';
  setStatus('', false);

  try {
    const res = await chrome.runtime.sendMessage({
      type: 'ASHBY_FILL',
      question,
      board: '',
      postingId: '',
      pageText: jd,
      remark: context,
    });
    if (res && res.text) {
      resultBox.value = res.text;
      resultField.style.display = 'block';
      try {
        await navigator.clipboard.writeText(res.text);
        setStatus('Answer copied to clipboard.', false);
      } catch (_) {
        setStatus('Answer generated (copy failed, select it manually).', true);
      }
    } else {
      setStatus((res && res.error) || 'No response from the extension.', true);
    }
  } catch (err) {
    setStatus(err.message || 'Unexpected error.', true);
  } finally {
    genBtn.disabled = false;
    genBtn.textContent = 'Generate answer';
  }
});

scanForKeywords();
