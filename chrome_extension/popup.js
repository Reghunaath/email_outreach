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
