const DEFAULT_MODEL = 'gemini-flash-latest';

const loadingEl = document.getElementById('loading');
const formEl = document.getElementById('form');
const apiKeyEl = document.getElementById('apiKey');
const modelEl = document.getElementById('model');
const profileEl = document.getElementById('profile');
const saveBtn = document.getElementById('saveBtn');
const statusEl = document.getElementById('status');

function setStatus(text, kind) {
  statusEl.textContent = text;
  statusEl.className = kind || '';
}

async function load() {
  try {
    const { apiKey = '', model = '', profile = '' } =
      await chrome.storage.local.get(['apiKey', 'model', 'profile']);
    apiKeyEl.value = apiKey;
    modelEl.value = model || DEFAULT_MODEL;
    profileEl.value = profile;
    loadingEl.style.display = 'none';
    formEl.style.display = 'block';
    if (!apiKey) {
      setStatus('Add your Gemini API key to enable AI fill.', 'muted');
    }
  } catch (err) {
    loadingEl.textContent = 'Could not load settings: ' + err.message;
  }
}

async function save() {
  saveBtn.disabled = true;
  setStatus('Saving…', 'muted');
  const apiKey = apiKeyEl.value.trim();
  const model = modelEl.value.trim() || DEFAULT_MODEL;
  const profile = profileEl.value.trim();

  try {
    await chrome.storage.local.set({ apiKey, model, profile });
    if (!apiKey) {
      setStatus('Saved — but no API key set, so AI fill won’t work yet.', 'error');
    } else {
      setStatus('Saved.', '');
    }
  } catch (err) {
    setStatus('Save failed: ' + err.message, 'error');
  } finally {
    saveBtn.disabled = false;
  }
}

saveBtn.addEventListener('click', save);
load();
