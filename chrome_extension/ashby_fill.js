(function () {
  const MARK = 'data-aifill';
  const JD_MAX = 6000;
  const IS_ASHBY = /(^|\.)ashbyhq\.com$/i.test(location.hostname);
  const IS_GREENHOUSE = /(^|\.)greenhouse\.io$/i.test(location.hostname);
  let debounceTimer = null;

  // --- Field detection -----------------------------------------------------

  function isVisible(el) {
    if (!el.offsetParent && el.offsetWidth === 0 && el.offsetHeight === 0) return false;
    const style = getComputedStyle(el);
    return style.visibility !== 'hidden' && style.display !== 'none';
  }

  function isFillable(el) {
    if (el.disabled || el.readOnly) return false;
    if (el.getAttribute('aria-hidden') === 'true') return false;
    if (el.getAttribute('role') === 'combobox') return false; // react-select dropdowns
    if (!isVisible(el)) return false;
    if (el.tagName === 'TEXTAREA') return true;
    if (el.tagName === 'INPUT') {
      const type = (el.getAttribute('type') || 'text').toLowerCase();
      return type === 'text'; // String fields; skip email/date/file/number/checkbox/tel/search/etc.
    }
    return false;
  }

  // --- Question resolution -------------------------------------------------

  function textFromIds(ids) {
    return ids
      .split(/\s+/)
      .map((id) => document.getElementById(id))
      .filter(Boolean)
      .map((n) => n.textContent.trim())
      .join(' ')
      .trim();
  }

  function getQuestion(el) {
    const labelledby = el.getAttribute('aria-labelledby');
    if (labelledby) {
      const t = textFromIds(labelledby);
      if (t) return t;
    }
    const ariaLabel = el.getAttribute('aria-label');
    if (ariaLabel && ariaLabel.trim()) return ariaLabel.trim();

    if (el.id) {
      const forLabel = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
      if (forLabel && forLabel.textContent.trim()) return forLabel.textContent.trim();
    }

    // Walk up a few container levels looking for a label/legend/heading.
    let node = el;
    for (let depth = 0; depth < 4 && node; depth++) {
      node = node.parentElement;
      if (!node) break;
      const candidate = node.querySelector('label, legend, h1, h2, h3, h4');
      if (candidate && candidate.textContent.trim()) return candidate.textContent.trim();
    }

    const placeholder = el.getAttribute('placeholder');
    if (placeholder && placeholder.trim()) return placeholder.trim();

    return '';
  }

  // --- Posting reference (for JD lookup via Ashby's posting API) -----------

  // URL is /{board}/{postingId}[/application]; board == posting-API job-board name.
  function getPostingRef() {
    const parts = location.pathname.split('/').filter(Boolean);
    const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    return {
      board: parts[0] || '',
      postingId: parts.find((p) => uuid.test(p)) || '',
    };
  }

  // Job-description text sent with each request.
  // - Ashby: this is a fallback only; the background worker normally fetches the
  //   JD from Ashby's posting API and uses this page text only if that fails.
  // - Greenhouse: the JD lives on the page, so we scrape the title + description
  //   here and it becomes the actual JD (no API lookup on Greenhouse).
  function getJobDescriptionText() {
    if (IS_GREENHOUSE) {
      const title = (document.querySelector('h1') || {}).innerText || '';
      const desc = (document.querySelector('.job__description') || {}).innerText || '';
      const combined = `${title}\n\n${desc}`.trim();
      if (combined) return combined.slice(0, JD_MAX);
    }
    return (document.body.innerText || '').trim().slice(0, JD_MAX);
  }

  // --- React-aware field fill ----------------------------------------------

  function setNativeValue(el, value) {
    const proto = el.tagName === 'TEXTAREA'
      ? window.HTMLTextAreaElement.prototype
      : window.HTMLInputElement.prototype;
    const setter = Object.getOwnPropertyDescriptor(proto, 'value')?.set;
    if (setter) {
      setter.call(el, value);
    } else {
      el.value = value;
    }
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // --- Button --------------------------------------------------------------

  function attachButton(field) {
    if (field.getAttribute(MARK)) return;
    if (field.nextElementSibling && field.nextElementSibling.hasAttribute('data-aifill-ctrl')) return;
    field.setAttribute(MARK, '1');

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = '✨ AI fill';
    btn.style.cssText = [
      'display:inline-flex', 'align-items:center', 'gap:4px',
      'margin:4px 0', 'padding:3px 9px',
      'font:500 12px -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
      'color:#1e1e2e', 'background:#89b4fa', 'border:none', 'border-radius:6px',
      'cursor:pointer', 'transition:background 0.15s',
    ].join(';');

    const setState = (state, label) => {
      btn.dataset.state = state;
      if (state === 'loading') {
        btn.disabled = true;
        btn.textContent = '⏳ Filling…';
        btn.style.background = '#45475a';
        btn.style.color = '#6c7086';
        btn.title = '';
      } else if (state === 'error') {
        btn.disabled = false;
        btn.textContent = '⚠ Retry';
        btn.style.background = '#f38ba8';
        btn.style.color = '#1e1e2e';
        btn.title = label || 'Error';
      } else {
        btn.disabled = false;
        btn.textContent = '✨ AI fill';
        btn.style.background = '#89b4fa';
        btn.style.color = '#1e1e2e';
        btn.title = '';
      }
    };

    btn.addEventListener('mouseenter', () => {
      if (btn.dataset.state !== 'loading') {
        btn.style.background = btn.dataset.state === 'error' ? '#eba0ac' : '#74c7ec';
      }
    });
    btn.addEventListener('mouseleave', () => {
      if (btn.dataset.state !== 'loading') {
        btn.style.background = btn.dataset.state === 'error' ? '#f38ba8' : '#89b4fa';
      }
    });
    btn.addEventListener('focus', () => {
      btn.style.outline = '2px solid #cdd6f4';
      btn.style.outlineOffset = '2px';
    });
    btn.addEventListener('blur', () => { btn.style.outline = 'none'; });

    // Second button + a toggleable box for per-field context/remarks.
    const ctxBtn = document.createElement('button');
    ctxBtn.type = 'button';
    ctxBtn.textContent = '💬 Add context';
    ctxBtn.style.cssText = [
      'display:inline-flex', 'align-items:center', 'gap:4px',
      'margin:4px 0', 'padding:3px 9px',
      'font:500 12px -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
      'color:#cdd6f4', 'background:#313244', 'border:none', 'border-radius:6px',
      'cursor:pointer', 'transition:background 0.15s',
    ].join(';');
    ctxBtn.addEventListener('mouseenter', () => { ctxBtn.style.background = '#45475a'; });
    ctxBtn.addEventListener('mouseleave', () => { ctxBtn.style.background = '#313244'; });
    ctxBtn.addEventListener('focus', () => {
      ctxBtn.style.outline = '2px solid #cdd6f4';
      ctxBtn.style.outlineOffset = '2px';
    });
    ctxBtn.addEventListener('blur', () => { ctxBtn.style.outline = 'none'; });

    const remarkBox = document.createElement('textarea');
    remarkBox.rows = 2;
    remarkBox.placeholder = 'Add context, then press Enter to fill (Shift+Enter for a new line)…';
    remarkBox.style.cssText = [
      'display:none', 'width:340px', 'max-width:100%', 'margin:0 0 4px',
      'padding:6px 8px',
      'font:400 12px -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif',
      'color:#cdd6f4', 'background:#181825', 'border:1px solid #313244',
      'border-radius:6px', 'resize:vertical',
    ].join(';');
    remarkBox.addEventListener('focus', () => { remarkBox.style.borderColor = '#89b4fa'; });
    remarkBox.addEventListener('blur', () => { remarkBox.style.borderColor = '#313244'; });

    async function runFill() {
      const question = getQuestion(field);
      if (!question) {
        setState('error', 'Could not find a question label for this field.');
        return;
      }
      if (btn.dataset.state === 'loading') return; // already in flight
      if (!chrome.runtime || !chrome.runtime.id) {
        setState('error', 'Extension was reloaded. Refresh this page and try again.');
        return;
      }
      setState('loading');
      try {
        const { board, postingId } = IS_ASHBY ? getPostingRef() : { board: '', postingId: '' };
        const res = await chrome.runtime.sendMessage({
          type: 'ASHBY_FILL',
          question,
          board,
          postingId,
          pageText: getJobDescriptionText(),
          remark: remarkBox.value.trim(),
        });
        if (res && res.text) {
          setNativeValue(field, res.text);
          field.focus();
          setState('idle');
          remarkBox.style.display = 'none'; // close the context box once filled
          ctxBtn.textContent = '💬 Add context';
        } else {
          setState('error', (res && res.error) || 'No response from the extension.');
        }
      } catch (err) {
        setState('error', err.message || 'Unexpected error.');
      }
    }

    btn.addEventListener('click', runFill);

    ctxBtn.addEventListener('click', () => {
      const hidden = remarkBox.style.display === 'none';
      remarkBox.style.display = hidden ? 'block' : 'none';
      ctxBtn.textContent = hidden ? '💬 Hide context' : '💬 Add context';
      if (hidden) remarkBox.focus();
    });

    // Enter submits (Shift+Enter inserts a newline).
    remarkBox.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        runFill();
      }
    });

    const row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:6px;align-items:center;';
    row.appendChild(btn);
    row.appendChild(ctxBtn);

    const wrap = document.createElement('div');
    wrap.setAttribute('data-aifill-ctrl', '1'); // so scan() ignores our own textarea
    wrap.style.cssText = 'display:flex;flex-direction:column;align-items:flex-start;';
    wrap.appendChild(row);
    wrap.appendChild(remarkBox);

    field.insertAdjacentElement('afterend', wrap);
  }

  // --- Scan + observe ------------------------------------------------------

  function scan() {
    document.querySelectorAll('input, textarea').forEach((el) => {
      if (el.closest('[data-aifill-ctrl]')) return; // ignore our own remark boxes
      if (isFillable(el)) attachButton(el);
    });
  }

  scan();

  new MutationObserver(() => {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(scan, 500);
  }).observe(document.body, { childList: true, subtree: true });
})();
