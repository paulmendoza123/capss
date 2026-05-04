// SPARK — exam.js
(function () {
  const SESSION_ID = parseInt(document.getElementById('exam-session-id')?.value);
  const EXAM_ID = parseInt(document.getElementById('exam-id')?.value);
  const CLASS_ID = parseInt(document.getElementById('exam-class-id')?.value);
  const DURATION_SECONDS = parseInt(document.getElementById('exam-duration')?.value) * 60;
  const TAB_LIMIT = parseInt(document.getElementById('tab-limit')?.value);
  const TAB_SWITCH_ENABLED = document.getElementById('tab-switch-enabled')?.value === '1';

  let tabSwitchCount = 0;
  let terminated = false;

  // Use server-calculated remaining time so timer persists across page reloads/re-entries
  const remainingEl = document.getElementById('time-remaining');
  let timerSeconds = remainingEl ? parseInt(remainingEl.value) : DURATION_SECONDS;

  // ── ANTI-COPY / ANTI-CHEAT ──
  document.addEventListener('contextmenu', e => e.preventDefault());
  document.addEventListener('copy', e => e.preventDefault());
  document.addEventListener('cut', e => e.preventDefault());
  document.addEventListener('selectstart', e => {
    // Allow selection inside textareas (short answer) but block question text
    if (e.target.tagName === 'TEXTAREA' || e.target.tagName === 'INPUT') return;
    e.preventDefault();
  });
  document.addEventListener('keydown', e => {
    const blocked = (
      (e.ctrlKey || e.metaKey) && ['c','x','u','s','p','a'].includes(e.key.toLowerCase())
    ) || e.key === 'F12' || (e.ctrlKey && e.shiftKey && ['i','j','c'].includes(e.key.toLowerCase()));
    if (blocked) { e.preventDefault(); e.stopPropagation(); }
  });

  // ── SCREENSHOT BLOCKING & DETECTION ──
  // 1. CSS-level: make content invisible to screen capture APIs (where supported)
  (function applyScreenshotBlock() {
    const style = document.createElement('style');
    style.id = 'spark-screenshot-block';
    style.textContent = `
      /* Chrome/Edge: hide from getDisplayMedia captures */
      body { -webkit-user-select: none !important; }
      #exam-content, .exam-question, .question-card, .choice-item {
        -webkit-user-select: none !important;
        user-select: none !important;
      }
      /* Overlay shown during detected capture */
      #screenshot-overlay {
        display: none;
        position: fixed; inset: 0; z-index: 99999;
        background: #111;
        align-items: center; justify-content: center;
        flex-direction: column; gap: 1rem;
        color: #fff; font-size: 1.2rem; font-weight: 600;
        text-align: center; padding: 2rem;
      }
      #screenshot-overlay.active { display: flex; }
    `;
    document.head.appendChild(style);

    // Inject dark overlay element
    const overlay = document.createElement('div');
    overlay.id = 'screenshot-overlay';
    overlay.innerHTML = `
      <div style="font-size:2.5rem">🚫</div>
      <div>Screenshots are not allowed during this exam.</div>
    `;
    document.body.appendChild(overlay);
  })();

  // 2. Detect PrintScreen / screenshot keyboard shortcuts
  let _screenshotFlashTimer = null;
  function onScreenshotDetected(method) {
    if (terminated) return;
    // Briefly show blackout overlay
    const overlay = document.getElementById('screenshot-overlay');
    if (overlay) {
      overlay.classList.add('active');
      clearTimeout(_screenshotFlashTimer);
      _screenshotFlashTimer = setTimeout(() => overlay.classList.remove('active'), 2500);
    }
    // Log to server
    logEvent('screenshot');
  }

  document.addEventListener('keyup', e => {
    // PrintScreen (all platforms)
    if (e.key === 'PrintScreen' || e.code === 'PrintScreen') {
      // Clear clipboard to prevent pasting the captured image
      try {
        navigator.clipboard.writeText('').catch(() => {});
      } catch(ex) {}
      onScreenshotDetected('printscreen');
    }
    // macOS: Cmd+Shift+3 (full screen), Cmd+Shift+4 (selection), Cmd+Shift+5 (toolbar)
    if (e.metaKey && e.shiftKey && ['3','4','5'].includes(e.key)) {
      onScreenshotDetected('mac_screenshot');
    }
    // Windows Snipping Tool: Win+Shift+S — detected via key combo approximation
    if (e.key === 'S' && e.shiftKey && (e.metaKey || e.getModifierState?.('OS'))) {
      onScreenshotDetected('snipping_tool');
    }
  });

  // 3. Detect screen capture via Page Visibility + visibilityState changes
  //    macOS screenshot briefly steals focus; catch it via blur without visibilitychange
  let _lastBlurTime = 0;
  let _screenshotBlurSuppressed = false;
  window.addEventListener('blur', () => { _lastBlurTime = Date.now(); _screenshotBlurSuppressed = false; });
  window.addEventListener('focus', () => {
    // Very short blur (<800ms) that isn't a full tab switch = likely screenshot tool
    const delta = Date.now() - _lastBlurTime;
    if (delta > 0 && delta < 800 && !document.hidden) {
      _screenshotBlurSuppressed = true;
      onScreenshotDetected('focus_blur_rapid');
    }
  });

  // ── TIMER ──
  const timerEl = document.getElementById('exam-timer');
  function updateTimer() {
    if (terminated) return;
    const m = Math.floor(timerSeconds / 60);
    const s = timerSeconds % 60;
    if (timerEl) {
      timerEl.textContent = `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
      timerEl.className = 'exam-timer';
      if (timerSeconds <= 60) timerEl.classList.add('danger');
      else if (timerSeconds <= 300) timerEl.classList.add('warning');
    }
    if (timerSeconds <= 0) { submitExam(); return; }
    timerSeconds--;
    setTimeout(updateTimer, 1000);
  }
  updateTimer();

  // ── BLUR OVERLAY ──
  const blurOverlay = document.getElementById('blur-overlay');
  let blurShown = false;

  function showBlur(reason) {
    if (terminated || !blurOverlay || blurShown) return;
    blurShown = true;
    blurOverlay.style.display = 'flex';
    const reasonEl = document.getElementById('blur-reason');
    if (reasonEl && reason) reasonEl.textContent = reason;
  }
  function hideBlur() {
    if (!blurOverlay) return;
    blurShown = false;
    blurOverlay.style.display = 'none';
  }

  // visibilitychange — tab switch, new tab, navigate away
  document.addEventListener('visibilitychange', () => {
    if (document.hidden) {
      showBlur('Tab switch or window hidden detected.');
      logEvent('tab_switch');
    } else {
      hideBlur();
    }
  });

  // window blur — app switch, minimize
  // Delayed 800ms so screenshot blurs (which return focus quickly) can be suppressed
  window.addEventListener('blur', () => {
    setTimeout(() => {
      if (!_screenshotBlurSuppressed && !terminated) {
        showBlur('Window minimized or focus lost.');
        logEvent('window_minimize');
      }
    }, 800);
  });
  window.addEventListener('focus', () => hideBlur());

  // Polling fallback every 500ms
  setInterval(() => {
    if (terminated) return;
    const inactive = document.hidden || !document.hasFocus();
    if (inactive) showBlur();
    else hideBlur();
  }, 500);

  blurOverlay?.addEventListener('click', () => {
    if (document.hasFocus()) hideBlur();
  });

  // ── EVENT LOGGING (tab_switch + window_minimize) ──
  let eventDebounce = {};

  function logEvent(eventType) {
    if (terminated) return;
    clearTimeout(eventDebounce[eventType]);
    eventDebounce[eventType] = setTimeout(async () => {
      try {
        const res = await fetch('/api/log-suspicious', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ session_id: SESSION_ID, event_type: eventType })
        });
        const data = await res.json();
        tabSwitchCount = data.count || tabSwitchCount + 1;

        const counter = document.getElementById('tab-switch-counter');
        if (counter) counter.textContent = tabSwitchCount;

        if (data.terminated) handleTerminated();
      } catch (e) {
        console.error('Log error:', e);
      }
    }, 300);
  }

  function handleTerminated() {
    terminated = true;
    const banner = document.getElementById('tab-warning-banner');
    if (banner) {
      banner.innerHTML = `🚫 Your exam has been terminated. Redirecting...`;
      banner.className = 'tab-warning danger';
      banner.style.display = 'block';
    }
    setTimeout(() => {
      window.location.href = CLASS_ID ? `/student/class/${CLASS_ID}` : '/student';
    }, 2500);
  }

  // ── STATUS CHECK every 5s — also syncs the shared exam timer ──
  setInterval(async () => {
    if (terminated) return;
    try {
      const res = await fetch(`/api/exam-status/${EXAM_ID}`);
      const data = await res.json();
      if (data.session_status === 'terminated') handleTerminated();
      // Sync timer with server's authoritative remaining time (shared exam clock)
      if (typeof data.time_remaining_seconds === 'number' && data.time_remaining_seconds >= 0) {
        // Only sync if drift is more than 3 seconds to avoid jitter
        if (Math.abs(timerSeconds - data.time_remaining_seconds) > 3) {
          timerSeconds = data.time_remaining_seconds;
        }
      }
    } catch (e) {}
  }, 5000);

  // ── HEARTBEAT every 4s — lets teacher see connected vs disconnected ──
  async function sendHeartbeat() {
    if (terminated) return;
    try {
      await fetch('/api/heartbeat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: SESSION_ID })
      });
    } catch (e) {}
  }
  sendHeartbeat(); // send immediately on load (logs "connected" on reconnect)
  setInterval(sendHeartbeat, 4000);

  // ── AUTO SAVE ──
  // localStorage key for this session's answers (backup against fast refresh)
  const ANSWERS_KEY = `exam_answers_${SESSION_ID}`;

  // Load any locally cached answers (survives refresh before server fetch completes)
  function getLocalAnswers() {
    try { return JSON.parse(localStorage.getItem(ANSWERS_KEY) || '{}'); } catch(e) { return {}; }
  }
  function setLocalAnswer(questionId, answerText) {
    try {
      const all = getLocalAnswers();
      all[questionId] = answerText;
      localStorage.setItem(ANSWERS_KEY, JSON.stringify(all));
    } catch(e) {}
  }

  // ── SAVE STATUS INDICATOR ──
  // Shows a brief "Saved ✓" badge near the top of the exam
  let _saveStatusTimer = null;
  function showSaveStatus(ok) {
    let el = document.getElementById('save-status-badge');
    if (!el) {
      el = document.createElement('div');
      el.id = 'save-status-badge';
      el.style.cssText = [
        'position:fixed', 'bottom:1.25rem', 'right:1.25rem', 'z-index:9998',
        'padding:0.4rem 0.85rem', 'border-radius:6px', 'font-size:0.8rem',
        'font-weight:600', 'pointer-events:none',
        'transition:opacity 0.3s ease'
      ].join(';');
      document.body.appendChild(el);
    }
    clearTimeout(_saveStatusTimer);
    if (ok) {
      el.textContent = '✓ Answer saved';
      el.style.background = 'rgba(52,199,89,0.18)';
      el.style.border = '1px solid rgba(52,199,89,0.45)';
      el.style.color = '#34c759';
    } else {
      el.textContent = '⚠ Save failed – retrying…';
      el.style.background = 'rgba(247,95,95,0.15)';
      el.style.border = '1px solid rgba(247,95,95,0.35)';
      el.style.color = '#f75f5f';
    }
    el.style.opacity = '1';
    _saveStatusTimer = setTimeout(() => { el.style.opacity = '0'; }, 2000);
  }

  // Save to server — primary: fetch with keepalive (sends session cookie reliably)
  // Fallback: sendBeacon (for page-unload survival)
  function saveAnswer(questionId, answerText) {
    if (terminated) return;
    // 1. Save locally immediately (instant, synchronous)
    setLocalAnswer(questionId, answerText);

    // 2. Primary: fetch with keepalive — sends cookies properly, works across browsers
    const payload = JSON.stringify({
      session_id: SESSION_ID,
      question_id: parseInt(questionId),
      answer_text: answerText
    });

    fetch('/api/save-answer', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      credentials: 'same-origin',   // always send session cookie
      keepalive: true,               // survives page unload (modern browsers)
      body: payload
    })
    .then(r => r.json())
    .then(data => {
      showSaveStatus(data.status === 'saved');
    })
    .catch(() => {
      // 3. Last-resort fallback: sendBeacon (no response, fire-and-forget)
      try {
        const blob = new Blob([payload], { type: 'application/json' });
        navigator.sendBeacon('/api/save-answer', blob);
      } catch(e) {}
      showSaveStatus(false);
    });
  }
  window._examSaveAnswer = saveAnswer;

  // Fix selected class on label when a choice is clicked
  function applySelectedClass(radio) {
    const name = radio.name;
    document.querySelectorAll(`input[type=radio][name="${name}"]`).forEach(r => {
      r.closest('.choice-item')?.classList.remove('selected');
    });
    radio.closest('.choice-item')?.classList.add('selected');
  }

  // Wire up radio buttons: save instantly on click (use 'click' not 'change'
  // so re-clicking the same option also triggers a save/visual feedback)
  document.querySelectorAll('.choice-item input[type=radio]').forEach(radio => {
    radio.addEventListener('click', () => {
      applySelectedClass(radio);
      saveAnswer(radio.dataset.qid, radio.value);
    });
  });

  // Wire up textareas: save instantly on every keystroke
  document.querySelectorAll('textarea.auto-save').forEach(ta => {
    ta.addEventListener('input', () => saveAnswer(ta.dataset.qid, ta.value));
  });

  // Re-apply selected class to any radios already checked (from server or localStorage restore)
  document.querySelectorAll('.choice-item input[type=radio]:checked').forEach(radio => {
    applySelectedClass(radio);
  });

  // ── SUBMIT ──
  function submitExam() {
    if (terminated) return;
    // Clear local backups on submit
    try {
      localStorage.removeItem('exam_section_' + SESSION_ID);
      localStorage.removeItem(ANSWERS_KEY);
    } catch(e) {}
    document.getElementById('exam-form')?.submit();
  }

  function confirmAndSubmit() {
    if (terminated) return;
    if (confirm('Are you sure you want to submit your exam? This cannot be undone.')) submitExam();
  }

  document.getElementById('submit-exam-btn-bottom')?.addEventListener('click', (e) => {
    e.preventDefault();
    confirmAndSubmit();
  });
})();
