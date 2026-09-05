// SPARK — exam.js
(function () {
  const SESSION_ID = parseInt(document.getElementById('exam-session-id')?.value);
  const EXAM_ID = parseInt(document.getElementById('exam-id')?.value);
  const CLASS_ID = parseInt(document.getElementById('exam-class-id')?.value);
  const DURATION_SECONDS = parseInt(document.getElementById('exam-duration')?.value) * 60;
  const TAB_LIMIT = parseInt(document.getElementById('tab-limit')?.value);
  const TAB_SWITCH_ENABLED = document.getElementById('tab-switch-enabled')?.value === '1';
  const FULLSCREEN_REQUIRED = document.getElementById('fullscreen-required')?.value === '1';
  let CONSENT_GIVEN = document.getElementById('consent-given')?.value === '1';

  let tabSwitchCount = 0;
  let terminated = false;
  let fullscreenIntentional = false; // true only right after we submit the exam ourselves
  let monitoringStarted = false;

  // Use server-calculated remaining time so timer persists across page reloads/re-entries
  const remainingEl = document.getElementById('time-remaining');
  let timerSeconds = remainingEl ? parseInt(remainingEl.value) : DURATION_SECONDS;
  let autoSubmitted = false;

  // ── EXAM TIMER — counts down from the teacher-set duration and auto-submits at 0 ──
  const timerDisplayEl = document.getElementById('exam-timer');
  function formatTime(totalSeconds) {
    const s = Math.max(0, totalSeconds);
    const mm = Math.floor(s / 60);
    const ss = s % 60;
    return String(mm).padStart(2, '0') + ':' + String(ss).padStart(2, '0');
  }
  function renderTimer() {
    if (!timerDisplayEl) return;
    timerDisplayEl.textContent = formatTime(timerSeconds);
    if (timerSeconds <= 60) timerDisplayEl.classList.add('timer-warning');
  }
  renderTimer(); // show the correct time immediately, don't wait for the first tick
  const timerIntervalId = setInterval(() => {
    if (terminated) { clearInterval(timerIntervalId); return; }
    if (timerSeconds <= 0) {
      renderTimer();
      clearInterval(timerIntervalId);
      if (!autoSubmitted) {
        autoSubmitted = true;
        submitExam();
      }
      return;
    }
    timerSeconds -= 1;
    renderTimer();
  }, 1000);

  // ── ANTI-COPY / ANTI-CHEAT (UI restriction only, no data logged — active immediately) ──
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

  // ── BLUR OVERLAY (shared utility — element/helpers only, no listeners attached yet) ──
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

  // ── CONSENT GATE ──
  // Nothing below this point (fullscreen enforcement, blur/tab-switch logging,
  // heartbeat, status polling) may run until the student has explicitly agreed
  // to the monitoring + data privacy statement.
  const consentModal = document.getElementById('consent-modal');
  const agreeBtn = document.getElementById('consent-agree-btn');
  const declineBtn = document.getElementById('consent-decline-btn');
  const consentErrorEl = document.getElementById('consent-error');

  function goToExamsList() {
    window.location.href = CLASS_ID ? `/student/class/${CLASS_ID}` : '/student';
  }

  if (!CONSENT_GIVEN) {
    if (consentModal) consentModal.style.display = 'flex';
  }

  agreeBtn?.addEventListener('click', async () => {
    agreeBtn.disabled = true;
    agreeBtn.textContent = 'Please wait…';
    try {
      const res = await fetch('/api/exam-consent', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: SESSION_ID })
      });
      const data = await res.json();
      if (res.ok && data.status === 'ok') {
        CONSENT_GIVEN = true;
        if (consentModal) consentModal.style.display = 'none';
        startMonitoring();
      } else {
        throw new Error(data.message || 'Could not record consent.');
      }
    } catch (e) {
      if (consentErrorEl) {
        consentErrorEl.textContent = 'Something went wrong recording your consent. Please try again.';
        consentErrorEl.style.display = 'block';
      }
      agreeBtn.disabled = false;
      agreeBtn.textContent = 'I Understand & Agree';
    }
  });

  declineBtn?.addEventListener('click', () => {
    goToExamsList();
  });

  // Entry point for everything monitoring-related — called immediately if consent
  // was already recorded (e.g. page refresh mid-exam), or after the Agree click.
  function startMonitoring() {
    if (monitoringStarted) return;
    monitoringStarted = true;
    initFullscreenEnforcement();
    initBlurAndTabTracking();
    initStatusPolling();
    initHeartbeat();
  }

  // ── FULLSCREEN MODE ENFORCEMENT ──
  function initFullscreenEnforcement() {
  const fsGate = document.getElementById('fullscreen-gate');
  const enterFsBtn = document.getElementById('enter-fullscreen-btn');

  function isFullscreen() {
    return !!(document.fullscreenElement || document.webkitFullscreenElement || document.msFullscreenElement);
  }

  function requestFullscreen() {
    const el = document.documentElement;
    const req = el.requestFullscreen || el.webkitRequestFullscreen || el.msRequestFullscreen;
    if (req) req.call(el).catch(() => {});
  }

  if (FULLSCREEN_REQUIRED) {
    if (!isFullscreen()) {
      // Block the exam behind a gate until the student clicks to enter fullscreen
      // (browsers require a user gesture to trigger the Fullscreen API).
      if (fsGate) fsGate.style.display = 'flex';
    }
    enterFsBtn?.addEventListener('click', () => {
      requestFullscreen();
      if (fsGate) fsGate.style.display = 'none';
    });

    document.addEventListener('fullscreenchange', handleFullscreenChange);
    document.addEventListener('webkitfullscreenchange', handleFullscreenChange);
    document.addEventListener('msfullscreenchange', handleFullscreenChange);
  }

  function handleFullscreenChange() {
    if (terminated || fullscreenIntentional) return;
    if (!isFullscreen()) {
      // Re-show the gate so the student must deliberately re-enter fullscreen
      if (fsGate) fsGate.style.display = 'flex';
      showBlur('You exited fullscreen mode. This has been logged as a violation.');
      logEvent('fullscreen_exit');
    } else {
      if (fsGate) fsGate.style.display = 'none';
      hideBlur();
    }
  }

  window._examIsFullscreen = isFullscreen;
  window._examRequestFullscreen = requestFullscreen;
  }

  // ── BLUR OVERLAY + visibility/focus tracking ──
  function initBlurAndTabTracking() {
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
  window.addEventListener('blur', () => {
    setTimeout(() => {
      if (!terminated) {
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
    if (FULLSCREEN_REQUIRED && !window._examIsFullscreen?.()) {
      window._examRequestFullscreen?.();
      return; // fullscreenchange handler will hide the blur once fullscreen is restored
    }
    if (document.hasFocus()) hideBlur();
  });
  }

  // ── STATUS CHECK every 5s — also syncs the shared exam timer ──
  function initStatusPolling() {
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
  }

  // ── HEARTBEAT every 4s — lets teacher see connected vs disconnected ──
  function initHeartbeat() {
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
  }

  if (CONSENT_GIVEN) startMonitoring();
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

  // ── DEBOUNCED SAVE FOR TEXT INPUT ──
  // Typing/deleting fast fires many 'input' events back-to-back. Each one used
  // to trigger its own independent network request, and those requests could
  // land at the server out of order (a slower request carrying OLD text could
  // finish after a newer request carrying the now-empty text), leaving stale
  // text saved even though the field looks empty. Debouncing collapses a burst
  // of keystrokes into a single save of the truly-final value, so there's only
  // ever one request in flight per question and no ordering race.
  const SAVE_DEBOUNCE_MS = 500;
  const _saveDebounceTimers = {};
  function debouncedSaveAnswer(questionId, answerText) {
    // Keep the local backup instant so a refresh never loses what was typed
    setLocalAnswer(questionId, answerText);
    clearTimeout(_saveDebounceTimers[questionId]);
    _saveDebounceTimers[questionId] = setTimeout(() => {
      delete _saveDebounceTimers[questionId];
      saveAnswer(questionId, answerText);
    }, SAVE_DEBOUNCE_MS);
  }
  // Immediately commit the current value and cancel any pending debounce for
  // this question — used at explicit checkpoints (switching sections) where
  // we want what's on screen saved right away rather than waiting out the
  // debounce window.
  function flushSaveAnswer(questionId, answerText) {
    clearTimeout(_saveDebounceTimers[questionId]);
    delete _saveDebounceTimers[questionId];
    saveAnswer(questionId, answerText);
  }
  window._examFlushSaveAnswer = flushSaveAnswer;

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

  // Wire up textareas: debounce saves so fast typing/deleting can't race itself
  document.querySelectorAll('textarea.auto-save').forEach(ta => {
    ta.addEventListener('input', () => debouncedSaveAnswer(ta.dataset.qid, ta.value));
  });

  // Wire up fill-in-the-blank inputs: on every keystroke, join all blanks for
  // that question (in order) with '|' into the hidden textarea, then save it
  // through the same pipeline as short-answer questions.
  document.querySelectorAll('.fib-blank-input').forEach(inp => {
    inp.addEventListener('input', () => {
      const qid = inp.dataset.qid;
      const blanks = Array.from(document.querySelectorAll(`.fib-blank-input[data-qid="${qid}"]`))
        .sort((a, b) => parseInt(a.dataset.blankIndex) - parseInt(b.dataset.blankIndex));
      // If every blank is empty, save an actual empty string — otherwise
      // join() still inserts '|' separators between the empty values
      // (e.g. "|" for 2 blanks), which is non-empty and wrongly counts
      // as "answered" even though nothing was typed.
      const allEmpty = blanks.every(b => b.value.trim() === '');
      const joined = allEmpty ? '' : blanks.map(b => b.value).join('|');
      const hidden = document.querySelector(`.fib-hidden[data-qid="${qid}"]`);
      if (hidden) hidden.value = joined;
      debouncedSaveAnswer(qid, joined);
    });
  });

  // Re-apply selected class to any radios already checked (from server or localStorage restore)
  document.querySelectorAll('.choice-item input[type=radio]:checked').forEach(radio => {
    applySelectedClass(radio);
  });

  // ── SUBMIT ──
  function submitExam() {
    if (terminated) return;
    fullscreenIntentional = true;
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
