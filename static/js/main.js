// SPARK — main.js

// ── CSRF: automatically attach the token to every same-origin fetch() that
// changes data (POST/PUT/PATCH/DELETE), so none of the app's existing
// fetch() calls (save-answer, heartbeat, reorder, log-suspicious, etc.) had
// to be edited individually. GET requests and requests to other origins are
// left untouched. Runs immediately (not on DOMContentLoaded) so it's in
// place before any other script on the page makes its first fetch() call.
(function () {
  var token = document.querySelector('meta[name="csrf-token"]');
  if (!token) return;
  token = token.content;
  var origFetch = window.fetch;
  window.fetch = function (input, init) {
    init = init || {};
    var method = (init.method || (input && input.method) || 'GET').toUpperCase();
    var url = typeof input === 'string' ? input : (input && input.url) || '';
    var sameOrigin = !/^([a-z]+:)?\/\//i.test(url) || url.indexOf(window.location.origin) === 0;
    if (sameOrigin && method !== 'GET' && method !== 'HEAD') {
      init = Object.assign({}, init);
      init.headers = new Headers(init.headers || (input && input.headers) || {});
      if (!init.headers.has('X-CSRFToken')) init.headers.set('X-CSRFToken', token);
    }
    return origFetch.call(this, input, init);
  };
})();

document.addEventListener('DOMContentLoaded', () => {
  // Card kebab menus (class cards, etc.)
  document.querySelectorAll('.card-menu-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      const dropdown = btn.nextElementSibling;
      const isOpen = dropdown.classList.contains('open');
      document.querySelectorAll('.card-menu-dropdown.open').forEach(d => d.classList.remove('open'));
      if (!isOpen) dropdown.classList.add('open');
    });
  });
  document.querySelectorAll('.card-menu-dropdown').forEach(dropdown => {
    dropdown.addEventListener('click', (e) => e.stopPropagation());
  });
  document.addEventListener('click', () => {
    document.querySelectorAll('.card-menu-dropdown.open').forEach(d => d.classList.remove('open'));
  });

  // Auto-dismiss alerts after 4s
  document.querySelectorAll('.alert').forEach(alert => {
    setTimeout(() => {
      alert.style.transition = 'opacity 0.4s';
      alert.style.opacity = '0';
      setTimeout(() => alert.remove(), 400);
    }, 4000);
  });

  // Tab switching — scoped to sibling tab-contents only
  document.querySelectorAll('.tabs').forEach(tabGroup => {
    const tabBtns = tabGroup.querySelectorAll('.tab-btn');

    // Collect the matching tab-content siblings that follow this .tabs element
    // They live as direct siblings in the same parent
    const parent = tabGroup.parentElement;
    const tabContents = parent ? Array.from(parent.querySelectorAll(':scope > .tab-content')) : [];

    tabBtns.forEach(btn => {
      btn.addEventListener('click', () => {
        const target = btn.dataset.tab;

        // Update active button within this tab group
        tabBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');

        // Show/hide only the sibling tab-contents
        tabContents.forEach(tc => {
          if (tc.dataset.tab === target) {
            tc.classList.add('active');
          } else {
            tc.classList.remove('active');
          }
        });

        // Update URL hash without scrolling
        history.replaceState(null, '', '#' + target);
      });
    });

    // Restore active tab from URL hash on page load
    const hash = window.location.hash.replace('#', '');
    if (hash) {
      const matchBtn = Array.from(tabBtns).find(b => b.dataset.tab === hash);
      if (matchBtn) matchBtn.click();
    }
  });

  // Choice selection highlight
  document.querySelectorAll('.choice-item').forEach(item => {
    item.addEventListener('click', () => {
      const radio = item.querySelector('input[type=radio]');
      if (radio) {
        const name = radio.name;
        document.querySelectorAll(`input[name="${name}"]`).forEach(r => {
          r.closest('.choice-item')?.classList.remove('selected');
        });
        radio.checked = true;
        item.classList.add('selected');
      }
    });
  });

  // Modal helpers
  window.openModal = id => document.getElementById(id)?.classList.add('open');
  window.closeModal = id => {
    const modal = document.getElementById(id);
    if (!modal) return;
    // Closing a modal (via ✕, Cancel, or clicking outside) should discard any
    // unsaved edits rather than leave them sitting in the form for next time.
    modal.querySelectorAll('form').forEach(form => {
      form.reset();
      form.querySelectorAll('.mc-choice-grid').forEach(grid => {
        // form.reset() restores each choice's text, but not the row's
        // show/hide state (that's a manual style, not a form value) —
        // clear it so mcInitGrid can recompute it from the restored text.
        grid.querySelectorAll('.mc-choice-row').forEach(row => row.style.removeProperty('display'));
        if (typeof window.mcInitGrid === 'function') window.mcInitGrid(grid);
      });
    });
    modal.classList.remove('open');
  };
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', e => {
      if (e.target === overlay) closeModal(overlay.id);
    });
  });

  // Show/hide password toggle buttons
  document.querySelectorAll('.password-toggle-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const input = document.getElementById(btn.dataset.target);
      if (!input) return;
      const isVisible = input.type === 'text';
      input.type = isVisible ? 'password' : 'text';
      btn.classList.toggle('is-visible', !isVisible);
      btn.setAttribute('aria-label', isVisible ? 'Show password' : 'Hide password');
    });
  });

  // Live password-requirements checklist (length + character variety)
  document.querySelectorAll('.pw-requirements').forEach(box => {
    const input = document.getElementById(box.dataset.target);
    if (!input) return;
    const lengthItem = box.querySelector('[data-rule="length"]');
    const varietyItem = box.querySelector('[data-rule="variety"]');
    const subLower = box.querySelector('[data-rule="lower"]');
    const subUpper = box.querySelector('[data-rule="upper"]');
    const subNumber = box.querySelector('[data-rule="number"]');
    const specialItem = box.querySelector('[data-rule="special"]');

    // On the signup pages, once the person has tried to submit (or the server
    // rejected the password), any requirement that is NOT met turns red.
    let showFailed = box.dataset.showFailed === '1';

    const check = () => {
      const val = input.value;
      const hasLower = /[a-z]/.test(val);
      const hasUpper = /[A-Z]/.test(val);
      const hasNumber = /[0-9]/.test(val);
      const hasSpecial = /[^A-Za-z0-9]/.test(val);
      // Special characters are their own requirement; the "at least 3 of the following"
      // group only covers lower case, upper case and numbers.
      const classesMet = [hasLower, hasUpper, hasNumber].filter(Boolean).length;
      const lengthOk = val.length >= 8;
      const specialOk = hasSpecial;
      const varietyOk = classesMet >= 3;

      lengthItem?.classList.toggle('met', lengthOk);
      subLower?.classList.toggle('met', hasLower);
      subUpper?.classList.toggle('met', hasUpper);
      subNumber?.classList.toggle('met', hasNumber);
      specialItem?.classList.toggle('met', specialOk);
      varietyItem?.classList.toggle('met', varietyOk);

      lengthItem?.classList.toggle('failed', showFailed && !lengthOk);
      specialItem?.classList.toggle('failed', showFailed && !specialOk);
      varietyItem?.classList.toggle('failed', showFailed && !varietyOk);
      // Only highlight the individual character types when the "3 of 4" rule itself failed
      const flagSub = (el, ok) => el?.classList.toggle('failed', showFailed && !varietyOk && !ok);
      flagSub(subLower, hasLower);
      flagSub(subUpper, hasUpper);
      flagSub(subNumber, hasNumber);

      return lengthOk && specialOk && varietyOk;
    };

    // Called by the signup form on submit: start showing red, return whether all rules pass
    box.enableFailed = () => { showFailed = true; return check(); };

    input.addEventListener('input', check);
    check(); // initial state, e.g. browser autofill
  });

  // Live password confirmation match check
  document.querySelectorAll('[data-confirm-target]').forEach(confirmInput => {
    const original = document.getElementById(confirmInput.dataset.confirmTarget);
    const msg = confirmInput.closest('.form-group')?.querySelector('.password-mismatch-msg');
    const form = confirmInput.closest('form');
    if (!original) return;

    const check = () => {
      const mismatch = confirmInput.value.length > 0 && confirmInput.value !== original.value;
      confirmInput.setCustomValidity(mismatch ? 'Passwords do not match' : '');
      if (msg) msg.classList.toggle('is-visible', mismatch);
      return !mismatch;
    };

    confirmInput.addEventListener('input', check);
    original.addEventListener('input', check);
    if (form && !form.hasAttribute('data-auth-form')) {
      form.addEventListener('submit', e => {
        if (!check()) {
          e.preventDefault();
          confirmInput.reportValidity();
        }
      });
    }
  });

  // Signup forms: validate on submit and turn ONLY the failing fields /
  // requirements red (no message at the top of the page).
  document.querySelectorAll('form[data-auth-form]').forEach(form => {
    const emailPattern = /^[^@\s]+@psu\.palawan\.edu\.ph$/i;
    const emailInput = form.querySelector('input[type="email"]');
    const emailHint = form.querySelector('[data-email-hint]');
    const pwInput = form.querySelector('.pw-requirements') &&
      document.getElementById(form.querySelector('.pw-requirements').dataset.target);
    const pwBox = form.querySelector('.pw-requirements');
    const confirmInput = form.querySelector('[data-confirm-target]');
    const mismatchMsg = confirmInput?.closest('.form-group')?.querySelector('.password-mismatch-msg');

    const resetEmailHint = () => {
      if (!emailHint) return;
      emailHint.classList.remove('is-error');
      emailHint.textContent = emailHint.dataset.default;
    };

    // Clear a field's red state as soon as the person edits it
    form.querySelectorAll('input, select').forEach(el => {
      const clear = () => {
        if (el === emailInput) resetEmailHint();
        if (el === pwInput || el === confirmInput) {
          if (confirmInput && confirmInput.value === pwInput.value) {
            confirmInput.classList.remove('is-invalid');
            mismatchMsg?.classList.remove('is-visible');
          }
        }
        if (el !== pwInput) el.classList.remove('is-invalid');
        else if (pwBox && pwBox.enableFailed && el.classList.contains('is-invalid')) {
          // keep the checklist live, clear the red border once everything passes
          const ok = pwBox.enableFailed();
          if (ok) el.classList.remove('is-invalid');
        }
      };
      el.addEventListener('input', clear);
      el.addEventListener('change', clear);
    });

    form.addEventListener('submit', e => {
      let firstBad = null;
      const flag = el => { el.classList.add('is-invalid'); firstBad = firstBad || el; };

      // Empty required fields (text, selects)
      form.querySelectorAll('input[required], select[required]').forEach(el => {
        const empty = el.type === 'password' ? !el.value : !el.value.trim();
        if (empty) flag(el);
      });

      // Email must be a @psu.palawan.edu.ph address
      if (emailInput && emailInput.value.trim() && !emailPattern.test(emailInput.value.trim())) {
        flag(emailInput);
        emailHint?.classList.add('is-error');
      }

      // Password requirements: unmet ones turn red
      if (pwBox && pwBox.enableFailed && pwInput) {
        if (!pwBox.enableFailed()) flag(pwInput);
      }

      // Confirm password must match
      if (confirmInput && pwInput && confirmInput.value && confirmInput.value !== pwInput.value) {
        flag(confirmInput);
        mismatchMsg?.classList.add('is-visible');
      }

      if (firstBad) {
        e.preventDefault();
        firstBad.focus();
      }
    });

    // If the server sent the page back with errors, jump to the first one
    const serverBad = form.querySelector('.is-invalid');
    if (serverBad) serverBad.focus();
  });

  // Forms outside the sign-up pages (admin: My Profile, Create User, Edit User,
  // Reset Password): block the submit and turn whatever failed red — the email
  // field and/or the unmet password requirements.
  document.querySelectorAll('form').forEach(form => {
    if (form.hasAttribute('data-auth-form')) return; // sign-up pages handle this themselves
    const pwBox = form.querySelector('.pw-requirements');
    const pwInput = pwBox && pwBox.enableFailed ? document.getElementById(pwBox.dataset.target) : null;
    const emailInput = form.querySelector('input[data-psu-email]');
    // Student accounts must have a program and a year level
    const roleSelect = form.querySelector('select[name="role"]');
    const programSel = form.querySelector('select[name="program"]');
    const yearSel = form.querySelector('select[name="year_level"]');
    const hasStudentFields = !!(roleSelect && programSel && yearSel);
    if (!pwInput && !emailInput && !hasStudentFields) return;

    const confirmInput = form.querySelector('[data-confirm-target]');
    const mismatchMsg = confirmInput?.closest('.form-group')?.querySelector('.password-mismatch-msg');
    const emailPattern = /^[^@\s]+@psu\.palawan\.edu\.ph$/i;
    const emailHint = form.querySelector('[data-email-hint]');

    // Email: must be @psu.palawan.edu.ph — unless it is an existing address
    // that is being left unchanged (data-original), e.g. the default admin.
    const emailIsBad = () => {
      const v = emailInput.value.trim();
      if (!v) return false; // empty is caught by "required"
      const original = (emailInput.dataset.original || '').toLowerCase();
      if (original && v.toLowerCase() === original) return false;
      return !emailPattern.test(v);
    };
    const clearEmailRed = () => {
      emailInput.classList.remove('is-invalid');
      if (emailHint) { emailHint.classList.remove('is-error'); emailHint.textContent = emailHint.dataset.default; }
    };

    // Capture phase + stopImmediatePropagation so inline/other submit prompts
    // ("Reset password for…?", "Grant full admin access…?") don't pop up for
    // a form that is going to be rejected anyway.
    form.addEventListener('submit', e => {
      let bad = null;
      if (emailInput && emailIsBad()) {
        emailInput.classList.add('is-invalid');
        emailHint?.classList.add('is-error');
        bad = emailInput;
      }
      if (pwInput && !pwBox.enableFailed()) { pwInput.classList.add('is-invalid'); bad = bad || pwInput; }
      if (hasStudentFields && roleSelect.value === 'student') {
        [programSel, yearSel].forEach(sel => {
          if (!sel.value) { sel.classList.add('is-invalid'); bad = bad || sel; }
        });
      }
      if (pwInput && confirmInput && confirmInput.value && confirmInput.value !== pwInput.value) {
        confirmInput.classList.add('is-invalid');
        mismatchMsg?.classList.add('is-visible');
        bad = bad || confirmInput;
      }
      if (bad) {
        e.preventDefault();
        e.stopImmediatePropagation();
        bad.focus();
      }
    }, true);

    // Any other field flagged red (e.g. "Current password is incorrect") clears once edited
    form.querySelectorAll('input.is-invalid').forEach(el => {
      if (el === pwInput || el === confirmInput || el === emailInput) return;
      el.addEventListener('input', () => el.classList.remove('is-invalid'));
    });

    if (hasStudentFields) {
      [programSel, yearSel].forEach(sel => sel.addEventListener('change', () => sel.classList.remove('is-invalid')));
      roleSelect.addEventListener('change', () => {
        if (roleSelect.value !== 'student') [programSel, yearSel].forEach(s => s.classList.remove('is-invalid'));
      });
    }

    if (emailInput) {
      emailInput.addEventListener('input', clearEmailRed);
      // Empty / not-an-email is stopped by the browser first — still show it red
      emailInput.addEventListener('invalid', () => {
        emailInput.classList.add('is-invalid');
        if (emailInput.value.trim()) emailHint?.classList.add('is-error');
      });
    }

    if (pwInput) {
      // An empty password is stopped by the browser's "required" check before
      // submit — still show the checklist in red in that case.
      pwInput.addEventListener('invalid', () => {
        pwBox.enableFailed();
        pwInput.classList.add('is-invalid');
      });
      // Clear the red border as soon as the password passes everything
      pwInput.addEventListener('input', () => {
        if (pwInput.classList.contains('is-invalid') && pwBox.enableFailed()) {
          pwInput.classList.remove('is-invalid');
        }
      });
      if (confirmInput) {
        const syncConfirm = () => {
          if (confirmInput.value === pwInput.value) {
            confirmInput.classList.remove('is-invalid');
            mismatchMsg?.classList.remove('is-visible');
          }
        };
        confirmInput.addEventListener('input', syncConfirm);
        pwInput.addEventListener('input', syncConfirm);
      }
    }
  });

  // If the server sent an admin page back with a failed email, jump to it
  const serverBad = document.querySelector('input.is-invalid, select.is-invalid');
  if (serverBad) serverBad.focus();
});


// ── Max upload size (10 MB) for every <input type="file"> ──
(function () {
  var MAX_MB = 10;
  var MAX_BYTES = MAX_MB * 1024 * 1024;
  document.addEventListener('change', function (e) {
    var input = e.target;
    if (!input || input.tagName !== 'INPUT' || input.type !== 'file') return;
    var tooBig = Array.prototype.filter.call(input.files || [], function (f) { return f.size > MAX_BYTES; });
    if (tooBig.length) {
      alert((tooBig.length === 1 ? '"' + tooBig[0].name + '" is' : 'Some files are') +
            ' too large. The maximum allowed size is ' + MAX_MB + ' MB.');
      input.value = '';
    }
  });
  // Safety net: block form submit if a file somehow slipped through.
  document.addEventListener('submit', function (e) {
    var inputs = e.target.querySelectorAll('input[type="file"]');
    for (var i = 0; i < inputs.length; i++) {
      var files = inputs[i].files || [];
      for (var j = 0; j < files.length; j++) {
        if (files[j].size > MAX_BYTES) {
          e.preventDefault();
          alert('"' + files[j].name + '" is too large. The maximum allowed size is ' + MAX_MB + ' MB.');
          return;
        }
      }
    }
  }, true);
})();


// ── Short answer: alternative answers ──
function addAltAnswer(btn) {
  var wrap = btn.closest('.alt-answers');
  if (!wrap) return;
  var first = wrap.querySelector('input[type="text"]');
  var row = document.createElement('div');
  row.className = 'alt-row';
  var input = document.createElement('input');
  input.type = 'text';
  input.name = first.name;
  input.className = first.className;
  input.removeAttribute('id');
  input.placeholder = 'Alternative answer';
  var remove = document.createElement('button');
  remove.type = 'button';
  remove.className = 'btn btn-ghost btn-sm alt-remove';
  remove.title = 'Remove';
  remove.setAttribute('aria-label', 'Remove alternative answer');
  remove.textContent = '\u2715';
  remove.onclick = function () { row.remove(); };
  row.appendChild(input);
  row.appendChild(remove);
  wrap.insertBefore(row, btn);
  input.focus();
}
// A form reset puts the box back to a single answer.
document.addEventListener('reset', function (e) {
  var extra = e.target.querySelectorAll ? e.target.querySelectorAll('.alt-answers .alt-row + .alt-row') : [];
  Array.prototype.forEach.call(extra, function (r) { r.remove(); });
});

// Profile cards: fields are read-only until the pen icon is clicked.
document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('form.profile-form').forEach(function (form) {
    var btn = document.querySelector('[data-profile-edit="' + form.id + '"]');
    var cancel = form.querySelector('[data-profile-cancel]');
    var fields = form.querySelectorAll('input[data-editable], select[data-editable]');
    var original = {};
    fields.forEach(function (f) { original[f.name] = f.value; });
    function setEditing(on) {
      form.classList.toggle('is-editing', on);
      if (btn) { btn.classList.toggle('is-active', on); btn.setAttribute('aria-pressed', on ? 'true' : 'false'); }
      fields.forEach(function (f) { f.disabled = !on; });
      if (on && fields[0]) { fields[0].focus(); if (fields[0].select) fields[0].select(); }
    }
    if (btn) btn.addEventListener('click', function () { setEditing(!form.classList.contains('is-editing')); if (!form.classList.contains('is-editing')) fields.forEach(function (f) { f.value = original[f.name]; }); });
    if (cancel) cancel.addEventListener('click', function () {
      fields.forEach(function (f) { f.value = original[f.name]; f.classList.remove('is-invalid'); });
      setEditing(false);
    });
    setEditing(form.classList.contains('is-editing'));
    if (!form.classList.contains('is-editing')) fields.forEach(function (f) { f.disabled = true; });
  });
});

// ── Separate date + time fields for exam scheduling ──────────────────────────
// Pages mark a group with data-dt-split: a date box, a time box and a hidden
// "scheduled_at" field. The two visible boxes use the browser's own date and
// time pickers (the same time picker the admin schedule editor uses); the hidden
// field is kept in sync as "YYYY-MM-DDTHH:MM", exactly what the server already expects.
(function () {
  function init() {
    document.querySelectorAll('[data-dt-split]').forEach(function (box) {
      var d = box.querySelector('[data-dt="date"]');
      var t = box.querySelector('[data-dt="time"]');
      var h = box.querySelector('input[type="hidden"]');
      if (!d || !t || !h) return;
      var row = box.closest('.form-row');
      if (row) row.classList.add('dt-wide');
      function sync() {
        h.value = (d.value && t.value) ? d.value + 'T' + t.value : '';
        d.required = !!t.value;      // if one is filled, the other must be too
        t.required = !!d.value;
      }
      var m = (h.value || '').match(/^(\d{4}-\d{2}-\d{2})[T ](\d{2}:\d{2})/);
      if (m) { d.value = m[1]; t.value = m[2]; }
      ['input', 'change'].forEach(function (ev) {
        d.addEventListener(ev, sync);
        t.addEventListener(ev, sync);
      });
      sync();
    });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();

// ── Exam timing fields (Create Exam + exam Settings) ─────────────────────
// Two independent timers that can be combined: the overall exam duration and
// the time per question. Either can be off, but not both unless the whole
// exam timer is switched off with the master "Exam Timer" toggle (no time limit).
(function () {
  function init() {
    const dToggle = document.getElementById('durationToggle');
    const qToggle = document.getElementById('questionTimeToggle');
    if (!dToggle || !qToggle) return;
    const dGroup = document.getElementById('durationGroup');
    const dInput = document.getElementById('durationInput');
    const dOff = document.getElementById('durationOffHint');
    const qGroup = document.getElementById('questionTimeGroup');
    const qInput = document.getElementById('questionTimeInput');
    const qUnit = document.getElementById('questionTimeUnit');
    const qOff = document.getElementById('questionTimeOffHint');
    const warnGroup = document.getElementById('timeWarningGroup');
    const errEl = document.getElementById('timingError');
    const master = document.getElementById('timerMasterToggle');
    const options = document.getElementById('timerOptions');
    let errTimer = null;

    function showError(msg) {
      if (!errEl) return;
      errEl.textContent = msg;
      errEl.style.display = msg ? '' : 'none';
      clearTimeout(errTimer);
      if (msg) errTimer = setTimeout(() => showError(''), 4000);
    }

    function sync() {
      // Master switch off = no time limit: hide the timers and don't submit/validate them.
      const timerOn = !master || master.checked;
      if (options) options.style.display = timerOn ? '' : 'none';
      const dOn = timerOn && dToggle.checked, qOn = timerOn && qToggle.checked;
      dGroup.style.display = dOn ? '' : 'none';
      if (dOff) dOff.style.display = dOn ? 'none' : '';
      // Disabled inputs are not submitted, and are skipped by required/min checks.
      dInput.disabled = !dOn;
      dInput.required = dOn;
      warnGroup.style.display = dOn ? '' : 'none';
      qGroup.style.display = qOn ? '' : 'none';
      if (qOff) qOff.style.display = qOn ? 'none' : '';
      qInput.disabled = !qOn;
      qUnit.disabled = !qOn;
      qInput.required = qOn;
    }

    if (master) master.addEventListener('change', () => {
      showError('');
      // Turning the timer back on with both timers off would be invalid: restore the duration.
      if (master.checked && !dToggle.checked && !qToggle.checked) dToggle.checked = true;
      sync();
    });

    function guard(toggle, other) {
      toggle.addEventListener('change', () => {
        if (!toggle.checked && !other.checked) {
          toggle.checked = true; // the exam must keep at least one time limit
          showError('At least one timer must stay on: the exam duration, the time per question, or both.');
        } else {
          showError('');
        }
        sync();
      });
    }
    guard(dToggle, qToggle);
    guard(qToggle, dToggle);

    // Keep the value inside 5 s – 60 min whichever unit is picked
    function checkQuestionRange() {
      if (!qToggle.checked) return true;
      const secs = (parseInt(qInput.value, 10) || 0) * (qUnit.value === 'minutes' ? 60 : 1);
      const ok = secs >= 5 && secs <= 3600;
      qInput.setCustomValidity(ok ? '' : 'Time per question must be between 5 seconds and 60 minutes.');
      return ok;
    }
    qInput.addEventListener('input', checkQuestionRange);
    qUnit.addEventListener('change', checkQuestionRange);
    const form = dToggle.closest('form');
    if (form) form.addEventListener('submit', (e) => {
      if (master && !master.checked) return; // untimed exam: nothing to validate
      if (!dToggle.checked && !qToggle.checked) {
        e.preventDefault();
        showError('Turn on at least one timer: the exam duration, the time per question, or both.');
        return;
      }
      if (!checkQuestionRange()) { e.preventDefault(); qInput.reportValidity(); }
    });
    sync();
    checkQuestionRange();
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
