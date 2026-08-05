// SPARK — main.js


document.addEventListener('DOMContentLoaded', () => {
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
  window.closeModal = id => document.getElementById(id)?.classList.remove('open');
  document.querySelectorAll('.modal-overlay').forEach(overlay => {
    overlay.addEventListener('click', e => {
      if (e.target === overlay) overlay.classList.remove('open');
    });
  });
});
