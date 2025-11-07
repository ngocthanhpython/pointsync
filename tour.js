(function () {
  // Helper: tìm heading theo text (cho "Danh sách học sinh chưa có điểm")
  function findHeadingStartsWith(txt) {
    const hs = Array.from(document.querySelectorAll('h2'));
    return hs.find(h => (h.textContent || '').trim().startsWith(txt));
  }

  // Danh sách các bước theo đúng thứ tự yêu cầu
  const steps = [
    { // 1. File nguồn (.xlsx) - nút Chọn tệp
      target: () => document.getElementById('src_xlsx'),
      title: 'Chọn file đã chấm trắc nghiệm',
      body: 'Hãy chọn tệp Excel chứa SBD và điểm đã chấm.'
    },
    { // 2. Cột SBD (file nguồn)
      target: () => document.getElementById('col_sbd'),
      title: 'Cột SBD (file nguồn)',
      body: 'Chọn cột có chứa Số Báo Danh trong file chấm.'
    },
    { // 3. Cột Điểm (file nguồn)
      target: () => document.getElementById('col_score'),
      title: 'Cột Điểm (file nguồn)',
      body: 'Chọn cột có chứa điểm (0–10).'
    },
    { // 4. File lớp (.xlsx) - nút Chọn tệp
      target: () => document.getElementById('dst_xlsx'),
      title: 'Chọn file lớp muốn đồng bộ',
      body: 'Hãy chọn tệp Excel danh sách lớp cần chép điểm.'
    },
    { // 5. 3. Tìm kiếm cột cần chép
      target: () => document.getElementById('analyze_btn'),
      title: 'Hiển thị các cột điểm',
      body: 'Bấm để tự động phát hiện các cột điểm của file lớp!'
    },
    { // 6. Cột điểm đích (áp cho tất cả sheet)
      target: () => document.getElementById('dest_label'),
      title: 'Chọn cột cần chép',
      body: 'Chọn cột điểm đích để chép vào tất cả các sheet lớp.'
    },
    { // 7. 4. Đồng bộ điểm
      target: () => document.getElementById('run_btn'),
      title: 'Đồng bộ điểm',
      body: 'Nhấn vào đây để thực hiện việc đồng bộ điểm.'
    },
    { // 8. Danh sách học sinh chưa có điểm
      target: () => {
        const h = findHeadingStartsWith('Danh sách học sinh chưa có điểm');
        return h ? (h.closest('section') || h) : null;
      },
      title: 'Danh sách học sinh chưa có điểm',
      body: 'Khu vực này hiển thị những học sinh chưa được chép điểm.'
    }
  ];

  // UI phần tour
  let overlay, highlight, tooltip, idx = 0;

  function ensureNodes() {
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.className = 'tour-overlay';
      document.body.appendChild(overlay);
      requestAnimationFrame(() => overlay.classList.add('show'));
    }
    if (!highlight) {
      highlight = document.createElement('div');
      highlight.className = 'tour-highlight';
      document.body.appendChild(highlight);
    }
    if (!tooltip) {
      tooltip = document.createElement('div');
      tooltip.className = 'tour-tooltip';
      tooltip.innerHTML = `
        <h3></h3>
        <p></p>
        <div class="tour-actions">
          <button class="btn" data-act="skip">Bỏ qua</button>
          <button class="btn" data-act="prev">Trước</button>
          <button class="btn primary" data-act="next">Tiếp</button>
        </div>
      `;
      document.body.appendChild(tooltip);
      tooltip.addEventListener('click', (e) => {
        const act = e.target && e.target.getAttribute('data-act');
        if (!act) return;
        if (act === 'skip') endTour(true);
        if (act === 'prev') prevStep();
        if (act === 'next') nextStep();
      });
    }
  }

  function positionStep() {
    const step = steps[idx];
    const el = step.target && step.target();
    if (!el) return nextStep(); // nếu chưa có element (do UI chưa sẵn), nhảy bước

    const r = el.getBoundingClientRect();
    const pad = 8;

    // Highlight box
    highlight.style.left   = (r.left - pad) + 'px';
    highlight.style.top    = (r.top - pad) + 'px';
    highlight.style.width  = (r.width + pad*2) + 'px';
    highlight.style.height = (r.height + pad*2) + 'px';

    // Tooltip content
    tooltip.querySelector('h3').textContent = step.title || 'Hướng dẫn';
    tooltip.querySelector('p').textContent  = step.body || '';

    // Tooltip position
    const tw = Math.min(320, Math.max(260, r.width));
    tooltip.style.maxWidth = tw + 'px';

    const spaceBelow = window.innerHeight - (r.bottom + 12);
    const spaceAbove = r.top - 12;
    let top;
    if (spaceBelow >= 120 || spaceBelow >= spaceAbove) {
      top = r.bottom + 12;
    } else {
      top = r.top - tooltip.offsetHeight - 12;
      if (top < 10) top = 10;
    }
    let left = Math.max(10, Math.min(r.left, window.innerWidth - tw - 10));

    tooltip.style.top = top + 'px';
    tooltip.style.left = left + 'px';

    // Buttons state
    tooltip.querySelector('[data-act="prev"]').disabled = (idx === 0);
    tooltip.querySelector('[data-act="next"]').textContent =
      (idx === steps.length - 1) ? 'Hoàn tất' : 'Tiếp';

    // Scroll into view nếu cần
    const needsScroll = r.top < 0 || r.bottom > window.innerHeight;
    if (needsScroll) {
      el.scrollIntoView({ behavior: 'smooth', block: 'center' });
      setTimeout(positionStep, 280);
    }
  }

  function nextStep() {
    if (idx < steps.length - 1) { idx++; positionStep(); }
    else { endTour(); }
  }
  function prevStep() {
    if (idx > 0) { idx--; positionStep(); }
  }
  function endTour(skip=false) {
    overlay && overlay.remove(); overlay = null;
    highlight && highlight.remove(); highlight = null;
    tooltip && tooltip.remove(); tooltip = null;
    if (!skip) try { localStorage.setItem('pointsync_tour_done', '1'); } catch(e){}
  }

  function startTour(force=false) {
    try {
      if (!force && localStorage.getItem('pointsync_tour_done') === '1') return;
    } catch (e) {}
    idx = 0;
    ensureNodes();
    positionStep();
    window.addEventListener('resize', positionStep, { passive: true });
    window.addEventListener('scroll', positionStep, { passive: true });
  }

  // Expose để gọi lại khi cần
  window.startPointSyncTour = () => startTour(true);

  // Tự khởi động tour 1 lần ở lần dùng đầu
  window.addEventListener('DOMContentLoaded', () => startTour(false));
})();
