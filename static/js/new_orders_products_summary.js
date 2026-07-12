document.addEventListener('DOMContentLoaded', function() {
  const openBtn = document.getElementById('open-new-products-summary');
  const modal = document.getElementById('new-products-summary-modal');
  const tbody = document.getElementById('new-products-summary-tbody');
  const loadingEl = document.getElementById('new-products-summary-loading');
  const errorEl = document.getElementById('new-products-summary-error');

  function getSelectedValue(name) {
    const el = document.querySelector(`select[name="${name}"]`);
    if (el) return el.value;
    // fallback: read from URL query string if select not present (standalone page)
    try {
      const params = new URLSearchParams(window.location.search);
      const v = params.get(name);
      return v === null ? '' : v;
    } catch (e) {
      return '';
    }
  }

  async function loadData() {
    loadingEl.style.display = 'block';
    errorEl.style.display = 'none';
    tbody.innerHTML = '';
    const year = getSelectedValue('year');
    const month = getSelectedValue('month');
    const params = new URLSearchParams();
    if (year) params.append('year', year);
    if (month) params.append('month', month);
    try {
      const res = await fetch(`/orders/new_products_summary?${params.toString()}`, { credentials: 'same-origin' });
      const data = await res.json();
      loadingEl.style.display = 'none';
      if (!data || !data.success) {
        errorEl.textContent = data && data.error ? data.error : 'حدث خطأ أثناء جلب البيانات';
        errorEl.style.display = 'block';
        return;
      }
      const rows = data.data || [];
      if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="4">لا توجد بيانات لهذا الشهر</td></tr>';
        return;
      }
      rows.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${idx + 1}</td>
          <td class="text-start" dir="rtl">${r.name}</td>
          <td>${r.total_quantity}</td>
          <td>${r.stock}</td>
        `;
        tbody.appendChild(tr);
      });
    } catch (err) {
      loadingEl.style.display = 'none';
      errorEl.textContent = err.message || 'حدث خطأ غير متوقع';
      errorEl.style.display = 'block';
    }
  }

  if (openBtn && modal) {
    openBtn.addEventListener('click', function(e) {
      // show bootstrap modal
      if (typeof bootstrap !== 'undefined') {
        const m = new bootstrap.Modal(modal);
        m.show();
      } else {
        modal.style.display = 'block';
      }
      loadData();
    });
  }
  // If running on the standalone page (no modal), auto-load data
  if (!modal && tbody) {
    loadData();
  }
});
