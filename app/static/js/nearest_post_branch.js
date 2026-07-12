document.addEventListener('DOMContentLoaded', function() {
  const checkbox = document.getElementById('nearestPostBranch');
  const deliveryFees = document.getElementById('deliveryFees');
  const governorateSelect = document.getElementById('customerGovernorate');

  let nearestPostFee = null;
  try {
    nearestPostFee = window.nearestPostBranchFee;
  } catch (e) {}

  if (checkbox && deliveryFees && governorateSelect) {
    checkbox.addEventListener('change', function() {
      if (checkbox.checked && nearestPostFee !== null) {
        deliveryFees.value = nearestPostFee;
        governorateSelect.value = 'لأقرب فرع بريد';
      } else {
        deliveryFees.value = 0;
        if (governorateSelect.value === 'لأقرب فرع بريد') {
          governorateSelect.value = '';
        }
      }
      if (typeof window.calculateTotals === 'function') {
        window.calculateTotals();
      }
    });
  }

  if (governorateSelect && checkbox) {
    governorateSelect.addEventListener('change', function() {
      if (governorateSelect.value === 'لأقرب فرع بريد') {
        checkbox.checked = true;
        if (nearestPostFee !== null) {
          deliveryFees.value = nearestPostFee;
        }
      } else {
        checkbox.checked = false;
        deliveryFees.value = 0;
      }
      if (typeof window.calculateTotals === 'function') {
        window.calculateTotals();
      }
    });
  }
});
