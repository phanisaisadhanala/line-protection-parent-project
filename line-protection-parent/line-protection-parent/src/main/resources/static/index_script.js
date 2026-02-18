document.addEventListener('DOMContentLoaded', () => {
  const form        = document.getElementById('projectForm');
  const downloadBtn = document.getElementById('downloadButton');
  const submitBtn   = form.querySelector('button[type="submit"], input[type="submit"], #submitButton');
  const csvInput    = document.getElementById('csvUpload');

  // ===== PRC-025 Synchronous modal wiring =====
  const relayLoad   = document.getElementById('relayLoadbility');
  const modal       = document.getElementById('criteriaModal');
  const btnClose    = document.getElementById('criteriaClose');
  const btnCancel   = document.getElementById('criteriaCancel');
  const btnSave     = document.getElementById('criteriaSave');

  const rowsHost     = document.getElementById('sheetRows');
  const addRowBtn    = document.getElementById('addRowBtn');
  const removeRowBtn = document.getElementById('removeRowBtn');
  const rowCountEl   = document.getElementById('rowCount');
  const totalMVAOut  = document.getElementById('footerTotalMVA');
  const totalMVARIn  = document.getElementById('footerTotalMVAR');
  const totalMWOut   = document.getElementById('footerTotalMW');
  const totalQCalc   = document.getElementById('footerTotalMVARcalc');

  const MAX_ROWS = 16;

  function updateRowCount() {
    rowCountEl.textContent = `${rowsHost.querySelectorAll('.row').length}/${MAX_ROWS}`;
  }

  function openModal() {
    resetModal();              // always start fresh
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';
  }
  function closeModal({resetSelect=false, clear=false} = {}) {
    modal.classList.remove('open');
    document.body.style.overflow = '';
    if (clear) resetModal();
    if (resetSelect && relayLoad) relayLoad.value = '';
  }

  // build a grid cell with an <input>
  function makeCellInput({type='text', step=null, min=null, max=null, placeholder='', readOnly=false, dataAttr=null}) {
    const wrap = document.createElement('div');
    wrap.className = 'cell';
    const input = document.createElement('input');
    input.type = type;
    if (step !== null) input.step = step;
    if (min  !== null) input.min  = min;
    if (max  !== null) input.max  = max;
    if (placeholder)  input.placeholder = placeholder;
    if (readOnly)     input.readOnly = true;
    if (dataAttr)     input.setAttribute(dataAttr, '');
    wrap.appendChild(input);
    return {wrap, input};
  }

  function recalcRow(rowEl) {
    const nameEl = rowEl.querySelector('[data-name]');
    const mvaEl  = rowEl.querySelector('[data-mva]');
    const qtyEl  = rowEl.querySelector('[data-qty]');
    const pfEl   = rowEl.querySelector('[data-pf]');
    const qInEl  = rowEl.querySelector('[data-mvar]');

    const totalEl = rowEl.querySelector('[data-total]');
    const mwEl    = rowEl.querySelector('[data-mw]');
    const qCalcEl = rowEl.querySelector('[data-qcalc]');

    const name = (nameEl?.value || '').trim();
    const mva  = parseFloat(mvaEl?.value);
    const qty  = parseFloat(qtyEl?.value);
    const pf   = parseFloat(pfEl?.value);

    const hasMVA = !Number.isNaN(mva);
    const hasQty = !Number.isNaN(qty);
    const hasPF  = !Number.isNaN(pf);

    // S = MVA × Qty (Total Asynchronous Generators MVA)
    const S = (hasMVA && hasQty) ? mva * qty : NaN;
    totalEl.value = Number.isNaN(S) ? '' : S.toFixed(1);

    // MW (calculated) = S × PF
    const P = (!Number.isNaN(S) && hasPF) ? (S * pf) : NaN;
    mwEl.value = Number.isNaN(P) ? '' : P.toFixed(1);

    // MVAR (calculated) — Excel logic:
    // IF(manual MVAR <> "", "", IF(Name="", "", S*SIN(ACOS(PF))))
    const userQEntered = (qInEl?.value || '').trim() !== '';

    if (userQEntered || name === '' || Number.isNaN(S) || !hasPF || pf < -1 || pf > 1) {
      qCalcEl.value = '';
    } else {
      const Q = S * Math.sin(Math.acos(pf)); // same as S*sqrt(1-pf^2)
      qCalcEl.value = Number.isFinite(Q) ? Q.toFixed(1) : '';
    }
  }

  // recompute FOOTER totals (sums)
  function recalcTotals() {
    let sumS = 0, sumMVARin = 0, sumP = 0, sumQcalc = 0;

    rowsHost.querySelectorAll('.row').forEach(row => {
      sumS      += parseFloat(row.querySelector('[data-total]').value) || 0;   // Total MVA
      sumMVARin += parseFloat(row.querySelector('[data-mvar]').value)  || 0;   // entered MVAR
      sumP      += parseFloat(row.querySelector('[data-mw]').value)    || 0;   // calc MW
      sumQcalc  += parseFloat(row.querySelector('[data-qcalc]').value) || 0;   // calc MVAR
    });

    totalMVAOut.value = sumS ? sumS.toFixed(1) : '';
    totalMVARIn.value = sumMVARin ? sumMVARin.toFixed(1) : '';
    totalMWOut.value  = sumP ? sumP.toFixed(1) : '';
    totalQCalc.value  = sumQcalc ? sumQcalc.toFixed(1) : '';

    updateRowCount();
  }

  // attach listeners for one row
  function wireRow(rowEl) {
    const name = rowEl.querySelector('[data-name]');
    const mva  = rowEl.querySelector('[data-mva]');
    const qty  = rowEl.querySelector('[data-qty]');
    const pf   = rowEl.querySelector('[data-pf]');
    const mvarIn = rowEl.querySelector('[data-mvar]');

    const trigger = () => { recalcRow(rowEl); recalcTotals(); };

    [name, mva, qty, pf, mvarIn].forEach(el => el.addEventListener('input', trigger));
  }

  // add a new row
  function addRow() {
    if (rowsHost.querySelectorAll('.row').length >= MAX_ROWS) return;

    const row = document.createElement('div');
    row.className = 'row';

    // 1) Name  (now tagged with data-name)
    row.appendChild(makeCellInput({type:'text', placeholder:'Gen / Device', dataAttr:'data-name'}).wrap);

    // 2) MVA
    row.appendChild(makeCellInput({type:'number', step:'0.01', min:'0', dataAttr:'data-mva'}).wrap);

    // 3) Qty
    row.appendChild(makeCellInput({type:'number', step:'1', min:'0', dataAttr:'data-qty'}).wrap);

    // 4) Total MVA (readonly)
    row.appendChild(makeCellInput({type:'text', readOnly:true, dataAttr:'data-total'}).wrap);

    // 5) PF (0..1 per your header)
    row.appendChild(makeCellInput({type:'number', step:'0.01', min:'0', max:'1', dataAttr:'data-pf'}).wrap);

    // 6) MVAR (entered)
    row.appendChild(makeCellInput({type:'number', step:'0.01', dataAttr:'data-mvar'}).wrap);

    // 7) MW (calc, readonly)
    row.appendChild(makeCellInput({type:'text', readOnly:true, dataAttr:'data-mw'}).wrap);

    // 8) MVAR calc (readonly)
    row.appendChild(makeCellInput({type:'text', readOnly:true, dataAttr:'data-qcalc'}).wrap);

    rowsHost.appendChild(row);
    wireRow(row);
    recalcRow(row);
    recalcTotals();
  }
  function clearRow(rowEl){
  rowEl.querySelectorAll('input').forEach(i => i.value = '');
}

function removeRow(){
  const rows = rowsHost.querySelectorAll('.row');
  if (rows.length === 0) return;

  // Prefer the row that currently has a focused input
  const focused = document.activeElement && document.activeElement.closest('.row');
  let target = focused instanceof Element ? focused : rows[rows.length - 1];

  if (rows.length === 1){
    // never leave the sheet empty: clear the only row instead
    clearRow(target);
    recalcRow(target);
    recalcTotals();
    rowCountEl.textContent = '1';
    return;
  }

  // Clear (so totals drop immediately), then remove
  clearRow(target);
  rowsHost.removeChild(target);

  // Update UI
  rowCountEl.textContent = rowsHost.querySelectorAll('.row').length;
  recalcTotals();
}

function resetModal() {
  rowsHost.innerHTML = '';
  totalMVAOut.value = '';
  totalMVARIn.value = '';
  totalMWOut.value = '';
  totalQCalc.value = '';
  addRow(); // start with one empty line
}

if (relayLoad) {
  relayLoad.addEventListener('change', e => {
    if (e.target.value === 'PRC_025_Synchronous') openModal();
  });
}
if (addRowBtn) addRowBtn.addEventListener('click', addRow);
if (removeRowBtn) removeRowBtn.addEventListener('click', removeRow);
if (btnClose)  btnClose.addEventListener('click', () => closeModal({resetSelect:true, clear:true}));
if (btnCancel) btnCancel.addEventListener('click', () => closeModal({resetSelect:true, clear:true}));

// Serialize rows when user clicks Save so you can POST them with the form
function serializeRows() {
  const out = [];
  rowsHost.querySelectorAll('.row').forEach(row => {
    const nameEl = row.querySelector('[data-name]');
    const mva  = row.querySelector('[data-mva]').value;
    const qty  = row.querySelector('[data-qty]').value;
    const pf   = row.querySelector('[data-pf]').value;
    const qIn  = row.querySelector('[data-mvar]').value;
    const S    = row.querySelector('[data-total]').value;
    const P    = row.querySelector('[data-mw]').value;
    const Qc   = row.querySelector('[data-qcalc]').value;

    if ((nameEl?.value || '').trim() || mva || qty || pf || qIn || S || P || Qc) {
      out.push({
        name: (nameEl?.value || ''),
        mva:  mva || '',
        qty:  qty || '',
        total: S || '',
        pf:   pf || '',
        q:    qIn || '',
        mw:   P || '',
        qcalc: Qc || ''
      });
    }
  });
  return out;
}

if (btnSave) {
  btnSave.addEventListener('click', () => {
    // light validation: if any numeric provided then require name, MVA, Qty, PF
    let ok = true;
    rowsHost.querySelectorAll('.row').forEach(row => {
      const name = row.querySelector('[data-name]');
      const mva  = row.querySelector('[data-mva]');
      const qty  = row.querySelector('[data-qty]');
      const pf   = row.querySelector('[data-pf]');
      const any  = [name, mva, qty, pf].some(el => (el.value && el.value.toString().trim() !== ''));
      if (any) {
        if (!name.value || !mva.value || !qty.value || !pf.value) {
          ok = false;
          [name, mva, qty, pf].forEach(el => { if (!el.value) el.reportValidity?.(); });
        }
      }
    });
    if (!ok) return;

    window._prc025Rows = serializeRows();
    closeModal();
  });
}

window.addEventListener('keydown', e => {
  if (e.key === 'Escape' && modal && modal.classList.contains('open')) {
    closeModal({resetSelect:true, clear:true});
  }
});
// ===== END modal wiring ===== //

  // Submit flow
  function toggleButtons(showDownload) {
    if (showDownload) {
      downloadBtn.style.display = 'inline-block';
      submitBtn.style.display   = 'none';
      submitBtn.disabled        = false;
    } else {
      downloadBtn.style.display = 'none';
      submitBtn.style.display   = '';
    }
  }

  function clearFormFields() {
    form.reset();
    if (csvInput) csvInput.value = '';
  }

  toggleButtons(false);

  form.addEventListener('submit', async (event) => {
    event.preventDefault();

    const data = {
      relayLocation:                  document.getElementById('relayLocation').value,
      lineNumber:                     document.getElementById('lineNumber').value,
      remoteLocation:                 document.getElementById('remoteLocation').value,
      noninalSystemVoltage:           document.getElementById('noninalSystemVoltage').value,
      breakerRating:                  document.getElementById('breakerRating').value,
      conductorRating:                document.getElementById('conductorRating').value,
      ctrW:                           document.getElementById('ctrW').value,
      ctrX:                           document.getElementById('ctrX').value,
      ptry:                           document.getElementById('ptry').value,
      prcApplicability:               document.getElementById('prcApplicability').value,
      busScheme:                      document.getElementById('busScheme').value,
      secondlines:                    document.getElementById('secondlines').value,
      numberOfTaps:                   document.getElementById('numberOfTaps').value,
      autoXfmrAtRemote:               document.getElementById('autoXfmrAtRemote').value,
      numberOfBreakers:               document.getElementById('numberOfBreakers').value,
      noOfDistributionTransformers:   document.getElementById('noOfDistributionTransformers').value,
      relayLoadbility:                document.getElementById('relayLoadbility').value,
      syncReference:                  document.getElementById('syncReference').value,
      syncSource:                     document.getElementById('syncSource').value,
      hotLineInd:                     document.getElementById('hotLineInd').value,
      vazPtRatio:                     document.getElementById('vazPtRatio').value,
      vbzPtRatio:                     document.getElementById('vbzPtRatio').value,
      vczPtRatio:                     document.getElementById('vczPtRatio').value,
      remoteCTR:                      document.getElementById('remoteCTR').value,
      remoteBFPU:                     document.getElementById('remoteBFPU').value,
      remoteBFGU:                     document.getElementById('remoteBFGU').value
    };

    if (!csvInput || csvInput.files.length === 0) {
      alert('Please choose a CSV file to upload.');
      return;
    }

    // attach PRC-025 rows (flattened)
    if (Array.isArray(window._prc025Rows) && window._prc025Rows.length) {
      data.generatorCount = String(window._prc025Rows.length);
      window._prc025Rows.forEach((r, i) => {
        const n = i + 1;
        data[`generatorName${n}`]        = r.name;
        data[`generatorMVA${n}`]         = r.mva;
        data[`generatorQty${n}`]         = r.qty;
        data[`generatorTotalMVA${n}`]    = r.total;
        data[`generatorRatedPF${n}`]     = r.pf;
        data[`staticReactivePower${n}`]  = r.q;
      });
    }

    const formData = new FormData();
    formData.append('formData', JSON.stringify(data));
    formData.append('csvFile', csvInput.files[0]);

    try {
      submitBtn.disabled = true;

      const res = await fetch('http://localhost:8080/upload', { method: 'POST', body: formData });
      if (!res.ok) throw new Error('Failed to generate Excel file');

      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);

      toggleButtons(true);
      clearFormFields();

      downloadBtn.onclick = () => {
        const a = document.createElement('a');
        a.href = url;
        a.download = 'Updated Line Protection Calculation Sheet.xlsm';
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);

        toggleButtons(false);
        submitBtn.disabled = false;
        submitBtn.focus();
      };
    } catch (err) {
      alert('Error: ' + err.message);
      console.error(err);
      submitBtn.disabled = false;
    }
  });

  form.addEventListener('input',  () => toggleButtons(false));
  form.addEventListener('change', () => toggleButtons(false));
});
