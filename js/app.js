(function(){
  'use strict';
  const G = window.Grouping;
  const fileInput = document.getElementById('file-input');
  const dropZone = document.getElementById('drop-zone');
  const colSelect = document.getElementById('col-select');
  const modeGroups = document.getElementById('mode-groups');
  const modeSize = document.getElementById('mode-size');
  const groupsCount = document.getElementById('groups-count');
  const groupSize = document.getElementById('group-size');
  const algoSelect = document.getElementById('algo-select');
  const seedInput = document.getElementById('seed-input');
  const seedRow = document.getElementById('seed-row');
  const preserveOrder = document.getElementById('preserve-order');
  const splitBtn = document.getElementById('split-btn');
  const exportXlsx = document.getElementById('export-xlsx');
  const exportCsv = document.getElementById('export-csv');
  const results = document.getElementById('results');
  const summary = document.getElementById('summary');
  const fileError = document.getElementById('file-error');
  const extraSelect = document.getElementById('extra-select');

  let currentWorkbook = null;
  let currentGroups = null;

  function showError(msg){
    fileError.textContent = msg;
  }
  function clearError(){ fileError.textContent = ''; }

  function setControlsForLoaded(has){
    colSelect.disabled = !has;
    splitBtn.disabled = !has;
  }

  function resetState(){
    currentWorkbook = null; currentGroups = null;
    colSelect.innerHTML = '';
    results.innerHTML = ''; summary.innerHTML = '';
    exportXlsx.disabled = true; exportCsv.disabled = true;
    splitBtn.disabled = true;
  }

  function handleFile(file){
    clearError();
    const name = file.name || '';
    if(!/\.xlsx$|\.xls$|\.csv$/i.test(name)){
      showError('Unsupported file type. Please upload .xlsx, .xls or .csv');
      return;
    }
    const reader = new FileReader();
    reader.onload = function(e){
      try{
        const data = e.target.result;
        const wb = XLSX.read(data, {type:'array'});
        currentWorkbook = wb;
        populateColumnsFromWorkbook(wb);
        setControlsForLoaded(true);
        const firstSheet = wb.SheetNames[0];
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[firstSheet], {header:1, blankrows:false});
        if(aoa.length > 5001) showError('Large file detected (>5000 rows) — processing may be slow.');
      }catch(err){
        console.error(err);
        showError('Failed to parse file. Ensure it is a valid Excel or CSV file.');
      }
    };
    reader.onerror = function(){ showError('Failed to read file'); };
    reader.readAsArrayBuffer(file);
  }

  function populateColumnsFromWorkbook(wb){
    const cols = G.parseColumnNames(wb) || [];
    colSelect.innerHTML = '';
    extraSelect.innerHTML = '';
    cols.forEach(c=>{
      const opt = document.createElement('option'); opt.value = c; opt.textContent = c || '(blank)';
      colSelect.appendChild(opt);
      const opt2 = document.createElement('option'); opt2.value = c; opt2.textContent = c || '(blank)';
      extraSelect.appendChild(opt2);
    });
    extraSelect.disabled = false;
  }

  function getSelectedExtras(){
    if(!extraSelect) return [];
    return Array.from(extraSelect.selectedOptions).map(o => o.value).filter(v => v && v.trim() !== '');
  }

  dropZone.addEventListener('click', (e)=>{
    if(e && e.target === fileInput) return;
    fileInput.click();
  });
  dropZone.addEventListener('dragover', (e)=>{ e.preventDefault(); dropZone.classList.add('dragover'); });
  dropZone.addEventListener('dragleave', ()=>{ dropZone.classList.remove('dragover'); });
  dropZone.addEventListener('drop', (e)=>{ e.preventDefault(); dropZone.classList.remove('dragover'); const f = e.dataTransfer.files[0]; if(f) handleFile(f); });

  fileInput.addEventListener('change', (e)=>{ const f = e.target.files && e.target.files[0]; if(f) handleFile(f); });

  function updateModeInputs(){
    const byGroups = modeGroups.checked;
    groupsCount.disabled = !byGroups;
    groupSize.disabled = byGroups;
  }
  modeGroups.addEventListener('change', updateModeInputs);
  modeSize.addEventListener('change', updateModeInputs);

  algoSelect.addEventListener('change', ()=>{
    const val = algoSelect.value;
    seedRow.style.display = (val === 'random') ? 'block' : 'none';
  });

  algoSelect.value = 'random';
  seedRow.style.display = 'block';

  function validateInputs(){
    clearError();
    if(!currentWorkbook) return false;
    const col = colSelect.value;
    if(!col) { showError('Please select a column'); return false; }
    const names = G.extractNames(currentWorkbook, col);
    if(!names || names.length === 0){ showError('Selected column contains no names'); return false; }
    const total = names.length;
    if(modeGroups.checked){
      const k = parseInt(groupsCount.value,10);
      if(isNaN(k) || k < 1){ showError('Number of groups must be a positive integer'); return false; }
      if(k > total){ showError('Number of groups cannot exceed number of names'); return false; }
    } else {
      const sz = parseInt(groupSize.value,10);
      if(isNaN(sz) || sz < 1){ showError('Group size must be a positive integer'); return false; }
    }
    return true;
  }

  splitBtn.addEventListener('click', ()=>{
    if(!validateInputs()) return;
    const col = colSelect.value;
    const names = G.extractNames(currentWorkbook, col);
    let numGroups;
    if(modeGroups.checked){ numGroups = parseInt(groupsCount.value,10); }
    else { const sz = parseInt(groupSize.value,10); numGroups = Math.max(1, Math.ceil(names.length / sz)); }
    const algo = algoSelect.value;
    if(algo === 'deterministic'){
      currentGroups = G.groupDeterministic(names, numGroups, { preserveOrder: preserveOrder.checked });
    } else {
      const seed = seedInput.value || undefined;
      currentGroups = G.groupRandom(names, numGroups, seed, {});
    }
    renderResults(currentGroups);
    exportXlsx.disabled = false; exportCsv.disabled = false;
  });

  function renderResults(groups){
    results.innerHTML = '';
    let total = 0;
    for(let i=0;i<groups.length;i++) total += groups[i].length;
    const algo = algoSelect.value;
    const seed = seedInput.value || '';
    const mode = modeGroups.checked ? `${groups.length} groups` : `~${groupSize.value} per group`;
    summary.innerHTML = `<strong>Total:</strong> ${total} students — <strong>Mode:</strong> ${mode} — <strong>Algorithm:</strong> ${algo}${seed? ' — seed: ' + seed : ''}`;
    for(let i=0;i<groups.length;i++){
      const g = groups[i];
      const card = document.createElement('section'); card.className = 'group-card';
      const h3 = document.createElement('h3'); h3.textContent = `Group ${i+1} (${g.length})`;
      card.appendChild(h3);
      const ul = document.createElement('ul');
      for(const m of g){ const li = document.createElement('li'); li.textContent = m.name; ul.appendChild(li); }
      card.appendChild(ul);
      results.appendChild(card);
    }
  }

  function timestamp(){
    const d = new Date();
    const pad = (n)=> String(n).padStart(2,'0');
    return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  }

  exportXlsx.addEventListener('click', ()=>{
    if(!currentGroups) return;
    const extras = getSelectedExtras();
    const wb = G.buildWorkbookForExport(currentGroups, { separateSheets: false, extras: extras });
    const fname = `groupify-results-${timestamp()}.xlsx`;
    XLSX.writeFile(wb, fname);
  });

  exportCsv.addEventListener('click', ()=>{
    if(!currentGroups) return;
    const extras = getSelectedExtras();
    const rows = G.formatGroupsAsTableRows(currentGroups, extras);
    const ws = XLSX.utils.json_to_sheet(rows);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a'); a.href = url; a.download = `groupify-results-${timestamp()}.csv`; document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  });

  resetState();

})();
