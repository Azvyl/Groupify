(function(window){
  'use strict';
  const Grouping = {};

  Grouping.sanitizeName = function(name){
    if(name === undefined || name === null) return '';
    return String(name).replace(/\s+/g,' ').trim();
  };

  Grouping.parseColumnNames = function(workbook){
    const firstSheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[firstSheetName];
    if(!ws) return [];
    const XLSX = window.XLSX;
    const aoa = XLSX.utils.sheet_to_json(ws, {header:1, range:0, blankrows:false});
    const first = aoa[0] || [];
    return first.map(c => (c===undefined || c===null) ? '' : String(c).trim());
  };

  Grouping.extractNames = function(workbook, columnName){
    const XLSX = window.XLSX;
    const firstSheetName = workbook.SheetNames[0];
    const ws = workbook.Sheets[firstSheetName];
    if(!ws) return [];
    const aoa = XLSX.utils.sheet_to_json(ws, {header:1, blankrows:true});
    const header = aoa[0] || [];
    const colIndex = header.findIndex(h => String(h).trim() === String(columnName).trim());
    if(colIndex === -1) return [];
    const out = [];
    for(let r=1;r<aoa.length;r++){
      const row = aoa[r] || [];
      const raw = row[colIndex];
      const name = Grouping.sanitizeName(raw);
      const rowValues = {};
      for(let c=0;c<header.length;c++){
        const key = (header[c] === undefined || header[c] === null) ? '' : String(header[c]).trim();
        rowValues[key] = (row[c] === undefined || row[c] === null) ? '' : row[c];
      }
      out.push({ name: name, rowIndex: r+1, originalIndex: r-1, rowValues: rowValues });
    }
    return out.filter(x => x.name !== '');
  };

  Grouping.seededRNG = function(seed){
    let h = 2166136261 >>> 0;
    const s = String(seed);
    for(let i=0;i<s.length;i++){
      h = Math.imul(h ^ s.charCodeAt(i), 16777619);
    }
    let state = h >>> 0;
    return function(){
      state += 0x6D2B79F5; state = state >>> 0;
      let t = Math.imul(state ^ state >>> 15, 1 | state);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  };

  Grouping.shuffleArray = function(arr, rng){
    const out = arr.slice();
    let random = rng || Math.random;
    for(let i=out.length-1;i>0;i--){
      const j = Math.floor(random() * (i+1));
      const tmp = out[i]; out[i] = out[j]; out[j] = tmp;
    }
    return out;
  };

  Grouping.groupDeterministic = function(namesArray, numGroups, opts = {}){
    if(numGroups < 1) return [];
    const n = namesArray.length;
    const groups = Array.from({length:numGroups}, ()=>[]);
    if(opts.preserveOrder){
      const base = Math.floor(n/numGroups);
      const rem = n % numGroups;
      let idx = 0;
      for(let g=0;g<numGroups;g++){
        const size = base + (g < rem ? 1 : 0);
        for(let k=0;k<size;k++){
          if(idx < n) groups[g].push(namesArray[idx++]);
        }
      }
    } else {
      for(let i=0;i<n;i++){
        const g = i % numGroups;
        groups[g].push(namesArray[i]);
      }
    }
    return groups;
  };

  Grouping.groupRandom = function(namesArray, numGroups, seed, _opts = {}){
    let rng = Math.random;
    if(seed !== undefined && seed !== null && seed !== ''){
      rng = Grouping.seededRNG(seed);
    }
    const shuffled = Grouping.shuffleArray(namesArray, rng);
    const groups = Array.from({length:numGroups}, ()=>[]);
    for(let i=0;i<shuffled.length;i++){
      const g = i % numGroups;
      groups[g].push(shuffled[i]);
    }
    return groups;
  };

  Grouping.formatGroupsAsTableRows = function(groupsArray, extras){
    extras = Array.isArray(extras) ? extras : [];
    const rows = [];
    for(let i=0;i<groupsArray.length;i++){
      const grp = groupsArray[i];
      const groupName = `Group ${i+1}`;
      for(const member of grp){
        const base = { Group: groupName, Student: member.name, OriginalRow: member.rowIndex };
        const values = member.rowValues || {};
        for(const ex of extras){
          base[ex] = (ex in values) ? values[ex] : '';
        }
        rows.push(base);
      }
    }
    return rows;
  };

  Grouping.buildWorkbookForExport = function(groupsArray, opts = { separateSheets: false }){
    const XLSX = window.XLSX;
    const wb = XLSX.utils.book_new();
    const extras = (opts && Array.isArray(opts.extras)) ? opts.extras : [];
    const rows = Grouping.formatGroupsAsTableRows(groupsArray, extras);
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Groups');
    if(opts.separateSheets){
      for(let i=0;i<groupsArray.length;i++){
        const grows = Grouping.formatGroupsAsTableRows([groupsArray[i]], extras).map(r => ({ Student: r.Student, OriginalRow: r.OriginalRow, ...extras.reduce((acc,e)=>{ acc[e]=r[e]; return acc; }, {}) }));
        const wsi = XLSX.utils.json_to_sheet(grows);
        XLSX.utils.book_append_sheet(wb, wsi, `Group ${i+1}`);
      }
    }
    return wb;
  };

  window.Grouping = Grouping;
})(window);
