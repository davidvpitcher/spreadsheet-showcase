export function analyzeColumnTypes(hot) {
    const cols = hot.countCols();
    const rows = hot.countRows();
    const types = [];
  
    for (let col = 0; col < cols; col++) {
      let num = 0, date = 0, text = 0;
  
      for (let row = 1; row < rows; row++) {
        const val = hot.getDataAtCell(row, col);
        if (!val) continue;
  
        const str = String(val).trim();
        if (!isNaN(str) && str.match(/^[\d.]+$/)) { num++; continue; }
        if (str.match(/^\d{4}-\d{2}-\d{2}$/) || str.match(/^\d{1,2}\/\d{1,2}\/\d{2,4}$/) || str.match(/^\d{4}\/\d{1,2}\/\d{1,2}$/)) {
          if (!isNaN(Date.parse(str))) { date++; continue; }
        }
        text++;
      }
  
      const total = num + date + text;
      types.push(total === 0 ? 'Empty' : num / total > 0.8 ? 'Number' : date / total > 0.8 ? 'Date' : 'Text');
    }
  
    return types;
  }
  
  export function countKeywordMatches(hot, keyword) {
    const data = hot.getData();
    let matchCount = 0;
  
    data.forEach(row => {
      if (row.some(cell => typeof cell === 'string' && cell.includes(keyword))) {
        matchCount++;
      }
    });
  
    return matchCount;
  }
  
  
  export function enableSearchPlugin(hot) {
    hot.updateSettings({
      search: true,
      columnSorting: {
        indicator: true,
        sortEmptyCells: false
      },
      manualColumnResize: true
    });
  }
  
  export function applySearch(hot, query) {
    const plugin = hot.getPlugin('search');
    const results = plugin.query(query);
    hot.render();
  
    if (results.length > 0) {
      const firstMatch = results[0];
      hot.scrollViewportTo(firstMatch.row, firstMatch.col, true, true);
      hot.selectCell(firstMatch.row, firstMatch.col);
      return true;
    }
    return false;
  }
  
  export function clearSearch(hot) {
    const plugin = hot.getPlugin('search');
    plugin.query('');
    hot.render();
  }
  
  export function toggleFreezeTopRow(hot) {
    const isFrozen = hot.getSettings().fixedRowsTop === 1;
    hot.updateSettings({ fixedRowsTop: isFrozen ? 0 : 1 });
    hot.render();
  }
  
  export function toggleFreezeFirstColumn(hot) {
    const isFrozen = hot.getSettings().fixedColumnsLeft === 1;
    hot.updateSettings({ fixedColumnsLeft: isFrozen ? 0 : 1 });
    hot.render();
  }
  
  export function resetFormatting(sheet, autoSaveCallback) {
    sheet.cellClassMap.clear();
    sheet.restoreFormatting();
    sheet.resetColumnOrder();
    if (typeof autoSaveCallback === 'function') {
      autoSaveCallback();
    }
  }
  
  export function toggleWrapText(sheet, selection, saveCallback) {
    const hot = sheet.hotInstance;
  
    
  if (!Array.isArray(selection) || selection.length === 0 || !Array.isArray(selection[0])) {
    console.warn('[setAlignment] No valid selection to align.');
    return;
  }

    selection.forEach(([startRow, startCol, endRow, endCol]) => {
      const rowStart = Math.min(startRow, endRow);
      const rowEnd = Math.max(startRow, endRow);
      const colStart = Math.min(startCol, endCol);
      const colEnd = Math.max(startCol, endCol);
  
      for (let row = rowStart; row <= rowEnd; row++) {
        if (row < 0) continue; // 🛡️ Skip invalid row (e.g. header)
        for (let col = colStart; col <= colEnd; col++) {
          if (col < 0) continue; // 🛡️ Skip invalid column (e.g. header)
  
          const key = `${row},${col}`;
          const current = sheet.cellClassMap.get(key) || '';
          const classList = new Set(current.trim().split(/\s+/).filter(Boolean));
  
          if (classList.has('wrap-text')) {
            classList.delete('wrap-text');
          } else {
            classList.add('wrap-text');
          }
  
          const finalClass = [...classList].join(' ');
          if (finalClass) {
            sheet.cellClassMap.set(key, finalClass);
            hot.setCellMeta(row, col, 'className', finalClass);
          } else {
            sheet.cellClassMap.delete(key);
            hot.removeCellMeta(row, col, 'className');
          }
        }
      }
    });
  
    sheet.restoreFormatting();
    if (typeof saveCallback === 'function') {
      saveCallback();
    }
  }
  

  
  export function setFontSize(sheet, selection, sizePx, saveCallback = null) {
    const hot = sheet.hotInstance;
  
    
  if (!Array.isArray(selection) || selection.length === 0 || !Array.isArray(selection[0])) {
    console.warn('[setAlignment] No valid selection to align.');
    return;
  }
  
    selection.forEach(([startRow, startCol, endRow, endCol]) => {
      const rowStart = Math.min(startRow, endRow);
      const rowEnd = Math.max(startRow, endRow);
      const colStart = Math.min(startCol, endCol);
      const colEnd = Math.max(startCol, endCol);
  
      for (let row = rowStart; row <= rowEnd; row++) {
        for (let col = colStart; col <= colEnd; col++) {
          if (row < 0 || col < 0) continue; // ✅ skip invalid header selections
      
          const key = `${row},${col}`;
          const className = `font-${sizePx}`;
          const existing = sheet.cellClassMap.get(key) || '';
          const classList = new Set(existing.trim().split(/\s+/).filter(Boolean));
      
          // Remove previous font size
          for (const cls of classList) {
            if (/^font-\d+$/.test(cls)) classList.delete(cls);
          }
      
          classList.add(className);
          const finalClass = [...classList].join(' ');
      
          sheet.cellClassMap.set(key, finalClass);
          hot.setCellMeta(row, col, 'className', finalClass);
        }
      }
      
    });
  
    injectFontSizeStyle(sizePx);
    hot.render();
    if (saveCallback) saveCallback();
  }

  
  export function injectFontSizeStyle(sizePx) {
    const className = `font-${sizePx}`;
    const styleTagId = 'dynamic-fontsize-styles';
    let styleTag = document.getElementById(styleTagId);
  
    if (!styleTag) {
      styleTag = document.createElement('style');
      styleTag.id = styleTagId;
      document.head.appendChild(styleTag);
    }
  
    if (!styleTag.textContent.includes(`.${className}`)) {
      styleTag.textContent += `
  td.${className} {
    font-size: ${sizePx}px !important;
  }
  `;
    }
  }
  export function setVerticalAlignment(sheet, selection, alignment, saveCallback = null) {
    const hot = sheet.hotInstance;
  
    if (!Array.isArray(selection) || selection.length === 0 || !Array.isArray(selection[0])) {
      console.warn('[setVerticalAlignment] Invalid or empty selection:', selection);
      return;
    }

    
    selection.forEach(([startRow, startCol, endRow, endCol]) => {
      const rowStart = Math.max(0, Math.min(startRow, endRow));
      const rowEnd = Math.max(0, Math.max(startRow, endRow));
      const colStart = Math.max(0, Math.min(startCol, endCol));
      const colEnd = Math.max(0, Math.max(startCol, endCol));
  
      for (let row = rowStart; row <= rowEnd; row++) {
        for (let col = colStart; col <= colEnd; col++) {
          const key = `${row},${col}`;
          const className = `valign-${alignment}`;
          const existing = sheet.cellClassMap.get(key) || '';
          const classList = new Set(existing.trim().split(/\s+/).filter(Boolean));
  
          ['valign-top', 'valign-middle', 'valign-bottom'].forEach(cls => classList.delete(cls));
          classList.add(className);
          const finalClass = [...classList].join(' ');
  
          sheet.cellClassMap.set(key, finalClass);
          hot.setCellMeta(row, col, 'className', finalClass);
        }
      }
    });
  
    injectVerticalAlignmentStyle();
    hot.render();
    if (saveCallback) saveCallback();
  }
  
  export function injectVerticalAlignmentStyle() {
    const styleTagId = 'dynamic-valign-styles';
    let styleTag = document.getElementById(styleTagId);
  
    if (!styleTag) {
      styleTag = document.createElement('style');
      styleTag.id = styleTagId;
      document.head.appendChild(styleTag);
    }
  
    const styles = `
      td.valign-top    { vertical-align: top !important; }
      td.valign-middle { vertical-align: middle !important; }
      td.valign-bottom { vertical-align: bottom !important; }
    `;
  
    if (!styleTag.textContent.includes('valign-top')) {
      styleTag.textContent += styles;
    }
  }
  
  export function highlightKeywordRows(sheet, autoSaveCallback) {
    const hot = sheet.hotInstance;
    const map = sheet.cellClassMap;
    const rowCount = hot.countRows();
    const colCount = hot.countCols();
  
    for (let row = 0; row < rowCount; row++) {
      const rowData = hot.getDataAtRow(row);
      const hasKeyword = rowData.some(cell => typeof cell === 'string' && cell.includes('KEYWORD_ROW_HIGHLIGHT'));
  
      if (hasKeyword) {
        for (let col = 0; col < colCount; col++) {
          map.set(`${row},${col}`, 'full-row-error');
        }
      }
    }
  
    sheet.restoreFormatting();
    if (typeof autoSaveCallback === 'function') {
      autoSaveCallback();
    }
  }
  
  export function highlightAlternatingRows(sheet, autoSaveCallback) {
    const hot = sheet.hotInstance;
    const map = sheet.cellClassMap;
    const rowCount = hot.countRows();
    const colCount = hot.countCols();
  
    for (let row = 0; row < rowCount; row++) {
      for (let col = 0; col < colCount; col++) {
        const key = `${row},${col}`;
        const existing = map.get(key) || '';
  
        // Remove old `color-` and `alt-row` classes
        const cleaned = existing
          .split(/\s+/)
          .filter(cls => !cls.startsWith('color-') && cls !== 'alt-row')
          .join(' ');
  
        // Add zebra style to every other row
        if (row % 2 === 1) {
          const newClass = (cleaned + ' alt-row').trim();
          map.set(key, newClass);
        } else {
          // Even rows: save cleaned class or delete if empty
          if (cleaned) {
            map.set(key, cleaned);
          } else {
            map.delete(key);
          }
        }
      }
    }
  
    sheet.restoreFormatting();
  
    if (typeof autoSaveCallback === 'function') {
      autoSaveCallback();
    }
  }
  
  
  export function highlightColumn2(sheet, autoSaveCallback) {
    const hot = sheet.hotInstance;
    const map = sheet.cellClassMap;
    const rowCount = hot.countRows();
  
    for (let row = 0; row < rowCount; row++) {
      const val = hot.getDataAtCell(row, 2);
      if (val && val.toString().includes('ERROR')) {
        const key = `${row},2`;
        const existing = map.get(key);
        const classes = new Set((existing || '').split(' ').filter(c => c));
        classes.add('error-cell');
        map.set(key, Array.from(classes).join(' '));
      }
    }
  
    sheet.restoreFormatting();
    if (typeof autoSaveCallback === 'function') {
      autoSaveCallback();
    }
  }
  
  export function exportToCSV(sheetRef) {
    const data = sheetRef.getData();
    const csv = data
      .map(row =>
        row
          .map(cell => {
            if (cell == null) return '';
            const cellStr = String(cell).replace(/"/g, '""');
            return `"${cellStr}"`;
          })
          .join(',')
      )
      .join('\n');
  
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
  
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', 'output.csv');
    link.click();
  
    URL.revokeObjectURL(url);
  }
  

  function reinjectDynamicColorStyles(cellClassMap) {
    const injectedStyles = new Set();
    let styleTag = document.getElementById('dynamic-color-styles');
    if (!styleTag) {
      styleTag = document.createElement('style');
      styleTag.id = 'dynamic-color-styles';
      document.head.appendChild(styleTag);
    }
  
    for (const [, classString] of cellClassMap) {
      const classes = classString.trim().split(/\s+/);
      for (const className of classes) {
        if (injectedStyles.has(className)) continue;
        injectedStyles.add(className);
  
        if (className.startsWith('color-')) {
          const hex = `#${className.slice(6)}`;
          styleTag.textContent += `
  td.${className} {
    background-color: ${hex} !important;
  }
  `;
        }
  
        if (className.startsWith('text-')) {
          const hex = `#${className.slice(5)}`;
          styleTag.textContent += `
  td.${className} {
    color: ${hex} !important;
  }
  `;
        }
      }
    }
  }
  
  
export function restoreSavedSpreadsheet(sheet, storageKey = 'tool60_state') {
  const raw = localStorage.getItem(storageKey);
  if (!raw) return;

  try {
    const state = JSON.parse(raw);
    const hot = sheet.hotInstance;

    // Enable borders before applying any data
    hot.updateSettings({
      manualColumnResize: true,
      manualRowResize: true,
      fixedRowsTop: state.frozenTop || 0,
      fixedColumnsLeft: state.frozenLeft || 0,
      colWidths: state.colWidths || [],
      rowHeights: state.rowHeights || [],
      customBorders: true
    });

    // Load spreadsheet data (must happen *after* enabling borders)
    hot.loadData(state.data);

    if (state.customBorders && Array.isArray(state.customBorders)) {
      const plugin = hot.getPlugin('customBorders');
      plugin.clearBorders(); // optional
      for (const border of state.customBorders) {
        plugin.setBorders([[border.row, border.col]], {
          ...(border.top && { top: border.top }),
          ...(border.right && { right: border.right }),
          ...(border.bottom && { bottom: border.bottom }),
          ...(border.left && { left: border.left }),
        });
      }
     // console.log('Restoring borders:', JSON.stringify(state.customBorders, null, 2));
    }
    

    // Restore cell class formatting
    if (state.cellClasses) {
      sheet.cellClassMap = new Map(state.cellClasses);
      sheet.restoreFormatting();
      reinjectDynamicColorStyles(sheet.cellClassMap);
    }

    hot.render();
  } catch (e) {
    console.warn("Failed to restore formatting:", e);
  }
}


export function saveSpreadsheet(sheet, storageKey = 'tool60_state') {
  const hot = sheet.hotInstance;
  const colCount = hot.countCols();
  const rowCount = hot.countRows();
  const colWidths = Array.from({ length: colCount }, (_, i) => hot.getColWidth(i));
  const rowHeights = Array.from({ length: rowCount }, (_, i) => hot.getRowHeight(i));
  const bordersPlugin = hot.getPlugin('customBorders');
  const savedBorders = [];

  if (bordersPlugin && typeof bordersPlugin.getBorders === 'function') {
    const borderData = bordersPlugin.getBorders();
    for (const item of borderData) {
      savedBorders.push({
        row: item.row,
        col: item.col,
        ...(item.top && { top: item.top }),
        ...(item.right && { right: item.right }),
        ...(item.bottom && { bottom: item.bottom }),
        ...(item.left && { left: item.left })
      });
    }
  
   // console.log('Saving borders (flat schema):', JSON.stringify(savedBorders, null, 2));
  }
  
  
  
  const state = {
    data: hot.getData(),
    colWidths,
    rowHeights,
    frozenTop: hot.getSettings().fixedRowsTop,
    frozenLeft: hot.getSettings().fixedColumnsLeft,
    cellClasses: Array.from(sheet.cellClassMap.entries()),
    customBorders: savedBorders
  };

  localStorage.setItem(storageKey, JSON.stringify(state));
}

  
  export function processData(inputHot, outputHot) {
    try {
      const inputData = inputHot.getData();
      const processed = inputData.map(row =>
        row.map(cell => typeof cell === 'string' ? cell.toUpperCase() : cell)
      );
      outputHot.loadData(processed);
    } catch (e) {
      console.error('Processing error:', e);
      alert('An error occurred during processing.');
    }
  }
  
  export function clearRecolors(sheet, autoSaveCallback) {
    if (sheet && sheet.clearRecolors) {
      sheet.clearRecolors();
    }
  
    const styleTag = document.getElementById('dynamic-color-styles');
    if (styleTag) {
      styleTag.remove();
    }
  
    if (typeof autoSaveCallback === 'function') {
      autoSaveCallback();
    }
  }
  
  export function fillDownSelection(hot) {
    const selected = hot.getSelected();
    if (!selected || selected.length === 0) {
      alert("Please select a range to fill down.");
      return;
    }
  
    const [startRow, startCol, endRow, endCol] = selected[selected.length - 1];
  
    for (let col = startCol; col <= endCol; col++) {
      const sourceValues = [];
  
      for (let row = startRow; row <= endRow; row++) {
        const val = hot.getDataAtCell(row, col);
        sourceValues.push(val);
      }
  
      const seeds = sourceValues.filter(v => v !== null && v !== '');
      if (seeds.length === 0) continue;
  
      let patternType = 'repeat';
      let step = 0;
      let textPrefix = '';
      let textNumberStart = null;
  
      const [first, second] = seeds;
      const firstNum = parseFloat(first);
      const secondNum = parseFloat(second);
  
      if (seeds.length >= 2) {
        if (!isNaN(firstNum) && !isNaN(secondNum)) {
          patternType = 'number-sequence';
          step = secondNum - firstNum;
        } else if (!isNaN(Date.parse(first)) && !isNaN(Date.parse(second))) {
          patternType = 'date-sequence';
          step = new Date(second).getTime() - new Date(first).getTime();
        } else {
          const match1 = typeof first === 'string' ? first.match(/^(.*?)(\d+)$/) : null;
          const match2 = typeof second === 'string' ? second.match(/^(.*?)(\d+)$/) : null;
          if (match1 && match2 && match1[1] === match2[1]) {
            patternType = 'text-sequence';
            textPrefix = match1[1];
            textNumberStart = parseInt(match2[2]);
            step = parseInt(match2[2]) - parseInt(match1[2]);
          }
        }
      }
  
      let lastKnown = seeds[seeds.length - 1];
      for (let rowOffset = seeds.length; (startRow + rowOffset) <= endRow; rowOffset++) {
        let row = startRow + rowOffset;
        let valueToSet = lastKnown;
  
        if (patternType === 'number-sequence') {
          valueToSet = (parseFloat(lastKnown) + step).toFixed(2);
          lastKnown = valueToSet;
        } else if (patternType === 'date-sequence') {
          const nextDate = new Date(new Date(lastKnown).getTime() + step);
          valueToSet = nextDate.toISOString().split('T')[0];
          lastKnown = valueToSet;
        } else if (patternType === 'text-sequence') {
          const nextNum = textNumberStart + step * (rowOffset - seeds.length + 1);
          valueToSet = textPrefix + nextNum;
        }
  
        hot.setDataAtCell(row, col, valueToSet);
      }
    }
  
    hot.render();
  }
  
  export function addMockData(sheet) {
    const rows = 101, cols = 26;
    const data = [];
    const sampleAddresses = [
      '0x1234567890abcdef1234567890abcdef12345678',
      '0xabcdefabcdefabcdefabcdefabcdefabcdefabcd',
      '0xdeadbeefdeadbeefdeadbeefdeadbeefdeadbeef',
      '0xcafebabecafebabecafebabecafebabecafebabe',
      '0xfeedfeedfeedfeedfeedfeedfeedfeedfeedfeed'
    ];
  
    const header = Array.from({ length: cols }, (_, j) => String.fromCharCode(65 + (j % 26)));
    data.push(header);
  
    for (let i = 1; i < rows; i++) {
      const row = [];
      for (let j = 0; j < cols; j++) {
        if (j === 0) row.push("Row " + i);
        else if (j === 1) row.push((Math.random() * 1000).toFixed(2));
        else if (j === 2 && i % 10 === 0) row.push("ERROR: MISSING VALUE");
        else if (j === 3 && i % 15 === 0) row.push("KEYWORD_ROW_HIGHLIGHT");
        else if (j === 4) row.push(i % 4 === 0 ? sampleAddresses[i % sampleAddresses.length] : '');
        else if (j === 5) row.push(sampleAddresses[(i * 3) % sampleAddresses.length]);
        else row.push('');
      }
      data.push(row);
    }
  
    sheet.loadData(data);
  }
  
  export function clearOutput(sheet, rowCount = 100, colCount = 26) {
    const empty = Array.from({ length: rowCount }, () => Array(colCount).fill(''));
    sheet.loadData(empty);
  }
  // SpreadsheetUtils.js
  export function toggleCellStyle(sheetComponent, className) {

    if (!sheetComponent || !sheetComponent.hotInstance || !sheetComponent.lastSelection) {
      console.warn('[toggleCellStyle] Invalid sheet reference or selection');
      return;
    }

    
    const hot = sheetComponent.hotInstance;
    const selection = sheetComponent.lastSelection;
    if (!selection || selection.length === 0) return;
  
    const maxRows = hot.countRows();
    const maxCols = hot.countCols();
  
    selection.forEach(([startRow, startCol, endRow, endCol]) => {
      const rowStart = Math.max(0, Math.min(startRow, endRow));
      const rowEnd = Math.min(Math.max(startRow, endRow), maxRows - 1);
      const colStart = Math.max(0, Math.min(startCol, endCol));
      const colEnd = Math.min(Math.max(startCol, endCol), maxCols - 1);
  
      for (let row = rowStart; row <= rowEnd; row++) {
        for (let col = colStart; col <= colEnd; col++) {
          const key = `${row},${col}`;
          const existingClass = sheetComponent.cellClassMap.get(key) || '';
          const classList = new Set(existingClass.trim().split(/\s+/).filter(Boolean));
  
          if (classList.has(className)) {
            classList.delete(className);
          } else {
            classList.add(className);
          }
  
          const updatedClass = [...classList].join(' ');
          if (updatedClass.length > 0) {
            sheetComponent.cellClassMap.set(key, updatedClass);
            hot.setCellMeta(row, col, 'className', updatedClass);
          } else {
            sheetComponent.cellClassMap.delete(key);
            hot.removeCellMeta(row, col, 'className');
          }
        }
      }
    });
  
    hot.render();
    if (selection[0]) hot.selectCell(...selection[0]);
  }
  

export function injectColorStyle(hex) {
  const className = `color-${hex.slice(1)}`;
  const styleTagId = 'dynamic-color-styles';
  let styleTag = document.getElementById(styleTagId);

  if (!styleTag) {
    styleTag = document.createElement('style');
    styleTag.id = styleTagId;
    document.head.appendChild(styleTag);
  }

  // Only set background-color — allow user to explicitly choose text color
  if (!styleTag.textContent.includes(`.${className}`)) {
    styleTag.textContent += `
td.${className} {
  background-color: ${hex} !important;
}
`;
  }
}


export function injectTextColorStyle(hex) {
  const className = `text-${hex.slice(1)}`;
  const styleTagId = 'dynamic-color-styles';
  let styleTag = document.getElementById(styleTagId);

  if (!styleTag) {
    styleTag = document.createElement('style');
    styleTag.id = styleTagId;
    document.head.appendChild(styleTag);
  }

  if (!styleTag.textContent.includes(`.${className}`)) {
    styleTag.textContent += `\ntd.${className} {\n  color: ${hex} !important;\n}\n`;
  }
}


















export function applyBackgroundColor(sheet, selection, color, saveCallback) {
  const hot = sheet.hotInstance;

  
  if (!Array.isArray(selection) || selection.length === 0 || !Array.isArray(selection[0])) {
    console.warn('[setAlignment] No valid selection to align.');
    return;
  }

  selection.forEach(([startRow, startCol, endRow, endCol]) => {
    const rowStart = Math.max(0, Math.min(startRow, endRow));
    const rowEnd = Math.max(0, Math.max(startRow, endRow));
    const colStart = Math.max(0, Math.min(startCol, endCol));
    const colEnd = Math.max(0, Math.max(startCol, endCol));

    for (let row = rowStart; row <= rowEnd; row++) {
      for (let col = colStart; col <= colEnd; col++) {
        if (row < 0 || col < 0) continue;

        const key = `${row},${col}`;
        const className = `color-${color.slice(1)}`;
        const existing = sheet.cellClassMap.get(key) || '';
        const classList = new Set(existing.trim().split(/\s+/).filter(Boolean));
        
        // Remove any previous color-* class
        for (const cls of [...classList]) {
          if (cls.startsWith('color-')) {
            classList.delete(cls);
          }
        }
        
        classList.add(className);
        
        const finalClass = [...classList].join(' ');

        sheet.cellClassMap.set(key, finalClass);
        hot.setCellMeta(row, col, 'className', finalClass);
      }
    }
  });

  injectColorStyle(color);
  hot.render();
  if (saveCallback) saveCallback();
}

export function applyTextColor(sheet, selection, color, saveCallback) {
  const hot = sheet.hotInstance;

  
  if (!Array.isArray(selection) || selection.length === 0 || !Array.isArray(selection[0])) {
    console.warn('[setAlignment] No valid selection to align.');
    return;
  }

  selection.forEach(([startRow, startCol, endRow, endCol]) => {
    const rowStart = Math.max(0, Math.min(startRow, endRow));
    const rowEnd = Math.max(0, Math.max(startRow, endRow));
    const colStart = Math.max(0, Math.min(startCol, endCol));
    const colEnd = Math.max(0, Math.max(startCol, endCol));

    for (let row = rowStart; row <= rowEnd; row++) {
      for (let col = colStart; col <= colEnd; col++) {
        if (row < 0 || col < 0) continue;

        const key = `${row},${col}`;
        const className = `text-${color.slice(1)}`;
        const existing = sheet.cellClassMap.get(key) || '';
        const classList = new Set(existing.trim().split(/\s+/).filter(Boolean));
        for (const cls of [...classList]) {
          if (cls.startsWith('text-')) {
            classList.delete(cls);
          }
        }
        
        classList.add(className);
        const finalClass = [...classList].join(' ');

        sheet.cellClassMap.set(key, finalClass);
        hot.setCellMeta(row, col, 'className', finalClass);
      }
    }
  });

  injectTextColorStyle(color);
  hot.render();
  if (saveCallback) saveCallback();
}
export function setAlignment(sheet, selection, alignment, saveCallback = null) {
  const hot = sheet.hotInstance;


  if (!Array.isArray(selection) || selection.length === 0 || !Array.isArray(selection[0])) {
    console.warn('[setAlignment] No valid selection to align.');
    return;
  }


  selection.forEach(([startRow, startCol, endRow, endCol]) => {
    const rowStart = Math.max(0, Math.min(startRow, endRow));
    const rowEnd = Math.max(0, Math.max(startRow, endRow));
    const colStart = Math.max(0, Math.min(startCol, endCol));
    const colEnd = Math.max(0, Math.max(startCol, endCol));

    for (let row = rowStart; row <= rowEnd; row++) {
      for (let col = colStart; col <= colEnd; col++) {
        if (row < 0 || col < 0) continue;

        const key = `${row},${col}`;
        const className = `align-${alignment}`;
        const existing = sheet.cellClassMap.get(key) || '';
        const classList = new Set(existing.trim().split(/\s+/).filter(Boolean));
        ['align-left', 'align-center', 'align-right'].forEach(cls => classList.delete(cls));
        classList.add(className);
        const finalClass = [...classList].join(' ');
        sheet.cellClassMap.set(key, finalClass);
        hot.setCellMeta(row, col, 'className', finalClass);
      }
    }
  });

  injectAlignmentStyle();
  hot.render();
  if (saveCallback) saveCallback();
}

export function injectAlignmentStyle() {
  const styleTagId = 'dynamic-alignment-styles';
  let styleTag = document.getElementById(styleTagId);

  if (!styleTag) {
    styleTag = document.createElement('style');
    styleTag.id = styleTagId;
    document.head.appendChild(styleTag);
  }

  const styles = `
td.align-left { text-align: left !important; }
td.align-center { text-align: center !important; }
td.align-right { text-align: right !important; }
`;

  if (!styleTag.textContent.includes('align-left')) {
    styleTag.textContent += styles;
  }
}
export function applyBorderStyle(sheet, selection, borderType, saveCallback) {
  console.log('[applyBorderStyle] type:', borderType, 'selection:', selection);

  const hot = sheet.hotInstance;
  const plugin = hot.getPlugin('customBorders');

  const maxRows = hot.countRows();
  const maxCols = hot.countCols();
  const defaultBorder = { width: 1, color: 'black' };

  if (!Array.isArray(selection)) {
    console.warn('[applyBorderStyle] No valid selection provided');
    return;
  }

  // Ensure consistent shape
  const normalized = Array.isArray(selection[0]) ? selection : [selection];

  // ❗ Skip invalid selections like full row/column headers (-1 index)
  const filtered = normalized.filter(([r1, c1, r2, c2]) => {
    const valid = [r1, c1, r2, c2].every(v => Number.isInteger(v) && v >= 0);
    if (!valid) console.warn('Skipping invalid border selection:', [r1, c1, r2, c2]);
    return valid;
  });

  if (filtered.length === 0) return;

  // Clear borders first
  filtered.forEach(([r1, c1, r2, c2]) => {
    plugin.clearBorders([[r1, c1, r2, c2]]);
  });

  // Apply borders
filtered.forEach(([r1, c1, r2, c2]) => {
  const startRow = Math.min(r1, r2);
  const endRow = Math.max(r1, r2);
  const startCol = Math.min(c1, c2);
  const endCol = Math.max(c1, c2);

  for (let row = startRow; row <= endRow; row++) {
    if (row < 0 || row >= maxRows) continue;

    for (let col = startCol; col <= endCol; col++) {
      if (col < 0 || col >= maxCols) continue;

      const borders = {};

      if (borderType === 'full' || borderType === 'all') {
        borders.top = defaultBorder;
        borders.bottom = defaultBorder;
        borders.left = defaultBorder;
        borders.right = defaultBorder;

      } else if (borderType === 'outer') {
        if (row === startRow) borders.top = defaultBorder;
        if (row === endRow) borders.bottom = defaultBorder;
        if (col === startCol) borders.left = defaultBorder;
        if (col === endCol) borders.right = defaultBorder;

      } else if (borderType === 'inner') {
        if (col < endCol) borders.right = defaultBorder;
        if (row < endRow) borders.bottom = defaultBorder;

      } else {
        if (borderType === 'top' && row === startRow) borders.top = defaultBorder;
        if (borderType === 'bottom' && row === endRow) borders.bottom = defaultBorder;
        if (borderType === 'left' && col === startCol) borders.left = defaultBorder;
        if (borderType === 'right' && col === endCol) borders.right = defaultBorder;
      }

      if (Object.keys(borders).length > 0) {
        plugin.setBorders([[row, col]], borders);
      }
    }
  }
});


  hot.render();
  if (saveCallback) saveCallback();
}
