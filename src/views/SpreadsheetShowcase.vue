<template>
  <div class="spreadsheet-tool">
    <h1>Spreadsheet Showcase</h1>





    <div class="section">
      <h2>Input Spreadsheet</h2>


<div class="showcase-toolbar">
  <!-- Existing Buttons -->
  <button @click="addMockData" title="Generate mock data">🧪 Mock</button>
  <button @click="highlightColumn2" title="Highlight errors in Column 2">⚠️</button>
<button
  class="striped-btn"
  @click="highlightAlternatingRows"
  title="Zebra stripe rows"
>
  &nbsp;&nbsp;&nbsp;&nbsp;
</button>


  <button @click="highlightKeywordRows" title="Highlight rows with keyword">🔍 Keywords</button>
  <button @click="resetFormatting" title="Reset all formatting">♻️</button>
<button
  @click="toggleFreezeTopRow"
  :title="freezeTop ? 'Unfreeze top row' : 'Freeze top row'"
  :class="{ 'active-freeze': freezeTop }"
>
  {{ freezeTop ? '🧊 Top' : '⬆️ Top' }}
</button>

<button
  @click="toggleFreezeFirstColumn"
  :title="freezeLeft ? 'Unfreeze left column' : 'Freeze left column'"
  :class="{ 'active-freeze': freezeLeft }"
>
  {{ freezeLeft ? '🧊 Left' : '⬅️ Left' }}
</button>


  <button @click="analyzeColumnTypes" title="Analyze column types">📊</button>

  <!-- Search Controls -->
<input
  v-model="searchQuery"
  placeholder="Search..."
  @keyup.enter="applySearch"
  style="padding: 6px; font-size: 13px; width: 160px;"
/>
<button @click="applySearch">🔎</button>
<button @click="clearSearch">❌</button>

<span v-if="matchCount > 0" style="font-size: 12px; margin-left: 4px;">
  ({{ matchCount }} matches)
</span>

<label style="font-size: 12px; display: flex; align-items: center; gap: 4px;">
  <input type="checkbox" v-model="autoSaveEnabled" @change="onAutoSaveToggle" />
  <span style="font-size: 20px; line-height: 1;">💾</span>

</label>


  
</div>



      <div class="toolbar">
<button
  @mousedown="pendingToolbarAction = true"
  @click="toggleBold"
  title="Bold"
><strong>B</strong></button>
<button
  @mousedown="pendingToolbarAction = true"
  @click="toggleItalic"
  title="Italic"
  aria-label="Italic"
>
  <em>I</em>
</button>
<button
  @mousedown="pendingToolbarAction = true"
  @click="toggleUnderline"
  title="Underline"
  aria-label="Underline"
><u>U</u></button>
<button
  @mousedown="pendingToolbarAction = true"
  @click="toggleStrikeThrough"
  title="Strike-through"
>
  <s>S</s>
</button>

<button
  @mousedown="pendingToolbarAction = true"
  @click="openColorPicker"
  title="Color Cell Background"
  ref="colorButton"
>
  🎨
</button>

<input
  ref="colorInput"
  type="color"
  class="color-input-overlay"
  @change="applyBackgroundColor"
/>
<button
  @mousedown="pendingToolbarAction = true"
  @click="openTextColorPicker"
  title="Text Color"
  ref="textColorButton"
>
  🖍️
</button>

<input
  ref="textColorInput"
  type="color"
  class="color-input-overlay"
  @change="applyTextColor"
/>


<button @click="setAlignment('left')" title="Align Left">⬅️</button>
<button @click="setAlignment('center')" title="Align Center">↔️</button>
<button @click="setAlignment('right')" title="Align Right">➡️</button>

<button
  @mousedown="pendingToolbarAction = true"
  @click="toggleWrap"
  title="Toggle Word Wrap"
>
  🔁
</button>


<button @click="changeFontSize(-1)" title="Decrease Font Size">➖</button>
<input
  type="number"
  min="8"
  max="72"
  v-model.number="currentFontSize"
  @change="applyFontSize"
  title="Set Font Size"
/>
<button @click="changeFontSize(1)" title="Increase Font Size">➕</button>

<button @click="setVertical('top')" title="Align Top">🔼</button>
<button @click="setVertical('middle')" title="Align Middle">⏺</button>
<button @click="setVertical('bottom')" title="Align Bottom">🔽</button>

<button @click="showBorderModal = true" title="Apply Border">🔲</button>





  <!-- Add more icons/functions here later -->
</div>


  <!--
<div class="formula-bar">
  <label>=</label>
  <input
    v-model="formulaInput"
    @keyup.enter="applyFormula"
    placeholder="Enter formula (e.g., =A1+B1 or =SUM(A1:A5))"
  />
  <span v-if="formulaResult !== null">→ {{ formulaResult }}</span>
</div>-->












<SpreadsheetComponent
  ref="inputSpreadsheet"
  button-label="Add 100 Rows"
  :highlight-header="true"
  :freeze-top-row="true"
  :auto-save-enabled="autoSaveEnabled"
  @request-save="saveSpreadsheetFromEvent"
  @update-last-selection="lastSelection = $event"  
    @deselected="handleDeselection"
/>
<BorderModal
  v-if="showBorderModal"
  @close="showBorderModal = false"
  @apply="applyBorderStyle"
/>




    </div>






<button class="process-btn" @click="processData" aria-label="Process input data and populate output">Process Data</button>

    <div class="section">
      <h2>Output Spreadsheet</h2>
      <SpreadsheetComponent 
      ref="outputSpreadsheet" 
      button-label="Add 100 Rows" 
      
  :highlight-header="false"
      />
    </div>

    <button class="export-btn" @click="exportToCSV">Export Output as CSV</button>
    <button class="clear-btn" @click="clearOutput">Clear Output</button>


  </div>
</template>

<script>
import SpreadsheetComponent from '../components/SpreadsheetComponent.vue';
import BorderModal from '../components/BorderModal.vue';

/* ideas

formatting palette
subsheets
formulas

*/

import { // function imports
  analyzeColumnTypes,
  enableSearchPlugin,
  applySearch,
  clearSearch,
  toggleFreezeTopRow,
  toggleFreezeFirstColumn,
  resetFormatting,
  fillDownSelection,
  highlightKeywordRows,
  highlightColumn2,
  highlightAlternatingRows,
  addMockData,
  clearRecolors,
  exportToCSV,
  clearOutput,
  processData,
  saveSpreadsheet,
  restoreSavedSpreadsheet,
  toggleCellStyle,
  applyTextColor,
  applyBackgroundColor,
setAlignment ,
toggleWrapText,
setFontSize ,
setVerticalAlignment,
applyBorderStyle
} from '../utils/SpreadsheetUtils';


export default {
  name: 'Tool60Page',
components: {
  SpreadsheetComponent,
  BorderModal
},

  mounted() {


  console.log('Tool60Page mounted — current handleDeselection:', this.handleDeselection);


  // this.addMockData(); // disabled to work on save/load feature
    enableSearchPlugin(this.$refs.inputSpreadsheet.hotInstance);

   if (this.autoSaveEnabled) {
     restoreSavedSpreadsheet(this.$refs.inputSpreadsheet);
    this.$nextTick(() => {
      const hot = this.$refs.inputSpreadsheet.hotInstance;
   this.saveHooks.afterChange = () => saveSpreadsheet(this.$refs.inputSpreadsheet);
this.saveHooks.afterColumnResize = () => saveSpreadsheet(this.$refs.inputSpreadsheet);
this.saveHooks.afterRowResize = () => saveSpreadsheet(this.$refs.inputSpreadsheet);

hot.addHook('afterChange', this.saveHooks.afterChange);
hot.addHook('afterColumnResize', this.saveHooks.afterColumnResize);
hot.addHook('afterRowResize', this.saveHooks.afterRowResize);

    });
  }
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && !e.altKey) {
    const key = e.key.toLowerCase();
    if (key === 'b') {
      e.preventDefault();
      this.toggleBold();
    } else if (key === 'i') {
      e.preventDefault();
      this.toggleItalic();
    } else if (key === 'u') {
      e.preventDefault();
      this.toggleUnderline();
    } else if (key === 'x' && e.shiftKey) {
      e.preventDefault();
      this.toggleStrikeThrough();
    }
  }
});



  },
  beforeUnmount() {
  document.removeEventListener('keydown', this.keyHandler);
},

data() { // our data
  return {
    searchQuery: '',
    matchCount: 0,
    autoSaveEnabled: JSON.parse(localStorage.getItem('tool60_autosave_enabled') || 'false'),
    lastSelection: null,
    pendingToolbarAction: false,
currentFontSize: 14,
showBorderModal: false,
freezeTop: false,
freezeLeft: false,
   // formulaInput: '',
  //  formulaResult: null,

deselectionTimeout: null,
     saveHooks: {
      afterChange: null,
      afterColumnResize: null,
      afterRowResize: null
    }
  };
}
,
  methods: {
    restoreLastSelection() {
  const sheet = this.$refs.inputSpreadsheet;
  const hot = sheet.hotInstance;
  const selection = sheet.lastSelection;

  if (selection) {
    hot.selectCells(selection);
  }
}
,
setAlignment(alignment) {
  setAlignment(this.$refs.inputSpreadsheet, this.lastSelection, alignment, this.autoSaveEnabled ? this.saveSpreadsheetFromEvent : null);
},
toggleBold() {
  if (this.deselectionTimeout) clearTimeout(this.deselectionTimeout);
  toggleCellStyle(this.$refs.inputSpreadsheet, 'bold-cell');
  if (this.autoSaveEnabled) this.saveSpreadsheetFromEvent();
},
toggleItalic() {
  if (this.deselectionTimeout) clearTimeout(this.deselectionTimeout);
  toggleCellStyle(this.$refs.inputSpreadsheet, 'italic-cell');
  if (this.autoSaveEnabled) this.saveSpreadsheetFromEvent();
},

toggleUnderline() {
  if (this.deselectionTimeout) clearTimeout(this.deselectionTimeout);
  toggleCellStyle(this.$refs.inputSpreadsheet, 'underline-cell');
  if (this.autoSaveEnabled) this.saveSpreadsheetFromEvent();
}
,



setVertical(alignment) {
  setVerticalAlignment(
    this.$refs.inputSpreadsheet,
    this.lastSelection,
    alignment,
    this.autoSaveEnabled ? this.saveSpreadsheetFromEvent : null
  );
},
applyBorderStyle(type) {
  const sheet = this.$refs.inputSpreadsheet;
  const selection = this.lastSelection;

  if (!selection) return;

  applyBorderStyle(
    sheet,
    selection,
    type,
    this.autoSaveEnabled ? () => this.saveSpreadsheetFromEvent() : null
  );

  this.showBorderModal = false;
}

,
changeFontSize(delta) {
  this.currentFontSize = Math.max(8, Math.min(72, this.currentFontSize + delta));
  this.applyFontSize();
},
applyFontSize() {
  setFontSize(
    this.$refs.inputSpreadsheet,
    this.lastSelection,
    this.currentFontSize,
    this.autoSaveEnabled ? () => this.saveSpreadsheetFromEvent() : null
  );
},

  onAutoSaveToggle() {
    const sheet = this.$refs.inputSpreadsheet;
    const hot = sheet.hotInstance;

    localStorage.setItem('tool60_autosave_enabled', JSON.stringify(this.autoSaveEnabled));

    if (this.autoSaveEnabled) {
      this.saveHooks.afterChange = () => saveSpreadsheet(sheet);
      this.saveHooks.afterColumnResize = () => saveSpreadsheet(sheet);
      this.saveHooks.afterRowResize = () => saveSpreadsheet(sheet);

      hot.addHook('afterChange', this.saveHooks.afterChange);
      hot.addHook('afterColumnResize', this.saveHooks.afterColumnResize);
      hot.addHook('afterRowResize', this.saveHooks.afterRowResize);

      this.saveHooks.afterChange();
    } else {
      hot.removeHook('afterChange', this.saveHooks.afterChange);
      hot.removeHook('afterColumnResize', this.saveHooks.afterColumnResize);
      hot.removeHook('afterRowResize', this.saveHooks.afterRowResize);
    }
  },
    
openColorPicker() {
  const button = this.$refs.colorButton;
  const input = this.$refs.colorInput;
// was   this.$refs.colorInput.click();
  // Position the color input directly over the button
  const rect = button.getBoundingClientRect();
  input.style.left = `${rect.left}px`;
  input.style.top = `${rect.top}px`;

  input.click();
}
,

    applyBackgroundColor(event) {
      applyBackgroundColor(this.$refs.inputSpreadsheet, this.lastSelection, event.target.value);
      if (this.autoSaveEnabled) this.saveSpreadsheetFromEvent();
    },


toggleStrikeThrough() {
  if (this.deselectionTimeout) clearTimeout(this.deselectionTimeout);
  toggleCellStyle(this.$refs.inputSpreadsheet, 'strike-cell');
  if (this.autoSaveEnabled) this.saveSpreadsheetFromEvent();
}
,
openTextColorPicker() {
  const button = this.$refs.textColorButton;
  const input = this.$refs.textColorInput;

  const rect = button.getBoundingClientRect();
  input.style.left = `${rect.left}px`;
  input.style.top = `${rect.top}px`;

  input.click();

  // was   this.$refs.textColorInput.click();
},


 applyTextColor(event) {
      applyTextColor(this.$refs.inputSpreadsheet, this.lastSelection, event.target.value);
      if (this.autoSaveEnabled) this.saveSpreadsheetFromEvent();
    },

 processData() {
      processData(this.$refs.inputSpreadsheet, this.$refs.outputSpreadsheet);
    },
  clearRecolors() {
      clearRecolors(this.$refs.inputSpreadsheet, this.autoSaveEnabled ? () => saveSpreadsheet(this.$refs.inputSpreadsheet) : null);
    },
saveSpreadsheetFromEvent() {
  saveSpreadsheet(this.$refs.inputSpreadsheet);
}
,
 analyzeColumnTypes() {
    const hot = this.$refs.inputSpreadsheet.hotInstance;
    const types = analyzeColumnTypes(hot);
    alert("Column types:\n" + types.map((t, i) => `Col ${i}: ${t}`).join('\n'));
  },

fillDownSelection() {
      fillDownSelection(this.$refs.inputSpreadsheet.hotInstance);
    },



applySearch() {
  const hot = this.$refs.inputSpreadsheet.hotInstance;
  const success = applySearch(hot, this.searchQuery);

  // Count how many matches are in the table
  const highlights = hot.getCellsMeta().filter(meta => meta.className?.includes('htSearchResult'));
  this.matchCount = highlights.length;

  if (!success) {
    alert('No results found.');
  }
}

,
clearSearch() {
  const hot = this.$refs.inputSpreadsheet.hotInstance;
  clearSearch(hot);
  this.searchQuery = '';
  this.matchCount = 0;
}
,
toggleFreezeTopRow() {
  const hot = this.$refs.inputSpreadsheet.hotInstance;
  this.freezeTop = !this.freezeTop;
  toggleFreezeTopRow(hot, this.freezeTop);
},

toggleFreezeFirstColumn() {
  const hot = this.$refs.inputSpreadsheet.hotInstance;
  this.freezeLeft = !this.freezeLeft;
  toggleFreezeFirstColumn(hot, this.freezeLeft);
},

resetFormatting() {
  
 resetFormatting(this.$refs.inputSpreadsheet, this.autoSaveEnabled ? () => saveSpreadsheet(this.$refs.inputSpreadsheet) : null);

 this.clearRecolors();
}


,
handleDeselection() {
  console.log('[Tool60Page] Deselect event received — clearing selection');
  this.lastSelection = null;
}



,

toggleWrap() {
  toggleWrapText(
    this.$refs.inputSpreadsheet,
    this.lastSelection,
    this.autoSaveEnabled ? () => this.saveSpreadsheetFromEvent() : null
  );
}
,
 
 highlightKeywordRows() {
      highlightKeywordRows(this.$refs.inputSpreadsheet, this.autoSaveEnabled ? () => saveSpreadsheet(this.$refs.inputSpreadsheet) : null);

    },


    highlightAlternatingRows() {
     highlightAlternatingRows(this.$refs.inputSpreadsheet, this.autoSaveEnabled ? () => saveSpreadsheet(this.$refs.inputSpreadsheet) : null);

    }

, addMockData() {
      addMockData(this.$refs.inputSpreadsheet);
    },

highlightColumn2() {
      highlightColumn2(this.$refs.inputSpreadsheet, this.autoSaveEnabled ? () => saveSpreadsheet(this.$refs.inputSpreadsheet) : null);
    },
/*
 applyFormula() {
    const input = this.formulaInput.trim();
    if (!input.startsWith('=')) {
      this.formulaResult = 'Invalid';
      return;
    }

    const formula = input.slice(1);
    try {
      const result = this.evaluateFormula(formula);
      this.formulaResult = isNaN(result) ? 'Error' : result;
    } catch (e) {
      this.formulaResult = 'Error';
    }
  },

  evaluateFormula(formula) {
    const sheet = this.$refs.inputSpreadsheet.hotInstance;

    // Convert A1-style references to actual cell values
    formula = formula.replace(/[A-Z]+\d+/g, (ref) => {
      const col = ref.charCodeAt(0) - 65;
      const row = parseInt(ref.slice(1), 10) - 1;
      const val = sheet.getDataAtCell(row, col);
      return parseFloat(val || 0);
    });

    // Handle SUM(A1:A5)
    formula = formula.replace(/SUM\((.*?)\)/gi, (_, range) => {
      const [start, end] = range.split(':');
      const col = start.charCodeAt(0) - 65;
      const rowStart = parseInt(start.slice(1), 10) - 1;
      const rowEnd = parseInt(end.slice(1), 10) - 1;
      let sum = 0;
      for (let r = rowStart; r <= rowEnd; r++) {
        const val = parseFloat(sheet.getDataAtCell(r, col) || 0);
        sum += val;
      }
      return sum;
    });

    // Evaluate final expression
    return Function(`"use strict"; return (${formula})`)();
  },*/
 clearOutput() {
      clearOutput(this.$refs.outputSpreadsheet);
    },
    exportToCSV() {
      exportToCSV(this.$refs.outputSpreadsheet);
    }
  }
};
</script>

<style scoped>
.spreadsheet-tool {
  max-width: 1100px;
  margin: 0 auto;
  padding: 2rem 1rem;
}

h1 {
  text-align: center;
  margin-bottom: 2rem;
}

.section {
  margin-bottom: 2rem;
}

h2 {
  text-align: left;
  margin-bottom: 0.5rem;
  color: #2c3e50;
}



.search-bar,
.persistence-bar {
  display: flex;
  gap: 8px;
  justify-content: center;
  align-items: center;
  margin-bottom: 1rem;
}


.search-bar input {
  padding: 8px;
  font-size: 16px;
  width: 250px;
}

.process-btn,
.export-btn,
.clear-btn,
.highlight-btn,
button {
  display: inline-block;
  margin: 0.5rem 0.5rem;
  padding: 10px 16px;
  font-size: 14px;
  background-color: #3498db;
  border: none;
  color: white;
  cursor: pointer;
  border-radius: 4px;
}
button:hover {
  background-color: #2980b9;
}
.generate-btn {
  display: block;
  margin: 1.5rem auto;
  padding: 12px 24px;
  font-size: 16px;
  background-color: #2ecc71;
  border: none;
  color: white;
  cursor: pointer;
  border-radius: 4px;
}
.generate-btn:hover {
  background-color: #27ae60;
}

.process-btn:hover,
.export-btn:hover {
  background-color: #2980b9;
}



.error-cell {
  background-color: #ffcccc !important;
  color: #900;
  font-weight: bold;
}
.toolbar {
  display: flex;
  flex-wrap: wrap;
  gap: 2px; /* reduce from default gap */
  margin-bottom: 4px;
  align-items: center;
}

.toolbar button,
.toolbar input[type="number"] {
  padding: 3px 5px;
  font-size: 14px;
  line-height: 1.2;
  border: 1px solid #ccc;
  border-radius: 4px;
  background: #f9f9f9;
  cursor: pointer;
}
.toolbar button[title="Align Middle"] {
  color: #444; /* darker for visibility */
  font-weight: bold; /* optional: gives ⏺ more presence */
}

.toolbar button:hover {
  background: #e6f0ff;
}

.toolbar input[type="number"] {
  width: 50px;
  text-align: center;
}
.toolbar button strong {
  color: black; /* fixes white-on-gray issue */
}

.toolbar button:hover {
  background-color: #ddd;
}
.toolbar button em {
  font-style: italic;
  color: black;
}
.toolbar button u {
  text-decoration: underline;
  color: black;
}

.color-picker {
  width: 36px;
  height: 36px;
  padding: 0;
  border: none;
  cursor: pointer;
}
td[class*="color-"] {
  font-weight: normal;
}
.toolbar button s {
  text-decoration: line-through;
  color: black;
}
.showcase-toolbar {
  display: flex;
  flex-wrap: wrap;
  gap: 2px;
  margin-bottom: 5px;
  justify-content: center;
  align-items: center;
}

.showcase-toolbar button {
  padding: 2px 5px;
  font-size: 13px;
  background-color: #eee;

  border: 1px solid #ccc;
  border-radius: 4px;
  cursor: pointer;
  transition: background-color 0.2s, color 0.2s;
  color: #000000; /* 🧠 explicitly set readable font color */
}

.showcase-toolbar button:hover {
  background-color: #cce6ff;
  color: #000000; /* optional but reinforces visibility */
}

.showcase-toolbar button.active-freeze {
  background-color: #aaa;
  color: white;
  border-color: #000;
}
.showcase-toolbar .striped-btn {
  background-image: repeating-linear-gradient(
    to bottom,
    #cccccc,
    #cccccc 4px,
    #eeeeee 4px,
    #eeeeee 8px
  );
  color: black;
  font-weight: bold;
}

.showcase-toolbar .striped-btn:hover {
  background-image: repeating-linear-gradient(
    to bottom,
    #bbbbbb,
    #bbbbbb 4px,
    #dddddd 4px,
    #dddddd 8px
  );
}
.color-input-overlay {
  position: absolute;
  opacity: 0;
  width: 36px;
  height: 36px;
  cursor: pointer;
  z-index: 10;
}
/*
.formula-bar {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 10px;
  justify-content: center;
}
.formula-bar input {
  width: 300px;
  padding: 6px;
  font-size: 14px;
}
*/

</style>
