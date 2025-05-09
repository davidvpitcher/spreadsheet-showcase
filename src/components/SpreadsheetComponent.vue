<template>
  <div>
    <div class="spreadsheet-container" ref="spreadsheetContainer"></div>
    <button @click="addRows">{{ buttonLabel }}</button>
  </div>
</template>

<script>
import Handsontable from 'handsontable';
import 'handsontable/dist/handsontable.full.css';



export default {
  name: 'SpreadsheetComponent',
  props: {
  
  buttonLabel: { type: String, required: true },
  width: { type: String, default: '800px' },
  height: { type: String, default: '400px' },
  highlightHeader: { type: Boolean, default: false }, 
  freezeTopRow: { type: Boolean, default: false },
  autoSaveEnabled: { type: Boolean, default: false },


  },
data() { // our data
  return {
    hotInstance: null,
    cellClassMap: new Map() // key = `${row},${col}`, value = className
  };
}
,


  mounted() {

     const vm = this;

    console.log('SpreadsheetComponent mounted');


    this.hotInstance = new Handsontable(this.$refs.spreadsheetContainer, {


 afterSelectionEnd: (row, col, row2, col2) => {
    this.lastSelection = [[row, col, row2, col2]];
    this.$emit('update-last-selection', this.lastSelection); // ✅ EMIT TO PARENT
  },
 
 afterColumnMove: () => { 
    console.log("AFTER COLUMN MOVE 1");
  if (this.autoSaveEnabled) {
    this.$emit('request-save');
  }
  },

 afterRowMove: () => { 
    console.log("AFTER Row MOVE 1");
  if (this.autoSaveEnabled) {
    this.$emit('request-save');
  }
  },

  
      contextMenu: {
    items: {
      ...Handsontable.plugins.ContextMenu.DEFAULT_ITEMS,  // include all default options
      'separator_fill': Handsontable.plugins.ContextMenu.SEPARATOR,
      'fill_down_custom': {
        name: '🠗 Fill Down Selection',
        callback: () => {
          const selected = this.hotInstance.getSelected();

          if (!selected || selected.length === 0) {
            alert("No selection.");
            return;
          }

          const [startRow, startCol, endRow] = selected[selected.length - 1];
  //const [startRow, startCol, endRow, endCol] = selected[selected.length - 1];

/*
          if (startCol !== endCol) {
            alert("Only one column can be selected.");
            return;
          }
*/
          const topValue = this.hotInstance.getDataAtCell(startRow, startCol);
          if (topValue === null || topValue === '') {
            alert("Top cell is empty.");
            return;
          }

          const isNumeric = !isNaN(topValue);
          for (let row = startRow + 1; row <= endRow; row++) {
            let valueToSet = topValue;
            if (isNumeric) {
              valueToSet = parseFloat(topValue) + (row - startRow);
            }
            this.hotInstance.setDataAtCell(row, startCol, valueToSet);
          }

          this.hotInstance.render();
        }
      },
'recolor_like_cells': {
          name: '🎨 Recolor Like Cells',
          callback: function () {
            const hot = vm.hotInstance;
            const selected = hot.getSelected();
            if (!selected || selected.length === 0) {
              alert("Please select a range to recolor.");
              return;
            }



         const [rawRow1, rawCol1, rawRow2, rawCol2] = selected[selected.length - 1];

const startRow = Math.min(rawRow1, rawRow2);
const endRow = Math.max(rawRow1, rawRow2);
const startCol = Math.min(rawCol1, rawCol2);
const endCol = Math.max(rawCol1, rawCol2);

            const cellMap = new Map();

            for (let row = startRow; row <= endRow; row++) {
              for (let col = startCol; col <= endCol; col++) {
                const val = hot.getDataAtCell(row, col)?.toString().trim();
                if (!val) continue;
                if (!cellMap.has(val)) cellMap.set(val, []);
                cellMap.get(val).push([row, col]);
              }
            }

            const usedColors = new Set();
            const getRandomColor = () => {
              let color;
              do {
                color = '#' + Math.floor(Math.random() * 0xffffff).toString(16).padStart(6, '0');
              } while (usedColors.has(color));
              usedColors.add(color);
              return color;
            };

            const isDarkColor = (hex) => {
              const rgb = parseInt(hex.slice(1), 16);
              const r = (rgb >> 16) & 0xff;
              const g = (rgb >> 8) & 0xff;
              const b = (rgb >> 0) & 0xff;
              const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
              return luminance < 128;
            };

            const cellClasses = {};

            cellMap.forEach((positions) => {
              if (positions.length < 2) return;
              const bgColor = getRandomColor();
              const className = `color-${bgColor.slice(1)}`;
              const textColor = isDarkColor(bgColor) ? '#ffffff' : '#000000';
              cellClasses[className] = { backgroundColor: bgColor, color: textColor };

              positions.forEach(([row, col]) => {
                const key = `${row},${col}`;
                vm.cellClassMap.set(key, className); // ✅ correct binding
              });
            });

            // Inject styles
            const styleTagId = 'dynamic-color-styles';
            let styleTag = document.getElementById(styleTagId);
            if (!styleTag) {
              styleTag = document.createElement('style');
              styleTag.id = styleTagId;
              document.head.appendChild(styleTag);
            }

            let css = '';
            for (const [className, styles] of Object.entries(cellClasses)) {
       css += `
  td.${className} {
    background-color: ${styles.backgroundColor} !important;
    color: ${styles.color} !important;
  }
`;

            }

            styleTag.textContent += css;
            console.log('Selected range:', selected[selected.length - 1]);
console.log('Unique values found:', [...cellMap.keys()]);
console.log('cellMap content:', cellMap);


            hot.render();

if (vm.autoSaveEnabled) {
  vm.$emit('request-save');
}



          }
        }
      }
    },



      data: Array(100).fill().map(() => Array(26).fill('')), // Start with 100 empty rows
      rowHeaders: true,
      colHeaders: true,
     // contextMenu: true,
      manualRowResize: true,
      manualColumnResize: true,
      minRows: 100,
      minCols: 26,
      colWidths: 100, // Set a default column width
      width: this.width,
      height: this.height,
      manualColumnMove:  true, // Apply saved order or enable move

manualRowMove: true,

  customBorders: [], 
      licenseKey: 'non-commercial-and-evaluation',
       fixedRowsTop: this.freezeTopRow ? 1 : 0,

cells: function (row, col) {
  const key = `${row},${col}`;
  const vueComponent = this.instance?.rootElement.__vue__;
  const customClassString = vueComponent?.cellClassMap?.get(key) || '';
  const customClasses = new Set(customClassString.split(/\s+/).filter(Boolean));

  const defaultMeta = Handsontable.plugins.MetaManager
    ? this.instance.getCellMeta(row, col)
    : {};

  const defaultClassString = defaultMeta?.className || '';
  const defaultClasses = new Set(
    typeof defaultClassString === 'string'
      ? defaultClassString.split(/\s+/).filter(Boolean)
      : []
  );

  // Merge default Handsontable class names (e.g., htSearchResult) with ours
  const allClasses = new Set([...customClasses, ...defaultClasses]);

  if (vueComponent?.highlightHeader && row === 0) {
    allClasses.add('header-highlight');
  }

  return {
    className: [...allClasses].join(' ')
  };
}





    });


 // ✅ New focusout-based logic
  const spreadsheetEl = this.$refs.spreadsheetContainer;
  spreadsheetEl.addEventListener('focusout', () => {
    setTimeout(() => {
      const newlyFocused = document.activeElement;
      const isStillInSpreadsheet = spreadsheetEl.contains(newlyFocused);
      const isInToolbar = newlyFocused?.closest('.toolbar');

     // console.log('[focusout] spreadsheet blurred to', newlyFocused);

      if (!isStillInSpreadsheet && !isInToolbar) {
   //     console.log('[focusout] Triggering deselected');
      this.$emit('deselected');

      }
    }, 0);
  });


/*
  this.$refs.spreadsheetContainer.addEventListener('focusout', () => {
    console.log('[focusout] spreadsheet blurred to', document.activeElement);
  });
*/

    this.$refs.spreadsheetContainer.addEventListener('paste', this.onPaste);
  },
  beforeUnmount() {
    console.log('SpreadsheetComponent beforeUnmount');
    if (this.hotInstance) {
      this.hotInstance.destroy();
    }
  },
  methods: {
    clearRecolors() {
  this.cellClassMap.clear();
  this.hotInstance.render();
}
,
 restoreFormatting() {
  const hot = this.hotInstance;

  hot.updateSettings({
    cells: (row, col) => {
      const key = `${row},${col}`;
      const classNames = this.cellClassMap.get(key);
      if (classNames) return { className: classNames };
      return {};
    }
  });

  // Create or get the style tag for dynamic classes
  const styleTagId = 'dynamic-color-styles';
  let styleTag = document.getElementById(styleTagId);
  if (!styleTag) {
    styleTag = document.createElement('style');
    styleTag.id = styleTagId;
    document.head.appendChild(styleTag);
  }

  let css = '';
  const seen = new Set();

  for (const classList of this.cellClassMap.values()) {
    for (const className of classList.split(' ')) {
      if (seen.has(className)) continue;
      seen.add(className);

    if (className.startsWith('color-')) {
  const hex = '#' + className.slice(6);
  css += `
td.${className} {
  background-color: ${hex} !important;
}
`;
}  else if (className.startsWith('text-')) {
     const hex = '#' + className.slice(5);
  css += `
td.${className} {
  color: ${hex} !important;
}
`;
      } else if (className === 'bold-cell') {
        css += `td.bold-cell { font-weight: bold !important; }\n`;
      } else if (className === 'italic-cell') {
        css += `td.italic-cell { font-style: italic !important; }\n`;
      } else if (className === 'underline-cell') {
        css += `td.underline-cell { text-decoration: underline !important; }\n`;
      } else if (className === 'strike-cell') {
        css += `td.strike-cell { text-decoration: line-through !important; }\n`;
      } else if (className === 'align-left') {
        css += `td.align-left { text-align: left !important; }\n`;
      } else if (className === 'align-center') {
        css += `td.align-center { text-align: center !important; }\n`;
      } else if (className === 'align-right') {
        css += `td.align-right { text-align: right !important; }\n`;
      }else if (/^font-\d+$/.test(className)) {
  const size = className.split('-')[1];
  css += `td.${className} { font-size: ${size}px !important; }\n`;
}else if (className === 'valign-top') {
  css += `td.valign-top { vertical-align: top !important; }\n`;
} else if (className === 'valign-middle') {
  css += `td.valign-middle { vertical-align: middle !important; }\n`;
} else if (className === 'valign-bottom') {
  css += `td.valign-bottom { vertical-align: bottom !important; }\n`;
}


    }
  }

  styleTag.textContent = css;
  hot.render();
}

,
resetColumnOrder() {
  localStorage.removeItem('spreadsheet_column_order');
  window.location.reload();
}
,
isDarkColor(hex) {
  const rgb = parseInt(hex.slice(1), 16);
  const r = (rgb >> 16) & 0xff;
  const g = (rgb >> 8) & 0xff;
  const b = (rgb >> 0) & 0xff;
  const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
  return luminance < 128;
}
,
onPaste(event) {
  event.preventDefault(); // ❗️ Prevent Handsontable's default paste

  const clipboardData = event.clipboardData || window.clipboardData;
  const pastedData = clipboardData.getData('Text');
  console.log('Raw pasted data:', JSON.stringify(pastedData));

  this.insertParsedPastedData(pastedData);
},
insertParsedPastedData(pastedData) {
  const trimmedData = pastedData.trim();
  const rows = this.customSplit(trimmedData);

  const parsedData = rows.map(row =>
    row.split('\t').map(cell => cell.trim())
  );

  const selection = this.hotInstance.getSelected();
  if (!selection || selection.length === 0) {
    alert("Please select a starting cell before pasting.");
    return;
  }

  const [startRow, startCol] = selection[0];

  // Paste into cells from starting point
  for (let i = 0; i < parsedData.length; i++) {
    for (let j = 0; j < parsedData[i].length; j++) {
      this.hotInstance.setDataAtCell(startRow + i, startCol + j, parsedData[i][j]);
    }
  }

  this.hotInstance.render();
},
  parsePastedData(pastedData) {
    const trimmedData = pastedData.trim();
    const rows = this.customSplit(trimmedData);

    console.log('Rows split from pasted data:', rows);

    const parsedData = rows.map(row => {
      const cells = row.split('\t');
      console.log('Cells split from row:', cells);
      return cells.map(cell => cell.trim());
    });

    console.log('Filtered parsed data before loading:', parsedData);
    this.hotInstance.loadData(parsedData);
  },
  customSplit(input) {
    let rows = [];
    let currentRow = '';
    let inQuotes = false;

    for (let i = 0; i < input.length; i++) {
      const currentChar = input[i];
      const nextChar = i < input.length - 1 ? input[i + 1] : '';

      currentRow += currentChar;

      if (currentChar === '"' && nextChar !== '"') { // Toggle the inQuotes state, ignore double quotes ""
        inQuotes = !inQuotes;
      } else if (currentChar === '\n' && !inQuotes) { // End of row
        rows.push(currentRow.trim());
        currentRow = '';
      }
    }

    if (currentRow.length > 0) {
      rows.push(currentRow.trim()); // Add the last row
    }

    return rows.filter(row => row !== ''); // Filter out any empty rows
  },
    addRows() {
      const currentData = this.hotInstance.getData();
      const newData = currentData.concat(Array(100).fill().map(() => Array(26).fill('')));
      this.hotInstance.loadData(newData);
    },
    getData() {
      return this.hotInstance.getData();
    },
    loadData(data) {
      this.hotInstance.loadData(data);
    }
  }
};
</script>

<style>
.spreadsheet-container {
  margin: 20px;
  border: 1px solid #ccc; /* Add border to see the container */
  overflow: hidden; /* Ensure overflow is handled */
}
button {
  display: block;
  margin: 20px auto;
  padding: 10px 20px;
  font-size: 16px;
}



</style>
