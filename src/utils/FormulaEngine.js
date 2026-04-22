import { HyperFormula } from 'hyperformula';

/**
 * Thin wrapper around HyperFormula to keep the dependency swappable.
 * Each grid tab creates one FormulaEngine instance with one sheet.
 */
export function createFormulaEngine() {
  const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  const sheetName = hf.addSheet('Sheet');
  const sheetId = hf.getSheetId(sheetName);

  return {
    hf,
    sheetId,

    setCellValue(row, col, rawValue) {
      hf.setCellContents({ sheet: sheetId, row, col }, [[rawValue]]);
    },

    getCellValue(row, col) {
      return hf.getCellValue({ sheet: sheetId, row, col });
    },

    getDisplayValue(row, col) {
      const val = hf.getCellValue({ sheet: sheetId, row, col });
      if (val instanceof Object && val.type === 'ERROR') {
        return String(val.value ?? '#ERROR!');
      }
      return val;
    },

    getCellFormula(row, col) {
      if (hf.doesCellHaveFormula({ sheet: sheetId, row, col })) {
        return hf.getCellFormula({ sheet: sheetId, row, col });
      }
      return null;
    },

    /**
     * Hydrate the engine from a cells map.
     * cells: { 'A1': { v: 10, f: '=B1+1' }, ... }
     */
    hydrate(cells, rows, cols) {
      const data = [];
      for (let r = 0; r < rows; r++) {
        const row = [];
        for (let c = 0; c < cols; c++) {
          const key = colLabel(c) + String(r + 1);
          const cell = cells[key];
          if (cell && cell.f) {
            row.push(cell.f);
          } else if (cell && cell.v != null) {
            row.push(cell.v);
          } else {
            row.push(null);
          }
        }
        data.push(row);
      }
      hf.setSheetContent(sheetId, data);
    },

    /**
     * After a cell edit, return a map of all cells whose computed value changed.
     * Returns: { 'A1': computedValue, 'B2': computedValue, ... }
     */
    getChangedCells() {
      const changes = {};
      const allChanges = hf.getExportedChanges?.() || [];
      for (const change of allChanges) {
        if (change.sheet === sheetId) {
          const key = colLabel(change.col) + String(change.row + 1);
          changes[key] = change.newValue;
        }
      }
      return changes;
    },

    destroy() {
      hf.destroy();
    },
  };
}

function colLabel(index) {
  let label = '';
  let n = index;
  while (n >= 0) {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  }
  return label;
}
