import { STRICT_FONT_FAMILY, STRICT_FONT_SIZE_PT } from './types';

export interface VirtualGridCell {
  cell: HTMLTableCellElement;
  originRow: number;
  originCol: number;
  rowSpan: number;
  colSpan: number;
  isOrigin: boolean;
}

export interface VirtualTableGrid {
  table: HTMLTableElement;
  rows: HTMLTableRowElement[];
  numRows: number;
  numCols: number;
  matrix: (VirtualGridCell | null)[][];
}

/**
 * Builds a complete 2D virtual grid matrix for an HTMLTableElement,
 * accurately resolving all rowSpan and colSpan placements.
 */
export function buildVirtualTableGrid(table: HTMLTableElement): VirtualTableGrid {
  const rows = Array.from(table.rows);
  const numRows = rows.length;
  const matrix: (VirtualGridCell | null)[][] = Array.from({ length: numRows }, () => []);
  let maxCols = 0;

  for (let r = 0; r < numRows; r++) {
    const tr = rows[r];
    let c = 0;
    const cells = Array.from(tr.cells);

    for (const cell of cells) {
      while (matrix[r][c] !== undefined && matrix[r][c] !== null) {
        c++;
      }
      const rowSpan = Math.max(1, cell.rowSpan || 1);
      const colSpan = Math.max(1, cell.colSpan || 1);

      for (let dr = 0; dr < rowSpan; dr++) {
        for (let dc = 0; dc < colSpan; dc++) {
          const targetRow = r + dr;
          const targetCol = c + dc;
          if (targetRow < numRows) {
            while (matrix[targetRow].length <= targetCol) {
              matrix[targetRow].push(null);
            }
            matrix[targetRow][targetCol] = {
              cell,
              originRow: r,
              originCol: c,
              rowSpan,
              colSpan,
              isOrigin: dr === 0 && dc === 0,
            };
          }
        }
      }
      c += colSpan;
      if (c > maxCols) maxCols = c;
    }
  }

  for (let r = 0; r < numRows; r++) {
    while (matrix[r].length < maxCols) {
      matrix[r].push(null);
    }
  }

  return {
    table,
    rows,
    numRows,
    numCols: maxCols,
    matrix,
  };
}

export function createDefaultCell(doc: Document = document): HTMLTableCellElement {
  const td = doc.createElement('td');
  td.style.border = '1px solid #000000';
  td.style.padding = '2px 4px';
  td.style.verticalAlign = 'top';
  td.style.fontFamily = STRICT_FONT_FAMILY;
  td.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
  td.style.lineHeight = '1';

  const p = doc.createElement('p');
  p.setAttribute('dir', 'rtl');
  p.style.margin = '0';
  p.style.lineHeight = '1';
  p.style.fontFamily = STRICT_FONT_FAMILY;
  p.style.fontSize = `${STRICT_FONT_SIZE_PT}pt`;
  p.appendChild(doc.createElement('br'));
  td.appendChild(p);
  return td;
}

export function createNotaryTableHTML(rowsCount: number, colsCount: number): string {
  const colWidthPct = (100 / Math.max(1, colsCount)).toFixed(2);
  let html = `<table dir="rtl" style="width:100%;max-width:120mm;border-collapse:collapse;table-layout:fixed;margin:0;font-family:Arial;font-size:13pt;line-height:1;"><tbody>`;
  for (let r = 0; r < rowsCount; r++) {
    html += `<tr>`;
    for (let c = 0; c < colsCount; c++) {
      html += `<td style="width:${colWidthPct}%;border:1px solid #000000;padding:2px 4px;vertical-align:top;font-family:Arial;font-size:13pt;line-height:1;"><p dir="rtl" style="margin:0;line-height:1;font-family:Arial;font-size:13pt;"><br></p></td>`;
    }
    html += `</tr>`;
  }
  html += `</tbody></table><p dir="rtl" style="margin:0;line-height:1;font-family:Arial;font-size:13pt;"><br></p>`;
  return html;
}

/**
 * Inserts a row above or below the row containing activeCell,
 * expanding rowSpan of any cells spanning across the insertion boundary.
 */
export function insertTableRowAtCell(
  activeCell: HTMLTableCellElement,
  position: 'above' | 'below'
): HTMLTableCellElement | null {
  const table = activeCell.closest('table') as HTMLTableElement | null;
  if (!table) return null;
  const grid = buildVirtualTableGrid(table);

  let targetRowIdx = -1;
  for (let r = 0; r < grid.numRows; r++) {
    for (let c = 0; c < grid.numCols; c++) {
      const slot = grid.matrix[r][c];
      if (slot && slot.cell === activeCell && slot.isOrigin) {
        targetRowIdx = position === 'above' ? r : r + slot.rowSpan - 1;
        break;
      }
    }
    if (targetRowIdx !== -1) break;
  }
  if (targetRowIdx === -1) return null;

  const newTr = document.createElement('tr');
  const handledCells = new Set<HTMLTableCellElement>();
  let firstCreatedCell: HTMLTableCellElement | null = null;

  for (let c = 0; c < grid.numCols; c++) {
    const slot = grid.matrix[targetRowIdx][c];
    if (!slot) {
      const td = createDefaultCell();
      if (!firstCreatedCell) firstCreatedCell = td;
      newTr.appendChild(td);
      continue;
    }

    // Check if this cell spans strictly across the insertion boundary
    const spansAcrossAbove =
      position === 'above' && slot.originRow < targetRowIdx && slot.originRow + slot.rowSpan > targetRowIdx;
    const spansAcrossBelow =
      position === 'below' && slot.originRow <= targetRowIdx && slot.originRow + slot.rowSpan > targetRowIdx + 1;

    if (spansAcrossAbove || spansAcrossBelow) {
      if (!handledCells.has(slot.cell)) {
        slot.cell.rowSpan = slot.rowSpan + 1;
        handledCells.add(slot.cell);
      }
    } else {
      const td = createDefaultCell();
      if (slot.cell.style.width) td.style.width = slot.cell.style.width;
      if (!firstCreatedCell) firstCreatedCell = td;
      newTr.appendChild(td);
    }
  }

  const refRow = grid.rows[targetRowIdx];
  if (position === 'above') {
    refRow.parentNode?.insertBefore(newTr, refRow);
  } else {
    refRow.parentNode?.insertBefore(newTr, refRow.nextSibling);
  }

  return firstCreatedCell;
}

/**
 * Deletes the row containing activeCell, adjusting rowSpans and transferring
 * origin cells to the next row if their top row was deleted.
 */
export function deleteTableRowAtCell(activeCell: HTMLTableCellElement): void {
  const table = activeCell.closest('table') as HTMLTableElement | null;
  if (!table) return;
  const grid = buildVirtualTableGrid(table);
  if (grid.numRows <= 1) {
    table.remove();
    return;
  }

  let rowIdx = -1;
  for (let r = 0; r < grid.numRows; r++) {
    for (let c = 0; c < grid.numCols; c++) {
      if (grid.matrix[r][c]?.cell === activeCell) {
        rowIdx = r;
        break;
      }
    }
    if (rowIdx !== -1) break;
  }
  if (rowIdx === -1) return;

  const handledCells = new Set<HTMLTableCellElement>();

  for (let c = 0; c < grid.numCols; c++) {
    const slot = grid.matrix[rowIdx][c];
    if (!slot || handledCells.has(slot.cell)) continue;
    handledCells.add(slot.cell);

    if (slot.rowSpan > 1) {
      if (slot.originRow === rowIdx && rowIdx + 1 < grid.numRows) {
        // Move this origin cell into the next row at the right DOM index
        const nextRow = grid.rows[rowIdx + 1];
        slot.cell.rowSpan = slot.rowSpan - 1;
        let insertBeforeCell: HTMLTableCellElement | null = null;
        for (let nc = c + slot.colSpan; nc < grid.numCols; nc++) {
          const nextSlot = grid.matrix[rowIdx + 1][nc];
          if (nextSlot && nextSlot.originRow === rowIdx + 1) {
            insertBeforeCell = nextSlot.cell;
            break;
          }
        }
        nextRow.insertBefore(slot.cell, insertBeforeCell);
      } else {
        slot.cell.rowSpan = slot.rowSpan - 1;
      }
    }
  }

  grid.rows[rowIdx].remove();
}

/**
 * Inserts a column to the right (before in RTL) or left (after in RTL) of activeCell.
 */
export function insertTableColumnAtCell(
  activeCell: HTMLTableCellElement,
  direction: 'before' | 'after'
): void {
  const table = activeCell.closest('table') as HTMLTableElement | null;
  if (!table) return;
  const grid = buildVirtualTableGrid(table);

  let targetColIdx = -1;
  for (let r = 0; r < grid.numRows; r++) {
    for (let c = 0; c < grid.numCols; c++) {
      const slot = grid.matrix[r][c];
      if (slot && slot.cell === activeCell && slot.isOrigin) {
        targetColIdx = direction === 'before' ? c : c + slot.colSpan - 1;
        break;
      }
    }
    if (targetColIdx !== -1) break;
  }
  if (targetColIdx === -1) return;

  const handledCells = new Set<HTMLTableCellElement>();

  for (let r = 0; r < grid.numRows; r++) {
    const slot = grid.matrix[r][targetColIdx];
    if (!slot) continue;

    const spansAcrossBefore =
      direction === 'before' && slot.originCol < targetColIdx && slot.originCol + slot.colSpan > targetColIdx;
    const spansAcrossAfter =
      direction === 'after' && slot.originCol <= targetColIdx && slot.originCol + slot.colSpan > targetColIdx + 1;

    if (spansAcrossBefore || spansAcrossAfter) {
      if (!handledCells.has(slot.cell)) {
        slot.cell.colSpan = slot.colSpan + 1;
        handledCells.add(slot.cell);
      }
    } else if (slot.originRow === r) {
      const newTd = createDefaultCell();
      const tr = grid.rows[r];
      if (direction === 'before') {
        tr.insertBefore(newTd, slot.cell);
      } else {
        tr.insertBefore(newTd, slot.cell.nextSibling);
      }
    }
  }

  // Rebalance column widths evenly
  rebalanceTableColumnWidths(table);
}

/**
 * Deletes the column containing activeCell, adjusting colSpans accordingly.
 */
export function deleteTableColumnAtCell(activeCell: HTMLTableCellElement): void {
  const table = activeCell.closest('table') as HTMLTableElement | null;
  if (!table) return;
  const grid = buildVirtualTableGrid(table);
  if (grid.numCols <= 1) {
    table.remove();
    return;
  }

  let colIdx = -1;
  for (let r = 0; r < grid.numRows; r++) {
    for (let c = 0; c < grid.numCols; c++) {
      if (grid.matrix[r][c]?.cell === activeCell) {
        colIdx = c;
        break;
      }
    }
    if (colIdx !== -1) break;
  }
  if (colIdx === -1) return;

  const handledCells = new Set<HTMLTableCellElement>();

  for (let r = 0; r < grid.numRows; r++) {
    const slot = grid.matrix[r][colIdx];
    if (!slot || handledCells.has(slot.cell)) continue;
    handledCells.add(slot.cell);

    if (slot.colSpan > 1) {
      slot.cell.colSpan = slot.colSpan - 1;
    } else {
      slot.cell.remove();
    }
  }

  rebalanceTableColumnWidths(table);
}

export function rebalanceTableColumnWidths(table: HTMLTableElement): void {
  const grid = buildVirtualTableGrid(table);
  if (grid.numCols <= 0) return;
  const unitPct = 100 / grid.numCols;
  const seen = new Set<HTMLTableCellElement>();

  for (let r = 0; r < grid.numRows; r++) {
    for (let c = 0; c < grid.numCols; c++) {
      const slot = grid.matrix[r][c];
      if (slot && !seen.has(slot.cell)) {
        seen.add(slot.cell);
        slot.cell.style.width = `${(unitPct * slot.colSpan).toFixed(2)}%`;
      }
    }
  }
}

/**
 * Merges a set of selected cells (or activeCell with its right/below neighbor)
 * into a single rectangular cell spanning minRow..maxRow and minCol..maxCol.
 */
export function mergeSelectedTableCells(cells: HTMLTableCellElement[]): HTMLTableCellElement | null {
  if (cells.length === 0) return null;
  const table = cells[0].closest('table') as HTMLTableElement | null;
  if (!table) return null;
  const grid = buildVirtualTableGrid(table);

  // If only 1 cell is passed, try merging it with its next sibling cell in the same row
  let targetCells = [...cells];
  if (targetCells.length === 1) {
    const single = targetCells[0];
    const nextSibling = single.nextElementSibling as HTMLTableCellElement | null;
    if (nextSibling) {
      targetCells.push(nextSibling);
    } else {
      return single;
    }
  }

  const cellSet = new Set(targetCells);
  let minR = Infinity;
  let maxR = -1;
  let minC = Infinity;
  let maxC = -1;

  // Expand bounding rectangle until stable (in case any intersecting cell spans outside)
  let changed = true;
  while (changed) {
    changed = false;
    for (let r = 0; r < grid.numRows; r++) {
      for (let c = 0; c < grid.numCols; c++) {
        const slot = grid.matrix[r][c];
        if (!slot) continue;
        const inBounds = r >= minR && r <= maxR && c >= minC && c <= maxC;
        if (cellSet.has(slot.cell) || inBounds) {
          if (!cellSet.has(slot.cell)) {
            cellSet.add(slot.cell);
            changed = true;
          }
          if (slot.originRow < minR) {
            minR = slot.originRow;
            changed = true;
          }
          if (slot.originRow + slot.rowSpan - 1 > maxR) {
            maxR = slot.originRow + slot.rowSpan - 1;
            changed = true;
          }
          if (slot.originCol < minC) {
            minC = slot.originCol;
            changed = true;
          }
          if (slot.originCol + slot.colSpan - 1 > maxC) {
            maxC = slot.originCol + slot.colSpan - 1;
            changed = true;
          }
        }
      }
    }
  }

  if (minR === Infinity || maxR < minR || maxC < minC) return null;

  const anchorSlot = grid.matrix[minR][minC];
  if (!anchorSlot) return null;
  const anchorCell = anchorSlot.cell;

  // Preserve non-empty content from merged cells into anchorCell
  const mergedCells = Array.from(cellSet);
  for (const c of mergedCells) {
    if (c === anchorCell) continue;
    const text = (c.textContent || '').trim();
    if (text.length > 0) {
      Array.from(c.childNodes).forEach((node) => {
        anchorCell.appendChild(node.cloneNode(true));
      });
    }
    c.remove();
  }

  anchorCell.rowSpan = maxR - minR + 1;
  anchorCell.colSpan = maxC - minC + 1;
  rebalanceTableColumnWidths(table);
  return anchorCell;
}

/**
 * Splits a merged cell (rowSpan > 1 or colSpan > 1) back into individual 1x1 cells.
 */
export function splitMergedTableCell(cell: HTMLTableCellElement): void {
  const table = cell.closest('table') as HTMLTableElement | null;
  if (!table) return;
  const rSpan = cell.rowSpan || 1;
  const cSpan = cell.colSpan || 1;
  if (rSpan <= 1 && cSpan <= 1) return;

  const grid = buildVirtualTableGrid(table);
  let originR = -1;
  let originC = -1;

  for (let r = 0; r < grid.numRows; r++) {
    for (let c = 0; c < grid.numCols; c++) {
      const slot = grid.matrix[r][c];
      if (slot && slot.cell === cell && slot.isOrigin) {
        originR = r;
        originC = c;
        break;
      }
    }
    if (originR !== -1) break;
  }

  if (originR === -1) return;

  cell.rowSpan = 1;
  cell.colSpan = 1;

  for (let dr = 0; dr < rSpan; dr++) {
    const targetRow = grid.rows[originR + dr];
    if (!targetRow) continue;

    // Find insertion anchor in targetRow
    let insertBeforeEl: HTMLTableCellElement | null = null;
    for (let nc = originC + cSpan; nc < grid.numCols; nc++) {
      const nextSlot = grid.matrix[originR + dr][nc];
      if (nextSlot && nextSlot.originRow === originR + dr) {
        insertBeforeEl = nextSlot.cell;
        break;
      }
    }

    const startDc = dr === 0 ? 1 : 0;
    for (let dc = startDc; dc < cSpan; dc++) {
      const newTd = createDefaultCell();
      targetRow.insertBefore(newTd, insertBeforeEl);
    }
  }

  rebalanceTableColumnWidths(table);
}
