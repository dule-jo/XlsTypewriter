const maxColumnWidth = 255;
const maxRowHeight = 409;
const defaultColumnWidth = 8.43;
const defaultFontSize = 11;
const columnPadding = 2;

function validateSize(size, max, name) {
  if (!Number.isFinite(size) || size <= 0 || size > max) {
    throw new RangeError(`${name} must be greater than 0 and at most ${max}.`);
  }
}

export function setColumnWidth(worksheet, column, width) {
  validateSize(width, maxColumnWidth, 'Column width');
  worksheet.getColumn(column).width = width;
}

export function setRowHeight(worksheet, row, height) {
  validateSize(height, maxRowHeight, 'Row height');
  worksheet.getRow(row).height = height;
}

function measureCell(cell) {
  if (!cell || cell.isMerged) {
    return undefined;
  }

  const value = cell.result ?? cell.value;
  const text = value instanceof Date
    ? value.toISOString().slice(0, 10)
    : String(cell.result ?? cell.text ?? '');
  if (text.length === 0) {
    return undefined;
  }

  const fontSize = cell.font?.size ?? defaultFontSize;
  const widthScale = (fontSize / defaultFontSize) * (cell.font?.bold ? 1.1 : 1);
  const lineWidths = text.split(/\r\n|\r|\n/).map(line => [...line].length * widthScale);
  return { fontSize, lineWidths };
}

export function fitColumns(worksheet, column) {
  const widths = new Map();
  function measureColumnCell(cell, columnIndex) {
    const measure = measureCell(cell);
    if (!measure) {
      return;
    }
    let width = widths.get(columnIndex) ?? 0;
    for (const lineWidth of measure.lineWidths) {
      width = Math.max(width, lineWidth + columnPadding);
    }
    widths.set(columnIndex, width);
  }

  worksheet.eachRow(row => {
    if (column === undefined) {
      row.eachCell(measureColumnCell);
    } else {
      measureColumnCell(row.findCell(column), column);
    }
  });

  for (const [columnIndex, width] of widths) {
    worksheet.getColumn(columnIndex).width = Math.min(maxColumnWidth, Math.ceil(width));
  }
}

export function fitRows(worksheet, rowNumber) {
  function fitRow(row) {
    let height = 0;
    row.eachCell((cell, columnIndex) => {
      const measure = measureCell(cell);
      if (!measure) {
        return;
      }

      let lineCount = measure.lineWidths.length;
      if (cell.alignment?.wrapText) {
        const columnWidth = worksheet.getColumn(columnIndex).width
          ?? worksheet.properties.defaultColWidth
          ?? defaultColumnWidth;
        const availableWidth = Math.max(1, columnWidth - columnPadding);
        lineCount = measure.lineWidths.reduce(
          (count, width) => count + Math.max(1, Math.ceil(width / availableWidth)),
          0,
        );
      }
      height = Math.max(height, lineCount * measure.fontSize * 1.35);
    });

    if (height > 0) {
      row.height = Math.min(maxRowHeight, Math.max(worksheet.properties.defaultRowHeight, Math.ceil(height)));
    }
  }

  if (rowNumber === undefined) {
    worksheet.eachRow(fitRow);
  } else {
    const row = worksheet.findRow(rowNumber);
    if (row) {
      fitRow(row);
    }
  }
}
