export function applyStyle(worksheet, { row, column, bottom, right }, style) {
  if (!style || typeof style !== 'object' || Array.isArray(style)) {
    throw new TypeError('Style must be an object.');
  }

  const snapshot = structuredClone(style);
  for (let rowIndex = row; rowIndex <= bottom; rowIndex++) {
    for (let columnIndex = column; columnIndex <= right; columnIndex++) {
      worksheet.getCell(rowIndex, columnIndex).style = structuredClone(snapshot);
    }
  }
}
