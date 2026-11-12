import { Cursor, validatePosition, writeAtCursor } from './cursor.js';
import { fitColumns, fitRows, setColumnWidth, setRowHeight } from './layout.js';
import { applyStyle } from './styles.js';

export class Worksheet extends Cursor {
  #worksheet;

  constructor(worksheet) {
    super();
    this.#worksheet = worksheet;
  }

  get name() {
    return this.#worksheet.name;
  }

  get pageSetup() {
    return this.#worksheet.pageSetup;
  }

  get headerFooter() {
    return this.#worksheet.headerFooter;
  }

  print(value) {
    this.#writeValue(value, 'column');
  }

  printAndGoDown(value) {
    this.#writeValue(value, 'row');
  }

  merge(rows, columns) {
    const { row, column, bottom, right } = this.#range(rows, columns);
    if (rows === 1 && columns === 1) {
      return;
    }

    this.#worksheet.mergeCells(row, column, bottom, right);
  }

  getStyle() {
    const { row, column } = this.position;
    validatePosition(column, row);
    return structuredClone(this.#worksheet.getCell(row, column).style);
  }

  setStyle(style, rows = 1, columns = 1) {
    applyStyle(this.#worksheet, this.#range(rows, columns), style);
  }

  setColumnWidth(width) {
    const { column } = this.position;
    validatePosition(column, 1);
    setColumnWidth(this.#worksheet, column, width);
  }

  setRowHeight(height) {
    const { row } = this.position;
    validatePosition(1, row);
    setRowHeight(this.#worksheet, row, height);
  }

  autoFitCurrentColumn() {
    const { column } = this.position;
    validatePosition(column, 1);
    fitColumns(this.#worksheet, column);
  }

  autoFitCurrentRow() {
    const { row } = this.position;
    validatePosition(1, row);
    fitRows(this.#worksheet, row);
  }

  autoFitColumns() {
    fitColumns(this.#worksheet);
  }

  autoFitRows() {
    fitRows(this.#worksheet);
  }

  #range(rows, columns) {
    if (!Number.isSafeInteger(rows) || rows < 1 || !Number.isSafeInteger(columns) || columns < 1) {
      throw new RangeError('Rows and columns must be positive integers.');
    }

    const { row, column } = this.position;
    const bottom = row + rows - 1;
    const right = column + columns - 1;
    validatePosition(right, bottom);

    return { row, column, bottom, right };
  }

  #writeValue(value, axis) {
    this[writeAtCursor](axis, (row, column) => {
      this.#worksheet.getCell(row, column).value = value;
    });
  }

  goToEnd() {
    const { right, bottom } = this.#worksheet.dimensions;
    this.goTo(right, bottom);
  }
}
