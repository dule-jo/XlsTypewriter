const maxColumn = 16_384;
const maxRow = 1_048_576;

function validateCount(count) {
  if (!Number.isSafeInteger(count) || count < 0) {
    throw new RangeError('Count must be a non-negative integer.');
  }
}

export function validatePosition(column, row) {
  if (!Number.isInteger(column) || column < 1 || column > maxColumn) {
    throw new RangeError(`Column must be an integer between 1 and ${maxColumn}.`);
  }
  if (!Number.isInteger(row) || row < 1 || row > maxRow) {
    throw new RangeError(`Row must be an integer between 1 and ${maxRow}.`);
  }
}

export const writeAtCursor = Symbol('writeAtCursor');

export class Cursor {
  #position = { column: 1, row: 1 };
  #positions = [];

  get position() {
    return { ...this.#position };
  }

  goTo(column, row) {
    validatePosition(column, row);
    this.#position = { column, row };
  }

  [writeAtCursor](axis, write) {
    const position = this.#position;
    const { column, row } = position;
    validatePosition(column, row);
    write(row, column);

    // Writing the final cell can leave the cursor one step past the sheet edge.
    this.#position = { ...position, [axis]: position[axis] + 1 };
  }

  goToStart() {
    this.goTo(1, 1);
  }

  skip(count = 1) {
    validateCount(count);
    this.goTo(this.#position.column + count, this.#position.row);
  }

  newRow(count = 1) {
    validateCount(count);
    this.goTo(1, this.#position.row + count);
  }

  newRowKeepColumn(count = 1) {
    validateCount(count);
    this.goTo(this.#position.column, this.#position.row + count);
  }

  pushPosition() {
    this.#positions.push(this.#position);
  }

  popPosition() {
    if (this.#positions.length === 0) {
      throw new RangeError('No saved cursor position.');
    }
    this.#position = this.#positions.pop();
  }

  purgePositions() {
    this.#positions.length = 0;
  }
}
