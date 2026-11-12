import { writeFile } from 'node:fs/promises';
import { PassThrough, Writable } from 'node:stream';
import { pipeline } from 'node:stream/promises';
import ExcelJS from 'exceljs';
import { addCustomProperties, documentProperties, serializeCustomProperties } from './document-properties.js';
import { Worksheet } from './worksheet.js';

export class Workbook {
  #workbook = new ExcelJS.Workbook();
  #worksheets = new Map();
  #properties = documentProperties(this.#workbook);
  #customProperties = new Map();

  get properties() {
    return this.#properties;
  }

  get customProperties() {
    return this.#customProperties;
  }

  addWorksheet(name) {
    const worksheet = new Worksheet(this.#workbook.addWorksheet(name));
    this.#worksheets.set(worksheet.name, worksheet);
    return worksheet;
  }

  getWorksheet(name) {
    return this.#worksheets.get(name);
  }

  async saveAs(target) {
    if (target instanceof Writable) {
      await this.#writeToStream(target);
    } else if (this.#customProperties.size > 0) {
      await writeFile(target, await this.toBuffer());
    } else {
      await this.#workbook.xlsx.writeFile(target);
    }
  }

  async toBuffer() {
    if (this.#customProperties.size === 0) {
      return this.#workbook.xlsx.writeBuffer();
    }

    const xml = serializeCustomProperties(this.#customProperties);
    const buffer = await this.#workbook.xlsx.writeBuffer();
    return addCustomProperties(buffer, xml);
  }

  async #writeToStream(target) {
    const output = new PassThrough();
    const delivery = pipeline(output, target);
    const write = this.#customProperties.size > 0
      ? this.toBuffer().then(buffer => { output.end(buffer); })
      : this.#workbook.xlsx.write(output);

    try {
      await Promise.all([delivery, write]);
    } catch (error) {
      output.destroy();
      throw error;
    }
  }
}
