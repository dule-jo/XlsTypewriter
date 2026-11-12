import assert from 'node:assert/strict';
import { Writable } from 'node:stream';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { Workbook } from 'xls-typewriter';

test('writes a worksheet with merged cells, styles and dimensions', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Inventory');

  sheet.merge(1, 2);
  sheet.print('Inventory');
  sheet.newRow();
  sheet.setColumnWidth(24);
  sheet.setRowHeight(20);
  sheet.setStyle({ font: { bold: true } }, 1, 2);
  sheet.print('Item');
  sheet.printAndGoDown('Quantity');
  sheet.print(42);

  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(await workbook.toBuffer());
  const inventory = reopened.getWorksheet('Inventory');

  assert.equal(inventory.getCell('A1').value, 'Inventory');
  assert.equal(inventory.getCell('B1').master.address, 'A1');
  assert.equal(inventory.getCell('A2').value, 'Item');
  assert.equal(inventory.getCell('B2').value, 'Quantity');
  assert.equal(inventory.getCell('B3').value, 42);
  assert.equal(inventory.getCell('A2').font.bold, true);
  assert.equal(inventory.getCell('B2').font.bold, true);
  assert.equal(inventory.getColumn(1).width, 24);
  assert.equal(inventory.getRow(2).height, 20);
});

test('preserves print settings and document properties when saving to a stream', async () => {
  const workbook = new Workbook();
  const sheet = workbook.addWorksheet('Inventory');
  sheet.pageSetup.orientation = 'landscape';
  sheet.headerFooter.oddFooter = '&RPage &P';
  workbook.properties.title = 'Quarterly inventory';
  workbook.customProperties.set('Department', 'Sales & Operations');

  const chunks = [];
  const stream = new Writable({
    write(chunk, encoding, callback) {
      chunks.push(chunk);
      callback();
    },
  });
  await workbook.saveAs(stream);
  assert.equal(stream.writableFinished, true);

  const buffer = Buffer.concat(chunks);
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(buffer);
  const inventory = reopened.getWorksheet('Inventory');
  assert.equal(reopened.title, 'Quarterly inventory');
  assert.equal(inventory.pageSetup.orientation, 'landscape');
  assert.equal(inventory.headerFooter.oddFooter, '&RPage &P');

  const archive = await JSZip.loadAsync(buffer);
  const custom = archive.file('docProps/custom.xml');
  assert.ok(custom);
  assert.match(await custom.async('string'), /name="Department"><vt:lpwstr>Sales &amp; Operations<\/vt:lpwstr>/);
  const relationships = await archive.file('_rels/.rels').async('string');
  assert.match(relationships, /relationships\/custom-properties" Target="docProps\/custom\.xml"/);
});
