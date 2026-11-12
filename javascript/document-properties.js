import JSZip from 'jszip';

const standardFields = [
  'title', 'subject', 'creator', 'keywords', 'description', 'category',
  'company', 'manager', 'language', 'revision', 'contentStatus',
  'created', 'modified', 'lastModifiedBy', 'lastPrinted',
];

export function documentProperties(workbook) {
  const properties = {};
  for (const field of standardFields) {
    Object.defineProperty(properties, field, {
      enumerable: true,
      get: () => workbook[field],
      set: value => { workbook[field] = value; },
    });
  }
  return properties;
}

const officeNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006';
const propertyFormat = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}';
const propertyPath = 'docProps/custom.xml';
const xmlEscapes = {
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;',
  '\t': '&#9;', '\n': '&#10;', '\r': '&#13;',
};

function escapeXml(text) {
  if (/[^\t\n\r\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u.test(text)) {
    throw new TypeError('Custom properties must contain valid XML characters.');
  }
  return text.replace(/[&<>"\t\n\r]/g, character => xmlEscapes[character]);
}

function variant(value) {
  if (typeof value === 'string') {
    return ['lpwstr', escapeXml(value)];
  }
  if (typeof value === 'boolean') {
    return ['bool', String(value)];
  }
  if (Number.isFinite(value)) {
    const isInt32 = Number.isInteger(value) && value >= -2_147_483_648 && value <= 2_147_483_647;
    return [isInt32 ? 'i4' : 'r8', String(value)];
  }
  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return ['filetime', value.toISOString()];
  }
  throw new TypeError('Custom property values must be strings, booleans, finite numbers, or valid Dates.');
}

export function serializeCustomProperties(properties) {
  const elements = [];
  for (const [name, value] of properties) {
    if (typeof name !== 'string' || !name.trim()) {
      throw new TypeError('Custom property names must be non-empty strings.');
    }
    const [type, text] = variant(value);
    const id = elements.length + 2;
    elements.push(
      `<property fmtid="${propertyFormat}" pid="${id}" name="${escapeXml(name)}">`
      + `<vt:${type}>${text}</vt:${type}></property>`,
    );
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + `<Properties xmlns="${officeNamespace}/custom-properties" xmlns:vt="${officeNamespace}/docPropsVTypes">`
    + elements.join('') + '</Properties>';
}

async function appendXml(archive, path, closingTag, element) {
  const xml = await archive.file(path)?.async('string');
  if (!xml?.includes(closingTag)) {
    throw new Error(`Invalid XLSX package entry: ${path}`);
  }
  archive.file(path, xml.replace(closingTag, element + closingTag));
}

export async function addCustomProperties(buffer, xml) {
  const archive = await JSZip.loadAsync(buffer);
  archive.file(propertyPath, xml);
  await Promise.all([
    appendXml(archive, '[Content_Types].xml', '</Types>',
      `<Override PartName="/${propertyPath}" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>`),
    appendXml(archive, '_rels/.rels', '</Relationships>',
      `<Relationship Id="rIdCustomProperties" Type="${officeNamespace}/relationships/custom-properties" Target="${propertyPath}"/>`),
  ]);
  return archive.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
