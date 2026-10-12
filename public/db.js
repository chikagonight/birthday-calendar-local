(function () {
"use strict";

const XLSX = (typeof require === 'function') ? require('./vendor/xlsx.full.min.js') : globalThis.XLSX;
const JSZip = (typeof require === 'function') ? require('./vendor/jszip.min.js') : globalThis.JSZip;

const COLUMNS = {
  employees: ['id', 'last_name', 'first_name', 'middle_name', 'birth_date', 'position', 'facility_id', 'note'],
  facilities: ['id', 'name', 'city', 'color'],
  gifts: ['id', 'employee_id', 'year', 'gift', 'gift_cost'],
};
const OPTIONAL = { employees: ['middle_name', 'position', 'note'], facilities: ['color'], gifts: ['gift', 'gift_cost'] };

function dateISO(value, date1904 = false) {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? '' : value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const parts = XLSX.SSF.parse_date_code(value, { date1904 });
    return parts ? `${String(parts.y).padStart(4, '0')}-${String(parts.m).padStart(2, '0')}-${String(parts.d).padStart(2, '0')}` : '';
  }
  return String(value ?? '');
}

function parseWorkbook(buffer) {
  const wb = XLSX.read(buffer, { type: (typeof Buffer !== 'undefined' && Buffer.isBuffer(buffer)) ? 'buffer' : 'array', cellDates: false });
  const result = {};
  const date1904 = Boolean(wb.Workbook?.WBProps?.date1904);
  for (const [name, columns] of Object.entries(COLUMNS)) {
    if (!wb.Sheets[name]) throw new Error(`Missing worksheet: ${name}`);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: true });
    const headers = rows[0] || [];
    const indices = columns.map(column => headers.indexOf(column));
    if (indices.includes(-1)) throw new Error(`Missing columns in ${name}`);
    result[name] = rows.slice(1).filter(row => indices.some(i => row[i] !== '' && row[i] != null)).map(row => {
      const record = {};
      columns.forEach((column, i) => {
        const value = row[indices[i]] ?? '';
        record[column] = column === 'birth_date' ? dateISO(value, date1904)
          : (column === 'year' || column === 'gift_cost') && value !== '' ? Number(value) : value;
      });
      return record;
    });
  }
  return result;
}

const present = value => typeof value === 'string' && value.trim().length > 0;
function validateFacility(row) {
  const errors = [];
  for (const field of ['name', 'city']) if (!present(row[field])) errors.push(`${field} is required`);
  if (row.color != null && row.color !== '' && !/^#[0-9a-f]{6}$/i.test(row.color)) errors.push('color must be a hex RGB value');
  return errors;
}
function validateEmployee(row, existingEmployees, facilities) {
  const errors = [];
  for (const field of ['last_name', 'first_name', 'facility_id']) if (!present(row[field])) errors.push(`${field} is required`);
  const date = row.birth_date;
  if (typeof date !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(date) ||
      Number.isNaN(Date.parse(`${date}T00:00:00Z`)) || new Date(`${date}T00:00:00Z`).toISOString().slice(0, 10) !== date) errors.push('birth_date must be YYYY-MM-DD');
  if (present(row.facility_id) && !facilities.some(f => f.id === row.facility_id)) errors.push('facility_id does not exist');
  return errors;
}
function validateGift(row, existingGifts, employees) {
  const errors = [];
  if (!present(row.employee_id) || !employees.some(e => e.id === row.employee_id)) errors.push('employee_id does not exist');
  if (row.year === '' || row.year == null || !Number.isInteger(Number(row.year))) errors.push('year must be an integer');
  if (row.gift_cost != null && row.gift_cost !== '' && (String(row.gift_cost).trim() === '' || !Number.isFinite(Number(row.gift_cost)) || Number(row.gift_cost) < 0)) errors.push('gift_cost must be nonnegative');
  if (existingGifts.some(g => g.id !== row.id && g.employee_id === row.employee_id && Number(g.year) === Number(row.year))) errors.push('year already exists for employee_id');
  return errors;
}

function checked(errors) { if (errors.length) throw new Error(errors.join('; ')); }
function create(db, data, name, validator) {
  const row = { ...data, id: crypto.randomUUID() };
  checked(validator(row, db));
  const next = { ...db, [name]: [...db[name], row] };
  return { db: next, [name.slice(0, -1).replace(/ie$/, 'y')]: row };
}
function update(db, id, patch, name, validator) {
  const old = db[name].find(row => row.id === id);
  if (!old) throw new Error(`${name} id not found`);
  if (Object.hasOwn(patch, 'id') && patch.id !== id) throw new Error('id cannot be changed');
  const row = { ...old, ...patch, id };
  checked(validator(row, db));
  return { db: { ...db, [name]: db[name].map(item => item.id === id ? row : item) },
    [name.slice(0, -1).replace(/ie$/, 'y')]: row };
}
function remove(db, id, name) {
  if (!db[name].some(row => row.id === id)) throw new Error(`${name} id not found`);
  return { ...db, [name]: db[name].filter(row => row.id !== id) };
}
const createFacility = (db, data) => create(db, data, 'facilities', (row) => validateFacility(row));
const createEmployee = (db, data) => create(db, data, 'employees', (row, state) => validateEmployee(row, state.employees, state.facilities));
const createGift = (db, data) => create(db, data, 'gifts', (row, state) => validateGift(row, state.gifts, state.employees));
const updateFacility = (db, id, patch) => update(db, id, patch, 'facilities', row => validateFacility(row));
const updateEmployee = (db, id, patch) => update(db, id, patch, 'employees', (row, state) => validateEmployee(row, state.employees, state.facilities));
const updateGift = (db, id, patch) => update(db, id, patch, 'gifts', (row, state) => validateGift(row, state.gifts, state.employees));
function deleteFacility(db, id) {
  if (db.employees.some(row => row.facility_id === id)) throw new Error('facility has employees');
  return remove(db, id, 'facilities');
}
function deleteEmployee(db, id, { cascade = false } = {}) {
  if (!cascade && db.gifts.some(row => row.employee_id === id)) throw new Error('employee has gifts');
  return { ...remove(db, id, 'employees'), gifts: cascade ? db.gifts.filter(row => row.employee_id !== id) : db.gifts };
}
const deleteGift = (db, id) => remove(db, id, 'gifts');

function escapeXml(value) {
  return String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}
function pathJoin(base, target) {
  if (target.startsWith('/')) return target.replace(/^\/+/, '');
  const parts = [];
  for (const part of `${base}/${target}`.split('/')) {
    if (part === '..') parts.pop(); else if (part && part !== '.') parts.push(part);
  }
  return parts.join('/');
}
function attributes(tag) {
  const result = {};
  for (const match of tag.matchAll(/([\w:.-]+)="([^"]*)"/g)) result[match[1]] = match[2];
  return result;
}
function dateSerial(value, date1904) {
  const day = Date.parse(`${value}T00:00:00Z`) / 86400000;
  if (!Number.isFinite(day)) throw new Error(`Invalid birth_date: ${value}`);
  return day + (date1904 ? 24107 : 25569);
}
function sheetXml(xml, name, records, date1904) {
  const dataMatch = /<(?:\w+:)?sheetData\b[^>]*>([\s\S]*?)<\/(?:\w+:)?sheetData>/.exec(xml);
  if (!dataMatch) throw new Error(`Missing sheetData in ${name}`);
  const rows = [...dataMatch[1].matchAll(/<(?:\w+:)?row\b[^>]*>[\s\S]*?<\/(?:\w+:)?row>/g)].map(m => m[0]);
  const header = rows.find(row => /\br="1"/.test(row));
  if (!header) throw new Error(`Missing header in ${name}`);
  const styles = new Map();
  for (const row of rows) {
    const number = Number(attributes(row.match(/<[^>]*row\b[^>]*>/)[0]).r);
    if (number < 2) continue;
    const cells = new Map();
    for (const cell of row.matchAll(/<(?:\w+:)?c\b[^>]*\br="([A-Z]+)\d+"[^>]*\/?>(?:[\s\S]*?<\/(?:\w+:)?c>)?/g)) {
      cells.set(cell[1], attributes(cell[0].split('>')[0]).s);
    }
    styles.set(number, cells);
  }
  const columns = COLUMNS[name];
  const dataRows = records.map((record, index) => {
    const number = index + 2;
    const style = styles.get(number) || styles.get(2) || new Map();
    const cells = columns.map((column, col) => {
      const ref = `${String.fromCharCode(65 + col)}${number}`;
      const value = record[column];
      const s = style.get(ref.replace(/\d+$/, '')) ?? (column === 'birth_date' ? '4' : undefined);
      const attr = ` r="${ref}"${s === undefined ? '' : ` s="${escapeXml(s)}"`}`;
      if (value == null || value === '') return `<c${attr}/>`;
      if (column === 'birth_date' || column === 'year' || column === 'gift_cost') {
        const num = column === 'birth_date' ? dateSerial(value, date1904) : Number(value);
        if (!Number.isFinite(num)) throw new Error(`Invalid ${column}`);
        return `<c${attr}><v>${num}</v></c>`;
      }
      const text = String(value);
      return `<c${attr} t="inlineStr"><is><t${/^\s|\s$/.test(text) ? ' xml:space="preserve"' : ''}>${escapeXml(text)}</t></is></c>`;
    }).join('');
    return `<row r="${number}">${cells}</row>`;
  });
  const prefix = /<([\w]+):sheetData\b/.exec(dataMatch[0])?.[1];
  const tag = prefix ? `${prefix}:sheetData` : 'sheetData';
  let result = xml.replace(dataMatch[0], `<${tag}>${header}${dataRows.join('')}</${tag}>`);
  result = result.replace(/(<(?:\w+:)?dimension\b[^>]*\bref=")[^"]+("[^>]*\/?>)/, `$1A1:${String.fromCharCode(64 + columns.length)}${Math.max(2, records.length + 1)}$2`);
  return result;
}

async function patchZip(input, db) {
  const zip = await JSZip.loadAsync(input);
  const workbook = await zip.file('xl/workbook.xml').async('string');
  const rels = await zip.file('xl/_rels/workbook.xml.rels').async('string');
  const date1904 = /\bdate1904="(?:1|true)"/.test(workbook);
  const relationships = new Map([...rels.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map(m => {
    const a = attributes(m[0]); return [a.Id, a.Target];
  }));
  const tables = Object.keys(zip.files).filter(path => /^xl\/tables\/[^/]+\.xml$/.test(path));
  for (const [name, columns] of Object.entries(COLUMNS)) {
    const sheet = [...workbook.matchAll(/<(?:\w+:)?sheet\s[^>]*\/?\s*>/g)].map(m => attributes(m[0])).find(a => a.name === name);
    if (!sheet) throw new Error(`Missing worksheet: ${name}`);
    const path = pathJoin('xl', relationships.get(sheet['r:id']));
    const entry = zip.file(path);
    if (!entry) throw new Error(`Missing worksheet file: ${path}`);
    zip.file(path, sheetXml(await entry.async('string'), name, db[name], date1904));
    const tableName = `${name}Table`;
    let found = false;
    for (const tablePath of tables) {
      const table = await zip.file(tablePath).async('string');
      if (attributes(table.match(/<(?:\w+:)?table\b[^>]*>/)?.[0] || '').name !== tableName) continue;
      const last = Math.max(1, db[name].length + 1);
      const ref = `A1:${String.fromCharCode(64 + columns.length)}${last}`;
      let updated = table.replace(/(\bref=")[^"]+(")/, `$1${ref}$2`);
      updated = updated.replace(/(<(?:\w+:)?autoFilter\b[^>]*\bref=")[^"]+(")/, `$1${ref}$2`);
      zip.file(tablePath, updated);
      found = true;
      break;
    }
    if (!found) throw new Error(`Missing Excel table: ${tableName}`);
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

async function serializeWorkbook(originalBuffer, db) {
  return patchZip(originalBuffer, db);
}

const api = { parseWorkbook, validateFacility, validateEmployee, validateGift,
  createFacility, createEmployee, createGift, updateFacility, updateEmployee, updateGift,
  deleteFacility, deleteEmployee, deleteGift, serializeWorkbook };
if (typeof module === 'object' && module.exports) {
  module.exports = api;
} else {
  globalThis.BirthdayDB = api;
}
})();
