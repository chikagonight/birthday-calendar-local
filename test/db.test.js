/*
 * Contract for public/db.js (CommonJS or browser global; no side effects on the supplied db):
 * @typedef {{employees: object[], facilities: object[], gifts: object[]}} Database
 * parseWorkbook(buffer: Buffer | Uint8Array): Database — Excel dates become YYYY-MM-DD; year
 * and populated gift_cost become numbers; empty data rows are ignored.
 * validateFacility(facility, existingFacilities): string[]
 * validateEmployee(employee, existingEmployees, facilities): string[]
 * validateGift(gift, existingGifts, employees): string[]
 * createFacility(db, data): {db, facility}
 * createEmployee(db, data): {db, employee}
 * createGift(db, data): {db, gift}
 * updateFacility(db, id, patch): {db, facility}
 * updateEmployee(db, id, patch): {db, employee}
 * updateGift(db, id, patch): {db, gift}
 * deleteFacility(db, id): Database — reject if employees refer to it.
 * deleteEmployee(db, id, {cascade = false} = {}): Database — cascade removes gifts.
 * deleteGift(db, id): Database
 * serializeWorkbook(originalBuffer: Buffer | Uint8Array, db: Database): Promise<Uint8Array> — preserve
 * the template's non-data content and Excel tables, updating their refs.
 * Creation generates a UUID once; updates retain it. Invalid writes throw Error.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const { webcrypto } = require('node:crypto');
const vm = require('node:vm');
const XLSX = require('../public/vendor/xlsx.full.min.js');
const JSZip = require('../public/vendor/jszip.min.js');
const {
  parseWorkbook, validateFacility, validateEmployee, validateGift,
  createFacility, createEmployee, createGift,
  updateFacility, updateEmployee, updateGift,
  deleteFacility, deleteEmployee, deleteGift, serializeWorkbook,
} = require('../public/db.js');

const template = '/mnt/c/Users/weed/Downloads/birthday_database.xlsx';
const emptyDb = () => ({ employees: [], facilities: [], gifts: [] });
const facility = { id: 'facility-1', name: 'Clinic', city: 'Moscow', color: '#AabbCC' };
const employee = {
  id: 'employee-1', last_name: 'Ivanov', first_name: 'Ivan', middle_name: 'I.',
  birth_date: '2024-01-01', position: 'Doctor', facility_id: facility.id, note: 'Note',
};
const gift = { id: 'gift-1', employee_id: employee.id, year: 2026, gift: 'Book', gift_cost: 123.5 };
const populatedDb = () => ({ employees: [{ ...employee }], facilities: [{ ...facility }], gifts: [{ ...gift }] });

function hasError(errors, field) {
  assert.ok(Array.isArray(errors), 'validation returns an array');
  assert.match(errors.join(' '), new RegExp(field), `expected an error for ${field}: ${errors}`);
}
function freezeDb(db) {
  for (const rows of Object.values(db)) {
    for (const row of rows) Object.freeze(row);
    Object.freeze(rows);
  }
  return Object.freeze(db);
}

test('parseWorkbook reads the real empty Excel template', () => {
  assert.deepEqual(parseWorkbook(readFileSync(template)), emptyDb());
});

test('parseWorkbook reads three populated sheets, normalizing numeric dates and gift types', () => {
  const wb = XLSX.utils.book_new();
  const sheets = {
    employees: [
      ['id', 'last_name', 'first_name', 'middle_name', 'birth_date', 'position', 'facility_id', 'note'],
      [employee.id, 'Ivanov', 'Ivan', 'I.', 45292, 'Doctor', facility.id, 'Note'],
      ['employee-2', 'Petrova', 'Anna', '', '1990-12-31', '', facility.id, ''],
    ],
    facilities: [
      ['id', 'name', 'city', 'color'],
      [facility.id, 'Clinic', 'Moscow', '#AabbCC'],
    ],
    gifts: [
      ['id', 'employee_id', 'year', 'gift', 'gift_cost'],
      [gift.id, employee.id, '2026', 'Book', '123.5'],
      ['gift-2', 'employee-2', 2027, 'Flowers', 0],
    ],
  };
  for (const [name, rows] of Object.entries(sheets)) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
  }
  const parsed = parseWorkbook(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  assert.deepEqual(parsed, {
    employees: [employee, {
      id: 'employee-2', last_name: 'Petrova', first_name: 'Anna', middle_name: '',
      birth_date: '1990-12-31', position: '', facility_id: facility.id, note: '',
    }],
    facilities: [facility],
    gifts: [gift, { id: 'gift-2', employee_id: 'employee-2', year: 2027, gift: 'Flowers', gift_cost: 0 }],
  });
  assert.equal(typeof parsed.employees[0].birth_date, 'string');
  assert.equal(typeof parsed.gifts[0].year, 'number');
  assert.equal(typeof parsed.gifts[0].gift_cost, 'number');
});

test('facility validation requires name and city and checks optional hex color', () => {
  assert.deepEqual(validateFacility({ name: 'Clinic', city: 'Moscow' }, []), []);
  assert.deepEqual(validateFacility(facility, []), []);
  for (const field of ['name', 'city']) hasError(validateFacility({ ...facility, [field]: '' }, []), field);
  for (const color of ['red', '#12345', '#1234567', '#GG0000']) {
    hasError(validateFacility({ ...facility, color }, []), 'color');
  }
});

test('employee validation requires identity, date and an existing facility', () => {
  assert.deepEqual(validateEmployee(employee, [], [facility]), []);
  assert.deepEqual(validateEmployee({ ...employee, middle_name: '', position: '', note: '' }, [], [facility]), []);
  for (const field of ['last_name', 'first_name', 'birth_date', 'facility_id']) {
    hasError(validateEmployee({ ...employee, [field]: '' }, [], [facility]), field);
  }
  hasError(validateEmployee({ ...employee, facility_id: 'unknown' }, [], [facility]), 'facility_id');
});

test('gift validation checks references, integer year, nonnegative cost, and employee/year uniqueness', () => {
  assert.deepEqual(validateGift(gift, [], [employee]), []);
  assert.deepEqual(validateGift({ ...gift, gift_cost: '' }, [], [employee]), []);
  for (const field of ['employee_id', 'year']) {
    hasError(validateGift({ ...gift, [field]: '' }, [], [employee]), field);
  }
  hasError(validateGift({ ...gift, employee_id: 'unknown' }, [], [employee]), 'employee_id');
  for (const year of [2026.5, '2026.5', 'not a year']) {
    hasError(validateGift({ ...gift, year }, [], [employee]), 'year');
  }
  for (const gift_cost of [-1, 'not a cost']) {
    hasError(validateGift({ ...gift, gift_cost }, [], [employee]), 'gift_cost');
  }
  hasError(validateGift({ ...gift, id: 'gift-2' }, [gift], [employee]), 'year');
  assert.deepEqual(validateGift(gift, [gift], [employee]), [], 'an unchanged gift may be revalidated');
});

test('create operations generate distinct UUIDs and never mutate the supplied db', () => {
  const initial = freezeDb(emptyDb());
  const { db: withFacility, facility: first } = createFacility(initial, { name: 'Clinic', city: 'Moscow' });
  const { db: withTwoFacilities, facility: second } = createFacility(withFacility, { name: 'School', city: 'Tula' });
  const { db: withEmployee, employee: person } = createEmployee(withTwoFacilities, {
    last_name: 'Ivanov', first_name: 'Ivan', birth_date: '2024-01-01', facility_id: first.id,
  });
  const { db: final, gift: present } = createGift(withEmployee, { employee_id: person.id, year: 2026 });
  assert.deepEqual(initial, emptyDb());
  assert.equal(withFacility.employees.length, 0);
  assert.equal(withFacility.facilities.length, 1);
  assert.equal(withTwoFacilities.facilities.length, 2);
  assert.equal(withEmployee.gifts.length, 0);
  const ids = [first.id, second.id, person.id, present.id];
  assert.equal(new Set(ids).size, ids.length);
  for (const id of ids) assert.match(id, /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i);
  assert.deepEqual(final, { facilities: [first, second], employees: [person], gifts: [present] });
});

test('create rejects invalid fields, dangling references and duplicate gifts', () => {
  const db = freezeDb(populatedDb());
  assert.throws(() => createFacility(db, { name: 'X', city: 'Y', color: 'red' }), /color/i);
  assert.throws(() => createEmployee(db, { ...employee, facility_id: 'missing' }), /facility_id/i);
  assert.throws(() => createEmployee(db, { ...employee, first_name: '' }), /first_name/i);
  assert.throws(() => createGift(db, { employee_id: 'missing', year: 2027 }), /employee_id/i);
  assert.throws(() => createGift(db, { employee_id: employee.id, year: 2026 }), /year/i);
  assert.throws(() => createGift(db, { employee_id: employee.id, year: 2027.2 }), /year/i);
  assert.deepEqual(db, populatedDb());
});

test('updates retain ids, validate merged records and do not change the input', () => {
  const db = freezeDb(populatedDb());
  const { db: changedFacility, facility: renamed } = updateFacility(db, facility.id, { name: 'Hospital' });
  const { db: changedEmployee, employee: renamedPerson } = updateEmployee(db, employee.id, { last_name: 'Sidorov' });
  const { db: changedGift, gift: changedPresent } = updateGift(db, gift.id, { year: 2027, gift_cost: 0 });
  assert.deepEqual(renamed, { ...facility, name: 'Hospital' });
  assert.deepEqual(renamedPerson, { ...employee, last_name: 'Sidorov' });
  assert.deepEqual(changedPresent, { ...gift, year: 2027, gift_cost: 0 });
  assert.deepEqual(changedFacility.facilities, [renamed]);
  assert.deepEqual(changedEmployee.employees, [renamedPerson]);
  assert.deepEqual(changedGift.gifts, [changedPresent]);
  assert.deepEqual(db, populatedDb());
  assert.throws(() => updateFacility(db, facility.id, { color: 'bad' }), /color/i);
  assert.throws(() => updateEmployee(db, employee.id, { facility_id: 'missing' }), /facility_id/i);
  assert.throws(() => updateGift(db, gift.id, { employee_id: 'missing' }), /employee_id/i);
  assert.throws(() => updateGift(db, gift.id, { gift_cost: -1 }), /gift_cost/i);
  assert.throws(() => updateFacility(db, facility.id, { id: 'changed' }), /id/i);
  assert.throws(() => updateEmployee(db, employee.id, { id: 'changed' }), /id/i);
  assert.throws(() => updateGift(db, gift.id, { id: 'changed' }), /id/i);
});

test('updating a gift cannot collide with another employee/year pair', () => {
  const db = freezeDb({ ...populatedDb(), gifts: [gift, { ...gift, id: 'gift-2', year: 2027 }] });
  assert.throws(() => updateGift(db, 'gift-2', { year: 2026 }), /year/i);
  assert.deepEqual(db.gifts[1].year, 2027);
});

test('deleteFacility always blocks referenced facilities; deletes unreferenced ones immutably', () => {
  const db = freezeDb(populatedDb());
  assert.throws(() => deleteFacility(db, facility.id), /employee|сотрудник/i);
  assert.deepEqual(db, populatedDb());
  const withExtra = freezeDb({ ...populatedDb(), facilities: [facility, { ...facility, id: 'facility-2' }] });
  assert.deepEqual(deleteFacility(withExtra, 'facility-2'), populatedDb());
  assert.equal(withExtra.facilities.length, 2);
});

test('deleteEmployee blocks gifts unless cascade is explicit, then removes only linked gifts', () => {
  const other = { ...employee, id: 'employee-2' };
  const otherGift = { ...gift, id: 'gift-2', employee_id: other.id };
  const db = freezeDb({ facilities: [facility], employees: [employee, other], gifts: [gift, otherGift] });
  assert.throws(() => deleteEmployee(db, employee.id), /gift|подар/i);
  assert.throws(() => deleteEmployee(db, employee.id, { cascade: false }), /gift|подар/i);
  assert.deepEqual(deleteEmployee(db, employee.id, { cascade: true }), {
    facilities: [facility], employees: [other], gifts: [otherGift],
  });
  assert.equal(db.employees.length, 2);
  assert.equal(db.gifts.length, 2);
  assert.deepEqual(deleteEmployee({ ...db, gifts: [otherGift] }, employee.id), {
    facilities: [facility], employees: [other], gifts: [otherGift],
  });
});

test('deleteGift removes only its target without mutating the original', () => {
  const db = freezeDb(populatedDb());
  assert.deepEqual(deleteGift(db, gift.id), { ...populatedDb(), gifts: [] });
  assert.deepEqual(db, populatedDb());
});

test('browser script exposes BirthdayDB without Node globals', async () => {
  const context = vm.createContext({ crypto: webcrypto, XLSX, JSZip });
  vm.runInContext('var XLSX = globalThis.XLSX; var JSZip = globalThis.JSZip;', context);
  vm.runInContext(readFileSync(require.resolve('../public/db.js'), 'utf8'), context);
  const api = context.BirthdayDB;
  assert.deepEqual(Object.keys(api).sort(), Object.keys(require('../public/db.js')).sort());
  assert.match(api.createFacility(emptyDb(), { name: 'Clinic', city: 'Moscow' }).facility.id,
    /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i);
  const original = readFileSync(template);
  const result = await api.serializeWorkbook(original, emptyDb());
  assert.ok(result instanceof Uint8Array);
  assert.equal(api.parseWorkbook(result).employees.length, 0);
});

test('serializeWorkbook round-trips new rows and expands actual Excel table refs', async () => {
  const original = readFileSync(template);
  const db = {
    facilities: [facility, { ...facility, id: 'facility-2', name: 'School' }],
    employees: [employee, { ...employee, id: 'employee-2', first_name: 'Anna', facility_id: 'facility-2' }],
    gifts: [gift, { ...gift, id: 'gift-2', employee_id: 'employee-2', year: 2027 }],
  };
  const result = await serializeWorkbook(original, db);
  assert.ok(result instanceof Uint8Array);
  assert.equal(Buffer.isBuffer(result), false);
  assert.deepEqual(parseWorkbook(result), db);
  const [before, after] = await Promise.all([JSZip.loadAsync(original), JSZip.loadAsync(result)]);
  // The fourth sheet is documentation, not a ListObject or a data sheet.
  assert.deepEqual(await after.file('xl/worksheets/sheet4.xml').async('nodebuffer'),
    await before.file('xl/worksheets/sheet4.xml').async('nodebuffer'));
  for (const [tableName, expectedRef] of [
    ['employeesTable', 'A1:H3'], ['facilitiesTable', 'A1:D3'], ['giftsTable', 'A1:E3'],
  ]) {
    const paths = Object.keys(after.files).filter((p) => /^xl\/tables\/[^/]+\.xml$/.test(p));
    const xmls = await Promise.all(paths.map((p) => after.file(p).async('string')));
    const xml = xmls.find((text) => new RegExp(`(?:name|displayName)="${tableName}"`).test(text));
    assert.ok(xml, `${tableName} must remain a real Excel table`);
    assert.match(xml, new RegExp(`\\bref="${expectedRef}"`));
  }
});
