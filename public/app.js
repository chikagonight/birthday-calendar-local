/* global XLSX, JSZip */
"use strict";

const APP_SHEET = "_BirthdayApp";
const MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
const MONTHS_GENITIVE = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
const WEEKDAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];
const DEFAULT_COLORS = ["#4c5bd4", "#d86b55", "#3b9278", "#9b6bc2", "#d49335", "#3b7ca6", "#bf5275", "#6d8d3c"];
const HEADERS = ["ID", "Город", "Объект", "Цвет объекта", "ФИО", "Должность", "День", "Месяц", "Год", "Заметка", "Подарок"];

const state = {
  employees: [], view: "calendar", cursor: new Date(), search: "", city: "", object: "",
  fileHandle: null, fileName: "", rawBytes: null, workbook: null, writable: false, saveQueue: Promise.resolve(),
};

const $ = (selector) => document.querySelector(selector);
const esc = (value) => String(value ?? "").replace(/[&<>'"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[c]);
const xmlEsc = (value) => String(value ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[c]);
const normalize = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

function showError(message) {
  const node = $("#error");
  node.innerHTML = `<span>${esc(message)}</span><button aria-label="Закрыть">×</button>`;
  node.classList.remove("hidden");
  node.querySelector("button").onclick = () => node.classList.add("hidden");
}

function toast(message) {
  const node = $("#toast");
  node.textContent = message;
  node.classList.remove("hidden");
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => node.classList.add("hidden"), 3200);
}

function colorFor(name) {
  let hash = 0;
  for (const char of String(name || "")) hash = ((hash << 5) - hash + char.charCodeAt(0)) | 0;
  return DEFAULT_COLORS[Math.abs(hash) % DEFAULT_COLORS.length];
}

function withAlpha(hex, alpha = "22") {
  return /^#[0-9a-f]{6}$/i.test(hex || "") ? `${hex}${alpha}` : `#4c5bd4${alpha}`;
}

function safeBirthday(year, month, day) {
  return new Date(year, month - 1, Math.min(day, new Date(year, month, 0).getDate()));
}

function daysUntil(employee, now = new Date()) {
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  let next = safeBirthday(today.getFullYear(), employee.birthMonth, employee.birthDay);
  if (next < today) next = safeBirthday(today.getFullYear() + 1, employee.birthMonth, employee.birthDay);
  return Math.round((next - today) / 86400000);
}

function nextAge(employee) {
  if (!employee.birthYear) return null;
  const now = new Date();
  const thisYear = safeBirthday(now.getFullYear(), employee.birthMonth, employee.birthDay);
  return (thisYear < new Date(now.getFullYear(), now.getMonth(), now.getDate()) ? now.getFullYear() + 1 : now.getFullYear()) - employee.birthYear;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return { birthDay: value.getDate(), birthMonth: value.getMonth() + 1, birthYear: value.getFullYear() };
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return { birthDay: date.d, birthMonth: date.m, birthYear: date.y };
  }
  const match = normalize(value).match(/(?:^|\D)(\d{1,2})[.\/-](\d{1,2})(?:[.\/-](\d{2,4}))?(?:\D|$)/);
  if (!match) return null;
  let year = match[3] ? Number(match[3]) : null;
  if (year !== null && year < 100) year += year >= 40 ? 1900 : 2000;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const check = new Date(year ?? 2000, month - 1, day);
  if (month < 1 || month > 12 || check.getMonth() !== month - 1 || check.getDate() !== day) return null;
  return { birthDay: day, birthMonth: month, birthYear: year };
}

function richBoldName(cell) {
  if (!cell?.r) return "";
  try {
    const xml = new DOMParser().parseFromString(`<root>${cell.r}</root>`, "application/xml");
    if (xml.querySelector("parsererror")) return "";
    const runs = [...xml.getElementsByTagName("r")];
    let name = "";
    let started = false;
    for (const run of runs) {
      const text = run.getElementsByTagName("t")[0]?.textContent || "";
      const bold = run.getElementsByTagName("b").length > 0;
      const onlyDate = /^\s*\d{1,2}[.\/-]\d{1,2}(?:[.\/-]\d{2,4})?\s*$/.test(text);
      if (!started && bold && !onlyDate) { started = true; name += text; continue; }
      if (started && bold && !onlyDate) { name += text; continue; }
      if (started && text.trim()) break;
      if (!started && text.trim() && !bold) break;
    }
    return normalize(name).replace(/[()[\]]+$/g, "").trim();
  } catch { return ""; }
}

function parseLegacyCell(cell, city) {
  const original = String(cell?.v ?? "");
  const date = parseDateValue(original);
  if (!date) return null;
  const datePattern = /\d{1,2}[.\/-]\d{1,2}(?:[.\/-]\d{2,4})?/g;
  let fullName = richBoldName(cell);
  const withoutDate = original.replace(datePattern, " ").replace(/\s+/g, " ").trim();
  if (!fullName) {
    const bracket = withoutDate.indexOf("(");
    if (bracket > 0) fullName = normalize(withoutDate.slice(0, bracket));
  }
  if (!fullName) return null;
  let position = withoutDate;
  const nameIndex = position.toLocaleLowerCase("ru").indexOf(fullName.toLocaleLowerCase("ru"));
  if (nameIndex >= 0) position = position.slice(nameIndex + fullName.length);
  position = position.trim().replace(/^[-–—,:;\s]+/, "").replace(/^\((.*)\)$/s, "$1").trim();
  return {
    id: 0, city, object: "", objectColor: colorFor(city), fullName, position,
    ...date, note: "", gift: "",
  };
}

function parseLegacyWorkbook(workbook) {
  const found = [];
  const seen = new Set();
  for (const city of workbook.SheetNames) {
    if (city === APP_SHEET) continue;
    const sheet = workbook.Sheets[city];
    for (const [address, cell] of Object.entries(sheet)) {
      if (address.startsWith("!") || !/^[A-Z]+\d+$/.test(address)) continue;
      const employee = parseLegacyCell(cell, city);
      if (!employee) continue;
      const key = `${city}|${employee.fullName.toLocaleLowerCase("ru")}|${employee.birthDay}|${employee.birthMonth}|${employee.birthYear ?? ""}`;
      if (seen.has(key)) continue;
      seen.add(key);
      found.push(employee);
    }
  }
  return found.map((employee, index) => ({ ...employee, id: index + 1 }));
}

function readAppSheet(workbook) {
  const sheet = workbook.Sheets[APP_SHEET];
  if (!sheet) return null;
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: true });
  const employees = rows.map((row, index) => ({
    id: Number(row["ID"]) || index + 1,
    city: normalize(row["Город"]),
    object: normalize(row["Объект"]),
    objectColor: /^#[0-9a-f]{6}$/i.test(row["Цвет объекта"]) ? row["Цвет объекта"] : colorFor(row["Объект"] || row["Город"]),
    fullName: normalize(row["ФИО"]),
    position: normalize(row["Должность"]),
    birthDay: Number(row["День"]),
    birthMonth: Number(row["Месяц"]),
    birthYear: row["Год"] === "" ? null : Number(row["Год"]),
    note: normalize(row["Заметка"]),
    gift: normalize(row["Подарок"]),
  })).filter((row) => row.fullName && row.birthDay && row.birthMonth);
  return employees;
}

function openHandleDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("birthday-calendar", 1);
    request.onupgradeneeded = () => request.result.createObjectStore("settings");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function rememberHandle(handle) {
  const db = await openHandleDb();
  await new Promise((resolve, reject) => {
    const tx = db.transaction("settings", "readwrite");
    tx.objectStore("settings").put(handle, "excelHandle");
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
  db.close();
}

async function recalledHandle() {
  const db = await openHandleDb();
  const value = await new Promise((resolve, reject) => {
    const tx = db.transaction("settings", "readonly");
    const request = tx.objectStore("settings").get("excelHandle");
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
  db.close();
  return value;
}

async function selectDatabase() {
  if (!window.showOpenFilePicker) {
    $("#file-input").click();
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({ multiple: false, types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }] });
    const permission = await handle.requestPermission({ mode: "readwrite" });
    if (permission !== "granted") throw new Error("Не предоставлено разрешение на изменение Excel-файла.");
    await rememberHandle(handle);
    await loadFromHandle(handle);
  } catch (e) { if (e.name !== "AbortError") showError(e.message); }
}

async function createDatabase() {
  if (!window.showSaveFilePicker) return showError("Создание базы поддерживается в актуальных Chrome и Edge.");
  try {
    const handle = await window.showSaveFilePicker({ suggestedName: "Дни рождения.xlsx", types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }] });
    const city = normalize(prompt("Название первого листа (город):", "Москва")) || "Город";
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([MONTHS]), city.slice(0, 31));
    const bytes = new Uint8Array(XLSX.write(workbook, { type: "array", bookType: "xlsx" }));
    const writable = await handle.createWritable();
    await writable.write(bytes);
    await writable.close();
    await rememberHandle(handle);
    await loadFromHandle(handle);
    await persistEmployees();
    toast("Excel-база создана");
  } catch (e) { if (e.name !== "AbortError") showError(e.message); }
}

async function loadFromHandle(handle) {
  const file = await handle.getFile();
  const bytes = new Uint8Array(await file.arrayBuffer());
  await loadWorkbook(bytes, file.name, handle, true);
}

async function loadWorkbook(bytes, name, handle = null, writable = false) {
  const workbook = XLSX.read(bytes, { type: "array", cellStyles: true, cellHTML: true, cellDates: true });
  const stored = readAppSheet(workbook);
  const employees = stored ?? parseLegacyWorkbook(workbook);
  state.rawBytes = bytes;
  state.workbook = workbook;
  state.fileHandle = handle;
  state.fileName = name;
  state.writable = writable;
  state.employees = employees;
  state.city = "";
  state.object = "";
  updateDbStatus(stored ? "База загружена" : `Распознано записей: ${employees.length}. Скрытый лист базы создастся при сохранении.`);
  render();
  toast(`Загружено сотрудников: ${employees.length}`);
}

function updateDbStatus(detail) {
  const node = $("#db-status");
  node.classList.toggle("connected", Boolean(state.rawBytes));
  node.innerHTML = `<span class="status-dot"></span><div><strong>${state.fileName ? esc(state.fileName) : "Excel-база не выбрана"}</strong><small>${esc(detail || "Выберите существующий файл или создайте новый")}</small></div>`;
  $("#add-btn").disabled = !state.rawBytes;
  $("#objects-btn").disabled = !state.rawBytes;
  $("#save-copy-btn").disabled = !state.rawBytes;
}

function parseXml(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("Повреждена структура Excel-файла.");
  return doc;
}

function serializeXml(doc) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${new XMLSerializer().serializeToString(doc.documentElement)}`;
}

function columnName(index) {
  let result = "";
  while (index > 0) { index--; result = String.fromCharCode(65 + index % 26) + result; index = Math.floor(index / 26); }
  return result;
}

function appSheetXml(employees) {
  const rows = [HEADERS, ...employees.map((e) => [e.id, e.city, e.object, e.objectColor, e.fullName, e.position, e.birthDay, e.birthMonth, e.birthYear ?? "", e.note, e.gift])];
  const sheetRows = rows.map((row, rowIndex) => `<row r="${rowIndex + 1}">${row.map((value, colIndex) => {
    const ref = `${columnName(colIndex + 1)}${rowIndex + 1}`;
    if (typeof value === "number" && Number.isFinite(value)) return `<c r="${ref}"><v>${value}</v></c>`;
    const text = xmlEsc(value);
    const preserve = /^\s|\s$/.test(String(value)) ? ' xml:space="preserve"' : "";
    return `<c r="${ref}" t="inlineStr"><is><t${preserve}>${text}</t></is></c>`;
  }).join("")}</row>`).join("");
  const last = `${columnName(HEADERS.length)}${rows.length}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${last}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData>${sheetRows}</sheetData></worksheet>`;
}

async function upsertAppSheet(bytes, employees) {
  const zip = await JSZip.loadAsync(bytes);
  const workbookPath = "xl/workbook.xml";
  const relsPath = "xl/_rels/workbook.xml.rels";
  const typesPath = "[Content_Types].xml";
  const workbookDoc = parseXml(await zip.file(workbookPath).async("string"));
  const relsDoc = parseXml(await zip.file(relsPath).async("string"));
  const typesDoc = parseXml(await zip.file(typesPath).async("string"));
  const sheets = [...workbookDoc.getElementsByTagNameNS("*", "sheet")];
  let targetSheet = sheets.find((sheet) => sheet.getAttribute("name") === APP_SHEET);
  let sheetPath;
  if (targetSheet) {
    const relId = targetSheet.getAttribute("r:id") || targetSheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
    const relationship = [...relsDoc.getElementsByTagNameNS("*", "Relationship")].find((rel) => rel.getAttribute("Id") === relId);
    if (!relationship) throw new Error("Не найдена связь скрытого листа Excel.");
    sheetPath = pathJoin("xl", relationship.getAttribute("Target"));
  } else {
    const usedIds = [...relsDoc.getElementsByTagNameNS("*", "Relationship")].map((rel) => Number((rel.getAttribute("Id") || "").replace(/\D/g, "")) || 0);
    const relId = `rId${Math.max(0, ...usedIds) + 1}`;
    const sheetNumbers = Object.keys(zip.files).map((name) => name.match(/^xl\/worksheets\/sheet(\d+)\.xml$/)?.[1]).filter(Boolean).map(Number);
    const sheetNumber = Math.max(0, ...sheetNumbers) + 1;
    sheetPath = `xl/worksheets/sheet${sheetNumber}.xml`;
    const sheetId = Math.max(0, ...sheets.map((sheet) => Number(sheet.getAttribute("sheetId")) || 0)) + 1;
    const sheetNode = workbookDoc.createElementNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "sheet");
    sheetNode.setAttribute("name", APP_SHEET);
    sheetNode.setAttribute("sheetId", String(sheetId));
    sheetNode.setAttribute("state", "veryHidden");
    sheetNode.setAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "r:id", relId);
    workbookDoc.getElementsByTagNameNS("*", "sheets")[0].appendChild(sheetNode);
    const relNode = relsDoc.createElementNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
    relNode.setAttribute("Id", relId);
    relNode.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
    relNode.setAttribute("Target", `worksheets/sheet${sheetNumber}.xml`);
    relsDoc.documentElement.appendChild(relNode);
    const override = typesDoc.createElementNS("http://schemas.openxmlformats.org/package/2006/content-types", "Override");
    override.setAttribute("PartName", `/${sheetPath}`);
    override.setAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
    typesDoc.documentElement.appendChild(override);
  }
  zip.file(workbookPath, serializeXml(workbookDoc));
  zip.file(relsPath, serializeXml(relsDoc));
  zip.file(typesPath, serializeXml(typesDoc));
  zip.file(sheetPath, appSheetXml(employees));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE", compressionOptions: { level: 6 } });
}

function pathJoin(base, target) {
  if (String(target).startsWith("/")) return String(target).replace(/^\/+/, "");
  const parts = `${base}/${target}`.split("/");
  const result = [];
  for (const part of parts) { if (!part || part === ".") continue; if (part === "..") result.pop(); else result.push(part); }
  return result.join("/");
}

async function persistEmployees() {
  const snapshot = state.employees.map((row) => ({ ...row }));
  state.saveQueue = state.saveQueue.catch(() => {}).then(async () => {
    const updated = await upsertAppSheet(state.rawBytes, snapshot);
    if (!state.fileHandle || !state.writable) throw new Error("Файл открыт только для чтения. Используйте «Сохранить копию» в Chrome или Edge.");
    const permission = await state.fileHandle.queryPermission({ mode: "readwrite" });
    if (permission !== "granted") throw new Error("Нет разрешения на запись в выбранный Excel-файл.");
    const writable = await state.fileHandle.createWritable();
    await writable.write(updated);
    await writable.close();
    state.rawBytes = updated;
    updateDbStatus("Все изменения сохранены в Excel");
  });
  return state.saveQueue;
}

async function saveCopy() {
  if (!state.rawBytes) return;
  try {
    const updated = await upsertAppSheet(state.rawBytes, state.employees);
    if (window.showSaveFilePicker) {
      const handle = await window.showSaveFilePicker({ suggestedName: state.fileName || "Дни рождения.xlsx", types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }] });
      const writable = await handle.createWritable();
      await writable.write(updated);
      await writable.close();
    } else {
      const link = document.createElement("a");
      link.href = URL.createObjectURL(new Blob([updated], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      link.download = state.fileName || "Дни рождения.xlsx";
      link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    }
    toast("Копия Excel сохранена");
  } catch (e) { if (e.name !== "AbortError") showError(e.message); }
}

async function commitEmployees(nextEmployees, message) {
  const before = state.employees;
  state.employees = nextEmployees;
  render();
  try { await persistEmployees(); toast(message); }
  catch (e) { state.employees = before; render(); showError(e.message); }
}

function filteredEmployees() {
  const query = state.search.toLocaleLowerCase("ru").trim();
  return state.employees.filter((e) => (!query || [e.fullName, e.position, e.object, e.city, e.note].some((v) => String(v).toLocaleLowerCase("ru").includes(query))) && (!state.city || e.city === state.city) && (!state.object || e.object === state.object));
}

function render() {
  const upcoming = state.employees.map((employee) => ({ employee, days: daysUntil(employee) })).filter((x) => x.days <= 30).sort((a, b) => a.days - b.days || a.employee.fullName.localeCompare(b.employee.fullName, "ru"));
  const urgent = upcoming.filter((x) => x.days <= 4);
  $("#stat-total").textContent = state.employees.length;
  $("#stat-upcoming").textContent = upcoming.length;
  $("#stat-gifts").textContent = urgent.filter((x) => !x.employee.gift).length;
  $("#gift-stat").classList.toggle("attention", urgent.some((x) => !x.employee.gift));
  $("#upcoming-count").textContent = upcoming.length;
  $("#calendar-view").classList.toggle("active", state.view === "calendar");
  $("#list-view").classList.toggle("active", state.view === "list");
  renderFilters();
  renderMain();
  renderUpcoming(upcoming);
}

function renderFilters() {
  const cities = [...new Set(state.employees.map((e) => e.city).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru"));
  const objects = [...new Set(state.employees.map((e) => e.object).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru"));
  $("#city-filter").innerHTML = `<option value="">Все города</option>${cities.map((v) => `<option value="${esc(v)}" ${v === state.city ? "selected" : ""}>${esc(v)}</option>`).join("")}`;
  $("#object-filter").innerHTML = `<option value="">Все объекты</option>${objects.map((v) => `<option value="${esc(v)}" ${v === state.object ? "selected" : ""}>${esc(v)}</option>`).join("")}`;
}

function renderMain() {
  const root = $("#main-view");
  if (!state.rawBytes) {
    root.className = "list-card";
    root.innerHTML = `<div class="empty-state"><span class="empty-icon">▣</span><h3>Выберите Excel-базу</h3><p>Приложение прочитает все городские листы и сохранит данные в этом же файле.</p><div><button class="btn btn-quiet" id="empty-select">Выбрать Excel</button><button class="btn btn-primary" id="empty-create">Создать базу</button></div></div>`;
    $("#empty-select").onclick = selectDatabase;
    $("#empty-create").onclick = createDatabase;
    return;
  }
  const employees = filteredEmployees();
  root.className = state.view === "calendar" ? "calendar-card" : "list-card";
  if (state.view === "list") {
    root.innerHTML = `<div class="list-header"><span>Сотрудник</span><span>Дата рождения</span><span>Должность</span><span>Город / объект</span><span></span></div>${employees.length ? [...employees].sort((a, b) => daysUntil(a) - daysUntil(b)).map((e) => `<div class="list-row"><div class="person"><span class="avatar" style="color:${esc(e.objectColor)};background:${withAlpha(e.objectColor)}">${esc(e.fullName[0]?.toUpperCase() || "?")}</span><div><strong>${esc(e.fullName)}</strong>${e.note ? `<small>${esc(e.note)}</small>` : ""}</div></div><span>${e.birthDay} ${MONTHS_GENITIVE[e.birthMonth - 1]}${e.birthYear ? ` ${e.birthYear}` : ""}</span><span>${esc(e.position || "—")}</span><span>${esc(e.city || "—")}${e.object ? `<small class="object-label"><i style="background:${esc(e.objectColor)}"></i>${esc(e.object)}</small>` : ""}</span><button class="icon-btn" data-edit="${e.id}">✎</button></div>`).join("") : `<div class="empty-state"><h3>Ничего не найдено</h3><p>Измените фильтры или добавьте сотрудника.</p></div>`}`;
  } else {
    const year = state.cursor.getFullYear();
    const month = state.cursor.getMonth();
    const offset = (new Date(year, month, 1).getDay() + 6) % 7;
    const count = new Date(year, month + 1, 0).getDate();
    const today = new Date();
    const cells = Array.from({ length: 42 }, (_, i) => {
      const day = i - offset + 1;
      if (day < 1 || day > count) return `<div class="day-cell empty"></div>`;
      const people = employees.filter((e) => e.birthMonth === month + 1 && e.birthDay === day);
      const isToday = year === today.getFullYear() && month === today.getMonth() && day === today.getDate();
      return `<div class="day-cell ${isToday ? "today" : ""}"><span class="day-number">${day}</span><div class="birthday-pills">${people.map((e) => { const distance = daysUntil(e); const status = distance === 0 ? "now" : distance <= 4 ? "soon" : distance <= 14 ? "near" : ""; const color = e.object ? e.objectColor : "#8a91a2"; return `<button class="birthday-pill ${status}" data-edit="${e.id}" style="border-left-color:${esc(color)};background:${withAlpha(color, "18")}"><span>${esc(e.fullName)}</span><small>${esc([e.city, e.object].filter(Boolean).join(" · "))}</small></button>`; }).join("")}</div></div>`;
    }).join("");
    root.innerHTML = `<div class="calendar-head"><button class="icon-btn" id="prev-month">‹</button><div><h2>${MONTHS[month]} ${year}</h2><button id="today-btn">Сегодня</button></div><button class="icon-btn" id="next-month">›</button></div><div class="week-row">${WEEKDAYS.map((d) => `<span>${d}</span>`).join("")}</div><div class="calendar-grid">${cells}</div>`;
    $("#prev-month").onclick = () => { state.cursor = new Date(year, month - 1, 1); renderMain(); };
    $("#next-month").onclick = () => { state.cursor = new Date(year, month + 1, 1); renderMain(); };
    $("#today-btn").onclick = () => { state.cursor = new Date(); renderMain(); };
  }
  root.querySelectorAll("[data-edit]").forEach((button) => button.onclick = () => openEmployee(Number(button.dataset.edit)));
}

function renderUpcoming(rows) {
  const root = $("#upcoming-list");
  if (!rows.length) return void (root.innerHTML = `<div class="side-empty"><span style="font-size:26px">♢</span><span>В ближайший месяц дней рождений нет</span></div>`);
  root.innerHTML = rows.map(({ employee: e, days }) => `<article class="upcoming-card ${days === 0 ? "is-today" : days <= 4 ? "is-urgent" : ""}" style="border-left:3px solid ${esc(e.object ? e.objectColor : "#dfe2e9")}"><button class="upcoming-main" data-edit="${e.id}"><span class="date-badge"><strong>${e.birthDay}</strong><small>${MONTHS_GENITIVE[e.birthMonth - 1].slice(0, 3)}</small></span><span><strong>${esc(e.fullName)}</strong><small>${esc(e.position || [e.city, e.object].filter(Boolean).join(" · ") || "Сотрудник")}</small><em>${days === 0 ? "Сегодня" : days === 1 ? "Завтра" : `Через ${days} дн.`}${nextAge(e) ? ` · ${nextAge(e)} лет` : ""}</em></span></button>${days <= 4 ? `<label class="gift-field">♢<input class="gift-input" data-id="${e.id}" value="${esc(e.gift)}" placeholder="Название подарка"></label>` : ""}</article>`).join("");
  root.querySelectorAll("[data-edit]").forEach((button) => button.onclick = () => openEmployee(Number(button.dataset.edit)));
  root.querySelectorAll(".gift-input").forEach((input) => input.onchange = async () => {
    const id = Number(input.dataset.id);
    await commitEmployees(state.employees.map((e) => e.id === id ? { ...e, gift: input.value.trim() } : e), "Подарок сохранён");
  });
}

function openEmployee(id) {
  const current = id ? state.employees.find((e) => e.id === id) : null;
  const firstCity = state.city || [...new Set(state.employees.map((e) => e.city).filter(Boolean))][0] || state.workbook?.SheetNames.find((name) => name !== APP_SHEET) || "";
  const value = current || { fullName: "", birthDay: 1, birthMonth: 1, birthYear: "", position: "", city: firstCity, object: "", objectColor: colorFor("new"), note: "", gift: "" };
  const cities = [...new Set([...state.workbook.SheetNames.filter((name) => name !== APP_SHEET), ...state.employees.map((e) => e.city)].filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru"));
  $("#modal-root").innerHTML = `<div class="modal-backdrop"><form class="modal" id="employee-form"><div class="modal-head"><div><p class="eyebrow">${current ? "РЕДАКТИРОВАНИЕ" : "НОВАЯ ЗАПИСЬ"}</p><h2>${current ? esc(current.fullName) : "Добавить сотрудника"}</h2></div><button type="button" class="icon-btn" data-close>×</button></div><label class="field span-2"><span>ФИО *</span><input name="fullName" required autofocus value="${esc(value.fullName)}" placeholder="Иванов Иван Иванович"></label><fieldset class="date-fields span-2"><legend>Дата рождения *</legend><label><span>День</span><input name="birthDay" type="number" min="1" max="31" required value="${value.birthDay}"></label><label><span>Месяц</span><select name="birthMonth">${MONTHS.map((m, i) => `<option value="${i + 1}" ${value.birthMonth === i + 1 ? "selected" : ""}>${m}</option>`).join("")}</select></label><label><span>Год (необязательно)</span><input name="birthYear" type="number" min="1900" max="${new Date().getFullYear()}" value="${value.birthYear ?? ""}" placeholder="1988"></label></fieldset><label class="field"><span>Должность</span><input name="position" value="${esc(value.position)}" placeholder="Инженер ПТО"></label><label class="field"><span>Город (лист Excel)</span><input name="city" list="cities-list" required value="${esc(value.city)}" placeholder="Москва"><datalist id="cities-list">${cities.map((city) => `<option value="${esc(city)}"></option>`).join("")}</datalist></label><label class="field"><span>Объект</span><input name="object" id="object-input" value="${esc(value.object)}" placeholder="Название объекта"></label><label class="field color-field"><span>Цвет объекта</span><div><input name="objectColor" id="object-color" type="color" value="${esc(value.objectColor || colorFor(value.object))}"><output id="color-code">${esc(value.objectColor || colorFor(value.object))}</output></div></label><label class="field span-2"><span>Подарок</span><input name="gift" value="${esc(value.gift)}" placeholder="Название подарка"></label><label class="field span-2"><span>Заметка</span><textarea name="note" rows="3" placeholder="Дополнительная информация">${esc(value.note)}</textarea></label><div class="modal-actions span-2">${current ? `<button type="button" class="btn btn-danger" id="delete-employee">Удалить</button>` : "<span></span>"}<div><button type="button" class="btn btn-quiet" data-close>Отмена</button><button class="btn btn-primary">Сохранить в Excel</button></div></div></form></div>`;
  const close = () => $("#modal-root").innerHTML = "";
  document.querySelectorAll("[data-close]").forEach((button) => button.onclick = close);
  $("#object-color").oninput = (event) => $("#color-code").textContent = event.target.value;
  $("#object-input").onchange = (event) => {
    const existing = state.employees.find((e) => e.object.toLocaleLowerCase("ru") === event.target.value.trim().toLocaleLowerCase("ru"));
    if (existing) { $("#object-color").value = existing.objectColor; $("#color-code").textContent = existing.objectColor; }
  };
  $("#employee-form").onsubmit = async (event) => {
    event.preventDefault();
    const raw = Object.fromEntries(new FormData(event.target));
    const nextId = Math.max(0, ...state.employees.map((e) => Number(e.id) || 0)) + 1;
    const employee = { id: current?.id || nextId, fullName: normalize(raw.fullName), position: normalize(raw.position), city: normalize(raw.city), object: normalize(raw.object), objectColor: raw.objectColor, birthDay: Number(raw.birthDay), birthMonth: Number(raw.birthMonth), birthYear: raw.birthYear ? Number(raw.birthYear) : null, note: normalize(raw.note), gift: normalize(raw.gift) };
    const check = new Date(employee.birthYear || 2000, employee.birthMonth - 1, employee.birthDay);
    if (!employee.fullName || !employee.city || check.getMonth() !== employee.birthMonth - 1 || check.getDate() !== employee.birthDay) return showError("Проверьте ФИО, город и дату рождения.");
    let next = current ? state.employees.map((e) => e.id === current.id ? employee : e) : [...state.employees, employee];
    if (employee.object) next = next.map((e) => e.object.toLocaleLowerCase("ru") === employee.object.toLocaleLowerCase("ru") ? { ...e, objectColor: employee.objectColor } : e);
    close();
    await commitEmployees(next, "Запись сохранена в Excel");
  };
  if (current) $("#delete-employee").onclick = async () => {
    if (!confirm(`Удалить запись «${current.fullName}»?`)) return;
    close();
    await commitEmployees(state.employees.filter((e) => e.id !== current.id), "Запись удалена из Excel");
  };
}

function openObjectColors() {
  const objects = [...new Set(state.employees.map((e) => e.object).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru"));
  $("#modal-root").innerHTML = `<div class="modal-backdrop"><form class="modal object-modal" id="objects-form"><div class="modal-head"><div><p class="eyebrow">НАСТРОЙКА</p><h2>Цвета объектов</h2></div><button type="button" class="icon-btn" data-close>×</button></div>${objects.length ? `<div class="object-colors span-2">${objects.map((name, index) => { const color = state.employees.find((e) => e.object === name)?.objectColor || colorFor(name); return `<label><input type="color" name="color-${index}" value="${esc(color)}"><span>${esc(name)}</span></label>`; }).join("")}</div>` : `<p class="import-note span-2">Объекты появятся здесь после добавления сотрудников.</p>`}<div class="modal-actions span-2"><span></span><div><button type="button" class="btn btn-quiet" data-close>Отмена</button><button class="btn btn-primary" ${objects.length ? "" : "disabled"}>Сохранить в Excel</button></div></div></form></div>`;
  const close = () => $("#modal-root").innerHTML = "";
  document.querySelectorAll("[data-close]").forEach((button) => button.onclick = close);
  $("#objects-form").onsubmit = async (event) => {
    event.preventDefault();
    const form = new FormData(event.target);
    const colors = Object.fromEntries(objects.map((name, index) => [name, form.get(`color-${index}`)]));
    close();
    await commitEmployees(state.employees.map((e) => e.object && colors[e.object] ? { ...e, objectColor: colors[e.object] } : e), "Цвета объектов сохранены");
  };
}

$("#select-db-btn").onclick = selectDatabase;
$("#create-db-btn").onclick = createDatabase;
$("#objects-btn").onclick = openObjectColors;
$("#save-copy-btn").onclick = saveCopy;
$("#add-btn").onclick = () => openEmployee();
$("#calendar-view").onclick = () => { state.view = "calendar"; render(); };
$("#list-view").onclick = () => { state.view = "list"; render(); };
$("#search").oninput = (event) => { state.search = event.target.value; renderMain(); };
$("#city-filter").onchange = (event) => { state.city = event.target.value; renderMain(); };
$("#object-filter").onchange = (event) => { state.object = event.target.value; renderMain(); };
$("#file-input").onchange = async (event) => {
  const file = event.target.files[0];
  if (!file) return;
  try { await loadWorkbook(new Uint8Array(await file.arrayBuffer()), file.name, null, false); }
  catch (e) { showError(e.message); }
  finally { event.target.value = ""; }
};

async function restorePrevious() {
  updateDbStatus();
  render();
  if (!window.showOpenFilePicker) return;
  try {
    const handle = await recalledHandle();
    if (!handle) return;
    const permission = await handle.queryPermission({ mode: "readwrite" });
    if (permission === "granted") await loadFromHandle(handle);
    else updateDbStatus(`Ранее выбран: ${handle.name}. Нажмите «Выбрать Excel», чтобы снова разрешить доступ.`);
  } catch { /* Handle may have become unavailable. */ }
}

restorePrevious();
