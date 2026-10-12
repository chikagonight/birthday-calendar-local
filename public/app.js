/* global BirthdayDB */
"use strict";

const MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"];
const MONTHS_GENITIVE = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
const WEEKDAYS = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];
const DEFAULT_COLORS = ["#4c5bd4", "#d86b55", "#3b9278", "#9b6bc2", "#d49335", "#3b7ca6", "#bf5275", "#6d8d3c"];
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const pickerOptions = { types: [{ description: "Excel", accept: { [XLSX_TYPE]: [".xlsx"] } }] };
const state = {
  db: { employees: [], facilities: [], gifts: [] }, view: "calendar", cursor: new Date(), search: "", city: "", object: "",
  fileHandle: null, fileName: "", rawBytes: null, writable: false, saveQueue: Promise.resolve(),
};

const $ = (selector) => document.querySelector(selector);
const esc = (value) => String(value ?? "").replace(/[&<>'"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[c]);
const normalize = (value) => String(value ?? "").replace(/\s+/g, " ").trim();
const fullName = (employee) => [employee.last_name, employee.first_name, employee.middle_name].filter(Boolean).join(" ");
const facilityFor = (employee) => state.db.facilities.find((facility) => facility.id === employee.facility_id);

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
function errorText(error) {
  let text = error.message || String(error);
  if (text.includes("year already exists for employee_id")) return "Подарок за этот год уже есть у сотрудника.";
  if (text.includes("facility has employees")) return "Нельзя удалить объект: к нему привязаны сотрудники. Сначала перенесите или удалите их.";
  if (text.includes("employee has gifts")) return "У сотрудника есть история подарков. Подтвердите удаление вместе с ней.";
  const labels = { last_name: "Фамилия", first_name: "Имя", birth_date: "Дата рождения", facility_id: "Объект", employee_id: "Сотрудник", gift_cost: "Стоимость", year: "Год", color: "Цвет", name: "Название", city: "Город" };
  for (const [field, label] of Object.entries(labels)) text = text.replace(new RegExp(`\\b${field}\\b`, "g"), label);
  return text.replace(/is required/g, "обязательно для заполнения").replace(/does not exist/g, "не найден(а)")
    .replace(/must be YYYY-MM-DD/g, "должна быть корректной датой").replace(/must be an integer/g, "должен быть целым числом")
    .replace(/must be nonnegative/g, "не может быть отрицательной").replace(/must be a hex RGB value/g, "должен быть в формате #RRGGBB")
    .replace(/id not found/g, "запись не найдена");
}
function colorFor(name) {
  let hash = 0;
  for (const char of String(name || "")) hash = ((hash << 5) - hash + char.charCodeAt(0)) | 0;
  return DEFAULT_COLORS[Math.abs(hash) % DEFAULT_COLORS.length];
}
const facilityColor = (facility) => /^#[0-9a-f]{6}$/i.test(facility?.color || "") ? facility.color : colorFor(facility?.name);
const withAlpha = (hex, alpha = "22") => /^#[0-9a-f]{6}$/i.test(hex || "") ? `${hex}${alpha}` : `#4c5bd4${alpha}`;
const birthdayParts = (employee) => employee.birth_date.split("-").map(Number);
function birthdayIn(year, month, day) {
  // 29 February is celebrated on 28 February in non-leap years.
  return Math.min(day, new Date(Date.UTC(year, month, 0)).getUTCDate());
}
function nextBirthday(employee, now = new Date()) {
  const [born, month, day] = birthdayParts(employee);
  const today = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  let year = now.getFullYear();
  let next = Date.UTC(year, month - 1, birthdayIn(year, month, day));
  if (next < today) next = Date.UTC(++year, month - 1, birthdayIn(year, month, day));
  return { days: Math.round((next - today) / 86400000), year, age: year - born };
}
const giftFor = (id, year) => state.db.gifts.find((gift) => gift.employee_id === id && Number(gift.year) === year);

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
  if (!window.showOpenFilePicker) { $("#file-input").click(); return; }
  try {
    const [handle] = await window.showOpenFilePicker({ multiple: false, ...pickerOptions });
    const permission = await handle.requestPermission({ mode: "readwrite" });
    if (permission !== "granted") throw new Error("Не предоставлено разрешение на изменение Excel-файла.");
    await state.saveQueue.catch(() => {});
    await loadFromHandle(handle);
    try { await rememberHandle(handle); } catch { showError("Файл загружен, но браузер не смог запомнить его для следующего запуска."); }
  } catch (e) { if (e.name !== "AbortError") showError(errorText(e)); }
}
async function createDatabase() {
  if (!window.showSaveFilePicker) return showError("Создание базы поддерживается в актуальных Chrome и Edge.");
  try {
    const handle = await window.showSaveFilePicker({ suggestedName: "Дни рождения.xlsx", ...pickerOptions });
    const response = await fetch("/template.xlsx");
    if (!response.ok) throw new Error("Не удалось загрузить шаблон базы.");
    const bytes = new Uint8Array(await response.arrayBuffer());
    BirthdayDB.parseWorkbook(bytes); // Do not overwrite a selected file with an invalid template.
    await state.saveQueue.catch(() => {});
    const writable = await handle.createWritable();
    await writable.write(bytes);
    await writable.close();
    await loadFromHandle(handle);
    try { await rememberHandle(handle); } catch { showError("База создана, но браузер не смог запомнить файл."); }
    toast("Excel-база создана");
  } catch (e) { if (e.name !== "AbortError") showError(errorText(e)); }
}
async function loadFromHandle(handle) {
  const file = await handle.getFile();
  await loadWorkbook(new Uint8Array(await file.arrayBuffer()), file.name, handle, true);
}
async function loadWorkbook(bytes, name, handle = null, writable = false) {
  const db = BirthdayDB.parseWorkbook(bytes);
  state.rawBytes = bytes;
  state.db = db;
  state.fileHandle = handle;
  state.fileName = name;
  state.writable = writable;
  state.city = "";
  state.object = "";
  updateDbStatus(writable ? "База загружена" : "Только чтение: для изменений откройте файл в Chrome/Edge");
  render();
  toast(`Загружено сотрудников: ${db.employees.length}`);
}
function updateDbStatus(detail) {
  const node = $("#db-status");
  node.classList.toggle("connected", Boolean(state.rawBytes));
  node.innerHTML = `<span class="status-dot"></span><div><strong>${state.fileName ? esc(state.fileName) : "Excel-база не выбрана"}</strong><small>${esc(detail || "Выберите существующий файл или создайте новый")}</small></div>`;
  $("#add-btn").disabled = !state.rawBytes || !state.writable || !state.db.facilities.length;
  $("#objects-btn").disabled = !state.rawBytes || !state.writable;
  $("#save-copy-btn").disabled = !state.rawBytes;
}

// Queue complete transactions (including the mutation), not just writes: rapid edits cannot overwrite each other.
function commitDb(change, message) {
  const pending = state.saveQueue.catch(() => {}).then(async () => {
    if (!state.fileHandle || !state.writable) throw new Error("Файл открыт только для чтения. Откройте его в Chrome/Edge для изменения.");
    const next = change(state.db);
    if (await state.fileHandle.queryPermission({ mode: "readwrite" }) !== "granted") throw new Error("Нет разрешения на запись в выбранный Excel-файл.");
    const bytes = await BirthdayDB.serializeWorkbook(state.rawBytes, next);
    const writable = await state.fileHandle.createWritable();
    await writable.write(bytes);
    await writable.close();
    state.rawBytes = bytes;
    state.db = next;
    updateDbStatus("Все изменения сохранены в Excel");
    render();
    toast(message);
  });
  state.saveQueue = pending;
  return pending.catch((error) => { showError(errorText(error)); throw error; });
}
async function saveCopy() {
  if (!state.rawBytes) return;
  try {
    await state.saveQueue.catch(() => {});
    const updated = await BirthdayDB.serializeWorkbook(state.rawBytes, state.db);
    if (window.showSaveFilePicker) {
      const handle = await window.showSaveFilePicker({ suggestedName: state.fileName || "Дни рождения.xlsx", ...pickerOptions });
      const writable = await handle.createWritable();
      await writable.write(updated);
      await writable.close();
    } else {
      const link = document.createElement("a");
      link.href = URL.createObjectURL(new Blob([updated], { type: XLSX_TYPE }));
      link.download = state.fileName || "Дни рождения.xlsx";
      link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 1000);
    }
    toast("Копия Excel сохранена");
  } catch (e) { if (e.name !== "AbortError") showError(errorText(e)); }
}
function filteredEmployees() {
  const query = state.search.toLocaleLowerCase("ru").trim();
  return state.db.employees.filter((employee) => {
    const facility = facilityFor(employee);
    return (!query || [fullName(employee), employee.position, facility?.name, facility?.city, employee.note].some((value) => String(value || "").toLocaleLowerCase("ru").includes(query)))
      && (!state.city || facility?.city === state.city) && (!state.object || facility?.name === state.object);
  });
}
function render() {
  const upcoming = state.db.employees.map((employee) => ({ employee, ...nextBirthday(employee) }))
    .filter(({ days }) => days <= 30).sort((a, b) => a.days - b.days || fullName(a.employee).localeCompare(fullName(b.employee), "ru"));
  const missing = upcoming.filter(({ employee, days, year }) => days <= 4 && !giftFor(employee.id, year));
  $("#stat-total").textContent = state.db.employees.length;
  $("#stat-upcoming").textContent = upcoming.length;
  $("#stat-gifts").textContent = missing.length;
  $("#gift-stat").classList.toggle("attention", missing.length > 0);
  $("#upcoming-count").textContent = upcoming.length;
  $("#calendar-view").classList.toggle("active", state.view === "calendar");
  $("#list-view").classList.toggle("active", state.view === "list");
  updateDbStatus(state.rawBytes ? (state.writable ? "Все изменения сохранены в Excel" : "Только чтение: для изменений откройте файл в Chrome/Edge") : undefined);
  renderFilters();
  renderMain();
  renderUpcoming(upcoming);
}
function renderFilters() {
  const cities = [...new Set(state.db.facilities.map((facility) => facility.city).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru"));
  const objects = [...new Set(state.db.facilities.map((facility) => facility.name).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ru"));
  $("#city-filter").innerHTML = `<option value="">Все города</option>${cities.map((value) => `<option value="${esc(value)}" ${value === state.city ? "selected" : ""}>${esc(value)}</option>`).join("")}`;
  $("#object-filter").innerHTML = `<option value="">Все объекты</option>${objects.map((value) => `<option value="${esc(value)}" ${value === state.object ? "selected" : ""}>${esc(value)}</option>`).join("")}`;
}
function renderMain() {
  const root = $("#main-view");
  if (!state.rawBytes || !state.db.facilities.length) {
    root.className = "list-card";
    root.innerHTML = state.rawBytes
      ? `<div class="empty-state"><span class="empty-icon">●</span><h3>Начало работы: сначала объекты, потом сотрудники, потом подарки</h3><p>Добавьте первый объект, чтобы создать сотрудника.</p><div><button class="btn btn-primary" id="empty-objects" ${state.writable ? "" : "disabled"}>Объекты</button></div></div>`
      : `<div class="empty-state"><span class="empty-icon">▣</span><h3>Выберите Excel-базу</h3><p>Откройте базу с таблицами сотрудников, объектов и подарков.</p><div><button class="btn btn-quiet" id="empty-select">Выбрать Excel</button><button class="btn btn-primary" id="empty-create">Создать базу</button></div></div>`;
    if (state.rawBytes) $("#empty-objects").onclick = openObjects;
    else { $("#empty-select").onclick = selectDatabase; $("#empty-create").onclick = createDatabase; }
    return;
  }
  const employees = filteredEmployees();
  root.className = state.view === "calendar" ? "calendar-card" : "list-card";
  if (state.view === "list") {
    root.innerHTML = `<div class="list-header"><span>Сотрудник</span><span>Дата рождения</span><span>Должность</span><span>Город / объект</span><span></span></div>${employees.length ? [...employees].sort((a, b) => nextBirthday(a).days - nextBirthday(b).days).map((employee) => {
      const facility = facilityFor(employee); const color = facilityColor(facility);
      const [year, month, day] = birthdayParts(employee);
      return `<div class="list-row"><div class="person"><span class="avatar" style="color:${esc(color)};background:${withAlpha(color)}">${esc(fullName(employee)[0]?.toUpperCase() || "?")}</span><div><strong>${esc(fullName(employee))}</strong>${employee.note ? `<small>${esc(employee.note)}</small>` : ""}</div></div><span>${day} ${MONTHS_GENITIVE[month - 1]} ${year}</span><span>${esc(employee.position || "—")}</span><span>${esc(facility?.city || "—")}<small class="object-label"><i style="background:${esc(color)}"></i>${esc(facility?.name || "—")}</small></span><button class="icon-btn" data-edit="${esc(employee.id)}" aria-label="Редактировать">✎</button></div>`;
    }).join("") : `<div class="empty-state"><h3>Ничего не найдено</h3><p>Измените фильтры или добавьте сотрудника.</p></div>`}`;
  } else {
    const year = state.cursor.getFullYear(); const month = state.cursor.getMonth();
    const offset = (new Date(year, month, 1).getDay() + 6) % 7;
    const count = new Date(year, month + 1, 0).getDate(); const today = new Date();
    const cells = Array.from({ length: 42 }, (_, index) => {
      const day = index - offset + 1;
      if (day < 1 || day > count) return `<div class="day-cell empty"></div>`;
      const people = employees.filter((employee) => { const [, m, d] = birthdayParts(employee); return m === month + 1 && birthdayIn(year, m, d) === day; });
      const isToday = year === today.getFullYear() && month === today.getMonth() && day === today.getDate();
      return `<div class="day-cell ${isToday ? "today" : ""}"><span class="day-number">${day}</span><div class="birthday-pills">${people.map((employee) => {
        const facility = facilityFor(employee); const color = facilityColor(facility); const distance = nextBirthday(employee).days;
        const status = distance === 0 ? "now" : distance <= 4 ? "soon" : distance <= 14 ? "near" : "";
        return `<button class="birthday-pill ${status}" data-edit="${esc(employee.id)}" style="border-left-color:${esc(color)};background:${withAlpha(color, "18")}"><span>${esc(fullName(employee))}</span><small>${esc([facility?.city, facility?.name].filter(Boolean).join(" · "))}</small></button>`;
      }).join("")}</div></div>`;
    }).join("");
    root.innerHTML = `<div class="calendar-head"><button class="icon-btn" id="prev-month">‹</button><div><h2>${MONTHS[month]} ${year}</h2><button id="today-btn">Сегодня</button></div><button class="icon-btn" id="next-month">›</button></div><div class="week-row">${WEEKDAYS.map((day) => `<span>${day}</span>`).join("")}</div><div class="calendar-grid">${cells}</div>`;
    $("#prev-month").onclick = () => { state.cursor = new Date(year, month - 1, 1); renderMain(); };
    $("#next-month").onclick = () => { state.cursor = new Date(year, month + 1, 1); renderMain(); };
    $("#today-btn").onclick = () => { state.cursor = new Date(); renderMain(); };
  }
  root.querySelectorAll("[data-edit]").forEach((button) => button.onclick = () => openEmployee(button.dataset.edit));
}
function renderUpcoming(rows) {
  const root = $("#upcoming-list");
  if (!rows.length) { root.innerHTML = `<div class="side-empty"><span style="font-size:26px">♢</span><span>В ближайший месяц дней рождений нет</span></div>`; return; }
  root.innerHTML = rows.map(({ employee, days, year, age }) => {
    const [, month, day] = birthdayParts(employee); const facility = facilityFor(employee); const gift = giftFor(employee.id, year);
    return `<article class="upcoming-card ${days === 0 ? "is-today" : days <= 4 ? "is-urgent" : ""}" style="border-left:3px solid ${esc(facilityColor(facility))}"><button class="upcoming-main" data-edit="${esc(employee.id)}"><span class="date-badge"><strong>${birthdayIn(year, month, day)}</strong><small>${MONTHS_GENITIVE[month - 1].slice(0, 3)}</small></span><span><strong>${esc(fullName(employee))}</strong><small>${esc(employee.position || [facility?.city, facility?.name].filter(Boolean).join(" · ") || "Сотрудник")}</small><em>${days === 0 ? "Сегодня" : days === 1 ? "Завтра" : `Через ${days} дн.`} · ${age} лет</em></span></button>${days <= 4 ? `<label class="gift-field">♢<input class="gift-input" data-id="${esc(employee.id)}" data-year="${year}" value="${esc(gift?.gift || "")}" placeholder="Подарок за ${year}" ${state.writable ? "" : "disabled"}></label>` : ""}</article>`;
  }).join("");
  root.querySelectorAll("[data-edit]").forEach((button) => button.onclick = () => openEmployee(button.dataset.edit));
  root.querySelectorAll(".gift-input").forEach((input) => input.onchange = async () => {
    const value = normalize(input.value); const id = input.dataset.id; const year = Number(input.dataset.year);
    if (value === (giftFor(id, year)?.gift || "") || (!value && !giftFor(id, year))) return;
    input.disabled = true;
    try {
      await commitDb((db) => {
        const existing = db.gifts.find((row) => row.employee_id === id && Number(row.year) === year);
        return existing ? BirthdayDB.updateGift(db, existing.id, { gift: value }).db
          : BirthdayDB.createGift(db, { employee_id: id, year, gift: value, gift_cost: "" }).db;
      }, "Подарок сохранён");
    } catch { input.disabled = false; }
  });
}
function closeModal() { $("#modal-root").innerHTML = ""; }
function modal(html) {
  $("#modal-root").innerHTML = `<div class="modal-backdrop">${html}</div>`;
  $("#modal-root").querySelectorAll("[data-close]").forEach((button) => button.onclick = closeModal);
}
function openObjects() {
  if (!state.writable) return;
  const rows = [...state.db.facilities].sort((a, b) => a.name.localeCompare(b.name, "ru"));
  modal(`<div class="modal"><div class="modal-head"><div><p class="eyebrow">БАЗА ДАННЫХ</p><h2>Объекты</h2></div><button class="icon-btn" data-close>×</button></div><div class="span-2 entity-list">${rows.length ? rows.map((facility) => `<div class="entity-row"><i class="facility-dot" style="background:${esc(facilityColor(facility))}"></i><div><strong>${esc(facility.name)}</strong><small>${esc(facility.city)}</small></div><button class="btn btn-quiet" data-facility-edit="${esc(facility.id)}">Изменить</button><button class="btn btn-danger" data-facility-delete="${esc(facility.id)}">Удалить</button></div>`).join("") : `<p class="hint">Начало работы: сначала объекты, потом сотрудники, потом подарки.</p>`}</div><div class="modal-actions span-2"><span></span><button class="btn btn-primary" id="new-facility">＋ Добавить объект</button></div></div>`);
  $("#new-facility").onclick = () => openFacility();
  $("#modal-root").querySelectorAll("[data-facility-edit]").forEach((button) => button.onclick = () => openFacility(button.dataset.facilityEdit));
  $("#modal-root").querySelectorAll("[data-facility-delete]").forEach((button) => button.onclick = async () => {
    const facility = state.db.facilities.find((row) => row.id === button.dataset.facilityDelete);
    if (!confirm(`Удалить объект «${facility.name}»?`)) return;
    try { await commitDb((db) => BirthdayDB.deleteFacility(db, facility.id), "Объект удалён"); openObjects(); } catch { /* Error shown by commitDb. */ }
  });
}
function openFacility(id) {
  const facility = state.db.facilities.find((row) => row.id === id);
  modal(`<form class="modal" id="facility-form"><div class="modal-head"><h2>${facility ? "Изменить объект" : "Добавить объект"}</h2><button type="button" class="icon-btn" data-close>×</button></div><label class="field"><span>Название *</span><input name="name" required value="${esc(facility?.name || "")}"></label><label class="field"><span>Город *</span><input name="city" required value="${esc(facility?.city || "")}"></label><div class="field color-field span-2"><span>Цвет (необязательно)</span><div><input name="color" type="color" value="${esc(facilityColor(facility))}"><label class="color-auto"><input name="autoColor" type="checkbox" ${facility?.color ? "" : "checked"}> Автоматический цвет</label></div></div><div class="modal-actions span-2"><button type="button" class="btn btn-quiet" id="back-objects">Назад</button><button class="btn btn-primary">Сохранить</button></div></form>`);
  $("#back-objects").onclick = openObjects;
  const form = $("#facility-form");
  form.elements.color.disabled = form.elements.autoColor.checked;
  form.elements.autoColor.onchange = () => { form.elements.color.disabled = form.elements.autoColor.checked; };
  form.elements.name.oninput = () => { if (form.elements.autoColor.checked) form.elements.color.value = colorFor(normalize(form.elements.name.value)); };
  form.onsubmit = async (event) => {
    event.preventDefault();
    const data = { name: normalize(form.elements.name.value), city: normalize(form.elements.city.value), color: form.elements.autoColor.checked ? "" : form.elements.color.value };
    try { await commitDb((db) => facility ? BirthdayDB.updateFacility(db, facility.id, data).db : BirthdayDB.createFacility(db, data).db, "Объект сохранён"); openObjects(); } catch { /* Keep form open. */ }
  };
}
function giftHistory(employee) {
  if (!employee) return `<section class="span-2"><h3>История подарков</h3><p class="hint">Сохраните сотрудника, чтобы добавлять подарки.</p></section>`;
  const gifts = state.db.gifts.filter((gift) => gift.employee_id === employee.id).sort((a, b) => Number(b.year) - Number(a.year));
  return `<section class="span-2"><div class="section-head"><h3>История подарков</h3><button type="button" class="btn btn-quiet" id="add-gift">＋ Добавить</button></div><div class="entity-list">${gifts.length ? gifts.map((gift) => `<div class="entity-row"><div><strong>${esc(gift.year)} · ${esc(gift.gift || "Без названия")}</strong><small>${gift.gift_cost === "" || gift.gift_cost == null ? "Стоимость не указана" : `${esc(gift.gift_cost)} ₽`}</small></div><button type="button" class="btn btn-quiet" data-gift-edit="${esc(gift.id)}">Изменить</button><button type="button" class="btn btn-danger" data-gift-delete="${esc(gift.id)}">Удалить</button></div>`).join("") : `<p class="hint">Подарков пока нет.</p>`}</div></section>`;
}
function openEmployee(id) {
  if (!state.rawBytes || !state.writable) return;
  const employee = state.db.employees.find((row) => row.id === id);
  if (!employee && !state.db.facilities.length) return showError("Сначала добавьте объект, затем сотрудника.");
  const selected = employee?.facility_id || state.db.facilities[0]?.id;
  modal(`<form class="modal" id="employee-form"><div class="modal-head"><div><p class="eyebrow">${employee ? "РЕДАКТИРОВАНИЕ" : "НОВАЯ ЗАПИСЬ"}</p><h2>${employee ? esc(fullName(employee)) : "Добавить сотрудника"}</h2></div><button type="button" class="icon-btn" data-close>×</button></div><label class="field"><span>Фамилия *</span><input name="last_name" required value="${esc(employee?.last_name || "")}"></label><label class="field"><span>Имя *</span><input name="first_name" required value="${esc(employee?.first_name || "")}"></label><label class="field"><span>Отчество</span><input name="middle_name" value="${esc(employee?.middle_name || "")}"></label><label class="field"><span>Дата рождения *</span><input name="birth_date" type="date" required value="${esc(employee?.birth_date || "")}"></label><label class="field"><span>Должность</span><input name="position" value="${esc(employee?.position || "")}"></label><label class="field"><span>Объект *</span><select name="facility_id" required ${state.db.facilities.length ? "" : "disabled"}>${state.db.facilities.map((facility) => `<option value="${esc(facility.id)}" ${selected === facility.id ? "selected" : ""}>${esc(facility.name)}</option>`).join("")}</select><small id="facility-city" class="hint"></small></label><label class="field span-2"><span>Заметка</span><textarea name="note" rows="3">${esc(employee?.note || "")}</textarea></label>${giftHistory(employee)}<div class="modal-actions span-2">${employee ? `<button type="button" class="btn btn-danger" id="delete-employee">Удалить</button>` : "<span></span>"}<div><button type="button" class="btn btn-quiet" data-close>Отмена</button><button class="btn btn-primary">Сохранить в Excel</button></div></div></form>`);
  const form = $("#employee-form");
  const updateCity = () => { $("#facility-city").textContent = state.db.facilities.length ? `Город: ${state.db.facilities.find((row) => row.id === form.elements.facility_id.value)?.city || "—"}` : "Сначала добавьте объект"; };
  form.elements.facility_id.onchange = updateCity;
  updateCity();
  form.onsubmit = async (event) => {
    event.preventDefault();
    const data = Object.fromEntries(["last_name", "first_name", "middle_name", "position", "note"].map((field) => [field, normalize(form.elements[field].value)]));
    data.birth_date = form.elements.birth_date.value;
    data.facility_id = form.elements.facility_id.value;
    try { await commitDb((db) => employee ? BirthdayDB.updateEmployee(db, employee.id, data).db : BirthdayDB.createEmployee(db, data).db, "Сотрудник сохранён"); closeModal(); } catch { /* Keep form open. */ }
  };
  if (employee) {
    $("#delete-employee").onclick = async () => {
      const count = state.db.gifts.filter((gift) => gift.employee_id === employee.id).length;
      if (!confirm(count ? `У сотрудника есть ${count} записей истории подарков, они тоже будут удалены. Продолжить?` : `Удалить сотрудника «${fullName(employee)}»?`)) return;
      try { await commitDb((db) => BirthdayDB.deleteEmployee(db, employee.id, { cascade: count > 0 }), "Сотрудник удалён"); closeModal(); } catch { /* Keep form open. */ }
    };
    $("#add-gift").onclick = () => openGift(employee.id);
    $("#modal-root").querySelectorAll("[data-gift-edit]").forEach((button) => button.onclick = () => openGift(employee.id, button.dataset.giftEdit));
    $("#modal-root").querySelectorAll("[data-gift-delete]").forEach((button) => button.onclick = async () => {
      if (!confirm("Удалить запись о подарке?")) return;
      try { await commitDb((db) => BirthdayDB.deleteGift(db, button.dataset.giftDelete), "Подарок удалён"); openEmployee(employee.id); } catch { /* Keep form open. */ }
    });
  }
}
function openGift(employeeId, giftId) {
  const gift = state.db.gifts.find((row) => row.id === giftId);
  modal(`<form class="modal" id="gift-form"><div class="modal-head"><h2>${gift ? "Изменить подарок" : "Добавить подарок"}</h2><button type="button" class="icon-btn" data-close>×</button></div><label class="field"><span>Год *</span><input name="year" type="number" step="1" required value="${esc(gift?.year ?? new Date().getFullYear())}"></label><label class="field"><span>Стоимость (₽, необязательно)</span><input name="gift_cost" type="number" min="0" step="any" value="${esc(gift?.gift_cost ?? "")}"></label><label class="field span-2"><span>Подарок</span><input name="gift" value="${esc(gift?.gift || "")}"></label><div class="modal-actions span-2"><button type="button" class="btn btn-quiet" id="back-employee">Назад</button><button class="btn btn-primary">Сохранить</button></div></form>`);
  $("#back-employee").onclick = () => openEmployee(employeeId);
  const form = $("#gift-form");
  form.onsubmit = async (event) => {
    event.preventDefault();
    const data = { employee_id: employeeId, year: Number(form.elements.year.value), gift: normalize(form.elements.gift.value), gift_cost: form.elements.gift_cost.value === "" ? "" : Number(form.elements.gift_cost.value) };
    try { await commitDb((db) => gift ? BirthdayDB.updateGift(db, gift.id, data).db : BirthdayDB.createGift(db, data).db, "Подарок сохранён"); openEmployee(employeeId); } catch { /* Keep form open. */ }
  };
}

$("#select-db-btn").onclick = selectDatabase;
$("#create-db-btn").onclick = createDatabase;
$("#objects-btn").onclick = openObjects;
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
  try { await state.saveQueue.catch(() => {}); await loadWorkbook(new Uint8Array(await file.arrayBuffer()), file.name); }
  catch (e) { showError(errorText(e)); }
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
