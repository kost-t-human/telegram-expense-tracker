/**
 * Telegram Mini App for tracking personal inflows and outflows.
 *
 * Records are stored in a Google Spreadsheet through a Google Apps Script web
 * app (see apps-script/Code.gs). Everything else — categories, settings, UI
 * state — lives in localStorage, namespaced by APP_ID so that several bot
 * deployments can share one browser origin, and is mirrored into Telegram's
 * per-user CloudStorage so the settings follow the user to another device.
 *
 * GAS_URL and APP_ID come from config.js (see config.js.example).
 */

const DEF_OUTFLOW_CATS = ['Other'];
const DEF_INFLOW_CATS  = ['Other'];

// Thousands are space-separated regardless of UI language, matching how the
// amount field formats what the user types.
const MONEY_LOCALE = 'ru-RU';

const $ = id => document.getElementById(id);

// ══════════════════════════════════════════════════════════════════
//  Persistence
// ══════════════════════════════════════════════════════════════════

const lsGet = (k, d = '') => localStorage.getItem(APP_ID + ':' + k) ?? d;
const lsSet = (k, v) => { localStorage.setItem(APP_ID + ':' + k, v); cloudPut(k, v); };
const save = (key, val) => lsSet(key, JSON.stringify(val));
const load = (key, def) => {
  try { const v = JSON.parse(lsGet(key, 'null')); return v !== null ? v : def; }
  catch { return def; }
};

/** The APP_ID-namespaced keys currently in localStorage, prefix stripped. */
const localKeys = () =>
  Array.from({ length: localStorage.length }, (_, i) => localStorage.key(i))
    .filter(k => k && k.startsWith(APP_ID + ':'))
    .map(k => k.slice(APP_ID.length + 1));

// ══════════════════════════════════════════════════════════════════
//  Cloud mirror
// ══════════════════════════════════════════════════════════════════

// Telegram keeps a small key/value store per bot and user, synced across that
// user's devices (Bot API 6.9+). Mirroring the settings there means a second
// device starts out already pointing at the right spreadsheet. The store is
// scoped to the bot, so keys go in under their bare name — ':' is not a legal
// cloud key character anyway. Outside Telegram, or on an older client, `cloud`
// stays null and everything below is a no-op: localStorage alone, as before.
const cloud = window.Telegram?.WebApp?.isVersionAtLeast?.('6.9')
  ? window.Telegram.WebApp.CloudStorage
  : null;

const CLOUD_MAX_BYTES = 4096;    // Telegram refuses anything longer
const CLOUD_RETRY_MS  = 5000;
const CLOUD_RETRY_MAX = 60000;

const cloudPending = new Set();  // keys whose last cloud write failed
let cloudRetryDelay = CLOUD_RETRY_MS;
let cloudRetryTimer = null;

/** Mirror one key upwards. Failures are queued, never surfaced to the user. */
function cloudPut(key, value) {
  if (!cloud) return;
  if (new TextEncoder().encode(value).length > CLOUD_MAX_BYTES) {
    // Too big to store. Drop whatever the cloud still holds under that key,
    // or another device would later restore the stale copy over its own.
    cloud.removeItem(key, () => {});
    cloudPending.delete(key);
    return;
  }
  cloud.setItem(key, value, err => {
    if (!err) { cloudPending.delete(key); cloudRetryDelay = CLOUD_RETRY_MS; return; }
    cloudPending.add(key);
    cloudRetryTimer ??= setTimeout(cloudFlush, cloudRetryDelay);
  });
}

/** Retry the queue, backing off while the writes keep failing. */
function cloudFlush() {
  cloudRetryTimer = null;
  cloudRetryDelay = Math.min(cloudRetryDelay * 2, CLOUD_RETRY_MAX);
  for (const key of [...cloudPending]) {
    cloudPending.delete(key);
    cloudPut(key, lsGet(key));   // the current value, not the one that failed
  }
}

/**
 * Reconcile localStorage with the cloud, once, before the state is read.
 * Keys only the cloud has win — that is the copy the user's other devices
 * already agreed on. Keys only this device has are pushed up, which is also
 * what migrates an install that predates the mirror.
 */
function cloudSync() {
  return new Promise(resolve => {
    if (!cloud) return resolve();
    cloud.getKeys((err, keys = []) => {
      if (err) return resolve();
      for (const key of localKeys()) {
        if (!keys.includes(key)) cloudPut(key, lsGet(key));
      }
      if (!keys.length) return resolve();
      cloud.getItems(keys, (err2, items) => {
        if (!err2) {
          for (const [key, value] of Object.entries(items || {})) {
            if (value !== '') localStorage.setItem(APP_ID + ':' + key, value);
          }
        }
        resolve();
      });
    });
  });
}

// ══════════════════════════════════════════════════════════════════
//  State
// ══════════════════════════════════════════════════════════════════

/** Everything persisted, read back out of localStorage. */
function storedState() {
  return {
    spreadsheetId: lsGet('spreadsheetId'),
    userName:      lsGet('userName'),
    expenseType:   lsGet('expenseType') || 'outflow',
    lang:          lsGet('lang'),
    summaryMode:   lsGet('summaryMode') || 'month',
    summaryView:   lsGet('summaryView') || 'category',
    historyType:   lsGet('historyType') || 'outflow',

    catCache: load('catCache', {}),                   // {outflow:[], inflow:[]}
    hidden:   load('hidden', { outflow: [], inflow: [] }),
    freq:     load('freq', {}),                       // {type: {category: uses}}
    accounts:       load('accounts', []),
    hiddenAccounts: load('hiddenAccounts', []),
    txnTypes:       load('txnTypes', []),
    hiddenTxnTypes: load('hiddenTxnTypes', []),
    subcats:        load('subcats', {}),              // {category: [subcategory]}
    hiddenSubcats:  load('hiddenSubcats', {}),

    useSubcats:  load('useSubcats', true),
    useAccounts: load('useAccounts', true),
    useTxnTypes: load('useTxnTypes', false),
    minusSign:   load('minusSign', false),
    setupDone:   load('setupDone', false),
  };
}

const S = {
  summaryPeriod: '',
  settingsCatType: 'outflow',
  historyRecords: [],   // the page(s) currently listed, newest entry first
  historyHasMore: false,
  recordPicked:  null,  // the record the actions modal is open for
  editing:       null,  // {row, sig} while the form edits an existing record
  tgUser: null,
  ...storedState(),
};

// v1 stored 'expense'/'income' where v2 uses 'outflow'/'inflow', and kept
// subcategories in a flat array instead of one bucket per category. Runs after
// the cloud sync, so that a v1 copy pulled from another device is fixed too.
function migrateV1() {
  let dirty = false;
  for (const bucket of [S.catCache, S.hidden, S.freq]) {
    if ('expense' in bucket) { bucket.outflow ??= bucket.expense; delete bucket.expense; dirty = true; }
    if ('income'  in bucket) { bucket.inflow  ??= bucket.income;  delete bucket.income;  dirty = true; }
  }
  if (dirty) { save('catCache', S.catCache); save('hidden', S.hidden); save('freq', S.freq); }
  if (S.expenseType === 'expense') { S.expenseType = 'outflow'; lsSet('expenseType', 'outflow'); }
  if (S.expenseType === 'income')  { S.expenseType = 'inflow';  lsSet('expenseType', 'inflow'); }
  if (Array.isArray(S.subcats))       { S.subcats = {};       save('subcats', S.subcats); }
  if (Array.isArray(S.hiddenSubcats)) { S.hiddenSubcats = {}; save('hiddenSubcats', S.hiddenSubcats); }
}

// ══════════════════════════════════════════════════════════════════
//  Localisation
// ══════════════════════════════════════════════════════════════════

let DICT = {};

/** t('key') → phrase; t('key', {name}) also fills `{name}` placeholders. */
function t(key, vars) {
  let phrase = DICT[key] ?? key;
  if (vars) for (const [k, v] of Object.entries(vars)) phrase = phrase.replaceAll(`{${k}}`, v);
  return phrase;
}

async function fetchDictionary(lang) {
  if (!lang) return;
  try {
    DICT = await (await fetch(`lang/${lang}.json`)).json();
  } catch (e) {
    console.error('Failed to load language', lang, e);
  }
}

function applyLang() {
  document.documentElement.lang = S.lang || 'en';
  const fill = (attr, apply) => document.querySelectorAll(`[${attr}]`).forEach(el => {
    apply(el, t(el.getAttribute(attr)));
  });
  fill('data-i18n',             (el, v) => el.textContent = v);
  fill('data-i18n-html',        (el, v) => el.innerHTML = v);
  fill('data-i18n-placeholder', (el, v) => el.placeholder = v);

  updateLangButtons();
  setType(S.expenseType);
  updateSubmitLabel();
  if ($('tab-settings').classList.contains('active')) renderCatList();
  if ($('tab-summary').classList.contains('active') && S.summaryPeriod) loadSummary();
  if ($('tab-history').classList.contains('active')) renderHistory();
}

function updateLangButtons() {
  for (const code of ['ru', 'en', 'vi']) {
    const btn = $('lang' + code[0].toUpperCase() + code.slice(1));
    if (btn) btn.className = 'lang-btn' + (S.lang === code ? ' active' : '');
  }
}

async function setLang(lang) {
  S.lang = lang;
  lsSet('lang', lang);
  $('lang-picker').style.display = 'none';
  await fetchDictionary(lang);
  applyLang();
  if (!S.setupDone) promptUserName();
}

// ══════════════════════════════════════════════════════════════════
//  Helpers
// ══════════════════════════════════════════════════════════════════

const esc = s => String(s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;')
  .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const formatMoney = n => Number(n).toLocaleString(MONEY_LOCALE);
const todayStr = () => new Date().toISOString().slice(0, 10);

/** Accepts either a spreadsheet URL or a bare ID. */
const extractSheetId = s => s.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/)?.[1] ?? s.trim();

/** Periods are 'YYYY' in year mode and 'YYYY-MM' in month mode. */
function currentPeriod(mode) {
  const now = new Date();
  return mode === 'year'
    ? String(now.getFullYear())
    : `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}
function shiftedPeriod(p, mode, step) {
  if (mode === 'year') return String(Number(p) + step);
  const [y, m] = p.split('-').map(Number);
  const d = new Date(y, m - 1 + step);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}
const prevPeriod = (p, mode) => shiftedPeriod(p, mode, -1);
const nextPeriod = (p, mode) => shiftedPeriod(p, mode, +1);

function periodLabel(p, mode) {
  if (mode === 'year') return t('year_format_val', { y: p });
  const [y, m] = p.split('-').map(Number);
  return t('month_format_val', { m: t('mon_' + m), y });
}

async function apiGet(params) {
  params._t = Date.now();
  const res = await fetch(GAS_URL + '?' + new URLSearchParams(params), { cache: 'no-store' });
  return res.json();
}

/**
 * Mirrors a local change to the spreadsheet. Returns false (after showing the
 * reason) when the server refused it, true when it went through or when no
 * spreadsheet is connected yet.
 */
async function pushToSheet(params, dupKey = 'item_exists') {
  if (!S.spreadsheetId) return true;
  try {
    const res = await apiGet({ ...params, spreadsheetId: S.spreadsheetId });
    if (res.status === 'duplicate') {
      showToast(t(dupKey, { name: res.existing }), 'warning');
      return false;
    }
    if (res.status !== 'ok') throw new Error(res.message);
    return true;
  } catch (e) {
    showToast(t('error_prefix') + e.message, 'error');
    return false;
  }
}

/**
 * Reconciles a locally ordered list with the spreadsheet: keeps the user's
 * order, drops what the sheet no longer has and appends new names the user has
 * not hidden. Names are compared case-insensitively.
 */
function mergeRemote(local, remote, hidden = []) {
  const has = (list, v) => list.some(x => x.toLowerCase() === v.toLowerCase());
  return [
    ...local.filter(v => has(remote, v)),
    ...remote.filter(v => !has(local, v) && !has(hidden, v)),
  ];
}

// ══════════════════════════════════════════════════════════════════
//  UI primitives
// ══════════════════════════════════════════════════════════════════

let toastTimer;
function showToast(msg, type = 'success') {
  const box = $('toast-box');
  box.textContent = msg;
  box.className = type === 'error' ? 't-err' : type === 'warning' ? 't-warn' : '';
  box.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => box.classList.remove('show'), 2800);
}

/** Repopulates a <select>, keeping the current choice when it still exists. */
function fillSelect(el, items, placeholder) {
  if (!el) return;
  const prev = el.value;
  el.innerHTML = (placeholder != null ? `<option value="">${esc(placeholder)}</option>` : '')
    + items.map(v => `<option value="${esc(v)}">${esc(v)}</option>`).join('');
  if (items.includes(prev)) el.value = prev;
}

/**
 * Selects a value in a <select>, adding the option first when the list has no
 * such entry. Assigning an unknown value silently keeps the previous choice,
 * which would quietly rewrite the field of a record being edited.
 */
function selectValue(el, value) {
  if (!el) return;
  const v = value || '';
  if (v && ![...el.options].some(o => o.value === v)) el.add(new Option(v, v));
  el.value = v;
}

/** Renders a reorderable list of names with rename/hide actions. */
function renderDraggableList(el, items, { onEdit, onHide, onReorder }) {
  el.innerHTML = '';
  items.forEach((item, i) => {
    const row = document.createElement('div');
    row.className = 'cat-item';
    row.draggable = true;
    row.dataset.index = i;
    row.innerHTML =
      `<span class="drag-handle"><i class="bi bi-grip-vertical"></i></span>`
      + `<span class="cat-label">${esc(item)}</span>`
      + `<button class="cat-action" title="${esc(t('rename_title'))}"><i class="bi bi-pencil"></i></button>`
      + `<button class="cat-action" title="${esc(t('hide_title'))}"><i class="bi bi-eye-slash"></i></button>`;

    const [btnEdit, btnHide] = row.querySelectorAll('button');
    btnEdit.addEventListener('click', () => onEdit(item));
    btnHide.addEventListener('click', () => onHide(item));

    const clearHighlights = () => el.querySelectorAll('.cat-item').forEach(r => r.style.background = '');
    row.addEventListener('dragstart', () => { row.dataset.dragging = '1'; row.style.opacity = '.4'; });
    row.addEventListener('dragend', () => { delete row.dataset.dragging; row.style.opacity = ''; clearHighlights(); });
    row.addEventListener('dragover', e => { e.preventDefault(); if (!row.dataset.dragging) row.style.background = 'rgba(0,122,255,.06)'; });
    row.addEventListener('dragleave', () => row.style.background = '');
    row.addEventListener('drop', e => {
      e.preventDefault();
      row.style.background = '';
      const from = el.querySelector('[data-dragging]');
      if (!from || from === row) return;
      onReorder(Number(from.dataset.index), Number(row.dataset.index));
    });
    el.appendChild(row);
  });
}

function showTab(tab, pushState = true) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  $('tab-' + tab).classList.add('active');
  $('nav-' + tab).classList.add('active');

  if (tab === 'summary') renderSummary();
  if (tab === 'history') renderHistory();
  if (tab === 'settings') {
    renderCatList();
    renderListSettings('accounts');
    renderListSettings('txnTypes');
    renderSubcatCatSelect();
  }
  if (pushState) history.pushState({ tab }, '', '#' + tab);

  const backBtn = window.Telegram?.WebApp?.BackButton;
  if (backBtn) tab === 'main' ? backBtn.hide() : backBtn.show();
}

// ══════════════════════════════════════════════════════════════════
//  Categories
// ══════════════════════════════════════════════════════════════════

function activeCategories(type) {
  const all = S.catCache[type] || [...(type === 'outflow' ? DEF_OUTFLOW_CATS : DEF_INFLOW_CATS)];
  const hidden = S.hidden[type] || [];
  return all.filter(c => !hidden.includes(c));
}

/** The three most used categories of the current type, for the quick pills. */
function topFreqCats(type) {
  const active = activeCategories(type);
  return Object.entries(S.freq[type] || {})
    .sort((a, b) => b[1] - a[1])
    .map(([cat]) => cat)
    .filter(cat => active.includes(cat))
    .slice(0, 3);
}

async function syncCategories(type) {
  if (!S.spreadsheetId) return;
  try {
    const res = await apiGet({ action: 'getCategories', spreadsheetId: S.spreadsheetId, type });
    if (res.status !== 'ok') return;
    const merged = mergeRemote(S.catCache[type] || [], res.categories, S.hidden[type] || []);
    if (JSON.stringify(merged) === JSON.stringify(S.catCache[type])) return;
    S.catCache[type] = merged;
    save('catCache', S.catCache);
    refreshCategories();
  } catch { /* offline — keep the cached list */ }
}

function renderQuickCats() {
  const top = topFreqCats(S.expenseType);
  const wrap = $('quickCats');
  const shown = [...wrap.querySelectorAll('.qpill')].map(b => b.textContent);
  if (JSON.stringify(top) === JSON.stringify(shown)) return;

  wrap.innerHTML = '';
  wrap.style.display = top.length ? 'flex' : 'none';
  const color = S.expenseType === 'outflow' ? 'var(--out)' : 'var(--in)';
  for (const cat of top) {
    const btn = document.createElement('button');
    btn.className = 'qpill';
    btn.textContent = cat;
    btn.style.color = btn.style.borderColor = color;
    btn.addEventListener('click', () => { $('fCategory').value = cat; renderSubcatSelect(); });
    wrap.appendChild(btn);
  }
}

function renderCategorySelect() {
  fillSelect($('fCategory'), activeCategories(S.expenseType));
  renderSubcatSelect();
}

/** Everything that has to follow a change to the category list. */
function refreshCategories() {
  renderCatList();
  renderCategorySelect();
  renderQuickCats();
  renderSubcatCatSelect();
}

function renderCatList() {
  const type = S.settingsCatType;
  renderDraggableList($('catList'), activeCategories(type), {
    onEdit: cat => renameCategory(type, cat),
    onHide: cat => hideCategory(type, cat),
    onReorder: (from, to) => reorderCats(type, from, to),
  });
}

function reorderCats(type, fromI, toI) {
  const active = activeCategories(type);
  active.splice(toI, 0, active.splice(fromI, 1)[0]);
  const hidden = (S.catCache[type] || []).filter(c => (S.hidden[type] || []).includes(c));
  S.catCache[type] = [...active, ...hidden];
  save('catCache', S.catCache);
  refreshCategories();
}

function hideCategory(type, cat) {
  S.hidden[type] ??= [];
  if (!S.hidden[type].includes(cat)) S.hidden[type].push(cat);
  save('hidden', S.hidden);
  refreshCategories();
}

function restoreHidden() {
  S.hidden = { outflow: [], inflow: [] };
  save('hidden', S.hidden);
  refreshCategories();
  showToast(t('all_restored'));
}

async function addCategory() {
  const input = $('sNewCat');
  const name = input.value.trim();
  const type = S.settingsCatType;
  if (!name) return;

  const dup = (S.catCache[type] || []).find(c => c.toLowerCase() === name.toLowerCase());
  if (dup) { showToast(t('cat_exists', { name: dup }), 'warning'); return; }

  input.disabled = true;
  const ok = await pushToSheet({ action: 'addCategory', name, type }, 'cat_exists');
  input.disabled = false;
  if (!ok) return;

  S.catCache[type] = [...(S.catCache[type] || []), name];
  save('catCache', S.catCache);
  input.value = '';
  refreshCategories();
  showToast(t('cat_added'));
}

async function renameCategory(type, oldName) {
  const newName = prompt(t('rename_prompt'), oldName)?.trim();
  if (!newName || newName === oldName) return;

  const existing = S.catCache[type] || [];
  const dup = existing.find(c => c.toLowerCase() === newName.toLowerCase() && c !== oldName);
  if (dup) { showToast(t('cat_exists', { name: dup }), 'warning'); return; }
  if (!await pushToSheet({ action: 'renameCategory', oldName, newName, type }, 'cat_exists')) return;

  S.catCache[type] = existing.map(c => c === oldName ? newName : c);
  save('catCache', S.catCache);
  if (S.freq[type]?.[oldName]) {
    S.freq[type][newName] = (S.freq[type][newName] || 0) + S.freq[type][oldName];
    delete S.freq[type][oldName];
    save('freq', S.freq);
  }
  refreshCategories();
}

// ══════════════════════════════════════════════════════════════════
//  Editable lists: accounts, transaction types, subcategories
//
//  All three behave identically in the settings tab; they differ only in where
//  they are stored and how they reach the spreadsheet, which the descriptors
//  below capture. Subcategories are nested per category, hence the `category`
//  hook and the extra argument on every accessor.
// ══════════════════════════════════════════════════════════════════

const LISTS = {
  accounts: {
    api: 'accounts',
    listEl: 'accountSettingsList', input: 'sNewAccount', field: 'fAccount',
    placeholder: 'no_account', msgAdded: 'account_added', msgRestored: 'accounts_restored',
    all:       () => S.accounts,
    setAll:    v  => { S.accounts = v; save('accounts', v); },
    hidden:    () => S.hiddenAccounts,
    setHidden: v  => { S.hiddenAccounts = v; save('hiddenAccounts', v); },
  },
  txnTypes: {
    api: 'txnTypes',
    listEl: 'txnTypeSettingsList', input: 'sNewTxnType', field: 'fTxnType',
    placeholder: 'no_txntype', msgAdded: 'txntype_added', msgRestored: 'txntypes_restored',
    all:       () => S.txnTypes,
    setAll:    v  => { S.txnTypes = v; save('txnTypes', v); },
    hidden:    () => S.hiddenTxnTypes,
    setHidden: v  => { S.hiddenTxnTypes = v; save('hiddenTxnTypes', v); },
  },
  subcats: {
    api: 'subcats',
    listEl: 'subcatSettingsList', input: 'sNewSubcat', field: 'fSubcategory',
    placeholder: 'no_subcat', msgAdded: 'subcat_added', msgRestored: 'subcats_restored',
    category:  () => $('sSubcatCat')?.value || '',
    all:       cat => S.subcats[cat] || [],
    setAll:    (v, cat) => { S.subcats[cat] = v; save('subcats', S.subcats); },
    hidden:    cat => S.hiddenSubcats[cat] || [],
    // Without a category there is nothing to scope to, so restore everything.
    setHidden: (v, cat) => {
      if (cat) S.hiddenSubcats[cat] = v; else S.hiddenSubcats = {};
      save('hiddenSubcats', S.hiddenSubcats);
    },
  },
};

/** Category the list is currently scoped to; '' for the flat lists. */
const listCat = kind => LISTS[kind].category?.() ?? '';

function activeList(kind, cat = listCat(kind)) {
  const list = LISTS[kind];
  if (list.category && !cat) return [];
  const hidden = list.hidden(cat);
  return list.all(cat).filter(v => !hidden.includes(v));
}

const activeSubcats = cat => activeList('subcats', cat);

const renderAccountSelect = () => fillSelect($('fAccount'), activeList('accounts'), t('no_account'));
const renderTxnTypeSelect = () => fillSelect($('fTxnType'), activeList('txnTypes'), t('no_txntype'));

// The record form picks subcategories of the category being recorded, which is
// not the one selected in settings — hence its own lookup.
function renderSubcatSelect() {
  const cat = $('fCategory').value;
  fillSelect($('fSubcategory'), cat ? activeSubcats(cat) : [], t('no_subcat'));
}

function renderListSettings(kind) {
  const list = LISTS[kind];
  const el = $(list.listEl);
  if (!el) return;
  if (list.category && !listCat(kind)) {
    el.innerHTML = `<div class="cat-item" style="color:var(--hint);font-size:13px;justify-content:center">${esc(t('select_cat_for_subcats'))}</div>`;
    return;
  }
  renderDraggableList(el, activeList(kind), {
    onEdit: name => renameListItem(kind, name),
    onHide: name => hideListItem(kind, name),
    onReorder: (from, to) => reorderList(kind, from, to),
  });
}

function refreshList(kind) {
  renderListSettings(kind);
  fillSelect($(LISTS[kind].field), activeList(kind), t(LISTS[kind].placeholder));
  if (kind === 'subcats') renderSubcatSelect();
}

function renderSubcatCatSelect() {
  const sel = $('sSubcatCat');
  if (!sel) return;
  fillSelect(sel, activeCategories(S.settingsCatType), t('select_cat_for_subcats'));
  renderListSettings('subcats');
}

function reorderList(kind, fromI, toI) {
  const list = LISTS[kind];
  const cat = listCat(kind);
  const active = activeList(kind, cat);
  active.splice(toI, 0, active.splice(fromI, 1)[0]);
  const hidden = list.hidden(cat);
  list.setAll([...active, ...list.all(cat).filter(v => hidden.includes(v))], cat);
  refreshList(kind);
}

function hideListItem(kind, name) {
  const list = LISTS[kind];
  const cat = listCat(kind);
  const hidden = list.hidden(cat);
  if (!hidden.includes(name)) list.setHidden([...hidden, name], cat);
  refreshList(kind);
}

function restoreHiddenList(kind) {
  LISTS[kind].setHidden([], listCat(kind));
  refreshList(kind);
  showToast(t(LISTS[kind].msgRestored));
}
const restoreHiddenAccounts = () => restoreHiddenList('accounts');
const restoreHiddenTxnTypes = () => restoreHiddenList('txnTypes');
const restoreHiddenSubcats  = () => restoreHiddenList('subcats');

async function addListItem(kind) {
  const list = LISTS[kind];
  const input = $(list.input);
  const name = input.value.trim();
  if (!name) return;

  const cat = listCat(kind);
  if (list.category && !cat) { showToast(t('select_cat_for_subcats'), 'warning'); return; }

  const dup = list.all(cat).find(v => v.toLowerCase() === name.toLowerCase());
  if (dup) { showToast(t('item_exists', { name: dup }), 'warning'); return; }

  input.disabled = true;
  const ok = await pushToSheet({ action: 'addListItem', list: list.api, name, category: cat });
  input.disabled = false;
  if (!ok) return;

  list.setAll([...list.all(cat), name], cat);
  input.value = '';
  refreshList(kind);
  showToast(t(list.msgAdded));
}

async function renameListItem(kind, oldName) {
  const newName = prompt(t('rename_prompt'), oldName)?.trim();
  if (!newName || newName === oldName) return;

  const list = LISTS[kind];
  const cat = listCat(kind);
  const dup = list.all(cat).find(v => v.toLowerCase() === newName.toLowerCase() && v !== oldName);
  if (dup) { showToast(t('item_exists', { name: dup }), 'warning'); return; }
  if (!await pushToSheet({ action: 'renameListItem', list: list.api, oldName, newName, category: cat })) return;

  const rename = v => v === oldName ? newName : v;
  list.setAll(list.all(cat).map(rename), cat);
  const hidden = list.hidden(cat);
  if (hidden.includes(oldName)) list.setHidden(hidden.map(rename), cat);
  refreshList(kind);
}

async function syncFlatList(kind, action) {
  if (!S.spreadsheetId) return;
  const list = LISTS[kind];
  try {
    const res = await apiGet({ action, spreadsheetId: S.spreadsheetId });
    if (res.status !== 'ok') return;
    const merged = mergeRemote(list.all(), res.values, list.hidden());
    if (JSON.stringify(merged) === JSON.stringify(list.all())) return;
    list.setAll(merged);
    refreshList(kind);
  } catch { /* offline — keep the cached list */ }
}
const syncAccounts = () => syncFlatList('accounts', 'getAccounts');
const syncTxnTypes = () => syncFlatList('txnTypes', 'getTxnTypes');

async function syncSubcats() {
  if (!S.spreadsheetId) return;
  try {
    const res = await apiGet({ action: 'getSubcats', spreadsheetId: S.spreadsheetId });
    if (res.status !== 'ok') return;
    if (!res.subcats) return; // pre-2.0 Apps Script: no per-category data, don't wipe local

    const merged = {};
    for (const cat of new Set([...Object.keys(S.subcats), ...Object.keys(res.subcats)])) {
      const list = mergeRemote(S.subcats[cat] || [], res.subcats[cat] || [], S.hiddenSubcats[cat] || []);
      if (list.length) merged[cat] = list;
    }
    if (JSON.stringify(merged) === JSON.stringify(S.subcats)) return;
    S.subcats = merged;
    save('subcats', S.subcats);
    refreshList('subcats');
  } catch { /* offline — keep the cached lists */ }
}

const syncAll = () => Promise.all([
  syncCategories('outflow'), syncCategories('inflow'),
  syncAccounts(), syncTxnTypes(), syncSubcats(),
]);

// ══════════════════════════════════════════════════════════════════
//  Record form
// ══════════════════════════════════════════════════════════════════

function setType(type) {
  S.expenseType = type;
  lsSet('expenseType', type);
  const isOut = type === 'outflow';
  $('btnOut').className = 'seg-btn' + (isOut ? ' active-out' : '');
  $('btnIn').className  = 'seg-btn' + (isOut ? '' : ' active-in');
  $('submitBtn').className = 'submit-btn ' + (isOut ? 'out' : 'in');
  renderQuickCats();
  renderCategorySelect();
  if (!S.catCache[type] && S.spreadsheetId) syncCategories(type);
}

async function submitForm() {
  const date        = $('fDate').value;
  const category    = $('fCategory').value;
  const amount      = $('fAmount').value.replace(/\s/g, '').replace(',', '.');
  const subcategory = $('fSubcategory').value.trim();
  const account     = $('fAccount').value.trim();
  const txnType     = $('fTxnType').value.trim();
  const note        = $('fNote').value.trim();

  if (!date || !category || !amount) { showToast(t('fill_required'), 'warning'); return; }
  if (!S.spreadsheetId) { showToast(t('set_table_warn'), 'warning'); showTab('settings'); return; }

  const btn = $('submitBtn');
  const busy = on => { btn.disabled = on; btn.classList.toggle('loading', on); };
  busy(true);
  try {
    if (S.editing) {
      // Not optimistic, unlike a new record: an edit that silently failed would
      // leave the list showing something the spreadsheet does not hold.
      const res = await apiGet({
        action: 'updateRecord', spreadsheetId: S.spreadsheetId,
        row: S.editing.row, sig: S.editing.sig,
        date, category, subcategory, amount: signedAmount(amount),
        account, txnType, note, type: S.expenseType,
      });
      busy(false);
      if (res.status === 'stale') { staleRecord(); return; }
      if (res.status !== 'ok') throw new Error(res.message || 'Server error');
      cancelEdit();
      showToast(t('record_updated'));
      showTab('history');
      return;
    }

    // Same-day duplicates are the ones worth questioning; older dates are
    // usually deliberate backfilling.
    if (date === todayStr()) {
      const dup = await apiGet({
        action: 'checkDuplicate', spreadsheetId: S.spreadsheetId,
        date, amount, category, subcategory, note, account, txnType,
        userName: S.userName || '', type: S.expenseType,
      });
      if (dup.duplicate) {
        busy(false);
        try { await confirmDuplicate(); } catch { return; }
        busy(true);
      }
    }

    // Optimistic: the form clears right away and the write finishes in the
    // background, so recording several entries in a row stays fast.
    S.freq[S.expenseType] ??= {};
    S.freq[S.expenseType][category] = (S.freq[S.expenseType][category] || 0) + 1;
    save('freq', S.freq);
    for (const id of ['fAmount', 'fSubcategory', 'fAccount', 'fTxnType', 'fNote']) $(id).value = '';
    renderQuickCats();
    showToast(t(S.expenseType === 'outflow' ? 'expense_saved' : 'inflow_saved'));
    busy(false);

    saveRecord({ date, category, subcategory, amount, account, txnType, note })
      .catch(e => showToast(t('error_prefix') + e.message, 'error'));
  } catch (e) {
    showToast(t('error_prefix') + e.message, 'error');
    busy(false);
  }
}

/** Resolves when the user confirms the duplicate, rejects when they dismiss it. */
function confirmDuplicate() {
  return new Promise((resolve, reject) => {
    const el = $('dupModal');
    const modal = bootstrap.Modal.getOrCreateInstance(el);
    const btn = $('dupConfirmBtn');
    const cleanup = () => {
      btn.removeEventListener('click', onOk);
      el.removeEventListener('hidden.bs.modal', onHide);
      modal.hide();
    };
    const onOk = () => { cleanup(); resolve(); };
    const onHide = () => { cleanup(); reject(new Error('cancelled')); };
    btn.addEventListener('click', onOk, { once: true });
    el.addEventListener('hidden.bs.modal', onHide, { once: true });
    modal.show();
  });
}

/** Some spreadsheet templates expect outflows as negative numbers. */
const signedAmount = amount =>
  S.minusSign && S.expenseType === 'outflow' ? -Math.abs(amount) : amount;

async function saveRecord({ date, category, subcategory, amount, account, txnType, note }) {
  const res = await apiGet({
    action: 'addRecord', spreadsheetId: S.spreadsheetId,
    date, category, subcategory, amount: signedAmount(amount), account, txnType, note,
    type: S.expenseType,
    firstName: S.tgUser?.first_name || '',
    lastName:  S.tgUser?.last_name || '',
    username:  S.tgUser?.username || '',
    userName:  S.userName || '',
  });
  if (res.status !== 'ok') throw new Error(res.message || 'Server error');
}

// ══════════════════════════════════════════════════════════════════
//  Summary
// ══════════════════════════════════════════════════════════════════

function setSummaryMode(mode) {
  S.summaryMode = mode;
  lsSet('summaryMode', mode);
  S.summaryPeriod = currentPeriod(mode);
  updateModeButtons();
  loadSummary();
}
function setSummaryView(view) {
  S.summaryView = view;
  lsSet('summaryView', view);
  updateModeButtons();
  loadSummary();
}
function setSummaryQuick(which) {
  const cur = currentPeriod(S.summaryMode);
  S.summaryPeriod = which === 'prev' ? prevPeriod(cur, S.summaryMode) : cur;
  loadSummary();
}
function shiftPeriod(dir) {
  S.summaryPeriod = shiftedPeriod(S.summaryPeriod, S.summaryMode, dir < 0 ? -1 : 1);
  loadSummary();
}
function updateModeButtons() {
  const pill = (id, on) => $(id).className = 'mode-pill' + (on ? ' active' : '');
  pill('modeMonth', S.summaryMode === 'month');
  pill('modeYear',  S.summaryMode === 'year');
  for (const view of ['category', 'subcategory', 'account', 'trend']) {
    pill('view' + view[0].toUpperCase() + view.slice(1), S.summaryView === view);
  }
}

let pickerYear = new Date().getFullYear();

function openPeriodPicker() {
  pickerYear = Number(S.summaryPeriod.split('-')[0]) || new Date().getFullYear();
  renderPeriodPicker();
  bootstrap.Modal.getOrCreateInstance($('periodModal')).show();
}
function pickerShiftYear(dir) {
  pickerYear += dir;
  renderPeriodPicker();
}
function renderPeriodPicker() {
  const byMonth = S.summaryMode === 'month';
  $('pickerYearSelect').style.display = byMonth ? 'flex' : 'none';
  $('monthGrid').style.display = byMonth ? 'grid' : 'none';
  $('yearList').style.display = byMonth ? 'none' : 'flex';

  if (byMonth) {
    const [curY, curM] = S.summaryPeriod.split('-');
    $('pickerYearLabel').textContent = pickerYear;
    $('monthGrid').innerHTML = Array.from({ length: 12 }, (_, i) => {
      const active = Number(curY) === pickerYear && Number(curM) === i + 1 ? 'active' : '';
      return `<button class="month-btn ${active}" onclick="selectPickerMonth(${i + 1})">${esc(t('mon_' + (i + 1)))}</button>`;
    }).join('');
  } else {
    const startY = new Date().getFullYear() + 1;
    $('yearList').innerHTML = Array.from({ length: 10 }, (_, i) => startY - i).map(y => {
      const active = Number(S.summaryPeriod) === y ? 'active' : '';
      return `<button class="year-btn ${active}" onclick="selectPickerYear(${y})">${esc(t('year_format_val', { y }))}</button>`;
    }).join('');
  }
}
function selectPickerMonth(m) {
  selectPeriod(`${pickerYear}-${String(m).padStart(2, '0')}`);
}
function selectPickerYear(y) {
  selectPeriod(String(y));
}
function selectPeriod(period) {
  S.summaryPeriod = period;
  bootstrap.Modal.getInstance($('periodModal')).hide();
  loadSummary();
}

function renderSummary() {
  S.summaryPeriod ||= currentPeriod(S.summaryMode);
  updateModeButtons();
  loadSummary();
}

async function loadSummary() {
  $('periodLabel').textContent = periodLabel(S.summaryPeriod, S.summaryMode);
  const box = $('summaryContent');
  if (!S.spreadsheetId) {
    box.innerHTML = `<div class="text-center text-muted py-5">${esc(t('setup_table'))}</div>`;
    return;
  }
  box.innerHTML = '<div class="text-center py-5"><div class="spinner-border spinner-border-sm" style="color:var(--hint)"></div></div>';

  const trend = S.summaryView === 'trend';
  try {
    const res = trend
      ? await apiGet({ action: 'getTrend', spreadsheetId: S.spreadsheetId, endPeriod: S.summaryPeriod, count: 6, groupBy: S.summaryMode })
      : await apiGet({ action: 'getSummary', spreadsheetId: S.spreadsheetId, period: S.summaryPeriod, groupBy: S.summaryMode, viewBy: S.summaryView });
    if (res.status !== 'ok') throw new Error(res.message);
    trend ? renderTrend(res.trend) : renderSummaryData(res);
  } catch (e) {
    box.innerHTML = `<div class="text-center text-danger py-4"><i class="bi bi-exclamation-circle"></i> ${esc(e.message)}</div>`;
  }
}

function renderTrend(trend) {
  const box = $('summaryContent');
  if (!trend?.length) {
    box.innerHTML = `<div class="text-center text-muted py-5">${esc(t('no_records'))}</div>`;
    return;
  }
  const max = Math.max(...trend.map(p => Math.max(p.outflow, p.income)), 1);
  const cols = trend.map(p => {
    const label = S.summaryMode === 'year' ? p.period : p.period.slice(5);
    return `<div class="trend-col">
        <div class="trend-bars">
          <div class="trend-bar out" style="height:${Math.round(p.outflow / max * 100)}%"></div>
          <div class="trend-bar in" style="height:${Math.round(p.income / max * 100)}%"></div>
        </div>
        <div class="trend-lbl">${esc(label)}</div>
      </div>`;
  }).join('');

  const legend = (color, key) =>
    `<div class="trend-legend-item"><div class="trend-legend-dot" style="background:${color}"></div>${esc(t(key))}</div>`;
  box.innerHTML = `<div class="trend-wrap">
      <div class="trend-title">${esc(t('trend_view'))}</div>
      <div class="trend-chart">${cols}</div>
      <div class="trend-legend">${legend('var(--out)', 'expenses_title')}${legend('var(--in)', 'inflow_title')}</div>
    </div>`;
}

function renderSummaryData(data) {
  const outflows = data.outflows || data.expenses || [];
  const inflows = data.income || data.inflow || [];
  if (!outflows.length && !inflows.length) {
    $('summaryContent').innerHTML = `<div class="text-center text-muted py-5">${esc(t('no_records'))}</div>`;
    return;
  }

  const totOut = outflows.reduce((sum, i) => sum + Math.abs(i.total), 0);
  const totIn = inflows.reduce((sum, i) => sum + i.total, 0);
  const balance = totIn - totOut;
  const days = data.daysInPeriod || 30;
  const perDay = total => `~${formatMoney(Math.round(Math.abs(total) / days))}/${esc(t('day_short'))}`;

  const savRate = totIn > 0 ? Math.round(balance / totIn * 100) : null;
  const savColor = savRate === null ? 'var(--hint)' : savRate >= 0 ? 'var(--in)' : 'var(--out)';
  const balColor = balance >= 0 ? 'var(--in)' : 'var(--out)';
  const kpiCard = (labelKey, icon, color, value, sub = '') => `<div class="kpi-card">
      <div class="kpi-top"><span class="kpi-lbl">${esc(t(labelKey))}</span><i class="bi ${icon} kpi-icon" style="color:${color}"></i></div>
      <div class="kpi-val" style="color:${color}">${value}</div>
      ${sub ? `<div class="kpi-sub">${sub}</div>` : ''}
    </div>`;
  const kpi = `<div class="kpi-grid">
      ${kpiCard('inflow_title', 'bi-arrow-down-circle', 'var(--in)', formatMoney(totIn))}
      ${kpiCard('expenses_title', 'bi-arrow-up-circle', 'var(--out)', formatMoney(totOut))}
      ${kpiCard('balance', 'bi-wallet2', balColor, (balance >= 0 ? '+' : '') + formatMoney(balance))}
      ${kpiCard('savings_rate', 'bi-piggy-bank', savColor, savRate !== null ? savRate + '%' : '—',
        `${esc(t('daily_avg_exp'))}: ${totOut > 0 ? formatMoney(Math.round(totOut / days)) : 0}`)}
    </div>`;

  // Change against the previous period. Spending less is the good direction,
  // so the arrow direction, not the colour, follows the sign.
  const deltaChip = (cur, prev, labelKey) => {
    if (!prev) return '';
    const diff = Math.round((cur - prev) / prev * 100);
    const cls = diff < 0 ? 'up' : diff > 0 ? 'down' : 'neu';
    return `<span class="delta-chip ${cls}"><i class="bi bi-arrow-${diff < 0 ? 'down' : 'up'}-short"></i>${diff > 0 ? '+' : ''}${diff}% ${esc(t(labelKey))}</span>`;
  };
  const chips = deltaChip(totOut, data.prevOutflowTotal || 0, 'expenses_title')
    + deltaChip(totIn, data.prevIncomeTotal || 0, 'inflow_title');
  const vs = chips
    ? `<div class="vs-prev"><span style="font-size:11px;color:var(--hint);align-self:center">${esc(t('vs_prev'))}</span>${chips}</div>`
    : '';

  let insight = '';
  if (outflows.length) {
    const top = outflows[0];
    const pct = totOut > 0 ? Math.round(Math.abs(top.total) / totOut * 100) : 0;
    insight = `<div class="insight-card">
        <div class="insight-icon" style="background:var(--out-soft);color:var(--out)"><i class="bi bi-fire"></i></div>
        <div class="insight-body">
          <div class="insight-title">${esc(t('top_expense'))}</div>
          <div class="insight-text">${esc(top.label || 'Other')}</div>
          <div class="insight-meta">${pct}% ${esc(t('of_total'))} · ${perDay(top.total)}</div>
        </div>
        <div style="font-size:15px;font-weight:700;color:var(--out)">${formatMoney(Math.abs(top.total))}</div>
      </div>`;
  }

  const section = (items, total, color, titleKey, icon) => {
    if (!items.length) return '';
    const rows = items.map(item => {
      const pct = total > 0 ? Math.round(Math.abs(item.total) / Math.abs(total) * 100) : 0;
      const label = item.label || item.category || 'Other';
      const meta = item.count != null
        ? `${item.count} ${esc(t('records_count'))} · ${perDay(item.total)}`
        : perDay(item.total);
      return `<div class="sum-row">
          <div class="sum-row-info">
            <div class="sum-row-name" title="${esc(label)}">${esc(label)}</div>
            <div class="sum-bar"><div class="sum-bar-fill" style="width:${pct}%;background:${color}"></div></div>
            <div class="sum-row-meta">${meta}</div>
          </div>
          <div class="sum-row-right">
            <span class="sum-row-amt" style="color:${color}">${formatMoney(Math.abs(item.total))}</span>
            <span class="sum-row-pct">${pct}%</span>
          </div>
        </div>`;
    }).join('');
    return `<div class="sum-section">
        <div class="sum-hdr">
          <div class="sum-hdr-title" style="color:${color}"><i class="bi ${icon}"></i> ${esc(t(titleKey))}</div>
          <div class="sum-hdr-total" style="color:${color}">${formatMoney(Math.abs(total))}</div>
        </div>${rows}
      </div>`;
  };

  $('summaryContent').innerHTML = kpi + vs + insight
    + section(outflows, totOut, 'var(--out)', 'expenses_title', 'bi-arrow-up-circle-fill')
    + section(inflows, totIn, 'var(--in)', 'inflow_title', 'bi-arrow-down-circle-fill');
}

// ══════════════════════════════════════════════════════════════════
//  History
// ══════════════════════════════════════════════════════════════════

const HISTORY_PAGE = 30;

function setHistoryType(type) {
  S.historyType = type;
  lsSet('historyType', type);
  updateHistoryButtons();
  loadHistory();
}

function updateHistoryButtons() {
  const isOut = S.historyType === 'outflow';
  $('btnHistOut').className = 'seg-btn' + (isOut ? ' active-out' : '');
  $('btnHistIn').className  = 'seg-btn' + (isOut ? '' : ' active-in');
}

function renderHistory() {
  updateHistoryButtons();
  loadHistory();
}

const loadMoreHistory = () => loadHistory(true);

async function loadHistory(append = false) {
  const box  = $('historyContent');
  const more = $('historyMoreBtn');

  if (!S.spreadsheetId) {
    more.style.display = 'none';
    box.innerHTML = `<div class="text-center text-muted py-5">${esc(t('setup_table'))}</div>`;
    return;
  }
  if (!append) {
    S.historyRecords = [];
    more.style.display = 'none';
    box.innerHTML = '<div class="text-center py-5"><div class="spinner-border spinner-border-sm" style="color:var(--hint)"></div></div>';
  }
  more.classList.toggle('loading', append);
  more.disabled = append;

  try {
    const res = await apiGet({
      action: 'getRecords', spreadsheetId: S.spreadsheetId,
      type: S.historyType, limit: HISTORY_PAGE, offset: S.historyRecords.length,
    });
    if (res.status !== 'ok') throw new Error(res.message);
    S.historyRecords = S.historyRecords.concat(res.records || []);
    S.historyHasMore = !!res.hasMore;
    renderHistoryList();
  } catch (e) {
    if (append) showToast(t('error_prefix') + e.message, 'error');
    else box.innerHTML = `<div class="text-center text-danger py-4"><i class="bi bi-exclamation-circle"></i> ${esc(e.message)}</div>`;
  } finally {
    more.classList.remove('loading');
    more.disabled = false;
  }
}

/** 'Today' / 'Yesterday' / '24 August' for a YYYY-MM-DD purchase date. */
function dateHeading(iso) {
  if (!iso) return '—';
  if (iso === todayStr()) return t('today');
  if (iso === new Date(Date.now() - 864e5).toISOString().slice(0, 10)) return t('yesterday');

  const [y, m, d] = iso.split('-').map(Number);
  const opts = { day: 'numeric', month: 'long' };
  if (y !== new Date().getFullYear()) opts.year = 'numeric';
  // toLocaleDateString gets the case right where a month name list would not:
  // Russian needs "24 августа", not the nominative "24 Август".
  return new Date(y, m - 1, d).toLocaleDateString(S.lang || 'en', opts);
}

/**
 * The stamp under the amount: when the row was entered, which is not always the
 * day it is filed under. Apps Script deployments older than 2.4.0 send no
 * entryDate, and those fall back to the bare time they always showed.
 */
function entryStamp(r) {
  if (!r.time || !r.entryDate) return r.time || '';
  const [y, m, d] = r.entryDate.split('-').map(Number);
  const opts = { day: '2-digit', month: '2-digit' };
  if (y !== new Date().getFullYear()) opts.year = '2-digit';
  return new Date(y, m - 1, d).toLocaleDateString(S.lang || 'en', opts) + ' ' + r.time;
}

function renderHistoryList() {
  const box  = $('historyContent');
  const recs = S.historyRecords;
  $('historyMoreBtn').style.display = S.historyHasMore ? 'flex' : 'none';

  if (!recs.length) {
    box.innerHTML = `<div class="text-center text-muted py-5">${esc(t('no_records_yet'))}</div>`;
    return;
  }

  const color = S.historyType === 'outflow' ? 'var(--out)' : 'var(--in)';
  let html = '', day = null;
  recs.forEach((r, i) => {
    if (r.date !== day) {
      if (day !== null) html += '</div>';
      html += `<div class="sec-hdr${day === null ? '' : ' pt-3'}">${esc(dateHeading(r.date))}</div><div class="rec-group">`;
      day = r.date;
    }
    const meta = [r.subcategory, r.account, r.txnType, r.note].filter(Boolean).join(' · ');
    html += `<div class="rec-row" onclick="openRecord(${i})">
        <div class="rec-info">
          <div class="rec-name">${esc(r.category || '—')}</div>
          ${meta ? `<div class="rec-meta">${esc(meta)}</div>` : ''}
        </div>
        <div class="rec-right">
          <span class="rec-amt" style="color:${color}">${formatMoney(r.amount)}</span>
          ${r.time ? `<span class="rec-time">${esc(entryStamp(r))}</span>` : ''}
        </div>
      </div>`;
  });
  box.innerHTML = html + '</div>';
}

function openRecord(i) {
  const r = S.historyRecords[i];
  if (!r) return;
  S.recordPicked = r;

  const color = r.type === 'inflow' ? 'var(--in)' : 'var(--out)';
  const parts = [r.category, r.subcategory, r.account, r.txnType, r.note].filter(Boolean);
  $('recSummary').innerHTML =
    `<div class="rec-modal-amt" style="color:${color}">${formatMoney(r.amount)}</div>`
    + `<div class="rec-modal-when">${esc(dateHeading(r.date))}${r.time ? ' · ' + esc(r.time) : ''}</div>`
    + (parts.length ? `<div class="rec-modal-parts">${esc(parts.join(' · '))}</div>` : '');

  showRecordActions();
  bootstrap.Modal.getOrCreateInstance($('recModal')).show();
}

function showRecordActions() {
  $('recActions').style.display = 'flex';
  $('recConfirm').style.display = 'none';
}
function askDeleteRecord() {
  $('recActions').style.display = 'none';
  $('recConfirm').style.display = 'block';
}
function closeRecordModal() {
  bootstrap.Modal.getOrCreateInstance($('recModal')).hide();
  S.recordPicked = null;
}

/** The row moved or is gone: drop what we hold and re-read the list. */
function staleRecord() {
  showToast(t('record_stale'), 'warning');
  closeRecordModal();
  if (S.editing) cancelEdit();
  showTab('history');
}

async function confirmDeleteRecord() {
  const r = S.recordPicked;
  if (!r) return;
  const btn = $('recDeleteBtn');
  btn.disabled = true;
  btn.classList.add('loading');
  try {
    const res = await apiGet({
      action: 'deleteRecord', spreadsheetId: S.spreadsheetId, row: r.row, sig: r.sig,
    });
    if (res.status === 'stale') { staleRecord(); return; }
    if (res.status !== 'ok') throw new Error(res.message || 'Server error');
    closeRecordModal();
    showToast(t('record_deleted'));
    loadHistory();
  } catch (e) {
    showToast(t('error_prefix') + e.message, 'error');
  } finally {
    btn.disabled = false;
    btn.classList.remove('loading');
  }
}

/** Loads the picked record into the record form, which then updates instead of appending. */
function startEditRecord() {
  const r = S.recordPicked;
  if (!r) return;
  S.editing = { row: r.row, sig: r.sig };
  closeRecordModal();

  setType(r.type === 'inflow' ? 'inflow' : 'outflow');
  $('fDate').value = r.date || todayStr();
  selectValue($('fCategory'), r.category);
  renderSubcatSelect();
  selectValue($('fSubcategory'), r.subcategory);
  selectValue($('fAccount'), r.account);
  selectValue($('fTxnType'), r.txnType);
  $('fNote').value = r.note || '';
  // Through the input handler, so the amount is grouped like a typed one.
  $('fAmount').value = String(Math.abs(r.amount));
  $('fAmount').dispatchEvent(new Event('input'));

  $('editBanner').style.display = 'flex';
  updateSubmitLabel();
  showTab('main');
  window.scrollTo(0, 0);
}

function cancelEdit() {
  S.editing = null;
  $('editBanner').style.display = 'none';
  updateSubmitLabel();
  for (const id of ['fAmount', 'fSubcategory', 'fAccount', 'fTxnType', 'fNote']) $(id).value = '';
  $('fDate').value = todayStr();
}

function updateSubmitLabel() {
  $('submitBtn').querySelector('.btn-label').textContent = t(S.editing ? 'save_changes' : 'save');
}

// ══════════════════════════════════════════════════════════════════
//  Settings
// ══════════════════════════════════════════════════════════════════

function saveSheetId() {
  const raw = $('sSheetInput').value.trim();
  if (!raw) { showToast(t('enter_table'), 'warning'); return; }
  const id = extractSheetId(raw);
  S.spreadsheetId = id;
  lsSet('spreadsheetId', id);
  showSheetId(id);
  $('no-id-banner').style.display = 'none';

  // A different spreadsheet means different lists — start from scratch.
  S.catCache = {};       save('catCache', S.catCache);
  S.accounts = [];       save('accounts', S.accounts);
  S.hiddenAccounts = []; save('hiddenAccounts', S.hiddenAccounts);
  S.txnTypes = [];       save('txnTypes', S.txnTypes);
  S.hiddenTxnTypes = []; save('hiddenTxnTypes', S.hiddenTxnTypes);
  S.subcats = {};        save('subcats', S.subcats);
  S.hiddenSubcats = {};  save('hiddenSubcats', S.hiddenSubcats);

  syncAll()
    .then(() => { showToast(t('table_connected')); refreshCategories(); })
    .catch(() => showToast(t('sync_error'), 'error'));
}

function showSheetId(id) {
  const hint = $('sIdHint');
  hint.textContent = 'ID: ' + id;
  hint.style.display = 'block';
}

function saveUserName() {
  S.userName = $('sUserNameInput').value.trim();
  lsSet('userName', S.userName);
  showToast(t('user_name_saved'));
}

function setSettingsCatType(type) {
  S.settingsCatType = type;
  $('sCatOut').className = 'cat-seg-btn' + (type === 'outflow' ? ' active-out' : '');
  $('sCatIn').className  = 'cat-seg-btn' + (type === 'inflow' ? ' active-in' : '');
  refreshCategories();
}

function toggleFeature(key, val) {
  S[key] = val;
  save(key, val);
  applyFeatures();
}

/** Shows or hides the parts of the UI belonging to the optional features. */
function applyFeatures() {
  const show = (id, on, display = 'block') => {
    const el = $(id);
    if (el) el.style.display = on ? display : 'none';
  };
  show('settings-subcats-wrap',   S.useSubcats);
  show('settings-accounts-wrap',  S.useAccounts);
  show('settings-txntypes-wrap',  S.useTxnTypes);
  show('fSubcategoryWrap', S.useSubcats,  'flex');
  show('fAccountWrap',     S.useAccounts, 'flex');
  show('fTxnTypeWrap',     S.useTxnTypes, 'flex');
  show('viewSubcategory',  S.useSubcats);
  show('viewAccount',      S.useAccounts);

  const check = (id, on) => { const el = $(id); if (el) el.checked = on; };
  check('sUseSubcats',  S.useSubcats);
  check('sUseAccounts', S.useAccounts);
  check('sUseTxnTypes', S.useTxnTypes);
  check('sMinusSign',   S.minusSign);

  if (!S.useSubcats  && S.summaryView === 'subcategory') setSummaryView('category');
  if (!S.useAccounts && S.summaryView === 'account')     setSummaryView('category');
}

// ══════════════════════════════════════════════════════════════════
//  First run
// ══════════════════════════════════════════════════════════════════

function promptUserName() {
  const tgName = S.tgUser && ([S.tgUser.first_name, S.tgUser.last_name].filter(Boolean).join(' ') || S.tgUser.username);
  if (tgName) $('iUserNameInput').value = tgName;
  $('user-name-prompt').style.display = 'flex';
}

/** Onboarding: name → optional features → spreadsheet. */
function setupStep(step) {
  if (step === 2) {
    const name = $('iUserNameInput').value.trim();
    if (!name) return;
    S.userName = name;
    lsSet('userName', name);
    $('user-name-prompt').style.display = 'none';
    $('features-prompt').style.display = 'flex';
  } else if (step === 3) {
    for (const [key, id] of Object.entries({
      useSubcats: 'iUseSubcats', useAccounts: 'iUseAccounts',
      useTxnTypes: 'iUseTxnTypes', minusSign: 'iMinusSign',
    })) {
      S[key] = $(id).checked;
      save(key, S[key]);
    }
    $('features-prompt').style.display = 'none';
    $('sheet-id-prompt').style.display = 'flex';
  } else if (step === 4) {
    const raw = $('iSheetIdInput').value.trim();
    if (!raw) return;
    S.spreadsheetId = extractSheetId(raw);
    lsSet('spreadsheetId', S.spreadsheetId);
    S.setupDone = true;
    save('setupDone', true);
    $('sheet-id-prompt').style.display = 'none';
    startApp();
    showToast(t('table_connected'));
  }
}

// ══════════════════════════════════════════════════════════════════
//  Bootstrap
// ══════════════════════════════════════════════════════════════════

function startApp() {
  applyLang();
  $('fDate').value = todayStr();
  renderAccountSelect();
  renderTxnTypeSelect();
  renderSubcatSelect();
  setSettingsCatType(S.settingsCatType);
  applyFeatures();

  if (S.userName) $('sUserNameInput').value = S.userName;
  if (S.spreadsheetId) {
    $('sSheetInput').value = S.spreadsheetId;
    showSheetId(S.spreadsheetId);
    syncAll();
  } else {
    $('no-id-banner').style.display = 'block';
  }
  setType(S.expenseType);
  bindAmountField();
}

function bindAmountField() {
  const amount = $('fAmount');
  const submit = $('submitBtn');

  // The on-screen keyboard covers the button, so pull it into view on focus.
  for (const el of [amount, $('fNote')]) {
    el.addEventListener('focus', () => setTimeout(() => submit.scrollIntoView({ behavior: 'smooth', block: 'center' }), 300));
  }

  amount.addEventListener('input', () => {
    const clean = amount.value
      .replace(/\s/g, '').replace(',', '.')
      .replace(/[^\d.]/g, '')
      .replace(/^(\d*\.?\d*).*$/, '$1'); // digits with at most one decimal point
    const [whole, fraction] = clean.split('.');
    const grouped = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    const formatted = fraction !== undefined ? `${grouped}.${fraction}` : grouped;
    if (amount.value !== formatted) amount.value = formatted;
  });
}

async function init() {
  const tg = window.Telegram?.WebApp;
  if (tg) {
    tg.ready();
    tg.expand();
    S.tgUser = tg.initDataUnsafe?.user || null;
    tg.BackButton?.onClick(() => history.back());
  }
  window.addEventListener('popstate', e => showTab(e.state?.tab || 'main', false));

  await cloudSync();
  Object.assign(S, storedState());
  migrateV1();

  if (!S.lang) { $('lang-picker').style.display = 'flex'; return; }
  await fetchDictionary(S.lang);
  if (!S.setupDone) {
    applyLang();
    promptUserName();
    return;
  }
  startApp();
}

init();
