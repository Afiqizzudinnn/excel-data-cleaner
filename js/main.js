/* global XLSX */
/**
 * Page code for the Excel Data Cleaner: load a file, run the rules, show the result and download it.
 * The rules themselves live in clean.js.
 */
import { cleanGrid } from './clean.js';
import { downloadWorkbook, readWorkbook, sheetToGrid, show } from './excel.js';
import { SAMPLE_FILE_NAME, sampleWorkbook } from './sample.js';

const PREVIEW_ROW_LIMIT = 200;
const LOG_ROW_LIMIT = 300;

const $ = (selector) => document.querySelector(selector);

const el = {
  pick: $('#pick'),
  fileInput: $('#file-input'),
  dropZone: $('#drop-zone'),
  sample: $('#sample'),
  fileName: $('#file-name'),
  sheetField: $('#sheet-field'),
  sheetSelect: $('#sheet'),
  dayOrder: $('#day-order'),
  clean: $('#clean'),
  notice: $('#notice'),
  results: $('#results'),
  summary: $('#summary'),
  download: $('#download'),
  tabs: $('#tabs'),
  preview: $('#preview'),
  log: $('#log'),
  footnote: $('#footnote'),
  options: {
    trim: $('#opt-trim'),
    empty: $('#opt-empty'),
    dupes: $('#opt-dupes'),
    dates: $('#opt-dates'),
    numbers: $('#opt-numbers'),
  },
};

const state = { workbook: null, fileName: '', isCsv: false, result: null };

const escapeHtml = (text) => String(text).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
const plural = (count, word) => `<b>${count}</b> ${word}${count === 1 ? '' : 's'}`;
const capitalise = (text) => text.charAt(0).toUpperCase() + text.slice(1);

function showNotice(message) {
  el.notice.textContent = message;
  el.notice.hidden = !message;
}

/** Wires a segmented control. `attr` is aria-checked for choices and aria-selected for tabs. */
function initSegmented(group, attr, onChange) {
  group.addEventListener('click', (event) => {
    const button = event.target.closest('button');
    if (!button) return;
    group.querySelectorAll('button').forEach((b) => b.setAttribute(attr, String(b === button)));
    onChange?.(button.dataset.value);
  });
}

const segmentedValue = (group) => group.querySelector('[aria-checked="true"]').dataset.value;

// ---------- Loading ----------

function useWorkbook(workbook, fileName, isCsv) {
  Object.assign(state, { workbook, fileName, isCsv, result: null });
  el.sheetSelect.innerHTML = workbook.SheetNames.map((name) => `<option>${escapeHtml(name)}</option>`).join('');
  el.sheetField.hidden = workbook.SheetNames.length < 2;
  el.fileName.textContent = fileName;
  el.fileName.classList.add('is-loaded');
  el.clean.disabled = false;
  el.results.hidden = true;
  showNotice('');
}

async function loadFile(file) {
  if (!file) return;
  try {
    const { workbook, isCsv } = await readWorkbook(file);
    useWorkbook(workbook, file.name, isCsv);
  } catch {
    showNotice(`Sorry, "${file.name}" could not be opened. Please use an .xlsx, .xls or .csv file.`);
  }
}

function loadSample({ scroll }) {
  useWorkbook(sampleWorkbook(), SAMPLE_FILE_NAME, false);
  runClean({ scroll });
}

// ---------- Cleaning ----------

function readOptions() {
  const switches = Object.fromEntries(Object.entries(el.options).map(([key, input]) => [key, input.checked]));
  return { ...switches, dayFirst: segmentedValue(el.dayOrder) === 'day' };
}

function runClean({ scroll }) {
  try {
    const grid = sheetToGrid(state.workbook.Sheets[el.sheetSelect.value]);
    state.result = cleanGrid(grid, readOptions(), state.isCsv);
  } catch (error) {
    showNotice(error.message);
    return;
  }
  renderResult(state.result);
  if (scroll) el.results.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ---------- Rendering ----------

function summarySentence({ before, rows, counts }) {
  const removed = [
    counts.dupes && plural(counts.dupes, 'duplicate row'),
    counts.empty && plural(counts.empty, 'empty row'),
  ].filter(Boolean);
  const actions = [];
  if (removed.length) actions.push(`removed ${removed.join(' and ')}`);
  if (counts.cells) actions.push(`fixed ${plural(counts.cells, 'cell')}`);
  const done = actions.join(removed.length > 1 ? ', and ' : ' and ');
  return `${plural(before, 'row')} went in, <b>${rows.length}</b> came out. ${done ? `${capitalise(done)}.` : 'Nothing needed fixing.'}`;
}

function renderPreview({ headers, rows, marks }) {
  const head = headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('');
  const body = rows.slice(0, PREVIEW_ROW_LIMIT).map((row, r) => {
    const cells = row.map((value, c) => {
      const before = marks[r][c];
      const classes = [before !== undefined && 'is-changed', typeof value === 'number' && 'num'].filter(Boolean).join(' ');
      const classAttr = classes ? ` class="${classes}"` : '';
      const titleAttr = before !== undefined ? ` title="Was: ${escapeHtml(before)}"` : '';
      return `<td${classAttr}${titleAttr}>${escapeHtml(show(value))}</td>`;
    });
    return `<tr>${cells.join('')}</tr>`;
  });
  el.preview.innerHTML = `<table class="data-table"><thead><tr>${head}</tr></thead><tbody>${body.join('')}</tbody></table>`;
}

function renderLog(log) {
  if (!log.length) {
    el.log.innerHTML = '<p class="empty-message">Nothing needed fixing. Your data was already clean.</p>';
    return;
  }
  const body = log.slice(0, LOG_ROW_LIMIT).map((entry) => `<tr>
    <td class="num">${entry.row}</td>
    <td>${escapeHtml(entry.column)}</td>
    <td>${escapeHtml(entry.change)}</td>
    <td>${escapeHtml(entry.before)}</td>
    <td>${escapeHtml(entry.after)}</td>
  </tr>`);
  el.log.innerHTML = `<table class="data-table">
    <thead><tr><th>Excel row</th><th>Column</th><th>What changed</th><th>Before</th><th>After</th></tr></thead>
    <tbody>${body.join('')}</tbody>
  </table>`;
}

function renderResult(result) {
  el.summary.innerHTML = summarySentence(result);
  renderPreview(result);
  renderLog(result.log);
  el.tabs.querySelector('[data-value="log"]').textContent = `Change log (${result.log.length})`;

  const notes = [];
  if (result.rows.length > PREVIEW_ROW_LIMIT) notes.push(`Showing the first ${PREVIEW_ROW_LIMIT} of ${result.rows.length} rows.`);
  if (result.log.length > LOG_ROW_LIMIT) notes.push(`The log shows the first ${LOG_ROW_LIMIT} of ${result.log.length} changes.`);
  notes.push('The download has 3 sheets: Cleaned Data, Change Log and Summary.');
  el.footnote.textContent = notes.join(' ');
  el.results.hidden = false;
}

// ---------- Download ----------

function downloadResult() {
  const { headers, rows, log, counts, before } = state.result;
  const baseName = state.fileName.replace(/\.(xlsx|xls|csv)$/i, '');
  downloadWorkbook(`${baseName}_cleaned.xlsx`, [
    ['Cleaned Data', [headers, ...rows]],
    ['Change Log', [
      ['Excel row', 'Column', 'What changed', 'Before', 'After'],
      ...log.map((entry) => [entry.row, entry.column, entry.change, entry.before, entry.after]),
    ]],
    ['Summary', [
      ['Summary', 'Count'],
      ['Data rows before', before],
      ['Data rows after', rows.length],
      ['Duplicate rows removed', counts.dupes],
      ['Empty rows removed', counts.empty],
      ['Cells with extra spaces fixed', counts.spaces],
      ['Text dates turned into real dates', counts.dates],
      ['Numbers stored as text fixed', counts.numbers],
      ['Source file', state.fileName],
      ['Cleaned on', new Date().toLocaleString('en-GB')],
    ]],
  ]);
}

// ---------- Start ----------

function init() {
  el.pick.addEventListener('click', () => el.fileInput.click());
  el.fileInput.addEventListener('change', () => {
    loadFile(el.fileInput.files[0]);
    el.fileInput.value = '';
  });

  ['dragenter', 'dragover'].forEach((type) => el.dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    el.dropZone.classList.add('is-dragover');
  }));
  ['dragleave', 'drop'].forEach((type) => el.dropZone.addEventListener(type, (event) => {
    event.preventDefault();
    el.dropZone.classList.remove('is-dragover');
  }));
  el.dropZone.addEventListener('drop', (event) => loadFile(event.dataTransfer.files[0]));

  el.sample.addEventListener('click', () => loadSample({ scroll: true }));
  el.clean.addEventListener('click', () => runClean({ scroll: true }));
  el.download.addEventListener('click', downloadResult);
  initSegmented(el.dayOrder, 'aria-checked');
  initSegmented(el.tabs, 'aria-selected', (tab) => {
    el.preview.hidden = tab !== 'preview';
    el.log.hidden = tab !== 'log';
  });

  if (typeof XLSX === 'undefined') {
    showNotice('The spreadsheet library did not load. Check your internet connection and refresh the page.');
    return;
  }
  if (new URLSearchParams(window.location.search).has('demo')) loadSample({ scroll: false });
}

init();
