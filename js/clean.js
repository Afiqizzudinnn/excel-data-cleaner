/**
 * The cleaning rules. Pure functions with no page code, so they can be read and tested on their own.
 */
import { findHeaderRow, headerNames, ISO_DATE_FORMAT, isBlank, isDate, show, ymdToSerial } from './excel.js';
import { parseDate, parseNumber } from './parse.js';

const FIX_LABELS = {
  spaces: 'Extra spaces removed',
  dates: 'Text date to real date',
  numbers: 'Text to number',
};

const WHOLE_ROW = '(whole row)';

/** Collapses runs of spaces, tabs and non-breaking spaces, keeping line breaks inside a cell. */
const tidySpaces = (text) => text.replace(/[ \t ]+/g, ' ').replace(/ ?\n ?/g, '\n').trim();

/** Key used to spot duplicate rows: same values, same types. */
const rowKey = (row) => JSON.stringify(row.map((value) => (isDate(value) ? `d${value.serial}` : `${typeof value}:${value}`)));

/**
 * Applies the chosen fixes to one cell.
 * @returns {{ value: *, fixes: string[] }} `fixes` lists keys of FIX_LABELS that changed the cell.
 */
function cleanCell(original, options, isCsv) {
  const fixes = [];
  let value = original === undefined ? '' : original;

  if (typeof value === 'string') {
    if (options.trim) {
      const tidied = tidySpaces(value);
      if (tidied !== value) {
        value = tidied;
        fixes.push('spaces');
      }
    }
    if (options.dates && value) {
      const date = parseDate(value, options.dayFirst);
      if (date) {
        value = { date: true, serial: ymdToSerial(date.y, date.m, date.d), z: ISO_DATE_FORMAT };
        fixes.push('dates');
      }
    }
    if (typeof value === 'string' && value) {
      const number = parseNumber(value);
      // A CSV has no number type, so plain numbers in a CSV are always converted, quietly.
      if (number && (options.numbers || (isCsv && !number.formatted))) {
        value = number.value;
        if (number.formatted || !isCsv) fixes.push('numbers');
      }
    }
  } else if (options.dates && isDate(value) && Number.isInteger(value.serial)) {
    // Real dates without a time get the same display format as converted ones.
    value = { ...value, z: ISO_DATE_FORMAT };
  }

  return { value, fixes };
}

/**
 * Cleans a sheet.
 * @param {{ n: number, cells: Array }[]} grid  rows from sheetToGrid()
 * @param {{ trim: boolean, empty: boolean, dupes: boolean, dates: boolean, dayFirst: boolean, numbers: boolean }} options
 * @param {boolean} isCsv
 * @returns {{
 *   headers: string[], rows: Array[], before: number,
 *   marks: Object<number, string>[],   per kept row: column index -> value before cleaning
 *   log: { row: number, column: string, change: string, before: string, after: string }[],
 *   counts: { spaces: number, dates: number, numbers: number, cells: number, empty: number, dupes: number }
 * }}
 */
export function cleanGrid(grid, options, isCsv) {
  const headerIndex = findHeaderRow(grid);
  if (headerIndex < 0) throw new Error('This sheet is empty. Pick another sheet or file.');

  const headers = headerNames(grid[headerIndex].cells);
  const data = grid.slice(headerIndex + 1);
  const counts = { spaces: 0, dates: 0, numbers: 0, cells: 0, empty: 0, dupes: 0 };
  const rows = [];
  const marks = [];
  const log = [];
  const firstSeen = new Map();

  for (const { n, cells } of data) {
    const cellFixes = [];
    const mark = {};
    const row = headers.map((column, c) => {
      const { value, fixes } = cleanCell(cells[c], options, isCsv);
      if (fixes.length) {
        cellFixes.push({ column, fixes, before: show(cells[c]), after: show(value) });
        mark[c] = show(cells[c]);
      }
      return value;
    });

    if (row.every(isBlank)) {
      if (options.empty) {
        counts.empty++;
        log.push({ row: n, column: WHOLE_ROW, change: 'Empty row removed', before: '', after: '' });
        continue;
      }
    } else if (options.dupes) {
      const key = rowKey(row);
      if (firstSeen.has(key)) {
        counts.dupes++;
        log.push({ row: n, column: WHOLE_ROW, change: `Duplicate of row ${firstSeen.get(key)} removed`, before: row.map(show).join(' | '), after: '' });
        continue;
      }
      firstSeen.set(key, n);
    }

    // Only count fixes for rows that are kept.
    for (const { column, fixes, before, after } of cellFixes) {
      fixes.forEach((key) => { counts[key]++; });
      counts.cells++;
      log.push({ row: n, column, change: fixes.map((key) => FIX_LABELS[key]).join('; '), before, after });
    }
    rows.push(row);
    marks.push(mark);
  }

  log.sort((a, b) => a.row - b.row);
  return { headers, rows, marks, log, counts, before: data.length };
}
