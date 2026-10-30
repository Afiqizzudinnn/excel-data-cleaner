/* global XLSX */
/**
 * A small, deliberately messy expense claims export for the "Try it with sample data" button.
 * All names and numbers are made up.
 */

export const SAMPLE_FILE_NAME = 'sample_expense_claims.xlsx';

const EMPTY_ROW = ['', '', '', '', '', ''];

const ROWS = [
  ['Claim ID', 'Staff Name', 'Department', 'Claim Date', 'Amount', 'Phone'],
  ['CL-1001', '  Nur Aisyah  binti Kamal', 'Finance', '03/03/2026', 'RM 1,250.00', '012-3456789'],
  ['CL-1002', 'Tan Wei Ming', 'Operations ', '2026-03-04', '350', '016-2223344'],
  ['CL-1003', 'Ravi Kumar', 'Sales', '5 Mar 2026', ' 89.90', '019-8765432'],
  EMPTY_ROW,
  ['CL-1004', 'Siti Hajar', '  Finance', '6/3/26', 'RM420.50', '011-23456789'],
  ['CL-1005', 'Lim Mei Ling', 'Marketing', '07-03-2026', '2,100', '017-5554433'],
  ['CL-1002', 'Tan Wei Ming ', 'Operations', '2026-03-04', '350', '016-2223344'],
  ['CL-1006', 'Hafiz Rahman', 'Sales', 'Mar 9, 2026', '(120.00)', '013-4447788'],
  EMPTY_ROW,
  ['CL-1007', 'Priya Devi', 'Human  Resources', '10.03.2026', 'RM 75', '014-9990011'],
  ['CL-1008', 'Daniel Wong', 'IT', '11 Mac 2026', '640.00', '012-7778899'],
  ['CL-1003', 'Ravi Kumar', 'Sales', '5 Mar 2026', '89.90', '019-8765432'],
  ['CL-1009', 'Aina Sofea', 'Finance', '12/03/2026', 'RM 1,980.00', '018-1212121'],
];

/** Every value is stored as text, the way a careless export would. */
export function sampleWorkbook() {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(ROWS), 'Claims');
  return workbook;
}
