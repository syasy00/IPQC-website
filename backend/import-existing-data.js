/**
 * One-time import tool.
 * Reads YOUR existing IPQC Excel tracker and converts its rows into
 * the format this app expects, so your historical audits show up
 * in the app (and therefore in Power BI) alongside new submissions.
 *
 * Usage:
 *   node import-existing-data.js "/path/to/Your Old IPQC Tracker.xlsx"
 *
 * It reads the FIRST worksheet of the file you point it at, matches
 * columns by header text (case-insensitive, ignores spacing/punctuation),
 * and writes everything into backend/data/ipqc-tracker.xlsx — replacing
 * the sample seed data. Re-run any time to re-import from scratch.
 */
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'ipqc-tracker.xlsx');
const SHEET_NAME = 'IPQC Tracker';

const COLUMNS = [
  { header: 'ID', key: 'id', width: 10 },
  { header: 'No', key: 'no', width: 8 },
  { header: 'Audit Date', key: 'auditDate', width: 14 },
  { header: 'WW', key: 'ww', width: 8 },
  { header: 'Shift', key: 'shift', width: 8 },
  { header: 'Auditors', key: 'auditors', width: 18 },
  { header: 'Person On Job', key: 'personOnJob', width: 18 },
  { header: 'Department', key: 'department', width: 16 },
  { header: 'Platform', key: 'platform', width: 14 },
  { header: 'Area / Station', key: 'areaStation', width: 20 },
  { header: 'Group Finding', key: 'groupFinding', width: 16 },
  { header: 'Category', key: 'category', width: 16 },
  { header: 'Details Findings', key: 'detailsFindings', width: 40 },
  { header: 'Picture URL', key: 'picture', width: 30 },
  { header: 'Remark', key: 'remark', width: 30 },
  { header: 'ICAR No.', key: 'icarNum', width: 16 },
  { header: 'MQE Engineer', key: 'mqeEngineer', width: 16 },
];

// Alternate header spellings your old tracker might use, mapped to our keys.
// Add more aliases here if your columns are named differently.
const HEADER_ALIASES = {
  no: ['no', 'no.', '#'],
  auditDate: ['auditdate', 'date', 'auditdt'],
  ww: ['ww', 'workweek', 'week'],
  shift: ['shift'],
  auditors: ['auditors', 'ipqcauditor', 'ipqcauditorname', 'auditor'],
  personOnJob: ['persononjob', 'picfinding', 'pic'],
  department: ['department', 'dept'],
  platform: ['platform'],
  areaStation: ['areastation', 'area/station', 'area', 'station'],
  groupFinding: ['groupfinding'],
  category: ['category'],
  detailsFindings: ['detailsfindings', 'findingdetails', 'details', 'findings'],
  picture: ['picture', 'pictureurl', 'photo', 'image'],
  remark: ['remark', 'remarks'],
  icarNum: ['icarnum', 'icarno', 'icarno.', 'icar'],
  mqeEngineer: ['mqeengineer', 'mqe'],
};

function normalize(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
}

function calculateWW(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  if (isNaN(date.getTime())) return '';
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
  return String(weekNo);
}

function cellToPlainValue(cell) {
  const v = cell.value;
  if (v == null) return '';
  if (v instanceof Date) return v.toISOString().split('T')[0];
  if (typeof v === 'object' && v.text) return v.text; // rich text
  if (typeof v === 'object' && v.result !== undefined) return v.result; // formula
  return v;
}

async function main() {
  const sourcePath = process.argv[2];
  if (!sourcePath) {
    console.error('Usage: node import-existing-data.js "/path/to/Your Old Tracker.xlsx"');
    process.exit(1);
  }
  if (!fs.existsSync(sourcePath)) {
    console.error('File not found:', sourcePath);
    process.exit(1);
  }

  const source = new ExcelJS.Workbook();
  await source.xlsx.readFile(sourcePath);
  const sourceSheet = source.worksheets[0];
  if (!sourceSheet) {
    console.error('No worksheet found in the source file.');
    process.exit(1);
  }

  // Map source columns -> our keys, by header text.
  const headerRow = sourceSheet.getRow(1);
  const colIndexToKey = {};
  headerRow.eachCell((cell, colNumber) => {
    const norm = normalize(cell.value);
    for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
      if (aliases.includes(norm)) { colIndexToKey[colNumber] = key; break; }
    }
  });

  const matchedKeys = new Set(Object.values(colIndexToKey));
  console.log('Matched columns:', [...matchedKeys].join(', ') || '(none — check your header names)');

  const records = [];
  let rowCounter = 0;
  sourceSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // header
    const rec = {
      id: '', no: '', auditDate: '', ww: '', shift: 'D', auditors: '', personOnJob: '',
      department: '', platform: '', areaStation: '', groupFinding: '', category: '',
      detailsFindings: '', picture: '', remark: '', icarNum: '', mqeEngineer: '',
    };
    let hasData = false;
    row.eachCell((cell, colNumber) => {
      const key = colIndexToKey[colNumber];
      if (!key) return;
      const val = cellToPlainValue(cell);
      if (val !== '') hasData = true;
      rec[key] = val;
    });
    if (!hasData) return;
    rowCounter += 1;
    rec.id = String(rowCounter);
    rec.no = rec.no || rowCounter;
    if (!rec.ww) rec.ww = calculateWW(rec.auditDate);
    records.push(rec);
  });

  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

  const target = new ExcelJS.Workbook();
  const sheet = target.addWorksheet(SHEET_NAME);
  sheet.columns = COLUMNS;
  sheet.getRow(1).font = { bold: true };
  records.forEach((r) => sheet.addRow(r));
  await target.xlsx.writeFile(DATA_FILE);

  console.log(`Imported ${records.length} records into ${DATA_FILE}`);
  if (records.length === 0) {
    console.log('Nothing imported — open the source file and check that its header row (row 1) has recognizable column names, or add aliases in HEADER_ALIASES.');
  }
}

main().catch((err) => { console.error(err); process.exit(1); });
