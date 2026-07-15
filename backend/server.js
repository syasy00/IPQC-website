/**
 * IPQC Tracker backend
 * ---------------------------------------------------------
 * No database. The Excel file at DATA_DIR/ipqc-tracker.xlsx IS the
 * data store. The frontend reads/writes through this API, and
 * Power BI can point straight at GET /api/download (or the raw
 * file if you mount a persistent disk) to pick up both the
 * historical rows you seed once and every new submission.
 * ---------------------------------------------------------
 */
const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 4000;

// On Render, mount a persistent disk and set DATA_DIR to its path
// (e.g. /var/data) so the Excel file survives restarts/redeploys.
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const DATA_FILE = path.join(DATA_DIR, 'ipqc-tracker.xlsx');
const UPLOADS_DIR = path.join(DATA_DIR, 'uploads');
const SHEET_NAME = 'IPQC Tracker';
const excelUpload = multer({
  dest: path.join(DATA_DIR, 'imports')
});

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
  { header: 'Status', key: 'status', width: 12 },
  { header: 'ICAR No.', key: 'icarNum', width: 16 },
  { header: 'Action Taken', key: 'actionTaken', width: 30 },
  { header: 'MQE Engineer', key: 'mqeEngineer', width: 16 },
];

const PLATFORM_MQE_MAPPING = { Apex: 'Siti Naimah', PDX: 'Larry', Navigator: 'Farid' };

function calculateWW(dateStr) {
  if (!dateStr) return '';
  const date = new Date(dateStr);
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
  return String(weekNo);
}


app.use(cors());
app.use(express.json({ limit: '5mb' }));

// ---------- Excel helpers ----------
async function ensureWorkbook() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, {
      recursive: true
    });
  }
  if (!fs.existsSync(UPLOADS_DIR)) {
    fs.mkdirSync(UPLOADS_DIR, {
      recursive: true
    });
  }
  const IMPORTS_DIR = path.join(
    DATA_DIR,
    'imports'
  );
  if (!fs.existsSync(IMPORTS_DIR)) {
    fs.mkdirSync(IMPORTS_DIR, {
      recursive: true
    });
  }
  if (fs.existsSync(DATA_FILE)) {
    return;
  }
  const workbook =
    new ExcelJS.Workbook();
  const sheet =
    workbook.addWorksheet(
      SHEET_NAME
    );
  sheet.columns = COLUMNS;
  sheet.getRow(1).font = {
    bold: true
  };
  await workbook.xlsx.writeFile(
    DATA_FILE
  );
  console.log(
    'Created empty workbook'
  );
}

function cellText(cell) {
  const v = cell.value;

  if (v == null) return '';

  if (typeof v === 'object') {
    if (v.text) return v.text;

    if (v.richText) {
      return v.richText.map(x => x.text).join('');
    }

    if (v.result) {
      return String(v.result);
    }
  }

  return String(v);
}

app.post(
  '/api/import',
  excelUpload.single('file'),
  async (req, res) => {
    try {

      const importedBook = new ExcelJS.Workbook();
      await importedBook.xlsx.readFile(req.file.path);

      const sourceSheet = importedBook.worksheets[0];
      const images = sourceSheet.getImages();

console.log(
  'IMAGES FOUND:',
  JSON.stringify(images, null, 2)
);

      const masterBook = new ExcelJS.Workbook();
      await masterBook.xlsx.readFile(DATA_FILE);

      const masterSheet =
        masterBook.getWorksheet(SHEET_NAME);
      const startNo = masterSheet.rowCount;

      let count = 0;
      

      sourceSheet.eachRow((row, rowNumber) => {

        if (rowNumber === 1) return;

        if (!cellText(row.getCell(1))) return;

        count++;



masterSheet.addRow([
  Date.now() + count,                 // ID
  startNo + count,                 // No

  cellText(row.getCell(1)),           // Date
  cellText(row.getCell(2)),           // WW
  cellText(row.getCell(3)),           // Shift
  cellText(row.getCell(4)),           // Auditor
  cellText(row.getCell(5)),           // PIC Finding
  cellText(row.getCell(6)),           // Department
  cellText(row.getCell(7)),           // Platform
  cellText(row.getCell(8)),           // Area / Station
  cellText(row.getCell(9)),           // Group Finding
  cellText(row.getCell(10)),          // Category
  cellText(row.getCell(11)),          // Finding Details

  '',                                 // Picture URL

  cellText(row.getCell(13)),          // Remark
  cellText(row.getCell(14)),          // Status
  cellText(row.getCell(15)),          // ICAR

  '',                                 // Action Taken
  ''                                  // MQE Engineer
]);

        
      });

      await masterBook.xlsx.writeFile(DATA_FILE);

      res.json({
        success: true,
        imported: count
      });

    } catch (err) {
      console.error(err);

      res.status(500).json({
        error: 'Import failed',
        details: err.message
      });
    }
  }
);



async function readAllRecords() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(DATA_FILE);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  const records = [];
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // header
    const rec = {};
    COLUMNS.forEach((col, idx) => {
      const cell = row.getCell(idx + 1);
      rec[col.key] = cell.value == null ? '' : cell.value;
    });
    if (rec.id !== '') records.push(rec);
  });
  return records;
}

async function appendRecord(record) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(DATA_FILE);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  if (!sheet) {
    throw new Error(`Worksheet "${SHEET_NAME}" not found`);
  }
  console.log("Before Add:", sheet.rowCount);
  sheet.addRow([
    record.id,
    record.no,
    record.auditDate,
    record.ww,
    record.shift,
    record.auditors,
    record.personOnJob,
    record.department,
    record.platform,
    record.areaStation,
    record.groupFinding,
    record.category,
    record.detailsFindings,
    record.picture,
    record.remark,
    record.status,
    record.icarNum,
    record.actionTaken,
    record.mqeEngineer
  ]);

  console.log("After Add:", sheet.rowCount);
  await workbook.xlsx.writeFile(DATA_FILE);
  console.log("Excel Saved");
  const verifyWorkbook = new ExcelJS.Workbook();
  await verifyWorkbook.xlsx.readFile(DATA_FILE);
  const verifySheet =
    verifyWorkbook.getWorksheet(SHEET_NAME);
  console.log(
    "Verify Rows:",
    verifySheet.rowCount
  );
}

async function updateRecord(id, patch) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(DATA_FILE);

  const sheet = workbook.getWorksheet(SHEET_NAME);

  let updated = null;

  for (let rowNumber = 2; rowNumber <= sheet.rowCount; rowNumber++) {

    const row = sheet.getRow(rowNumber);

    const currentId = String(
      row.getCell(1).value || ''
    ).trim();

    if (currentId === String(id).trim()) {

      patch.id = currentId;

      if (!patch.no) {
        patch.no = row.getCell(2).value;
      }

      COLUMNS.forEach((col, idx) => {

        if (patch[col.key] !== undefined) {

          row.getCell(idx + 1).value =
            patch[col.key];

        }

      });

      updated = {};

      COLUMNS.forEach((col, idx) => {

        updated[col.key] =
          row.getCell(idx + 1).value;

      });

      break;
    }
  }

  if (updated) {
    await workbook.xlsx.writeFile(DATA_FILE);
  }

  return updated;
}

async function deleteRecord(id) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(DATA_FILE);
  const sheet = workbook.getWorksheet(SHEET_NAME);
  let rowToDelete = null;
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (String(row.getCell(1).value) === String(id)) rowToDelete = rowNumber;
  });
  if (rowToDelete) {
    sheet.spliceRows(rowToDelete, 1);
    await workbook.xlsx.writeFile(DATA_FILE);
    return true;
  }
  return false;
}

async function nextIdentity() {
  const records = await readAllRecords();
  const maxNo = records.reduce((m, r) => Math.max(m, Number(r.no) || 0), 0);
  return { id: String(Date.now()), no: maxNo + 1 };
}

// ---------- Image upload ----------

const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOADS_DIR),
    filename: (req, file, cb) => {
      const ext = path.extname(file.originalname) || '.jpg';
      cb(null, `${Date.now()}-${Math.round(Math.random() * 1e6)}${ext}`);
    },
  }),
  limits: { fileSize: 8 * 1024 * 1024 },
});

app.use('/uploads', express.static(UPLOADS_DIR));

app.post('/api/upload', upload.single('image'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No image provided' });
  const publicUrl = `${req.protocol}://${req.get('host')}/uploads/${req.file.filename}`;
  res.json({ url: publicUrl });
});

// ---------- Records API ----------

app.get('/api/records', async (req, res) => {
  try {
    const records = await readAllRecords();
    res.json(records);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to read Excel data file' });
  }
});

app.post('/api/records', async (req, res) => {
  try {
    console.log('POST BODY:', req.body);
    const body = req.body || {};
    if (!body.auditDate || !body.auditors || !body.department) {
      return res.status(400).json({
        error: 'auditDate, auditors and department are required'
      });
    }

    const { id, no } = await nextIdentity();
    const record = {
      id,
      no,
      auditDate: body.auditDate,
      ww: calculateWW(body.auditDate),
      shift: body.shift || 'D',
      auditors: body.auditors,
      personOnJob: body.personOnJob || '',
      department: body.department,
      platform: body.platform || '',
      areaStation: body.areaStation || '',
      groupFinding: body.groupFinding || '',
      category: body.category || '',
      detailsFindings: body.detailsFindings || '',
      picture: body.picture || '',
      remark: body.remark || '',
      status: body.status || 'Open',
      icarNum: body.icarNum || '',
      actionTaken: body.actionTaken || '',
      mqeEngineer:
        body.mqeEngineer ||
        PLATFORM_MQE_MAPPING[body.platform] ||
        '',
    };

    console.log('NEW RECORD:', record);
    await appendRecord(record);
    res.status(201).json(record);
  } catch (err) {
    console.error('CREATE ERROR:', err);
    res.status(500).json({
      error: 'Failed to append to Excel data file'
    });

  }

});

app.put('/api/records/:id', async (req, res) => {

  console.log('UPDATE ID:', req.params.id);
  console.log('UPDATE BODY:', req.body);

  try {

    const updated =
      await updateRecord(
        req.params.id,
        req.body || {}
      );

    console.log('UPDATED RECORD:', updated);

    if (!updated) {
      return res.status(404).json({
        error: 'Record not found'
      });
    }

    res.json(updated);

  } catch (err) {

    console.error(err);

    res.status(500).json({
      error: 'Failed to update Excel data file'
    });
  }

});


app.delete('/api/records/:id', async (req, res) => {
  try {
    const ok = await deleteRecord(req.params.id);
    if (!ok) return res.status(404).json({ error: 'Record not found' });
    res.json({ deleted: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to delete from Excel data file' });
  }
});

// Direct link for Power BI (Get Data -> Web) or manual download.
app.get('/api/download', (req, res) => {
  res.download(DATA_FILE, 'ipqc-tracker.xlsx');
});

app.get('/api/health', (req, res) => res.json({ ok: true }));

ensureWorkbook().then(() => {
  app.listen(PORT, () => console.log(`IPQC backend running on port ${PORT}`));
});

app.get('/api/debug', async (req, res) => {
  try {

    const exists = fs.existsSync(DATA_FILE);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(DATA_FILE);

    const sheet = workbook.getWorksheet(SHEET_NAME);

    res.json({
      dataFile: DATA_FILE,
      exists,
      rows: sheet.rowCount
    });

  } catch (err) {
    res.status(500).json({
      error: err.message
    });
  }
});

app.get('/api/test-write', async (req, res) => {

  fs.writeFileSync(
    path.join(DATA_DIR, 'test.txt'),
    new Date().toISOString()
  );

  const content = fs.readFileSync(
    path.join(DATA_DIR, 'test.txt'),
    'utf8'
  );

  res.json({
    content
  });

});
