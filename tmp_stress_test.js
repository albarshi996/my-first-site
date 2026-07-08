const fs = require('fs');
const path = require('path');
const vm = require('vm');
const XLSX = require('xlsx');

const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
const scriptMatch = html.match(/<script>([\s\S]*)<\/script>/);
if (!scriptMatch) throw new Error('No script found in index.html');

const script = scriptMatch[1];
const startMarker = 'function addSheet(wb,data,name) {';
const endMarker = "window.addEventListener('DOMContentLoaded', () => {";
const start = script.indexOf(startMarker);
const end = script.indexOf(endMarker);
if (start === -1 || end === -1) throw new Error('Could not locate export functions block');
const exportBlock = script.slice(start, end);

class MockElement {
  constructor(tagName = 'div') {
    this.tagName = tagName.toUpperCase();
    this.children = [];
    this.style = {};
    this.className = '';
    this.textContent = '';
    this.innerHTML = '';
    this.id = '';
    this.value = '';
    this.disabled = false;
    this.classList = {
      add: () => {},
      remove: () => {},
      toggle: () => {},
      contains: () => false
    };
    this.dataset = {};
  }
  appendChild(child) {
    this.children.push(child);
    if (child && child.id) {
      if (!this.ownerDocument) return child;
      this.ownerDocument._elementsById.set(child.id, child);
    }
    return child;
  }
  setAttribute(name, value) {
    this[name] = value;
  }
}

class MockDocument {
  constructor() {
    this._elementsById = new Map();
    this.body = new MockElement('body');
    this.body.ownerDocument = this;
  }
  createElement(tagName) {
    const el = new MockElement(tagName);
    el.ownerDocument = this;
    return el;
  }
  getElementById(id) {
    return this._elementsById.get(id) || null;
  }
  addEventListener() {}
  removeEventListener() {}
  querySelectorAll() { return []; }
  querySelector() { return null; }
}

const document = new MockDocument();
const window = { addEventListener() {}, removeEventListener() {} };
const context = {
  console,
  window,
  document,
  navigator: {},
  setTimeout: (fn) => { fn(); },
  clearTimeout: () => {},
  XLSX,
  DB: {
    users: [],
    documents: [],
    versions: [],
    audit: [],
    trash: [],
    categories: [],
    structure: { departments: [], projects: [], branches: [] },
    settings: {}
  },
  currentUser: { name: 'admin' },
  isDirty: false,
  STATUS_LABELS: { active: 'نشط', expired: 'منتهي', expiring: 'ينتهي' },
  docStatus: () => 'active',
  fmtDate: (v) => v || '',
  toast: () => {},
  markClean: () => {},
  encPw: (v) => v,
  genId: () => String(Math.random()),
  safeJSON: (val, def) => { try { return val ? JSON.parse(val) : def; } catch { return def; } },
  applyWorkbook: null,
  bh_showExportProgress: null,
  bh_hideExportProgress: null,
  bh_buildWorkbookChunkedAsync: null,
  bh_writeWorkbookChunked: null,
  bh_runStressTestExport: null,
  addSheet: null,
  bh_sanitizeExcelCell: null,
  bh_sanitizeSheetRows: null,
  bh_buildDocumentExportRow: null,
  bh_buildWorkbookChunked: null
};
context.window.document = document;
context.window.window = window;
context.global = context;
context.globalThis = context;
vm.createContext(context);
vm.runInContext(exportBlock, context);

const { bh_runStressTestExport, bh_buildWorkbookChunkedAsync, bh_writeWorkbookChunked, bh_buildWorkbookChunked, addSheet, bh_sanitizeExcelCell, bh_sanitizeSheetRows, bh_buildDocumentExportRow, bh_buildWorkbookChunkedAsync: buildAsync } = context;

context.applyWorkbook = function(wb, cb) {
  try {
    const sn = wb.SheetNames;
    if (sn.includes('Users')) {
      const rows = XLSX.utils.sheet_to_json(wb.Sheets['Users']);
      if (rows.length) context.DB.users = rows.map(r => ({ id:r.id||'u', name:r.name||'', password:r.password||'', dept:r.dept||'', section:r.section||'', role:r.role||'مشاهدة', permissions: {}, twoFA:!!r.twoFA, twoFACode:r.twoFACode||'', lastLogin:r.lastLogin||'' }));
    }
    if (sn.includes('Documents')) {
      const importedDocs = XLSX.utils.sheet_to_json(wb.Sheets['Documents']).map(r => ({ id:r.id||'d', docNum:r.docNum||'', title:r.title||'', type:r.type||'', category:r.category||'', dept:r.dept||'', section:r.section||'', project:r.project||'', source:r.source||'', client:r.client||'', confidential:r.confidential||'عام', issue:r.issue||'', expiry:r.expiry||'', status:r.status||'active', approvedBy:r.approvedBy||'', approvedDate:r.approvedDate||'', files:r.files||'', keywords:r.keywords||'', notes:r.notes||'', ocr:r.ocr||'', filePath:r.filePath||'', signature: r.signature || null, tracking: r.tracking || [] }));
      context.DB.documents = importedDocs;
    }
    if (cb) cb();
  } catch (err) {
    console.error(err);
  }
};

const outDir = path.join(__dirname, 'tmp-stress-output');
fs.mkdirSync(outDir, { recursive: true });

const longOcrBase = 'OCR '.repeat(20000) + '\n';
const testDocs = [];
for (let i = 0; i < 2000; i++) {
  testDocs.push({
    id: 'stress-' + i,
    docNum: `STRESS-${String(i + 1).padStart(4, '0')}`,
    title: `وثيقة اختبار ${i + 1}`,
    type: 'اختبار',
    category: 'stress',
    dept: 'الاختبار',
    section: 'التحمل',
    project: 'Stress',
    source: 'Auto',
    client: 'QA',
    confidential: 'عام',
    issue: '2026-01-01',
    expiry: '2030-01-01',
    status: 'active',
    approvedBy: 'admin',
    approvedDate: '2026-01-01',
    files: `file_${i}.pdf`,
    keywords: `اختبار,تحمل,${i}`,
    notes: `ملاحظة ${i}`,
    ocr: `${longOcrBase}محتوى OCR طويل للغاية للوثيقة ${i}`,
    filePath: `/${i}/test.pdf`,
    signature: { name: 'QA', title: 'اختبار' },
    tracking: [{ user: 'admin', action: 'إضافة', date: new Date().toISOString() }]
  });
}

context.DB.documents = testDocs;
(async () => {
  const wb = await buildAsync({ onProgress: () => {} });
  const filename = path.join(outDir, 'Brandzo_StressTest.xlsx');
  const writeBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  fs.writeFileSync(filename, writeBuffer);

  const loaded = XLSX.readFile(filename);
  const rows = XLSX.utils.sheet_to_json(loaded.Sheets['Documents']);
  const roundTrip = [];
  context.DB.documents = [];
  context.applyWorkbook(loaded, () => {
    roundTrip.push(...context.DB.documents);
  });

  const ok = rows.length === 2000 && roundTrip.length === 2000 && !!rows[0] && !!roundTrip[0];
  const summary = { ok, exported: rows.length, reImported: roundTrip.length, fileSize: fs.statSync(filename).size, firstOcrLength: String(rows[0]?.ocr || '').length };
  console.log(JSON.stringify(summary));
  if (!ok) process.exit(1);
})();
