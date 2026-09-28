const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('code.gs', 'utf8');
const brainSource = fs.readFileSync('BrainEngine.gs', 'utf8');
const html = fs.readFileSync('supplier_registration_form.html', 'utf8');
const adminHtml = fs.readFileSync('admin_message.html', 'utf8');
const manifest = JSON.parse(fs.readFileSync('appsscript.json', 'utf8'));

new Function(source);
[html, adminHtml].forEach(page => {
  [...page.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi)]
    .forEach(match => new Function(match[1]));
});
assert(!source.includes('GmailApp.sendEmail'));
assert(!source.includes('Session.getEffectiveUser().getEmail()'));
assert(source.includes('scheduleBackgroundAIQueue()'));
assert(source.includes("const SUBMISSION_ID_COLUMN = 82;"));
assert(!source.includes('function adminMessageRecordKey'));
assert(!html.includes('google.script.run.runAIOnSpecificRow'));
assert(!html.includes("lastFolderUrl = 'https://drive.google.com'"));
assert(source.includes('assertFileInSubmissionRoot(file);'));
assert(source.includes('assertFolderInSubmissionRoot(folder);'));
assert(brainSource.includes('var ss = getDataSpreadsheet();'));
assert.strictEqual((source.match(/SpreadsheetApp\.getActiveSpreadsheet\(\)/g) || []).length, 2);
assert.deepStrictEqual(manifest.webapp, {
  access: 'ANYONE_ANONYMOUS',
  executeAs: 'USER_DEPLOYING'
});
assert(html.includes("type: 'select', name: 'currency', label: 'สกุลเงิน / Currency', required: true"));
assert(html.includes("{ value: 'THB', label: 'THB / บาทไทย' }"));
assert(html.includes("{ value: 'USD', label: 'USD / ดอลลาร์สหรัฐ' }"));

const currencyContext = {};
vm.createContext(currencyContext);
const currencyStart = source.indexOf('const CURRENCY_CODES');
const currencyEnd = source.indexOf('\nfunction label', currencyStart);
vm.runInContext(source.slice(currencyStart, currencyEnd), currencyContext);
assert.strictEqual(currencyContext.normalizeCurrency('TBH'), 'THB');
assert.strictEqual(currencyContext.normalizeCurrency(' baht '), 'THB');
assert.strictEqual(currencyContext.normalizeCurrency('random'), '');

const columnsContext = {
  val: (answers, key) => answers && answers[key] != null ? answers[key] : '',
  normalizeCurrency: currencyContext.normalizeCurrency
};
vm.createContext(columnsContext);
vm.runInContext(source.slice(source.indexOf('const LABELS'), source.indexOf('function labelJoin')), columnsContext);
const columnsStart = source.indexOf('function buildInitialColumns');
const columnsEnd = source.indexOf('\n/**\n * Isolated writer', columnsStart);
vm.runInContext(source.slice(columnsStart, columnsEnd), columnsContext);
assert.strictEqual(columnsContext.buildInitialColumns().length, 66);
assert.strictEqual(columnsContext.buildInitialColumns()[63]({ answers: { currency: 'TBH' } }), 'THB / บาทไทย');
for (const option of html.matchAll(/\{ value: '([A-Z]{3})', label: '([^']+)' \}/g)) {
  const saved = columnsContext.buildInitialColumns()[63]({ answers: { currency: option[1] } });
  assert.strictEqual(saved, option[2]);
  assert.strictEqual(currencyContext.normalizeCurrency(saved), option[1]);
}

// Exercise the actual checklist evaluator, without calling Google or sending mail.
let auditRules = [];
let auditResult = { passed: true };
let auditBlob = null;
let auditedAttachments = null;
const auditContext = {
  getRulesFromSheetOnly: () => auditRules,
  getAIModelConfig: () => ({ modelName: 'test-model' }),
  determineVendorEntityType: () => false,
  getBlobFromDriveUrlOrId: () => auditBlob,
  parseDriveId: () => '',
  logLiveJob: () => {},
  Logger: { log: () => {} },
  BrainEngine: { checkIfBlankOrEmptyDocument: blob => {
    assert(!Array.isArray(blob), 'Do not treat a file array as a corrupt blob');
    return { isBlank: false };
  } },
  verifySingleDocWithGemini: blobs => { auditedAttachments = blobs; return auditResult; }
};
vm.createContext(auditContext);
vm.runInContext(source.slice(source.indexOf('function evaluateDocumentRulesFromSheet'), source.indexOf('function generateFM021PdfBlob')), auditContext);
const completeAnswers = { companyName: 'Test', taxId: '123', ownerName: 'Owner', accountName: 'Test', accountNumber: '12345678', currency: 'THB', signedName: 'Owner' };
auditRules = [{ order: 1, docKey: 'fm021', isMandatory: true }];
let audit = auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, { fm021: {} }, 'test');
assert.strictEqual(audit.overallStatus, 'PASSED');
audit = auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, {}, 'test');
assert.strictEqual(audit.overallStatus, 'NEED_REVIEW');
assert.strictEqual(audit.items[0].modelUsed, 'RuleEngine-FM021');
auditRules = [{ order: 3, docKey: 'docAffidavit', isMandatory: true }];
auditResult = { passed: null, note: 'API unavailable' };
audit = auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, { docAffidavit: {} }, 'test');
assert.strictEqual(audit.overallStatus, 'NEED_REVIEW');
auditResult = null;
assert.strictEqual(auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, { docAffidavit: {} }, 'test').overallStatus, 'NEED_REVIEW');
auditResult = { passed: true };
const attachments = [{}, {}];
assert.strictEqual(auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, { docAffidavit: attachments }, 'test').overallStatus, 'PASSED');
assert.strictEqual(auditedAttachments, attachments);
audit = auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers, fileUrls: { docAffidavit: 'unreadable' } }, {}, 'test');
assert.strictEqual(audit.overallStatus, 'NEED_REVIEW');
assert.strictEqual(auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, {}, 'test').overallStatus, 'FAILED');
auditResult = { passed: false, note: 'Confirmed mismatch' };
assert.strictEqual(auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, { docAffidavit: {} }, 'test').overallStatus, 'FAILED');

const start = source.indexOf('function sendSubmissionEmail');
const end = source.indexOf('\nfunction getSubmissionFolder', start);
const sent = [];
const context = {
  NOTIFY_ON_SUBMIT: true,
  NOTIFY_EMAILS: ['supplier@example.com', 'ADMIN@example.com', 'admin@example.com'],
  ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/example/exec' }) },
  Utilities: { formatDate: () => '17/09/2026 12:00:00' },
  Session: { getScriptTimeZone: () => 'Asia/Bangkok' },
  MailApp: {
    getRemainingDailyQuota: () => 100,
    sendEmail: message => sent.push(message)
  },
  Logger: { log: () => {} },
  parseDriveId: () => '',
  isValidEmailAddress: email => /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(email),
  isValidSubmissionId: value => /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || '').trim()),
  createAdminMessageToken_: () => 'signed-token',
  emailHtml: value => String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;'),
  emailHttpsUrl: value => /^https:\/\/[^\s"'<>]+$/i.test(String(value || '').trim()) ? String(value).trim() : '#',
  result: null
};
vm.createContext(context);
vm.runInContext(source.slice(start, end) + '\nresult = sendSubmissionEmail("ACME", "123", {items: []}, "", "", null, null, "", 2, "supplier@example.com", "123e4567-e89b-42d3-a456-426614174000");', context);

assert.strictEqual(context.result, 1);
assert.strictEqual(sent.length, 1);
assert.strictEqual(sent[0].to, 'admin@example.com');
assert(sent[0].htmlBody && sent[0].body);
assert(sent[0].htmlBody.includes('action=admin_message'));
assert(!sent[0].to.includes('supplier@example.com'));

const emailChecklist = { overallStatus: 'FAILED', summaryText: 'ชื่อผู้ส่งมอบในเอกสาร ESG ไม่ตรงกับแบบฟอร์ม และมีเอกสารรอเจ้าหน้าที่ตรวจสอบ', items: [
  { order: 1, name: 'ทะเบียนประวัติผู้ส่งมอบ (FM021)', statusBool: true, statusText: 'ตรวจสอบแล้ว', isMandatory: true, note: 'กรอกข้อมูลบังคับครบถ้วน และสร้างเอกสารแล้ว' },
  { order: 8, name: 'ESG Self-Assessment', statusBool: false, statusText: 'เอกสารไม่ผ่านเงื่อนไข', isMandatory: true, condition: 'ชื่อผู้ส่งมอบตรงกับฟอร์ม', modelUsed: 'gemini-3.5-flash-lite', duration: 3.4,
    note: 'ไม่ผ่าน: ชื่อผู้ส่งมอบในเอกสารไม่ตรงกับชื่อในฟอร์ม\nต้องส่งแบบประเมินของกิจการที่ลงทะเบียน\n\nหลักฐานประกอบจาก AI (อาจคลาดเคลื่อน):\n1. ชื่อผู้ส่งมอบ\nคาดหวัง: บริษัท ตัวอย่าง ซัพพลาย จำกัด\nพบ: บริษัท อื่น จำกัด — ไฟล์ 1 หน้า 2 ช่องชื่อผู้ส่งมอบ\nเหตุผล: เป็นคนละกิจการ จึงไม่ผ่านแม้ตอบครบทุกข้อ\n\n2. ความครบถ้วน\nพบการตอบหัวข้อสิ่งแวดล้อม สังคม และธรรมาภิบาล แต่ไม่ชดเชยชื่อบริษัทที่ไม่ตรง', viewUrl: 'https://example.com/esg', downloadUrl: 'https://example.com/esg.pdf' },
  { order: 3, name: 'สำเนาหนังสือรับรองบริษัท', statusBool: false, needsManualReview: true, statusText: 'รอเจ้าหน้าที่ตรวจสอบ', isMandatory: true, note: 'ระบบอ่านไฟล์ไม่สำเร็จ จึงยังยืนยันวันที่ออกเอกสารไม่ได้', viewUrl: 'https://example.com/certificate' },
  { order: 10, name: 'Company Profile', statusBool: false, statusText: 'ยังไม่ส่ง', isMandatory: false, note: 'ไม่ได้แนบเอกสารเพิ่มเติม' }
] };
context.sendSubmissionEmail('บริษัท ตัวอย่าง ซัพพลาย จำกัด', '0100000000000', emailChecklist, 'https://example.com/fm021.pdf', 'https://example.com/files', {}, {}, 'https://example.com/report.pdf', 2, 'supplier@example.com', '123e4567-e89b-42d3-a456-426614174000');
const redesignedEmail = sent[sent.length - 1];
assert(redesignedEmail.htmlBody.indexOf('ESG Self-Assessment') < redesignedEmail.htmlBody.indexOf('ทะเบียนประวัติผู้ส่งมอบ (FM021)'));
assert(redesignedEmail.htmlBody.includes('รอตรวจ / Review'));
assert(redesignedEmail.htmlBody.includes('ข้อมูลประกอบ / Other'));
assert(redesignedEmail.htmlBody.includes('ไฟล์ 1 หน้า 2'));
assert(!redesignedEmail.htmlBody.includes('<thead>'));
assert(!redesignedEmail.htmlBody.includes('href="#"'));
assert.strictEqual(redesignedEmail.attachments.length, 2);
if (process.argv.includes('--email-preview')) fs.writeFileSync('admin-email-preview.html', redesignedEmail.htmlBody);
context.sendSubmissionEmail('<script>alert(1)</script>', '123', { overallStatus: 'NEED_REVIEW', items: [{ name: '<img src=x>', note: '<b>untrusted</b>', needsManualReview: true }] }, '', '', null, {}, '', 0, '', '');
assert(!sent[sent.length - 1].htmlBody.includes('<script>'));
assert(sent[sent.length - 1].htmlBody.includes('&lt;b&gt;untrusted&lt;/b&gt;'));

const reportStart = source.indexOf('function saveOrUpdateSingleAiReport');
const reportEnd = source.indexOf('\nconst DEFAULT_RULES', reportStart);
let requestedReportName = '';
let trashedOldReport = false;
let savedReportName = '';
const oldReport = { setTrashed: () => { trashedOldReport = true; } };
const reportContext = {
  isValidSubmissionId: context.isValidSubmissionId,
  DriveApp: { Access: { ANYONE_WITH_LINK: 1 }, Permission: { VIEW: 1 } },
  Logger: { log: () => {} }
};
vm.createContext(reportContext);
vm.runInContext(source.slice(reportStart, reportEnd), reportContext);
const reportBlob = { setName: name => { savedReportName = name; } };
const savedReport = reportContext.saveOrUpdateSingleAiReport({
  getFilesByName: name => {
    requestedReportName = name;
    let available = true;
    return { hasNext: () => available, next: () => { available = false; return oldReport; } };
  },
  createFile: () => ({ getId: () => 'report-id', setSharing: () => {} })
}, reportBlob, 'ACME', '123e4567-e89b-42d3-a456-426614174000');
assert.strictEqual(requestedReportName, 'ACME_123e4567-e89b-42d3-a456-426614174000_AI_Audit_Report.pdf');
assert.strictEqual(savedReportName, requestedReportName);
assert.strictEqual(trashedOldReport, true);
assert.strictEqual(savedReport.fileId, 'report-id');

function makeSheet(initialRows) {
  const rows = initialRows.map(row => row.slice());
  let maxColumns = Math.max(1, ...rows.map(row => row.length));
  return {
    rows,
    getLastRow: () => rows.length,
    getMaxColumns: () => maxColumns,
    insertColumnsAfter: (_column, count) => { maxColumns += count; },
    appendRow: row => rows.push(row.slice()),
    getRange(row, column, rowCount = 1, columnCount = 1) {
      return {
        getValue: () => (rows[row - 1] || [])[column - 1],
        getValues: () => Array.from({ length: rowCount }, (_, r) =>
          Array.from({ length: columnCount }, (_, c) => (rows[row - 1 + r] || [])[column - 1 + c] ?? '')
        ),
        setValue(value) {
          if (!rows[row - 1]) rows[row - 1] = [];
          rows[row - 1][column - 1] = value;
          maxColumns = Math.max(maxColumns, column);
        },
        setValues(values) {
          values.forEach((valueRow, r) => {
            if (!rows[row - 1 + r]) rows[row - 1 + r] = [];
            valueRow.forEach((value, c) => { rows[row - 1 + r][column - 1 + c] = value; });
          });
          maxColumns = Math.max(maxColumns, column + columnCount - 1);
        }
      };
    }
  };
}

const firstSubmissionId = crypto.randomUUID();
const secondSubmissionId = crypto.randomUUID();
const supplierRow = Array(82).fill('');
supplierRow[0] = '2026-09-17T12:00:00+07:00';
supplierRow[2] = 'ACME';
supplierRow[7] = '123';
supplierRow[16] = 'supplier@example.com';
supplierRow[81] = firstSubmissionId;
const duplicateSupplierRow = supplierRow.slice();
duplicateSupplierRow[81] = secondSubmissionId;
const supplierHeaders = Array(82).fill('Header');
supplierHeaders[81] = 'Submission ID';
const supplierSheet = makeSheet([supplierHeaders, supplierRow, duplicateSupplierRow]);
let logSheet = null;
const adminSent = [];
let mailQuota = 10;
const scriptProperties = new Map();
const adminContext = {
  ADMIN_MESSAGE_SECRET_PROPERTY: 'secret-property',
  ADMIN_MESSAGE_LOG_SHEET: 'Admin_Message_Log',
  ADMIN_MESSAGE_TOKEN_DAYS: 30,
  MAX_ADMIN_MESSAGE_LENGTH: 2000,
  SHEET_NAME: 'Form1_e777',
  SUBMISSION_ID_COLUMN: 82,
  AI_PENDING_STATUS: '⏳ รอคิวตรวจ AI',
  SUBMISSION_ID_HEADER: 'Submission ID',
  Utilities: {
    DigestAlgorithm: { SHA_256: 'sha256' },
    Charset: { UTF_8: 'utf8' },
    base64EncodeWebSafe: bytes => Buffer.from(bytes).toString('base64url'),
    computeDigest: (_algorithm, value) => [...crypto.createHash('sha256').update(String(value)).digest()],
    computeHmacSha256Signature: (value, key) => [...crypto.createHmac('sha256', key).update(value).digest()],
    getUuid: () => crypto.randomUUID()
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: key => scriptProperties.get(key) || null,
      setProperty: (key, value) => scriptProperties.set(key, value)
    })
  },
  LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
  getDataSpreadsheet: () => ({
    getSheetByName: name => name === 'Form1_e777' ? supplierSheet : logSheet,
    insertSheet: () => (logSheet = makeSheet([]))
  }),
  SpreadsheetApp: { flush: () => {} },
  MailApp: {
    getRemainingDailyQuota: () => mailQuota,
    sendEmail: mail => adminSent.push(mail)
  }
};
vm.createContext(adminContext);
const adminStart = source.indexOf('function isValidEmailAddress');
const adminEnd = source.indexOf('\nfunction sendSubmissionEmail', adminStart);
vm.runInContext(source.slice(adminStart, adminEnd), adminContext);
assert.strictEqual(adminContext.emailHtml('<b>"x"</b>'), '&lt;b&gt;&quot;x&quot;&lt;/b&gt;');
assert.strictEqual(adminContext.emailHttpsUrl('javascript:alert(1)'), '#');
const legacyHeader = ['Request ID', 'Created At', 'Token Fingerprint', 'Submission Key', 'Supplier Row', 'Company', 'Supplier Email', 'Message', 'Status', 'Completed At', 'Error'];
const legacyRow = ['old-request', 'date', 'fingerprint', 'legacy-key', 2, 'Old Co', 'old@example.com', 'Old message', 'SEND_ACCEPTED', 'date', ''];
logSheet = makeSheet([legacyHeader, legacyRow]);
adminContext.getAdminMessageLogSheet();
assert.strictEqual(logSheet.rows[0][3], 'Submission ID / Legacy Key');
assert.deepStrictEqual(logSheet.rows[1], legacyRow);
logSheet = null;
const copiedId = crypto.randomUUID();
const idHeaders = Array(82).fill('');
idHeaders[81] = 'Submission ID';
const idSheet = makeSheet([
  idHeaders,
  Array(81).fill('').concat(''),
  Array(81).fill('').concat(copiedId),
  Array(81).fill('').concat(copiedId)
]);
assert.strictEqual(adminContext.ensureSubmissionIds(idSheet), 3);
assert.strictEqual(idSheet.rows[0][81], 'Submission ID');
assert(adminContext.isValidSubmissionId(idSheet.rows[1][81]));
assert.notStrictEqual(idSheet.rows[2][81], copiedId);
assert.notStrictEqual(idSheet.rows[3][81], copiedId);
assert.notStrictEqual(idSheet.rows[2][81], idSheet.rows[3][81]);
const generatedId = idSheet.rows[1][81];
idSheet.rows[1][81] = idSheet.rows[2][81];
assert.throws(() => adminContext.findSubmissionRowById(idSheet, idSheet.rows[2][81]), /Duplicate submission ID/);
idSheet.rows[1][81] = generatedId;
const occupiedIdColumn = makeSheet([Array(82).fill(''), Array(81).fill('').concat('existing value')]);
assert.throws(() => adminContext.ensureSubmissionIds(occupiedIdColumn), /contains data/);

const duplicateToken = adminContext.createAdminMessageToken_(3, secondSubmissionId);
[supplierSheet.rows[1], supplierSheet.rows[2]] = [supplierSheet.rows[2], supplierSheet.rows[1]];
const duplicateSubmission = adminContext.findAdminMessageSubmission(duplicateToken);
assert.strictEqual(duplicateSubmission.submissionId, secondSubmissionId);
assert.strictEqual(duplicateSubmission.rowIndex, 2);

const messageToken = adminContext.createAdminMessageToken_(2, firstSubmissionId);
assert.strictEqual(adminContext.getAdminMessageContext(messageToken).ok, true);
assert.strictEqual(adminContext.getAdminMessageContext(messageToken + 'x').ok, false);
const caseSheets = new Map();
adminContext.getDataSpreadsheet = () => ({
  getSheetByName: name => name === 'Form1_e777' ? supplierSheet : name === 'Admin_Message_Log' ? logSheet : caseSheets.get(name),
  insertSheet: name => { const sheet = makeSheet([]); caseSheets.set(name, sheet); return sheet; }
});
adminContext.NOTIFY_EMAILS = ['admin@example.com', 'SUPPLIER@example.com'];
adminContext.ScriptApp = { getService: () => ({ getUrl: () => 'https://script.google.com/test/exec' }) };
adminContext.validateUploadPayload = () => {};
adminContext.MAX_UPLOAD_FILE_BYTES = 5 * 1024 * 1024;
const privateFiles = new Map();
adminContext.Utilities.base64Decode = data => Buffer.from(data, 'base64');
adminContext.Utilities.base64Encode = data => Buffer.from(data).toString('base64');
adminContext.Utilities.newBlob = (bytes, _type, name) => ({ getBytes: () => bytes, getName: () => name });
const privateFolder = {
  getId: () => 'private-folder',
  createFile: blob => {
    const id = crypto.randomUUID();
    const file = { getId: () => id, getBlob: () => blob, getParents: () => {
      let available = true;
      return { hasNext: () => available, next: () => { available = false; return privateFolder; } };
    } };
    privateFiles.set(id, file);
    return file;
  }
};
adminContext.DriveApp = { createFolder: () => privateFolder, getFolderById: () => privateFolder, getFileById: id => privateFiles.get(id) };
vm.runInContext(fs.readFileSync('Conversation.gs', 'utf8'), adminContext);
const openedCase = adminContext.getCaseContext(messageToken);
assert.strictEqual(openedCase.ok, true);
assert.strictEqual(openedCase.role, 'admin');
assert.strictEqual(adminContext.getCaseContext(messageToken).caseId, openedCase.caseId);
const adminCaseToken = openedCase.token;
const supplierCaseToken = adminContext.caseToken_(openedCase.caseId, 'supplier');
assert.strictEqual(adminContext.getCaseContext(supplierCaseToken).role, 'supplier');
assert.strictEqual(adminContext.getCaseContext(supplierCaseToken.replace('.supplier.', '.admin.')).ok, false);
const caseRequest = 'request_1234567890';
assert.strictEqual(adminContext.sendAdminMessage(messageToken, 'Please reply here.', caseRequest).ok, true);
assert.strictEqual(adminSent.length, 1);
assert.strictEqual(adminSent[0].to, 'supplier@example.com');
assert(!adminSent[0].body.includes(adminCaseToken));
assert(adminSent[0].body.includes(encodeURIComponent(supplierCaseToken)));
assert.strictEqual(adminContext.postCaseMessage(adminCaseToken, 'Please reply here.', [], caseRequest).duplicate, true);
assert.strictEqual(adminSent.length, 1);
assert.strictEqual(adminContext.postCaseMessage(adminCaseToken, 'Changed text', [], caseRequest).ok, false);
const attachment = { name: 'reply.pdf', data: Buffer.from('%PDF-document').toString('base64') };
assert.strictEqual(adminContext.postCaseMessage(supplierCaseToken, 'Supplier reply', [attachment], 'supplier_request_0001').ok, true);
assert.strictEqual(adminSent[1].to, 'admin@example.com');
assert(adminSent[1].htmlBody.includes('Supplier reply'));
assert(adminSent[1].htmlBody.includes('Please reply here.'));
assert.strictEqual(adminSent[1].attachments.length, 1);
const safeEmail = adminContext.caseEmail_({ id: 'test-case', values: ['', '', '<script>company</script>'] }, [
  { role: 'supplier', text: '<img src=x onerror=alert(1)>', files: [{ name: '<file>.pdf' }] }
], 'https://example.com/?token=a&b=c');
assert(!safeEmail.html.includes('<script>'));
assert(!safeEmail.html.includes('<img src=x'));
assert(safeEmail.html.includes('&lt;file&gt;.pdf'));
assert(safeEmail.html.includes('token=a&amp;b=c'));
assert(adminSent[1].body.includes(encodeURIComponent(adminCaseToken)));
const history = adminContext.getCaseContext(supplierCaseToken);
assert.strictEqual(history.messages.length, 2);
assert.strictEqual(history.messages[1].role, 'supplier');
assert.strictEqual(history.messages[1].notification, '');
const attachmentId = history.messages[1].files[0].id;
assert.strictEqual(adminContext.downloadCaseAttachment(adminCaseToken, attachmentId).ok, true);
assert.strictEqual(adminContext.downloadCaseAttachment(supplierCaseToken, 'other-file').ok, false);
const otherCase = adminContext.getCaseContext(duplicateToken);
assert.strictEqual(adminContext.downloadCaseAttachment(otherCase.token, attachmentId).ok, false);
assert.strictEqual(adminContext.closeSupplierCase(supplierCaseToken).ok, false);
mailQuota = 0;
assert.strictEqual(adminContext.postCaseMessage(adminCaseToken, 'Saved without email', [], 'request_without_mail').notification, 'FAILED');
assert.strictEqual(adminContext.getCaseContext(adminCaseToken).messages.length, 3);
mailQuota = 10;
assert.strictEqual(adminContext.closeSupplierCase(adminCaseToken).ok, true);
assert.strictEqual(adminContext.closeSupplierCase(adminCaseToken).ok, true);
assert.strictEqual(adminContext.getCaseContext(supplierCaseToken).status, 'CLOSED');
assert.strictEqual(adminContext.postCaseMessage(supplierCaseToken, 'Late reply', [], 'request_after_close').ok, false);
assert.strictEqual(adminContext.sendAdminMessage(messageToken, 'Old endpoint cannot bypass close', 'request_legacy_close').ok, false);
assert.strictEqual(adminContext.downloadCaseAttachment(supplierCaseToken, attachmentId).ok, true);
assert.strictEqual(adminContext.postCaseMessage(adminCaseToken, 'Please reply here.', [], caseRequest).duplicate, true);
assert.strictEqual(adminContext.postCaseMessage(otherCase.token, 'Other case', [], caseRequest).ok, true);

const validationStart = source.indexOf('function validateUploadPayload');
const validationEnd = source.indexOf('\nfunction isFolderInSubmissionRoot', validationStart);
const validationContext = {
  ALLOWED_UPLOAD_KEYS: new Set(['bookBank']),
  ALLOWED_UPLOAD_EXTENSIONS: new Set(['pdf']),
  MAX_FILES_PER_FIELD: 10,
  MAX_UPLOAD_FILE_BYTES: 5 * 1024 * 1024,
  MAX_UPLOAD_FILES: 25,
  MAX_UPLOAD_TOTAL_BYTES: 50 * 1024 * 1024,
  Utilities: { base64Decode: value => Buffer.from(value, 'base64') }
};
vm.createContext(validationContext);
vm.runInContext(source.slice(validationStart, validationEnd), validationContext);
const pdfBytes = Buffer.alloc(100);
pdfBytes.write('%PDF-');
assert.doesNotThrow(() => validationContext.validateUploadPayload({
  bookBank: { name: 'bank.pdf', data: pdfBytes.toString('base64') }
}));
assert.throws(() => validationContext.validateUploadPayload({
  bookBank: { name: 'malware.exe', data: Buffer.alloc(100).toString('base64') }
}), /Unsupported file type/);
assert.throws(() => validationContext.validateUploadPayload({
  bookBank: { name: 'fake.pdf', data: Buffer.alloc(100).toString('base64') }
}), /does not match/);

const queueStart = source.indexOf('function processPendingAIQueue');
const queueEnd = source.indexOf('\nfunction cleanUpQueueTriggers', queueStart);
const queueHeader = Array(82).fill('');
const queueRowA = Array(82).fill('');
const queueRowB = Array(82).fill('');
const queueIdA = crypto.randomUUID();
const queueIdB = crypto.randomUUID();
queueRowA[66] = '⏳ รอคิวตรวจ AI';
queueRowA[81] = queueIdA;
queueRowB[66] = '⏳ รอคิวตรวจ AI';
queueRowB[81] = queueIdB;
const queueSheet = makeSheet([queueHeader, queueRowA, queueRowB]);
const queuedCalls = [];
const queueContext = {
  SHEET_NAME: 'Form1_e777',
  SUBMISSION_ID_COLUMN: 82,
  AI_PENDING_STATUS: '⏳ รอคิวตรวจ AI',
  getDataSpreadsheet: () => ({ getSheetByName: () => queueSheet }),
  ensureSubmissionIds: () => {},
  LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
  runAIOnSpecificRow: (rowIndex, interactive, submissionId) => {
    queuedCalls.push({ rowIndex, interactive, submissionId });
    if (queuedCalls.length === 1) [queueSheet.rows[1], queueSheet.rows[2]] = [queueSheet.rows[2], queueSheet.rows[1]];
    return { status: 'ok' };
  },
  cleanUpQueueTriggers: () => {},
  ScriptApp: { newTrigger: () => ({ timeBased: () => ({ after: () => ({ create: () => {} }) }) }) },
  Logger: { log: () => {} }
};
vm.createContext(queueContext);
vm.runInContext(source.slice(queueStart, queueEnd), queueContext);
assert.strictEqual(queueContext.processPendingAIQueue(), 2);
assert.deepStrictEqual(queuedCalls.map(call => call.submissionId), [queueIdA, queueIdB]);
assert(queuedCalls.every(call => call.rowIndex === 0 && call.interactive === false));

const doPostStart = source.indexOf('function doPost');
const doPostEnd = source.indexOf('\nfunction jsonOut', doPostStart);
let queueCalls = 0;
const submitId = crypto.randomUUID();
let generatedSubmissionId = '';
let savedCurrency = '';
const sheet = makeSheet([Array(81).fill('')]);
const folder = {
  getUrl: () => 'https://drive.example/folder',
  createFile: () => ({
    getId: () => 'pdf-id',
    setSharing: () => {}
  })
};
const submitContext = {
  SHEET_NAME: 'Form1_e777',
  SUBMISSION_ID_COLUMN: 82,
  AI_PENDING_STATUS: '⏳ รอคิวตรวจ AI',
  getDataSpreadsheet: () => ({ getSheetByName: () => sheet }),
  normalizeCurrency: currencyContext.normalizeCurrency,
  validateUploadPayload: () => {},
  getSubmissionFolder: () => folder,
  buildInitialColumns: () => [data => { savedCurrency = data.answers.currency; return data.answers.companyName; }],
  ensureSubmissionIds: () => {},
  findSubmissionRowById: () => 2,
  generateFM021PdfBlob: (_companyName, _answers, submissionId) => {
    generatedSubmissionId = submissionId;
    return {};
  },
  isValidEmailAddress: email => /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(email),
  scheduleBackgroundAIQueue: () => { queueCalls++; return true; },
  jsonOut: value => value,
  Utilities: { getUuid: () => submitId },
  LockService: { getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} }) },
  SpreadsheetApp: { flush: () => {} },
  DriveApp: { Access: { ANYONE_WITH_LINK: 1 }, Permission: { VIEW: 1 } },
  Logger: { log: () => {} }
};
vm.createContext(submitContext);
vm.runInContext(source.slice(doPostStart, doPostEnd), submitContext);
const submitResult = submitContext.doPost({ postData: { contents: JSON.stringify({
  answers: { companyName: 'ACME', taxId: '123', email: 'buyer@example.com', currency: 'TBH' },
  files: {}
}) } });
assert.strictEqual(submitResult.status, 'ok');
assert.strictEqual(submitResult.submissionId, submitId);
assert.strictEqual(generatedSubmissionId, submitId);
assert.strictEqual(sheet.rows[1][0], 'ACME');
assert.strictEqual(sheet.rows[1][81], submitId);
assert.strictEqual(savedCurrency, 'THB');
assert.strictEqual(queueCalls, 1);
assert.strictEqual(submitContext.doPost({ postData: { contents: '{}' } }).status, 'error');
assert.strictEqual(submitContext.doPost({ postData: { contents: JSON.stringify({
  answers: { companyName: 'ACME', taxId: '123', email: 'buyer@example.com' },
  files: {}
}) } }).status, 'error');
assert.strictEqual(submitContext.doPost({ postData: { contents: JSON.stringify({
  answers: { companyName: 'ACME', taxId: '123', email: 'not-an-email', currency: 'THB' },
  files: {}
}) } }).status, 'error');

// Multi-file fields must never point at the shared submission folder.
const uploadTargets = [];
folder.createFolder = key => ({ getUrl: () => 'https://drive.google.com/drive/folders/' + key });
submitContext.saveFileToDrive = (target, file) => {
  uploadTargets.push(target.getUrl());
  return { blob: {}, url: target.getUrl() + '/' + file.name };
};
let savedFileUrls;
submitContext.buildInitialColumns = () => [data => { savedFileUrls = data.fileUrls; return data.answers.companyName; }];
const multiResult = submitContext.doPost({ postData: { contents: JSON.stringify({
  answers: { companyName: 'Test', taxId: '123', email: 'test@example.com', currency: 'THB / บาทไทย' },
  files: { docIdCard: [{ name: 'front.pdf' }, { name: 'back.pdf' }], bookBank: [{ name: 'bank.pdf' }] }
}) } });
assert.strictEqual(multiResult.status, 'ok');
assert.notStrictEqual(savedFileUrls.docIdCard, savedFileUrls.bookBank);
assert.deepStrictEqual(uploadTargets, [savedFileUrls.docIdCard, savedFileUrls.docIdCard, savedFileUrls.bookBank]);

const esgContext = {};
vm.createContext(esgContext);
vm.runInContext(source.slice(source.indexOf('function esgIdentityInstruction_'), source.indexOf('function evaluateDocumentRulesFromSheet')), esgContext);
vm.runInContext(brainSource, esgContext);
for (const enforce of [esgContext.enforceLocalDeterministicAudit, esgContext.BrainEngine.enforceDeterministicAudit]) {
 for (const docKey of ['docESG', 'docESGInd', 'docCompanyProfile', 'docCompanyProfileInd']) {
  const rule = { docKey };
  const expected = { companyName: 'บริษัท คิวทีซี พาวเวอร์ ทดสอบ จำกัด (มหาชน)' };
  const sample = { passed: true, isBlankOrEmpty: false, extractedDocumentType: 'ESG Self Assessment',
    extractedCompanyName: 'บริษัท ริงเลฟเวิล จำกัด', supplierNameEvidence: 'ชื่อบริษัทผู้ส่งมอบ: บริษัท ริงเลฟเวิล จำกัด',
    verbatimSnippet: 'แบบประเมินตนเองด้าน ESG', note: 'กรอกครบทุกข้อ' };
  assert.strictEqual(enforce({ ...sample }, rule, expected).passed, false);
  assert.strictEqual(enforce({ ...sample, extractedCompanyName: 'NOT_FOUND', supplierNameEvidence: '' }, rule, expected).passed, false);
  assert.strictEqual(enforce({ ...sample, extractedCompanyName: 'บริษัท คิวทีซี เอนเนอร์ยี่ จำกัด (มหาชน)', supplierNameEvidence: 'QTC template' }, rule, expected).passed, false);
  const sameCompany = { companyName: 'บจก. ริงเลฟเวิล' };
  assert.strictEqual(enforce({ ...sample }, rule, sameCompany).passed, true);
  assert.strictEqual(enforce({ ...sample, passed: false, note: 'ข้อมูล ESG ไม่ครบ' }, rule, sameCompany).passed, false);
  assert.strictEqual(enforce({ ...sample, supplierNameEvidence: '' }, rule, sameCompany).passed, false);
  assert.strictEqual(enforce({ ...sample }, { docKey: 'docESGInd' }, expected).passed, false);
  assert.strictEqual(enforce({ ...sample }, rule, { companyName: 'Ring Level Co., Ltd.' }).passed, false);
 }
}
auditRules = [{ order: 10, docKey: 'docCompanyProfile', isMandatory: false }];
auditResult = { passed: false, identityRejected: true, note: 'Wrong supplier' };
assert.strictEqual(auditContext.evaluateDocumentRulesFromSheet({ answers: completeAnswers }, { docCompanyProfile: {} }, 'test').overallStatus, 'FAILED');
const perFileContext = { BrainEngine: { auditDocument: blob => blob.result } };
vm.createContext(perFileContext);
vm.runInContext(source.slice(source.indexOf('function verifySingleDocWithGemini'), source.indexOf('function esgIdentityInstruction_')), perFileContext);
for (const docKey of ['docESG', 'docCompanyProfile']) {
  const mixed = perFileContext.verifySingleDocWithGemini([{ result: { passed: true } }, { result: { passed: false, identityRejected: true } }], { docKey }, {});
  assert.strictEqual(mixed.passed, false);
  assert.strictEqual(mixed.identityRejected, true);
}
const evidence = '1. ชื่อผู้ส่งมอบ: คาดหวัง QTC; พบ Ring Level; ไฟล์ 1 หน้า 2 ช่องชื่อบริษัท; ชื่อไม่ตรง\nแก้ไข: ส่งเอกสารของผู้ส่งมอบในฟอร์ม';
const explained = esgContext.appendAuditEvidence_({ passed: false, identityRejected: true, note: 'ไม่ผ่าน: คนละบริษัท' }, evidence);
assert.strictEqual(explained.passed, false);
assert.strictEqual(explained.identityRejected, true);
assert(explained.note.startsWith('ไม่ผ่าน: คนละบริษัท'));
assert(explained.note.includes(evidence));
assert.strictEqual(esgContext.appendAuditEvidence_({ passed: null, note: 'API unavailable' }, '').note, 'API unavailable');
assert.strictEqual(esgContext.appendAuditEvidence_(null, evidence), null);
assert(esgContext.auditEvidenceInstruction_().includes('EVERY applicable rule criterion'));
for (const auditSource of [source, brainSource]) {
  assert(auditSource.includes("'note', 'evidenceReport']"));
  assert(auditSource.includes('maxOutputTokens: 4096'));
  assert(auditSource.includes('auditEvidenceInstruction_()'));
  assert(auditSource.includes('parsed.evidenceReport'));
}
assert(source.includes("emailHtml(item.note || '-').replace(/\\n/g, '<br>')"));
console.log('All checks passed');
