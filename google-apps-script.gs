// Corridor's Plus — Google Sheets Sync + Receipt Upload
// Paste this entire file into: Google Sheet → Extensions → Apps Script
// Then: run setup() once, then deploy as a Web App (Deploy → New deployment → Web app → Anyone)

const SPREADSHEET_ID     = '1u13Ai9uvKcau_cB4NjubcFnHkn2xXnsW3UmduOgbA0c';
const MASTER_SHEET       = 'Materials List';
const RECEIPTS_FOLDER_ID = '19TGCyD8v0j0UyahztXXuwIEGLVyk3LVJ';

const HEADERS = [
  'App ID', 'Trade', 'Material', 'Brand/Type', 'Unit',
  'On Hand', 'Min Stock',
  'Last Delivery Date', 'Last Delivery Qty', 'Ordered By',
  'To Order', 'Order Status', 'Delivery ETA',
  'Link', 'Notes',
  'All-time Delivered', 'All-time Used'
];

// ── GET — pull all materials into the app ─────────────────────────────────────
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MASTER_SHEET);
    if (!sheet) {
      return json({ error: 'Master sheet not found — run setup() first from the Apps Script editor.' });
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return json([]);

    const headers   = data[0].map(h => String(h).trim().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' '));
    const materials = data.slice(1)
      .filter(r => String(r[1] || r[0]).trim())
      .map(r => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = (r[i] !== undefined && r[i] !== '') ? String(r[i]) : ''; });
        return obj;
      });
    return json(materials);
  } catch (err) {
    return json({ error: err.toString() });
  }
}

// ── POST — route to receipt upload or materials sync ─────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.type === 'receipt')       return uploadReceipt(payload);
    if (payload.type === 'daily_report')  return generateDailyReport(payload);
    if (payload.type === 'payroll')       return handlePayroll(payload);
    if (payload.type === 'weekly_recap')  return generateWeeklyRecap(payload);
    return syncMaterials(payload);

  } catch (err) {
    return json({ error: err.toString() });
  }
}

// ── Receipt upload — analyze with Claude then save to subfolder ───────────────
function uploadReceipt(payload) {
  try {
    if (!payload.data) return json({ error: 'No image data received.' });

    const parts    = payload.data.split(',');
    const mimeType = parts[0].split(';')[0].split(':')[1] || 'image/jpeg';
    const base64   = parts[1];

    // Ask Claude to analyze all pages of the receipt together
    const extraPages = Array.isArray(payload.extraPages) ? payload.extraPages : [];
    const analysis = analyzeReceiptWithClaude(base64, mimeType, extraPages);

    // Build filename from analysis
    const vendor   = normalizeVendor(analysis.vendor   || payload.vendor   || 'Unknown');
    const amount   = analysis.amount   || payload.amount   || null;
    const category = analysis.category || payload.category || 'Other';
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const dateStr  = analysis.date     || todayStr;
    const amtStr   = amount ? '_$' + Number(amount).toFixed(2) : '';
    const filename = `${category.replace(/\s+/g,'-')}_${dateStr}_${vendor.replace(/\s+/g,'-')}${amtStr}.jpg`;

    // Save files to Drive
    const bytes      = Utilities.base64Decode(base64);
    const blob       = Utilities.newBlob(bytes, mimeType, filename);
    const monthFolder = getOrCreateSubfolder(category, dateStr);
    let   driveUrl;

    if (extraPages.length > 0) {
      // Multi-page receipt: create a per-receipt subfolder so all pages are together
      const safeVendor    = vendor.replace(/[^a-zA-Z0-9\s]/g, '').replace(/\s+/g, '-').substring(0, 30);
      const rcFolderName  = `${safeVendor}_${dateStr}`;
      const existingRcFolders = monthFolder.getFoldersByName(rcFolderName);
      const rcFolder      = existingRcFolders.hasNext() ? existingRcFolders.next() : monthFolder.createFolder(rcFolderName);
      rcFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      const mainFile = rcFolder.createFile(blob);
      mainFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      extraPages.forEach(function(page) {
        try {
          const ep     = page.data.split(',');
          const epMime = ep[0].split(';')[0].split(':')[1] || 'image/jpeg';
          const epB64  = ep[1];
          const epBlob = Utilities.newBlob(Utilities.base64Decode(epB64), epMime, page.filename || 'page.jpg');
          const epFile = rcFolder.createFile(epBlob);
          epFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch(e) {}
      });

      driveUrl = rcFolder.getUrl(); // link to folder so all pages are visible
    } else {
      // Single-page: upload directly to month folder
      const file = monthFolder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      driveUrl = file.getUrl();
    }

    // Log to the Receipts sheet
    const items = Array.isArray(analysis.items) ? analysis.items : [];
    logReceiptToSheet({ date: dateStr, vendor, category, amount, url: driveUrl, items });

    return json({
      success:  true,
      url:      driveUrl,
      vendor,
      amount,
      category,
      date:     dateStr,
    });
  } catch (err) {
    return json({ error: err.toString() });
  }
}

// ── Call Claude API to read all pages of the receipt ─────────────────────────
// Maps known name variants to a single canonical vendor name
const VENDOR_ALIASES = {
  'janovic paint & decorating center': 'Janovic Paint & Decorating',
  'janovic':                           'Janovic Paint & Decorating',
};

function normalizeVendor(raw) {
  if (!raw) return 'Unknown';
  let v = raw.trim();
  // Strip leading "The "
  v = v.replace(/^The\s+/i, '');
  // Strip trailing store/location numbers like "#1234" or "Store 1234"
  v = v.replace(/\s+(?:Store\s*)?\#\d+$/i, '').trim();
  // Apply alias map
  const alias = VENDOR_ALIASES[v.toLowerCase()];
  if (alias) v = alias;
  return v;
}

function generateDailyReport(payload) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('RECEIPT_API_KEY');
    if (!apiKey) return json({ error: 'API key not set in Script Properties.' });

    const { date, projectName, crew, floors, notes, defaultItemOrder } = payload;

    // Build context for Claude
    let ctx = `Project: ${projectName}\nDate: ${date}\n\n`;

    ctx += `CREW ON SITE (${crew.length}):\n`;
    crew.forEach(function(w) { ctx += `- ${w.name} (${w.trade || 'unspecified'})\n`; });

    ctx += `\nWORK LOGGED TODAY:\n`;
    floors.forEach(function(f) {
      if (!f.entries || f.entries.length === 0) return;
      ctx += `\n${f.name}:\n`;
      f.entries.forEach(function(e) {
        const parts = [];
        if (e.skilled > 0)  parts.push(`${e.skilled} skilled × ${e.skilledHours} hrs`);
        if (e.laborer > 0)  parts.push(`${e.laborer} laborer × ${e.laborerHours} hrs`);
        ctx += `  - ${e.itemName} [${e.status}]: ${parts.join(', ')} = ${e.manHours} man-hrs`;
        if (e.workerNames && e.workerNames.length) ctx += ` (${e.workerNames.join(', ')})`;
        if (e.note) ctx += ` — "${e.note}"`;
        ctx += '\n';
      });
    });

    ctx += `\nITEMS STILL IN PROGRESS:\n`;
    var anyInProgress = false;
    floors.forEach(function(f) {
      const ip = (f.allItems || []).filter(function(i) { return i.status === 'active'; });
      if (ip.length) { ctx += `  ${f.name}: ${ip.map(function(i){return i.name;}).join(', ')}\n`; anyInProgress = true; }
    });
    if (!anyInProgress) ctx += '  None\n';

    ctx += `\nSTANDARD WORK SEQUENCE: ${(defaultItemOrder||[]).join(' → ')}\n`;
    if (notes) ctx += `\nFOREMAN NOTES:\n${notes}\n`;

    const prompt = `You are writing sections of a daily field report for a NYC hallway contractor. Use the data below.

${ctx}

Respond with ONLY a JSON object in this exact shape — no markdown, no explanation:
{"summary":"...","nextDayPlan":"..."}

Rules:
- summary: 2-3 sentences describing what was accomplished today. Incorporate any foreman notes naturally. Professional tone, past tense.
- nextDayPlan: 2-3 sentences on what should happen tomorrow, based on what is still in progress and the standard work sequence. Be specific about floors and tasks.
- Do not repeat crew names or man-hour numbers (those are shown separately in the report).
- Plain paragraph prose, no bullet points.`;

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      payload: JSON.stringify({ model: 'claude-haiku-4-5-20251001', max_tokens: 512, messages: [{ role: 'user', content: prompt }] }),
      muteHttpExceptions: true,
    });

    const result  = JSON.parse(response.getContentText());
    const text    = result?.content?.[0]?.text || '{}';
    const clean   = text.replace(/```json|```/g, '').trim();
    const report  = JSON.parse(clean);
    return json({ success: true, report });
  } catch(err) {
    return json({ error: err.toString() });
  }
}

// ── Weekly recap — summarize the week with Claude ────────────────────────────
function generateWeeklyRecap(payload) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('RECEIPT_API_KEY');
    if (!apiKey) return json({ error: 'API key not configured.' });

    const d = payload.weekData || {};

    const prompt = `You are writing a weekly recap paragraph for a NYC hallway renovation contractor named Corridor's Plus, working at Georgetown Plaza.

WEEK: ${d.label || ''}

LABOR LOGGED:
${d.labor || 'None'}

ATTENDANCE:
${d.attendance || 'None'}

DAILY NOTES:
${d.notes || 'None'}

FLOOR & ITEM STATUS:
${d.floorStatus || 'None'}

Write exactly 3 sentences — no more — like a straight-talking job site supervisor giving a quick verbal update to the owner. First sentence: what got done across the floors. Second sentence: how the crew performed and any issues or setbacks from the notes. Third sentence: what needs to happen next week.

STRICT RULES:
- Plain sentences only. No markdown. No pound signs, asterisks, underscores, or backticks.
- No headers, titles, bullet points, or numbered lists.
- Do not open with the company name, project name, or the week date.
- Do not repeat raw numbers already in the tables.
- Past tense throughout.`;

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      payload: JSON.stringify({
        model: 'claude-haiku-4-5-20251001',
        max_tokens: 250,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true,
    });

    const result = JSON.parse(response.getContentText());
    const recap  = result?.content?.[0]?.text || '';
    return json({ recap });
  } catch(err) {
    return json({ error: err.toString() });
  }
}

function analyzeReceiptWithClaude(base64, mimeType, extraPages) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('RECEIPT_API_KEY');
    if (!apiKey) return {};

    // Build content array: one image block per page, then the prompt
    const content = [];
    content.push({ type: 'image', source: { type: 'base64', media_type: mimeType, data: base64 } });
    (extraPages || []).forEach(function(page) {
      try {
        const ep     = page.data.split(',');
        const epMime = ep[0].split(';')[0].split(':')[1] || 'image/jpeg';
        const epB64  = ep[1];
        content.push({ type: 'image', source: { type: 'base64', media_type: epMime, data: epB64 } });
      } catch(e) {}
    });
    content.push({
      type: 'text',
      text: `You are reading a contractor's receipt or invoice. There may be multiple pages — treat them as one document. Extract the key details and respond ONLY with a JSON object — no explanation, no markdown.

Return exactly this shape:
{
  "vendor": "store or company name",
  "date": "YYYY-MM-DD",
  "amount": 123.45,
  "category": "one of: Materials & Supplies | Tools & Equipment | Subcontractors | Food & Meals | Other",
  "items": ["qty x item name — $price", "qty x item name — $price"]
}

Rules:
- vendor: copy the exact business name from the top of the receipt (logo/header). Include the full name as printed, e.g. "M&D Shapiro True Value Hardware" not just "True Value Hardware". Strip leading "The " (so "The Home Depot" → "Home Depot"). Strip trailing store numbers or "#1234".
- date: the date printed on the receipt in YYYY-MM-DD format. null if not visible.
- amount: the grand total paid across all pages as a number, no $ sign. null if not visible.
- category: pick the best match from the five options above based on what was purchased.
- items: list each line item as "qty x description — $price". Max 20 items across all pages. Empty array if not readable.`
    });

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:  'post',
      headers: {
        'x-api-key':         apiKey,
        'anthropic-version': '2023-06-01',
        'content-type':      'application/json',
      },
      payload: JSON.stringify({
        model:      'claude-haiku-4-5-20251001',
        max_tokens: 1024,
        messages:   [{ role: 'user', content: content }],
      }),
      muteHttpExceptions: true,
    });

    const result = JSON.parse(response.getContentText());
    const text   = result?.content?.[0]?.text || '{}';
    // Strip any accidental markdown fences
    const clean  = text.replace(/```json|```/g, '').trim();
    return JSON.parse(clean);
  } catch (err) {
    Logger.log('Claude analysis failed: ' + err.toString());
    return {};
  }
}

// ── Get or create category → month subfolder inside the receipts folder ───────
function getOrCreateSubfolder(category, dateStr) {
  const root      = DriveApp.getFolderById(RECEIPTS_FOLDER_ID);
  const dateObj   = dateStr ? new Date(dateStr + 'T12:00:00') : new Date();
  const monthName = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM MMMM');

  // category folder
  const catFolders = root.getFoldersByName(category);
  const catFolder  = catFolders.hasNext() ? catFolders.next() : root.createFolder(category);
  catFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // month folder inside category
  const monFolders = catFolder.getFoldersByName(monthName);
  const monFolder  = monFolders.hasNext() ? monFolders.next() : catFolder.createFolder(monthName);
  monFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return monFolder;
}

// ── Log a receipt row to the Receipts sheet, grouped by vendor ────────────────
const RECEIPT_SHEET   = 'Receipts';
const RECEIPT_HEADERS = ['Date', 'Vendor', 'Category', 'Amount', 'Drive Link'];

// Vendor color palette — assigned dynamically as new vendors appear
const VENDOR_PALETTE = [
  '#dbeafe', '#fce7f3', '#fef3c7', '#ccfbf1', '#ede9fe',
  '#ffedd5', '#dcfce7', '#fef2f2', '#e0f2fe', '#fdf4ff',
];


function getVendorColor(vendor) {
  const props     = PropertiesService.getScriptProperties();
  const mapJson   = props.getProperty('VENDOR_COLOR_MAP') || '{}';
  const map       = JSON.parse(mapJson);
  const key       = vendor.trim().toLowerCase();
  if (!map[key]) {
    const usedCount = Object.keys(map).length;
    map[key] = VENDOR_PALETTE[usedCount % VENDOR_PALETTE.length];
    props.setProperty('VENDOR_COLOR_MAP', JSON.stringify(map));
  }
  return map[key];
}

function logReceiptToSheet({ date, vendor, category, amount, url, items }) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(RECEIPT_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(RECEIPT_SHEET);
    setupReceiptSheet(sheet);
  }

  // Ensure the header row matches RECEIPT_HEADERS
  const currentHdr = sheet.getRange(1, 1, 1, RECEIPT_HEADERS.length).getValues()[0];
  const hdrMatch   = RECEIPT_HEADERS.every(function(h, i) { return String(currentHdr[i] || '').trim() === h; });
  if (!hdrMatch) {
    const hdrRange = sheet.getRange(1, 1, 1, RECEIPT_HEADERS.length);
    hdrRange.setValues([RECEIPT_HEADERS]);
    hdrRange.setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff').setFontSize(12);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 180);
    sheet.setColumnWidth(4, 100);
    sheet.setColumnWidth(5, 130);
  }

  const data       = sheet.getDataRange().getValues();
  const numRows    = data.length;
  const vendorCol  = 1;
  const dateCol    = 0;
  const bgColor    = getVendorColor(vendor);
  const vendorKey  = vendor.trim().toLowerCase();

  // Find the correct chronological insertion point within this vendor's rows.
  const tz2 = Session.getScriptTimeZone();
  const toDateStr = function(val) {
    return val instanceof Date ? Utilities.formatDate(val, tz2, 'yyyy-MM-dd') : String(val || '');
  };
  let insertAfter = -1;
  for (let i = 1; i < numRows; i++) {
    const rowVendor = String(data[i][vendorCol] || '').trim().toLowerCase();
    const rowDate   = toDateStr(data[i][dateCol]);
    const isTotal   = rowDate.startsWith('TOTAL:');
    if (rowVendor === vendorKey && !isTotal && rowDate <= date) {
      insertAfter = i + 1;
    }
  }
  if (insertAfter === -1) {
    for (let i = 1; i < numRows; i++) {
      const rowVendor = String(data[i][vendorCol] || '').trim().toLowerCase();
      const rowDate   = toDateStr(data[i][dateCol]);
      const isTotal   = rowDate.startsWith('TOTAL:');
      if (rowVendor === vendorKey && !isTotal) {
        insertAfter = i;
        break;
      }
    }
  }

  const dateObj = new Date(date + 'T12:00:00'); // store as Date so Sheets can format it
  const newRow = [dateObj, vendor, category, amount ? Number(amount) : '', url];
  let   range;

  if (insertAfter > 0) {
    sheet.insertRowAfter(insertAfter);
    range = sheet.getRange(insertAfter + 1, 1, 1, RECEIPT_HEADERS.length);
    range.setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
    const lastRow = sheet.getLastRow();
    range = sheet.getRange(lastRow, 1, 1, RECEIPT_HEADERS.length);
  }

  styleReceiptRow(range, bgColor, items);
  rebuildVendorMonthlyTotals(sheet, vendor);
  ensureReceiptFilter(sheet);
}

// ── Rebuild monthly subtotal rows for a specific vendor ───────────────────────
function rebuildVendorMonthlyTotals(sheet, vendor) {
  const vendorKey = vendor.trim().toLowerCase();

  // Remove existing subtotal rows for this vendor first
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    const rowVendor = String(data[i][1] || '').trim().toLowerCase();
    const isTotal   = String(data[i][0] || '').startsWith('TOTAL:');
    if (isTotal && rowVendor === vendorKey) {
      sheet.deleteRow(i + 1);
    }
  }

  // Re-read after deletions
  const fresh     = sheet.getDataRange().getValues();
  const dateCol   = 0;
  const vendorCol = 1;
  const amountCol = 3;

  // Collect all receipt rows for this vendor with their sheet row index
  const tz = Session.getScriptTimeZone();
  const vendorRows = [];
  for (let i = 1; i < fresh.length; i++) {
    const v      = String(fresh[i][vendorCol] || '').trim().toLowerCase();
    const rawD   = fresh[i][dateCol];
    const d      = rawD instanceof Date ? Utilities.formatDate(rawD, tz, 'yyyy-MM-dd') : String(rawD || '');
    if (v === vendorKey && !d.startsWith('TOTAL:')) {
      vendorRows.push({ sheetRow: i + 1, date: d, amount: parseFloat(fresh[i][amountCol]) || 0 });
    }
  }
  if (vendorRows.length === 0) return;

  // Group by YYYY-MM
  const byMonth = {};
  vendorRows.forEach(r => {
    const month = r.date.substring(0, 7); // "2026-05"
    if (!byMonth[month]) byMonth[month] = [];
    byMonth[month].push(r);
  });

  // Insert a subtotal row after each month's last row (process in reverse to keep row numbers valid)
  const months = Object.keys(byMonth).sort().reverse();
  months.forEach(month => {
    const rows      = byMonth[month];
    const lastRow   = Math.max(...rows.map(r => r.sheetRow));
    const total     = rows.reduce((s, r) => s + r.amount, 0);
    const label     = Utilities.formatDate(new Date(month + '-15'), Session.getScriptTimeZone(), 'MMMM yyyy');
    sheet.insertRowAfter(lastRow);
    const range = sheet.getRange(lastRow + 1, 1, 1, RECEIPT_HEADERS.length);
    range.setValues([[`TOTAL: ${label}`, vendor, '', total, '']]);
    range.setBackground('#1a3a5c');
    range.setFontColor('#ffffff');
    range.setFontWeight('bold');
    range.setFontSize(11);
    range.getCell(1, 4).setNumberFormat('$#,##0.00');
  });
}

function styleReceiptRow(range, bgColor, items) {
  range.setBackground(bgColor);
  range.setFontColor('#111111');
  range.setFontSize(11);
  range.setVerticalAlignment('middle');
  range.getCell(1, 1).setNumberFormat('mmm d, yyyy'); // e.g. "May 8, 2026"
  range.getCell(1, 4).setNumberFormat('$#,##0.00');

  // Add line items as a hover note on the Amount cell
  if (items && items.length > 0) {
    const note = 'Items purchased:\n' + items.map(i => '• ' + i).join('\n');
    range.getCell(1, 4).setNote(note);
  }

  // Make Drive Link clickable (column 5)
  const linkCell = range.getCell(1, 5);
  const url      = linkCell.getValue();
  if (url) {
    linkCell.setFormula(`=HYPERLINK("${url}","View Receipt")`);
    linkCell.setFontColor('#1a73e8');
  }
}

function ensureReceiptFilter(sheet) {
  if (!sheet.getFilter()) {
    sheet.getRange(1, 1, sheet.getLastRow(), RECEIPT_HEADERS.length).createFilter();
  }
}


function setupReceiptSheet(sheet) {
  // Header row
  sheet.appendRow(RECEIPT_HEADERS);
  const hdr = sheet.getRange(1, 1, 1, RECEIPT_HEADERS.length);
  hdr.setFontWeight('bold');
  hdr.setBackground('#1a3a5c');
  hdr.setFontColor('#ffffff');
  hdr.setFontSize(12);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120); // Date
  sheet.setColumnWidth(2, 160); // Vendor
  sheet.setColumnWidth(3, 180); // Category
  sheet.setColumnWidth(4, 100); // Amount
  sheet.setColumnWidth(5, 130); // Drive Link
  sheet.getRange(2, 1, 1000, 1).setNumberFormat('mmm d, yyyy');
}

// ── Materials sync — update/add/delete/archive rows in Master sheet ───────────
function syncMaterials(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName(MASTER_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(MASTER_SHEET, 0);
      sheet.appendRow(HEADERS);
    }

    const allData  = sheet.getDataRange().getValues();
    const hdr      = allData[0].map(h => String(h).trim().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' '));
    const matIdx   = hdr.indexOf('Material');
    const tradeIdx = hdr.indexOf('Trade');

    if (matIdx === -1) {
      return json({ error: 'Material column not found in sheet headers. Re-run setup() or check the header row.' });
    }

    // Build three maps: App ID, trade+name (exact), and name-only (fallback)
    const idMap       = {};  // "appId" → 1-based row
    const rowMap      = {};  // "trade||name" → 1-based row
    const nameOnlyMap = {};  // "name" → 1-based row
    const appIdIdx    = hdr.indexOf('App ID');
    allData.slice(1).forEach((r, i) => {
      const name   = String(r[matIdx]   || '').trim().toLowerCase();
      const trade  = tradeIdx  >= 0 ? String(r[tradeIdx]  || '').trim().toLowerCase() : '';
      const appId  = appIdIdx  >= 0 ? String(r[appIdIdx]  || '').trim() : '';
      if (!name) return;
      if (appId) idMap[appId] = i + 2;
      rowMap[trade + '||' + name] = i + 2;
      if (!nameOnlyMap[name]) nameOnlyMap[name] = i + 2;
    });

    // Fields the app is allowed to overwrite — Notes and Delivery ETA are sheet-managed
    const APP_FIELDS = HEADERS.filter(h => h !== 'Notes' && h !== 'Delivery ETA');

    const updates = payload.materials || [];
    const newRows = [];

    updates.forEach(mat => {
      const name     = String(mat['Material'] || '').trim().toLowerCase();
      const trade    = String(mat['Trade']    || '').trim().toLowerCase();
      const appId    = String(mat['App ID']   || '').trim();
      // Match by App ID first, then trade+name, then name only
      const rowNum   = (appId && idMap[appId]) || rowMap[trade + '||' + name] || nameOnlyMap[name];
      if (rowNum) {
        APP_FIELDS.forEach(field => {
          const colIdx = hdr.indexOf(field);
          if (colIdx >= 0 && mat[field] !== undefined) {
            sheet.getRange(rowNum, colIdx + 1).setValue(mat[field]);
          }
        });
        // Backfill App ID if the row doesn't have one yet
        if (appIdIdx >= 0 && appId && !allData[rowNum - 1][appIdIdx]) {
          sheet.getRange(rowNum, appIdIdx + 1).setValue(appId);
        }
      } else {
        newRows.push(HEADERS.map(h => (mat[h] !== undefined ? mat[h] : '')));
      }
    });

    newRows.forEach(row => sheet.appendRow(row));

    // Handle archives — move rows from Master to Archived tab
    const archives = payload.archives || [];
    if (archives.length > 0) {
      const archiveExactKeys = new Set(archives.map(d => (String(d.trade||'') + '||' + String(d.name||'')).toLowerCase()));
      const archiveNameKeys  = new Set(archives.map(d => String(d.name||'').trim().toLowerCase()));
      let archSheet = ss.getSheetByName('Archived');
      if (!archSheet) {
        archSheet = ss.insertSheet('Archived');
        archSheet.appendRow(HEADERS);
        const ah = archSheet.getRange(1, 1, 1, HEADERS.length);
        ah.setFontWeight('bold'); ah.setBackground('#4a1942'); ah.setFontColor('#ffffff');
        archSheet.setFrozenRows(1);
      }
      const archHdr    = archSheet.getDataRange().getValues()[0].map(h => String(h).trim());
      const archMatIdx = archHdr.indexOf('Material');
      const freshData  = sheet.getDataRange().getValues();
      const freshHdr   = freshData[0].map(h => String(h).trim().replace(/[\r\n]+/g,' ').replace(/\s+/g,' '));
      const fTrade = freshHdr.indexOf('Trade');
      const fMat   = freshHdr.indexOf('Material');
      for (let i = freshData.length - 1; i >= 1; i--) {
        const rowName  = String(fMat   >= 0 ? freshData[i][fMat]   : '').trim().toLowerCase();
        const rowTrade = String(fTrade >= 0 ? freshData[i][fTrade] : '').trim().toLowerCase();
        const exactKey = rowTrade + '||' + rowName;
        const matched  = archiveExactKeys.has(exactKey) || archiveNameKeys.has(rowName);
        if (matched && rowName) {
          const alreadyArchived = archSheet.getDataRange().getValues().slice(1).some(r => {
            return String(archMatIdx >= 0 ? r[archMatIdx] : '').trim().toLowerCase() === rowName;
          });
          if (!alreadyArchived) {
            archSheet.appendRow(HEADERS.map(h => { const ci = freshHdr.indexOf(h); return ci >= 0 ? freshData[i][ci] : ''; }));
          }
          sheet.deleteRow(i + 1);
        }
      }
    }

    // Handle deletions — remove sheet rows for materials deleted in the app
    // Skip any item that is also present in the materials payload (re-added after deletion)
    const deletions = payload.deletions || [];
    if (deletions.length > 0) {
      const activeKeys = new Set(updates.map(m => (String(m['Trade']||'') + '||' + String(m['Material']||'')).toLowerCase()));
      const delKeys = new Set(
        deletions
          .filter(d => !activeKeys.has((String(d.trade||'') + '||' + String(d.name||'')).toLowerCase()))
          .map(d => (String(d.trade||'') + '||' + String(d.name||'')).toLowerCase())
      );
      if (delKeys.size > 0) {
        const freshData     = sheet.getDataRange().getValues();
        const freshHdrTrade = freshData[0].map(h => String(h).trim()).indexOf('Trade');
        const freshHdrMat   = freshData[0].map(h => String(h).trim()).indexOf('Material');
        for (let i = freshData.length - 1; i >= 1; i--) {
          const key = (String(freshData[i][freshHdrTrade]||'') + '||' + String(freshData[i][freshHdrMat]||'')).toLowerCase();
          if (delKeys.has(key)) sheet.deleteRow(i + 1);
        }
      }
    }

    return json({ success: true, updated: updates.length, deleted: deletions.length, archived: archives.length });
  } catch (err) {
    return json({ error: err.toString() });
  } finally {
    try { lock.releaseLock(); } catch(e) {}
  }
}

// ── Run this ONCE to create and populate the Master sheet ─────────────────────
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MASTER_SHEET);

  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(MASTER_SHEET, 0);
  }

  const data = [
    HEADERS,
    // Plaster
    ['Plaster','Sealer/Primer','Zinsser Gardz2300','5 Gallon Buckets',5,'','04/28/2025',5,'jv','','','','Before skim coat'],
    ['Plaster','Joint Compound AP','All-Purpose','5 Gallon Buckets',20,'','04/28/2025',20,'jv','','','','1 coat'],
    ['Plaster','Joint Compound LW','Light Weight','5 Gallon Buckets',30,'','04/28/2025',30,'jv','','','','1-2 coats'],
    ['Plaster','Plaster of Paris','','Bags',10,'','04/28/2025',10,'jv','','','','-'],
    ['Plaster','Primer (Post-skim) Ceiling','Valspar High Build','5 Gallon Buckets',5,'','','','','','','','After sanding, ON CEILING ONLY'],
    ['Plaster','Primer (Post-skim) Walls','Roman Pro-977 Ultra Prime','5 Gallon Buckets',5,'','05/08/2025',5,'jv','','ordered del 5/11','','After sanding, ON WALLS ONLY'],
    ['Plaster','Sandpaper','180-400 grit','Multi Pack',10,'','04/28/2025',10,'jv','','','','Purchase 180-400 Grit'],
    ['Plaster','Drywall Sanding Screen','','10 pack',5,'','04/28/2025',5,'jv','','','','-'],
    ['Plaster','Hyde Sanding Screen','','',5,'','05/08/2025',5,'jv','','ordered','','Order Amazon version'],
    ['Plaster','Roller Covers 3/4"','3/4" nap','Multi-pack',12,'','','','','','','','Primer application'],
    ['Plaster','Roller Covers 1/2"','1/2" nap','Multi-pack',12,'','','','','','','','Primer application'],
    ['Plaster','Skim Rollers','1"-1.25" nap','Each',5,'','','','','','','','For rolling compound'],
    ['Plaster','PlasterWeld','','1 Gallon',15,'','','','','','','','For repairs'],
    ['Plaster','Mesh Tape','','rolls','','','','','','','','',''],
    ['Plaster','Corner Bead','','pcs',0,'','','','','','','',''],
    // Paint
    ['Paint','Door Paint','','1 Gallon',25,'','','','','','','','Benjamin Moore Command Newburg Green HC-158 Satin Finish'],
    ['Paint','Ceiling Paint','','5 Gallon Bucket',5,'','05/08/2025',5,'jv','','ordered del 5/11','','Benjamin Moore Brilliant White OC-51 Flat Finish'],
    ['Paint','Raceway Paint','','5 Gallon Bucket',5,'','05/08/2025',5,'jv','','ordered del 5/11','','Benjamin Moore Brilliant White OC-51 Eggshell Finish'],
    ['Paint','Door and Raceway Primer','Stix','1 Gallon',15,'','','','','','','','Stix primer'],
    ['Paint','Painting Trays','','Single',8,'','04/28/2025',8,'jv','','','',''],
    ['Paint','Painting Tray Liners','','10 Pack',10,'','04/28/2025',10,'jv','','','',''],
    ['Paint','Ready Patch','','1 gallon',5,'','04/28/2025',5,'jv','','','',''],
    ['Paint','Caulk','DAP Alex Fast Dry','Box',5,'','04/28/2025',5,'jv','','','https://www.homedepot.com/p/DAP-Alex-Fast-Dry-10-1-oz-White-Acrylic-Latex-Plus-Silicone-Caulk-Contractor-12-Pack-18426/100634331','-'],
    ['Paint','Roller Covers 1/4"','1/4" Nap','Multi Pack',12,'','','','','','','','Order with door paint'],
    ['Paint','Door/Frame Paint','','gal','','','','','','','','',''],
    ['Paint','Primer','','buckets','','','','','','','','',''],
    ['Paint','Stix Primer','','gal',15,'','','','','','','',''],
    ['Paint','Gardz 2300 Primer','','5/gal Buckets',3,2,'','','',5,'','',''],
    ['Paint','Primer (Ceiling)','','buckets',5,'','','','','','','',''],
    ['Paint','Primer (Walls)','','buckets',5,'','','','','','','',''],
    // Wallcovering
    ['Wallcovering','Wallcovering','','rolls','','','','','','','','',''],
    ['Wallcovering','Wallcovering Paste','','gal','','','','','','','','',''],
    // Carpentry
    ['Carpentry','Baseboard','','pcs','','','','','','','','',''],
    ['Carpentry','Blocking','','pcs','','','','','','','','',''],
    ['Carpentry','Raceway','','pcs','','','','','','','','',''],
    // Electrical
    ['Electrical','Wire Nuts','3M','Multi Pack',3,'','','','','','Ordered','','Johnny will pick up'],
    ['Electrical','MC Cable','','Feet',500,'','','','','','Ordered','','Johnny will pick up'],
    ['Electrical','Single BX Connectors 3/8"','','',100,'','','','','','Ordered','','Johnny will pick up'],
    ['Electrical','1900 Box (shallow)','','',90,'','','','','','Ordered','','Johnny will pick up, Majic to confirm'],
    ['Electrical','1900 Box Blank Plates','','',90,'','','','','','Ordered','',''],
    ['Electrical','Octagon Boxes','','pcs',60,'','','','','','Ordered','','Johnny will pick up, Majic to confirm'],
    ['Electrical','#12 Stranded Wire (White)','','Feet',500,'','','','','','Ordered','',''],
    ['Electrical','#12 Stranded Wire (Black)','','Feet',500,'','','','','','Ordered','',''],
    ['Electrical','Ground Tails','','Pieces',100,'','','','','','Ordered','',''],
    ['Electrical','1/2" EMT Pipe','','10 Foot Lengths',40,'','','','','','Ordered','',''],
    ['Electrical','1/2" EMT Connectors','','Pieces',120,'','','','','','Ordered','',''],
    ['Electrical','1/2" EMT LR','','Pieces',30,'','','','','','Ordered','https://www.lightbulbwholesaler.com/topaz-elr1','-'],
    ['Electrical','1/2" EMT LL','','Pieces',30,'','','','','','Ordered','https://www.lightbulbwholesaler.com/topaz-ell1','-'],
    ['Electrical','Light Retrofit Kits','','units','','','','','','','','',''],
    ['Electrical','Outlets','','pcs',0,'','','','','','','',''],
    ['Electrical','Device Covers','','pcs',0,'','','','','','','',''],
    ['Electrical','Sconces','','units',0,'','','','','','','',''],
    ['Electrical','EM Devices','','pcs',0,'','','','','','','',''],
    ['Electrical','MC Connectors','','packs',100,'','','','','','','',''],
    ['Electrical','1900 Box','','pcs',30,'','','','','','','',''],
    ['Electrical','Mud Plates','','pcs',30,'','','','','','','',''],
    ['Electrical','Blank Covers','','pcs',30,'','','','','','','',''],
    // Misc
    ['Misc.','Garbage Bags','-','boxes',20,'','04/28/2025',20,'jv','','','',''],
    ["Misc.","Painter's Plastic",'-','boxes',10,'','04/28/2025',10,'jv','','','',''],
    ['Misc.','Masking Tape','3M','Rolls',45,20,'','','','','','','-'],
    ['Misc.','Frog Tape','','rolls',60,15,'','','',45,'','https://www.homedepot.com/p/FrogTape-Pro-Grade-1-41-in-x-60-yds-Blue-Painter-s-Tape-with-PaintBlock-4-Pack-104982/312886403','-'],
    ['Misc.','Drop Cloths','','packs',10,'','04/28/2025',10,'jv','','','','Size is 6X9'],
    ['Misc.','N95 Masks','3M','boxes',22,5,'','','','','','','-'],
    ['Misc.','Zipties','','boxes',20,3,'','','','','','','-'],
    ['Misc.','Construction Adhesive','PL 400','Box',3,'','04/28/2025',3,'jv','','','https://www.homedepot.com/p/Loctite-PL-400-Subfloor-28-oz-All-Weather-Latex-Construction-Adhesive-Tan-Cartridge-12-pack-2136221/300752201','-'],
    ['Misc.','Painting Rags','','boxes',3,'','04/28/2025',3,'jv','','','','-'],
    ['Misc.','Microfiber Rags','','packs',2,'','04/28/2025',2,'jv','','','','-'],
    ['Misc.','Finish Nails','','boxes',5,'','04/28/2025',5,'jv','','','','-'],
    ['Misc.','Air Filters for Scrubber','','',4,'','05/08/2025',4,'jv','','ordered','','Size is 13.5X14.5'],
    ['Misc.','Extension Cords','Ridgid 50ft','',4,'','05/03/2025',8,'jr','','','',''],
    ['Misc.','Splitters with Surge Protection','Ridgid','',2,'','05/03/2025',4,'jr','','','',''],
    ['Misc.','Bulbs','','packs',5,'','04/29/2025',5,'jv','','','','100 Watts'],
    ['Misc.','PL400','','tubes',36,'','','','','','','',''],
    ['Misc.','Ready Patch','','Gallons',5,2,'','','','','','',''],
    ['Misc.','Storage Cabinet','','',2,'','04/28/2025',2,'jv','','','','-'],
    // Tools
    ['Tools','Grinder','','',1,'','05/08/2025',1,'jv','','ordered','','Majic needs to cut channels for sconces'],
    ['Tools','Dust Free Pole Sander','Hyde','',4,'','05/08/2025',4,'jv','','ordered','','-'],
  ];

  const paddedData = data.map(row => {
    while (row.length < HEADERS.length) row.push('');
    return row.slice(0, HEADERS.length);
  });
  sheet.getRange(1, 1, paddedData.length, HEADERS.length).setValues(paddedData);

  const hdrRange = sheet.getRange(1, 1, 1, HEADERS.length);
  hdrRange.setFontWeight('bold');
  hdrRange.setBackground('#1a3a5c');
  hdrRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);

  Logger.log('Setup complete! ' + (data.length - 1) + ' materials loaded into the Master sheet.');
}

// ── Run this ONCE to apply color-coding rules to the Master sheet ─────────────
function setRowColors() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_SHEET);
  if (!sheet) { Logger.log('Master sheet not found — run setup() first.'); return; }

  sheet.clearConditionalFormatRules();

  const lastRow = Math.max(sheet.getLastRow(), 200);
  const numCols = HEADERS.length;
  const range   = sheet.getRange(2, 1, lastRow - 1, numCols);

  const onHandCol   = HEADERS.indexOf('On Hand')     + 1;
  const minStockCol = HEADERS.indexOf('Min Stock')    + 1;
  const statusCol   = HEADERS.indexOf('Order Status') + 1;
  const tradeCol    = 1;

  const onHandLetter   = columnLetter(onHandCol);
  const minStockLetter = columnLetter(minStockCol);
  const statusLetter   = columnLetter(statusCol);
  const tradeLetter    = columnLetter(tradeCol);

  const rules = [];

  // Priority 1: Ordered — green (highest priority, overrides low stock)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISNUMBER(SEARCH("ordered",$${statusLetter}2))`)
    .setBackground('#008000')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build());

  // Priority 2: Review — red (needs attention)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISNUMBER(SEARCH("review",$${statusLetter}2))`)
    .setBackground('#FF0000')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build());

  // Priority 3: Low stock — red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${onHandLetter}2<>"",$${minStockLetter}2<>"",$${minStockLetter}2>0,VALUE($${onHandLetter}2)<VALUE($${minStockLetter}2))`)
    .setBackground('#FF0000')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build());

  // Trade group colors (lower priority)
  const tradeColors = [
    ['Plaster',      '#dbeafe'],
    ['Paint',        '#ede9fe'],
    ['Wallcovering', '#fce7f3'],
    ['Carpentry',    '#fef3c7'],
    ['Electrical',   '#cffafe'],
    ['Misc.',        '#f3f4f6'],
    ['Tools',        '#fed7aa'],
  ];

  tradeColors.forEach(([trade, color]) => {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${tradeLetter}2="${trade}"`)
      .setBackground(color)
      .setRanges([range])
      .build());
  });

  sheet.setConditionalFormatRules(rules);
  Logger.log('Row colors applied! Low stock = red, Ordered = green, trade groups color-coded.');
}

// ── Add missing columns to Master sheet without touching existing data ────────
function addMissingColumns() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MASTER_SHEET);
  if (!sheet) { Logger.log('Master sheet not found.'); return; }

  const hdrRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existing = hdrRow.map(h => String(h).trim());

  let added = 0;
  HEADERS.forEach(h => {
    if (!existing.includes(h)) {
      const newCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, newCol).setValue(h);
      sheet.getRange(1, newCol).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
      existing.push(h);
      added++;
      Logger.log('Added column: ' + h);
    }
  });

  Logger.log(added > 0 ? added + ' column(s) added.' : 'No missing columns — sheet is up to date.');
}

// ── Custom menu ───────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Corridor's Plus")
    .addItem('Export Monthly Receipts PDF', 'exportMonthlyPDF')
    .addItem('Recalculate Receipt Totals', 'recalculateAllTotals')
    .addItem('Reapply Row Colors', 'setRowColors')
    .addToUi();
}

// ── Rebuild TOTAL rows for every vendor (run after manual row edits/deletions) ─
function recalculateAllTotals() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(RECEIPT_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('No Receipts sheet found.'); return; }

  const data = sheet.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();

  // Collect unique vendors from receipt rows (exclude TOTAL rows and header)
  const vendors = new Set();
  for (let i = 1; i < data.length; i++) {
    const rawD = data[i][0];
    const d    = rawD instanceof Date ? Utilities.formatDate(rawD, tz, 'yyyy-MM-dd') : String(rawD || '');
    if (!d.startsWith('TOTAL:') && data[i][1]) {
      vendors.add(String(data[i][1]).trim());
    }
  }

  vendors.forEach(vendor => rebuildVendorMonthlyTotals(sheet, vendor));
  SpreadsheetApp.getUi().alert('Totals recalculated for ' + vendors.size + ' vendor(s).');
}

// ── Export current month's receipts as a PDF saved to Drive ───────────────────
function exportMonthlyPDF() {
  const ui       = SpreadsheetApp.getUi();
  const tz       = Session.getScriptTimeZone();
  const now      = new Date();
  const monthKey = Utilities.formatDate(now, tz, 'yyyy-MM');       // "2026-05"
  const label    = Utilities.formatDate(now, tz, 'MMMM yyyy');     // "May 2026"

  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const srcSheet = ss.getSheetByName(RECEIPT_SHEET);
  if (!srcSheet) { ui.alert('No Receipts sheet found.'); return; }

  const allData = srcSheet.getDataRange().getValues();
  const headers = allData[0];

  // Filter rows that belong to this month (Date column starts with yyyy-MM)
  const monthRows = allData.slice(1).filter(r => String(r[0]).startsWith(monthKey));
  if (monthRows.length === 0) {
    ui.alert('No receipts found for ' + label + '.');
    return;
  }

  // Create a temporary sheet for the export
  const tmpName  = '_PDF_Export_Tmp';
  let   tmpSheet = ss.getSheetByName(tmpName);
  if (tmpSheet) ss.deleteSheet(tmpSheet);
  tmpSheet = ss.insertSheet(tmpName);

  // Write title, headers, and data rows
  tmpSheet.appendRow(['Corridor\'s Plus — Receipts: ' + label]);
  tmpSheet.appendRow([]);
  tmpSheet.appendRow(headers);
  monthRows.forEach(r => tmpSheet.appendRow(r));

  // Style title
  const titleRange = tmpSheet.getRange(1, 1, 1, headers.length);
  titleRange.merge();
  titleRange.setValue('Corridor\'s Plus — Receipts: ' + label);
  titleRange.setFontSize(14).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');

  // Style header row
  const hdrRange = tmpSheet.getRange(3, 1, 1, headers.length);
  hdrRange.setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff').setFontSize(11);

  // Format amount column
  const amtCol = headers.indexOf('Amount') + 1;
  if (amtCol > 0 && monthRows.length > 0) {
    tmpSheet.getRange(4, amtCol, monthRows.length, 1).setNumberFormat('$#,##0.00');
  }

  // Total row
  let totalAmt = 0;
  monthRows.forEach(r => { totalAmt += parseFloat(r[headers.indexOf('Amount')]) || 0; });
  const totalRow = new Array(headers.length).fill('');
  totalRow[0] = 'TOTAL';
  if (amtCol > 0) totalRow[amtCol - 1] = totalAmt;
  tmpSheet.appendRow(totalRow);
  const totRange = tmpSheet.getRange(4 + monthRows.length, 1, 1, headers.length);
  totRange.setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
  if (amtCol > 0) totRange.getCell(1, amtCol).setNumberFormat('$#,##0.00');

  tmpSheet.autoResizeColumns(1, headers.length);
  SpreadsheetApp.flush();

  // Export the temp sheet as PDF via the Sheets export URL
  const sheetId = tmpSheet.getSheetId();
  const pdfUrl  = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID +
    '/export?format=pdf&gid=' + sheetId +
    '&size=letter&portrait=true&fitw=true' +
    '&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';

  const token    = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(pdfUrl, { headers: { Authorization: 'Bearer ' + token } });
  const pdfBlob  = response.getBlob().setName('Receipts_' + monthKey + '.pdf');

  // Save to Drive receipts folder
  const folder = DriveApp.getFolderById(RECEIPTS_FOLDER_ID);
  const file   = folder.createFile(pdfBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Clean up temp sheet
  ss.deleteSheet(tmpSheet);

  ui.alert('PDF exported!\n\nFile: Receipts_' + monthKey + '.pdf\n\nOpen it here:\n' + file.getUrl());
}

// ── Helper: convert column number to letter (e.g. 1 → A, 27 → AA) ─────────────
function columnLetter(col) {
  let letter = '';
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ── Payroll export — push a pay period to the payroll spreadsheet ─────────────
const PAYROLL_SHEET_PROP = 'PAYROLL_SHEET_ID';

function handlePayroll(payload) {
  try {
    const props = PropertiesService.getScriptProperties();
    let   sheetId = props.getProperty(PAYROLL_SHEET_PROP);
    let   ss;

    if (sheetId) {
      try { ss = SpreadsheetApp.openById(sheetId); } catch(e) { sheetId = null; }
    }

    if (!sheetId) {
      // First run: create the payroll spreadsheet
      ss = SpreadsheetApp.create("Corridor's Plus — Payroll");
      ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
      DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheetId = ss.getId();
      props.setProperty(PAYROLL_SHEET_PROP, sheetId);
      // Remove the default blank sheet
      const blank = ss.getSheetByName('Sheet1');
      if (blank && ss.getSheets().length > 1) ss.deleteSheet(blank);
    }

    const { start, end, dates, workers } = payload;

    // Tab name: e.g. "05/13–05/19"
    const tabName = shortDateGs(start) + '–' + shortDateGs(end);

    // Overwrite if tab already exists (re-push)
    const existing = ss.getSheetByName(tabName);
    if (existing) ss.deleteSheet(existing);

    // Insert at front so most recent period is the first tab
    const sheet = ss.insertSheet(tabName, 0);

    // Column headers: Worker | Trade | Wed 5/13 | Thu 5/14 | … | Tue 5/19 | Total Hours
    const DOW = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const dayLabels = dates.map(function(d) {
      const parts = d.split('-').map(Number);
      const dow   = DOW[new Date(parts[0], parts[1] - 1, parts[2]).getDay()];
      const mm    = parts[1] < 10 ? '0' + parts[1] : String(parts[1]);
      const dd    = parts[2] < 10 ? '0' + parts[2] : String(parts[2]);
      return dow + ' ' + mm + '/' + dd;
    });

    const headers = ['Worker', 'Trade'].concat(dayLabels).concat(['Total Hours']);
    const numCols = headers.length;

    // Data rows
    const rows = [headers];
    workers.forEach(function(w) {
      rows.push([w.name, w.trade].concat(w.dayHours).concat([w.total]));
    });

    // Totals row
    const totalsRow = ['TOTAL', ''];
    dates.forEach(function(_, i) {
      totalsRow.push(workers.reduce(function(s, w) { return s + (w.dayHours[i] || 0); }, 0));
    });
    totalsRow.push(workers.reduce(function(s, w) { return s + w.total; }, 0));
    rows.push(totalsRow);

    sheet.getRange(1, 1, rows.length, numCols).setValues(rows);

    // ── Styling ───────────────────────────────────────────────────────────────
    const numDataRows = rows.length; // includes header + workers + totals

    // Header
    const hdr = sheet.getRange(1, 1, 1, numCols);
    hdr.setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(2);

    // Worker rows — alternate shading
    for (var i = 2; i <= numDataRows - 1; i++) {
      sheet.getRange(i, 1, 1, numCols).setBackground(i % 2 === 0 ? '#f8fafc' : '#ffffff').setFontSize(12);
    }

    // Totals row
    const tot = sheet.getRange(numDataRows, 1, 1, numCols);
    tot.setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13);

    // Center-align day columns and total column
    sheet.getRange(1, 3, numDataRows, numCols - 2).setHorizontalAlignment('center');

    // Highlight zero-hour cells in day columns (red tint) for at-a-glance gaps
    for (var r = 2; r <= numDataRows - 1; r++) {
      for (var c = 3; c <= numCols - 1; c++) {
        if (sheet.getRange(r, c).getValue() === 0) {
          sheet.getRange(r, c).setBackground('#fee2e2').setFontColor('#dc2626');
        }
      }
    }

    // Column widths
    sheet.setColumnWidth(1, 238);
    sheet.setColumnWidth(2, 159);
    for (var ci = 3; ci <= numCols - 1; ci++) sheet.setColumnWidth(ci, 113);
    sheet.setColumnWidth(numCols, 136);

    // Remove unused rows and columns so the sheet is clean
    const maxRows = sheet.getMaxRows();
    if (maxRows > numDataRows) sheet.deleteRows(numDataRows + 1, maxRows - numDataRows);
    const maxCols = sheet.getMaxColumns();
    if (maxCols > numCols) sheet.deleteColumns(numCols + 1, maxCols - numCols);

    // Convert to a native Sheets table (adds filter dropdowns + Table1 label)
    sheet.getRange(1, 1, numDataRows, numCols).createFilter();

    // Protect all cells — script can still write, nobody can edit manually
    const protection = sheet.protect().setDescription('Managed by Corridor\'s Plus — do not edit manually');
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) protection.setDomainEdit(false);

    // Sort payroll tabs: most recent start date on the left
    const tabPattern = /^\d{2}\/\d{2}–\d{2}\/\d{2}$/;
    const payrollTabs = ss.getSheets().filter(function(s) { return tabPattern.test(s.getName()); });
    payrollTabs.sort(function(a, b) {
      return b.getName().split('–')[0].localeCompare(a.getName().split('–')[0]);
    });
    payrollTabs.forEach(function(s, i) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(i + 1);
    });

    SpreadsheetApp.flush();

    const spreadsheetUrl = ss.getUrl();
    return json({ success: true, tabName: tabName, spreadsheetUrl: spreadsheetUrl });
  } catch(err) {
    return json({ error: err.toString() });
  }
}

function shortDateGs(iso) {
  var parts = iso.split('-');
  return parts[1] + '/' + parts[2];
}

// ── Helper: return JSON response ──────────────────────────────────────────────
function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
                       .setMimeType(ContentService.MimeType.JSON);
}
