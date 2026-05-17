// Corridor's Plus — Google Sheets Sync + Receipt Upload
// Paste this entire file into: Google Sheet → Extensions → Apps Script
// Then: run setup() once, then deploy as a Web App (Deploy → New deployment → Web app → Anyone)

const SPREADSHEET_ID     = '1u13Ai9uvKcau_cB4NjubcFnHkn2xXnsW3UmduOgbA0c';
const MASTER_SHEET       = 'Master';
const RECEIPTS_FOLDER_ID = '19TGCyD8v0j0UyahztXXuwIEGLVyk3LVJ';

const HEADERS = [
  'Trade', 'Material', 'Brand/Type', 'Unit',
  'On Hand', 'Min Stock',
  'Last Delivery Date', 'Last Delivery Qty', 'Ordered By',
  'To Order', 'Order Status', 'Delivery ETA',
  'Link', 'Notes'
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
    if (payload.type === 'status_update') return updateReceiptStatus(payload);
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

    // Ask Claude to analyze the receipt
    const analysis = analyzeReceiptWithClaude(base64, mimeType);

    // Build filename from analysis
    const vendor   = analysis.vendor   || payload.vendor   || 'Unknown';
    const amount   = analysis.amount   || payload.amount   || null;
    const category = analysis.category || payload.category || 'Other';
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const dateStr  = analysis.date     || todayStr;
    const amtStr   = amount ? '_$' + Number(amount).toFixed(2) : '';
    const filename = `${category.replace(/\s+/g,'-')}_${dateStr}_${vendor.replace(/\s+/g,'-')}${amtStr}.jpg`;

    // Save to the appropriate subfolder (use receipt date for correct month folder)
    const bytes  = Utilities.base64Decode(base64);
    const blob   = Utilities.newBlob(bytes, mimeType, filename);
    const folder = getOrCreateSubfolder(category, dateStr);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Upload any extra pages to the same folder
    const extraPages = Array.isArray(payload.extraPages) ? payload.extraPages : [];
    extraPages.forEach(function(page) {
      try {
        const ep     = page.data.split(',');
        const epMime = ep[0].split(';')[0].split(':')[1] || 'image/jpeg';
        const epB64  = ep[1];
        const epBlob = Utilities.newBlob(Utilities.base64Decode(epB64), epMime, page.filename || 'page.jpg');
        const epFile = folder.createFile(epBlob);
        epFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch(e) { /* skip failed extra page */ }
    });

    // Log to the Receipts sheet
    const items = Array.isArray(analysis.items) ? analysis.items : [];
    logReceiptToSheet({ date: dateStr, vendor, category, amount, url: file.getUrl(), items });

    return json({
      success:  true,
      url:      file.getUrl(),
      fileId:   file.getId(),
      vendor,
      amount,
      category,
      date:     dateStr,
    });
  } catch (err) {
    return json({ error: err.toString() });
  }
}

// ── Call Claude API to read the receipt image ─────────────────────────────────
function analyzeReceiptWithClaude(base64, mimeType) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('RECEIPT_API_KEY');
    if (!apiKey) return {};

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method:  'post',
      headers: {
        'x-api-key':         apiKey,
        'anthropic-version': '2023-06-01',
        'content-type':      'application/json',
      },
      payload: JSON.stringify({
        model:      'claude-haiku-4-5-20251001',
        max_tokens: 512,
        messages: [{
          role: 'user',
          content: [
            {
              type:   'image',
              source: { type: 'base64', media_type: mimeType, data: base64 },
            },
            {
              type: 'text',
              text: `You are reading a contractor's receipt or invoice. Extract the key details and respond ONLY with a JSON object — no explanation, no markdown.

Return exactly this shape:
{
  "vendor": "store or company name",
  "date": "YYYY-MM-DD",
  "amount": 123.45,
  "category": "one of: Materials & Supplies | Tools & Equipment | Subcontractors | Food & Meals | Other",
  "items": ["qty x item name — $price", "qty x item name — $price"]
}

Rules:
- vendor: the business name on the receipt (e.g. "Home Depot", "Grainger")
- date: the date printed on the receipt in YYYY-MM-DD format. null if not visible.
- amount: the total amount paid as a number, no $ sign. null if not visible.
- category: pick the best match from the five options above based on what was purchased.
- items: list each line item as "qty x description — $price". Max 10 items. Empty array if not readable.`
            }
          ]
        }]
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
const RECEIPT_HEADERS = ['Date', 'Vendor', 'Category', 'Amount', 'Status', 'Drive Link'];

// Vendor color palette — assigned dynamically as new vendors appear
const VENDOR_PALETTE = [
  '#dbeafe', '#fce7f3', '#fef3c7', '#ccfbf1', '#ede9fe',
  '#ffedd5', '#dcfce7', '#fef2f2', '#e0f2fe', '#fdf4ff',
];

// Category colors used in Summary tab only
const CATEGORY_COLORS = {
  'Materials & Supplies': '#dbeafe',
  'Tools & Equipment':    '#ccfbf1',
  'Subcontractors':       '#fce7f3',
  'Food & Meals':         '#fef3c7',
  'Other':                '#f3f4f6',
};

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

  const data       = sheet.getDataRange().getValues();
  const numRows    = data.length;
  const vendorCol  = 1;
  const dateCol    = 0;
  const bgColor    = getVendorColor(vendor);
  const vendorKey  = vendor.trim().toLowerCase();

  // Find the correct chronological insertion point within this vendor's rows.
  // We want to insert after the last vendor receipt row whose date <= new date,
  // skipping TOTAL subtotal rows. If vendor not present yet, append.
  let insertAfter = -1;
  for (let i = 1; i < numRows; i++) {
    const rowVendor = String(data[i][vendorCol] || '').trim().toLowerCase();
    const rowDate   = String(data[i][dateCol]   || '');
    const isTotal   = rowDate.startsWith('TOTAL:');
    if (rowVendor === vendorKey && !isTotal && rowDate <= date) {
      insertAfter = i + 1; // 1-based sheet row after this data row
    }
  }
  // If no earlier-or-equal date row found but vendor exists, insert before first vendor row
  if (insertAfter === -1) {
    for (let i = 1; i < numRows; i++) {
      const rowVendor = String(data[i][vendorCol] || '').trim().toLowerCase();
      const rowDate   = String(data[i][dateCol]   || '');
      const isTotal   = rowDate.startsWith('TOTAL:');
      if (rowVendor === vendorKey && !isTotal) {
        insertAfter = i; // insert BEFORE this row (sheet row i+1, so insertRowAfter(i) = before i+1)
        break;
      }
    }
  }

  const newRow = [date, vendor, category, amount ? Number(amount) : '', 'Pending', url];
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
  refreshSummarySheet(ss);
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
  const vendorRows = [];
  for (let i = 1; i < fresh.length; i++) {
    const v = String(fresh[i][vendorCol] || '').trim().toLowerCase();
    const d = String(fresh[i][dateCol]   || '');
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
    range.setValues([[`TOTAL: ${label}`, vendor, '', total, '', '']]);
    range.setBackground('#1a3a5c');
    range.setFontColor('#ffffff');
    range.setFontWeight('bold');
    range.setFontSize(11);
    range.getCell(1, 4).setNumberFormat('$#,##0.00');
  });
}

function styleReceiptRow(range, bgColor, items) {
  range.setBackground(bgColor);
  range.setFontSize(11);
  range.setVerticalAlignment('middle');
  range.getCell(1, 4).setNumberFormat('$#,##0.00');

  // Add line items as a hover note on the Amount cell
  if (items && items.length > 0) {
    const note = 'Items purchased:\n' + items.map(i => '• ' + i).join('\n');
    range.getCell(1, 4).setNote(note);
  }

  // Style Status cell
  const statusCell = range.getCell(1, 5);
  const status     = statusCell.getValue() || 'Pending';
  statusCell.setValue(status);
  statusCell.setFontWeight('bold');
  statusCell.setHorizontalAlignment('center');
  statusCell.setBackground(status === 'Paid' ? '#dcfce7' : '#fef3c7');
  statusCell.setFontColor(status === 'Paid' ? '#16a34a' : '#d97706');

  // Make Drive Link clickable (now column 6)
  const linkCell = range.getCell(1, 6);
  const url      = linkCell.getValue();
  if (url) {
    linkCell.setFormula(`=HYPERLINK("${url}","View Receipt")`);
    linkCell.setFontColor('#1a73e8');
  }
}

// ── Update receipt status in the sheet when toggled in the app ────────────────
function updateReceiptStatus(payload) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet   = ss.getSheetByName(RECEIPT_SHEET);
    if (!sheet) return json({ error: 'Receipts sheet not found' });

    const data      = sheet.getDataRange().getValues();
    const hdr       = data[0].map(h => String(h).trim());
    const linkCol   = hdr.indexOf('Drive Link');
    const statusCol = hdr.indexOf('Status');
    if (linkCol < 0 || statusCol < 0) return json({ error: 'Column not found' });

    const newStatus  = payload.status === 'paid' ? 'Paid' : 'Pending';
    const driveUrl   = payload.driveUrl || '';

    for (let i = 1; i < data.length; i++) {
      const cellVal = String(data[i][linkCol] || '');
      if (cellVal.includes(driveUrl) || driveUrl.includes(cellVal.replace(/.*"(https[^"]+)".*/, '$1'))) {
        const statusCell = sheet.getRange(i + 1, statusCol + 1);
        statusCell.setValue(newStatus);
        statusCell.setFontWeight('bold');
        statusCell.setHorizontalAlignment('center');
        statusCell.setBackground(newStatus === 'Paid' ? '#dcfce7' : '#fef3c7');
        statusCell.setFontColor(newStatus === 'Paid' ? '#16a34a' : '#d97706');
        return json({ success: true });
      }
    }
    return json({ error: 'Receipt row not found' });
  } catch (err) {
    return json({ error: err.toString() });
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
  sheet.setColumnWidth(1, 110); // Date
  sheet.setColumnWidth(2, 160); // Vendor
  sheet.setColumnWidth(3, 180); // Category
  sheet.setColumnWidth(4, 100); // Amount
  sheet.setColumnWidth(5, 100); // Status
  sheet.setColumnWidth(6, 130); // Drive Link
}

// ── Summary tab — totals by vendor and by category ───────────────────────────
const SUMMARY_SHEET = 'Receipt Summary';

function refreshSummarySheet(ss) {
  let sheet = ss.getSheetByName(SUMMARY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SUMMARY_SHEET);
  } else {
    sheet.clearContents();
    sheet.clearFormats();
  }

  const src  = ss.getSheetByName(RECEIPT_SHEET);
  if (!src) return;
  const data = src.getDataRange().getValues().slice(1).filter(r => r[0]); // skip header + blanks

  // Aggregate by vendor
  const byVendor = {};
  const byCat    = {};
  let   grandTotal = 0;

  data.forEach(r => {
    const vendor   = String(r[1] || 'Unknown').trim();
    const category = String(r[2] || 'Other').trim();
    const amount   = parseFloat(r[3]) || 0;
    byVendor[vendor]   = (byVendor[vendor]   || 0) + amount;
    byCat[category]    = (byCat[category]    || 0) + amount;
    grandTotal        += amount;
  });

  let row = 1;

  // Title
  sheet.getRange(row, 1, 1, 3).merge().setValue('Receipt Summary').setFontWeight('bold').setFontSize(14).setBackground('#1a3a5c').setFontColor('#ffffff');
  sheet.getRange(row, 4).setValue(new Date()).setNumberFormat('mmm d, yyyy').setFontColor('#888888');
  row += 2;

  // By Category
  sheet.getRange(row, 1, 1, 2).setValues([['CATEGORY', 'TOTAL']]);
  sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
  row++;
  Object.entries(byCat).sort((a,b) => b[1]-a[1]).forEach(([cat, total]) => {
    const bg = CATEGORY_COLORS[cat] || '#f9fafb';
    sheet.getRange(row, 1).setValue(cat).setBackground(bg);
    sheet.getRange(row, 2).setValue(total).setNumberFormat('$#,##0.00').setBackground(bg);
    row++;
  });
  sheet.getRange(row, 1).setValue('TOTAL').setFontWeight('bold');
  sheet.getRange(row, 2).setValue(grandTotal).setNumberFormat('$#,##0.00').setFontWeight('bold');
  row += 2;

  // By Vendor
  sheet.getRange(row, 1, 1, 2).setValues([['VENDOR', 'TOTAL']]);
  sheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
  row++;
  Object.entries(byVendor).sort((a,b) => b[1]-a[1]).forEach(([vendor, total]) => {
    sheet.getRange(row, 1).setValue(vendor);
    sheet.getRange(row, 2).setValue(total).setNumberFormat('$#,##0.00');
    row++;
  });

  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 120);
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

    // Build two maps: trade+name (exact) and name-only (fallback)
    const rowMap      = {};  // "trade||name" → 1-based row
    const nameOnlyMap = {};  // "name" → 1-based row (fallback when trade is missing/mismatched)
    allData.slice(1).forEach((r, i) => {
      const name  = String(r[matIdx] || '').trim().toLowerCase();
      const trade = tradeIdx >= 0 ? String(r[tradeIdx] || '').trim().toLowerCase() : '';
      if (!name) return;
      rowMap[trade + '||' + name] = i + 2;
      if (!nameOnlyMap[name]) nameOnlyMap[name] = i + 2;
    });

    // Fields the app is allowed to overwrite — Notes is sheet-managed
    const APP_FIELDS = HEADERS.filter(h => h !== 'Notes');

    const updates = payload.materials || [];
    const newRows = [];

    updates.forEach(mat => {
      const name     = String(mat['Material'] || '').trim().toLowerCase();
      const trade    = String(mat['Trade']    || '').trim().toLowerCase();
      const exactKey = trade + '||' + name;
      const rowNum   = rowMap[exactKey] || nameOnlyMap[name];
      if (rowNum) {
        APP_FIELDS.forEach(field => {
          const colIdx = hdr.indexOf(field);
          if (colIdx >= 0 && mat[field] !== undefined) {
            sheet.getRange(rowNum, colIdx + 1).setValue(mat[field]);
          }
        });
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
    const deletions = payload.deletions || [];
    if (deletions.length > 0) {
      const delKeys = new Set(deletions.map(d => (String(d.trade||'') + '||' + String(d.name||'')).toLowerCase()));
      const freshData     = sheet.getDataRange().getValues();
      const freshHdrTrade = freshData[0].map(h => String(h).trim()).indexOf('Trade');
      const freshHdrMat   = freshData[0].map(h => String(h).trim()).indexOf('Material');
      for (let i = freshData.length - 1; i >= 1; i--) {
        const key = (String(freshData[i][freshHdrTrade]||'') + '||' + String(freshData[i][freshHdrMat]||'')).toLowerCase();
        if (delKeys.has(key)) sheet.deleteRow(i + 1);
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

  // Priority 1: Low stock — red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${onHandLetter}2<>"",$${minStockLetter}2<>"",$${minStockLetter}2>0,VALUE($${onHandLetter}2)<VALUE($${minStockLetter}2))`)
    .setBackground('#FF0000')
    .setFontColor('#FFFFFF')
    .setRanges([range])
    .build());

  // Priority 2: Ordered — green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=ISNUMBER(SEARCH("ordered",$${statusLetter}2))`)
    .setBackground('#008000')
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
    ['Tools',        '#ccfbf1'],
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

// ── Custom menu ───────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Corridor's Plus")
    .addItem('Export Monthly Receipts PDF', 'exportMonthlyPDF')
    .addToUi();
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

// ── Helper: return JSON response ──────────────────────────────────────────────
function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
                       .setMimeType(ContentService.MimeType.JSON);
}
