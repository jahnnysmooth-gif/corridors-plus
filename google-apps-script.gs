// Corridor's Plus — Google Sheets Sync + Receipt Upload
// Paste this entire file into: Google Sheet → Extensions → Apps Script
// Then: run setup() once, then deploy as a Web App (Deploy → New deployment → Web app → Anyone)

const SPREADSHEET_ID     = '1u13Ai9uvKcau_cB4NjubcFnHkn2xXnsW3UmduOgbA0c';
const MASTER_SHEET       = 'Master';
const RECEIPTS_FOLDER_ID = '1GxTFqBDmGK6AruHeFmNLOXH2zqUlmXZK';

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

    // Receipt upload request — handle separately (no sheet lock needed)
    if (payload.type === 'receipt') {
      return uploadReceipt(payload);
    }

    // Materials sync request
    return syncMaterials(payload);

  } catch (err) {
    return json({ error: err.toString() });
  }
}

// ── Receipt upload — save image to Drive folder ───────────────────────────────
function uploadReceipt(payload) {
  try {
    if (!payload.data) return json({ error: 'No image data received.' });

    const parts    = payload.data.split(',');
    const mimeType = parts[0].split(';')[0].split(':')[1] || 'image/jpeg';
    const base64   = parts[1];
    const bytes    = Utilities.base64Decode(base64);
    const blob     = Utilities.newBlob(bytes, mimeType, payload.filename || 'receipt.jpg');

    const folder = DriveApp.getFolderById(RECEIPTS_FOLDER_ID);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return json({ success: true, url: file.getUrl(), fileId: file.getId() });
  } catch (err) {
    return json({ error: err.toString() });
  }
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
