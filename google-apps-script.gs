// Corridor's Plus — Google Sheets Sync
// Paste this entire file into: Google Sheet → Extensions → Apps Script
// Then: run setup() once, then deploy as a Web App (Deploy → New deployment → Web app → Anyone)

const SPREADSHEET_ID = '1u13Ai9uvKcau_cB4NjubcFnHkn2xXnsW3UmduOgbA0c';
const MASTER_SHEET  = 'Master';
const HEADERS = ['Trade', 'Material', 'Brand/Type', 'Unit', 'Qty', 'Status',
                 'Last Purchase Date', 'Last Purchase Qty', 'Purchased By', 'Notes', 'Link'];

// Called by the app to pull all materials (GET request)
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(MASTER_SHEET);
    if (!sheet) {
      return json({ error: 'Master sheet not found — run setup() first from the Apps Script editor.' });
    }
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return json([]);

    const headers   = data[0].map(h => String(h).trim());
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

// Called by the app to push material updates (POST request)
function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    const payload = JSON.parse(e.postData.contents);
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet   = ss.getSheetByName(MASTER_SHEET);

    if (!sheet) {
      sheet = ss.insertSheet(MASTER_SHEET, 0);
      sheet.appendRow(HEADERS);
    }

    const allData  = sheet.getDataRange().getValues();
    const hdr      = allData[0].map(h => String(h).trim());
    const matIdx   = hdr.indexOf('Material');
    const tradeIdx = hdr.indexOf('Trade');

    // Build a map: "trade||material" → 1-based row number
    const rowMap = {};
    allData.slice(1).forEach((r, i) => {
      const key = [String(r[tradeIdx] || ''), String(r[matIdx] || '')].join('||').toLowerCase();
      if (r[matIdx]) rowMap[key] = i + 2;
    });

    const updates = payload.materials || [];
    const newRows = [];

    updates.forEach(mat => {
      const key = [String(mat['Trade'] || ''), String(mat['Material'] || '')].join('||').toLowerCase();
      const row = HEADERS.map(h => (mat[h] !== undefined ? mat[h] : ''));
      if (rowMap[key]) {
        sheet.getRange(rowMap[key], 1, 1, HEADERS.length).setValues([row]);
      } else {
        newRows.push(row);
      }
    });

    newRows.forEach(row => sheet.appendRow(row));
    lock.releaseLock();
    return json({ success: true, updated: updates.length });
  } catch (err) {
    return json({ error: err.toString() });
  }
}

// ── Run this ONCE from the Apps Script editor to create the Master sheet ──────
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MASTER_SHEET);

  if (sheet) {
    sheet.clearContents();
  } else {
    sheet = ss.insertSheet(MASTER_SHEET, 0);
  }

  // Insert all rows at once (much faster than appendRow in a loop)
  const data = [
    HEADERS,
    // Plaster
    ['Plaster','Sealer/Primer','Zinsser Gardz2300','5 Gallon Buckets',5,'','04/28/2025',5,'jv','Before skim coat',''],
    ['Plaster','Joint Compound AP','All-Purpose','5 Gallon Buckets',20,'','04/28/2025',20,'jv','1 coat',''],
    ['Plaster','Joint Compound LW','Light Weight','5 Gallon Buckets',30,'','04/28/2025',30,'jv','1-2 coats',''],
    ['Plaster','Plaster of Paris','','Bags',10,'','04/28/2025',10,'jv','-',''],
    ['Plaster','Primer (Post-skim) Ceiling','Valspar High Build','5 Gallon Buckets',5,'','','','','After sanding, ON CEILING ONLY',''],
    ['Plaster','Primer (Post-skim) Walls','Roman Pro-977 Ultra Prime','5 Gallon Buckets',5,'ordered del 5/11','05/08/2025',5,'jv','After sanding, ON WALLS ONLY',''],
    ['Plaster','Sandpaper','180-400 grit','Multi Pack',10,'','04/28/2025',10,'jv','Purchase 180-400 Grit',''],
    ['Plaster','Drywall Sanding Screen','','10 pack',5,'','04/28/2025',5,'jv','-',''],
    ['Plaster','Hyde Sanding Screen','','',5,'ordered','05/08/2025',5,'jv','Order Amazon version',''],
    ['Plaster','Roller Covers 3/4"','3/4" nap','Multi-pack',12,'','','','','Primer application',''],
    ['Plaster','Roller Covers 1/2"','1/2" nap','Multi-pack',12,'','','','','Primer application',''],
    ['Plaster','Skim Rollers','1"-1.25" nap','Each',5,'','','','','For rolling compound',''],
    ['Plaster','PlasterWeld','','1 Gallon',15,'','','','','For repairs',''],
    // Paint
    ['Paint','Door Paint','','1 Gallon',25,'','','','','Benjamin Moore Command Newburg Green HC-158 Satin Finish',''],
    ['Paint','Ceiling Paint','','5 Gallon Bucket',5,'ordered del 5/11','05/08/2025',5,'jv','Benjamin Moore Brilliant White OC-51 Flat Finish',''],
    ['Paint','Raceway Paint','','5 Gallon Bucket',5,'ordered del 5/11','05/08/2025',5,'jv','Benjamin Moore Brilliant White OC-51 Eggshell Finish',''],
    ['Paint','Door and Raceway Primer','Stix','1 Gallon',15,'','','','','Stix primer',''],
    ['Paint','Painting Trays','','Single',8,'','04/28/2025',8,'jv','',''],
    ['Paint','Painting Tray Liners','','10 Pack',10,'','04/28/2025',10,'jv','',''],
    ['Paint','Ready Patch','','1 gallon',5,'','04/28/2025',5,'jv','',''],
    ['Paint','Caulk','DAP Alex Fast Dry','Box',5,'','04/28/2025',5,'jv','-',''],
    ['Paint','Roller Covers 1/4"','1/4" Nap','Multi Pack',12,'','','','','Order with door paint',''],
    // Wallcovering
    ['Wallcovering','Wallcovering','','rolls','','','','','','',''],
    ['Wallcovering','Wallcovering Paste','','gal','','','','','','',''],
    // Carpentry
    ['Carpentry','Baseboard','','pcs','','','','','','',''],
    ['Carpentry','Blocking','','pcs','','','','','','',''],
    ['Carpentry','Raceway','','pcs','','','','','','',''],
    // Electrical
    ['Electrical','Wire Nuts','3M','Multi Pack',3,'','','','','Johnny will pick up',''],
    ['Electrical','MC Cable','','Feet',250,'','','','','Johnny will pick up',''],
    ['Electrical','Single BX Connectors 3/8"','','',100,'','','','','Johnny will pick up',''],
    ['Electrical','1900 Box (shallow)','','',90,'','','','','Johnny will pick up, Majic to confirm',''],
    ['Electrical','1900 Box Blank Plates','','',90,'','','','','',''],
    ['Electrical','Octagon Boxes','','',60,'','','','','Johnny will pick up, Majic to confirm',''],
    ['Electrical','#12 Stranded Wire (White)','','Feet',500,'','','','','',''],
    ['Electrical','#12 Stranded Wire (Black)','','Feet',500,'','','','','',''],
    ['Electrical','Ground Tails','','Pieces',100,'','','','','',''],
    ['Electrical','1/2" EMT Pipe','','10 Foot Lengths',40,'','','','','',''],
    ['Electrical','1/2" EMT Connectors','','Pieces',120,'','','','','',''],
    ['Electrical','1/2" EMT LR','','Pieces',30,'','','','',' -',''],
    ['Electrical','1/2" EMT LL','','Pieces',30,'','','','',' -',''],
    ['Electrical','Light Retrofit Kits','','units','','','','','','',''],
    ['Electrical','Outlets','','pcs','','','','','','',''],
    ['Electrical','Device Covers','','pcs','','','','','','',''],
    ['Electrical','Sconces','','units','','','','','','',''],
    ['Electrical','EM Devices','','pcs','','','','','','',''],
    // Misc
    ['Misc.','Garbage Bags','-','Box',20,'','04/28/2025',20,'jv','',''],
    ["Misc.","Painter's Plastic",'-','400 Feet Box',10,'','04/28/2025',10,'jv','',''],
    ['Misc.','Masking Tape','3M','Multi-Pack',15,'','04/28/2025',15,'jv','-',''],
    ['Misc.','Frog Tape','','Multi-Pack',15,'','04/28/2025',15,'jv','-',''],
    ['Misc.','Drop Cloths','','Multi-Pack',10,'','04/28/2025',10,'jv','Size is 6X9',''],
    ['Misc.','N95 Masks','3M','Box',3,'','04/28/2025',3,'jv','-',''],
    ['Misc.','Zipties','','Box',1,'','04/28/2025',1,'jv','-',''],
    ['Misc.','Construction Adhesive','PL 400','Box',3,'','04/28/2025',3,'jv','-',''],
    ['Misc.','Painting Rags','','Box',3,'','04/28/2025',3,'jv','-',''],
    ['Misc.','Microfiber Rags','','',2,'','04/28/2025',2,'jv','-',''],
    ['Misc.','Finish Nails','','',5,'','04/28/2025',5,'jv','-',''],
    ['Misc.','Air Filters for Scrubber','','',4,'ordered','05/08/2025',4,'jv','Size is 13.5X14.5',''],
    ['Misc.','Extension Cords','Ridgid 50ft','',4,'','05/03/2025',8,'jr','-',''],
    ['Misc.','Splitters with Surge Protection','Ridgid','',2,'','05/03/2025',4,'jr','-',''],
    ['Misc.','Bulbs','','Multipack',5,'','04/29/2025',5,'jv','100 Watts',''],
    // Tools
    ['Tools','Grinder','','',1,'ordered','05/08/2025',1,'jv','Majic needs to cut channels for sconces',''],
    ['Tools','Dust Free Pole Sander','Hyde','',4,'ordered','05/08/2025',4,'jv','-',''],
    // Storage
    ['Misc.','Storage Cabinet','','',2,'','04/28/2025',2,'jv','-',''],
  ];

  sheet.getRange(1, 1, data.length, HEADERS.length).setValues(data);

  // Style the header row
  const hdrRange = sheet.getRange(1, 1, 1, HEADERS.length);
  hdrRange.setFontWeight('bold');
  hdrRange.setBackground('#1a3a5c');
  hdrRange.setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);

  Logger.log('Setup complete! ' + (data.length - 1) + ' materials loaded into the Master sheet.');
}

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
                       .setMimeType(ContentService.MimeType.JSON);
}
