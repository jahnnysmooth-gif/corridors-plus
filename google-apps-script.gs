// Corridor's Plus — Google Sheets Sync
// Paste this entire file into: Google Sheet → Extensions → Apps Script
// Then: run setup() once, then deploy as a Web App (Deploy → New deployment → Web app → Anyone)

const SPREADSHEET_ID = '1u13Ai9uvKcau_cB4NjubcFnHkn2xXnsW3UmduOgbA0c';
const MASTER_SHEET  = 'Master';
const HEADERS = [
  'Trade', 'Material', 'Brand/Type', 'Unit',
  'On Hand', 'Min Stock',
  'Last Delivery Date', 'Last Delivery Qty', 'Ordered By',
  'To Order', 'Order Status',
  'Link', 'Notes'
];

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
    const hdr      = allData[0].map(h => String(h).trim().replace(/[\r\n]+/g, ' ').replace(/\s+/g, ' '));
    const matIdx   = hdr.indexOf('Material');
    const tradeIdx = hdr.indexOf('Trade');

    // Build a map: "trade||material" → 1-based row number
    const rowMap = {};
    allData.slice(1).forEach((r, i) => {
      const key = [String(r[tradeIdx] || ''), String(r[matIdx] || '')].join('||').toLowerCase();
      if (r[matIdx]) rowMap[key] = i + 2;
    });

    // All fields are app-managed (partial update still protects against full row wipe)
    const APP_FIELDS = HEADERS;

    const updates = payload.materials || [];
    const newRows = [];

    updates.forEach(mat => {
      const key = [String(mat['Trade'] || ''), String(mat['Material'] || '')].join('||').toLowerCase();
      if (rowMap[key]) {
        // Partial update — only overwrite app-managed columns, leave Min Stock, Brand, Link, Notes etc. alone
        APP_FIELDS.forEach(field => {
          const colIdx = hdr.indexOf(field);
          if (colIdx >= 0 && mat[field] !== undefined) {
            sheet.getRange(rowMap[key], colIdx + 1).setValue(mat[field]);
          }
        });
      } else {
        // New row — push all fields
        newRows.push(HEADERS.map(h => (mat[h] !== undefined ? mat[h] : '')));
      }
    });

    newRows.forEach(row => sheet.appendRow(row));

    // Handle deletions — remove sheet rows for materials deleted in the app
    const deletions = payload.deletions || [];
    if (deletions.length > 0) {
      const delKeys = new Set(deletions.map(d => (String(d.trade||'') + '||' + String(d.name||'')).toLowerCase()));
      const freshData = sheet.getDataRange().getValues();
      const freshHdrTrade = freshData[0].map(h => String(h).trim()).indexOf('Trade');
      const freshHdrMat   = freshData[0].map(h => String(h).trim()).indexOf('Material');
      // Work backwards so row deletion doesn't shift indices
      for (let i = freshData.length - 1; i >= 1; i--) {
        const key = (String(freshData[i][freshHdrTrade]||'') + '||' + String(freshData[i][freshHdrMat]||'')).toLowerCase();
        if (delKeys.has(key)) sheet.deleteRow(i + 1);
      }
    }

    lock.releaseLock();
    return json({ success: true, updated: updates.length, deleted: deletions.length });
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

  // Columns: Trade | Material | Brand/Type | Unit | On Hand | Min Stock |
  //          Last Delivery Date | Last Delivery Qty | Ordered By |
  //          To Order | Order Status | Link | Notes
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
