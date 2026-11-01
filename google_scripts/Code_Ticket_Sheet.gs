// =============================================================
// SECTION 1 — GLOBALS
// Sheet references are assigned lazily inside each function via
// getSheetVariables() so they are never null when setupSheets()
// has not yet been run.
// =============================================================

var ss            = SpreadsheetApp.getActiveSpreadsheet();
var trackingSheet = null;
var summarySheet  = null;
var ticketSheet   = null;
var CONFIG        = {};

// Column constants (1-based)
var COL_A = 1;
var COL_B = 2;
var COL_C = 3;
var COL_D = 4;
var COL_E = 5;
var COL_F = 6;
var COL_G = 7;
var COL_H = 8;
var COL_I = 9;  // Reservation ID
var COL_J = 10; // Total Price
var COL_K = 11; // Status
var COL_L = 12; // Reservation Received Email flag
var COL_M = 13; // Reservation Email Date/Time
var COL_N = 14; // Reminder Email flag
var COL_O = 15; // Confirmed Email flag
var COL_P = 16; // Waitlisted Email flag
var COL_Q = 17; // Cancelled Email flag
var COL_R = 18; // Early Bird flag
var COL_S = 19; // Row number (autofill)

// Resolves and caches sheet references. Called at the top of
// every function that accesses the sheets.
function getSheetVariables() {
  trackingSheet = ss.getSheetByName("Tracking");
  summarySheet  = ss.getSheetByName("Counters");
  ticketSheet   = ss.getSheetByName("Tickets");
}


// =============================================================
// SECTION 1b — REFERENCES, TEMPLATE IDS AND IMAGE ASSETS
// The "References" tab is written by createReservationForm.gs.
// It holds the folder URLs and the Drive file ID of every HTML
// template. Template IDs are read from there (no hardcoded list)
// and written to the Counters template rows by setupSheets(); loadConfig()
// then loads them into CONFIG at runtime.
// =============================================================

// Base URL used to build image links from a Drive file ID.
var imgLink = "https://drive.google.com/uc?export=view&id=";

// Counters key -> template file name as listed in the References tab
var TEMPLATE_FILE_NAMES = {
  bank:      "ReservationReceived_BANK.html",
  gcash:     "ReservationReceived_GCash.html",
  paypal:    "ReservationReceived_Paypal.html",
  waitlist:  "Waitlisted.html",
  reminder:  "Reminder.html",
  confirmed: "TicketConfirmed.html",
  cancelled: "ReservationCancelled.html",
  ticketPDF: "TicketPDF.html"
};

// Image base names (without extension) looked up in the asset folders
var ASSET_IMAGE_NAMES = [
  "facebook", "instagram", "twitter", "header", "banner",
  "bank-info", "gcash-info", "gcash-qr", "bank-qr", "maya-info", "maya-qr"
];

// Reads the References tab into a { item: value } map.
// Returns null if the tab does not exist.
function readReferences(spreadsheet) {
  var refSheet = spreadsheet.getSheetByName("References");
  if (!refSheet) return null;

  var last = refSheet.getLastRow();
  var map  = {};
  if (last < 2) return map;

  refSheet.getRange(2, 1, last - 1, 2).getValues().forEach(function(row) {
    if (row[0] !== "") map[String(row[0])] = String(row[1]);
  });
  return map;
}

// Returns { bank, gcash, paypal, waitlist, reminder, confirmed,
// cancelled, ticketPDF } with Drive file IDs from the References tab.
// Throws a clear error listing anything missing.
function getTemplateIdsFromReferences(spreadsheet) {
  var refs = readReferences(spreadsheet);
  if (!refs) {
    throw new Error('The "References" tab was not found. It is created by createReservationForm.gs; template IDs are read from it.');
  }

  var ids     = {};
  var missing = [];
  Object.keys(TEMPLATE_FILE_NAMES).forEach(function(key) {
    var fileName = TEMPLATE_FILE_NAMES[key];
    var id       = refs["Template ID: " + fileName];
    if (!id || id === "NOT FOUND") {
      missing.push(fileName);
    } else {
      ids[key] = id;
    }
  });

  if (missing.length > 0) {
    throw new Error('Template ID(s) missing in the "References" tab: ' + missing.join(", "));
  }
  return ids;
}

// =============================================================
// COUNTERS CONFIG LAYOUT (columns D = label, E = value)
// The single source of truth for where each config field lives. Every
// block is a bordered group; one grey spacer row separates blocks.
// setupSheets() writes the sheet from this list and loadConfig() reads
// it back through CFG_ROW, so changing the order here (e.g. adding or
// moving a field) needs no other row-number edits.
// =============================================================

var COUNTERS_LAYOUT = [
  [
    { id: "headerLink",       label: "Header Link",                       def: "" },
    { id: "showName",         label: "Show Name",                         def: "" },
    { id: "bannerLink",       label: "Banner Link",                       def: "" },
    { id: "showPrefix",       label: "Show Prefix",                       def: "" }
  ],
  [
    { id: "ticketPrice",      label: "Ticket Price",                      def: 0 },
    { id: "ticketLimit",      label: "Ticket Limit",                      def: 0 }
  ],
  [
    { id: "bankName",         label: "Bank Name",                         def: "" },
    { id: "bankAccountName",  label: "Bank Account Name",                 def: "" },
    { id: "bankAccountNumber",label: "Bank Account Number",               def: "" },
    { id: "bankQR",           label: "Bank QR Link",                      def: "" },
    { id: "bankInfographic",  label: "Bank Infographic Link",             def: "" }
  ],
  [
    { id: "gcashNumber",      label: "GCash Number",                      def: "" },
    { id: "gcashQR",          label: "GCash QR Link",                     def: "" },
    { id: "gcashInfographic", label: "GCash Infographic Link",            def: "" }
  ],
  [
    { id: "paypalLink",       label: "Paypal Link",                       def: "" },
    { id: "paypalFee",        label: "Paypal Transaction Fee",            def: 0 }
  ],
  [
    { id: "mayaQR",           label: "Maya QR Link",                      def: "" },
    { id: "mayaInfographic",  label: "Maya Infographic Link",             def: "" }
  ],
  [
    { id: "reminderDays",     label: "Days before reminder email",        def: 1 },
    { id: "showDate",         label: "Show Date",                         def: "" },
    { id: "waitText",         label: "Reservation Waiting Period Text",   def: "by 10 PM tomorrow" },
    { id: "waitDays",         label: "Reservation Waiting Period (Days)", def: 1 }
  ],
  [
    { id: "startRow",         label: "Start Processing From Row",         def: 2 }
  ],
  [
    { id: "earlyBirdEnabled", label: "Early Bird Enabled",                def: "No" },
    { id: "earlyBirdLimit",   label: "Early Bird Limit",                  def: 0 },
    { id: "earlyBirdPrice",   label: "Early Bird Ticket Price",           def: 0 }
  ],
  [
    { id: "webAppUrl",        label: "Web App URL",                       def: "" },
    { id: "qrFolderId",       label: "QR Folder ID",                      def: "" }
  ],
  [
    { id: "tplBank",          label: "Template: Reservation Bank",        tpl: "bank" },
    { id: "tplGCash",         label: "Template: Reservation GCash",       tpl: "gcash" },
    { id: "tplPayPal",        label: "Template: Reservation PayPal",      tpl: "paypal" },
    { id: "tplWaitlisted",    label: "Template: Waitlisted",              tpl: "waitlist" },
    { id: "tplReminder",      label: "Template: Reminder",                tpl: "reminder" },
    { id: "tplConfirmed",     label: "Template: Ticket Confirmed",        tpl: "confirmed" },
    { id: "tplCancelled",     label: "Template: Reservation Cancelled",   tpl: "cancelled" },
    { id: "tplTicketPDF",     label: "Template: Ticket PDF",              tpl: "ticketPDF" }
  ],
  [
    { id: "facebookIcon",     label: "Facebook Link",                     def: "" },
    { id: "instagramIcon",    label: "Instagram Link",                    def: "" },
    { id: "twitterIcon",      label: "Twitter Link",                      def: "" }
  ],
  [
    { id: "facebookUrl",      label: "Facebook URL",                      def: "" },
    { id: "instagramUrl",     label: "Instagram URL",                     def: "" },
    { id: "twitterUrl",       label: "Twitter URL",                       def: "" }
  ],
  [
    { id: "emailSenderName",  label: "Email Sender Name",                 def: "" }
  ]
];

var CFG_ROW     = {};   // field id -> row number in Counters
var CFG_LABEL   = {};   // field id -> label
var CFG_FIELDS  = [];   // flat list with row numbers
var CFG_BLOCKS  = [];   // { first, last } row of each block
var CFG_LAST_ROW = 0;   // last row that holds a config value

(function buildCountersLayout() {
  var r = 1;
  COUNTERS_LAYOUT.forEach(function(block) {
    var first = r;
    block.forEach(function(f) {
      CFG_ROW[f.id]   = r;
      CFG_LABEL[f.id] = f.label;
      CFG_FIELDS.push({ id: f.id, label: f.label, def: f.def, tpl: f.tpl, row: r });
      r++;
    });
    CFG_BLOCKS.push({ first: first, last: r - 1 });
    r++;                     // grey spacer row between blocks
  });
  CFG_LAST_ROW = r - 2;
})();

// A1 address of a config value cell, e.g. counterCell("showDate") -> "E26"
function counterCell(id) {
  return "E" + CFG_ROW[id];
}

// Logs a warning if an already-populated Counters tab does not match the
// layout above (e.g. it was created by an older version of this script).
function checkCountersLayout(counters) {
  var mismatches = [];
  CFG_FIELDS.forEach(function(f) {
    var actual = String(counters.getRange(f.row, 4).getValue());
    if (actual !== f.label) {
      mismatches.push("D" + f.row + ": expected '" + f.label + "', found '" + actual + "'");
    }
  });
  if (mismatches.length > 0) {
    Logger.log("WARNING: the Counters tab does not match the current layout, so config values " +
               "may be read from the wrong cells. Delete the Counters tab and run setupSheets() " +
               "again (first note ID COUNT in B1 and EARLY BIRD TICKETS SOLD in B12). Mismatches: " +
               mismatches.join(" | "));
  } else {
    Logger.log("Counters layout check: OK.");
  }
}

// Counters field <- References tab item. type "number" writes a number,
// "text" writes plain text (so leading zeros / dates are never converted).
var REFERENCE_TO_COUNTERS = [
  { id: "showName",          key: "Show Name",           type: "text"   },
  { id: "showPrefix",        key: "Show Prefix",         type: "text"   },
  { id: "ticketPrice",       key: "Ticket Price",        type: "number" },
  { id: "ticketLimit",       key: "Ticket Limit",        type: "number" },
  { id: "bankName",          key: "Bank Name",           type: "text"   },
  { id: "bankAccountName",   key: "Bank Account Name",   type: "text"   },
  { id: "bankAccountNumber", key: "Bank Account Number", type: "text"   },
  { id: "gcashNumber",       key: "GCash Number",        type: "text"   },
  { id: "paypalLink",        key: "Paypal Link",         type: "text"   },
  { id: "paypalFee",         key: "Paypal Transaction Fee", type: "number" },
  { id: "showDate",          key: "Show Date",           type: "text"   },
  { id: "qrFolderId",        key: "QR Folder ID",        type: "text"   },
  { id: "facebookUrl",       key: "Facebook URL",        type: "text"   },
  { id: "instagramUrl",      key: "Instagram URL",       type: "text"   },
  { id: "twitterUrl",        key: "Twitter URL",         type: "text"   },
  { id: "emailSenderName",   key: "Email Sender Name",   type: "text"   }
];

// True for References values that are still the unedited placeholders
// from Create_Ticket_Form.gs (e.g. "<BANK NAME HERE>", "Month dd, YYYY").
function isReferencePlaceholder(value) {
  var v = String(value).trim();
  return v === "" || /^<.*>$/.test(v) || v === "Month dd, YYYY";
}

// Copies the matching References values into Counters column E.
// A cell is only written when it is empty (or still the 0 default of a
// number field), so values typed in by hand are never overwritten.
function fillCountersFromReferences(counters, refs) {
  REFERENCE_TO_COUNTERS.forEach(function(entry) {
    var cell  = counterCell(entry.id);
    var label = CFG_LABEL[entry.id];
    var raw   = refs[entry.key];
    if (raw === undefined || isReferencePlaceholder(raw)) {
      Logger.log(label + ": no usable value in References ('" + entry.key + "'), left unchanged.");
      return;
    }

    var target  = counters.getRange(cell);
    var current = target.getValue();
    var isEmpty = (current === "" || current === null) ||
                  (entry.type === "number" && Number(current) === 0);
    if (!isEmpty) {
      Logger.log(label + " (" + cell + ") already has a value, skipping.");
      return;
    }

    if (entry.type === "number") {
      var num = Number(String(raw).replace(/,/g, ""));
      if (isNaN(num)) {
        Logger.log(label + ": '" + raw + "' is not a number, left unchanged.");
        return;
      }
      target.setValue(num);
    } else {
      target.setNumberFormat("@");   // keep as text
      target.setValue(String(raw));
    }
    Logger.log(label + " (" + cell + ") set from References.");
  });
}

// Extracts a Drive folder ID from a folder URL.
function folderIdFromUrl(url) {
  var m = String(url || "").match(/\/folders\/([A-Za-z0-9_-]+)/);
  return m ? m[1] : null;
}

// Builds { lowercased file name: file ID } for the files directly inside
// `folder` (first file wins if a name occurs twice).
function indexFolderFiles(folder) {
  var index = {};
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f    = files.next();
    var name = f.getName().toLowerCase();
    if (!index.hasOwnProperty(name)) index[name] = f.getId();
  }
  return index;
}

// Returns the file ID for baseName.png / .jpg / .jpeg (in that order,
// case-insensitive, so "Header.PNG" also matches) from an index built by
// indexFolderFiles(), or null.
function findImageId(index, baseName) {
  var exts = [".png", ".jpg", ".jpeg"];
  for (var i = 0; i < exts.length; i++) {
    var key = (baseName + exts[i]).toLowerCase();
    if (index.hasOwnProperty(key)) return index[key];
  }
  return null;
}

// Looks up every image in ASSET_IMAGE_NAMES. The show_assets folder
// wins over global_assets when a file exists in both.
// Returns { baseName: fileId | null }. Folder URLs come from the
// "References" tab ("Show Assets URL" / "Global Assets URL").
function getAssetImageIds(spreadsheet) {
  var ids = {};
  ASSET_IMAGE_NAMES.forEach(function(n) { ids[n] = null; });

  var refs = readReferences(spreadsheet);
  if (!refs) {
    Logger.log('Assets: "References" tab not found, cannot locate asset folders.');
    return ids;
  }

  // Priority order: show_assets first, then global_assets.
  var sources = [
    { label: "show_assets",   key: "Show Assets URL"   },
    { label: "global_assets", key: "Global Assets URL" }
  ];
  var folders = [];
  sources.forEach(function(src) {
    var folderId = folderIdFromUrl(refs[src.key]);
    if (!folderId) {
      Logger.log('Assets: "' + src.key + '" missing or not a folder URL in the References tab (' + refs[src.key] + ').');
      return;
    }
    try {
      var index = indexFolderFiles(DriveApp.getFolderById(folderId));
      Logger.log("Assets: " + src.label + " contains " + Object.keys(index).length + " file(s).");
      folders.push({ label: src.label, index: index });
    } catch (err) {
      Logger.log("Assets: could not open " + src.label + " (" + folderId + "): " + err);
    }
  });

  ASSET_IMAGE_NAMES.forEach(function(name) {
    for (var i = 0; i < folders.length; i++) {
      var id = findImageId(folders[i].index, name);
      if (id) {
        ids[name] = id;
        Logger.log("Asset " + name + ": " + id + " (from " + folders[i].label + ")");
        return;
      }
    }
    Logger.log("Asset " + name + ": NOT FOUND (.png/.jpg/.jpeg) in show_assets or global_assets.");
  });

  return ids;
}


// =============================================================
// SECTION 2 — SHEET SETUP
// Must come before loadConfig() and all core functions since
// everything depends on the sheets existing first.
// Run setupSheets() once manually from the Apps Script editor
// after creating the Google Form and its linked spreadsheet.
// Safe to re-run — each block is only written if not already
// present.
// =============================================================

function setupSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // ── TRACKING ──────────────────────────────────────────────
  var tracking = spreadsheet.getSheetByName("Tracking");
  if (!tracking) {
    tracking = spreadsheet.insertSheet("Tracking");
    Logger.log("Created sheet: Tracking");
  } else {
    Logger.log("Sheet already exists, skipping creation: Tracking");
  }

  if (!tracking.getRange("A1").getValue()) {
    // Cols A-H are populated via QUERY from Form Responses 1.
    tracking.getRange("A1").setFormula(
      '=IFERROR(QUERY(\'Form Responses 1\'!A:H), "Timestamp")'
    );
    // Cols I-S are script-managed — set plain text headers.
    tracking.getRange("I1:S1").setValues([[
      "Reservation ID",
      "Total Price",
      "Status",
      "Reservation Received Email",
      "Reservation Email Date/Time",
      "Reminder Email",
      "Confirmed Email",
      "Waitlisted Email",
      "Cancelled Email",
      "Early Bird?",
      "Row"
    ]]);
    // Seed ROW formula in S2; ensureTrackingRowNumbers() fills
    // it down as data grows.
    tracking.getRange("S2").setFormula("=ROW()");

    // ── Header formatting ──
    // A–H: form response columns (light green to distinguish from script cols)
    var trkFormHeader = tracking.getRange("A1:H1");
    trkFormHeader.setBackground("#B7E1CD")
                 .setFontWeight("bold")
                 .setFontColor("#000000")
                 .setBorder(true, true, true, true, null, null, "#4A86C8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // I–S: script-managed columns (light blue, same palette as Counters)
    var trkScriptHeader = tracking.getRange("I1:S1");
    trkScriptHeader.setBackground("#CFE2F3")
                   .setFontWeight("bold")
                   .setFontColor("#000000")
                   .setBorder(true, true, true, true, null, null, "#4A86C8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Freeze header row and set column widths
    tracking.setFrozenRows(1);
    tracking.setColumnWidth(1,  180); // A  Timestamp
    tracking.setColumnWidth(2,  120); // B  First Name
    tracking.setColumnWidth(3,  120); // C  Last Name
    tracking.setColumnWidth(4,  200); // D  Email
    tracking.setColumnWidth(5,  130); // E  Mobile
    tracking.setColumnWidth(6,  120); // F  City
    tracking.setColumnWidth(7,  60);  // G  # Tickets
    tracking.setColumnWidth(8,  120); // H  Payment Option
    tracking.setColumnWidth(9,  160); // I  Reservation ID
    tracking.setColumnWidth(10, 100); // J  Total Price
    tracking.setColumnWidth(11, 160); // K  Status
    tracking.setColumnWidth(12, 200); // L  Reservation Received Email
    tracking.setColumnWidth(13, 200); // M  Reservation Email Date/Time
    tracking.setColumnWidth(14, 130); // N  Reminder Email
    tracking.setColumnWidth(15, 140); // O  Confirmed Email

    // Format col A (rows 2 onward) as date/time — matches the
    // timestamp values written by the Google Form on submission.
    tracking.getRange("A2:A").setNumberFormat("m/d/yyyy h:mm:ss");
    tracking.setColumnWidth(16, 140); // P  Waitlisted Email
    tracking.setColumnWidth(17, 130); // Q  Cancelled Email
    tracking.setColumnWidth(18, 100); // R  Early Bird?
    tracking.setColumnWidth(19, 60);  // S  Row

    Logger.log("Tracking: headers written.");
  } else {
    Logger.log("Tracking: headers already present, skipping.");
  }

  // ── TICKETS ───────────────────────────────────────────────
  var tickets = spreadsheet.getSheetByName("Tickets");
  if (!tickets) {
    tickets = spreadsheet.insertSheet("Tickets");
    Logger.log("Created sheet: Tickets");
  } else {
    Logger.log("Sheet already exists, skipping creation: Tickets");
  }

  if (!tickets.getRange("A1").getValue()) {
    tickets.getRange("A1:J1").setValues([[
      "First Name",
      "Last Name",
      "Email Address",
      "Mobile Number",
      "Number of Tickets",
      "Reservation ID",
      "Status",
      "UUID",
      "Ticket Number",
      "Ticket ID"
    ]]);

    // ── Header formatting ──
    var tktHeader = tickets.getRange("A1:J1");
    tktHeader.setBackground("#CFE2F3")
             .setFontWeight("bold")
             .setFontColor("#000000")
             .setBorder(true, true, true, true, null, null, "#4A86C8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Freeze header row and set column widths
    tickets.setFrozenRows(1);
    tickets.setColumnWidth(1,  120); // A  First Name
    tickets.setColumnWidth(2,  120); // B  Last Name
    tickets.setColumnWidth(3,  200); // C  Email Address
    tickets.setColumnWidth(4,  130); // D  Mobile Number
    tickets.setColumnWidth(5,  120); // E  Number of Tickets
    tickets.setColumnWidth(6,  160); // F  Reservation ID
    tickets.setColumnWidth(7,  160); // G  Status
    tickets.setColumnWidth(8,  240); // H  UUID
    tickets.setColumnWidth(9,  110); // I  Ticket Number
    tickets.setColumnWidth(10, 180); // J  Ticket ID

    Logger.log("Tickets: headers written.");
  } else {
    Logger.log("Tickets: headers already present, skipping.");
  }

  // ── COUNTERS ──────────────────────────────────────────────
  var counters = spreadsheet.getSheetByName("Counters");
  if (!counters) {
    counters = spreadsheet.insertSheet("Counters");
    Logger.log("Created sheet: Counters");
  } else {
    Logger.log("Sheet already exists, skipping creation: Counters");
  }

  if (!counters.getRange("A1").getValue()) {

    // Template file IDs come from the References tab (throws if missing).
    var TEMPLATE_IDS = getTemplateIdsFromReferences(spreadsheet);

    // Col A: labels, Col B: values / formulas
    var abData = [
      ["ID COUNT",                                0],           // A1/B1 — managed by setReservation()
      [null,                                      null],        // A2/B2
      ["ON HOLD COUNT",                           '=SUMIF(Tracking!K:K,"*On Hold*",Tracking!G:G)'],
      ["PAYMENT CONFIRMED/TICKETS CLAIMED COUNT", '=SUMIF(Tracking!K:K,"Payment Confirmed",Tracking!G:G)+SUMIF(Tracking!K:K,"Ticket/s Claimed",Tracking!G:G)'],
      [null,                                      null],
      ["CANCELLED STATUS COUNT",                  '=SUMIF(Tracking!K:K,"Cancelled",Tracking!G:G)'],
      ["CANCELLED EMAIL SENT COUNT",              '=SUMIF(Tracking!Q:Q,"SENT",Tracking!G:G)'],
      [null,                                      null],
      ["TICKET COUNTER (PENDING & CONFIRMED)",    "=B3+B4+(B6-B7)"],
      ["TICKET SLOTS REMAINING",                  "=" + counterCell("ticketLimit") + "-B9"],
      [null,                                      null],
      ["EARLY BIRD TICKETS SOLD",                 0]            // A12/B12 — managed by setReservation()
    ];

    for (var r = 0; r < abData.length; r++) {
      var rowNum = r + 1;
      if (abData[r][0]) counters.getRange(rowNum, 1).setValue(abData[r][0]);
      if (abData[r][1] !== null) {
        var bVal = abData[r][1];
        if (typeof bVal === 'string' && bVal.charAt(0) === '=') {
          counters.getRange(rowNum, 2).setFormula(bVal);
        } else {
          counters.getRange(rowNum, 2).setValue(bVal);
        }
      }
    }

    // Col D: labels, Col E: default values — rows come from COUNTERS_LAYOUT.
    CFG_FIELDS.forEach(function(f) {
      counters.getRange(f.row, 4).setValue(f.label);
      counters.getRange(f.row, 5).setValue(f.tpl ? TEMPLATE_IDS[f.tpl] : f.def);
    });

    // Early Bird Enabled dropdown
    var ebRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Yes", "No", "Exceeded"])
      .build();
    counters.getRange(counterCell("earlyBirdEnabled")).setDataValidation(ebRule);

    // Format Show Date as plain text so entries like
    // "July 12, 2025" are never auto-converted to a date value.
    counters.getRange(counterCell("showDate")).setNumberFormat("@");

    // ── FORMATTING ────────────────────────────────────────────
    // Colour palette
    var COLOR_LABEL   = "#CFE2F3"; // light blue  — label cells
    var COLOR_VALUE   = "#FFFFFF"; // white        — value/formula cells
    var COLOR_SKIP    = "#F3F3F3"; // light grey   — skipped/blank rows

    // Border styles shorthand
    var BORDER_SOLID  = SpreadsheetApp.BorderStyle.SOLID;
    var BORDER_MEDIUM = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;

    // Helper: apply section formatting to a two-column block.
    function formatSection(rows, colLabel, colValue) {
      rows.forEach(function(entry) {
        var r       = entry.row;
        var lblCell = counters.getRange(r, colLabel);
        var valCell = counters.getRange(r, colValue);
        lblCell.setFontWeight("bold").setBackground(COLOR_LABEL);
        valCell.setFontWeight("normal").setBackground(COLOR_VALUE);
        lblCell.setBorder(true, true, true, true, null, null, "#B0C4DE", BORDER_SOLID);
        valCell.setBorder(true, true, true, true, null, null, "#B0C4DE", BORDER_SOLID);
      });
    }

    // Helper: draw a medium outer border around a range
    function outerBorder(range) {
      range.setBorder(true, true, true, true, null, null, "#4A86C8", BORDER_MEDIUM);
    }

    // ── Column A/B — Counter sections ────────────────────────
    formatSection([{ row: 1, type: "label" }], 1, 2);
    outerBorder(counters.getRange("A1:B1"));

    formatSection([{ row: 3, type: "label" }, { row: 4, type: "label" }], 1, 2);
    outerBorder(counters.getRange("A3:B4"));

    formatSection([{ row: 6, type: "label" }, { row: 7, type: "label" }], 1, 2);
    outerBorder(counters.getRange("A6:B7"));

    formatSection([{ row: 9, type: "label" }, { row: 10, type: "label" }], 1, 2);
    outerBorder(counters.getRange("A9:B10"));

    formatSection([{ row: 12, type: "label" }], 1, 2);
    outerBorder(counters.getRange("A12:B12"));

    // Grey out skipped rows in A/B
    [2, 5, 8, 11].forEach(function(r) {
      counters.getRange(r, 1, 1, 2).setBackground(COLOR_SKIP);
    });

    // ── Column D/E — Config sections (from COUNTERS_LAYOUT) ──
    CFG_BLOCKS.forEach(function(b, i) {
      var rows = [];
      for (var rr = b.first; rr <= b.last; rr++) rows.push({ row: rr });
      formatSection(rows, 4, 5);
      outerBorder(counters.getRange("D" + b.first + ":E" + b.last));
      // grey spacer row below every block except the last
      if (i < CFG_BLOCKS.length - 1) {
        counters.getRange(b.last + 1, 4, 1, 2).setBackground(COLOR_SKIP);
      }
    });

    // Set column widths for readability
    counters.setColumnWidth(1, 280); // A — counter labels
    counters.setColumnWidth(2, 120); // B — counter values
    counters.setColumnWidth(3, 30);  // C — spacer
    counters.setColumnWidth(4, 220); // D — config labels
    counters.setColumnWidth(5, 300); // E — config values

    Logger.log("Counters: formatting applied.");
    Logger.log("Counters: scaffold written.");
  } else {
    Logger.log("Counters: already populated, skipping.");
  }

  try {
  // ── LOCK SCRIPT / FORMULA-MANAGED COUNTER CELLS ───────────
  // B1, B3, B4, B6, B7, B9, B10 and B12 are written by the script or by
  // formulas and must not be edited in the sheet UI. Only the account
  // running the script (the owner) keeps edit rights, so the timed
  // triggers, which run as that account, can still write to them.
  // Re-running is safe: ranges that are already protected are skipped.
  var lockedRanges = ["B1", "B3:B4", "B6:B7", "B9:B10", "B12"];
  var lockDescription = "Script-managed counter - do not edit";
  var existingLocks = {};
  counters.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p) {
    if (p.getDescription() === lockDescription) {
      existingLocks[p.getRange().getA1Notation()] = true;
    }
  });

  var lockOwner = Session.getEffectiveUser();
  lockedRanges.forEach(function(a1) {
    if (existingLocks[a1]) {
      Logger.log("Counters " + a1 + ": already locked, skipping.");
      return;
    }
    var protection = counters.getRange(a1).protect().setDescription(lockDescription);
    protection.addEditor(lockOwner);
    protection.removeEditors(protection.getEditors().filter(function(u) {
      return u.getEmail() !== lockOwner.getEmail();
    }));
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
    Logger.log("Counters " + a1 + ": locked.");
  });
  } catch (lockErr) {
    Logger.log("Locking counter cells failed, continuing: " + lockErr);
  }

  // Layout check (warns if an older Counters tab does not match)
  checkCountersLayout(counters);

  try {
  // ── VALUES FROM THE REFERENCES TAB ────────────────────────
  // Show name/prefix/date, ticket price/limit, bank, GCash number,
  // PayPal link, QR folder ID, social URLs and email sender name.
  var references = readReferences(spreadsheet);
  if (references) {
    fillCountersFromReferences(counters, references);
  } else {
    Logger.log('References tab not found, Counters values were not filled from it.');
  }
  } catch (refErr) {
    Logger.log("Filling Counters from References failed, continuing: " + refErr);
  }

  // ── IMAGE LINKS (from show_assets / global_assets) ────────
  // Fills the link cells in Counters column E. A cell is only
  // written when it is empty, so manual edits are never overwritten.
  var assetIds = getAssetImageIds(spreadsheet);

  var linkCells = [
    { id: "headerLink",       image: "header"     },
    { id: "bannerLink",       image: "banner"     },
    { id: "bankQR",           image: "bank-qr"    },
    { id: "bankInfographic",  image: "bank-info"  },
    { id: "gcashQR",          image: "gcash-qr"   },
    { id: "gcashInfographic", image: "gcash-info" },
    { id: "mayaQR",           image: "maya-qr"    },
    { id: "mayaInfographic",  image: "maya-info"  },
    { id: "facebookIcon",     image: "facebook"   },
    { id: "instagramIcon",    image: "instagram"  },
    { id: "twitterIcon",      image: "twitter"    }
  ];

  linkCells.forEach(function(entry) {
    var label = CFG_LABEL[entry.id];
    var cell  = counterCell(entry.id);
    var id    = assetIds[entry.image];
    if (!id) {
      Logger.log(label + ": no " + entry.image + ".png/.jpg/.jpeg found, left unchanged.");
      return;
    }
    var target = counters.getRange(cell);
    if (target.getValue() === "" || target.getValue() === null) {
      target.setValue(imgLink + id);
      Logger.log(label + " (" + cell + ") set.");
    } else {
      Logger.log(label + " (" + cell + ") already has a value, skipping.");
    }
  });

  // ── SHARING ───────────────────────────────────────────────
  // The emails embed these images and sendReservationEmails() fetches the
  // QR images without signing in, so every image found is set to
  // "Anyone with the link can view". Only the files actually selected
  // (show_assets preferred over global_assets) are changed.
  Object.keys(assetIds).forEach(function(name) {
    var id = assetIds[name];
    if (!id) return;
    try {
      DriveApp.getFileById(id).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      Logger.log("Sharing set to 'Anyone with the link can view': " + name + " (" + id + ")");
    } catch (err) {
      Logger.log("Could not change sharing for " + name + " (" + id + "): " + err +
                 " — your Workspace admin may restrict public link sharing.");
    }
  });

  Logger.log("setupSheets() complete.");
}


// =============================================================
// SECTION 3 — SHEET / UI UTILITIES
// Small sheet-level helpers. Depend on sheets existing but do
// not depend on CONFIG, so they sit before loadConfig().
// =============================================================

function ensureTrackingRowNumbers() {
  getSheetVariables();
  var last = trackingSheet.getLastRow();
  if (last < 2) return;
  var s2 = trackingSheet.getRange("S2").getValue();
  if (!s2) {
    trackingSheet.getRange("S2").setValue("=ROW()");
  }
  var src  = trackingSheet.getRange("S2");
  var dest = trackingSheet.getRange("S2:S" + last);
  src.autoFill(dest, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function createDropdown() {
  getSheetVariables();
  var last = trackingSheet.getLastRow();
  if (last < 2) return;
  var cell = trackingSheet.getRange("K2:K" + last);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["On Hold", "Payment Confirmed", "Cancelled", "Ticket/s Claimed", "Waitlisted", "On Hold - From Waitlist", "Invalid"])
    .build();
  cell.setDataValidation(rule);
}


// =============================================================
// SECTION 4 — TRIGGER SETUP
// Run createAllTriggers() once manually. Safe to re-run —
// checks for an existing trigger per function before creating.
// =============================================================

function createAllTriggers() {
  var existing = {};
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getEventType() === ScriptApp.EventType.CLOCK) {
      existing[t.getHandlerFunction()] = true;
    }
  });

  var triggers = [
    // ['sendReminderEmail',     function(b) { return b.everyHours(1);    }],
    ['sendWaitlistedEmails',  function(b) { return b.everyMinutes(5);  }],
    ['setReservation',        function(b) { return b.everyMinutes(15); }],
    ['sendReservationEmails', function(b) { return b.everyMinutes(30); }],
    ['sendConfirmationEmail', function(b) { return b.everyHours(1);    }],
    ['sendCancellationEmail', function(b) { return b.everyHours(1);    }]
  ];

  triggers.forEach(function(entry) {
    var fnName  = entry[0];
    var builder = entry[1];
    if (existing[fnName]) {
      Logger.log("Trigger already exists, skipping: " + fnName);
    } else {
      builder(ScriptApp.newTrigger(fnName).timeBased()).create();
      Logger.log("Trigger created: " + fnName);
    }
  });
}


// =============================================================
// SECTION 5 — CONFIG
// Reads all configuration from Counters column E in one batch
// (rows defined by COUNTERS_LAYOUT).
// Called explicitly at the top of every core function.
// =============================================================

function loadConfig() {
  getSheetVariables();
  // Single range read — configValues[n][0] is Counters row n+1.
  // Row numbers come from COUNTERS_LAYOUT (see CFG_ROW), never hardcoded.
  var configValues = summarySheet.getRange(1, 5, CFG_LAST_ROW, 1).getValues();
  function cfg(id) { return configValues[CFG_ROW[id] - 1][0]; }

  CONFIG.headerLink           = cfg("headerLink");
  CONFIG.mailSubject          = cfg("showName");
  CONFIG.bannerLink           = cfg("bannerLink");
  CONFIG.preFix               = cfg("showPrefix");
  CONFIG.ticketPrice          = Number(cfg("ticketPrice"));
  CONFIG.waitlistThreshold    = Number(cfg("ticketLimit"));
  CONFIG.bankName             = cfg("bankName");
  CONFIG.bankAccountName      = cfg("bankAccountName");
  CONFIG.accountNumber        = String(cfg("bankAccountNumber"));
  CONFIG.bankQR               = cfg("bankQR");
  CONFIG.bankInfographic      = cfg("bankInfographic");
  CONFIG.gCashNumber          = String(cfg("gcashNumber"));
  CONFIG.gcashQR              = cfg("gcashQR");
  CONFIG.gcashInfographic     = cfg("gcashInfographic");
  CONFIG.payPalLink           = cfg("paypalLink");
  CONFIG.paypalFee            = Number(cfg("paypalFee"));
  CONFIG.mayaQR               = cfg("mayaQR");
  CONFIG.mayaInfographic      = cfg("mayaInfographic");
  CONFIG.reminderThreshold    = Number(cfg("reminderDays"));
  CONFIG.showDate             = String(cfg("showDate"));
  CONFIG.reservationPeriod    = String(cfg("waitText"));
  CONFIG.resAllowance         = Number(cfg("waitDays"));
  CONFIG.startRow             = Number(cfg("startRow"));
  CONFIG.earlyBirdEnabledRef  = counterCell("earlyBirdEnabled"); // cell address for read/write-back
  CONFIG.earlyBirdThreshold   = Number(cfg("earlyBirdLimit"));
  CONFIG.earlyBirdTicketPrice = Number(cfg("earlyBirdPrice"));
  CONFIG.webAppURL            = cfg("webAppUrl");
  CONFIG.qrFolderID           = cfg("qrFolderId");
  CONFIG.tplBank              = cfg("tplBank");
  CONFIG.tplGCash             = cfg("tplGCash");
  CONFIG.tplPayPal            = cfg("tplPayPal");
  CONFIG.tplWaitlisted        = cfg("tplWaitlisted");
  CONFIG.tplReminder          = cfg("tplReminder");
  CONFIG.tplConfirmed         = cfg("tplConfirmed");
  CONFIG.tplCancelled         = cfg("tplCancelled");
  CONFIG.tplTicketPDF         = cfg("tplTicketPDF");
  // Social image links (icons), built from imgLink + file ID
  CONFIG.facebookIcon         = cfg("facebookIcon");
  CONFIG.instagramIcon        = cfg("instagramIcon");
  CONFIG.twitterIcon          = cfg("twitterIcon");
  // Social profile URLs
  CONFIG.facebookUrl          = cfg("facebookUrl");
  CONFIG.instagramUrl         = cfg("instagramUrl");
  CONFIG.twitterUrl           = cfg("twitterUrl");
  CONFIG.emailSenderName      = cfg("emailSenderName");

  // PH mobile numbers are 11 digits; prepend 0 if stored without it
  if (CONFIG.gCashNumber && CONFIG.gCashNumber.length === 10) {
    CONFIG.gCashNumber = "0" + CONFIG.gCashNumber;
  }

  // Derived
  CONFIG.qrFolder = DriveApp.getFolderById(CONFIG.qrFolderID);

  // Regex patterns
  CONFIG.regexBank   = /(bank|maya)/i;
  CONFIG.regexMaya   = /maya/i;
  CONFIG.regexGCash  = /gcash/i;
  CONFIG.regexPaypal = /paypal/i;

  // Resolve sending alias (prefer one containing "tickets")
  CONFIG.me    = Session.getActiveUser().getEmail();
  CONFIG.alias = CONFIG.me;
  try {
    var aliases = GmailApp.getAliases();
    if (aliases && aliases.length > 0) {
      for (var a = 0; a < aliases.length; a++) {
        var al = String(aliases[a]);
        if (al.toLowerCase().indexOf("tickets") !== -1) {
          CONFIG.alias = al;
          break;
        }
      }
    }
  } catch (err) {
    CONFIG.alias = CONFIG.me;
  }
}


// =============================================================
// SECTION 6 — SMALL UTILITIES
// Pure helpers with no sheet or config dependencies.
// =============================================================

function pad(d) {
  return (d < 10) ? '0' + d.toString() : d.toString();
}

function lpad(value, padding) {
  var zeroes = new Array(padding + 1).join("0");
  return (zeroes + value).slice(-padding);
}

function getCustomDateTime() {
  var d       = new Date();
  var month   = d.getMonth() + 1;
  var day     = d.getDate();
  var year    = d.getFullYear();
  var hours   = pad(d.getHours());
  var minutes = pad(d.getMinutes());
  var seconds = pad(d.getSeconds());
  return month + "/" + day + "/" + year + " " + hours + ":" + minutes + ":" + seconds;
}

function getPaymentDeadline() {
  var d = new Date();
  d.setDate(d.getDate() + CONFIG.resAllowance);
  return (d.getMonth() + 1) + "/" + d.getDate() + "/" + d.getFullYear() + " 10:00 PM";
}

// Replaces {placeholder} tokens. The social placeholders are always
// available (from CONFIG, so loadConfig() must have run):
//   {plcFacebookLink} / {plcInstagramLink} / {plcTwitterLink}  -> profile URL
//   {plcFacebookIcon} / {plcInstagramIcon} / {plcTwitterIcon}  -> icon image link
//   {plcPaypalLink}                                            -> PayPal link (Counters)
//   {plcPaypalTransactionFee}                                  -> PayPal transaction fee per ticket (Counters)
// Keys in `map` win over these defaults.
function tpl(html, map) {
  var all = {
    "plcFacebookLink":  CONFIG.facebookUrl,
    "plcFacebookIcon":  CONFIG.facebookIcon,
    "plcInstagramLink": CONFIG.instagramUrl,
    "plcInstagramIcon": CONFIG.instagramIcon,
    "plcTwitterLink":   CONFIG.twitterUrl,
    "plcTwitterIcon":   CONFIG.twitterIcon,
    "plcPaypalLink":    CONFIG.payPalLink,
    "plcPaypalTransactionFee": CONFIG.paypalFee
  };
  for (var m in map) {
    if (map.hasOwnProperty(m)) all[m] = map[m];
  }

  for (var k in all) {
    if (all.hasOwnProperty(k)) {
      var value = (all[k] === undefined || all[k] === null) ? "" : String(all[k]);
      // function replacer: "$" in a value is never treated as a special pattern
      html = html.replace(new RegExp("\\{" + k + "\\}", "g"), function() { return value; });
    }
  }
  return html;
}

// Fetches HTML template content from a Drive file by ID.
// File IDs come from the References tab, are stored in Counters
// the Counters template rows by setupSheets() and loaded into CONFIG by loadConfig(). Throws clearly if the ID is missing
// so misconfiguration is easy to diagnose in the logs.
function readTemplate(fileId) {
  if (!fileId) {
    throw new Error("readTemplate: file ID is empty — check the Counters template rows.");
  }
  return DriveApp.getFileById(fileId).getBlob().getDataAsString();
}

function randomStr(m) {
  var length = m || 15;
  var s = '';
  var r  = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  for (var i = 0; i < length; i++) { s += r.charAt(Math.floor(Math.random() * r.length)); }
  return s;
}


// =============================================================
// SECTION 7 — SHEET HELPERS
// Helpers that read/write sheets and may depend on CONFIG.
// =============================================================

function writeColumnValues(sheet, col, startRow, values) {
  if (!values || values.length === 0) return;
  var twoD = values.map(function(v) { return [v]; });
  sheet.getRange(startRow, col, twoD.length, 1).setValues(twoD);
}

function fileExistsInFolder(fileName) {
  var folder = DriveApp.getFolderById(CONFIG.qrFolderID);
  var files  = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next().getId() : null;
}

function deleteFilesByName(searchString) {
  try {
    var files = CONFIG.qrFolderID
      ? DriveApp.getFolderById(CONFIG.qrFolderID).getFilesByName(searchString)
      : DriveApp.getFilesByName(searchString);
    while (files.hasNext()) {
      var file = files.next();
      try {
        file.setTrashed(true);
        Logger.log("Deleted: " + file.getName() + " | ID: " + file.getId());
      } catch (fileError) {
        Logger.log("Failed to delete: " + file.getName() + " | " + fileError);
      }
    }
  } catch (error) {
    Logger.log("deleteFilesByName error: " + error);
  }
}

function renameFilesWithError(searchString) {
  try {
    var files    = CONFIG.qrFolderID
      ? DriveApp.getFolderById(CONFIG.qrFolderID).getFiles()
      : DriveApp.getFiles();
    var dateTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd-HH:mm:ss");
    while (files.hasNext()) {
      var file = files.next();
      try {
        var fileName = file.getName();
        if (fileName.indexOf(searchString) !== -1) {
          if (fileName.indexOf("ERROR") === 0) {
            Logger.log("Skipping already-errored file: " + fileName);
          } else {
            file.setName("ERROR-" + dateTime + "-" + fileName);
            Logger.log("Renamed: " + file.getName() + " | ID: " + file.getId());
          }
        }
      } catch (fileError) {
        Logger.log("Failed to rename: " + file.getName() + " | " + fileError);
      }
    }
  } catch (error) {
    Logger.log("renameFilesWithError error: " + error);
  }
}

// Finds the row in the Tickets sheet whose ticketId (col J)
// matches the given value and sets status (col G) to "Invalid".
function markTicketInvalid(ticketId) {
  try {
    var last = ticketSheet.getLastRow();
    if (last < 2) return;
    var data = ticketSheet.getRange(2, 1, last - 1, 10).getValues();
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][9]) === String(ticketId)) {
        ticketSheet.getRange(r + 2, 7).setValue("Invalid");
        Logger.log("Marked ticket as Invalid in Tickets sheet: " + ticketId);
        return;
      }
    }
    Logger.log("markTicketInvalid: no Tickets sheet row found for: " + ticketId);
  } catch (err) {
    Logger.log("markTicketInvalid error for " + ticketId + ": " + err);
  }
}

// Searches the Tickets sheet for a row matching ticketId (col J)
// and status (col G). Returns the 1-based row number or null.
function findTicketSheetRow(ticketId, status) {
  try {
    var last = ticketSheet.getLastRow();
    if (last < 2) return null;
    var data = ticketSheet.getRange(2, 1, last - 1, 10).getValues();
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][9]) === String(ticketId) &&
          String(data[r][6]) === String(status)) {
        return r + 2;
      }
    }
  } catch (err) {
    Logger.log("findTicketSheetRow error for " + ticketId + ": " + err);
  }
  return null;
}

function createQRCode(val, fileName) {
  var maxRetries = 2;
  var delayMs    = 5000;
  if (!val) {
    Logger.log("createQRCode error: val is empty or null");
    return null;
  }
  for (var attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      var qrUrl    = "https://quickchart.io/qr?text=" + encodeURIComponent(val) + "&size=300&ecLevel=Q";
      var response = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        throw new Error("HTTP " + response.getResponseCode() + " from QuickChart API");
      }
      var imageData = response.getBlob();
      if (!imageData.getContentType().startsWith("image/")) {
        throw new Error("Invalid content type: " + imageData.getContentType());
      }
      var file = CONFIG.qrFolder.createFile(imageData);
      file.setName(fileName + ".png");
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getId();
    } catch (error) {
      Logger.log("createQRCode attempt " + attempt + " failed: " + error);
      if (attempt < maxRetries) {
        Logger.log("Retrying in 5 seconds...");
        Utilities.sleep(delayMs);
      } else {
        Logger.log("createQRCode failed after " + maxRetries + " attempts.");
        return null;
      }
    }
  }
}

function imgToHtmlTag(fileId) {
  if (!fileId) return "";
  var file     = DriveApp.getFileById(fileId);
  var fileBlob = file.getBlob();
  return "data:" + fileBlob.getContentType() + ';base64,' + Utilities.base64Encode(fileBlob.getBytes());
}


// =============================================================
// SECTION 8 — CORE LOGIC
// =============================================================


// =============================================================
// SECTION 8.1 — UPDATE WAITLIST
// Promotes notified waitlisted rows to "On Hold - From Waitlist"
// while seats are available. Each promotion is written straight
// to its own row (column K) as soon as it is decided.
// =============================================================
 
function updateWaitlist() {
  loadConfig();
  var last = trackingSheet.getLastRow();
  if (last < CONFIG.startRow) return;
  var numRows = last - CONFIG.startRow + 1;
  var table   = trackingSheet.getRange(CONFIG.startRow, 1, numRows, COL_S).getValues();
 
  // Read the seat counter once and track it locally, so we do not
  // depend on the sheet formulas recalculating between row writes.
  var seatsUsed = Number(summarySheet.getRange('B9').getValue());
  var changed   = false;
 
  for (var i = 0; i < table.length; i++) {
    var status   = table[i][COL_K - 1];
    var sentFlag = table[i][COL_P - 1];
 
    if (String(status) === 'Waitlisted' && String(sentFlag) === 'SENT') {
      if (seatsUsed < CONFIG.waitlistThreshold) {
        trackingSheet.getRange(CONFIG.startRow + i, COL_K)
                     .setValue("On Hold - From Waitlist");
        seatsUsed += Number(table[i][COL_G - 1]) || 0;
        changed = true;
      }
    }
  }
 
  // Make sure the Counters formulas (B9 etc.) reflect the new
  // statuses before the caller reads them.
  if (changed) SpreadsheetApp.flush();
}
 
 
// =============================================================
// SECTION 8.2 — SET RESERVATION
// Assigns Reservation ID, Total Price and Status (and the Early
// Bird flag) to new rows. Values are written per row, together
// with the counters, so a crash or timeout midway leaves the
// sheet consistent: finished rows are complete and the counters
// match them, unfinished rows are still blank and get picked up
// on the next run.
// =============================================================
 
function setReservation() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log("setReservation could not acquire lock: " + e);
    return;
  }
 
  try {
    ensureTrackingRowNumbers();
    loadConfig();
    updateWaitlist();
 
    var last = trackingSheet.getLastRow();
    if (last < CONFIG.startRow) return;
    var numRows = last - CONFIG.startRow + 1;
    var table   = trackingSheet.getRange(CONFIG.startRow, 1, numRows, COL_S).getValues();
 
    var anyDropdown = false;
 
    // Counters: read once, then keep local copies and write them
    // back right after each row is processed.
    var newCount         = Number(summarySheet.getRange('B1').getValue());
    var earlyBirdEnabled = summarySheet.getRange(CONFIG.earlyBirdEnabledRef).getValue();
    var earlyBirdCount   = Number(summarySheet.getRange('B12').getValue());
 
    // Seats used (on hold + confirmed + unsent cancellations).
    // Read after updateWaitlist() has flushed its changes, then
    // tracked locally as new "On Hold" rows are assigned.
    var seatsUsed = Number(summarySheet.getRange('B9').getValue());
 
    for (var i = 0; i < table.length; i++) {
      var numberTicket     = Number(table[i][COL_G - 1]);
      var payOption        = table[i][COL_H - 1];
      var reservationId    = table[i][COL_I - 1];
      var totalTicketPrice = table[i][COL_J - 1];
 
      if ((reservationId === '' || reservationId === null) &&
          (totalTicketPrice === '' || totalTicketPrice === null)) {
 
        var sheetRow = CONFIG.startRow + i;
        anyDropdown  = true;
 
        newCount++;
        var resId = CONFIG.preFix + lpad(newCount, 6);
 
        // Early bird threshold check
        if (String(earlyBirdEnabled) === "Yes" && earlyBirdCount >= CONFIG.earlyBirdThreshold) {
          summarySheet.getRange(CONFIG.earlyBirdEnabledRef).setValue("Exceeded");
          earlyBirdEnabled = "Exceeded";
        }
 
        var finalTicketPrice;
        var isEarlyBird = false;
        if (String(earlyBirdEnabled) === "Yes") {
          finalTicketPrice = CONFIG.earlyBirdTicketPrice;
          earlyBirdCount  += numberTicket;
          isEarlyBird      = true;
        } else {
          finalTicketPrice = CONFIG.ticketPrice;
        }
 
        var grossTicketPrice = CONFIG.regexPaypal.test(String(payOption))
          ? finalTicketPrice + CONFIG.paypalFee
          : finalTicketPrice;
 
        var totalPrice = grossTicketPrice * numberTicket;
 
        var newStatus;
        if (seatsUsed < CONFIG.waitlistThreshold) {
          newStatus  = "On Hold";
          seatsUsed += numberTicket;
        } else {
          newStatus = "Waitlisted";
        }
 
        // 1) Counters first: if we fail after this point we only lose
        //    an ID number (a gap), never reuse one (a duplicate).
        summarySheet.getRange('B1').setValue(newCount);
        if (isEarlyBird) summarySheet.getRange('B12').setValue(earlyBirdCount);
 
        // 2) Row values. I:K (ID, Total, Status) are adjacent, so one write.
        trackingSheet.getRange(sheetRow, COL_I, 1, 3)
                     .setValues([[resId, totalPrice, newStatus]]);
        if (isEarlyBird) {
          trackingSheet.getRange(sheetRow, COL_R).setValue("Yes");
        }
      }
    }
 
    if (anyDropdown) createDropdown();
 
  } finally {
    lock.releaseLock();
  }
}


// =============================================================
// SECTION 8.3 — SEND RESERVATION EMAILS
// =============================================================

function sendReservationEmails() {
  loadConfig();

  var last = trackingSheet.getLastRow();
  if (last < CONFIG.startRow) return;
  var numRows = last - CONFIG.startRow + 1;
  var table   = trackingSheet.getRange(CONFIG.startRow, 1, numRows, COL_S).getValues();

  var colL = table.map(function(r) { return r[COL_L - 1]; });
  var colM = table.map(function(r) { return r[COL_M - 1]; });

  var tplBank   = readTemplate(CONFIG.tplBank);
  var tplGcash  = readTemplate(CONFIG.tplGCash);
  var tplPaypal = readTemplate(CONFIG.tplPayPal);

  for (var i = 0; i < table.length; i++) {
    var firstName     = table[i][COL_B - 1];
    var lastName      = table[i][COL_C - 1];
    var email         = table[i][COL_D - 1];
    var phone         = table[i][COL_E - 1];
    var numberTickets = table[i][COL_G - 1];
    var payOption     = table[i][COL_H - 1];
    var reservationId = table[i][COL_I - 1];
    var total         = table[i][COL_J - 1];
    var status        = table[i][COL_K - 1];
    var sentFlag      = table[i][COL_L - 1];

    if (!reservationId || !payOption) continue;
    if (!(status === 'On Hold' || status === 'On Hold - From Waitlist')) continue;
    if (sentFlag && String(sentFlag).length > 0) continue;

    var subjPrefix = (status === "On Hold") ? "Reservation Received" : "Reservation Opened";
    var subj       = reservationId + " - " + subjPrefix + ": " + CONFIG.mailSubject;
    var html       = "";
    var qrCodeBlob = null;

    if (CONFIG.regexBank.test(String(payOption))) {
      // Maya uses the same email template as bank transfers but with its own
      // QR and infographic. A label such as "Bank Transfer (Maya Bank)" still
      // counts as a bank transfer, so "bank" wins over "maya".
      var isMaya = CONFIG.regexMaya.test(String(payOption)) && !/bank/i.test(String(payOption));
      html = tpl(tplBank, {
        "plcBankName":        CONFIG.bankName,
        "plcBankAccountName": CONFIG.bankAccountName,
        "plcAccountNumber":   CONFIG.accountNumber,
        "plcInfographic":     isMaya ? CONFIG.mayaInfographic : CONFIG.bankInfographic
      });
      try {
        qrCodeBlob = isMaya
          ? UrlFetchApp.fetch(CONFIG.mayaQR).getBlob().setName("maya-QR.jpg")
          : UrlFetchApp.fetch(CONFIG.bankQR).getBlob().setName("bank-QR.jpg");
      }
      catch (err) { Logger.log("Failed to fetch " + (isMaya ? "mayaQR" : "bankQR") + ": " + err); }
    } else if (CONFIG.regexGCash.test(String(payOption))) {
      html = tpl(tplGcash, {
        "plcGCashNumber": CONFIG.gCashNumber,
        "plcInfographic": CONFIG.gcashInfographic
      });
      try { qrCodeBlob = UrlFetchApp.fetch(CONFIG.gcashQR).getBlob().setName("GCash-QR.jpg"); }
      catch (err) { Logger.log("Failed to fetch gcashQR: " + err); }
    } else {
      html = tplPaypal;
    }

    var fullName   = (firstName || "") + " " + (lastName || "");
    var isEarly    = (table[i][COL_R - 1] === "Yes");
    var finalPrice = isEarly ? CONFIG.earlyBirdTicketPrice : CONFIG.ticketPrice;

    html = tpl(html, {
      "plcHeader":          CONFIG.headerLink,
      "plcBanner":          CONFIG.bannerLink,
      "plcReservationId":   reservationId,
      "plcFullName":        fullName,
      "plcName":            fullName,
      "plcEmail":           email || "",
      "plcPhone":           phone || "",
      "plcTicketPrice":     finalPrice,
      "plcNumberofTickets": numberTickets,
      "plcPaymentMethod":   payOption,
      "plcWaitingTime":     CONFIG.reservationPeriod,
      "plcDeadline":        getPaymentDeadline(),
      "plcTotal":           (total !== '' && total !== null) ? (total + ".00") : ""
    });

    try {
      var mailOptions = { htmlBody: html, name: CONFIG.emailSenderName, from: CONFIG.alias };
      if (qrCodeBlob && !CONFIG.regexPaypal.test(String(payOption))) {
        mailOptions.attachments = [qrCodeBlob];
      }
      GmailApp.sendEmail(email, subj, html, mailOptions);
      colL[i] = "SENT";
      colM[i] = getCustomDateTime();
    } catch (err) {
      Logger.log("sendReservationEmails failed for row " + (CONFIG.startRow + i) + ": " + err);
    }
  }

  writeColumnValues(trackingSheet, COL_L, CONFIG.startRow, colL);
  writeColumnValues(trackingSheet, COL_M, CONFIG.startRow, colM);
}


// =============================================================
// SECTION 8.4 — SEND CONFIRMATION EMAILS
// =============================================================

function sendConfirmationEmail() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    Logger.log("sendConfirmationEmail could not acquire lock: " + e);
    return;
  }

  try {
    loadConfig();

    var last = trackingSheet.getLastRow();
    if (last < CONFIG.startRow) return;
    var numRows = last - CONFIG.startRow + 1;
    var table   = trackingSheet.getRange(CONFIG.startRow, 1, numRows, COL_S).getValues();
    var colO    = table.map(function(r) { return r[COL_O - 1]; });

    var tplTicketConfirmed = readTemplate(CONFIG.tplConfirmed);

    for (var i = 0; i < table.length; i++) {
      var firstName            = table[i][COL_B - 1];
      var lastName             = table[i][COL_C - 1];
      var email                = table[i][COL_D - 1];
      var phone                = table[i][COL_E - 1];
      var numberTickets        = Number(table[i][COL_G - 1]) || 0;
      var reservationId        = table[i][COL_I - 1];
      var status               = table[i][COL_K - 1];
      var reservationEmailSent = table[i][COL_L - 1];
      var confirmationSentFlag = table[i][COL_O - 1];

      if (!reservationId) continue;
      if (!reservationEmailSent) continue;
      if (!(status === 'Payment Confirmed')) continue;
      if (confirmationSentFlag && String(confirmationSentFlag).length > 0) continue;

      try {
        var attachmentArr    = [];
        var proceedWithEmail = 0;

        for (var t = 0; t < numberTickets; t++) {
          try {
            var ticketCtr   = t + 1;
            var ticketId    = reservationId + '-' + ticketCtr;
            var pdfFileName = ticketId + ".pdf";
            var pngFileName = ticketId + ".png";

            var existingPDF = fileExistsInFolder(pdfFileName);

            if (!existingPDF) {
              // ── New ticket: check for orphaned PNG first ──
              var existingPNG = fileExistsInFolder(pngFileName);
              if (existingPNG) renameFilesWithError(pngFileName);

              var uIDBase    = reservationId || "";
              var mid        = Math.floor(uIDBase.length / 2);
              var uIDStr     = uIDBase.slice(0, mid) + randomStr(4) + uIDBase.slice(mid) + ticketCtr.toString();
              var uUID       = Utilities.base64EncodeWebSafe(uIDStr);
              var qrID       = createQRCode(uUID, ticketId);

              if (qrID) {
                var pdfTemplate   = readTemplate(CONFIG.tplTicketPDF);
                var headerIdMatch = String(CONFIG.headerLink).match(/(?<=id=).*$/);
                var bannerIdMatch = String(CONFIG.bannerLink).match(/(?<=id=).*$/);

                pdfTemplate = tpl(pdfTemplate, {
                  'plcHeader':        imgToHtmlTag(headerIdMatch ? headerIdMatch[0] : null),
                  'plcBanner':        imgToHtmlTag(bannerIdMatch ? bannerIdMatch[0] : null),
                  'plcQRImg':         imgToHtmlTag(qrID),
                  'plcTicketIdBig':   ticketId,
                  'plcReservationId': reservationId,
                  'plcTicketNumber':  ticketCtr + " of " + numberTickets,
                  'plcFullName':      (firstName || "") + " " + (lastName || ""),
                  'plcEmail':         email || "",
                  'plcPhone':         phone || ""
                });

                var blob = HtmlService.createHtmlOutput(pdfTemplate).getBlob();
                var file = CONFIG.qrFolder.createFile(blob.getAs("application/pdf"));
                file.setName(pdfFileName);
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

                // Verify PDF is queryable in Drive before proceeding
                var confirmedPDF = fileExistsInFolder(pdfFileName);
                if (!confirmedPDF) {
                  Logger.log("PDF creation could not be confirmed for: " + pdfFileName);
                  proceedWithEmail++;
                  renameFilesWithError(pdfFileName);
                  break;
                }

                var pdfBlob = UrlFetchApp.fetch(
                  "https://drive.google.com/uc?export=view&id=" + file.getId()
                ).getBlob().setName(pdfFileName);
                attachmentArr.push(pdfBlob);

                // Write to Tickets sheet only after PDF is confirmed in Drive
                var writeRow = ticketSheet.getLastRow() + 1;
                ticketSheet.getRange(writeRow, 1, 1, 10).setValues([[
                  firstName, lastName, email, phone, numberTickets,
                  reservationId, status, uUID, ticketCtr, ticketId
                ]]);

              } else {
                proceedWithEmail++;
              }

            } else {
              // ── Existing PDF found: validate against Tickets sheet ──
              var ticketSheetRow = findTicketSheetRow(ticketId, "Payment Confirmed");

              if (!ticketSheetRow) {
                // Stale or orphaned PDF — tag it and regenerate
                Logger.log("Existing PDF found but no valid Tickets sheet row for: " + ticketId + " — regenerating.");
                renameFilesWithError(pdfFileName);

                // Check for orphaned PNG before regenerating QR
                var existingPNG2 = fileExistsInFolder(pngFileName);
                if (existingPNG2) renameFilesWithError(pngFileName);

                var uIDBase2   = reservationId || "";
                var mid2       = Math.floor(uIDBase2.length / 2);
                var uIDStr2    = uIDBase2.slice(0, mid2) + randomStr(4) + uIDBase2.slice(mid2) + ticketCtr.toString();
                var uUID2      = Utilities.base64EncodeWebSafe(uIDStr2);
                var qrID2      = createQRCode(uUID2, ticketId);

                if (qrID2) {
                  var pdfTemplate2   = readTemplate(CONFIG.tplTicketPDF);
                  var headerIdMatch2 = String(CONFIG.headerLink).match(/(?<=id=).*$/);
                  var bannerIdMatch2 = String(CONFIG.bannerLink).match(/(?<=id=).*$/);

                  pdfTemplate2 = tpl(pdfTemplate2, {
                    'plcHeader':        imgToHtmlTag(headerIdMatch2 ? headerIdMatch2[0] : null),
                    'plcBanner':        imgToHtmlTag(bannerIdMatch2 ? bannerIdMatch2[0] : null),
                    'plcQRImg':         imgToHtmlTag(qrID2),
                    'plcTicketIdBig':   ticketId,
                    'plcReservationId': reservationId,
                    'plcTicketNumber':  ticketCtr + " of " + numberTickets,
                    'plcFullName':      (firstName || "") + " " + (lastName || ""),
                    'plcEmail':         email || "",
                    'plcPhone':         phone || ""
                  });

                  var blob2 = HtmlService.createHtmlOutput(pdfTemplate2).getBlob();
                  var file2 = CONFIG.qrFolder.createFile(blob2.getAs("application/pdf"));
                  file2.setName(pdfFileName);
                  file2.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

                  var confirmedPDF2 = fileExistsInFolder(pdfFileName);
                  if (!confirmedPDF2) {
                    Logger.log("Regenerated PDF could not be confirmed for: " + pdfFileName);
                    proceedWithEmail++;
                    renameFilesWithError(pdfFileName);
                    break;
                  }

                  var pdfBlob2 = UrlFetchApp.fetch(
                    "https://drive.google.com/uc?export=view&id=" + file2.getId()
                  ).getBlob().setName(pdfFileName);
                  attachmentArr.push(pdfBlob2);

                  var writeRow2 = ticketSheet.getLastRow() + 1;
                  ticketSheet.getRange(writeRow2, 1, 1, 10).setValues([[
                    firstName, lastName, email, phone, numberTickets,
                    reservationId, status, uUID2, ticketCtr, ticketId
                  ]]);

                } else {
                  Logger.log("QR code generation failed during PDF regeneration for: " + ticketId);
                  proceedWithEmail++;
                  break;
                }

              } else {
                // Tickets sheet row is valid — reuse the existing PDF
                var pdfBlob3 = UrlFetchApp.fetch(
                  "https://drive.google.com/uc?export=view&id=" + existingPDF.toString()
                ).getBlob().setName(pdfFileName);
                attachmentArr.push(pdfBlob3);
              }
            }

          } catch (err) {
            Logger.log("Ticket generation error at row " + (CONFIG.startRow + i) + ", ticket " + (t + 1) + ": " + err);
            proceedWithEmail++;
            renameFilesWithError(pdfFileName);
            renameFilesWithError(pngFileName);
            markTicketInvalid(ticketId);
            // `break` is intentional: an incomplete ticket set must not
            // be emailed. col O stays blank so the next run retries.
            break;
          }

        } // end per-ticket loop

        Logger.log('Proceed with Email Flag: ' + proceedWithEmail);

        if (proceedWithEmail === 0) {
          var subject           = reservationId + " - Ticket Purchase Confirmed: " + CONFIG.mailSubject;
          var htmlEmailTemplate = tpl(tplTicketConfirmed, {
            'plcShowDate': CONFIG.showDate,
            'plcHeader':   CONFIG.headerLink,
            'plcBanner':   CONFIG.bannerLink
          });
          GmailApp.sendEmail(email, subject, htmlEmailTemplate, {
            htmlBody:    htmlEmailTemplate,
            name:        CONFIG.emailSenderName,
            from:        CONFIG.alias,
            attachments: attachmentArr
          });

          colO[i] = "SENT";
        }

      } catch (err) {
        // `continue` on the outer reservation loop — log and move on to
        // the next reservation rather than aborting the entire run.
        Logger.log("sendConfirmationEmail error at row " + (CONFIG.startRow + i) + ": " + err);
        continue;
      }

    } // end per-reservation loop

    writeColumnValues(trackingSheet, COL_O, CONFIG.startRow, colO);

  } finally {
    lock.releaseLock();
  }
}


// =============================================================
// SECTION 8.5 — SEND REMINDER EMAILS
// =============================================================

function sendReminderEmail() {
  loadConfig();

  var last = trackingSheet.getLastRow();
  if (last < CONFIG.startRow) return;
  var numRows = last - CONFIG.startRow + 1;
  var table   = trackingSheet.getRange(CONFIG.startRow, 1, numRows, COL_S).getValues();

  var colN        = table.map(function(r) { return r[COL_N - 1]; });
  var now         = new Date().getTime();
  var tplReminder = readTemplate(CONFIG.tplReminder);

  for (var i = 0; i < table.length; i++) {
    var firstName            = table[i][COL_B - 1];
    var lastName             = table[i][COL_C - 1];
    var email                = table[i][COL_D - 1];
    var phone                = table[i][COL_E - 1];
    var numberTickets        = table[i][COL_G - 1];
    var payOption            = table[i][COL_H - 1];
    var reservationId        = table[i][COL_I - 1];
    var total                = table[i][COL_J - 1];
    var status               = table[i][COL_K - 1];
    var reservationEmailSent = table[i][COL_L - 1];
    var payDeadlineSentTime  = table[i][COL_M - 1];
    var reminderSentFlag     = table[i][COL_N - 1];

    if (!reservationEmailSent) continue;
    if (!reservationId || !payDeadlineSentTime) continue;
    if (!(status === 'On Hold' || status === 'On Hold - From Waitlist')) continue;
    if (reminderSentFlag && String(reminderSentFlag).length > 0) continue;

    var daysDiff = Math.floor(
      Math.abs(now - new Date(payDeadlineSentTime).getTime()) / (1000 * 60 * 60 * 24)
    );

    if (daysDiff > CONFIG.reminderThreshold) {
      var subj = reservationId + " - Reservation Reminder: " + CONFIG.mailSubject;
      var html = tpl(tplReminder, {
        'plcHeader':          CONFIG.headerLink,
        'plcBanner':          CONFIG.bannerLink,
        'plcReservationId':   reservationId,
        'plcFullName':        (firstName || "") + " " + (lastName || ""),
        'plcName':            (firstName || "") + " " + (lastName || ""),
        'plcEmail':           email || "",
        'plcPhone':           phone || "",
        'plcTicketPrice':     (table[i][COL_R - 1] === "Yes") ? CONFIG.earlyBirdTicketPrice : CONFIG.ticketPrice,
        'plcNumberofTickets': numberTickets,
        'plcPaymentMethod':   payOption,
        'plcTotal':           (total !== '' && total !== null) ? (total + ".00") : ""
      });

      try {
        GmailApp.sendEmail(email, subj, html, {
          htmlBody: html,
          name:     CONFIG.emailSenderName,
          from:     CONFIG.alias
        });
        colN[i] = "SENT";
      } catch (err) {
        Logger.log("sendReminderEmail failed for row " + (CONFIG.startRow + i) + ": " + err);
      }
    }
  }

  writeColumnValues(trackingSheet, COL_N, CONFIG.startRow, colN);
}


// =============================================================
// SECTION 8.6 — SEND WAITLISTED EMAILS
// =============================================================

function sendWaitlistedEmails() {
  loadConfig();

  var last = trackingSheet.getLastRow();
  if (last < CONFIG.startRow) return;
  var numRows = last - CONFIG.startRow + 1;
  var table   = trackingSheet.getRange(CONFIG.startRow, 1, numRows, COL_S).getValues();

  var colP          = table.map(function(r) { return r[COL_P - 1]; });
  var tplWaitlisted = readTemplate(CONFIG.tplWaitlisted);

  for (var i = 0; i < table.length; i++) {
    var firstName        = table[i][COL_B - 1];
    var lastName         = table[i][COL_C - 1];
    var email            = table[i][COL_D - 1];
    var reservationId    = table[i][COL_I - 1];
    var status           = table[i][COL_K - 1];
    var waitlistSentFlag = table[i][COL_P - 1];

    if (!reservationId || !email) continue;
    if (!(status === 'Waitlisted')) continue;
    if (waitlistSentFlag && String(waitlistSentFlag).length > 0) continue;

    var subj = reservationId + " - Waitlisted: " + CONFIG.mailSubject;
    var html = tpl(tplWaitlisted, {
      'plcHeader':        CONFIG.headerLink,
      'plcBanner':        CONFIG.bannerLink,
      'plcReservationId': reservationId,
      'plcFullName':      (firstName || "") + " " + (lastName || "")
    });

    try {
      GmailApp.sendEmail(email, subj, html, {
        htmlBody: html,
        name:     CONFIG.emailSenderName,
        from:     CONFIG.alias
      });
      colP[i] = "SENT";
    } catch (err) {
      Logger.log("sendWaitlistedEmails failed for row " + (CONFIG.startRow + i) + ": " + err);
    }
  }

  writeColumnValues(trackingSheet, COL_P, CONFIG.startRow, colP);
}


// =============================================================
// SECTION 8.7 — SEND CANCELLATION EMAILS
// =============================================================

function sendCancellationEmail() {
  loadConfig();

  var last = trackingSheet.getLastRow();
  if (last < CONFIG.startRow) return;
  var numRows = last - CONFIG.startRow + 1;
  var table   = trackingSheet.getRange(CONFIG.startRow, 1, numRows, COL_S).getValues();

  var colQ         = table.map(function(r) { return r[COL_Q - 1]; });
  var tplCancelled = readTemplate(CONFIG.tplCancelled);

  for (var i = 0; i < table.length; i++) {
    var firstName            = table[i][COL_B - 1];
    var lastName             = table[i][COL_C - 1];
    var email                = table[i][COL_D - 1];
    var reservationId        = table[i][COL_I - 1];
    var status               = table[i][COL_K - 1];
    var reservationEmailSent = table[i][COL_L - 1];
    var cancelledSentFlag    = table[i][COL_Q - 1];

    if (!reservationEmailSent) continue;
    if (!reservationId || !email) continue;
    if (!(status === 'Cancelled')) continue;
    if (cancelledSentFlag && String(cancelledSentFlag).length > 0) continue;

    var subj = reservationId + " - Ticket Reservation Cancelled: " + CONFIG.mailSubject;
    var html = tpl(tplCancelled, {
      'plcHeader':        CONFIG.headerLink,
      'plcBanner':        CONFIG.bannerLink,
      'plcReservationId': reservationId,
      'plcFullName':      (firstName || "") + " " + (lastName || "")
    });

    try {
      GmailApp.sendEmail(email, subj, html, {
        htmlBody: html,
        name:     CONFIG.emailSenderName,
        from:     CONFIG.alias
      });
      colQ[i] = "SENT";
    } catch (err) {
      Logger.log("sendCancellationEmail failed for row " + (CONFIG.startRow + i) + ": " + err);
    }
  }

  writeColumnValues(trackingSheet, COL_Q, CONFIG.startRow, colQ);
}