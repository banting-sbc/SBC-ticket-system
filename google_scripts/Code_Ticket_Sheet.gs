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
// SECTION 1b — TEMPLATE CONSTANTS
// Drive file IDs for all HTML email and PDF templates.
// These are written to Counters E36:E43 by setupSheets() and
// loaded into CONFIG at runtime by loadConfig().
// Update these values here if a template file is ever replaced.
// =============================================================

var TEMPLATE_IDS = {
  bank:      "129T-hTaGn6DZn4xtTLS2U17HR89qcnQd", // Template: Reservation Bank
  gcash:     "1M4UJO97-pnE5vOMo5420IE9Ii8BhtLuF", // Template: Reservation GCash
  paypal:    "1oSaoUr4aBGmdsvbeQy2atwIGbZuU9Nr4", // Template: Reservation PayPal
  waitlist:  "1wF-qcFD8KfXAm4oVDqh3RtDJ6wM3oZze", // Template: Waitlisted
  reminder:  "18PMft48dE_qx2nzp9Jht0V7x_OraTa-z", // Template: Reminder
  confirmed: "1ZJLtDSrtYYF1fH-3ZgGUsymq6XxgLKzj", // Template: Ticket Confirmed
  cancelled: "1ahY23WWsEaicvSH7dpiXfmnd1l_KSAZw", // Template: Reservation Cancelled
  ticketPDF: "1n8tKodqEsc08mluFMzWEfu-dY3r2oCpE"  // Template: Ticket PDF
};


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
      ["TICKET SLOTS REMAINING",                  "=E7-B9"],
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

    // Col D: labels, Col E: default values.
    // Row numbers correspond exactly to loadConfig() offsets.
    var deData = [
      ["Header Link",                       ""],            // D1/E1
      ["Show Name",                         ""],            // D2/E2
      ["Banner Link",                       ""],            // D3/E3
      ["Show Prefix",                       ""],            // D4/E4
      [null,                                null],          // D5/E5 — skipped
      ["Ticket Price",                      0],             // D6/E6
      ["Ticket Limit",                      0],             // D7/E7
      [null,                                null],          // D8/E8 — skipped
      ["Bank Name",                         ""],            // D9/E9
      ["Bank Account Name",                 ""],            // D10/E10
      ["Bank Account Number",               ""],            // D11/E11
      ["Bank QR Link",                      ""],            // D12/E12
      ["Bank Infographic Link",             ""],            // D13/E13
      [null,                                null],          // D14/E14 — skipped
      ["GCash Number",                      ""],            // D15/E15
      ["GCash QR Link",                     ""],            // D16/E16
      ["GCash Infographic Link",            ""],            // D17/E17
      [null,                                null],          // D18/E18 — skipped
      ["Paypal Link",                       ""],            // D19/E19
      ["Paypal Transaction Fee",            0],             // D20/E20
      [null,                                null],          // D21/E21 — skipped
      ["Days before reminder email",        1],             // D22/E22
      ["Show Date",                         ""],            // D23/E23
      ["Reservation Waiting Period Text",   "by 10 PM tomorrow"], // D24/E24
      ["Reservation Waiting Period (Days)", 1],             // D25/E25
      [null,                                null],          // D26/E26 — skipped
      ["Start Processing From Row",         2],             // D27/E27
      [null,                                null],          // D28/E28 — skipped
      ["Early Bird Enabled",                "No"],          // D29/E29 — dropdown
      ["Early Bird Limit",                  0],             // D30/E30
      ["Early Bird Ticket Price",           0],             // D31/E31
      [null,                                null],          // D32/E32 — skipped
      ["Web App URL",                       ""],            // D33/E33
      ["QR Folder ID",                      ""],            // D34/E34
      [null,                                null],          // D35/E35 — skipped
      ["Template: Reservation Bank",        TEMPLATE_IDS.bank],      // D36/E36
      ["Template: Reservation GCash",       TEMPLATE_IDS.gcash],     // D37/E37
      ["Template: Reservation PayPal",      TEMPLATE_IDS.paypal],    // D38/E38
      ["Template: Waitlisted",              TEMPLATE_IDS.waitlist],  // D39/E39
      ["Template: Reminder",                TEMPLATE_IDS.reminder],  // D40/E40
      ["Template: Ticket Confirmed",        TEMPLATE_IDS.confirmed], // D41/E41
      ["Template: Reservation Cancelled",   TEMPLATE_IDS.cancelled], // D42/E42
      ["Template: Ticket PDF",              TEMPLATE_IDS.ticketPDF]  // D43/E43
    ];

    for (var s = 0; s < deData.length; s++) {
      var rowNum2 = s + 1;
      if (deData[s][0]) counters.getRange(rowNum2, 4).setValue(deData[s][0]);
      if (deData[s][1] !== null) counters.getRange(rowNum2, 5).setValue(deData[s][1]);
    }

    // Early Bird Enabled dropdown on E29
    var ebRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Yes", "No", "Exceeded"])
      .build();
    counters.getRange("E29").setDataValidation(ebRule);

    // Format Show Date (E23) as plain text so entries like
    // "July 12, 2025" are never auto-converted to a date value.
    counters.getRange("E23").setNumberFormat("@");

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

    // ── Column D/E — Config sections ─────────────────────────
    formatSection([
      { row: 1 }, { row: 2 }, { row: 3 }, { row: 4 }
    ], 4, 5);
    outerBorder(counters.getRange("D1:E4"));

    formatSection([{ row: 6 }, { row: 7 }], 4, 5);
    outerBorder(counters.getRange("D6:E7"));

    formatSection([
      { row: 9 }, { row: 10 }, { row: 11 }, { row: 12 }, { row: 13 }
    ], 4, 5);
    outerBorder(counters.getRange("D9:E13"));

    formatSection([{ row: 15 }, { row: 16 }, { row: 17 }], 4, 5);
    outerBorder(counters.getRange("D15:E17"));

    formatSection([{ row: 19 }, { row: 20 }], 4, 5);
    outerBorder(counters.getRange("D19:E20"));

    formatSection([
      { row: 22 }, { row: 23 }, { row: 24 }, { row: 25 }
    ], 4, 5);
    outerBorder(counters.getRange("D22:E25"));

    formatSection([{ row: 27 }], 4, 5);
    outerBorder(counters.getRange("D27:E27"));

    formatSection([{ row: 29 }, { row: 30 }, { row: 31 }], 4, 5);
    outerBorder(counters.getRange("D29:E31"));

    formatSection([{ row: 33 }, { row: 34 }], 4, 5);
    outerBorder(counters.getRange("D33:E34"));

    formatSection([
      { row: 36 }, { row: 37 }, { row: 38 }, { row: 39 },
      { row: 40 }, { row: 41 }, { row: 42 }, { row: 43 }
    ], 4, 5);
    outerBorder(counters.getRange("D36:E43"));

    // Grey out skipped rows in D/E
    [5, 8, 14, 18, 21, 26, 28, 32, 35].forEach(function(r) {
      counters.getRange(r, 4, 1, 2).setBackground(COLOR_SKIP);
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
// Reads all configuration from Counters!E1:E43 in one batch.
// Called explicitly at the top of every core function.
// =============================================================

function loadConfig() {
  getSheetVariables();
  // Single range read — configValues[n][0] == cell E(n+1)
  var configValues = summarySheet.getRange('E1:E43').getValues();

  CONFIG.headerLink           = configValues[0][0];   // E1
  CONFIG.mailSubject          = configValues[1][0];   // E2
  CONFIG.bannerLink           = configValues[2][0];   // E3
  CONFIG.preFix               = configValues[3][0];   // E4
  // E5 — skipped
  CONFIG.ticketPrice          = Number(configValues[5][0]);   // E6
  CONFIG.waitlistThreshold    = Number(configValues[6][0]);   // E7
  // E8 — skipped
  CONFIG.bankName             = configValues[8][0];   // E9
  CONFIG.bankAccountName      = configValues[9][0];   // E10
  CONFIG.accountNumber        = String(configValues[10][0]);  // E11
  CONFIG.bankQR               = configValues[11][0];  // E12
  CONFIG.bankInfographic      = configValues[12][0];  // E13
  // E14 — skipped
  CONFIG.gCashNumber          = String(configValues[14][0]);  // E15
  CONFIG.gcashQR              = configValues[15][0];  // E16
  CONFIG.gcashInfographic     = configValues[16][0];  // E17
  // E18 — skipped
  CONFIG.payPalLink           = configValues[18][0];  // E19
  CONFIG.paypalFee            = Number(configValues[19][0]);  // E20
  // E21 — skipped
  CONFIG.reminderThreshold    = Number(configValues[21][0]);  // E22
  CONFIG.showDate             = String(configValues[22][0]);  // E23
  CONFIG.reservationPeriod    = String(configValues[23][0]);  // E24
  CONFIG.resAllowance         = Number(configValues[24][0]);  // E25
  // E26 — skipped
  CONFIG.startRow             = Number(configValues[26][0]);  // E27
  // E28 — skipped
  CONFIG.earlyBirdEnabledRef  = 'E29'; // cell address for read/write-back
  CONFIG.earlyBirdThreshold   = Number(configValues[29][0]);  // E30
  CONFIG.earlyBirdTicketPrice = Number(configValues[30][0]);  // E31
  // E32 — skipped
  CONFIG.webAppURL            = configValues[32][0];  // E33
  CONFIG.qrFolderID           = configValues[33][0];  // E34
  // E35 — skipped
  CONFIG.tplBank              = configValues[35][0];  // E36
  CONFIG.tplGCash             = configValues[36][0];  // E37
  CONFIG.tplPayPal            = configValues[37][0];  // E38
  CONFIG.tplWaitlisted        = configValues[38][0];  // E39
  CONFIG.tplReminder          = configValues[39][0];  // E40
  CONFIG.tplConfirmed         = configValues[40][0];  // E41
  CONFIG.tplCancelled         = configValues[41][0];  // E42
  CONFIG.tplTicketPDF         = configValues[42][0];  // E43

  // PH mobile numbers are 11 digits; prepend 0 if stored without it
  if (CONFIG.gCashNumber && CONFIG.gCashNumber.length === 10) {
    CONFIG.gCashNumber = "0" + CONFIG.gCashNumber;
  }

  // Derived
  CONFIG.qrFolder = DriveApp.getFolderById(CONFIG.qrFolderID);

  // Regex patterns
  CONFIG.regexBank   = /(bank|maya)/i;
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

function tpl(html, map) {
  for (var k in map) {
    if (map.hasOwnProperty(k)) {
      html = html.replace(new RegExp("\\{" + k + "\\}", "g"), map[k]);
    }
  }
  return html;
}

// Fetches HTML template content from a Drive file by ID.
// File IDs are stored in Counters E36:E43 and loaded into
// CONFIG by loadConfig(). Throws clearly if the ID is missing
// so misconfiguration is easy to diagnose in the logs.
function readTemplate(fileId) {
  if (!fileId) {
    throw new Error("readTemplate: file ID is empty — check Counters E36:E43.");
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
      html = tpl(tplBank, {
        "plcBankName":        CONFIG.bankName,
        "plcBankAccountName": CONFIG.bankAccountName,
        "plcAccountNumber":   CONFIG.accountNumber,
        "plcInfographic":     CONFIG.bankInfographic
      });
      try { qrCodeBlob = UrlFetchApp.fetch(CONFIG.bankQR).getBlob().setName("bank-QR.jpg"); }
      catch (err) { Logger.log("Failed to fetch bankQR: " + err); }
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
      var mailOptions = { htmlBody: html, name: "Sleeping Boy Collective Tickets", from: CONFIG.alias };
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
            name:        "Sleeping Boy Collective Tickets",
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
          name:     "Sleeping Boy Collective Tickets",
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
        name:     "Sleeping Boy Collective Tickets",
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
        name:     "Sleeping Boy Collective Tickets",
        from:     CONFIG.alias
      });
      colQ[i] = "SENT";
    } catch (err) {
      Logger.log("sendCancellationEmail failed for row " + (CONFIG.startRow + i) + ": " + err);
    }
  }

  writeColumnValues(trackingSheet, COL_Q, CONFIG.startRow, colQ);
}