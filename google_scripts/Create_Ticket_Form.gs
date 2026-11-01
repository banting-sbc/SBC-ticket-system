// ============================================================
// EDIT THIS BEFORE EACH RUN
// ============================================================
var showName          = '<SHOW NAME HERE>';
var showPrefix        = '<SHOW PREFIX HERE>';
var ticketPriceOnline = 0;                          // number, e.g. 500
var ticketPriceDoor   = 0;                          // number, e.g. 600
var ticketLimit       = 0;                          // number, e.g. 150
var bankName          = '<BANK NAME HERE>';          // if set, the option reads "Bank Transfer (Bank Name)"
var bankAccountName   = '<BANK ACCOUNT NAME HERE>';
var bankAccountNumber = '<BANK ACCOUNT NUMBER HERE>'; // keep in quotes so leading zeros survive
var gcashNumber       = '<GCASH NUMBER HERE>';      // keep in quotes so the leading 0 survives
var paypalLink        = '<PAYPAL LINK HERE>';
var paypalTransactionFee = 0;                       // number, e.g. 25
var showDate          = 'Month dd, YYYY';           // e.g. 'July 12, 2025'
var facebookUrl       = '<FACEBOOK URL HERE>';
var twitterUrl        = '<TWITTER URL HERE>';
var instagramUrl      = '<INSTAGRAM URL HERE>';
var emailSenderName   = '<EMAIL SENDER NAME HERE>'; // e.g. 'Sleeping Boy Collective Tickets'

// Payment options offered in the form: 'Yes' or 'No' (at least one must be 'Yes')
var enableBankTransfer = 'Yes';
var enableMaya         = 'No';
var enableGCash        = 'No';
var enablePaypal       = 'No';
// ============================================================

var SBC_FOLDER_NAME = 'SBC-ticket-system';

var EMAIL_TEMPLATES_FOLDER_NAME = 'email_html_templates';
var GLOBAL_ASSETS_FOLDER_NAME   = 'global_assets';
var SHOW_ASSETS_FOLDER_NAME     = 'show_assets';
var GITHUB_OWNER  = 'banting-sbc';
var GITHUB_REPO   = 'SBC-ticket-system';
var GITHUB_BRANCH = 'main';

// Images (base names, .png or .jpg) that must exist inside global_assets
var GLOBAL_ASSET_IMAGE_NAMES = [
  'bank-info', 'bank-qr', 'banner', 'facebook', 'gcash-info',
  'gcash-qr', 'header', 'instagram', 'maya-info', 'maya-qr', 'twitter'
];
var IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg'];

// Template files that must exist inside the email_html_templates folder
var REQUIRED_EMAIL_TEMPLATES = [
  'Reminder.html',
  'ReservationCancelled.html',
  'ReservationReceived_BANK.html',
  'ReservationReceived_GCash.html',
  'ReservationReceived_Paypal.html',
  'TicketConfirmed.html',
  'TicketPDF.html',
  'Waitlisted.html'
];

function createReservationForm() {
  // --- Step 0: validate the payment option inputs before anything is created ---
  var paymentChoices = getEnabledPaymentOptions();

  // --- Step 1: SBC-ticket-system folder must exist, otherwise exit ---
  var sbcFolder = checkSbcTicketSystemFolder();
  if (!sbcFolder) {
    throw new Error('Folder "' + SBC_FOLDER_NAME + '" was not found in Google Drive. Please create it first, then run this script again. No files were created.');
  }

  // --- Step 2: email_html_templates folder (created if missing) ---
  var templatesFolder = ensureEmailTemplatesFolder(sbcFolder);

  // --- Step 3: required HTML templates (any missing ones are downloaded from GitHub) ---
  ensureEmailTemplateFiles(templatesFolder);

  // --- Step 4: global_assets folder inside SBC-ticket-system (created if missing) ---
  var globalAssetsFolder = ensureGlobalAssetsFolder(sbcFolder);

  // --- Step 5: required global images (any missing ones are downloaded from GitHub) ---
  ensureGlobalAssetFiles(globalAssetsFolder);

  // --- Step 6: form, sheet and subfolder creation (below) ---

  // --- Stop if a folder for this show already exists (before creating anything) ---
  var existingShowFolder = checkShowFolderExists(sbcFolder, showName);
  if (existingShowFolder) {
    throw new Error('A folder named "' + showName + '" already exists inside "' + SBC_FOLDER_NAME + '": ' + existingShowFolder.getUrl() + '. Change showName or rename/remove the existing folder, then run again. No files were created.');
  }

  // --- Create Form ---
  var form = FormApp.create(showName);

  form.setDescription('Show passes are priced at PHP ' + ticketPriceOnline + ' online and PHP ' + ticketPriceDoor + ' at the door (if still available).\n\nShow passes reserved online shall be redeemed at the venue proper on ' + showDate + '. Reservations are given until 10pm of the next day to settle payment before getting automatically cancelled. Maximum of four (4) tickets only per transaction.\n\nPlease note that all ticket purchases are final and non-refundable.\n\n***************************************************************\n\nREMINDER: After completing this form, please check your Spam, Updates, Social, and Promotions folders.\n\nYou will be receiving an email regarding the status of your reservation at least a few hours after submitting this form, we need this time to process your reservation. Please DO NOT double book. The payment information and other important details will be included in the email.\n\nPlease check Spam, Updates, Social, and Promotions folders for our emails, since it sometimes lands there. You can also add your email address here to your address book so you don\'t miss any important emails from us.');

  form.addTextItem().setTitle('First Name').setRequired(true);
  form.addTextItem().setTitle('Last Name').setRequired(true);
  
  var emailValidation = FormApp.createTextValidation()
    .setHelpText('Please enter a valid email address (e.g. name@example.com).')
    .requireTextIsEmail()
    .build();

  form.addTextItem()
    .setTitle('Email Address')
    .setRequired(true)
    .setValidation(emailValidation);

  var mobileValidation = FormApp.createTextValidation()
    .setHelpText('Please enter a valid 11-digit mobile number starting with 0 (e.g. 09123456789).')
    .requireTextMatchesPattern('^0\\d{10}$')
    .build();

  form.addTextItem()
    .setTitle('Mobile Number')
    .setRequired(true)
    .setValidation(mobileValidation);
  form.addParagraphTextItem().setTitle('City').setRequired(true);

  form.addListItem()
    .setTitle('Number of Tickets')
    .setChoiceValues(['1', '2', '3', '4'])
    .setRequired(true);

  var paymentItem = form.addMultipleChoiceItem()
    .setTitle('Payment Option')
    .setChoiceValues(paymentChoices)
    .setRequired(true);

  // Help text only applies when Bank Transfer is one of the options
  if (paymentChoices.some(function (c) { return c.indexOf('Bank Transfer') === 0; })) {
    paymentItem.setHelpText('For bank transfers, you can use the "Bank Transfer" option of GCash or Instapay from other banks.');
  }

  form.addCheckboxItem()
    .setTitle('OPTIONAL: Would you like to get updates on upcoming shows and new releases via email? (Leave blank to decline)')
    .setChoiceValues(['Yes']);

  form.addCheckboxItem()
    .setTitle('I have read this form, understood its contents and consent to the processing of my personal data. I understand that my consent does not preclude the existence of other criteria for lawful processing of personal data, and does not waive any of my rights under the Data Privacy Act of 2012 and other applicable laws.\n\nPLEASE CLICK SUBMIT TO SEND YOUR RESERVATION')
    .setChoiceValues(['I agree.'])
    .setRequired(true);

  // --- Create Spreadsheet named after show ---
  var sheet = SpreadsheetApp.create(showName);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());

  // Delete default 'Sheet1' tab — form linking creates 'Form Responses 1' automatically
  var defaultSheet = sheet.getSheetByName('Sheet1');
  if (defaultSheet) sheet.deleteSheet(defaultSheet);

  // --- Create show folder inside SBC-ticket-system and move both files into it ---
  var folder    = sbcFolder.createFolder(showName);
  var qrFolder  = folder.createFolder('QR');
  var showAssetsFolder = folder.createFolder(SHOW_ASSETS_FOLDER_NAME);
  var root     = DriveApp.getRootFolder();
  var formFile  = DriveApp.getFileById(form.getId());
  var sheetFile = DriveApp.getFileById(sheet.getId());

  folder.addFile(formFile);
  root.removeFile(formFile);

  folder.addFile(sheetFile);
  root.removeFile(sheetFile);

  // --- Add a "References" tab to the spreadsheet with all the links and IDs ---
  var templateIds = getEmailTemplateIds(templatesFolder);

  var referenceRows = [
    ['SBC Folder URL',       sbcFolder.getUrl()],
    ['Templates Folder URL', templatesFolder.getUrl()],
    ['Global Assets URL',    globalAssetsFolder.getUrl()],
    ['Show Folder URL',      folder.getUrl()],
    ['Show Assets URL',      showAssetsFolder.getUrl()],
    ['QR Folder URL',        qrFolder.getUrl()],
    ['QR Folder ID',         qrFolder.getId()],
    ['Form URL',             form.getPublishedUrl()],
    ['Form Edit URL',        form.getEditUrl()],
    ['Spreadsheet URL',      sheet.getUrl()],
    ['Show Name',            showName],
    ['Show Prefix',          showPrefix],
    ['Show Date',            showDate],
    ['Ticket Price',         ticketPriceOnline],
    ['Ticket Price Door',    ticketPriceDoor],
    ['Ticket Limit',         ticketLimit],
    ['Bank Name',            bankName],
    ['Bank Account Name',    bankAccountName],
    ['Bank Account Number',  bankAccountNumber],
    ['GCash Number',         gcashNumber],
    ['Paypal Link',          paypalLink],
    ['Paypal Transaction Fee', paypalTransactionFee],
    ['Facebook URL',         facebookUrl],
    ['Twitter URL',          twitterUrl],
    ['Instagram URL',        instagramUrl],
    ['Email Sender Name',    emailSenderName],
    ['Payment Options',      paymentChoices.join(', ')]
  ];
  REQUIRED_EMAIL_TEMPLATES.forEach(function (fileName) {
    referenceRows.push(['Template ID: ' + fileName, templateIds[fileName]]);
  });

  writeReferencesSheet(sheet, referenceRows);

  Logger.log('SBC Folder: '      + sbcFolder.getUrl());
  Logger.log('Templates Folder: ' + templatesFolder.getUrl());
  Logger.log('Global Assets: '   + globalAssetsFolder.getUrl());
  Logger.log('Folder: '          + folder.getUrl());
  Logger.log('Show Assets: '     + showAssetsFolder.getUrl());
  Logger.log('QR Folder: '       + qrFolder.getUrl());
  Logger.log('QR Folder ID: '    + qrFolder.getId());
  Logger.log('Form URL: '        + form.getPublishedUrl());
  Logger.log('Form edit URL: '   + form.getEditUrl());
  Logger.log('Spreadsheet URL: ' + sheet.getUrl());
}

/**
 * Checks whether a folder named "SBC-ticket-system" exists in Google Drive.
 * Trashed folders are ignored by DriveApp.getFoldersByName().
 *
 * @return {GoogleAppsScript.Drive.Folder|null} The folder if found, otherwise null.
 */
function checkSbcTicketSystemFolder() {
  var folders = DriveApp.getFoldersByName(SBC_FOLDER_NAME);

  if (!folders.hasNext()) {
    Logger.log('Folder "' + SBC_FOLDER_NAME + '" was NOT found in Google Drive.');
    return null;
  }

  var found = folders.next();

  if (folders.hasNext()) {
    Logger.log('WARNING: More than one folder named "' + SBC_FOLDER_NAME + '" exists. Using: ' + found.getUrl());
  }

  Logger.log('Folder "' + SBC_FOLDER_NAME + '" found: ' + found.getUrl());
  return found;
}

/**
 * Step 2 check: looks for the "email_html_templates" folder directly inside the given parent folder.
 *
 * @param {GoogleAppsScript.Drive.Folder} parentFolder
 * @return {GoogleAppsScript.Drive.Folder|null} The folder if found, otherwise null.
 */
function checkEmailTemplatesFolder(parentFolder) {
  var folders = parentFolder.getFoldersByName(EMAIL_TEMPLATES_FOLDER_NAME);

  if (!folders.hasNext()) {
    Logger.log('Folder "' + EMAIL_TEMPLATES_FOLDER_NAME + '" was NOT found inside "' + parentFolder.getName() + '".');
    return null;
  }

  var found = folders.next();
  Logger.log('Folder "' + EMAIL_TEMPLATES_FOLDER_NAME + '" found: ' + found.getUrl());
  return found;
}

/**
 * Step 3 check: lists which files from REQUIRED_EMAIL_TEMPLATES are missing from the folder.
 *
 * @param {GoogleAppsScript.Drive.Folder} templatesFolder
 * @return {string[]} Names of missing files (empty array if all are present).
 */
function checkEmailTemplateFiles(templatesFolder) {
  var missingFiles = REQUIRED_EMAIL_TEMPLATES.filter(function (fileName) {
    return !templatesFolder.getFilesByName(fileName).hasNext();
  });

  if (missingFiles.length > 0) {
    Logger.log('Missing template file(s): ' + missingFiles.join(', '));
  } else {
    Logger.log('All ' + REQUIRED_EMAIL_TEMPLATES.length + ' required template files are present.');
  }

  return missingFiles;
}

/**
 * Downloads a single file from the GitHub email_html_templates folder into a Drive folder.
 *
 * @param {string} fileName e.g. "Reminder.html"
 * @param {GoogleAppsScript.Drive.Folder} targetFolder
 */
function downloadGithubFileToDrive(fileName, targetFolder) {
  var rawUrl = 'https://raw.githubusercontent.com/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/' +
               GITHUB_BRANCH + '/' + EMAIL_TEMPLATES_FOLDER_NAME + '/' + encodeURIComponent(fileName);

  saveUrlToDrive_(rawUrl, fileName, targetFolder);
  Logger.log('Downloaded missing file: ' + fileName);
}

/**
 * Fetches a URL and saves the response as a file in the given Drive folder.
 */
function saveUrlToDrive_(url, fileName, targetFolder) {
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error('Failed to download "' + fileName + '" (HTTP ' + response.getResponseCode() + ').');
  }

  var blob = response.getBlob().setName(fileName);
  if (/\.html?$/i.test(fileName))       blob.setContentType('text/html');
  else if (/\.png$/i.test(fileName))    blob.setContentType('image/png');
  else if (/\.jpe?g$/i.test(fileName))  blob.setContentType('image/jpeg');

  return targetFolder.createFile(blob);
}

/**
 * Step 2: returns the "email_html_templates" folder inside SBC-ticket-system,
 * creating it first if it does not exist.
 *
 * @param {GoogleAppsScript.Drive.Folder} sbcFolder
 * @return {GoogleAppsScript.Drive.Folder}
 */
function ensureEmailTemplatesFolder(sbcFolder) {
  var folder = checkEmailTemplatesFolder(sbcFolder);
  if (folder) return folder;

  folder = sbcFolder.createFolder(EMAIL_TEMPLATES_FOLDER_NAME);
  Logger.log('Created folder "' + EMAIL_TEMPLATES_FOLDER_NAME + '": ' + folder.getUrl());
  return folder;
}

/**
 * Step 3: makes sure every file in REQUIRED_EMAIL_TEMPLATES exists in the templates folder,
 * downloading any missing ones from GitHub. Existing files are never overwritten.
 * Tries every missing file, then throws one error listing any that failed.
 *
 * @param {GoogleAppsScript.Drive.Folder} templatesFolder
 */
function ensureEmailTemplateFiles(templatesFolder) {
  var missingFiles = checkEmailTemplateFiles(templatesFolder);
  if (missingFiles.length === 0) return;

  var failures = [];
  missingFiles.forEach(function (fileName) {
    try {
      downloadGithubFileToDrive(fileName, templatesFolder);
    } catch (e) {
      failures.push(fileName + ' (' + e.message + ')');
    }
  });

  if (failures.length > 0) {
    throw new Error('Could not download missing template file(s): ' + failures.join('; '));
  }
}

/**
 * Returns the "global_assets" folder inside SBC-ticket-system,
 * creating it first if it does not exist.
 *
 * @param {GoogleAppsScript.Drive.Folder} sbcFolder
 * @return {GoogleAppsScript.Drive.Folder}
 */
function ensureGlobalAssetsFolder(sbcFolder) {
  var folders = sbcFolder.getFoldersByName(GLOBAL_ASSETS_FOLDER_NAME);

  if (folders.hasNext()) {
    var existing = folders.next();
    Logger.log('Folder "' + GLOBAL_ASSETS_FOLDER_NAME + '" found: ' + existing.getUrl());
    return existing;
  }

  var created = sbcFolder.createFolder(GLOBAL_ASSETS_FOLDER_NAME);
  Logger.log('Created folder "' + GLOBAL_ASSETS_FOLDER_NAME + '": ' + created.getUrl());
  return created;
}

/**
 * Reads the Yes/No payment option inputs at the top of the script and returns
 * the labels of the enabled options, in this order: Bank Transfer, Maya,
 * GCash, PayPal. If bankName is filled in, the first label becomes
 * "Bank Transfer (Bank Name)". Input is trimmed and not case-sensitive. Throws if a value
 * is not Yes/No or if no option is enabled.
 *
 * @return {string[]}
 */
function getEnabledPaymentOptions() {
  var inputs = [
    { label: 'Bank Transfer', variable: 'enableBankTransfer', value: enableBankTransfer },
    { label: 'Maya',          variable: 'enableMaya',         value: enableMaya },
    { label: 'GCash',         variable: 'enableGCash',        value: enableGCash },
    { label: 'PayPal',        variable: 'enablePaypal',       value: enablePaypal }
  ];

  var enabled = [];
  inputs.forEach(function (input) {
    var v = String(input.value).trim().toLowerCase();
    if (v === 'yes') {
      var label = input.label;
      if (input.variable === 'enableBankTransfer') {
        var bank = String(bankName).trim();
        if (bank !== '' && !/^<.*>$/.test(bank)) label = 'Bank Transfer (' + bank + ')';
      }
      enabled.push(label);
    } else if (v !== 'no') {
      throw new Error('Invalid value "' + input.value + '" for ' + input.variable + ' - use \'Yes\' or \'No\'. No files were created.');
    }
  });

  if (enabled.length === 0) {
    throw new Error('No payment option is enabled. Set at least one of enableBankTransfer, enableMaya, enableGCash or enablePaypal to \'Yes\'. No files were created.');
  }
  return enabled;
}

/**
 * Returns { lowercased file name: true } for the files directly inside a folder.
 */
function indexFileNames_(folder) {
  var names = {};
  var files = folder.getFiles();
  while (files.hasNext()) names[files.next().getName().toLowerCase()] = true;
  return names;
}

/**
 * Returns true if baseName exists as .png, .jpg or .jpeg (case-insensitive)
 * in a name index built by indexFileNames_().
 */
function hasImage_(nameIndex, baseName) {
  return IMAGE_EXTENSIONS.some(function (ext) {
    return nameIndex.hasOwnProperty((baseName + ext).toLowerCase());
  });
}

/**
 * Checks which of GLOBAL_ASSET_IMAGE_NAMES are missing from the global_assets folder.
 * An image counts as present as a .png, .jpg or .jpeg file.
 *
 * @param {GoogleAppsScript.Drive.Folder} globalAssetsFolder
 * @return {string[]} Base names of the missing images (empty if all are present).
 */
function checkGlobalAssetFiles(globalAssetsFolder) {
  var index = indexFileNames_(globalAssetsFolder);
  var missing = GLOBAL_ASSET_IMAGE_NAMES.filter(function (baseName) {
    return !hasImage_(index, baseName);
  });

  if (missing.length > 0) {
    Logger.log('Missing global asset image(s): ' + missing.join(', '));
  } else {
    Logger.log('All ' + GLOBAL_ASSET_IMAGE_NAMES.length + ' required global asset images are present.');
  }
  return missing;
}

/**
 * Makes sure every image in GLOBAL_ASSET_IMAGE_NAMES exists in global_assets.
 * Missing ones are downloaded from the GitHub global_assets folder (the repo
 * file may be .png, .jpg or .jpeg; its own name is kept). Existing files are
 * never overwritten. Throws one error listing anything that could not be
 * found in the repo or downloaded.
 *
 * @param {GoogleAppsScript.Drive.Folder} globalAssetsFolder
 */
function ensureGlobalAssetFiles(globalAssetsFolder) {
  var missing = checkGlobalAssetFiles(globalAssetsFolder);
  if (missing.length === 0) return;

  // One call lists the GitHub folder
  var apiUrl = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO +
               '/contents/' + GLOBAL_ASSETS_FOLDER_NAME + '?ref=' + GITHUB_BRANCH;
  var response = UrlFetchApp.fetch(apiUrl, {
    headers: { 'Accept': 'application/vnd.github+json' },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('GitHub API request failed for "' + GLOBAL_ASSETS_FOLDER_NAME + '" (HTTP ' +
                    response.getResponseCode() + '): ' + response.getContentText());
  }

  var repoFiles = {};   // lowercased name -> { name, download_url }
  JSON.parse(response.getContentText()).forEach(function (item) {
    if (item.type === 'file') repoFiles[item.name.toLowerCase()] = item;
  });

  var problems = [];
  missing.forEach(function (baseName) {
    var match = null;
    for (var i = 0; i < IMAGE_EXTENSIONS.length && !match; i++) {
      match = repoFiles[(baseName + IMAGE_EXTENSIONS[i]).toLowerCase()] || null;
    }
    if (!match) {
      problems.push(baseName + ' (not found in the GitHub ' + GLOBAL_ASSETS_FOLDER_NAME + ' folder)');
      return;
    }
    try {
      saveUrlToDrive_(match.download_url, match.name, globalAssetsFolder);
      Logger.log('Downloaded missing global asset: ' + match.name);
    } catch (e) {
      problems.push(baseName + ' (' + e.message + ')');
    }
  });

  if (problems.length > 0) {
    throw new Error('Could not get global asset image(s): ' + problems.join('; '));
  }
}

/**
 * Looks up the Drive file ID of every required template in the templates folder.
 *
 * @param {GoogleAppsScript.Drive.Folder} templatesFolder
 * @return {Object<string, string>} Map of file name -> file ID ("NOT FOUND" if absent)
 */
function getEmailTemplateIds(templatesFolder) {
  var ids = {};
  REQUIRED_EMAIL_TEMPLATES.forEach(function (fileName) {
    var files = templatesFolder.getFilesByName(fileName);
    ids[fileName] = files.hasNext() ? files.next().getId() : 'NOT FOUND';
  });
  return ids;
}

/**
 * Creates (or replaces) a "References" tab in the spreadsheet and fills it with
 * an Item / Value table.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {Array<Array<string>>} rows [item, value] pairs
 */
function writeReferencesSheet(spreadsheet, rows) {
  var existing = spreadsheet.getSheetByName('References');
  if (existing) spreadsheet.deleteSheet(existing);

  var refSheet = spreadsheet.insertSheet('References');
  var data = [['Item', 'Value']].concat(rows);

  // Plain text format first, so values like "July 12, 2025" or an account
  // number with leading zeros are never converted to a date / number.
  refSheet.getRange(1, 1, data.length, 2).setNumberFormat('@');
  refSheet.getRange(1, 1, data.length, 2).setValues(data);
  refSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  refSheet.setFrozenRows(1);
  refSheet.autoResizeColumn(1);
  refSheet.setColumnWidth(2, 600);
  SpreadsheetApp.flush();
}

/**
 * Checks whether a folder named after the show already exists inside SBC-ticket-system.
 *
 * @param {GoogleAppsScript.Drive.Folder} sbcFolder
 * @param {string} name The show name
 * @return {GoogleAppsScript.Drive.Folder|null} The existing folder, or null if none.
 */
function checkShowFolderExists(sbcFolder, name) {
  var folders = sbcFolder.getFoldersByName(name);

  if (!folders.hasNext()) {
    Logger.log('No existing folder named "' + name + '" inside "' + sbcFolder.getName() + '".');
    return null;
  }

  var existing = folders.next();
  Logger.log('A folder named "' + name + '" already exists: ' + existing.getUrl());
  return existing;
}