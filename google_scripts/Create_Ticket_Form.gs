function createReservationForm() {
  var showName = '<SHOW NAME HERE>';

  // --- Create Form ---
  var form = FormApp.create(showName);

  form.setDescription('<ADD YOUR SHOW INFO HERE>\n\nShow passes are priced at PHP XXXX online and PHP XXXX at the door (if still available).\n\nShow passes reserved online shall be redeemed at the venue proper on Month XX, XXXX. Reservations are given until 10pm of the next day to settle payment before getting automatically cancelled. Maximum of four (4) tickets only per transaction.\n\nPlease note that all ticket purchases are final and non-refundable.\n\n***************************************************************\n\nREMINDER: After completing this form, please check your Spam, Updates, Social, and Promotions folders.\n\nYou will be receiving an email regarding the status of your reservation at least a few hours after submitting this form, we need this time to process your reservation. Please DO NOT double book. The payment information and other important details will be included in the email.\n\nPlease check Spam, Updates, Social, and Promotions folders for our emails, since it sometimes lands there. You can also add your email address here to your address book so you don\'t miss any important emails from us.');

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

  form.addMultipleChoiceItem()
    .setTitle('Payment Option')
    .setChoiceValues(['Bank'])
    .setHelpText('This is the only payment channel for this show. You can use the "Bank Transfer" option of GCash or Instapay from other banks.')
    .setRequired(true);

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

  // --- Create folder and move both files into it ---
  var folder    = DriveApp.createFolder(showName);
  var qrFolder  = folder.createFolder('QR');
  var root      = DriveApp.getRootFolder();
  var formFile  = DriveApp.getFileById(form.getId());
  var sheetFile = DriveApp.getFileById(sheet.getId());

  folder.addFile(formFile);
  root.removeFile(formFile);

  folder.addFile(sheetFile);
  root.removeFile(sheetFile);

  Logger.log('Folder: '          + folder.getUrl());
  Logger.log('QR Folder: '       + qrFolder.getUrl());
  Logger.log('QR Folder ID: '    + qrFolder.getId());
  Logger.log('Form URL: '        + form.getPublishedUrl());
  Logger.log('Form edit URL: '   + form.getEditUrl());
  Logger.log('Spreadsheet URL: ' + sheet.getUrl());
}