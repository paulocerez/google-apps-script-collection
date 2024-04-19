function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('PDF Generator')
      .addItem('Generate Customer Savings PDF', 'promptForEmail')
      .addToUi();
}

function promptForEmail() {
  let ui = SpreadsheetApp.getUi();
  let result = ui.prompt(
      'Enter Recipient Email',
      'Please enter the email address of the recipient:',
      ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response
  let button = result.getSelectedButton();
  let recipientEmail = result.getResponseText();
  
  if (button === ui.Button.OK && validateEmail(recipientEmail)) {
    // User clicked "OK" and entered a valid email
    generateCustomerSavingsPDF(recipientEmail);
  } else if (button === ui.Button.CANCEL) {
    // User clicked "Cancel"
    ui.alert('PDF generation was canceled.');
  } else {
    // User clicked "OK" but did not enter a valid email
    ui.alert('You did not enter a valid email address.');
  }
}

function validateEmail(email) {
  let emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Simple validation for email pattern
  return emailPattern.test(email);
}

function generateCustomerSavingsPDF(recipientEmail) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Wärmepumpen Rechner');
  let range = sheet.getRange('F21:F33');
  let values = range.getValues();
  
  let doc = DocumentApp.create('Kundenersparnis PDF');
  let body = doc.getBody();
  
  body.appendParagraph('Einsparungen für Kunden\n');
  body.appendParagraph('Erwartete Heizlast: ' + values[0][0] + ' kW');
  body.appendParagraph('WP Preis: ' + values[2][0] + ' €');
  body.appendParagraph('Nach Förderung: ' + values[3][0] + ' €');
  body.appendParagraph('Break-even nach (Jahren): ' + values[5][0]);
  body.appendParagraph('Einsparungen nach 20 Jahren: ' + values[7][0] + ' €');
  body.appendParagraph('CO2-Einsparungen nach 20 Jahren (kg): ' + values[9][0]);
  body.appendParagraph('- Anzahl der benötigten Bäume pro Jahr: ' + values[10][0]);
  body.appendParagraph('- Gefahrene KM mit dem Auto: ' + values[11][0]);
  body.appendParagraph('- Anzahl Flüge: Berlin - Mallorca - Berlin: ' + values[12][0]);
  
  doc.saveAndClose();
  
  let docFile = DriveApp.getFileById(doc.getId());
  let pdfBlob = docFile.getAs('application/pdf').setName(doc.getName() + '.pdf');
  
  // PDF Customer Savings Test Folder
  let folder = DriveApp.getFolderById('1rVg3RpDdMMV3Q3Z2xLgPa__GaHbDX4vP');
  folder.createFile(pdfBlob);
  
  // Send PDF to the recipient's email address
  MailApp.sendEmail(recipientEmail, 'Customer Savings PDF', 'Please find the attached PDF.', {attachments: [pdfBlob]});
  
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}