function onFormSubmit(e) {

  if (!e) {
    Logger.log("No event object. Run via form submission.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const responses = e.namedValues;

  const name = responses["Full Name"][0];
  const email = responses["Email"][0];
  const address = responses["Address"][0];
  const timestamp = responses["Timestamp"][0];
  const amount = responses["Donation Amount"][0];
  const receiptId = generateReceiptId(sheet);

  const receiptCol = getColumnIndexByName(sheet, "Receipt Name");
  sheet.getRange(row, receiptCol).setValue(receiptId);
  
  const pdfFile = createPDF({
    receiptId,
    date: timestamp,
    name,
    address,
    amount
  });

  const pdfCol = getColumnIndexByName(sheet, "Receipt PDF");
  sheet.getRange(row, pdfCol).setValue(pdfFile.getUrl());

  MailApp.sendEmail({
    to: email,
    subject: "Donation Receipt - " + receiptId,
    htmlBody: `
      Dear ${name},<br><br>
      Thank you for your donation of ${amount}.<br><br>
      Please find your receipt attached.<br><br>
      Regards,<br>
      NGO Team
    `,
    attachments: [pdfFile.getBlob()]
  });
}

function generateReceiptId(sheet) {
  const year = new Date().getFullYear();

  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  const counterCell = configSheet.getRange("A1");

  let lastNumber = parseInt(counterCell.getValue(), 10);
  if (isNaN(lastNumber)) lastNumber = 0;

  const newNumber = lastNumber + 1;
  counterCell.setValue(newNumber);

  const formatted = String(newNumber).padStart(4, '0');

  return `RCPT-${year}-${formatted}`;
}

function createPDF(data) {
  const templateId = "15q5eBB5NQ2nzodwD7mc-FTXhBJHJRciAl-Jpt8yfnhw";
  const folderId = "1lquqwuZEnFeaomuapyhJ14Ujz6sSYGLe"

  const folder = DriveApp.getFolderById(folderId);
  const templateFile = DriveApp.getFileById(templateId);

    // Create copy directly in folder
  const copy = templateFile.makeCopy("Receipt-" + data.receiptId, folder);

  // IMPORTANT: slight delay to ensure file availability
  Utilities.sleep(15000);

  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  body.replaceText("{{receipt_id}}", data.receiptId);
  body.replaceText("{{date}}", new Date(data.date).toLocaleDateString());
  body.replaceText("{{name}}", data.name);
  body.replaceText("{{address}}", data.address);
  body.replaceText("{{amount}}", data.amount);

  doc.saveAndClose();

  const pdfBlob = copy.getAs("application/pdf");
  const pdfFile = folder.createFile(pdfBlob).setName(data.receiptId + ".pdf");

  // Cleanup
  copy.setTrashed(true);

  return pdfFile;
}

function getColumnIndexByName(sheet, columnName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === columnName) {
      return i + 1; // 1-based index
    }
  }

  throw new Error("Column not found: " + columnName);
}
