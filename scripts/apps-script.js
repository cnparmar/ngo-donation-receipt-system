function initialSetup() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Create Google Form
  const form = FormApp.create('Donation Form');

  form.addTextItem().setTitle('Full Name');
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addTextItem().setTitle('Phone Number');
  form.addParagraphTextItem().setTitle('Address');
  form.addTextItem().setTitle('Donation Amount');

  const currencyItem = form.addMultipleChoiceItem();
  currencyItem.setTitle('Currency');
  currencyItem.setChoices([
    currencyItem.createChoice('INR'),
    currencyItem.createChoice('USD')
  ]);

  const paymentItem = form.addMultipleChoiceItem();
  paymentItem.setTitle('Payment Mode');
  paymentItem.setChoices([
    paymentItem.createChoice('UPI'),
    paymentItem.createChoice('Bank Transfer'),
    paymentItem.createChoice('Cash')
  ]);

  form.addTextItem().setTitle('Transaction Reference');

  // Link Form to Sheet
  form.setDestination(FormApp.DestinationType.SPREADSHEET, SpreadsheetApp.getActiveSpreadsheet().getId());

  // Store Form Link
  sheet.getRange("D1").setValue("Form Link:");
  sheet.getRange("D2").setValue(form.getPublishedUrl());

  try {
    SpreadsheetApp.getUi().alert("✅ Setup complete! Form created.");
  } catch (e) {
    Logger.log("Setup complete!");
  }
}

function onFormSubmit(e) {

  if (!e) {
    Logger.log("No event object. Run via form submission.");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = e.range.getRow();
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  const timestamp = data[0];
  const name = data[1];
  const email = data[2];
  const address = data[4];
  const amount = parseFloat(data[5]);
  const currency = data[6];
  const paymentMode = data[7];
  const reference = data[8];

  const receiptId = generateReceiptId(sheet);

  sheet.getRange(row, 9).setValue(receiptId);

  const conversion = convertToINR(amount, currency);
  sheet.getRange(row, 11).setValue(conversion.convertedAmount);
  sheet.getRange(row, 12).setValue(conversion.exchangeRate);

  const symbol = getCurrencySymbol(currency);

  const pdfFile = createPDF({
    receiptId,
    date: timestamp,
    name,
    address,
    amount,
    currency,
    currencySymbol: symbol,
    paymentMode,
    reference
  });

  sheet.getRange(row, 10).setValue(pdfFile.getUrl());

  MailApp.sendEmail({
    to: email,
    subject: "Donation Receipt - " + receiptId,
    htmlBody: `
      Dear ${name},<br><br>
      Thank you for your donation.<br><br>
      Please find your receipt attached.<br><br>
      Regards,<br>
      NGO Team
    `,
    attachments: [pdfFile.getBlob()]
  });
}

function generateReceiptId(sheet) {
  const year = sheet.getRange("B1").getValue();
  const counterCell = sheet.getRange("B2");

  let lastNumber = counterCell.getValue();
  if (!lastNumber) lastNumber = 0;

  const newNumber = lastNumber + 1;
  counterCell.setValue(newNumber);

  const formatted = String(newNumber).padStart(4, '0');

  return `RCPT-${year}-${formatted}`;
}

function convertToINR(amount, currency) {
  const rates = {
    "USD": 83,
    "INR": 1
  };

  const rate = rates[currency] || 1;

  return {
    convertedAmount: amount * rate,
    exchangeRate: rate
  };
}

function getCurrencySymbol(currency) {
  if (currency.includes("INR")) return "₹";
  if (currency.includes("USD")) return "$";
  return "";
}

function createPDF(data) {
  const templateId = "15q5eBB5NQ2nzodwD7mc-FTXhBJHJRciAl-Jpt8yfnhw";
  const folderId = "1e15e_pJlwYjdkJIZrAnUxVCXM9rZqUVH";

  const templateFile = DriveApp.getFileById(templateId);
  const copy = templateFile.makeCopy("Receipt-" + data.receiptId);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  body.replaceText("{{receipt_id}}", data.receiptId);
  body.replaceText("{{date}}", new Date(data.date).toLocaleDateString());
  body.replaceText("{{name}}", data.name);
  body.replaceText("{{address}}", data.address);
  body.replaceText("{{amount}}", data.amount);
  body.replaceText("{{currency}}", data.currency);
  body.replaceText("{{currency_symbol}}", data.currencySymbol);
  body.replaceText("{{payment_mode}}", data.paymentMode || "-");
  body.replaceText("{{reference}}", data.reference || "-");

  doc.saveAndClose();

  const pdf = DriveApp.getFileById(copy.getId()).getAs("application/pdf");
  const folder = DriveApp.getFolderById(folderId);
  const pdfFile = folder.createFile(pdf);

  DriveApp.getFileById(copy.getId()).setTrashed(true);

  return pdfFile;
}

