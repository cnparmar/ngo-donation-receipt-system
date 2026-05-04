# Setup Guide (Step-by-Step)

## Step 1: Create Google Form
Add fields:
- Name
- Email
- Address
- Amount
- Currency (INR/USD)
- Payment Mode
- Reference

---

## Step 2: Link to Google Sheet
Click "Responses → Link to Sheets"

---

## Step 3: Prepare Sheet

Add columns:
- Receipt ID
- PDF Link
- Converted Amount (INR)
- Exchange Rate

Top cells:
A1: Year → 2026  
A2: Last Receipt Number → 0  

---

## Step 4: Create Receipt Template

Create Google Doc and paste:

Receipt No: {{receipt_id}}  
Date: {{date}}  

Name: {{name}}  
Address: {{address}}  

Amount: {{currency_symbol}}{{amount}}  
Currency: {{currency}}  

---

## Step 5: Copy IDs

- Copy Google Doc ID
- Copy Google Drive folder ID

---

## Step 6: Add Script

Go to:
Extensions → Apps Script

Paste code from `/scripts/apps-script.js`

Replace:
- TEMPLATE_ID
- FOLDER_ID

---

## Step 7: Add Trigger

Function: onFormSubmit  
Event: On form submit  

---

## Step 8: Test

Submit test entry and verify:
- Email received
- PDF generated
- Data saved

---

Done 🎉
