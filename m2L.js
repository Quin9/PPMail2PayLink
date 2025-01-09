// Include xlsx library for reading XLSX files
import * as XLSX from 'xlsx';
import fs from 'fs';
import path from 'path';

// Base URL templates
const urlTemplates = {
  xclick: "https://www.paypal.com/cgi-bin/webscr?business={email}&cmd=_xclick&currency_code=&amount=0.02&item_name=&return=&cancel_return=",
  donations: "https://www.paypal.com/cgi-bin/webscr?business={email}&cmd=_donations&currency_code=&amount=0.02&item_name=&return=&cancel_return="
};

// Function to generate PayPal URLs
function generatePayPalUrls(emailList, mode) {
  const baseUrl = urlTemplates[mode] || urlTemplates.xclick; // Default to xclick if mode is invalid
  return emailList.map(email => ({ email, url: baseUrl.replace("{email}", email) }));
}

// Function to read emails from an XLSX file
function readEmailsFromFile(filePath) {
  const fileBuffer = fs.readFileSync(filePath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0]; // Assuming emails are in the first sheet
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Read sheet as 2D array

  // Assuming emails are in the first column
  return data.map(row => row[0]).filter(email => typeof email === 'string');
}

// Function to save generated URLs to a new Excel file
function saveUrlsToFile(emailUrlPairs) {
  const sheetData = emailUrlPairs.map(pair => [pair.email, pair.url]); // Include both email and URL
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(sheetData);
  XLSX.utils.book_append_sheet(workbook, sheet, "PayPal URLs");

  const randomNum = Math.floor(100 + Math.random() * 900); // Generate random 3-digit number
  const outputDir = process.pkg ? path.dirname(process.execPath) : process.cwd(); // Adjust output path for EXE
  const fileName = path.join(outputDir, `Newppmail2paylink${randomNum}.xlsx`);

  XLSX.writeFile(workbook, fileName);
  console.log(`File saved as ${fileName}`);
}

// Example usage
const filePath = process.argv[2]; // Accept file path as a command-line argument
const mode = process.argv[3] === 'jz' ? 'donations' : 'xclick'; // Determine mode based on argument
const emails = readEmailsFromFile(filePath);
const emailUrlPairs = generatePayPalUrls(emails, mode);

// Save the generated URLs to a new file
saveUrlsToFile(emailUrlPairs);
