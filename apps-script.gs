/**
 * WASTE TO WATTS — Google Apps Script
 * =====================================
 * This script receives form submissions from the website
 * and writes them to this Google Sheet with calculated impact columns.
 *
 * SETUP INSTRUCTIONS (do this once):
 * ────────────────────────────────────
 * 1. Go to https://script.google.com
 * 2. Click "New Project"
 * 3. Delete any existing code and paste this entire file
 * 4. Click Save (name it "Waste To Watts Logger")
 * 5. Click Deploy → New Deployment
 * 6. Type: Web App
 * 7. Execute as: Me
 * 8. Who has access: Anyone
 * 9. Click Deploy → copy the Web App URL
 * 10. Open index.html in a text editor
 *     Find: const APPS_SCRIPT_URL = "YOUR_APPS_SCRIPT_URL_HERE";
 *     Replace with: const APPS_SCRIPT_URL = "paste_your_url_here";
 * 11. Save index.html — done!
 *
 * The sheet will be created automatically on the first submission.
 */

// ── Sheet column headers ──────────────────────────────────────────────────────
const HEADERS = [
  "Timestamp",
  "Name",
  "Email",
  "Location Type",
  "Food Type",
  "Food Waste (kg)",
  "Meals Rescued",
  "Contamination (%)",
  "Est. kWh Generated",       // kg × 0.2
  "CO₂e Avoided (kg)",        // kg × 0.55
  "Home Power Days",          // kWh ÷ 10
  "Equivalent Car km Avoided" // CO₂e ÷ 0.21 kg/km
];

// ── Conversion factors ────────────────────────────────────────────────────────
const KWH_PER_KG  = 0.2;   // kWh per kg food waste (prototype estimate)
const CO2_PER_KG  = 0.55;  // kg CO₂e avoided per kg diverted from landfill
const KG_CO2_PER_KM = 0.21; // avg passenger car CO₂e per km (Canada)

// ── Main handler ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const kg            = Number(data.kgWaste)      || 0;
    const kwh           = +(kg * KWH_PER_KG).toFixed(3);
    const co2           = +(kg * CO2_PER_KG).toFixed(3);
    const homePowerDays = +(kwh / 10).toFixed(3);
    const carKm         = +(co2 / KG_CO2_PER_KM).toFixed(1);

    const sheet = getOrCreateSheet();

    sheet.appendRow([
      new Date(data.timestamp || Date.now()), // Timestamp
      data.name            || "",
      data.email           || "",
      data.locationType    || "",
      data.foodType        || "",
      kg,
      Number(data.mealsRescued)   || 0,
      Number(data.contamination)  || 0,
      kwh,
      co2,
      homePowerDays,
      carKm
    ]);

    // Auto-format new row (alternating row colour, number format)
    const lastRow = sheet.getLastRow();
    formatRow(sheet, lastRow);

    return jsonResponse({ success: true, row: lastRow });

  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

// ── Sheet setup ──────────────────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName("Donations");

  if (!sheet) {
    sheet = ss.insertSheet("Donations");
    setupHeaders(sheet);
  }
  return sheet;
}

function setupHeaders(sheet) {
  // Write headers
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  // Style header row
  const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setBackground("#0d4f2c");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Column widths
  sheet.setColumnWidth(1, 160);   // Timestamp
  sheet.setColumnWidth(2, 140);   // Name
  sheet.setColumnWidth(3, 190);   // Email
  sheet.setColumnWidth(4, 150);   // Location Type
  sheet.setColumnWidth(5, 160);   // Food Type
  sheet.setColumnWidth(6, 120);   // kg
  sheet.setColumnWidth(7, 120);   // Meals Rescued
  sheet.setColumnWidth(8, 140);   // Contamination
  sheet.setColumnWidth(9, 150);   // kWh
  sheet.setColumnWidth(10, 150);  // CO₂e
  sheet.setColumnWidth(11, 150);  // Home Power Days
  sheet.setColumnWidth(12, 180);  // Car km
}

function formatRow(sheet, rowIndex) {
  const range = sheet.getRange(rowIndex, 1, 1, HEADERS.length);
  const bg    = rowIndex % 2 === 0 ? "#f1faf2" : "#ffffff";
  range.setBackground(bg);

  // Format numeric columns
  sheet.getRange(rowIndex, 6).setNumberFormat("0.0");   // kg
  sheet.getRange(rowIndex, 9).setNumberFormat("0.000"); // kWh
  sheet.getRange(rowIndex, 10).setNumberFormat("0.000"); // CO₂e
  sheet.getRange(rowIndex, 11).setNumberFormat("0.000"); // home days
  sheet.getRange(rowIndex, 12).setNumberFormat("0.0");   // car km
  sheet.getRange(rowIndex, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
}

// ── Utility ──────────────────────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── CORS preflight (GET request returns OK) ───────────────────────────────────
function doGet() {
  return jsonResponse({ status: "Waste To Watts logger is active." });
}

/**
 * TEST FUNCTION — run this manually in the Apps Script editor
 * to verify the sheet is created and formatted correctly.
 */
function testInsert() {
  doPost({
    postData: {
      contents: JSON.stringify({
        name: "Test User",
        email: "test@example.com",
        locationType: "School cafeteria",
        foodType: "Fruit & Vegetables",
        kgWaste: 12.5,
        mealsRescued: 4,
        contamination: 3,
        timestamp: new Date().toISOString()
      })
    }
  });
}
