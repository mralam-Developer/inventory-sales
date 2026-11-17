/**
 * Google Apps Script — Shah and Company Stock API
 * Deploy as a Web App (GET, accessible by anyone / anonymous)
 * Returns sheet stock data + hardcoded product metadata as JSON.
 *
 * The POS loads this URL instead of static data.json.
 * Any change to the Stock sheet is reflected on the next page load.
 */

// ──────────────────────────────────────────────────────────
//  Product metadata — NOT in the sheet (sheet has no pricing,
//  only stock details). These values are used internally by
//  the POS for cart totals / checkout but are never shown.
//  Update these here when prices/suppliers change.
// ──────────────────────────────────────────────────────────
var META = {
  "STK001": { "sup": "ABC Traders",    "minStock": 30, "cost": 35,  "sell": 55,  "unit": "Pcs" },
  "STK002": { "sup": "XYZ Wholesale",  "minStock": 30, "cost": 20,  "sell": 35,  "unit": "Pcs" },
  "STK003": { "sup": "ABC Traders",    "minStock": 40, "cost": 12,  "sell": 20,  "unit": "Pcs" },
  "STK004": { "sup": "Local Market",   "minStock": 20, "cost": 45,  "sell": 70,  "unit": "Pcs" },
  "STK005": { "sup": "XYZ Wholesale",  "minStock": 20, "cost": 18,  "sell": 30,  "unit": "Pcs" },
  "STK006": { "sup": "ABC Traders",    "minStock": 30, "cost": 22,  "sell": 38,  "unit": "Pcs" },
  "STK007": { "sup": "Gas Co. Ltd",    "minStock": 30, "cost": 60,  "sell": 95,  "unit": "Pcs" },
  "STK008": { "sup": "Jewellery Sup.", "minStock": 25, "cost": 150, "sell": 220, "unit": "Pcs" },
  "STK009": { "sup": "ABC Traders",    "minStock": 15, "cost": 80,  "sell": 120, "unit": "Pcs" },
  "STK010": { "sup": "ABC Traders",    "minStock": 20, "cost": 8,   "sell": 15,  "unit": "Pcs" },
  "STK011": { "sup": "XYZ Wholesale",  "minStock": 20, "cost": 25,  "sell": 40,  "unit": "Pcs" },
  "STK012": { "sup": "ABC Traders",    "minStock": 15, "cost": 10,  "sell": 18,  "unit": "Pcs" },
  "STK013": { "sup": "XYZ Wholesale",  "minStock": 10, "cost": 55,  "sell": 85,  "unit": "Pcs" },
  "STK014": { "sup": "ABC Traders",    "minStock": 25, "cost": 15,  "sell": 25,  "unit": "Pcs" },
  "STK015": { "sup": "Local Market",   "minStock": 20, "cost": 110, "sell": 160, "unit": "Pcs" },
  "STK016": { "sup": "ABC Traders",    "minStock": 20, "cost": 40,  "sell": 60,  "unit": "Pcs" },
  "STK017": { "sup": "ABC Traders",    "minStock": 20, "cost": 30,  "sell": 45,  "unit": "Pcs" },
  "STK018": { "sup": "Pipe Co.",       "minStock": 30, "cost": 70,  "sell": 110, "unit": "Pcs" },
  "STK019": { "sup": "ABC Traders",    "minStock": 20, "cost": 20,  "sell": 32,  "unit": "Pcs" },
  "STK020": { "sup": "ABC Traders",    "minStock": 20, "cost": 28,  "sell": 42,  "unit": "Pcs" },
  "STK021": { "sup": "ABC Traders",    "minStock": 20, "cost": 35,  "sell": 50,  "unit": "Pcs" },
  "STK022": { "sup": "Pipe Co.",       "minStock": 15, "cost": 50,  "sell": 75,  "unit": "Pcs" },
  "STK023": { "sup": "XYZ Wholesale",  "minStock": 10, "cost": 45,  "sell": 65,  "unit": "Pcs" },
  "STK024": { "sup": "ABC Traders",    "minStock": 15, "cost": 90,  "sell": 130, "unit": "Pcs" }
};

// ──────────────────────────────────────────────────────────
//  doGet — entry point for the deployed Web App
// ──────────────────────────────────────────────────────────
function doGet(e) {
  var sheet = getStockSheet();
  if (!sheet) {
    return errorJson("Stock sheet not found");
  }

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return errorJson("Stock sheet is empty");
  }

  var items = [];
  var stockMap = {};

  // Row 0 = header: SKU | Part Name | Country | Shop Stock | Store Stock | Total Stock | Last Updated
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var sku = row[0];
    if (!sku || sku === "SKU") continue;

    var name     = row[1] || "";
    var country  = row[2] || "";
    var shopStr  = row[3];
    var storeStr = row[4];
    var totalVal = row[5];

    var shopStock  = (typeof shopStr === "number") ? shopStr : parseInt(shopStr, 10) || 0;
    var storeStock = (typeof storeStr === "number") ? storeStr : parseInt(storeStr, 10) || 0;
    var totalStock = (typeof totalVal === "number") ? totalVal : parseInt(totalVal, 10) || (shopStock + storeStock);

    var meta = META[sku] || {};
    var item = {
      "sku":         sku,
      "name":        name,
      "country":     country,
      "sup":         meta.sup         || "",
      "minStock":    meta.minStock    || 0,
      "cost":        meta.cost        || 0,
      "sell":        meta.sell        || 0,
      "unit":        meta.unit        || "Pcs",
      "shopStock":   shopStock,
      "storeStock":  storeStock,
      "totalStock":  totalStock
    };
    items.push(item);
    stockMap[sku] = totalStock;
  }

  var result = {
    "generated": new Date().toISOString(),
    "source":    "https://docs.google.com/spreadsheets/d/1yDxKuf4yT_dBP4fb2KT_uJ9PerrF0Cc2gvOJmZrpR8I/edit",
    "items":     items,
    "stock":     stockMap
  };

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ──────────────────────────────────────────────────────────
//  Spreadsheet ID — the sheet this script reads from.
//  Find it in the sheet URL:
//  https://docs.google.com/spreadsheets/d/1yDxKuf4yT_dBP4fb2KT_uJ9PerrF0Cc2gvOJmZrpR8I/edit
//  The ID is the part between /d/ and /edit
// ──────────────────────────────────────────────────────────
var SPREADSHEET_ID = "1yDxKuf4yT_dBP4fb2KT_uJ9PerrF0Cc2gvOJmZrpR8I";

// ──────────────────────────────────────────────────────────
//  Helpers
// ──────────────────────────────────────────────────────────
function getStockSheet() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!ss) return null;
  return ss.getSheetByName("Stock");
}

function errorJson(msg) {
  var r = { "error": msg, "items": [], "stock": {} };
  return ContentService.createTextOutput(JSON.stringify(r))
    .setMimeType(ContentService.MimeType.JSON);
}
