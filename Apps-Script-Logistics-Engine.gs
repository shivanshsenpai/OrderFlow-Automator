/**
 * ==============================================================
 * 🚀 GOOGLE SHEETS ORDER MANAGEMENT & SYNC AUTOMATION
 * ==============================================================
 * * SETUP INSTRUCTIONS:
 * Update the CONFIG object below with your specific Sheet names, 
 * API endpoints, and Column Headers before running the script.
 */

const CONFIG = {
  sheets: {
    newOrders: "New_Orders",       
    allOrders: "All Orders",       
    productCost: "Product Cost Master",   
    shipping: "Shipping Rules"   
  },
  api: {
    exchangeRateUrl: "https://open.er-api.com/v6/latest/INR",
    baseCurrency: "INR" // Ensure this matches your API endpoint base
  },
  colors: {
    errorBackground: "#ff0000", 
    defaultBackground: "#ffffff" 
  },
  columns: {
    // Primary Identifiers
    orderId: "Order ID",
    internalRef: "Internal Reference ID",
    brandName: "Store/Brand Name",
    salesChannel: "Sales Channel",
    customerName: "Customer Name",
    orderDate: "Order Date",
    sku: "Item SKU",
    
    // Product & Shipping
    productCost: "Unit Cost",
    country: "Destination Country",
    shippingCost: "Shipping Cost",
    
    // Media
    imageUrl: "Image Source URL",
    imageFormula: "Image Preview",
    
    // Financials
    currency: "Order Currency",
    originalPrice: "Original Price",
    exchangeRate: "Exchange Rate",
    basePrice: "Base Currency Price",
    maxExpense: "Estimated Max Expense",
    actualExpense: "Actual Total Expense",
    maxProfit: "Estimated Profit",
    actualProfit: "Actual Net Profit",
    
    // Shipping Sheet Specifics
    brandPrefix: "Brand Prefix",
    lastSequence: "Last Sequence Number",
    uniqueChannel: "Exclusive Channel"
  },
  // Columns that are NOT required. If a column is missing from this list, 
  // it is treated as MANDATORY and will highlight red if left empty.
  optionalColumns: [
    "Delivery Status",
    "Payment Status",
    "Staff Notes",
    "Customer Phone",
    "Image Preview",
    "Courier Partner",
    "Tracking Number",
    "Order Status",
    "Shipping Cost",
    "Product URL",
    "Updated Tracking",
    "Exchange Rate",
    "Base Currency Price",
    "Product Shipping Cost",
    "Unit Cost",
    "Estimated Max Expense",
    "Actual Total Expense",
    "Estimated Profit",
    "Actual Net Profit",
    "Platform Fees & Taxes",
    "Return Date",
    "Return Courier",
    "Return Tracking Number",
    "Return Received Date",
    "Dispute Claims",
    "Internal Reference ID"
  ]
};

// ==============================
// 🚀 FULL SYNC: Rates + Product Cost (All Orders sheet)
// ==============================
function runIncompleteSync() {
  updateRatesAndINR();
  updateProductCostFromSheet();
  SpreadsheetApp.getUi().alert("✅ Full sync complete!\n\n• Exchange rates updated\n• Product costs updated");
}

function updateRatesAndINR() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.newOrders);
  if (!sheet) return;

  const startRow = 3;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < startRow) return;

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  
  const currencyColIdx = headers.indexOf(CONFIG.columns.currency) + 1;
  const priceColIdx = headers.indexOf(CONFIG.columns.originalPrice) + 1;
  const conversionColIdx = headers.indexOf(CONFIG.columns.exchangeRate) + 1;
  const basePriceColIdx = headers.indexOf(CONFIG.columns.basePrice) + 1;

  if (currencyColIdx === 0 || priceColIdx === 0 || conversionColIdx === 0 || basePriceColIdx === 0) {
     console.log("❌ Missing required financial columns for sync.");
     return;
  }

  const currencyValues = sheet.getRange(startRow, currencyColIdx, lastRow - 2, 1).getValues();
  const priceValues = sheet.getRange(startRow, priceColIdx, lastRow - 2, 1).getValues();

  let rates = {};

  // 🔥 Get exchange rates safely
  try {
    const response = UrlFetchApp.fetch(CONFIG.api.exchangeRateUrl, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());

    if (data && data.rates) {
      rates = data.rates;
    } else {
      throw new Error("Invalid API response");
    }
  } catch (e) {
    console.log("Failed to fetch exchange rates: " + e);
    return;
  }

  const conversionRates = [];
  const priceBase = [];

  for (let i = 0; i < currencyValues.length; i++) {
    const currency = (currencyValues[i][0] || "").toString().trim().toUpperCase();
    const price = parseFloat(priceValues[i][0]);

    let rate = "";

    if (!currency) {
      conversionRates.push([""]);
      priceBase.push([""]);
      continue;
    }

    if (currency === CONFIG.api.baseCurrency) {
      rate = 1;
    } else if (rates[currency]) {
      // 🔁 Convert correctly (Base → invert)
      rate = 1 / rates[currency];
    } else {
      rate = ""; // Unknown currency
    }

    conversionRates.push([rate]);

    // 💰 Calculate Base price
    if (!isNaN(price) && rate !== "") {
      priceBase.push([price * rate]);
    } else {
      priceBase.push([""]);
    }
  }

  // ✅ Write once (fast + safe)
  sheet.getRange(startRow, conversionColIdx, conversionRates.length, 1).setValues(conversionRates);
  sheet.getRange(startRow, basePriceColIdx, priceBase.length, 1).setValues(priceBase);

  console.log("✅ Rates & Base Price updated successfully");
}

// manual push for images to turn from link to image 
function applyImageFormulasToAllOrders(startRow, numRows) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheets.allOrders);
  if (!sheet) return;
  
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  const imageUrlCol = headers.indexOf(CONFIG.columns.imageUrl) + 1;
  const imageCol = headers.indexOf(CONFIG.columns.imageFormula) + 1;

  if (imageUrlCol === 0 || imageCol === 0) return;

  const imageUrls = sheet.getRange(startRow, imageUrlCol, numRows, 1).getValues();

  const formulas = imageUrls.map((row, i) => {
    const url = (row[0] || "").toString().trim();
    if (!url) return [""];

    const cellRef = sheet.getRange(startRow + i, imageUrlCol).getA1Notation();
    return [`=IMAGE(${cellRef})`];
  });

  sheet.getRange(startRow, imageCol, numRows, 1).setFormulas(formulas);
}

// bring in products
function updateProductCostFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName(CONFIG.sheets.allOrders);
  const costSheet = ss.getSheetByName(CONFIG.sheets.productCost);

  if (!orderSheet || !costSheet) return;

  const startRow = 3;
  const orderLastRow = orderSheet.getLastRow();
  const orderLastCol = orderSheet.getLastColumn();

  if (orderLastRow < startRow) return;

  const headers = orderSheet.getRange(2, 1, 1, orderLastCol).getValues()[0];

  const skuIndex = headers.indexOf(CONFIG.columns.sku);
  const productCostIndex = headers.indexOf(CONFIG.columns.productCost);

  if (skuIndex === -1 || productCostIndex === -1) {
    console.log("❌ SKU or Product Cost column not found!");
    return;
  }

  const skuData = orderSheet.getRange(startRow, skuIndex + 1, orderLastRow - 2, 1).getValues();

  const costLastRow = costSheet.getLastRow();
  if (costLastRow < 2) return;

  const costData = costSheet.getRange(2, 1, costLastRow - 1, 2).getValues();

  const costMap = {};
  costData.forEach((row) => {
    const sku = (row[0] || "").toString().trim();
    if (sku) costMap[sku] = row[1];
  });

  const output = [];
  let updatedCount = 0;

  for (let i = 0; i < skuData.length; i++) {
    const sku = (skuData[i][0] || "").toString().trim();

    if (sku && costMap.hasOwnProperty(sku)) {
      output.push([costMap[sku]]);
      updatedCount++;
    } else {
      output.push([""]);
    }
  }

  orderSheet.getRange(startRow, productCostIndex + 1, output.length, 1).setValues(output);
  console.log("✅ Product cost updated. Rows processed: " + updatedCount);
}

// ==============================
// 🔄 UNIFIED onEdit (BOTH SHEETS)
// ==============================
function onEdit(e) {
  if (!e) return;

  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  const validSheets = [CONFIG.sheets.allOrders, CONFIG.sheets.newOrders];
  if (!validSheets.includes(sheetName)) return;
  if (row <= 2) return;

  const ss = e.source;
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  // ==============================
  // ✅ COLOR-BASED VALIDATION (New_Orders ONLY)
  // ==============================
  if (sheetName === CONFIG.sheets.newOrders) {
    const orderIdCol = headers.indexOf(CONFIG.columns.orderId) + 1;
    const redColor = CONFIG.colors.errorBackground;
    const whiteColor = CONFIG.colors.defaultBackground;

    if (col === orderIdCol) {
      const numEditedRows = e.range.getNumRows();
      const range = sheet.getRange(row, 1, numEditedRows, headers.length);
      const values = range.getValues();
      const backgrounds = range.getBackgrounds();

      for (let r = 0; r < numEditedRows; r++) {
        const pid = (values[r][orderIdCol - 1] || "").toString().trim();

        for (let c = 0; c < headers.length; c++) {
          if (pid) {
            if (!CONFIG.optionalColumns.includes(headers[c])) {
              const val = (values[r][c] || "").toString().trim();
              backgrounds[r][c] = val === "" ? redColor : whiteColor;
            } else {
              backgrounds[r][c] = whiteColor;
            }
          } else {
            backgrounds[r][c] = whiteColor;
          }
        }
      }
      range.setBackgrounds(backgrounds);

    } else {
      const rowOrderId = (sheet.getRange(row, orderIdCol).getValue() || "").toString().trim();
      if (rowOrderId) {
        updateCellColor(sheet, row, col, headers, CONFIG.optionalColumns);
      }
    }

    // ==============================
    // 🔢 AUTO INTERNAL ORDER NUMBER (New_Orders ONLY)
    // ==============================
    const brandNameCol = headers.indexOf(CONFIG.columns.brandName) + 1;
    const orderIdColForION = headers.indexOf(CONFIG.columns.orderId) + 1;
    const internalRefCol = headers.indexOf(CONFIG.columns.internalRef) + 1;
    const customerNameCol = headers.indexOf(CONFIG.columns.customerName) + 1;
    const salesChannelCol = headers.indexOf(CONFIG.columns.salesChannel) + 1;
    const orderDateCol = headers.indexOf(CONFIG.columns.orderDate) + 1;

    const editStartCol = e.range.getColumn();
    const editEndCol = editStartCol + e.range.getNumColumns() - 1;
    
    const brandOrPortalEdited =
      (brandNameCol > 0 && brandNameCol >= editStartCol && brandNameCol <= editEndCol) ||
      (orderIdColForION > 0 && orderIdColForION >= editStartCol && orderIdColForION <= editEndCol) ||
      (customerNameCol > 0 && customerNameCol >= editStartCol && customerNameCol <= editEndCol) ||
      (salesChannelCol > 0 && salesChannelCol >= editStartCol && salesChannelCol <= editEndCol) ||
      (orderDateCol > 0 && orderDateCol >= editStartCol && orderDateCol <= editEndCol);

    const internalOrderEdited = internalRefCol > 0 && internalRefCol >= editStartCol && internalRefCol <= editEndCol;

    if (brandOrPortalEdited && brandNameCol > 0 && orderIdColForION > 0) {
      const numEditedRows = e.range.getNumRows();
      generateInternalOrderNo(ss, sheet, row, numEditedRows, headers);
    } else if (internalOrderEdited && !brandOrPortalEdited) {
      recalcInternalOrderNos(ss, sheet, row, headers);
    }
  }

  // ==============================
  // 📍 COLUMN INDEXES
  // ==============================
  const skuCol = headers.indexOf(CONFIG.columns.sku) + 1;
  const productCostCol = headers.indexOf(CONFIG.columns.productCost) + 1;
  const countryCol = headers.indexOf(CONFIG.columns.country) + 1;
  const imageUrlCol = headers.indexOf(CONFIG.columns.imageUrl) + 1;
  const imageCol = headers.indexOf(CONFIG.columns.imageFormula) + 1;
  const shippingCol = headers.indexOf(CONFIG.columns.shippingCost) + 1;
  const currencyCol = headers.indexOf(CONFIG.columns.currency) + 1;
  const currencyPriceCol = headers.indexOf(CONFIG.columns.originalPrice) + 1;
  const conversionCol = headers.indexOf(CONFIG.columns.exchangeRate) + 1;
  const basePriceCol = headers.indexOf(CONFIG.columns.basePrice) + 1;
  const maxExpenseCol = headers.indexOf(CONFIG.columns.maxExpense) + 1;

  // ==============================
  // 🔹 SKU → PRODUCT COST AUTO-FILL
  // ==============================
  if (col === skuCol && skuCol > 0) {
    const costSheet = ss.getSheetByName(CONFIG.sheets.productCost);
    if (costSheet && productCostCol > 0) {
      const sku = (e.range.getValue() || "").toString().trim();

      if (!sku) {
        sheet.getRange(row, productCostCol).setValue("");
      } else {
        const costData = costSheet.getRange(2, 1, costSheet.getLastRow() - 1, 2).getValues();
        const costMap = {};
        costData.forEach((r) => {
          const key = (r[0] || "").toString().trim();
          if (key) costMap[key] = r[1];
        });
        sheet.getRange(row, productCostCol).setValue(costMap[sku] || "");
      }
    }
  }

  // ==============================
  // 🖼️ IMAGE URL → IMAGE FORMULA
  // ==============================
  if (col === imageUrlCol && imageUrlCol > 0 && imageCol > 0) {
    const imageUrl = (e.range.getValue() || "").toString().trim();
    if (!imageUrl) {
      sheet.getRange(row, imageCol).setValue("");
    } else {
      const formula = `=IMAGE(${sheet.getRange(row, imageUrlCol).getA1Notation()})`;
      sheet.getRange(row, imageCol).setFormula(formula);
    }
  }

  // ==============================
  // 🔹 COUNTRY → SHIPPING
  // ==============================
  if (col === countryCol && countryCol > 0) {
    const shippingSheet = ss.getSheetByName(CONFIG.sheets.shipping);
    if (shippingSheet && shippingCol > 0) {
      let country = (e.range.getValue() || "").toString().trim().toLowerCase();

      if (!country || country === "") {
        sheet.getRange(row, shippingCol).setValue("");
      } else {
        const data = shippingSheet.getRange(2, 1, shippingSheet.getLastRow() - 1, 2).getValues();
        const chargeMap = {};
        data.forEach((r) => {
          const c = (r[0] || "").toString().trim().toLowerCase();
          if (c) chargeMap[c] = r[1];
        });
        sheet.getRange(row, shippingCol).setValue(chargeMap[country] || "");
      }
    }
  }

  // ==============================
  // 🔹 CURRENCY → CONVERSION + BASE PRICE
  // ==============================
  if ((col === currencyCol || col === currencyPriceCol) && currencyCol > 0) {
    const currency = (sheet.getRange(row, currencyCol).getValue() || "").toString().trim().toUpperCase();
    const price = parseFloat(sheet.getRange(row, currencyPriceCol).getValue());

    let rate = "";
    if (currency) {
      if (currency === CONFIG.api.baseCurrency) {
        rate = 1;
      } else {
        try {
          const response = UrlFetchApp.fetch(CONFIG.api.exchangeRateUrl, { muteHttpExceptions: true });
          const data = JSON.parse(response.getContentText());
          if (data?.rates?.[currency]) rate = 1 / data.rates[currency];
        } catch (err) {
          rate = "";
        }
      }
    }

    if (conversionCol > 0) sheet.getRange(row, conversionCol).setValue(rate);
    
    if (basePriceCol > 0) {
      if (!isNaN(price) && rate !== "") {
        sheet.getRange(row, basePriceCol).setValue(price * rate);
      } else {
        sheet.getRange(row, basePriceCol).setValue("");
      }
    }
  }

  // ==============================
  // 💰 AUTO-CALCULATE FINANCIALS
  // ==============================
  if (
    (col === basePriceCol || col === productCostCol || col === shippingCol ||
      col === countryCol || col === currencyCol || col === currencyPriceCol || col === conversionCol) &&
    maxExpenseCol > 0
  ) {
    calculateFinancials(sheet, row, ss, headers);
  }
}

// ==============================
// 💰 HELPER: Calculate Financials
// ==============================
function calculateFinancials(sheet, row, ss, headers) {
  const shippingSheet = ss.getSheetByName(CONFIG.sheets.shipping);
  if (!shippingSheet) return;

  const basePriceCol = headers.indexOf(CONFIG.columns.basePrice) + 1;
  const productCostCol = headers.indexOf(CONFIG.columns.productCost) + 1;
  const shippingCol = headers.indexOf(CONFIG.columns.shippingCost) + 1;
  const countryCol = headers.indexOf(CONFIG.columns.country) + 1;
  const maxExpenseCol = headers.indexOf(CONFIG.columns.maxExpense) + 1;
  const actualExpenseCol = headers.indexOf(CONFIG.columns.actualExpense) + 1;
  const maxProfitCol = headers.indexOf(CONFIG.columns.maxProfit) + 1;
  const actualProfitCol = headers.indexOf(CONFIG.columns.actualProfit) + 1;

  if (basePriceCol === 0 || productCostCol === 0 || shippingCol === 0 || countryCol === 0 || maxExpenseCol === 0) return;

  const price = parseFloat(sheet.getRange(row, basePriceCol).getValue()) || 0;
  const productCost = parseFloat(sheet.getRange(row, productCostCol).getValue()) || 0;
  const shipping = parseFloat(sheet.getRange(row, shippingCol).getValue()) || 0;
  let country = (sheet.getRange(row, countryCol).getValue() || "").toString().trim().toLowerCase();

  if (!price || !productCost || !shipping || !country) return;

  const shipData = shippingSheet.getRange(2, 1, shippingSheet.getLastRow() - 1, 3).getValues();
  let magic = 0;
  for (let i = 0; i < shipData.length; i++) {
    if ((shipData[i][0] || "").toString().trim().toLowerCase() === country) {
      magic = parseFloat(shipData[i][2]) || 0;
      break;
    }
  }

  const maxExpense = price * magic;
  const actualExpense = productCost + shipping;
  const maxProfit = price * 0.2;
  const actualProfit = maxExpense - actualExpense + maxProfit;

  if (maxExpenseCol > 0) sheet.getRange(row, maxExpenseCol).setValue(maxExpense);
  if (actualExpenseCol > 0) sheet.getRange(row, actualExpenseCol).setValue(actualExpense);
  if (maxProfitCol > 0) sheet.getRange(row, maxProfitCol).setValue(maxProfit);
  if (actualProfitCol > 0) sheet.getRange(row, actualProfitCol).setValue(actualProfit);
}

// ==============================
// 🎨 HELPER: Update Cell Color on Edit
// ==============================
function updateCellColor(sheet, row, col, headers, optionalFields) {
  const header = headers[col - 1];
  if (!header) return;

  const cellRange = sheet.getRange(row, col);
  const cellValue = (cellRange.getValue() || "").toString().trim();

  if (!optionalFields.includes(header)) {
    cellRange.setBackground(cellValue === "" ? CONFIG.colors.errorBackground : CONFIG.colors.defaultBackground);
  } else {
    cellRange.setBackground(CONFIG.colors.defaultBackground);
  }
}

// ==============================
// 📤 MANUAL: Push All New Orders to All Orders
// ==============================
function pushAllNewOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newOrdersSheet = ss.getSheetByName(CONFIG.sheets.newOrders);
  if (!newOrdersSheet) {
    SpreadsheetApp.getUi().alert(`❌ ${CONFIG.sheets.newOrders} sheet not found!`);
    return;
  }

  const startRow = 3;
  const lastRow = newOrdersSheet.getLastRow();
  const lastCol = newOrdersSheet.getLastColumn();

  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert(`ℹ️ No data rows in ${CONFIG.sheets.newOrders}.`);
    return;
  }

  const headers = newOrdersSheet.getRange(2, 1, 1, lastCol).getValues()[0];

  // ─── Step 1: Force Financial recalculation ───
  for (let row = startRow; row <= lastRow; row++) {
    calculateFinancials(newOrdersSheet, row, ss, headers);
  }
  SpreadsheetApp.flush();

  // ─── Step 2: Validate all rows ───
  const orderIdColIdx = headers.indexOf(CONFIG.columns.orderId);
  const incompleteIds = [];
  const numDataRows = lastRow - startRow + 1;
  const sheetData = newOrdersSheet.getRange(startRow, 1, numDataRows, lastCol).getValues();

  for (let r = 0; r < numDataRows; r++) {
    const rowValues = sheetData[r];
    const orderId = orderIdColIdx !== -1 ? (rowValues[orderIdColIdx] || "").toString().trim() : "";
    if (!orderId) continue;

    let rowComplete = true;
    for (let c = 0; c < headers.length; c++) {
      if (CONFIG.optionalColumns.includes(headers[c])) continue; 
      if ((rowValues[c] || "").toString().trim() === "") {
        rowComplete = false;
        break;
      }
    }

    if (!rowComplete) incompleteIds.push(orderId);
  }

  if (incompleteIds.length > 0) {
    SpreadsheetApp.getUi().alert(
      `⚠️ The following ${CONFIG.columns.orderId}s have empty mandatory fields:\n\n` +
      incompleteIds.join("\n") + "\n\nPlease fill them before pushing."
    );
    return;
  }

  // ─── Step 3: All rows valid – copy to All Orders ───
  const allOrdersSheet = ss.getSheetByName(CONFIG.sheets.allOrders);
  if (!allOrdersSheet) return;

  const existingRowSet = new Set();
  const aoLastRow = allOrdersSheet.getLastRow();
  const imageColIdx = headers.indexOf(CONFIG.columns.imageFormula);

  if (aoLastRow >= 3) {
    const aoData = allOrdersSheet.getRange(3, 1, aoLastRow - 2, lastCol).getValues();
    aoData.forEach(row => {
      const key = row.map((cell, idx) => idx === imageColIdx ? "" : (cell || "").toString().trim()).join("|||");
      existingRowSet.add(key);
    });
  }

  const freshData = newOrdersSheet.getRange(startRow, 1, numDataRows, lastCol).getValues();

  const rowsToPush = freshData
    .filter((row) => {
      const oid = orderIdColIdx !== -1 ? (row[orderIdColIdx] || "").toString().trim() : "";
      if (oid === "") return false;

      const key = row.map((cell, idx) => idx === imageColIdx ? "" : (cell || "").toString().trim()).join("|||");
      return !existingRowSet.has(key);
    })
    .map(row => {
      if (imageColIdx === -1) return row;
      const newRow = [...row];
      newRow[imageColIdx] = "";
      return newRow;
    });

  if (rowsToPush.length === 0) {
    SpreadsheetApp.getUi().alert(`ℹ️ No new orders to copy. Everything already exists in ${CONFIG.sheets.allOrders}.`);
    return;
  }

  const insertAt = allOrdersSheet.getLastRow() + 1;
  allOrdersSheet.getRange(insertAt, 1, rowsToPush.length, headers.length).setValues(rowsToPush);
  applyImageFormulasToAllOrders(insertAt, rowsToPush.length);

  SpreadsheetApp.getUi().alert(`✅ ${rowsToPush.length} new order(s) copied to ${CONFIG.sheets.allOrders}!\n\nOriginals kept in ${CONFIG.sheets.newOrders}.`);
}

// ==============================
// 🔄 SYNC: Pull Last Sequence from All Orders → Shipping
// ==============================
function syncLastMaxFromAllOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shippingSheet = ss.getSheetByName(CONFIG.sheets.shipping);
  const allOrdersSheet = ss.getSheetByName(CONFIG.sheets.allOrders);

  if (!shippingSheet || !allOrdersSheet) return;

  const shipLastCol = shippingSheet.getLastColumn();
  const shipLastRow = shippingSheet.getLastRow();
  const shipHeaders = shippingSheet.getRange(1, 1, 1, shipLastCol).getValues()[0];

  const shipInitialCol = shipHeaders.indexOf(CONFIG.columns.brandPrefix);
  const shipLastMaxCol = shipHeaders.indexOf(CONFIG.columns.lastSequence);

  if (shipInitialCol === -1 || shipLastMaxCol === -1 || shipLastRow < 2) return;

  const shipData = shippingSheet.getRange(2, 1, shipLastRow - 1, shipLastCol).getValues();
  const prefixMaxMap = {}; 

  shipData.forEach(row => {
    const initial = (row[shipInitialCol] || "").toString().trim().toUpperCase();
    if (initial) prefixMaxMap[initial] = 10000; 
  });

  const scanSheetForMax = (targetSheet) => {
    if (targetSheet.getLastRow() >= 3) {
      const headers = targetSheet.getRange(2, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      const internalCol = headers.indexOf(CONFIG.columns.internalRef);
      if (internalCol !== -1) {
        const data = targetSheet.getRange(3, internalCol + 1, targetSheet.getLastRow() - 2, 1).getValues();
        data.forEach(row => {
          const match = (row[0] || "").toString().trim().match(/^([A-Za-z]+)(\d+)$/);
          if (match && prefixMaxMap.hasOwnProperty(match[1].toUpperCase())) {
            const num = parseInt(match[2], 10);
            if (num > prefixMaxMap[match[1].toUpperCase()]) {
              prefixMaxMap[match[1].toUpperCase()] = num;
            }
          }
        });
      }
    }
  };

  scanSheetForMax(allOrdersSheet);
  scanSheetForMax(ss.getSheetByName(CONFIG.sheets.newOrders));

  const lastMaxOutput = shipData.map(row => {
    const initial = (row[shipInitialCol] || "").toString().trim().toUpperCase();
    return initial && prefixMaxMap.hasOwnProperty(initial) ? [prefixMaxMap[initial]] : [""];
  });

  shippingSheet.getRange(2, shipLastMaxCol + 1, lastMaxOutput.length, 1).setValues(lastMaxOutput);

  const summary = Object.entries(prefixMaxMap).map(([prefix, max]) => `${prefix} → ${max}`).join("\n");
  SpreadsheetApp.getUi().alert(`✅ Sequences synced from ${CONFIG.sheets.allOrders}!\n\n` + summary);
}

// ==============================
// 🔢 AUTO INTERNAL ORDER NUMBER GENERATOR
// ==============================
function generateInternalOrderNo(ss, sheet, startRow, numRows, headers) {
  const brandNameColIdx = headers.indexOf(CONFIG.columns.brandName);
  const orderIdColIdx = headers.indexOf(CONFIG.columns.orderId);
  const internalRefColIdx = headers.indexOf(CONFIG.columns.internalRef);
  const salesChannelColIdx = headers.indexOf(CONFIG.columns.salesChannel);
  const customerNameColIdx = headers.indexOf(CONFIG.columns.customerName);
  const orderDateColIdx = headers.indexOf(CONFIG.columns.orderDate);

  if (internalRefColIdx === -1 || brandNameColIdx === -1 || orderIdColIdx === -1) return;

  const shippingSheet = ss.getSheetByName(CONFIG.sheets.shipping);
  if (!shippingSheet) return;

  const shipLastRow = shippingSheet.getLastRow();
  const shipLastCol = shippingSheet.getLastColumn();
  if (shipLastRow < 2) return;

  const shipHeaders = shippingSheet.getRange(1, 1, 1, shipLastCol).getValues()[0];
  const shipBrandCol = shipHeaders.indexOf(CONFIG.columns.brandName);
  const shipInitialCol = shipHeaders.indexOf(CONFIG.columns.brandPrefix);
  const shipLastMaxCol = shipHeaders.indexOf(CONFIG.columns.lastSequence);
  const shipUniqueSCCol = shipHeaders.indexOf(CONFIG.columns.uniqueChannel);

  if (shipBrandCol === -1 || shipInitialCol === -1 || shipLastMaxCol === -1) return;

  const shipData = shippingSheet.getRange(2, 1, shipLastRow - 1, shipLastCol).getValues();

  const brandInitialMap = {};  
  const brandMaxMap = {};      
  const uniqueSalesChannels = new Set(); 

  shipData.forEach(row => {
    const brand = (row[shipBrandCol] || "").toString().trim().toLowerCase();
    const initial = (row[shipInitialCol] || "").toString().trim();
    if (brand && initial) {
      brandInitialMap[brand] = initial;
      brandMaxMap[brand] = parseInt(row[shipLastMaxCol]) || 10000;
    }
    if (shipUniqueSCCol !== -1) {
      const sc = (row[shipUniqueSCCol] || "").toString().trim().toLowerCase();
      if (sc) uniqueSalesChannels.add(sc);
    }
  });

  const dataStartRow = 3;
  if (sheet.getLastRow() < dataStartRow) return;

  const allNewData = sheet.getRange(dataStartRow, 1, sheet.getLastRow() - dataStartRow + 1, headers.length).getValues();
  const groupToNumber = {};  
  const fullOutput = [];

  for (let i = 0; i < allNewData.length; i++) {
    const row = allNewData[i];
    const orderId = (row[orderIdColIdx] || "").toString().trim();
    const brandName = (row[brandNameColIdx] || "").toString().trim();
    const existingVal = (row[internalRefColIdx] || "").toString().trim();
    
    if (!orderId || !brandName || !brandInitialMap[brandName.toLowerCase()]) {
      fullOutput.push([existingVal]);
      continue;
    }

    const brandLower = brandName.toLowerCase();
    const prefix = brandInitialMap[brandLower].toUpperCase();
    const portalKey = `portal|||${orderId.toLowerCase()}|||${brandLower}`;

    let assignedNumber = groupToNumber[portalKey] || null;

    if (!assignedNumber && i > 0 && salesChannelColIdx !== -1 && customerNameColIdx !== -1) {
      const salesChannel = (row[salesChannelColIdx] || "").toString().trim().toLowerCase();
      if ([...uniqueSalesChannels].some(usc => salesChannel.includes(usc))) {
        const customerName = (row[customerNameColIdx] || "").toString().trim().toLowerCase();
        const purchaseDate = orderDateColIdx !== -1 ? (row[orderDateColIdx] || "").toString().trim().toLowerCase() : "";

        if (customerName) {
          const prevRow = allNewData[i - 1];
          const prevChannel = (prevRow[salesChannelColIdx] || "").toString().trim().toLowerCase();
          
          if (
            (prevRow[brandNameColIdx] || "").toString().trim().toLowerCase() === brandLower &&
            (prevRow[customerNameColIdx] || "").toString().trim().toLowerCase() === customerName &&
            [...uniqueSalesChannels].some(usc => prevChannel.includes(usc)) &&
            (orderDateColIdx !== -1 ? (prevRow[orderDateColIdx] || "").toString().trim().toLowerCase() === purchaseDate : true) &&
            fullOutput[i - 1] && fullOutput[i - 1][0]
          ) {
            assignedNumber = fullOutput[i - 1][0];
          }
        }
      }
    }

    if (assignedNumber) {
      fullOutput.push([assignedNumber]);
      groupToNumber[portalKey] = assignedNumber;
    } else {
      const newNum = (brandMaxMap[brandLower] || 10000) + 1;
      const newOrderNo = prefix + newNum;
      brandMaxMap[brandLower] = newNum;
      groupToNumber[portalKey] = newOrderNo;
      fullOutput.push([newOrderNo]);
    }
  }

  sheet.getRange(dataStartRow, internalRefColIdx + 1, fullOutput.length, 1).setValues(fullOutput);
}

// ==============================
// 🔄 RECALC: Close gaps in Internal OrderNo
// ==============================
function recalcInternalOrderNos(ss, sheet, editedRow, headers) {
  const brandNameColIdx = headers.indexOf(CONFIG.columns.brandName);
  const internalRefColIdx = headers.indexOf(CONFIG.columns.internalRef);

  if (internalRefColIdx === -1 || brandNameColIdx === -1) return;

  const dataStartRow = 3;
  if (sheet.getLastRow() < dataStartRow) return;

  const allData = sheet.getRange(dataStartRow, 1, sheet.getLastRow() - dataStartRow + 1, headers.length).getValues();
  const editedBrand = (allData[editedRow - dataStartRow]?.[brandNameColIdx] || "").toString().trim().toLowerCase();
  if (!editedBrand) return;

  const shippingSheet = ss.getSheetByName(CONFIG.sheets.shipping);
  if (!shippingSheet) return;

  const shipHeaders = shippingSheet.getRange(1, 1, 1, shippingSheet.getLastColumn()).getValues()[0];
  const shipBrandCol = shipHeaders.indexOf(CONFIG.columns.brandName);
  const shipInitialCol = shipHeaders.indexOf(CONFIG.columns.brandPrefix);
  const shipLastMaxCol = shipHeaders.indexOf(CONFIG.columns.lastSequence);

  if (shipBrandCol === -1 || shipInitialCol === -1 || shipLastMaxCol === -1) return;

  const shipData = shippingSheet.getRange(2, 1, shippingSheet.getLastRow() - 1, shippingSheet.getLastColumn()).getValues();
  let prefix = "", shipRowIdx = -1;

  for (let i = 0; i < shipData.length; i++) {
    if ((shipData[i][shipBrandCol] || "").toString().trim().toLowerCase() === editedBrand) {
      prefix = (shipData[i][shipInitialCol] || "").toString().trim().toUpperCase();
      shipRowIdx = i;
      break;
    }
  }
  if (!prefix) return;

  const brandRows = []; 
  for (let i = 0; i < allData.length; i++) {
    if ((allData[i][brandNameColIdx] || "").toString().trim().toLowerCase() !== editedBrand) continue;
    const match = (allData[i][internalRefColIdx] || "").toString().trim().match(/^([A-Za-z]+)(\d+)$/);
    if (match) brandRows.push(parseInt(match[2], 10));
  }

  if (brandRows.length === 0) return;

  const uniqueNumbers = [...new Set(brandRows)];
  const sortedUnique = [...uniqueNumbers].sort((a, b) => a - b);
  const base = sortedUnique[0]; 
  const sortedMapping = {};
  sortedUnique.forEach((num, idx) => sortedMapping[num] = base + idx);

  let newMax = 0;
  const fullOutput = allData.map(row => {
    if ((row[brandNameColIdx] || "").toString().trim().toLowerCase() !== editedBrand) return [row[internalRefColIdx]];

    const match = (row[internalRefColIdx] || "").toString().trim().match(/^([A-Za-z]+)(\d+)$/);
    if (match) {
      const newNum = sortedMapping[parseInt(match[2], 10)] || parseInt(match[2], 10);
      if (newNum > newMax) newMax = newNum;
      return [prefix + newNum];
    }
    return [row[internalRefColIdx]];
  });

  sheet.getRange(dataStartRow, internalRefColIdx + 1, fullOutput.length, 1).setValues(fullOutput);

  const aoSheet = ss.getSheetByName(CONFIG.sheets.allOrders);
  if (aoSheet && aoSheet.getLastRow() >= 3) {
    const aoHeaders = aoSheet.getRange(2, 1, 1, aoSheet.getLastColumn()).getValues()[0];
    const aoInternalCol = aoHeaders.indexOf(CONFIG.columns.internalRef);
    if (aoInternalCol !== -1) {
      const aoData = aoSheet.getRange(3, aoInternalCol + 1, aoSheet.getLastRow() - 2, 1).getValues();
      aoData.forEach(r => {
        const m = (r[0] || "").toString().trim().match(/^([A-Za-z]+)(\d+)$/);
        if (m && m[1].toUpperCase() === prefix && parseInt(m[2], 10) > newMax) newMax = parseInt(m[2], 10);
      });
    }
  }

  if (newMax > 0) shippingSheet.getRange(2 + shipRowIdx, shipLastMaxCol + 1).setValue(newMax);
}
