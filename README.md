# 🚀 OrderFlow Automator for Google Sheets

A powerful, unified Google Apps Script designed to automate and streamline e-commerce order management within Google Sheets. This script handles everything from real-time currency conversion and automatic cost-mapping to strict data validation and automated internal order numbering.

## ✨ Features

* **Real-Time Currency Conversion:** Automatically fetches the latest exchange rates (via Open Exchange Rates API) and converts international orders to a base currency (INR).
* **Automated SKU Mapping:** Instantly pulls product costs from a centralized "ProductCost" sheet the moment an SKU is entered.
* **Dynamic Shipping Calculation:** Calculates shipping charges automatically based on the destination country.
* **Smart Financials:** Automatically computes Maximum Expense, Actual Expense, Maximum Profit, and Actual Profit based on live data.
* **Intelligent Row Validation:** Colors mandatory empty fields red dynamically on edit. Prevents incomplete rows from being pushed to the master sheet.
* **Auto-Internal Order Numbering:** Generates sequential, brand-specific internal order numbers based on dynamic prefixes, closing sequence gaps if rows are manually merged or edited.
* **1-Click Sync & Push:** Includes functions to safely batch-push validated "New Orders" to an "All Orders" master sheet, preventing exact duplicates.

## 📋 Prerequisites

This script requires a Google Spreadsheet with specifically named sheets (customizable in the code) and specific column headers.

### Required Sheets
1.  `New_Orders`: Where new, unprocessed orders arrive.
2.  `All Orders`: The master database of processed orders.
3.  `ProductCost`: Contains at least two columns: `SKU` and `Product Cost`.
4.  `Shipping_Charges`: Contains logic for shipping calculation, brand prefixes (`Initial code`), and numbering counters (`Last Maximum`).

## 🛠️ Setup & Installation

1. Create a new [Google Sheet](https://sheets.new/).
2. Go to **Extensions > Apps Script** in the top menu.
3. Delete any code in the script editor and paste the code from `Code.gs` found in this repository.
4. Save the project (Ctrl+S / Cmd+S).
5. (Optional but recommended) Set up an Installable Trigger for the `onEdit` function if your sheet will be edited by API integrations or service accounts. Otherwise, the simple `onEdit(e)` trigger will run automatically for human edits.

## ⚙️ Configuration

At the very top of the script, you will find a `CONFIG` object. Update this object to match your specific sheet names and preferred API endpoints.

```javascript
const CONFIG = {
  sheets: {
    newOrders: "New_Orders",       
    allOrders: "All Orders",       
    productCost: "ProductCost",   
    shipping: "Shipping_Charges"   
  },
  api: {
    exchangeRateUrl: "[https://open.er-api.com/v6/latest/INR](https://open.er-api.com/v6/latest/INR)" 
  },
  colors: {
    errorBackground: "#ff0000", 
    defaultBackground: "#ffffff" 
  }
};
