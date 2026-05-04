---

📦 Scalar ERP v1.0 – Serverless MES on Google Sheets

"Spreadsheet? No. This is a state‑machine based inventory & settlement engine that runs on Google infrastructure."

Scalar ERP turns Google Sheets into a lightweight, state‑driven ERP engine.
It runs on Python (gspread) or pure Google Apps Script – no servers, no databases, no expensive licenses.
This is a production implementation of the Scalar Architecture applied to inventory, BOM explosion, and settlement workflows.

---

🚀 How It Works (For Normal Humans)

Concept What It Means
State Machine Purchase Order → BOM → Material need → Stock check → Shortage alert → Payment settlement
Serverless No setup. No hosting. Just a Google Sheet and 3 minutes.
Two Engines Python (for batch/automation) OR Apps Script (click‑to‑run inside Sheets)
Self‑Checking Built‑in validator automatically tells you if data is wrong

---

🧩 What's Inside

```
Scalar-ERP-Core/
├── service_account.json      # GCP key (you generate)
├── app.py                    # Auto‑creates sheets + headers
├── seed_data.py              # Loads BOM + sample orders
├── engine.py                 # Calculates shortages
├── verification_engine.py    # Checks if everything matches
├── update_and_settle.py      # Shipped → payment entries
├── setup_security.py         # Dropdowns + header protection
└── Scalar_ERP_Code.gs        # Full Apps Script (copy‑paste)
```

---

🔐 One‑Click Apps Script (Copy, Paste, Run)

Open your Google Sheet → Extensions → Apps Script → paste this:

```javascript
/**
 * Scalar ERP Kernel v2.0
 * Google Apps Script + MTP Scalar Architecture
 * License: MIT
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 SCALAR ERP')
      .addItem('📈 Run Inventory Engine', 'runInventoryEngine')
      .addItem('💰 Process Settlements', 'runSettlementProcess')
      .addSeparator()
      .addItem('🔄 Refresh Dashboard', 'refreshDashboard')
      .addToUi();
}

// Inventory Engine: Active POs → BOM → Shortage (with safety stock)
function runInventoryEngine() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orderSheet = ss.getSheetByName("TRX_발주");
  var bomSheet = ss.getSheetByName("MST_BOM");
  var invSheet = ss.getSheetByName("TRX_재고");
  
  if (!orderSheet || !bomSheet || !invSheet) {
    SpreadsheetApp.getUi().alert('❌ Missing sheets: TRX_발주, MST_BOM, TRX_재고');
    return;
  }

  var orders = orderSheet.getDataRange().getValues();
  var boms = bomSheet.getDataRange().getValues();
  var invData = invSheet.getDataRange().getValues();

  if (orders.length < 2) {
    SpreadsheetApp.getUi().alert('⚠️ No purchase order data found.');
    return;
  }
  if (boms.length < 2) {
    SpreadsheetApp.getUi().alert('⚠️ No BOM data found.');
    return;
  }

  // Auto-detect column headers (no hardcoded positions)
  var orderHeaders = orders[0];
  var bomHeaders = boms[0];
  var invHeaders = invData[0];
  
  var orderItemIdx = orderHeaders.indexOf('품목');
  var orderQtyIdx = orderHeaders.indexOf('수량');
  var orderStatusIdx = orderHeaders.indexOf('상태');
  
  var bomItemIdx = bomHeaders.indexOf('완성품');
  var bomMaterialIdx = bomHeaders.indexOf('자재명');
  var bomQtyIdx = bomHeaders.indexOf('소요량(Qty)');
  
  var invCodeIdx = invHeaders.indexOf('품목코드');
  var invStockIdx = invHeaders.indexOf('현재고');
  var invSafetyIdx = invHeaders.indexOf('안전재고');
  var invShortageIdx = invHeaders.indexOf('부족수량');
  var invUpdateIdx = invHeaders.indexOf('최종갱신');
  
  // Validate columns
  var missing = [];
  if (orderItemIdx === -1) missing.push('TRX_발주.품목');
  if (orderQtyIdx === -1) missing.push('TRX_발주.수량');
  if (orderStatusIdx === -1) missing.push('TRX_발주.상태');
  if (bomItemIdx === -1) missing.push('MST_BOM.완성품');
  if (bomMaterialIdx === -1) missing.push('MST_BOM.자재명');
  if (bomQtyIdx === -1) missing.push('MST_BOM.소요량(Qty)');
  if (invCodeIdx === -1) missing.push('TRX_재고.품목코드');
  if (invStockIdx === -1) missing.push('TRX_재고.현재고');
  if (invSafetyIdx === -1) missing.push('TRX_재고.안전재고');
  
  if (missing.length > 0) {
    SpreadsheetApp.getUi().alert('❌ Missing columns:\n' + missing.join('\n'));
    return;
  }

  // Step 1: Calculate material needs from '진행중' (in progress) orders
  var requirements = {};
  var processedOrders = 0;
  
  for (var i = 1; i < orders.length; i++) {
    var status = orders[i][orderStatusIdx];
    if (status === '진행중') {
      processedOrders++;
      var item = orders[i][orderItemIdx];
      var qty = Number(orders[i][orderQtyIdx]) || 0;
      if (qty <= 0) continue;
      
      for (var j = 1; j < boms.length; j++) {
        if (boms[j][bomItemIdx] === item) {
          var material = boms[j][bomMaterialIdx];
          var consumption = Number(boms[j][bomQtyIdx]) || 0;
          requirements[material] = (requirements[material] || 0) + (qty * consumption);
        }
      }
    }
  }
  
  if (processedOrders === 0) {
    SpreadsheetApp.getUi().alert('⚠️ No orders with status "진행중".\n\nSet at least one order to "진행중" and retry.');
    return;
  }

  // Step 2: Update inventory with shortage = max(0, safety - (stock - needed))
  var updatedCount = 0;
  var shortageMap = {};
  
  for (var k = 1; k < invData.length; k++) {
    var materialName = invData[k][invCodeIdx];
    var currentStock = Number(invData[k][invStockIdx]) || 0;
    var safetyStock = Number(invData[k][invSafetyIdx]) || 0;
    var needed = requirements[materialName] || 0;
    
    var shortage = Math.max(0, safetyStock - (currentStock - needed));
    
    if (shortage > 0) {
      shortageMap[materialName] = shortage;
    }
    
    if (invShortageIdx !== -1) {
      invSheet.getRange(k + 1, invShortageIdx + 1).setValue(shortage);
      updatedCount++;
    }
    if (invUpdateIdx !== -1) {
      invSheet.getRange(k + 1, invUpdateIdx + 1).setValue(new Date());
    }
  }
  
  // Step 3: Result message
  var message = '✅ Inventory calculation complete.\n';
  message += '   Processed orders: ' + processedOrders + '\n';
  message += '   Updated materials: ' + updatedCount + '\n\n';
  
  if (Object.keys(shortageMap).length > 0) {
    message += '⚠️ Shortage items:\n';
    for (var material in shortageMap) {
      message += '   • ' + material + ': ' + shortageMap[material] + '\n';
    }
  } else {
    message += '✅ All materials have sufficient stock.';
  }
  
  SpreadsheetApp.getUi().alert(message);
}

// Settlement: '발송완결' (shipped complete) → payment entries
function runSettlementProcess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orderSheet = ss.getSheetByName("TRX_발주");
  var settleSheet = ss.getSheetByName("TRX_결제마감");
  
  if (!orderSheet || !settleSheet) {
    SpreadsheetApp.getUi().alert('❌ Missing sheets: TRX_발주 or TRX_결제마감');
    return;
  }
  
  var orders = orderSheet.getDataRange().getValues();
  var orderHeaders = orders[0];
  
  var orderNoIdx = orderHeaders.indexOf('발주번호');
  var orderQtyIdx = orderHeaders.indexOf('수량');
  var orderStatusIdx = orderHeaders.indexOf('상태');
  var orderItemIdx = orderHeaders.indexOf('품목');
  
  if (orderNoIdx === -1 || orderStatusIdx === -1) {
    SpreadsheetApp.getUi().alert('❌ Missing columns: "발주번호" or "상태" in TRX_발주');
    return;
  }
  
  var newSettlements = [];
  var settleData = settleSheet.getDataRange().getValues();
  var existingOrders = new Set();
  
  for (var s = 1; s < settleData.length; s++) {
    if (settleData[s][1]) existingOrders.add(settleData[s][1].toString());
  }
  
  // Price mapping by product (customize here)
  var priceMap = {
    '워커': 85000,
    '구두': 120000,
    '운동화': 65000,
    '샌들': 45000
  };
  var defaultPrice = 50000;
  
  var completedCount = 0;
  var totalAmount = 0;
  
  for (var i = 1; i < orders.length; i++) {
    var status = orders[i][orderStatusIdx];
    if (status === '발송완결') {
      completedCount++;
      var orderNo = orders[i][orderNoIdx].toString();
      
      if (!existingOrders.has(orderNo)) {
        var qty = Number(orders[i][orderQtyIdx]) || 0;
        var item = orders[i][orderItemIdx] || '';
        var unitPrice = priceMap[item] || defaultPrice;
        var amount = qty * unitPrice;
        totalAmount += amount;
        
        newSettlements.push([
          "SHIP-" + orderNo,
          orderNo,
          amount,
          "결제대기",
          new Date()
        ]);
        existingOrders.add(orderNo);
      }
    }
  }
  
  if (newSettlements.length > 0) {
    if (settleData.length < 2 || settleData[0][0] !== '정산번호') {
      settleSheet.getRange(1, 1, 1, 5).setValues([['정산번호', '발주번호', '금액', '결제상태', '생성일']]);
    }
    var lastRow = Math.max(settleSheet.getLastRow(), 1);
    settleSheet.getRange(lastRow + 1, 1, newSettlements.length, 5).setValues(newSettlements);
    
    SpreadsheetApp.getUi().alert(
      '✅ Settlement entries created\n\n' +
      '• Entries: ' + newSettlements.length + '\n' +
      '• Total amount: ₩' + totalAmount.toLocaleString() + '\n' +
      '• Completed orders: ' + completedCount
    );
  } else {
    var msg = '💡 No new settlements to process.\n\n';
    msg += 'Condition: TRX_발주 status = "발송완결"\n';
    if (completedCount === 0) {
      msg += '\n※ No orders with "발송완결" status found.';
    } else {
      msg += '\n※ ' + completedCount + ' completed orders already settled.';
    }
    SpreadsheetApp.getUi().alert(msg);
  }
}

function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('🔄 Syncing dashboard...', 'Scalar ERP', 3);
  SpreadsheetApp.flush();
  ss.toast('✅ Dashboard refreshed', 'Scalar ERP', 2);
}
```

---

🧪 For Normal Users (No Code)

1. Create a new Google Sheet
2. Extensions → Apps Script
3. Paste the code above (replace everything)
4. Save and refresh the sheet
5. Click the 🚀 SCALAR ERP menu → Run Inventory Engine

That's it. No terminal. No Python. Just clicks.

---

🐍 For Automation Nerds (Python Version)

```bash
pip install gspread pandas google-auth
python engine.py
```

---

✅ Integrity Check

Run this to compare Python output vs. sheet data:

```bash
python verification_engine.py
```

Expected:

```
🟢 PASSED – All data integrity constraints satisfied.
```

---

📜 License

MIT – Use it, break it, fix it, ship it.

---

✍️ Author's Note

"What you're looking at looks like an Excel file. But it's not Excel.
That's the point – minimal surface, maximum compression.
No expensive licenses. No complicated setup. Just a sheet that thinks it's an ERP."

---

Built with ☕ and state machines.
Questions? Open an issue. Improvements? Send a PR.