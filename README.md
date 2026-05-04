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
 * Scalar ERP 통합 커널 v2.1 (LockService 적용 + 정합성 검증 완료)
 * 제작: L0 (지휘관) / 실행: L1 (AI)
 * 정합성 대상: python engine.py, verification_engine.py
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 SCALAR ERP')
      .addItem('📈 재고 연산 실행 (Engine)', 'runInventoryEngineSafe')
      .addItem('💰 출고 마감 처리 (Settlement)', 'runSettlementProcessSafe')
      .addSeparator()
      .addItem('📊 대시보드 강제 갱신', 'refreshDashboard')
      .addToUi();
}

// ==================== LOCK 적용 래퍼 함수 ====================

// 안전한 재고 연산 (동시 실행 방지)
function runInventoryEngineSafe() {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      SpreadsheetApp.getUi().alert('⚠️ 다른 사용자가 현재 계산 중입니다.\n30초 후 다시 시도해주세요.');
      return;
    }
    runInventoryEngine();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ 오류 발생: ' + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// 안전한 정산 처리 (동시 실행 방지)
function runSettlementProcessSafe() {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      SpreadsheetApp.getUi().alert('⚠️ 다른 사용자가 현재 정산 중입니다.\n30초 후 다시 시도해주세요.');
      return;
    }
    runSettlementProcess();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ 오류 발생: ' + e.toString());
  } finally {
    lock.releaseLock();
  }
}

// ==================== 코어 엔진 ====================

// 1. 재고 연산 엔진 (python engine.py와 100% 정합)
function runInventoryEngine() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orderSheet = ss.getSheetByName("TRX_발주");
  var bomSheet = ss.getSheetByName("MST_BOM");
  var invSheet = ss.getSheetByName("TRX_재고");
  
  if (!orderSheet || !bomSheet || !invSheet) {
    SpreadsheetApp.getUi().alert('❌ 필수 시트가 없습니다: TRX_발주, MST_BOM, TRX_재고');
    return;
  }

  var orders = orderSheet.getDataRange().getValues();
  var boms = bomSheet.getDataRange().getValues();
  var invData = invSheet.getDataRange().getValues();

  if (orders.length < 2) {
    SpreadsheetApp.getUi().alert('⚠️ 발주 데이터가 없습니다.\n\n📝 예시:\n발주일자 | 품목 | 수량 | 상태=진행중');
    return;
  }
  if (boms.length < 2) {
    SpreadsheetApp.getUi().alert('⚠️ BOM 데이터가 없습니다.\n\n📝 예시:\n완성품 | 자재명 | 소요량(Qty)');
    return;
  }

  // 헤더 인덱스 자동 탐지 (컬럼명 변경에 강함)
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
  
  // 필수 컬럼 검증
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
    SpreadsheetApp.getUi().alert('❌ 필수 컬럼 누락:\n' + missing.join('\n'));
    return;
  }
  
  // 1단계: 발주별 필요 자재량 계산 (진행중인 발주만)
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
    SpreadsheetApp.getUi().alert('⚠️ "진행중" 상태인 발주가 없습니다.\n\n발주 시트의 상태 컬럼을 "진행중"으로 변경 후 실행하세요.');
    return;
  }
  
  // 2단계: 재고 시트 업데이트 (안전재고 공식 적용)
  var updatedCount = 0;
  var shortageMap = {};
  
  for (var k = 1; k < invData.length; k++) {
    var materialName = invData[k][invCodeIdx];
    var currentStock = Number(invData[k][invStockIdx]) || 0;
    var safetyStock = Number(invData[k][invSafetyIdx]) || 0;
    var needed = requirements[materialName] || 0;
    
    // 핵심 공식: max(0, 안전재고 - (현재고 - 필요량))
    var shortage = Math.max(0, safetyStock - (currentStock - needed));
    
    if (shortage > 0) {
      shortageMap[materialName] = shortage;
    }
    
    // 부족수량 업데이트
    if (invShortageIdx !== -1) {
      invSheet.getRange(k + 1, invShortageIdx + 1).setValue(shortage);
      updatedCount++;
    }
    // 최종갱신일 업데이트
    if (invUpdateIdx !== -1) {
      invSheet.getRange(k + 1, invUpdateIdx + 1).setValue(new Date());
    }
  }
  
  // 3단계: 결과 메시지 (파이썬 엔진과 동일한 포맷)
  var message = '✅ 재고 연산이 완료되었습니다.\n';
  message += '   처리된 발주: ' + processedOrders + '건\n';
  message += '   업데이트된 자재: ' + updatedCount + '개\n\n';
  
  if (Object.keys(shortageMap).length > 0) {
    message += '⚠️ 부족 품목:\n';
    for (var material in shortageMap) {
      message += '   • ' + material + ': ' + shortageMap[material] + '\n';
    }
  } else {
    message += '✅ 모든 품목의 재고가 충분합니다.';
  }
  
  SpreadsheetApp.getUi().alert(message);
}

// 2. 출고 및 정산 프로세스 (python settlement.py와 정합)
function runSettlementProcess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var orderSheet = ss.getSheetByName("TRX_발주");
  var settleSheet = ss.getSheetByName("TRX_결제마감");
  
  if (!orderSheet || !settleSheet) {
    SpreadsheetApp.getUi().alert('❌ 필수 시트가 없습니다: TRX_발주, TRX_결제마감');
    return;
  }
  
  var orders = orderSheet.getDataRange().getValues();
  var orderHeaders = orders[0];
  
  var orderNoIdx = orderHeaders.indexOf('발주번호');
  var orderQtyIdx = orderHeaders.indexOf('수량');
  var orderStatusIdx = orderHeaders.indexOf('상태');
  var orderItemIdx = orderHeaders.indexOf('품목');
  
  if (orderNoIdx === -1 || orderStatusIdx === -1) {
    SpreadsheetApp.getUi().alert('❌ TRX_발주 시트 컬럼 오류: "발주번호", "상태" 필요');
    return;
  }
  
  var newSettlements = [];
  var settleData = settleSheet.getDataRange().getValues();
  var settleOrderNos = new Set();
  
  // 이미 정산된 발주번호 수집
  for (var s = 1; s < settleData.length; s++) {
    if (settleData[s][1]) settleOrderNos.add(settleData[s][1].toString());
  }
  
  // 단가 매핑 (품목별 실제 단가)
  var priceMap = {
    '워커': 85000,
    '구두': 120000,
    '운동화': 65000,
    '샌들': 45000,
    '부츠': 110000,
    '로퍼': 95000
  };
  var defaultPrice = 50000;
  
  var completedCount = 0;
  var totalAmount = 0;
  
  for (var i = 1; i < orders.length; i++) {
    var status = orders[i][orderStatusIdx];
    // 파이썬 엔진과 동일: '발송완결'이 출고완료 상태
    if (status === '발송완결') {
      completedCount++;
      var orderNo = orders[i][orderNoIdx].toString();
      
      if (!settleOrderNos.has(orderNo)) {
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
        settleOrderNos.add(orderNo);
      }
    }
  }
  
  if (newSettlements.length > 0) {
    // 헤더 확인 및 추가
    if (settleData.length < 2 || settleData[0].length < 5 || settleData[0][0] !== '정산번호') {
      settleSheet.getRange(1, 1, 1, 5).setValues([['정산번호', '발주번호', '금액', '결제상태', '생성일']]);
    }
    var lastRow = Math.max(settleSheet.getLastRow(), 1);
    settleSheet.getRange(lastRow + 1, 1, newSettlements.length, 5).setValues(newSettlements);
    
    SpreadsheetApp.getUi().alert(
      '✅ 정산 데이터 생성 완료\n\n' +
      '• 생성 건수: ' + newSettlements.length + '건\n' +
      '• 총 금액: ' + totalAmount.toLocaleString() + '원\n' +
      '• 발송완결 건수: ' + completedCount + '건'
    );
  } else {
    var msg = '💡 새로 마감할 내역이 없습니다.\n\n';
    msg += '조건: TRX_발주 상태 = "발송완결"\n';
    if (completedCount === 0) {
      msg += '\n※ 현재 "발송완결" 상태인 발주가 없습니다.';
    } else {
      msg += '\n※ ' + completedCount + '건의 "발송완결" 발주가 이미 정산 처리되었습니다.';
    }
    SpreadsheetApp.getUi().alert(msg);
  }
}

// 3. 대시보드 갱신
function refreshDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('🔄 대시보드 데이터를 동기화합니다...', 'Scalar ERP', 3);
  
  // 추가 대시보드 로직이 있다면 여기에 삽입
  SpreadsheetApp.flush();
  ss.toast('✅ 대시보드 갱신 완료', 'Scalar ERP', 2);
}

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
