function updateSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("Summary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Summary");
  }
  
  // ── 大会シート一覧を取得 ──
  var eventSheets = ss.getSheets().filter(function(sh){
    return sh.getName() !== summarySheet.getName();
  });
  
  // ── 各シートのデータを取得 ──
  var sheetDataList = eventSheets.map(function(sh){
    var lastRow = sh.getLastRow();
    var data = lastRow >= 2
      ? sh.getRange(2, 1, lastRow - 1, 3).getValues()
      : [];
    return { name: sh.getName(), data: data };
  });
  
  // ── 全シートから名前をユニーク抽出 ──
  var nameSet = {};
  sheetDataList.forEach(function(sheetObj){
    sheetObj.data.forEach(function(row){
      if (row[0]) nameSet[row[0]] = true;
    });
  });
  var allNames = Object.keys(nameSet);
  
  // ── Summary シートの A 列から既存順を取得 ──
  var lastA = summarySheet.getLastRow();
  var existingA = lastA >= 1
    ? summarySheet.getRange(1, 1, lastA, 1).getValues().flat().filter(function(n){ return n; })
    : [];
  
  // ── A列にある名前順 ──
  var orderedNames = existingA.filter(function(n){
    return allNames.indexOf(n) !== -1;
  });
  // ── A列にない名前は末尾にソートして追加 ──
  var newNames = allNames
    .filter(function(n){ return existingA.indexOf(n) === -1; })
    .sort();
  var names = orderedNames.concat(newNames);
  
  // ── Summary 用データを組み立て ──
  var summaryValues = names.map(function(name){
    var totalUnpaid = 0;
    var unpaidEvents = [];
    sheetDataList.forEach(function(sheetObj){
      for (var i = 0; i < sheetObj.data.length; i++) {
        var row = sheetObj.data[i];
        if (row[0] === name) {
          var amount = parseFloat(row[1]) || 0;
          if (!row[2]) {
            totalUnpaid += amount;
            unpaidEvents.push(sheetObj.name);
          }
          break;
        }
      }
    });
    return [ name, totalUnpaid ].concat(unpaidEvents);
  });
  
  if (summaryValues.length === 0) {
    return;
  }
  
  // ── 最大列数に合わせてパディング ──
  var maxCols = summaryValues.reduce(function(max, row){
    return Math.max(max, row.length);
  }, 0);
  summaryValues.forEach(function(row){
    while (row.length < maxCols) row.push("");
  });
  
  // ── B列以降をクリア（A列は untouched） ──
  var lastRow = Math.max(summarySheet.getMaxRows(), summaryValues.length);
  summarySheet.getRange(1, 2, lastRow, maxCols).clearContent();
  
  // ── B列から書き込み ──
  summarySheet
    .getRange(1, 2, summaryValues.length, maxCols)
    .setValues(summaryValues);
}

