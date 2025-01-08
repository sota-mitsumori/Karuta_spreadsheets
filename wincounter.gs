function onEdit(e) {
  // 編集イベントの情報を取得
  const range = e.range;
  const sheet = range.getSheet();
  const editedCell = range.getA1Notation();
  const sheetName = sheet.getName();

  // 結果を出力するシート名を指定
  const outputSheetName = "Stats"; // シート名が「Stats」であることを確認してください

  // 編集されたシートが「Stats」であり、編集されたセルがA3であるかを確認
  if (sheetName === outputSheetName && editedCell === "A3") {
    // セルA3から対象の名前を取得
    const targetName = sheet.getRange("A3").getValue().toString().trim();

    // 名前が空の場合、結果をクリアして終了
    if (targetName === "") {
      clearResults(sheet);
      return;
    }

    // 勝ちと負けを示すシンボル
    const winSymbol = "○";
    const lossSymbol = "×";

    // スプレッドシート全体を取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();

    let totalWins = 0;
    let totalLosses = 0;

    // すべてのシートをループ（「Stats」シートを除く）
    sheets.forEach((sht) => {
      if (sht.getName() === outputSheetName) {
        return; // 「Stats」シートはスキップ
      }

      const data = sht.getDataRange().getValues();

      data.forEach(row => {
        // 行の最初のセルが対象の名前か確認
        if (row[0].toString().trim() === targetName) {
          // 列B～J（インデックス1～9）をチェック
          for (let i = 1; i < Math.min(row.length, 20); i++) {
            const cell = row[i].toString().trim();
            if (cell === winSymbol) {
              totalWins += 1;
            } else if (cell === lossSymbol) {
              totalLosses += 1;
            }
          }
        }
      });
    });

    // 勝率の計算
    const totalMatches = totalWins + totalLosses;
    const winRate = totalMatches > 0 ? (totalWins / totalMatches) * 100 : 0;

    // 結果を「Stats」シートに出力
    outputResults(sheet, totalWins, totalLosses, winRate, targetName);
  }
}

// 結果を「Stats」シートに出力する関数
function outputResults(sheet, wins, losses, rate, name) {
  // ヘッダーの設定
  sheet.getRange("A2").setValue("名前");
  sheet.getRange("B2").setValue("勝ち数");
  sheet.getRange("C2").setValue("負け数");
  sheet.getRange("D2").setValue("勝率");

  // 既存の結果をクリア（2行目以降）
  sheet.getRange("A3:D").clearContent();

  // 結果の入力
  sheet.getRange("A3").setValue(name);
  sheet.getRange("B3").setValue(wins);
  sheet.getRange("C3").setValue(losses);
  sheet.getRange("D3").setValue(rate.toFixed(2) + "%");
}

// 結果をクリアする関数
function clearResults(sheet) {
  // ヘッダーの設定（保持）
  sheet.getRange("A2:D2").setValues([["名前", "勝ち数", "負け数", "勝率"]]);

  // 既存の結果をクリア（2行目以降）
  sheet.getRange("A3:D").clearContent();
}
