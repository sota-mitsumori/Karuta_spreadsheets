function onEdit(e) {
  updateStats();  // 変更があったら勝敗の統計を自動更新
}

function updateStats() {
  const outputSheetName = "Stats";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = ss.getSheetByName(outputSheetName);
  const sheets = ss.getSheets();

　// 勝敗の記号 (pull-down format is preferred to be used)
  const winSymbol = "○";
  const lossSymbol = "×";

  let nameStats = {};

  // Loop through all sheets (excluding "Stats")
  sheets.forEach(sheet => {
    if (sheet.getName() === outputSheetName) return;

    // データの取得先
    const data = sheet.getRange("A4:Z").getValues();

    data.forEach(row => {
      // A列にある名前を探索
      const name = row[0].toString().trim();
      if (!name) return;

      // initiate the array
      if (!nameStats[name]) {
        nameStats[name] = { wins: 0, losses: 0 };
      }

      // Check columns B to Z for win/loss symbols
      for (let i = 1; i < Math.min(row.length, 26); i++) {
        const cell = row[i].toString().trim();
        if (cell === winSymbol) {
          nameStats[name].wins += 1;
        } else if (cell === lossSymbol) {
          nameStats[name].losses += 1;
        }
      }
    });
  });

  // Convert nameStats to an array and calculate win rates
  let resultArray = [];
  for (let [name, stats] of Object.entries(nameStats)) {
    const totalMatches = stats.wins + stats.losses;
    const winRate = totalMatches > 0 ? (stats.wins / totalMatches) * 100 : 0;
    resultArray.push([name, stats.wins, stats.losses, winRate]);
  }

  // Sort the array by win rate in descending order, then by number of wins if win rates are equal
  resultArray.sort((a, b) => {
    if (b[3] === a[3]) {
      return b[1] - a[1];  // Sort by wins if win rates are equal
    }
    return b[3] - a[3];  // Sort by win rate
  });

  // Clear existing results in "Stats" sheet
  outputSheet.getRange("A3:D").clearContent();

  // Set headers
  outputSheet.getRange("A2:D2").setValues([["名前", "勝ち数", "負け数", "勝率"]]);

  // Write sorted results to the "Stats" sheet
  outputSheet.getRange(3, 1, resultArray.length, 4).setValues(
    resultArray.map(row => [row[0], row[1], row[2], row[3].toFixed(2) + "%"])
  );
}

