/**
 * このスクリプトは、スプレッドシートが編集されたときに自動的に実行されます。
 * 主に5つの機能を持っています。
 * 1. 対戦相手の略称が入力された際、相手の対戦相手欄にも自動で入力する。
 * 2. 勝敗(○×)が入力された際、相手の勝敗欄にも自動で逆の結果を入力する。
 * 3. 枚数差が入力された際、相手の枚数差欄にも自動で同じ数値を入力する。
 * 4. 特定のセルが変更された際に、色付きセルの数から参加者をカウントする。
 * 5. 編集があるたびに「Stats」シートの勝敗記録を更新する。
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = range.getValue().toString().trim();

  const sheetName = sheet.getName();
  // 「Stats」や「name_correspondance」シートでは実行しない
  if (sheetName === "Stats" || sheetName === "name_correspondance") {
    return;
  }

  // --- データ行（4行目以降）での自動入力ロジック ---
  if (row >= 4) {
    // Case 1: 対戦相手の列（B, E, H...）が編集された場合
    if (col > 1 && (col - 2) % 3 === 0) {
      try {
        autoFillOpponent(sheet, row, col, value);
      } catch (err) {
        SpreadsheetApp.getUi().alert(`対戦相手の自動入力中にエラーが発生しました: ${err.toString()}`);
      }
    }
    // Case 2: 勝敗結果の列（C, F, I...）が編集された場合
    else if (col > 2 && col % 3 === 0) {
      try {
        autoFillResult(sheet, row, col, value);
      } catch (err) {
        SpreadsheetApp.getUi().alert(`勝敗の自動入力中にエラーが発生しました: ${err.toString()}`);
      }
    }
    // Case 3: 枚数差の列 (D, G, J...) が編集された場合
    else if (col > 3 && (col - 1) % 3 === 0) {
      try {
        autoFillScoreDifference(sheet, row, col, value);
      } catch (err) {
        SpreadsheetApp.getUi().alert(`枚数差の自動入力中にエラーが発生しました: ${err.toString()}`);
      }
    }
  }

  // --- 参加者数のカウント機能 ---
  if (value.toLowerCase() === 'on') {
    sheet.getRange(3, 1).setValue('off'); 
    const color = '#cccccc';
    for (var i = 0; i < 8; i++) {
      var columnletter = String.fromCharCode(64 + 3 * i + 3);
      var backgroundColors = sheet.getRange(columnletter + ":" + columnletter).getBackgrounds();
      var count = 0;
      for (var j = 0; j < backgroundColors.length; j++) {
        if (backgroundColors[j][0] === color) {
          count++;
        }
      }
      sheet.getRange(3, 3 * i + 3).setValue(count);
    }
  }

  // --- Statsシートの更新 ---
  // スクリプトによる自動編集でない場合のみStatsを更新し、無限ループを防ぐ
  if (e.authMode !== ScriptApp.AuthMode.NONE) {
    updateStats();
  }
}

/**
 * 対戦相手の略称が入力された際、相手の対戦相手欄にも自動でこちらの略称を入力する
 */
function autoFillOpponent(sheet, playerRow, opponentCol, opponentNickname) {
  const { nicknameToFull, fullToNickname } = getNicknameMaps();
  if (!nicknameToFull || !fullToNickname) return;

  const playerFullName = sheet.getRange(playerRow, 1).getValue().toString().trim();
  const playerNickname = fullToNickname[playerFullName];

  if (!playerNickname) {
    console.log(`プレイヤー「${playerFullName}」の略称が見つかりませんでした。`);
    return;
  }

  if (!opponentNickname) {
    sheet.getRange(playerRow, opponentCol).setBackground(null);
    return;
  }

  const opponentFullName = nicknameToFull[opponentNickname];
  if (!opponentFullName) {
    sheet.getRange(playerRow, opponentCol).setBackground('#f4cccc');
    return;
  }
  
  sheet.getRange(playerRow, opponentCol).setBackground(null);

  const opponentRowInSheet = findPlayerRowByFullName(sheet, opponentFullName);
  if (opponentRowInSheet === -1) {
    console.log(`対戦相手「${opponentFullName}」がこのシートのA列に見つかりませんでした。`);
    return;
  }

  const opponentCell = sheet.getRange(opponentRowInSheet, opponentCol);
  if (opponentCell.getValue().toString().trim() !== playerNickname) {
    opponentCell.setValue(playerNickname);
  }
}

/**
 * 勝敗が入力された際、相手の勝敗欄に逆の結果を自動入力する
 */
function autoFillResult(sheet, playerRow, resultCol, resultValue) {
  const winSymbol = "○";
  const lossSymbol = "×";

  let oppositeResult;
  if (resultValue === winSymbol) {
    oppositeResult = lossSymbol;
  } else if (resultValue === lossSymbol) {
    oppositeResult = winSymbol;
  } else if (resultValue === "") {
    oppositeResult = "";
  } else {
    return;
  }

  const opponentNickname = sheet.getRange(playerRow, resultCol - 1).getValue().toString().trim();
  if (!opponentNickname) return;

  const { nicknameToFull } = getNicknameMaps();
  if (!nicknameToFull) return;
  const opponentFullName = nicknameToFull[opponentNickname];
  if (!opponentFullName) return;
  
  const opponentRowInSheet = findPlayerRowByFullName(sheet, opponentFullName);
  if (opponentRowInSheet === -1) return;

  const opponentResultCell = sheet.getRange(opponentRowInSheet, resultCol);
  if (opponentResultCell.getValue().toString().trim() !== oppositeResult) {
    opponentResultCell.setValue(oppositeResult);
  }
}

/**
 * 枚数差が入力された際、相手の枚数差欄に同じ数値を自動入力する
 */
function autoFillScoreDifference(sheet, playerRow, scoreCol, scoreValue) {
  // 1. 2つとなりのセルから対戦相手の略称を取得
  const opponentNickname = sheet.getRange(playerRow, scoreCol - 2).getValue().toString().trim();
  if (!opponentNickname) return;

  // 2. 略称から対戦相手のフルネームと行を探す
  const { nicknameToFull } = getNicknameMaps();
  if (!nicknameToFull) return;
  const opponentFullName = nicknameToFull[opponentNickname];
  if (!opponentFullName) return;
  
  const opponentRowInSheet = findPlayerRowByFullName(sheet, opponentFullName);
  if (opponentRowInSheet === -1) return;

  // 3. 対戦相手の枚数差セルを更新
  const opponentScoreCell = sheet.getRange(opponentRowInSheet, scoreCol);
  if (opponentScoreCell.getValue().toString().trim() !== scoreValue.toString().trim()) {
    opponentScoreCell.setValue(scoreValue);
  }
}


/**
 * フルネームを基に、シートのA列からプレイヤーの行番号を検索するヘルパー関数
 */
function findPlayerRowByFullName(sheet, fullName) {
  if (!fullName) return -1;
  const namesRange = sheet.getRange("A4:A" + sheet.getLastRow());
  const namesList = namesRange.getValues();
  for (let i = 0; i < namesList.length; i++) {
    if (namesList[i][0].toString().trim() === fullName) {
      return i + 4;
    }
  }
  return -1;
}

/**
 * 「name_correspondance」シートからデータを読み込み、対応表オブジェクトを作成する
 */
function getNicknameMaps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const correspondenceSheet = ss.getSheetByName("name_correspondance");
  if (!correspondenceSheet) {
    console.error("シート「name_correspondance」が見つかりません。");
    return {};
  }
  
  const lastRow = correspondenceSheet.getLastRow();
  if (lastRow < 2) return {};
  const data = correspondenceSheet.getRange("A2:B" + lastRow).getValues();
  
  const nicknameToFull = {};
  const fullToNickname = {};
  
  data.forEach(row => {
    const fullName = row[0].toString().trim();
    const nickname = row[1].toString().trim();
    if (fullName && nickname) {
      nicknameToFull[nickname] = fullName;
      fullToNickname[fullName] = nickname;
    }
  });

  return { nicknameToFull, fullToNickname };
}

/**
 * 全シートを横断して勝敗を集計し、「Stats」シートを更新する
 */
function updateStats() {
  const outputSheetName = "Stats";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let outputSheet = ss.getSheetByName(outputSheetName);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(outputSheetName);
  }
  
  const sheets = ss.getSheets();
  const winSymbol = "○";
  const lossSymbol = "×";
  let nameStats = {};

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    if (sheetName === outputSheetName || sheetName === "name_correspondance") return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) return;
    const data = sheet.getRange("A4:Z" + lastRow).getValues();

    data.forEach(row => {
      const name = row[0].toString().trim();
      if (!name) return;

      if (!nameStats[name]) {
        nameStats[name] = { wins: 0, losses: 0 };
      }

      for (let i = 2; i < Math.min(row.length, 26); i += 3) {
        const cell = row[i].toString().trim();
        if (cell === winSymbol) {
          nameStats[name].wins += 1;
        } else if (cell === lossSymbol) {
          nameStats[name].losses += 1;
        }
      }
    });
  });

  let resultArray = [];
  for (let [name, stats] of Object.entries(nameStats)) {
    const totalMatches = stats.wins + stats.losses;
    const winRate = totalMatches > 0 ? (stats.wins / totalMatches) * 100 : 0;
    resultArray.push([name, stats.wins, stats.losses, winRate]);
  }

  resultArray.sort((a, b) => {
    if (b[3] === a[3]) {
      return b[1] - a[1];
    }
    return b[3] - a[3];
  });

  if (outputSheet.getLastRow() > 1) {
    outputSheet.getRange("A2:D" + outputSheet.getLastRow()).clearContent();
  }

  outputSheet.getRange("A2:D2").setValues([["名前", "勝ち数", "負け数", "勝率"]]);

  if (resultArray.length > 0) {
    outputSheet.getRange(3, 1, resultArray.length, 4).setValues(
      resultArray.map(row => [row[0], row[1], row[2], row[3].toFixed(2) + "%"])
    );
  }
}

