function onEdit(e) {
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();
  const nameSheetName = 'Summary';  // 名前リストを置いているシート名
  if (sheet.getName() === nameSheetName || range.getColumn() !== 1) return;

  const ss = e.source;
  const nameSheet = ss.getSheetByName(nameSheetName);
  if (!nameSheet) return;

  // 名前リストを一度だけ取得
  const allNames = nameSheet.getRange('A:A')
                    .getValues()
                    .flat()
                    .filter(n => n);

  // 編集された範囲の値（2D配列）を取得
  const values = range.getValues();

  // 自動補完後の値を格納する2D配列を作成
  const newValues = values.map(row => {
    const input = row[0];
    if (!input) return [''];
    const matches = allNames.filter(n => n.startsWith(input));
    // 一意に1件だけ & まだ補完されていなければ置き換え
    if (matches.length === 1 && matches[0] !== input) {
      return [matches[0]];
    }
    return [input];
  });

  // 値が変わっていたらまとめて上書き
  range.setValues(newValues);
}

