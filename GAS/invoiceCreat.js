function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('GAS実行')
  .addItem('請求書作成', 'addSheet')
  .addToUi();
}

function addSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('請求者情報');
  const lastRow = sheet.getLastRow();
  // const lastColumn = sheet.getLastColumn();
  const resValues = {
    name: sheet.getRange(2, 3, lastRow - 1).getValues(),
    email: sheet.getRange(2, 2, lastRow - 1).getValues(),
    tel: sheet.getRange(2, 5, lastRow - 1).getValues(),
    address: sheet.getRange(2, 6, lastRow - 1, 2).getValues(),
    company: sheet.getRange(2, 8, lastRow - 1, 2).getValues(),
    bank: sheet.getRange(2, 10, lastRow - 1, 5).getValues()
  };
  const copySheet = ss.getSheetByName('コピー元');
  const cpShtLastRow = copySheet.getLastRow();
  const cpShtLastColumn = copySheet.getLastColumn();
  const cpShtRng = copySheet.getRange(1, 1, cpShtLastRow, cpShtLastColumn)
  const cpData = cpShtRng.getValues();
  const names = resValues['name'].flat(Infinity);

  names.forEach((name,i) => {
    let newSheet = ss.insertSheet(name + '請求書');
    newSheet.getRange(1, 1, cpShtLastRow, cpShtLastColumn).setValues(cpData);
    let targetToCopy = newSheet.getRange(1, 1, cpShtLastRow, cpShtLastColumn);

    //PASTE_NORMAL値、数式、書式、結合
    cpShtRng.copyTo(targetToCopy, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    newSheet.getRange(8, 5).setValue(name);
    newSheet.getRange(12, 5).setValue(resValues['email'][i]);
    newSheet.getRange(11, 5).setValue(resValues['tel'][i]);
    newSheet.getRange(9, 5).setValue(resValues['address'][i][0]);
    newSheet.getRange(10, 5).setValue(resValues['address'][i][1]);
    newSheet.getRange(9, 5).setValue(resValues['address'][i][0]);
    newSheet.getRange(5, 5).setValue(resValues['company'][i][0]);
    newSheet.getRange(6, 5).setValue(resValues['company'][i][1]);
    newSheet.getRange(38, 1).setValue(resValues['bank'][i][0]);
    newSheet.getRange(38, 3).setValue(resValues['bank'][i][1]);
    newSheet.getRange(39, 2).setValue(resValues['bank'][i][2]);
    newSheet.getRange(41, 2).setValue(resValues['bank'][i][3]);
    newSheet.getRange(40, 1).setValue(resValues['bank'][i][4]);
  });

}
