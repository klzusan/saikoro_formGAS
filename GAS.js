const ss = SpreadsheetApp.getActiveSpreadsheet();
const s_imp = ss.getSheetByName('import_shopForm').getDataRange().getValues();
const csv = ss.getSheetByName('csv');
const s_csv = csv.getDataRange().getValues();
let goods_csvList = [];

function main() {
  // by ChatGPT
  const startRow = csv.getLastRow();
  const numRows = goods_csvList.length;
  const numCols = s_csv[0].length;
  //
  
  for (let i = 1; i < s_imp.length; i++) {
    proc_row(s_imp[i]);
  }

  // by ChatGPT
  let filledList = goods_csvList.map(row => {
    let newRow = new Array(numCols).fill("");
    for (let i = 0; i < row.length; i++) {
      if (row[i] !== undefined) newRow[i] = row[i];
    }
    return newRow;
  });

  csv.getRange(startRow + 1, 1, filledList.length, numCols).setValues(filledList);
  //
}

// この中で一つの商品を処理
function proc_row(item) {
  const col_kindName = getNumOfColumn(s_imp, '種類名を「/」区切りで入力してください');
  const col_kindQuan = getNumOfColumn(s_imp, '在庫数を入力してください');
  const col_image = getNumOfColumn(s_imp, '商品画像をアップロードしてください');

  const imagesName = getImagesName(item, col_image);

  if (isMultiple(item, col_kindName)) {
    // 種類分けあり
    const singles = item[col_kindName].split("/");
    const quantities = item[col_kindQuan].split("/");

    for (let i = 0; i < singles.length; i++) {
      // 商品の種類ごとに商品情報リストを作成
      let list = [];
      list[getNumOfColumn(s_csv, '商品名')] = item[getNumOfColumn(s_imp, '商品名を入力してください')];
      list[getNumOfColumn(s_csv, '種類名')] = singles[i];
      list[getNumOfColumn(s_csv, '説明')] = item[getNumOfColumn(s_imp, '商品概要を記述してください')];
      list[getNumOfColumn(s_csv, '価格')] = item[getNumOfColumn(s_imp, '希望価格')];
      list[getNumOfColumn(s_csv, '税率')] = 1;
      list[getNumOfColumn(s_csv, '公開状態')] = 0;
      list[getNumOfColumn(s_csv, '種類在庫数')] = +quantities[i];
      
      for (let j = 0; j < imagesName.length; j++) {
        list[getNumOfColumn(s_csv, `画像${j+1}`)] = imagesName[j];
      }

      // list[getNumOfColumn(s_csv, '')] = item[getNumOfColumn(s_imp, '')];

      goods_csvList.push(list);
    }
  } else {
    // 種類分けなし
    let list = [];
    list[getNumOfColumn(s_csv, '商品名')] = item[getNumOfColumn(s_imp, '商品名を入力してください')];
    list[getNumOfColumn(s_csv, '説明')] = item[getNumOfColumn(s_imp, '商品概要を記述してください')];
    list[getNumOfColumn(s_csv, '価格')] = item[getNumOfColumn(s_imp, '希望価格')];
    list[getNumOfColumn(s_csv, '税率')] = 1;
    list[getNumOfColumn(s_csv, '公開状態')] = 0;
    list[getNumOfColumn(s_csv, '在庫数')] = item[getNumOfColumn(s_imp, '在庫数')];

    for (let j = 0; j < imagesName.length; j++) {
      list[getNumOfColumn(s_csv, `画像${j+1}`)] = imagesName[j];
    }

    goods_csvList.push(list);
  }
}

function getNumOfColumn(sheet, include_str) {
  var h = sheet[0];
  for (let i = 0; i < h.length; i++) {
    if (typeof h[i] == "string" && h[i].includes(include_str)) {
      return i;
    }
  }

  console.log('String not found.');
  return -1;
}

function getImagesName(item, col_image) {
  const urls = item[col_image].split(", ");
  let fileName = [];
  
  for (let i = 0; i < urls.length; i++) {
    let url = urls[i];

    let fileId = url.split("id=")[1];

    let file = DriveApp.getFileById(fileId);
    fileName[i] = file.getName();
  }

  return fileName;
}

function isMultiple(item, col_kind) {
  if(!item[col_kind]) {
    return false;
  }
  return true;
}
