// Google Apps Script (コード.gs)
let submitSheetName = '提出データ';
let staffIdList = '社員番号一覧';
let operationSheet = '操作シート';



function confirmAndCreate() {
  // 1. 確認ポップアップを表示
  // 第1引数: タイトル, 第2引数: メッセージ, 第3引数: ボタンの種類
  const response = Browser.msgBox("実行確認", "データベースを初期化して新しく作成しますか？", Browser.Buttons.YES_NO);

  // 2. 「はい (yes)」が押された時だけ実行
  if (response === "yes") {
    CreateNewDatabase(); // 本番の関数を呼び出す
    Browser.msgBox("完了", "新しいデータベースを作成しました。", Browser.Buttons.OK);
  } else {
    // 「いいえ」または「×」で閉じた場合
    Logger.log("実行がキャンセルされました");
  }
}

function CreateNewDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("提出DB原本");
  const targetSheet = ss.getSheetByName(submitSheetName);


  // コピー元の値を二次元配列で取得
  const values = sourceSheet.getDataRange().getValues();

  // コピー先を一度クリア（値のみ）
  targetSheet.clearContents();

  // コピー先のA1セルを起点に、取得した配列と同じサイズで貼り付け
  targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);

  const period = ss.getSheetByName(operationSheet).getRange(2,8).getValue();

  targetSheet.getRange(2,1,1,2).merge();
  targetSheet.getRange(2,1).setValue(period);


  const startDate = new Date(ss.getSheetByName(operationSheet).getRange(2,2).getValue());
  const endDate = new Date(ss.getSheetByName(operationSheet).getRange(2,4).getValue());
  const NumofDays = getDayDiff(startDate,endDate);
  var date = new Date(startDate.getTime());
  var datename;
  for(var i = 3;i <= NumofDays*2+1;i+=2){
    datename = "'" + Utilities.formatDate(date,"JST","M/d");
    targetSheet.getRange(2,i).setValue(datename);
    date.setDate(date.getDate() + 1);
  }
}

function getDayDiff(startDate,endDate) {
  // 2. ミリ秒の差分を計算
  const diffTime = endDate.getTime() - startDate.getTime();

  // 3. ミリ秒を「日」に変換 (1日 = 1000ms * 60s * 60m * 24h)
  // +1 するのは「当日を含める（1日〜3日なら3日間）」ため。不要なら消してください。
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1;

  console.log(diffDays); // 数値で出力される
  return diffDays;
}


function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index_1');
  
  // 「操作シート」から日付を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const opSheet = ss.getSheetByName(operationSheet); // シート名: 操作シート
  
  // B2 (開始日) と D2 (終了日) を取得
  // getDisplayValue() を使うことで、スプシ上の見た目通りの文字列（YYYY-MM-DD等）で取得できます
  template.startDate = opSheet.getRange(2,2).getDisplayValue();
  template.endDate = opSheet.getRange(2,4).getDisplayValue();
  
  var output = template.evaluate();
  output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}


/**
 * 社員番号から氏名を検索する
 */
function getStaffName(staffId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(submitSheetName); // '社員番号一覧'
  const data = sheet.getDataRange().getValues();
  
  // 2行目から最終行までループ（ヘッダー飛ばし）
  for (let i = 2; i < data.length; i++) {
    if (String(data[i][1]) === String(staffId)) { // B列が社員番号
      return data[i][0]; // C列の氏名を返す
    }
  }
  return "登録なし";
}


/**
 * HTMLフォームから送られたデータを受け取り、スプレッドシートに書き込みます。
 * @param {Array<string>} data フォームから送られたデータ（最初の要素が社員番号）
 * @returns {string} 成功メッセージ
 */
function processSelection(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(submitSheetName);
    const staffid = data[0]; // 社員番号
    const dataToWrite = [data]; // 書き込むデータ（二次元配列）
    
    // 現在のシートの全データ (社員番号列, B列のみを読み込むのがベストだが、ここでは簡単のため全データ)
    // 実際に運用する場合は、最終行から読み取るべき範囲を限定することで高速化できます。
    const lastRow = sheet.getLastRow();
    
    // B列 (列番号2) のデータのみを取得。ヘッダー行を考慮して2行目から最終行まで
    let targetRow = 3; // データが始まる行
    let rowToOverwrite = -1; // 上書き対象の行

    if (lastRow >= targetRow) {
      // B列のデータを一括で取得
      const staffIdsInSheet = sheet.getRange(targetRow, 2, lastRow - targetRow + 1, 1).getValues(); 
      
      // 取得した配列の中から社員番号と一致する行を検索
      for (let i = 0; i < staffIdsInSheet.length; i++) {
        // staffIdsInSheet[i][0] は B列の値
        if (String(staffIdsInSheet[i][0]) === String(staffid)) {
          // 見つかった場合、その行番号を計算 (i は0から始まるインデックス)
          rowToOverwrite = targetRow + i; 
          break; // 見つかったのでループを終了
        }
      }
    }

    let writeRow;
    let action;

    if (rowToOverwrite !== -1) {
      // 社員番号が見つかった場合、その行に上書き
      writeRow = rowToOverwrite;
      action = "上書き";
    } else {
      // 見つからなかった場合、最終行の次に追記
      writeRow = lastRow + 1;
      action = "新規作成";
    }

    sheet.getRange(writeRow, 2, 1, data.length).setValues(dataToWrite);
    
    return `データが正常にSSに${action}されました。（${writeRow}行目）`;

  } catch (error) {
    Logger.log(`データの書き込み中にエラー: ${error.message}`);
    throw new Error(`データの書き込みに失敗しました: ${error.message}`);
  }
}




