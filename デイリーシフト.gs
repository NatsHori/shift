let daily_spredsheet = SpreadsheetApp.getActiveSpreadsheet();
let sousa_sheet = daily_spredsheet.getSheetByName('操作用');
let genpontest_sheet = daily_spredsheet.getSheetByName('新原本');
let ASshift_sheet = SpreadsheetApp.openById('1IUUAyU2G4CqacM2g2BLWetBPD0UIV-bxGxuuQt0w69g').getSheetByName('新ASシフト');


function A_click() {
  const password = "1234567890";
  const write_pass = Browser.inputBox("パスワードを入力");
  if (write_pass != password){
    Browser.msgBox("パスワードが違います");
    return;
  }else{
    Browser.msgBox("実行します");
  }


  const n = sousa_sheet.getRange(7,2).getValue();
  const date_value = (sousa_sheet.getRange(1,1).getValue());
  const date_value2 = date_value;
  date_value2.setDate(date_value.getDate());
  sousa_sheet.getRange(1,13).setValue(date_value);
  const weekday_name = ['月','火','水','木','金','土','日'];
  var w = sousa_sheet.getRange(5,15).getValue();
  var m ;
  var d ;
  var sheet_name;
  var sheetToMove;
  date_value.setDate(date_value2.getDate()-1);
  for(i = 0;i < n;i++){;
    date_value.setDate(date_value2.getDate()+1);
    sousa_sheet.getRange(1,13).setValue(date_value);
    m = date_value.getMonth()+1;
    d = date_value.getDate();
    w = sousa_sheet.getRange(5,15).getValue();
    sheet_name = m + '/' + d + weekday_name[w];
    genpontest_sheet.copyTo(daily_spredsheet).setName(sheet_name);
    B_click(sheet_name);
    sheetToMove = daily_spredsheet.getSheetByName(sheet_name);
    daily_spredsheet.setActiveSheet(sheetToMove);
    daily_spredsheet.moveActiveSheet(8);
    sheetToMove.getRange(1,5).setValue(sheet_name);
  }
  
}

function B_click(sheet_name){
  const target_sheet = daily_spredsheet.getSheetByName(sheet_name);
  var i;
  var a;
  const range_start = sousa_sheet.getRange(5,13,1,2).getValues();
  const range_end = sousa_sheet.getRange(5,9,1,2).getValues();

  const  member_range= range_end[0][0]-range_start[0][0]-4;

  const staff_member = ASshift_sheet.getRange(range_start[0][0]+2,range_start[0][1],member_range,5).getValues()

  var l = staff_member.length;
  for(i = l-1;i>=0;i--){
    if(staff_member[i][2] == ""){
      staff_member.splice(i,1);
    }else{
      if(staff_member[i][2] < 6){
        a = staff_member[i][2];
        a = 6-a;
        staff_member[i][2] = 6;
        staff_member[i][4] -= a;
      }else if(staff_member[i][2]>staff_member[i][3]){
        staff_member[i][3]+=24;
      }
    }
  }

  l = staff_member.length;
  for(i = l-1;i>=0;i--){
    if (staff_member[i][3] > 26){
      var insert_row = 5;
      var minus = 23;
    }else{
      var insert_row = 3;
      var minus = 7;
    }
    target_sheet.insertRowAfter(insert_row);

    drowing(staff_member[i],insert_row+1,minus,target_sheet);
  }
  return 0;
}

function drowing(staff_member,row,minus,target_sheet){
  target_sheet.getRange(row,1).setValue(staff_member[0]);
  target_sheet.getRange(row,1).setFontSize(20);
  target_sheet.setRowHeight(row,43);
  target_sheet.getRange(row,1,1,2).merge();
  target_sheet.getRange(row,(staff_member[2]*2)-minus,1,staff_member[4]*2).setBackground('#00ffff');
  //target_sheet.getRange(row,4).insertCheckboxes().check();
  return 0;
}

//移行用----------------------------------------------------------------------------------------------------------------//
function Search(sheet,target) {//汎用
  // getValues ではなく getDisplayValues を使う（見たままを取得）
  const data = sheet.getDataRange().getDisplayValues(); 
  
  // シートの見た目が「01/18」なら、ここも合わせる 
  
  console.log("--- 検索開始 ---");

  for (let i = 0; i < data.length; i++) {
    for (let j = 0; j < data[i].length; j++) {
      // 前後の空白を除去して完全一致チェック
      const cellText = data[i][j].trim(); 
      
      if (cellText === target) {
        const row = i + 1;
        const col = j + 1;
        const address = sheet.getRange(row, col).getA1Notation();
        
        console.log(`発見！ 番地: ${address} (行:${row}, 列:${col})`);
        
        return {
          address: address,
          row: row,
          col: col
        };
      }
    }
  }
  console.log("見つかりませんでした。");
  return null;
}


function OneDayDrowing(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const Datasheet = ss.getSheetByName('ASシフト_テスト用');
  const Targetsheet = ss.getSheetByName('仮塗りtest');

  const DayName = Targetsheet.getRange(1,5).getDisplayValue();
  const result = Search(DayName);
  const row = result.row;
  const col = result.col;
  const Data = Datasheet.getRange(row+2,col,22,5).getValues();
  const DrowArrey = Array.from({length:22},()=> Array(40).fill("#ffffff"));
  const Names = Datasheet.getRange(row+2,col,22,1).getValues();

  for (var i = 0;i<22;i++){
    if(Data[i][0] = ""){
      break;
    }
    const s = Data[i][2];
    const n = Data[i][3];
    for(var j = (s-6)*2;j<(n-6)*2;j++){
      DrowArrey[i][j] = "#00ffff";
    } 
  }
  Targetsheet.getRange(4,5,22,40).setBackgrounds(DrowArrey);
  Targetsheet.getRange(4,1,22,1).setValues(Names);
}
