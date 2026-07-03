function Create_work_schedule(){
  const sheet = SpreadsheetApp.openById('1p4Ra9Qw1k0OtpehblNJAEqELuJm1bnied1gCSpBLZyQ').getSheetByName('提出データ');
  const ASsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('新ASシフト');
  const data = sheet.getDataRange().getValues();

  var i,j,k,l;
  var a,b,c;
  var Temp_Array;
  const N1 = data.length;
  const N2 = data[0].length-1;
  var DayDataArray = new Array(30).fill(["","","",""]);

  for (k = 2;k<N2;k += 2){
    j = 0;
    DayDataArray = new Array(30).fill(["","","",""]);
    for (i=2;i<N1;i++){
      if ((data[i][k] != "")&&(data[i][k] != -1)&&(data[i][k]!=data[i][k+1])){
      DayDataArray[j]=[data[i][0],"",data[i][k],data[i][k+1]];
      j ++;
      }
    }
    for (i = 0;i<j-1;i++){
      a = i;
      for(l = i+1;l<j;l++){
        if(DayDataArray[l][2]==""){
          break;
        }
        if (DayDataArray[l][2]<DayDataArray[a][2]){
          a=l;
        }else if(DayDataArray[l][2]==DayDataArray[a][2]){
          if(DayDataArray[l][3]<DayDataArray[a][3]){
            a=l;
          }
        }
      }
      Temp_Array = DayDataArray[i];
      DayDataArray[i]=DayDataArray[a];
      DayDataArray[a]=Temp_Array;
    }


    const result = Search(ASsheet,data[1][k]);

    const row = result.row;
    const col = result.col;
    console.log(`行: ${row}, 列: ${col}`);

    ASsheet.getRange(row+2,col,30,4).setValues(DayDataArray);
  }
  return;
}

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

function OneDayDrowing(){//人数調整用
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const Datasheet = ss.getSheetByName('新ASシフト');
  const Targetsheet = ss.getSheetByName('人数調整用');
  const Statussheet = ss.getSheetByName('AS属性表');

  const DayName = Targetsheet.getRange(1,5).getDisplayValue();
  const result = Search(Datasheet,DayName);
  const row = result.row;
  const col = result.col;
  const Data = Datasheet.getRange(row+2,col,30,5).getValues();
  const DrowArrey = Array.from({length:30},()=> Array(40).fill("#ffffff"));
  const Names = Datasheet.getRange(row+2,col,30,1).getValues();
  const MemberStatus = Statussheet.getDataRange().getValues();
  const MemberStatusLengh = MemberStatus.length;
  const DrowingMode = Targetsheet.getRange(1,9).getValue();
  var DrowColor = "#00ffff";

  for (var i = 0;i<30;i++){
    if(Data[i][0] == ""){
      break;
    }
    if (DrowingMode){
      DrowColor = "#00ffff";
      for(var j = 0;j<MemberStatusLengh;j++){
        if(MemberStatus[j][0]==Data[i][0]){
          if(MemberStatus[j][1]=="男"){
            DrowColor = "#1ea3ff";
          }else if(MemberStatus[j][1]=="女"){
            DrowColor = "#fd4484";
          }
        }
      }
    }

    var s = Data[i][2];
    if (s < 6){s = 6;}
    var n = Data[i][3];
    if(n>26){n = 26;}
    for(var j = (s-6)*2;j<(n-6)*2;j++){
      DrowArrey[i][j] = DrowColor;
    } 
  }
  Targetsheet.getRange(4,5,30,40).setBackgrounds(DrowArrey);
  Targetsheet.getRange(4,1,30,1).setValues(Names);
}
