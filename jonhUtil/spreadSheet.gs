/**
 * 앱스크립트가 실행되는 문서를 가져온다.
 */
function getSpreadSheetByActive() {
  console.log("[ S ] - getSpreadSheetByActive");
  console.log("[ E ] - getSpreadSheetByActive");
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 문서 ID를 기반으로 해당 문서를 가져온다.
 */
function getSpreadSheetById( id ) {
  console.log("[ S ] - getSpreadSheetByActive");
  consoleLogDateTime( "Target SpreadSheet Id : " + id );
  console.log("[ E ] - getSpreadSheetByActive");
  return  SpreadsheetApp.openById(id)
}


/**
 * 새로운 스프레드시트문서룰 생성한다.
 */
function createSpreadSheetByName( nm ){
  console.log("[ S ] - createSpreadSheetByName");

  consoleLogDateTime( "Target Sheet Name : " + nm );
  var file = SpreadsheetApp.create( nm );

  console.log("[ E ] - createSpreadSheetByName");
  return file;
 }

/**
 * 스프레드시트의 시트를 순서기준으로 가져온다.
 */
function getSheetByNum( ss, num ) {
  console.log("[ S ] - createSpreadSheetByName");
  ss = ss || getSpreadSheetByActive();
  consoleLogDateTime( "Target Sheet Num : " + num );
  console.log("[ E ] - createSpreadSheetByName");
  return ss.getSheets()[num];
}
/**
 * 스프레드시트의 시트를 이름 기준으로 가져온다.
 */
function getSheetByNm( ss, Nm ) {
  console.log("[ S ] - getSheetByNm");
  ss = ss || getSpreadSheetByActive();
  consoleLogDateTime( "Target Sheet Name : " + nm );
  console.log("[ E ] - getSheetByNm");
  return ss.getSheetByName(nm);//시트를 가져온다;
}

/**
 * 스프레드시트의 이름을 배열로 가져온다.
 */
function getSheetNamesToArray( ss ){
  console.log("[ S ] - getSheetNamesToArray");
  ss = ss || getSpreadSheetByActive();
  var sheets = ss.getSheets();
  
  var r = [];
  var i = 0,iLen = sheets.length,io;
  for(;i<iLen;++i){
    io = sheets[ i ];
    r.push( io.getName() )
  }
  
  consoleLogDateTime(r)
  console.log("[ E ] - getSheetNamesToArray");
  return r
}

/**
 * 시트를 생성한다.
 */
function createSheetByTemplate( sss,templateSheetNm, tss, createSheetNm ){

  console.log("[ S ] - createSheetByTemplate");
  ss = sss || getSpreadSheetByActive();
  var templateSheet = ss.getSheetByName( templateSheetNm );

  //시트가 존재한다면 덮어쓴다.
  var copy = tss.getSheetByName( createSheetNm );
  if(copy)
  {
    tss.deleteSheet( copy )
  }

  var sheets = getSheetNamesToArray( tss );

  tss.insertSheet( createSheetNm, sheets.length, { template:templateSheet } );
  consoleLogDateTime( "Created Sheet : " + createSheetNm );

  console.log("[ E ] - createSheetByTemplate");

  return 
}


/**
 * 시트를 생성한다.
 */
function createSheet( ss, createSheetNm ){

  console.log("[ S ] - createSheet");
  ss = ss || getSpreadSheetByActive();

  //시트가 존재한다면 덮어쓴다.
  var copy = ss.getSheetByName( createSheetNm );
  if(copy)
  {
    ss.deleteSheet( copy )
  }

  var sheets = getSheetNamesToArray( ss );

  ss.insertSheet( createSheetNm, sheets.length );
  consoleLogDateTime( "Created Sheet : " + createSheetNm );

  console.log("[ E ] - createSheet");

  return 
}

/**
 * 데이터를 범위에 입력한다.
 */
function dataInertSheet( ss, targetSheetNm, row, column, optNumRows, optNumColumns, data ){
  console.log("[ S ] - dataInertSheet");
  ss = ss || getSpreadSheetByActive();
  var targetSheet = ss.getSheetByName( targetSheetNm )
  targetSheet.getRange(row, column, optNumRows, optNumColumns).setValues( data )
  console.log("[ E ] - dataInertSheet");
}


/**
 * 시트를 초기화한다.( 생성한 시트를 삭제한다.기본시트는 제외)
 */
function sheetDelete(ss, arr){
  
  console.log("[ S ] - sheetDelete");
  var ss = ss || SpreadsheetApp.getActiveSpreadsheet()

  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    ss.deleteSheet( ss.getSheetByName( io ) );
    consoleLogDateTime( "Deleted : " + io );

  }
  console.log("[ E ] - sheetDelete");
}


/**
 * 스프레드시트 데이터를 가져오는 함수
 *@return {array} r
 <CODE>
 [
   [1,2,3,4],
   [1,2,3,4],
   ...
 ]
 </CODE>
 */
function getDataByRange( ss, sheetNm , row, column, optNumRows, optNumColumns ){
  
  console.log("[ S ] - getDataByRange");
  ss = ss || getSpreadSheetByActive();
  var targetSheet = ss.getSheetByName( sheetNm )
  var r = targetSheet.getRange( sheetNm , row, column, optNumRows, optNumColumns ).getValues();
  console.log("[ E ] - getDataByRange");
  return r;
}

/** 
 * 
  */
function getSheetToJSON(o){
  var o = { spredadSheetId : "1UjPk_er30s5AeOLWQh_1YK76PGAFY8BMqVbfVbYzO7c", sheetNm : "총괄시트(읽기전용)_john_test", range : "",};
 
  var arr = getDataByRange( o );

  var header = arr[ 2 ];
  var _arr = []; 
  var i = 3,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    //console.log( typeof(io[0]) )
    var tmpO = {};
    if( typeof(io[0]) != "object" ) continue;
    //var date = Utilities.formatDate(new Date(io[0].toString()),"GMT+09:00", "yyyyMMDD")
    var j = 0,jLen = io.length,jo;
    for(;j<jLen;++j){
    jo = io[j];
    tmpO[ header[j] ] = jo;
    }
    _arr[ i ] = tmpO
  }
  console.log( _arr )
  return o;
}

function getSheetDataToArrayByRowIndex( o, n ){
  var n = 4
  var o = { spredadSheetId : "1UjPk_er30s5AeOLWQh_1YK76PGAFY8BMqVbfVbYzO7c", sheetNm : "총괄시트(읽기전용)_john_test", range : "",};
  var ss = null;
  if( o.spredadSheetId && o.spredadSheetId != "" ) ss = SpreadsheetApp.openById(o.spredadSheetId);
  else ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var targetSheet = null;
  if( o.sheetNm && o.sheetNm != "" ) targetSheet = ss.getSheetByName( o.sheetNm )
  else targetSheet = ss.getSheets()[ 0 ];

  var r = null;
  targetSheet.getLastColumn()
  r = targetSheet.getRange(n, 1, 1, targetSheet.getLastColumn()).getValues()
  console.log( r )
  return r[0];

}

function updateDataCell( o, row, col, d ){
  var n = 4
  var o = { spredadSheetId : "1UjPk_er30s5AeOLWQh_1YK76PGAFY8BMqVbfVbYzO7c", sheetNm : "총괄시트(읽기전용)_john_test", range : "",};
  var ss = null;
  if( o.spredadSheetId && o.spredadSheetId != "" ) ss = SpreadsheetApp.openById(o.spredadSheetId);
  else ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var targetSheet = null;
  if( o.sheetNm && o.sheetNm != "" ) targetSheet = ss.getSheetByName( o.sheetNm )
  else targetSheet = ss.getSheets()[ 0 ];

  //targetSheet.getRange(row, 1, 1, col).setValue("test")
  var preNote = targetSheet.getRange(4,19).getNote();
  var preValue = targetSheet.getRange(4,19).getValue();
  var d = d || "완료"
  note = `🛠 modified: ${Utilities.formatDate(new Date(),"GMT+09:00", "yyyy-MM-DD HH:mm:ss")}\n${preValue}\n↪${d}\n\n` + preNote
  targetSheet.getRange(4,19).setNote( note )
  targetSheet.getRange(4,19).setValue(d)
  
  return;

}