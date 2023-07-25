
function onEdit(e){
  console.log(e)
  var addrBase = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK', 'BL', 'BM', 'BN', 'BO', 'BP', 'BQ', 'BR', 'BS', 'BT', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CD', 'CE', 'CF', 'CG', 'CH', 'CI', 'CJ', 'CK', 'CL', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT', 'CU', 'CV' ];
  console.log(range.getColumn() - 1)
  var range = e.range;
  var key = addrBase[range.getColumn() - 1];
  var row = range.getRow();
  var as = e.source.getActiveSheet();
  var _id = as.getRange(row, as.getLastColumn() ).getValue();
  var value = e.value;

var r = {}
  r[ "_id" ] =  _id;
  r[ key ] = value;

    var url_winter = "https://swcamp-api-server.run.goorm.app/updateOne"
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(r)
    };
    
    console.log( r)
    var a = UrlFetchApp.fetch(url_winter,options);
}
function sheetToJSON() {

  var sheetNm = "john_test";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tsheet = ss.getSheetByName(sheetNm);//시트를 가져온다;
  var tV = tsheet.getDataRange().getValues()

  var r = [];

  var keys = _buildColumnsArray(tV[ 0 ].length)//tV[ 0 ];
  var i = 1,iLen = tV.length,io;
  //var i = 1,iLen = 2,io;

  for(;i<iLen;++i){
    io = tV[ i ];
    var o = {}
    var j = 0,jLen = io.length,jo;
    for(;j<jLen;++j){
      jo = io[ j ];
      o[ keys[ j ] ] = jo
    }
    //r.push( o )
    var url_winter = "https://swcamp-api-server.run.goorm.app/insertOne"
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(o)
    };
    
    var a = UrlFetchApp.fetch(url_winter,options);
    console.log( a.getContentText() )
    tsheet.getRange(i+1,tsheet.getLastColumn()).setValue( a.getContentText().replace(/\n/gi,"") )
  }

  

  //console.log( r );
  //saveAsJSON( r );
}

function saveAsJSON( d ) {
  var blob,file,fileSets,d;
  
  fileSets = {
    title: 'AAA_Test.json',
    mimeType: 'application/json'
  };
  
  blob = Utilities.newBlob(JSON.stringify(d,null,4), "application/vnd.google-apps.script+json");
  file = Drive.Files.insert(fileSets, blob);
  Logger.log('ID: %s, File size (bytes): %s, type: %s', file.id, file.fileSize, file.mimeType);

}

function alphaToNum(alpha) {
  var i = 0,
      num = 0,
      len = alpha.length;

  for (; i < len; i++) {
    num = num * 26 + alpha.charCodeAt(i) - 0x40;
  }

  return num - 1;
}

function numToAlpha(num) {
  var alpha = '';

  for (; num >= 0; num = parseInt(num / 26, 10) - 1) {
    alpha = String.fromCharCode(num % 26 + 0x41) + alpha;
  }

  return alpha;
}

function _buildColumnsArray(num) {
  var num = 100;

  var res = [];
  for (i = 0; i < num ; i++) {
    res.push(numToAlpha(i));
  }
  console.log(res)
  return res;
}