function getFileLsitByFolderIdToArr( folderId ){
    var folderId = folderId || "1cmTAtuA0CjCz9qQuOdXlcYI40SXsWhP7do6yzg8-APmtXg4Kl1MOYJ40b4cLhEPpg4bXNdTG";
    var daforder = DriveApp.getFolderById(folderId);
    var dafiles = daforder.getFiles();
    
    var r = [];
    while(dafiles.hasNext()){
      var dafile = dafiles.next();
      var o = {
        id : dafile.getId(),
        //nm : dafile.getName(),
        //cdt : dafile.getDateCreated(),
        //udt : dafile.getLastUpdated(),
        //owner : dafile.getOwner().getName(),
        //mime : dafile.getMimeType(),
        //url : dafile.getUrl(),
      }
      r.push( o )
    }
    console.log( r )
    return r
}

function get_filelist() {
  
   
  // get This Folder ID
  var folderId = "1cmTAtuA0CjCz9qQuOdXlcYI40SXsWhP7do6yzg8-APmtXg4Kl1MOYJ40b4cLhEPpg4bXNdTG";
  var daforder = DriveApp.getFolderById(folderId);
  var dafiles = daforder.getFiles();
 
  //var sheet = SpreadsheetApp.getActiveSheet();
  var targetSheetNm = "출석부리스트_20230127";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheetNm);//시트를 가져온다;
  //sheet.clear();
 
  var srow = 1;
  // Write Header
  sheet.getRange(srow, 1).setValue("file name");
  sheet.getRange(srow, 2).setValue("만들어진 날짜");
  sheet.getRange(srow, 3).setValue("마지막으로 수정한 날짜");
  sheet.getRange(srow, 4).setValue("소유자");
  sheet.getRange(srow, 5).setValue("file type");
  sheet.getRange(srow, 6).setValue("documentId");
  sheet.getRange(srow, 7).setValue("Link");
  // Set Header color
  var range = sheet.getRange("A1:F1");
  range.setBackground("#f3f3f3");
   
  // Get file names
  while(dafiles.hasNext()){
    var dafile = dafiles.next();
    var file_name = dafile.getName();
    console.log( file_name );
    srow = srow + 1;
    // Write file info
    sheet.getRange(srow, 1).setValue(file_name);
    sheet.getRange(srow, 2).setValue(dafile.getDateCreated());
    sheet.getRange(srow, 3).setValue(dafile.getLastUpdated());
    sheet.getRange(srow, 4).setValue(dafile.getOwner().getName());
    sheet.getRange(srow, 5).setValue(dafile.getMimeType().replace('application/vnd.google-apps.', 'google '));
    sheet.getRange(srow, 6).setValue(dafile.getId())
    sheet.getRange(srow, 7).setValue(dafile.getUrl());
    
  }
   
  range = sheet.getRange("F:F");
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
   
}

function getSheetData(FileId){

  var FileId = FileId || "1gFge1R4oRk10Pw2AqSLekm3katqt7OEn";
  var excelfile = DriveApp.getFileById(FileId);
  
  var blob = excelfile.getBlob();

  var newFile = {
      title : '_converted',
      parents: [{id: "17PpNwyq5RV8KVQWRpelz3WEsTUc1jy6K"}] 
    }; 
    
  var tempFile = Drive.Files.insert(newFile, blob, { convert: true });
  
  //console.log( tempFile.getId() );
  
  var tt = tempFile.getId()

  var targetSheetNm = "출석자명단통합";
  var targetSheetError = "errorList";
  var sourceSheetNm = "참여자 명단";

  //var ssDB = SpreadsheetApp.openById(tt).getSheetByName(sourceSheetNm)
  var ssDB = SpreadsheetApp.openById(tt).getSheets()[0];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet00 = ss.getSheetByName(targetSheetNm);//시트를 가져온다;
  
  //var targetSheet01 = ss.getSheetByName(targetSheetError)
  var sourceSheet00 = ssDB;//시트를 가져온다;
  
  var lastcol = sourceSheet00.getLastColumn(); //자료가 있는 마지막 행
  var values = sourceSheet00.getDataRange().getValues(); 

  targetSheet00.setFrozenRows(1);
  var d00 = "";
  var d01 = "";
  for(n=0; n<values.length;++n){ 

    var io = values[n]; 
    var lastRow = targetSheet00.getLastRow()+1;

   
    if( io[7].toString().indexOf("담당교사") != -1 ) d00 = io[8]
    if( io[7].toString().indexOf("수업일시") != -1 ) d01 = io[8]
    if( io[2].toString().indexOf("이름") != -1 ) continue;
    if( io[5].toString().indexOf("00초등학교") != -1 ) continue;
    if( io[1].toString() == "" || io[1].toString().indexOf("파일 저장명") != -1 ) continue;

    io[0] = d00;
    io[1] = d01;

    targetSheet00.appendRow(io)
    // var j = 0,jLen = io.length,jo;
    // for(;j<jLen;++j){
    //   jo = io[ j ];
  
    //   targetSheet00.getRange(lastRow, j+1 ).setValue(jo)
    // }
    
  }
  Drive.Files.remove(tt);
}


function getSheetDataAndDocId(FileId){

  var FileId = FileId || "1gFge1R4oRk10Pw2AqSLekm3katqt7OEn";
  var excelfile = DriveApp.getFileById(FileId);
  
  var blob = excelfile.getBlob();

  var newFile = {
      title : '_converted',
      parents: [{id: "17PpNwyq5RV8KVQWRpelz3WEsTUc1jy6K"}] 
    }; 
    
  var tempFile = Drive.Files.insert(newFile, blob, { convert: true });
  
  //console.log( tempFile.getId() );
  
  var tt = tempFile.getId()
  var ssDB = SpreadsheetApp.openById(tt).getSheets()[0];
  
  var sourceSheet00 = ssDB;//시트를 가져온다;
  
  
  var d00 = "";
  var d01 = "";
  var d02 = "";
  try
  {
    var values = sourceSheet00.getDataRange().getValues(); 

    for(n=0; n<values.length;++n){ 

      var io = values[n]; 
    
      if( io[7].toString().indexOf("담당교사") != -1 )
      {
        if(io[8] == "" ) d00 = io[9]
        else d00 = io[8]
      }
      if( io[7].toString().indexOf("수업일시") != -1 )
      {
        if(io[8] == "" ) d01 = io[9]
        else d01 = io[8]
      }
      if( io[5] != "" && io[5] != "학교명" && io[5] != "00초등학교" ) d02 = io[5]

      if( d00 != "" && d01 != "" && d02 != "" )
      {
        break;
      }
    
    }
    
  }
  catch(e)
  {
    console.log(e)
  }
  Drive.Files.remove(tt);
  console.log([ d00, d01, d02 ])
  return [ d00, d01, d02 ];
}

function get_filelistAndInfo() {
   
  // get This Folder ID
  var folderId = "1cmTAtuA0CjCz9qQuOdXlcYI40SXsWhP7do6yzg8-APmtXg4Kl1MOYJ40b4cLhEPpg4bXNdTG";
  var daforder = DriveApp.getFolderById(folderId);
  var dafiles = daforder.getFiles();
 
  var targetSheetNm = "출석부리스트_학교_선생님_20230213";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(targetSheetNm);//시트를 가져온다;
  sheet.clear();
 
  var srow = 1;
  // Write Header
  sheet.getRange(srow, 1).setValue("선생님");
  sheet.getRange(srow, 2).setValue("수업일자");
  sheet.getRange(srow, 3).setValue("학교명");
  sheet.getRange(srow, 4).setValue("file name");
  sheet.getRange(srow, 5).setValue("만들어진 날짜");
  sheet.getRange(srow, 6).setValue("마지막으로 수정한 날짜");
  sheet.getRange(srow, 7).setValue("소유자");
  sheet.getRange(srow, 8).setValue("file type");
  sheet.getRange(srow, 9).setValue("documentId");
  sheet.getRange(srow, 10).setValue("Link");
  sheet.getRange(srow, 11).setValue("error");
  // Set Header color
  var range = sheet.getRange("A1:I1");
  range.setBackground("#f3f3f3");
   
  // Get file names
  while(dafiles.hasNext()){
    var dafile = dafiles.next();
    var file_name = dafile.getName();
    console.log( file_name );
    srow = srow + 1;
    // Write file info
    var cell11 = ""
    try
    {
      var tmp000 = getSheetDataAndDocId(dafile.getId())
    }
    catch(e)
    {
      cell11= "error - " + e 
    }
    console.log( tmp000 )
    if( !tmp000 )
    { 
      tmp000 = [ "","",""]
    }
    sheet.getRange(srow, 1).setValue(tmp000[0]);
    sheet.getRange(srow, 2).setValue(tmp000[1]);
    sheet.getRange(srow, 3).setValue(tmp000[2]);
    sheet.getRange(srow, 4).setValue(file_name);
    sheet.getRange(srow, 5).setValue(dafile.getDateCreated());
    sheet.getRange(srow, 6).setValue(dafile.getLastUpdated());
    sheet.getRange(srow, 7).setValue(dafile.getOwner().getName());
    sheet.getRange(srow, 8).setValue(dafile.getMimeType().replace('application/vnd.google-apps.', 'google '));
    sheet.getRange(srow, 9).setValue(dafile.getId())
    sheet.getRange(srow, 10).setValue(dafile.getUrl());
    sheet.getRange(srow, 11).setValue(cell11);
    
  }
   
  range = sheet.getRange("J:J");
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
   
}

function getList(){
    var targetSheetNm = "출석부리스트_20230127";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetSheet00 = ss.getSheetByName(targetSheetNm);//시트를 가져온다;
    var arr = targetSheet00.getDataRange().getValues(); 

    var i = 1,iLen = arr.length,io;
    for(;i<iLen;++i){
      io = arr[ i ];
      console.log( io )
      if(io[4] == "application/haansoftxlsx" || io[4] == "image/jpeg" || io[7]) continue;
      console.log( io[5] )
      try
      {
        getSheetData(io[5])
        targetSheet00.getRange(i+1,8).setValue(true)
      }
      catch(e)
      {
        var m = "docId : " + io[5] + "\n" + "Error : " +  e; 
        targetSheet00.getRange(i+1,9).setValue(m)
      }

      
    }

}