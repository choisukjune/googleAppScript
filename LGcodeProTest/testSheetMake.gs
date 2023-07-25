/**
 * LGcodePro 테스트 관리 파일을 생성하는 class
 */
class LGcodeProTest{
  constructor(option){
    this.sourceId = option.sourceId;
    this.sheet1Nm = option.sheet1Nm;
    this.sheet2Nm = option.sheet2Nm;
    this.pageSize = option.pageSize;

    this.sSheet = johnUtil.getSpreadSheetById( this.sourceId );
    this.targetSheet = this.sSheet.getSheetByName( this.sheet1Nm );
    this.templateSheet = this.sSheet.getSheetByName( this.sheet2Nm );
  }

  getSourceData(){
    this.dataSource = this.targetSheet.getDataRange().getValues();
  }

  creatNewFile(){
    this.newFileNm = "[" + johnUtil.dateTime__YYMMDDhhmmss() + "] - " + johnUtil.getFileNmById( this.sourceId )
    this.newFile = johnUtil.createSpreadSheetByName( this.newFileNm );
    this.tSheet = johnUtil.getSpreadSheetById( this.newFile.getId() )
  }

  copyTemplate(){
    this.templateSheet.copyTo(this.tSheet).setName(this.sheet2Nm);
    this.tTemplateSheet = this.tSheet.getSheetByName( this.sheet2Nm )
  }
  
  refineData(){
    this.refineData = johnUtil.arrTo2depthByNum( this.dataSource, this.pageSize)
  }
  
  deleteSheet( sheeNm ){
    this.tSheet.deleteSheet( this.tSheet.getSheetByName( sheeNm ) )
  }

  insertRefineDataToTemplate(){
    //var i = 0,iLen = this.refineData.length,io;
    var i = 0,iLen = 2,io;
    for(;i<iLen;++i){
      io = this.refineData[ i ];
      this.tSheet.insertSheet( i.toString(), { template:this.tTemplateSheet } );
      console.log( "시트생성" );
      this.tSheet.getSheetByName(i.toString()).getRange(6, 2, io.length, io[0].length).setValues( io )
    }
  }

  logic(){
    this.getSourceData(); //원본소스를 가져온다;
    this.refineData(); //원본데이터를 2차원배열로 정제한다.;
    this.creatNewFile(); //신규스프레드시트를 만든다;
    this.copyTemplate();//템플릿을 복사한다.;
    this.insertRefineDataToTemplate();//템플릿을 복사하면 데이터를 입력한다.;
    this.deleteSheet( "시트1" );//시트를삭제;
    this.deleteSheet( "Template" );//시트를삭제;
  }
}

function main(){
  
  var option = {
    "sourceId" : "1BoEEFAV9XCkMYIKOyINaTYM5MGthDSDzqaj-qPNgLEM",
    "sheet1Nm" : "RawData",
    "sheet2Nm" : "Template",
    "pageSize" : 16
  };

  var _t = new LGcodeProTest( option )
  _t.logic();
  console.log( _t.dataSource )
}

//---------------------------------------------------------------------------;
//---------------------------------------------------------------------------;
// [2023.04.13 - 초석준 ] 사용하지 않는파일 정리필요함 ;
//---------------------------------------------------------------------------;
//---------------------------------------------------------------------------;

// /**
//  * 시트를 생성한다.
//  */
// function createSheetByTemplate( templateSheetNm, createSheetNm ){
//   var ss = SpreadsheetApp.getActiveSpreadsheet()
//   var templateSheet = ss.getSheetByName( templateSheetNm );

//   //시트가 존재한다면 덮어쓴다.
//   var copy = ss.getSheetByName( createSheetNm );
//   if(copy)
//   {
//     ss.deleteSheet( copy )
//   }
  
//   var sheets = getSheetNamesToArray()

//   ss.insertSheet( createSheetNm, sheets.length, { template:templateSheet } );
//   console.log( "시트생성" );

//   return 
// }

// /**
//  * 데이터를 범위에 입력한다.
//  */
// function dataInertSheet( targetSheetNm, row, column, optNumRows, optNumColumns, data ){

//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var targetSheet = ss.getSheetByName( targetSheetNm )

//   var values = targetSheet.getDataRange().getValues()

//   targetSheet.getRange(row, column, optNumRows, optNumColumns).setValues( data )
// }



// /**
//  * 시트를 초기화한다.( 생성한 시트를 삭제한다.기본시트는 제외)
//  */
// function sheetInit(){
//   var sheets = getSheetNamesToArray()

//   var ss = SpreadsheetApp.getActiveSpreadsheet()

//   var i = 0,iLen = sheets.length,io;
//   for(;i<iLen;++i){
//     io = sheets[ i ];
//   console.log( io )
//     if( io != "RawData" && io != "Template" )
//     {
      
//       var _dSheet = ss.getSheetByName( io )
//       ss.deleteSheet( _dSheet );
//       console.log( "Deleted : " + io )
//     }
//   }
// }

//   // var ss = getSpreadSheetById( tId )
//   // var templateSheet = ss.getSheetByName( tSheetNm );



//   // ts.insertSheet( createSheetNm, sheets.length, { template:templateSheet } );
//   // console.log( "시트생성" );

//   // var ss = SpreadsheetApp.getActiveSpreadsheet();
//   // var targetSheet = ss.getSheetByName( targetSheetNm )

//   // var values = targetSheet.getDataRange().getValues()

//   // targetSheet.getRange(row, column, optNumRows, optNumColumns).setValues( data )

// /**
//  * 실행로직
//  */
// /*/
// function main(){
//   var targetSheetNm = "RawData";
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var targetSheet = ss.getSheetByName( targetSheetNm )
//   var arr = targetSheet.getDataRange().getValues();

//   console.log( arrTo2depthByNum( arr, 16) )
//   var sheetsData = arrTo2depthByNum( arr, 16)

//   var i = 0,iLen = sheetsData.length,io;
//   for(;i<iLen;++i){
//     io = sheetsData[ i ];
//     console.log( io )
//     createSheetByTemplate( "Template", i.toString() );
//     dataInertSheet( i.toString(), 6, 2, io.length, io[0].length, io )
//   }
  
// }
// */
