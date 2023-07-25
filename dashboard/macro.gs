function sort() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A3:BO1000').activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort({column: 1, ascending: true});
  spreadsheet.getRange('A2').activate();
};
function block_sort(){
  SpreadsheetApp.getActive().toast("추가된 학교는 데이터 하단에 덧붙여주시면 점검 완료 후 자동으로 날짜 순 정렬이 됩니다.","잠시 매크로 점검중 입니다")
};

function new_sort1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4').activate();

  // 플랫폼 정렬
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('플랫폼'), true);
  spreadsheet.getRange('A3').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('A3:X3').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);

  // OP 정렬
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('OP'), true);
  spreadsheet.getRange('A3').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getRange('A3:V3').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);

  // 굿즈 정렬
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('굿즈'), true);
  spreadsheet.getRange('A3:S1000').activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);

  // 에듀비즈 정렬
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('에듀비즈'), true);
  spreadsheet.getRange('A3').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);

  //수료증 정렬
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('수료증'), true);
  spreadsheet.getRange('A3:R1000').activate();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);

  // 학교 정렬
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('학교 데이터'), true);
  spreadsheet.getRange('A3').activate();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.getActiveRange().offset(1, 0, spreadsheet.getActiveRange().getNumRows() - 1).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);

  // 플랫폼 학교데이터 참조 리셋
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('플랫폼'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.getCurrentCell().setFormula('=\'학교 데이터\'!A4');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A4:Q4'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A4:Q4').activate();
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A4:Q1003'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('A4:Q1003').activate();

  // OP 학교데이터 참조 리셋
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('OP'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.getRange('\'플랫폼\'!A4:Q1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // 굿즈 학교데이터 참조 리셋
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('굿즈'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.getRange('\'플랫폼\'!A4:Q1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // 에듀비즈 학교데이터 참조 리셋
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('에듀비즈'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.getRange('\'플랫폼\'!A4:Q1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // 수료증 학교데이터 참조 리셋
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('수료증'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.getRange('\'플랫폼\'!A4:Q1000').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);


  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('총괄시트(읽기전용)'), true);
  spreadsheet.getRange('A4').activate();
  spreadsheet.toast('정렬이 완료되었습니다')
};


function sort_example() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4:X1000').activate()
  .sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
};

function new_sort() {
  var spreadsheet = SpreadsheetApp.getActive();
  const sheetNames = ['플랫폼', 'OP', '굿즈', '에듀비즈', '수료증', '학교 데이터']
  const numOfCols = [26, 15, 21, 58, 23, 19]
  const minRow = 4, maxRow = 900;
  var lastCol = 0;

  // 정렬
  for(var i  = 0; i<sheetNames.length; i++){
    var sheet = spreadsheet.getSheetByName(sheetNames[i])
    lastCol = numOfCols[i]
    sheet.getRange(minRow, 1, maxRow, lastCol).sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
  }
  
  spreadsheet.toast('정렬이 완료되었습니다')
};

function fetchDashboard(){
  var criteriaDate = SpreadsheetApp.getActive().getSheetByName('대시보드').getRange(15,3).getValue()
  var formattedDate = toStringByFormatting(criteriaDate)
  var result = [
    [ // 초등
      [0,0], // 서울
      [0,0], // 경기
      [0,0], // 부산
      [0,0], // 대구
      [0,0], // 경남
      [0,0], // 전북
      [0,0]  // 전남
    ],
    [ // 중등
      [0,0],// 서울
      [0,0],// 경기
      [0,0],// 부산
      [0,0],// 대구
      [0,0],// 경남
      [0,0],// 전북
      [0,0] // 전남
    ],
    [ // 고등
      [0,0],// 서울
      [0,0],// 경기
      [0,0],// 부산
      [0,0],// 대구
      [0,0],// 경남
      [0,0],// 전북
      [0,0] // 전남
    ],
  ]  // result[초,중,고][지역][신청,이수]

  var schoolData = SpreadsheetApp.getActive().getSheetByName('학교 데이터').getRange(4,1,900, 19).getValues()
  schoolData.forEach(school =>{
    // 취소 된 학교거나 아직 종료되지 않은 학교는 무시
    if(school[1].includes('x')) {return;}
    else if(school[10] === '' && school[9].substr(0,10) >= formattedDate) {return;}
    else if(school[10] !== '' && school[10].substr(0,10) >= formattedDate) {return;}
    else if(school[10] === '' && school[9] === '') {return;}

    var schoolType = 0;
    switch(school[2]){
      case '초등' :
        schoolType = 0;
        break;
      case '중등' :
        schoolType = 1;
        break;
      case '고등' :
        schoolType = 2;
        break;
    }

    switch(school[4]){
      case '서울' :
        result[schoolType][0][0] += Number(school[12]);
        result[schoolType][0][1] += Number(school[13]);
        break;
      case '경기' :
        result[schoolType][1][0] += Number(school[12]);
        result[schoolType][1][1] += Number(school[13]);
        break;
      case '부산' :
        result[schoolType][2][0] += Number(school[12]);
        result[schoolType][2][1] += Number(school[13]);
        break;
      case '대구' :
        result[schoolType][3][0] += Number(school[12]);
        result[schoolType][3][1] += Number(school[13]);
        break;
      case '경남' :
        result[schoolType][4][0] += Number(school[12]);
        result[schoolType][4][1] += Number(school[13]);
        break;
      case '전북' :
        result[schoolType][5][0] += Number(school[12]);
        result[schoolType][5][1] += Number(school[13]);
        break;
      case '전남' :
        result[schoolType][6][0] += Number(school[12]);
        result[schoolType][6][1] += Number(school[13]);
        break;
    }
  })
  SpreadsheetApp.getActive().getSheetByName('대시보드').getRange('D19:E25').setValues(result[0]);
  SpreadsheetApp.getActive().getSheetByName('대시보드').getRange('G19:H25').setValues(result[1]);
  SpreadsheetApp.getActive().getSheetByName('대시보드').getRange('J19:K25').setValues(result[2]);
}
