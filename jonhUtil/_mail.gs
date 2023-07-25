var _this = this;
    _this.ActiveUser = Session.getActiveUser().getEmail();
    _this.MailType = null;
    _this.MailInfo =[];
    _this.MailTypeByMailInfo ={};
    //_this.AllData = getSheetToJSON({ spredadSheetId : "1UjPk_er30s5AeOLWQh_1YK76PGAFY8BMqVbfVbYzO7c", sheetNm : "총괄시트(읽기전용)_john_test", range : "",})


/**
 * 메일을 발송 할 리스트를 가져온다.
 * contact.gs (V3)
 * @return {arr}
 */
function getMailList() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var targetSheet00 = ss.getSheetByName(targetSheetNm);//시트를 가져온다;
    var mailList = ss.getSheets()[0].getDataRange().getValues()
    return mailList;
}
/**
 * 발송할 메일의 컨테츠를 가져오는 함수.
 * contact.gs (V3)
 * @param {String} p 메일컨테츠의 키값.
 * @return {String} r
 */
function getContentsList( p ) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    //var targetSheet00 = ss.getSheetByName(targetSheetNm);//시트를 가져온다;
    var arr = ss.getSheets()[1].getDataRange().getValues()
    var o = {};
    console.log(arr)
    var i = 0,iLen = arr.length,io;
    for(;i<iLen;++i){
      io = arr[ i ];
      if( !o[ io[0] ] ) o[ io[0] ] = io[ 1 ];
    }
    console.log( o );
    var r = o[ p ];
    return r;
}


/**
 * 메일을 발송한다.
 * gs (V3)
 * @param {object} p 메일을 보내기위한 정보.
 <CODE>
  {
    to : "12k4@naver.com",
    subject : "mail TEST 입니다.",
    htmlBody : "<html>메일테스트 발송 입니다. <br> 오늘도 즐거운하루! ok? </html>",
  }
  </CODE>
 * @return {void}
 */
function sendMail( p ) {
  try
  {
    console.log( "[S] - " + p.to + " - " +  p.subject )
    var o = {
      to : p.to,
      subject : p.subject,
      htmlBody : p.htmlBody
      
    }
    console.log( o )
    MailApp.sendEmail( o ,{from: "name@domain.com"} )
    console.log( "[E] - " + p.to + " - " +  p.subject )
  }
  catch(e)
  {
    console.log( "[e] - " + e );
  }
  return;
}

/**
 * 메일을 발송한다.
 * gs (V3)
 * @param {object} p 메일을 보내기위한 정보.
 <CODE>
  {
    to : "12k4@naver.com",
    subject : "mail TEST 입니다.",
    htmlBody : "<html>메일테스트 발송 입니다. <br> 오늘도 즐거운하루! ok? </html>",
  }
  </CODE>
 * @return {void}
 */
function sendMailByContentId( p ) {
  try
  {
    var p = p || { to : "12k4@naver.com", contentType : "case1" }
    var content =  getDataByRange({ spredadSheetId : null, sheetNm : "mailContents", range : null,})
    
    var _o = {};
    
    var i = 0,iLen = content.length,io;
    for(;i<iLen;++i){
      io = content[ i ]
      if( !_o[ io[0] ] ) _o[ io[0] ] = io;
    }
    console.log( _o )

    var o = {
      dateTime : Utilities.formatDate(new Date(),"GMT+09:00", "yyyy-MM-DD HH:mm:ss"),
      to : p.to,
      bcc : _this.ActiveUser,//숨은참조를 발송자 자신으로 설정해 혹시 모를 내용오류에 대해 확인할수 있도록 처리,
      subject : _o[p.contentType][1],
      htmlBody : _o[p.contentType][2],
      domIdx : p.domIdx
      
    }
    MailApp.sendEmail( o )
    return o;
  }
  catch(e)
  {
    console.log( "[error] - " + e );
  }
  return;
}

/**
 * 전체로직을 실행하는 함수
 * contact.gs (V3)
 * @return {void}
 */
function logic(){
  
  var arr = getMailList();
  var contentsO = getContentsList( "case1" );
  var newDate = Utilities.formatDate(new Date(),"GMT+09:00", "MM/dd/yyyy hh:mm:ss");
  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    var o = {
      to : io[0],
      subject : "mail TEST 입니다.",
      htmlBody : contentsO.replace( "{=CONTENTS=}", newDate )
    }
    //sendMail( o )
  }
  return;
}

function insertHistory( arr ) {
  try
  {
    var ss =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("history")
    ss.appendRow( arr );
  }
  catch(e)
  {
    console.log( "[error] - " + e );
  }
  return;
}

function renderDialog(){

  //init;
  _this.MailType = "mail-1";
  _this.MailInfoByMailType = getMailinfoCode();
  _this.MailInfo = getMailListBySheetNm( _this.MailInfoByMailType[ _this.MailType ] );

  console.log( _this );
  
  var html = `
  <!DOCTYPE html>
  <html>
    <head>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/semantic-ui/2.5.0/semantic.min.js" integrity="sha512-Xo0Jh8MsOn72LGV8kU5LsclG7SUzJsWGhXbWcYs2MAmChkQzwiW/yTQwdJ8w6UA9C6EVG18GHb/TrYpYCjyAQw==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/semantic-ui/2.5.0/semantic.min.css" integrity="sha512-KXol4x3sVoO+8ZsWPFI/r5KBVB/ssCGB5tsv2nVOKwLg33wTFP3fmnXa47FdSVIshVTgsYk/1734xSk9aFIa4A==" crossorigin="anonymous" referrerpolicy="no-referrer" />
    <script src="https://code.jquery.com/jquery-3.6.3.min.js" integrity="sha256-pvPw+upLPUjgMXY0G+8O0xUf+/Im1MZjXxxgOcBQBXU=" crossorigin="anonymous"></script>
  <script>
    var mailCode = ${JSON.stringify(this.MailInfoByMailType)};
    console.log(mailCode)
    function sendEmailByMailType() { 
      
      var targetEmailInfo = document.getElementsByClassName("targetEmailInfo");

      var _el_result = document.getElementById('result');
      var mailType =  "${_this.MailType}"

      var arr = [];
      var i = 0,iLen = targetEmailInfo.length,io;
      for(;i<iLen;++i){
        
        io = targetEmailInfo[ i ];
        var tmpArr = [];

        var j = 0,jLen = io.children.length,jo;
        for(;j<jLen;++j){
          jo = io.children[j]
          tmpArr.push(jo.innerText);
        }
        arr.push(tmpArr);
     }

      var i = 0,iLen = arr.length,io;
      for(;i<iLen;++i){
        
        io = arr[ i ];
        
        var p = {
          to : io[ 6 ],
          subject : "",
          htmlBody : "",
          contentType : mailType
        }
        google.script.run.sendMailByContentId( p )
        

        var now = new Date();
        io.shift();
        _el_result.innerHTML = _el_result.innerHTML + "<tr style='font-size:12px;'><td>" + now + "</td><td>" + io.join(" / ") + "</td><td>" + mailType + "</td><td>발송완료</td></tr>"
        
        
        var historyO = {}
        io.push( mailType );
        io.push( "발송완료" );
        historyArr = historyArr.concat(io)
        google.script.run.insertHistory(historyArr)
 
        
  
     }    
    }

    var deleteRow = function(e){
      var i = e.parentNode.parentNode.rowIndex;
      var _table00 = document.getElementById("table00");
      _table00.deleteRow(i)
      if(_table00.children[1].childElementCount == 0 )
      {
        _table00.children[1].innerHTML = "<tr><td colspan='7' style='text-align:center''><h3>발송대상 이메일이 존재하지 않습니다.</h3></tr>"
      }
    }

    window.addEventListener('DOMContentLoaded', function(){
      
      var _btns_deletEmail = document.getElementsByClassName("deletEmail");
      var select_contentType = document.getElementById("_select_contentType");
      
      var i = 0,iLen = _btns_deletEmail.length,io;
      for(;i<iLen;++i){
        io = _btns_deletEmail[ i ];
        io.addEventListener("click", function(e){ deleteRow(e.currentTarget) });
      }

      select_contentType.addEventListener('change', function(e){  
        
        var r =  e.target.options[e.target.selectedIndex].value;
        var _table00 = document.getElementById("table00");
        var _tbody = document.getElementById("table00");
        
        loaderOn();
        
        _table00.children[1] = "";
        
        var sheetNm = mailCode[ r ];
        
        google.script.run.withSuccessHandler (function (result) {
        
          console.log(result)
          google.script.run.withSuccessHandler (function (result) {
          
            console.log(result)
            _table00.children[1].innerHTML = result;
            
            var _btns_deletEmail = document.getElementsByClassName("deletEmail");
            var i = 0,iLen = _btns_deletEmail.length,io;
            for(;i<iLen;++i){
              io = _btns_deletEmail[ i ];
              io.addEventListener("click", function(e){ deleteRow(e.currentTarget);});
            }
            loaderOff();
          }).updateMailInfoListToHtml(result)
        }).getMailListBySheetNm(sheetNm)

      });
    })
    function loaderOn(){
      var loader = document.getElementById("dimmer");
          loader.classList.add("active");
    }
    function loaderOff(){
      var loader = document.getElementById("dimmer");
          loader.classList.remove("active");
    }
  </script>
  <style>
  td{
    font-size:12px;
  }
  </style>
  </head>
  <body>
  
    <div class="ui dimmer blurring" id="dimmer">
        <div class="ui massive text loader">
            <h3>Loading</h3>
        </div>
    </div>
    <div class="ui stackable grid">
      <div class="sixteen wide column">
        <div class="ui success message">
          <i class="close icon"></i>
          <div class="header">
            SWCAMP mail Service
          </div>
          <p>메일발송 테스트중입니다..</p>
        </div>
      </div>
      <div class="sixteen wide column">
          <div class="ui form">
            <div class="field">
              <label>메일타입 선택</label>
              <select id="_select_contentType">
              {=OPTOINS=}
              </select>
            </div>
          </div>
      </div>
      <div class="sixteen wide column">
        <div class="ui form">
          <div class="field">
            <label>발송대상 이메일 리스트</label>
            <div>{=EAMIL_LIST=}</div>
          </div>
        </div>
      </div>
      <div class="sixteen wide column">
        <label>발송대상 이메일 리스트</label>
        <div>
        <table class="ui compact celled table" style="min-height: 200px;">
          <thead>
            <tr>
              <th>DateTime</th>
              <th>info</th>
              <th>Mail Type</th>
              <th>Result</th>
            </tr>
          </thead>
          <tbody id="result">
          </tbody>
        </table>
        </div>
      </div>
      <div class="sixteen wide column">
        <button class="ui green basic button" onclick="sendEmailByMailType()"><i class="icon envelope outline"></i>발송</button>
        <button class="ui orange basic button" onclick="google.script.host.close()"/><i class="icon close icon"></i>취소</button><br><br>
      </div>
    </div>

  </body>
  </html>
  `;

  var htmlOption = getOptionListToHtml();
  var tableThtml = mailInfoListToHtml( _this.MailInfo );

  var html = html.replace( "{=EAMIL_LIST=}", tableThtml )
    .replace( "{=OPTOINS=}", htmlOption );;
  
  //다이알로그 그리기
  var htmlOutput = HtmlService.createHtmlOutput(html).setWidth(1000).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "SWCAMP 메일발송");
  
}

function renderSidebar(){

  //init;
  _this.MailType = "mail-1";
  _this.MailInfoByMailType = getMailinfoCode();
  _this.MailInfo = getMailListBySheetNm( _this.MailInfoByMailType[ _this.MailType ] );

  console.log( _this );
  
  var html = `
 
    <script src="https://cdnjs.cloudflare.com/ajax/libs/semantic-ui/2.5.0/semantic.min.js" integrity="sha512-Xo0Jh8MsOn72LGV8kU5LsclG7SUzJsWGhXbWcYs2MAmChkQzwiW/yTQwdJ8w6UA9C6EVG18GHb/TrYpYCjyAQw==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/semantic-ui/2.5.0/semantic.min.css" integrity="sha512-KXol4x3sVoO+8ZsWPFI/r5KBVB/ssCGB5tsv2nVOKwLg33wTFP3fmnXa47FdSVIshVTgsYk/1734xSk9aFIa4A==" crossorigin="anonymous" referrerpolicy="no-referrer" />
    <script src="https://code.jquery.com/jquery-3.6.3.min.js" integrity="sha256-pvPw+upLPUjgMXY0G+8O0xUf+/Im1MZjXxxgOcBQBXU=" crossorigin="anonymous"></script>
  <script>
    var mailCode = ${JSON.stringify(this.MailInfoByMailType)};
    var mailType = "${_this.MailType}";
    console.log(mailCode)
    
    function sendEmailByMailType() { 
      
      var targetEmailInfo = document.getElementsByClassName("targetEmailInfo");
      var select_contentType = document.getElementById("_select_contentType");
      var _el_result = document.getElementById('result');
      var _tMailType = select_contentType.options[select_contentType.selectedIndex].value;
      var arr = [];
      var i = 0,iLen = targetEmailInfo.length,io;
      for(;i<iLen;++i){
        
        io = targetEmailInfo[ i ];
        console.log( io )
        arr.push(io.innerText);
     }

      var i = 0,iLen = arr.length,io;
      for(;i<iLen;++i){
        
        io = arr[ i ];
        console.log(io)
        var p = {
          to : io,
          subject : "",
          htmlBody : "",
          contentType : _tMailType,
          domIdx : i
        }
        google.script.run.withSuccessHandler (function (result) {
          
          console.log(result)
          
          var _table00 = document.getElementById("table00");
          var _tDom = _table00.children[ result.domIdx ];


          _tDom.style.backgroundColor = "#eee"
          _tDom.children[0].children[1].children[0].innerText += " ✅"


          var message = "[" + result.dateTime + "] " + result.to + " - " + result.subject + " - Success\\n-------------------------\\n";
          _el_result.textContent +=  message
          var historyArr = [result.dateTime,result.to,result.bcc,result.subject,result.htmlBody,"Success"];
          
          
          google.script.run.insertHistory(historyArr)
          
        }).sendMailByContentId( p ) 
     }    
    }

    var deleteRow = function(e){
      
      var _emailListEmpty = document.getElementById("emailListEmpty");
      var target =  e.currentTarget.parentNode.parentNode;
      target.parentNode.removeChild(target);
      
      if( table00.childElementCount == 0 )
      {
        _emailListEmpty.classList.remove("hidden");        
      }
    }

    window.addEventListener('DOMContentLoaded', function(){
      
      renderMailList( mailType )

      var _btns_deletEmail = document.getElementsByClassName("deletEmail");
      var select_contentType = document.getElementById("_select_contentType");
      var _emailListEmpty = document.getElementById("emailListEmpty");

      var i = 0,iLen = _btns_deletEmail.length,io;
      for(;i<iLen;++i){
        io = _btns_deletEmail[ i ];
        io.addEventListener("click", function(e){ 
          deleteRow(e.currentTarget) 
        });
      }

      select_contentType.addEventListener('change', function(e){  
        
        var r =  e.target.options[e.target.selectedIndex].value;
        renderMailList( r )
      });
    })
    function loaderOn(){
      var loader = document.getElementById("dimmer");
          loader.classList.add("active");
    }
    function loaderOff(){
      var loader = document.getElementById("dimmer");
          loader.classList.remove("active");
    }
    function renderMailList( mailType ){
        var _table00 = document.getElementById("table00");
        var _emailListEmpty = document.getElementById("emailListEmpty");      
        var select_contentType = document.getElementById("_select_contentType");
        var _tMailType = select_contentType.options[select_contentType.selectedIndex].value;

        if( _emailListEmpty.classList.value.indexOf("hidden") == -1 ) _emailListEmpty.classList.add("hidden");

        
        loaderOn();
        
        _table00.children[1] = "";
        
        var sheetNm = mailCode[ _tMailType ];
        
        google.script.run.withSuccessHandler (function (result) {
        
          console.log(result)
          google.script.run.withSuccessHandler (function (result) {
          
            console.log(result)
            _table00.innerHTML = result;
            
            var _btns_deletEmail = document.getElementsByClassName("deletEmail");
            var i = 0,iLen = _btns_deletEmail.length,io;
            for(;i<iLen;++i){
              io = _btns_deletEmail[ i ];
              io.addEventListener("click", function(e){ deleteRow(e);});
            }
            loaderOff();
          }).updateMailInfoListToHtml(result)
        }).getMailListBySheetNm(sheetNm)
    }
  </script>
  <style>
  td{
    font-size:12px;
  }
  .column{
    padding-left : 0.5rem!important;
    padding-right : 2rem!important;
  }
  .description{
    font-size:11px;
  }
  .meta{
    font-size:0.9rem!important;
  }
  </style>

  
    <div class="ui dimmer blurring" id="dimmer">
        <div class="ui massive text loader">
            <h4>Loading</h4>
        </div>
    </div>
    <div class="ui equal width padded grid">
      <div class="row">
        <div class="column">
          <div class="ui info message">
            <div class="header">
              SWCAMP mail Service
            </div>
            <p>메일발송 테스트중입니다..</p>
          </div>
        </div>
      </div>
      <div class="row">
      
        <div class="column">
            <div class="ui form">
              <div class="field">
                <label>메일타입 선택</label>
                <select id="_select_contentType">
                {=OPTOINS=}
                </select>
              </div>
            </div>
        </div>
      
      </div>
      <div class="row">
        <div class="column">
          <div class="ui form">
            <div class="field">
              <label>발송대상 이메일 리스트</label>
              <div id="emailListEmpty" class="ui negative message hidden"><div class="header">발송메일이 존재하지 않습니다.</div><p>메일리스트를 확인해주세요</p></div>
              <div class="ui cards" id="table00"></div>
            </div>
          </div>
        </div>
      
      </div>

      <div class="row">

        <div class="column">
          <div class="ui form">
            <div class="field">
              <label>발송결과</label>
              <textarea id="result" style="font-size: 12px; height: 300px;"></textarea>
            </div>
          </div>
        </div>

      </div>

      <div class="row">

        <div class="column">
          <button class="ui green basic button" onclick="sendEmailByMailType()"><i class="icon envelope outline"></i>발송</button>
          <button class="ui orange basic button" onclick="google.script.host.close()"/><i class="icon close icon"></i>취소</button><br><br>
        </div>
    
      </div>
    </div>

  `;

  var htmlOption = getOptionListToHtml();

  var html = html.replace( "{=OPTOINS=}", htmlOption );;
  
  //다이알로그 그리기
  var htmlOutput = HtmlService.createHtmlOutput(html).setTitle('SWCAMP Mail Service');//.setWidth(1000).setHeight(800);
  //SpreadsheetApp.getUi().showModalDialog(htmlOutput, "SWCAMP 메일발송");
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(htmlOutput);
  
}

function getMailinfoCode(){
  var p = { spredadSheetId : null, sheetNm : "mailContents", range : null };
  var arr = getDataByRange( p );
  var o = {}
  
  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    o[ io[0] ] = io[3]; 
  }

  return o;
}

function getOptionListToHtml(){
  var p = {
    spredadSheetId : null,
    sheetNm : "mailContents", 
    range : null,
  }
  var arr = getDataByRange( p )

  var htmlOption = '<option value="{=type=}">{=type=}</option>';
  var _htmlOption = "";
  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    _htmlOption += htmlOption.replace( /{=type=}/gi, io[ 0 ] ) + "\n";
  }
  return _htmlOption;
}

function getMailListBySheetNm( SheetNm ){
  console.log(SheetNm)
  var p00 = {
    spredadSheetId : null,
    sheetNm : SheetNm,//"test_mail_address", 
    range : null,
  };
  var arr = getDataByRange( p00 )
  console.log(arr)
  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    io[ 0 ] = io[ 0 ].toString() 
  }
  return arr
}

function mailInfoListToHtml( arr ){

  var tmpHtml = ``;
  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    tmpHtml += `
      <div class="card">
        <div class="content">
          <i class="icon close icon deletEmail right floated"></i>
          <div class="header">
            <h5 class="ui header">${io[ 5 ]}</h5>
          </div>
          <div class="meta">
            ${Utilities.formatDate(new Date(io[ 0 ].toString()),"GMT+09:00", "yyyy-MM-dd")}
          </div>
          <div class="description">
            <span>${io[ 1 ]}</span> / 
            <span>${io[ 2 ]}</span> / 
            <span>${io[ 3 ]}</span> / 
            <span>${io[ 4 ]}</span> / 
            <span class="targetEmailInfo">${io[ 5 ]}</span>
          </div>
        </div>
        <!--div class="extra content">
          <div class="ui two buttons">
            <div class="ui basic green button">Approve</div>
            <div class="ui basic red button">Decline</div>
          </div>
        </div-->
      </div>\n
    `
  }

  return tmpHtml;
}
function updateMailInfoListToHtml( arr ){

  var tmpHtml = ``;
  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    tmpHtml += `
      <div class="card">
        <div class="content">
          <i class="icon close icon deletEmail right floated"></i>
          <div class="header">
            <h5 class="ui header">${io[ 5 ]}</h5>
          </div>
          <div class="meta">
            ${Utilities.formatDate(new Date(io[ 0 ].toString()),"GMT+09:00", "yyyy-MM-dd")}
          </div>
          <div class="description">
            <span>${io[ 1 ]}</span> / 
            <span>${io[ 2 ]}</span> / 
            <span>${io[ 3 ]}</span> / 
            <span>${io[ 4 ]}</span> / 
            <span class="targetEmailInfo">${io[ 5 ]}</span>
          </div>
        </div>
        <!--div class="extra content">
          <div class="ui two buttons">
            <div class="ui basic green button">Approve</div>
            <div class="ui basic red button">Decline</div>
          </div>
        </div-->
      </div>\n
      `
  }

  return tmpHtml;
}
//----------------------------------------------------------------------------------------------------;
//----------------------------------------------------------------------------------------------------;
// Event.
//----------------------------------------------------------------------------------------------------;
//----------------------------------------------------------------------------------------------------;
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('연락')
    .addItem('이메일 - 다이얼로그', 'renderDialog')
    //.addItem('이메일 - 사이드바', 'renderSidebar')
    .addToUi();
    renderSidebar();
}
function tyt(){
    console.log(_this.AllData)

}

//-----------------------------------------------------------------------------------------------;
//-----------------------------------------------------------------------------------------------;
//SAMPLE_CODE;
//-----------------------------------------------------------------------------------------------;
//-----------------------------------------------------------------------------------------------;


