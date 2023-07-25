function sendMail() {
  var option = {
      to : "12k4@naver.com",//메일주소;
      subject : "테스트메일입니다.",//메일제목;
      htmlBody : "<html>메일 <br>오늘도 즐거운 하루</html>",//메일내용;
    };
  try{
    MailApp.sendEmail(option)
    console.log( "[+] Success : " + JSON.stringify(option) )
  }catch(e){
    console.log( "[+] Fail : " + e )
  }
}
