/**
 * 날짜를 여러포맷으로 변환하는함수
 * return {String}
 */
function formatDate( dateObj,format)
{
    format = format || 0;
    var dateObj = dateObj || new Date();
    var monthNames = [ "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December" ];
    var curr_date = dateObj.getDate();
    var curr_month = dateObj.getMonth();
    curr_month = curr_month + 1;
    var curr_year = dateObj.getFullYear();
    var curr_min = dateObj.getMinutes();
    var curr_hr= dateObj.getHours();
    var curr_sc= dateObj.getSeconds();
    if(curr_month.toString().length == 1)
    curr_month = '0' + curr_month;      
    if(curr_date.toString().length == 1)
    curr_date = '0' + curr_date;
    if(curr_hr.toString().length == 1)
    curr_hr = '0' + curr_hr;
    if(curr_min.toString().length == 1)
    curr_min = '0' + curr_min;

    if(format ==1)//dd-mm-yyyy
    {
        return curr_date + "-"+curr_month+ "-"+curr_year;       
    }
    else if(format ==2)//yyyy-mm-dd
    {
        return curr_year + "-"+curr_month+ "-"+curr_date;       
    }
    else if(format ==3)//dd/mm/yyyy
    {
        return curr_date + "/"+curr_month+ "/"+curr_year;       
    }
    else if(format ==4)// MM/dd/yyyy HH:mm:ss
    {
        return curr_year+"-"+curr_month +"-"+curr_date+ " "+curr_hr+":"+curr_min+":"+curr_sc;       
    }
}
function onEdit(e) {
    // Prevent errors if no object is passed.
    if (!e) return;
    // Get the active sheet.
    var newDate = Utilities.formatDate(new Date(),"GMT+09:00", "MM/dd/yyyy hh:mm:ss");
  var updteUserEmail = e.user.getEmail();
  var text  = 'Updated By ' + updteUserEmail;

    e.source.getActiveSheet()
        // Set the cell you want to update with the date.
        .getRange('A2')
        // Update the date.
        .setValue("Laste Updated : " + newDate + " | " + text);
    // Get the active sheet.


  // var bold  = e.source.getActiveSheet().newTextStyle().setBold(true).build();
  // var value = e.source.getActiveSheet().newRichTextValue().setText(updteUserEmail).setTextStyle(text.indexOf(updteUserEmail), text.length, bold).build();
  // e.source.getActiveSheet().getRange('A3').setRichTextValue(value);
}

function test(){
  var response = UrlFetchApp.fetch("https://swcamp-api-server.run.goorm.app/findAll");
  console.log(response.getContentText())
}