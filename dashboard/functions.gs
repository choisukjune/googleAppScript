function checkStatus(firstDate, secondDate, thirdDate){
  first = firstDate.substr(0,10)
  second = secondDate.substr(0,10)
  third = thirdDate.substr(0,10)
  td = new Date()
  today = toStringByFormatting(td)

  if(first===""){
    return '미정'
  }

  if(today < first){
    return "수업 전"
  }
  else if(today === first){
    return "1일차"
  }

  if(third === ""){
    if(today < second){
      return "진행중"
    }
    else if(today === second){
      return "2일차"
    }
    else{
      return "수업 완료"
    }
  }
  else{
    if(today < third){
      return "진행중"
    }
    else if(today === third){
      return "3일차"
    }
    else{
      return "수업 완료"
    }
  }

};

function leftPad(value) {
    if (value >= 10) {
        return value;
    }

    return `0${value}`;
}

function toStringByFormatting(source, delimiter = '-') {
    const year = source.getFullYear();
    const month = leftPad(source.getMonth() + 1);
    const day = leftPad(source.getDate());

    return [year, month, day].join(delimiter);
}

function refresh(){
  SpreadsheetApp.flush()
}