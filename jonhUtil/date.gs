 
 /**
  * 닐찌를 yyyy-MM-dd HH:mm:ss 타입으로 리턴한다.
  */
 function getDateTime( date, dateType ){
  console.log( "[ S ] - getDateTime" )
  date = date || new Date();
  dateType = dateType || "YYYY-MM-DD hh:mm:ss";
  var r = Utilities.formatDate(date,"GMT+09:00", dateType )
  consoleLogDateTime( r );
  console.log( "[ E ] - getDateTime" )
  return r;
 }

function excelSerialDateToJSDate (excelSerialDate) {
  // "Excel serial date" is just
  // the count of days since `01/01/1900`
  // (seems that it may be even fractional).
  //
  // The count of days elapsed
  // since `01/01/1900` (Excel epoch)
  // till `01/01/1970` (Unix epoch).
  // Accounts for leap years
  // (19 of them, yielding 19 extra days).
  const daysBeforeUnixEpoch = 70 * 365 + 19;

  // An hour, approximately, because a minute
  // may be longer than 60 seconds, see "leap seconds".
  const hour = 60 * 60 * 1000;

  // "In the 1900 system, the serial number 1 represents January 1, 1900, 12:00:00 a.m.
  //  while the number 0 represents the fictitious date January 0, 1900".
  // These extra 12 hours are a hack to make things
  // a little bit less weird when rendering parsed dates.
  // E.g. if a date `Jan 1st, 2017` gets parsed as
  // `Jan 1st, 2017, 00:00 UTC` then when displayed in the US
  // it would show up as `Dec 31st, 2016, 19:00 UTC-05` (Austin, Texas).
  // That would be weird for a website user.
  // Therefore this extra 12-hour padding is added
  // to compensate for the most weird cases like this
  // (doesn't solve all of them, but most of them).
  // And if you ask what about -12/+12 border then
  // the answer is people there are already accustomed
  // to the weird time behaviour when their neighbours
  // may have completely different date than they do.
  //
  // `Math.round()` rounds all time fractions
  // smaller than a millisecond (e.g. nanoseconds)
  // but it's unlikely that an Excel serial date
  // is gonna contain even seconds.
  //
  return new Date(Math.round((excelSerialDate - daysBeforeUnixEpoch) * 24 * hour) + 12 * hour);
};