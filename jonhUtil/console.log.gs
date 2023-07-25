/**
 * 
 */
function consoleLogDateTime( msg ){
  var now = Utilities.formatDate(new Date(),"GMT+09:00", "yyyy-MM-dd HH:mm:ss")
  var txt = [
    "[ " + now + " ]"
  ]
  if( msg ) txt.push( msg );
  console.log( txt.join( " - " ))
}

function consoleLogDateTimeReturnStr( msg ){
  var now = Utilities.formatDate(new Date(),"GMT+09:00", "yyyy-MM-dd HH:mm:ss")
  var txt = [
    "[ " + now + " ]"
  ]
  if( msg ) txt.push( msg );
  var r = txt.join( " - " )
  console.log( r );
  return r;
}