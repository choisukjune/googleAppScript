/**
 * 1차원배열을 2차원배열로 만든다
 */
function arrTo2depthByNum( arr, num ){
  console.log("[ S ] - arrTo2depthByNum");
  var r = [];
  var _ta = [];
  var i = 0,iLen = arr.length,io;
  for(;i<iLen;++i){
    io = arr[ i ];
    _ta.push( io )
    if( _ta.length == num ){
      r.push( _ta );
      _ta = [];
    }
  }
  r.push( _ta );
  // consoleLogDateTime( r );
  console.log("[ E ] - arrTo2depthByNum");
  return r
}