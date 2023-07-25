function pad(n, width){
	n = n + '';
	return n.length >= width ? n : new Array(width - n.length + 1).join('0') + n;
}
/**
 * @title fcLPad 왼쪽 문자열 채움 
 * @param src    원래 문자열
 * @param len    pad후 자리수
 * @param padStr src가 모자랄 경우 왼쪽에 채울 문자
 * @OpenIssues   pad후 자리수 < 입력문자열길이 인 경우 pad후 자리수로 자른다.
 * @사용방법
 *      var ret = fcLPad("200", 5, "0"); //00200
 * @return 변환문자열
 */
function fcLPad(src, len, padStr) {
    var retStr = "";
    var padCnt = Number(len) - String(src).length;
    if (Number(padCnt) < 1) {
    	return String(src).substring(0, Number(len));
    }
    for(var i=0;i<padCnt;i++) {
    	retStr += String(padStr);
    }
    return retStr+src;
}
/**
 * fcRPad 오른쪽 문자열 채움 
 * @param src    원래 문자열
 * @param len    pad후 자리수
 * @param padStr src가 모자랄 경우 오른쪽에 채울 문자
 * @OpenIssues   pad후 자리수 < 입력문자열길이 인 경우 pad후 자리수로 자른다.
 * @사용방법
 *      var ret = fcRPad("123", 5, "0"); //12300
 * @return 변환문자열
 */
function fcRPad(src, len, padStr){
    var retStr = "";
    var padCnt = Number(len) - String(src).length;
    if (Number(padCnt) < 1) {
    	return String(src).substring(0,Number(len));
    }
    for(var i=0;i<padCnt;i++) {
    	retStr += String(padStr);
    }
    return src+retStr;
}