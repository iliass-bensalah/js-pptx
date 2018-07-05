

module.exports ={
		
	moveObjectElement: function(currentKey, afterKey, obj) {
		var temp=obj;
    var result = {};
    var val = obj[currentKey];
    delete obj[currentKey];
    var next = -1;
    var i = 0;
    if(typeof afterKey == 'undefined' || afterKey == null) afterKey = '';
    Object.keys(obj).forEach(function(k, v) {
        if((afterKey == '' && i == 0) || next == 1) {
            result[currentKey] = val; 
            next = 0;
        }
        if(k == afterKey) { next = 1; }
        result[k] = temp[k];
        ++i;
    },this);
    if(next == 1) {
        result[currentKey] = val; 
    }
    if(next !== -1) return result; else return obj;
}
		
		
}