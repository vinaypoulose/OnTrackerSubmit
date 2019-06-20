// Looks up a dictionary of key value pairs (2d array) and returns the first value corresponding to the key specified
// 
// key: value to look up
//
// dictionary: two dimensional array with at least two columns and one row
//
// return: the value corresponding to the specified key. NULL if the key is not found.

function getValueFromDictionary(key, dictionary) {
  
  return vLookUp(key, dictionary, 1);
}

// Looks up a key in a two dimensional array and  returns the first value in the column specified correspoending to the key
// 
// key: value to look up
//
// range: two dimensional array with at least two columns and one row
//
// index: the column number whose value is to be returned
//
// return: the value corresponding to the specified key. NULL if the key is not found.

function vLookUp(key, range, index) {
  
  var i;
  
  for (i = 0; i < range.length; i++) {
    
    if (range[i][0].toString().toLowerCase() == key.toString().toLowerCase()) return range[i][index];
  }
  
  return null;
}


function vLookUpByColumnName(key, range, columnName) {
  
  var i;
  var columnNames = range[0];
  
  for (i = 0; i < columnNames.length; i++) {
    
    if (columnNames[i].toString().toLowerCase() == columnName.toString().toLowerCase()) {
      
      return vLookUp(key, range, i);
    }
  }
  
  return null;
}