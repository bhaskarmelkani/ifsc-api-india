var XLSX = require('xlsx');
var elasticsearch = require('elasticsearch');
var fileName = '68774.xls';

var workbook = XLSX.readFile(fileName);

var sheets = workbook.Sheets,
  sheetKey,
cellKey;

var elasticClient = new elasticsearch.Client({
  host: 'localhost:9200',
  log: 'info'
});

var indexName = "bank_details";

function addDocument(document, id, callback) {
  try{
  if(false){
    return elasticClient.index({
      index: indexName,
      type: "spreadsheet",
      id: id,
      body: document
    }, function(){
      if(callback){
        callback(); 
      }
    });
  }else{
    return elasticClient.index({
      index: indexName,
      type: "spreadsheet",
      body: document
    }, function(){
      if(callback){
        callback(); 
      }
    });
  }
  }catch(e){
    console.log('Exception');
    console.log(e);
  }

}

var dataToIndex = [];


for(sheetKey in sheets){
  var sheet = sheets[sheetKey];
  var dataKeys = {};
  var cellObj = undefined;
  var cellStart = undefined,
    cellEnd = undefined,
    cellMatch;
  for(cellKey in sheet){
    var cell = sheet[cellKey];
    if(cellMatch = cellKey.match(/^([a-zA-Z]+)1$/)){
      dataKeys[cellKey] = cell['v'];
      if(!cellStart){
        cellStart = cellMatch[1];
      }
      cellEnd = cellMatch[1];
    }else{
      var indexMatch = cellKey.match(/^([a-zA-Z]+)[0-9]+$/);
      if(indexMatch){
        var index = indexMatch[1] + '1';
        if(indexMatch[1] === cellStart){
          cellObj = {};
        }
        cellObj[dataKeys[index]] = cell['v'];
        if(indexMatch[1] === cellEnd){
          dataToIndex.push([[cellObj['IFSC']], cellObj]);
        }
      }
    }
  }
}


function iterateAndIndex(){
  console.log('iterateAndIndex');
  console.log(dataToIndex.length);
  var obj = dataToIndex.pop();
  console.log(obj);
  if(!obj){
    console.log('returning');
    return; 
  }
  addDocument(obj[1], obj[0], iterateAndIndex);
};
iterateAndIndex();
