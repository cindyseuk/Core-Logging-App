//shows sidebar when the From column is editted
function onEdit(e){
 var range = e.range;

 var column = range.getColumn();
 if (column==2){
   //check for validation
   var row = range.getRow();
   if (validateMetres(row)){
     showPredictionSidebar();
   }
 }
};
/**
 * Validation for From and To column.
 * The To value must be greater than From
 */
function validateMetres(row){
  var cell = sheet.getRange(row,2);
  var fromCol = sheet.getRange(row,1).getValue();
  var toCol = cell.getValue();
  
  if (fromCol>=toCol){
    var rule= SpreadsheetApp.newDataValidation().
              requireNumberGreaterThan(fromCol).
              build();
    cell.setDataValidation(rule);
    return false;
  }
  return true;
}

function showPredictionSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('predictionSidebar')
      .setTitle('Lithology Prediction');
  SpreadsheetApp.getUi() 
      .showSidebar(html);

}

function getParentFolder(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var folders = file.getParents();
  while (folders.hasNext()){
    Logger.log('folder name = '+folders.next().getName());
  }
}

/**
 * Function that retrieves the lithologies and probabilities from the master sheet 
 * and returns as an array in descending order
 */
  
function getLithologyFromMasterSheet(){
  //get parent folder 
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  var projectFolder = spreadsheetFile.getParents().next();

  var masterFile = projectFolder.getFilesByName("Master").next();
  var masterSheet = SpreadsheetApp.open(masterFile);
  SpreadsheetApp.setActiveSpreadsheet(masterSheet);
  
  var sheet = masterSheet.getSheets()[0];
  var lithologySheet = SpreadsheetApp.setActiveSheet(sheet);

  //sort probabilities by descending order
  lithologySheet.sort(5,false);
  
  var lithArray = [];
  var probability = lithologySheet.getRange(1,5).getValue();
  var lithName = lithologySheet.getRange(1,2).getValue();
  lithArray.push([lithName, probability]);
  
  for (var i=1; i<lithologySheet.getLastRow(); i++){
      var probability = lithologySheet.getRange(i,5).getValue();
      var lithName = lithologySheet.getRange(i,2).getValue();
      if (probability>0){
        var array = [lithName, probability];
        lithArray.push(array);
      }
  }
  Logger.log(lithArray);
  return lithArray;
}

/**
 * Function that retrieves the descriptions of the lithology
 * 
 * @params{lithology}
 */

function getLithologyDescription(lithology){
 // lithology="garnet";
  var spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  var spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  var projectFolder = spreadsheetFile.getParents().next();

  var masterFile = projectFolder.getFilesByName("Master").next();
  var masterSheet = SpreadsheetApp.open(masterFile);
  SpreadsheetApp.setActiveSpreadsheet(masterSheet);
  
  var sheet = masterSheet.getSheets()[0];
  var lithologySheet = SpreadsheetApp.setActiveSheet(sheet);

  var values = lithologySheet.getDataRange().getValues();
  var rowIndex;

  for (var i=1; i<values.length; i++){
    if (values[i][1]==lithology){
      rowIndex=i;
      break;
    }
  }
  var description = lithologySheet.getRange(rowIndex+1,3).getValue();
  return description;
}






