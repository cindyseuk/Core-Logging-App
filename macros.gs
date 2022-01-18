
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheet = spreadsheet.getActiveSheet();

function onLoad() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hole Logs')
      .addItem('Search', 'searchHole')
      .addItem('Create a new Hole', 'createNewHole')
      .addSeparator()
      .addItem('Toggle Row Height', 'changeRowHeightRelativeToMetres')
      .addItem('Reset Row Height', 'changeRowHeightToDefault')
      .addToUi();
    
  format();
  freezeHeader();
  colourHeaders();
}
/**
 * 
 */
function changeRowHeightRelativeToMetres(){
  var row=8;
  var toCol = sheet.getRange(row,2).getValue();
  var fromCol = sheet.getRange(row,1).getValue();
  var distance = toCol-fromCol;
  sheet.setRowHeight(row, distance*20);
  row++;
  toCol = sheet.getRange(row,2).getValue();
  fromCol = sheet.getRange(row,1).getValue();

  while(toCol!=0 && fromCol!=0){
    distance = toCol-fromCol;
    sheet.setRowHeight(row, distance*20);
    row++;
    toCol = sheet.getRange(row,2).getValue();
    fromCol = sheet.getRange(row,1).getValue();
  }
}

function changeRowHeightToDefault(){
  var row=8;
  sheet.setRowHeight(row, 10);
  row++;
  var toCol = sheet.getRange(row,2).getValue();
  var fromCol = sheet.getRange(row,1).getValue();
  while(toCol!=0 && fromCol!=0){
    sheet.setRowHeight(row, 10);
    row++;
    toCol = sheet.getRange(row,2).getValue();
    fromCol = sheet.getRange(row,1).getValue();
  }
}
/**
 * Prompt user for a Hole ID to search for and sets it as the active sheet
 */
function searchHole(){
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Search for Hole',
    'Please enter the Hole ID:',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();

  //check for valid input
  var checkValid = spreadsheet.getSheetByName(text);
  if (checkValid==null){
    ui.alert("A hole with that ID does not exist. Please try again");
    createNewHole();
    return;
  }
  var sheet = spreadsheet.getSheetByName(text);
  spreadsheet.setActiveSheet(sheet);
}

/**
 * Create a new hole by giving it a new ID. 
 * If a hole already exists with that id, report error
 */
function createNewHole(){
  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(
    'Creating a new hole',
    'Please enter the Hole ID:',
      ui.ButtonSet.OK_CANCEL);

  var button = result.getSelectedButton();
  var text = result.getResponseText();

  //check for valid input
  var checkValid = spreadsheet.getSheetByName(text);
  if (checkValid!=null){
    ui.alert("A hole with that ID already exists. Please enter another one.");
    createNewHole();
    return;
  }

  if (button == ui.Button.OK) {
    makeDuplicateOfTemplate();
    var totalSheets = spreadsheet.getNumSheets();
    var newSheet = spreadsheet.getSheets()[totalSheets-1];
    var currentSheet = SpreadsheetApp.setActiveSheet(newSheet);
    var date = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
    currentSheet.getRange(1,13).setValue(date);
    currentSheet.setName(text);   //rename
    currentSheet.getRange(1,5).setValue(text)
                .setHorizontalAlignment('left');
    currentSheet.getRange(1,2)
                .setValue(Session.getActiveUser().getEmail())
                .setHorizontalAlignment('left');
    format();
  } 
}

/**
 * Creates duplicate of master template
 */
function makeDuplicateOfTemplate(){
  var spreadsheetId = spreadsheet.getId();
  var spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  var projectFolder = spreadsheetFile.getParents().next();

  var tempFile = projectFolder.getFilesByName("Template").next();
  var templateSheet = SpreadsheetApp.open(tempFile);
  var sheet = templateSheet.getSheets()[0];
  var destination = SpreadsheetApp.openById(spreadsheetId);
  sheet.copyTo(destination); 
}

function colourHeaders() {
  helper_col(1,2,'#b6d7a8');
  helper_col(6,2,'#b6d7a8');
  helper_col(3,3,'#d9ead3');
  helper_col(7,1,'#d9ead3');
};

/**
 * Helper function to colour rows
 * 
 * @params{row}
 * @params{numRows}
 * @params{colour}
 */
function helper_col(row, numRows, colour){
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(row,1,numRows,sheet.getLastColumn()).activate();
  spreadsheet.getActiveRangeList().setBackground(colour);
}
/**
 * function to freeze header and metres column
 */
function freezeHeader(){
  var sheet = spreadsheet.getActiveSheet();
  var headerRange = sheet.getRange(1, 1, 7, sheet.getLastColumn()).activate();
  spreadsheet.getActiveRange().protect();

  //freeze rows
  sheet.getRange(1,1, 7, sheet.getLastColumn()).activate();
  spreadsheet.getActiveSheet().setFrozenRows(7);

  //freeze columns
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1,1, sheet.getMaxRows(), 2).activate();
  spreadsheet.getActiveSheet().setFrozenColumns(2);
}
/**
 * function to format header
 */
function format() {
  outline_border(2);
  outline_border(6);
  //set centre alignment
  var currentSheet = spreadsheet.getActiveSheet();
  var range= currentSheet.getRange(8,1, sheet.getMaxRows(), 13);
  range.setHorizontalAlignment("center");
  create_dropdown();
};

function outline_border(row_idx){
  var currentSheet= spreadsheet.getActiveSheet();
   currentSheet.getRange(row_idx,1,1,currentSheet.getLastColumn()).activate();
   spreadsheet.getActiveRangeList()
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBorder(true, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK); 
}
/**
 * Checks for numeric input type 
 */
function create_dropdown() {
  numeric_warning(8,1,3);
  numeric_warning(8,10,2);
  dropdown("L3:L5",8,4);  //oxidation
  dropdown("A3:G5",8,5);  //lithology
  dropdown('H3:I5',8,6);  //texture
  dropdown("K3:K5",8,8);  //alteration
  dropdown("K3:K5",8,9);  //alteration
  dropdown("I3:J5",8,7);  //grain size
};

function numeric_warning(row, col, numCols){
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(row,col, sheet.getMaxRows(), numCols).activate();
  spreadsheet.getActiveRange().setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(false)
  .requireNumberBetween(0, 100000000)
  .build());
}

/**
 * Creates dropdown menus corresponding to legend
 * 
 * @params{range} 
 * @params{row} 
 * @params{col} 
 */
function dropdown(range, row, col) {
  var sheet = spreadsheet.getActiveSheet();
  var values = sheet.getRange(range).getValues();
  var list=[];

  for (var i=0; i<values.length; i++){
    for (var j=0; j<values[i].length; j++){
      var value = values[i][j].slice(0,values[i][j].search(":"));
      list.push(value);
    }
  }

  sheet.getRange(row,col, sheet.getMaxRows(), 1).activate();
  spreadsheet.getActiveRange()
  .setDataValidation(SpreadsheetApp.newDataValidation()
  .setAllowInvalid(true)
  .requireValueInList(list)
  .build());
};

//unfinished
function onSelectionChange(e) {
  var range = e.range;
  if (range.getNumRows() === 1 &&
      range.getNumColumns() === 1 &&
      range.getCell(1, 1).getValue() === '') {
  //  range.setBackground('red');
    //sheet.setRowHeight(1, 30);
  }
};

//the header cannot be edited 
function protect_header() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1,1, 7, sheet.getLastColumn()).activate();
  var protection = spreadsheet.getActiveRange().protect();
  protection.setWarningOnly(true);
};

function hide_rows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  if (sheet.isRowHiddenByUser(1)){
    sheet.showRows(1,5);
  } else {
    sheet.hideRows(1, 5);
  }
};


