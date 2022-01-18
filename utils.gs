
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function render(file, argsObject){
  var temp=HtmlService.createTemplateFromFile(file);
  if (argsObject){
    var keys = Object.keys(argsObject);
      keys.forEach(function(key){
        temp[key]=argsObject[key];
      });
  }
  return temp.evaluate();//.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}