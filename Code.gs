var url="https://docs.google.com/spreadsheets/d/1dnboA_luJiv5PsJgfCbiJ0ZcL93M_2wxOBR4g0D-4D4/edit#gid=0";
var Route = {};

Route.path = function(route, callback){
  Route[route] = callback;
}

function doGet(e) {

  Route.path("newProject", projectForm);
  Route.path("projects", loadProject);

 if (Route[e.parameters.v]){
   return Route[e.parameters.v]();
 }  else {
   return render("home");
 }
}

function projectForm(){
  return HtmlService.createTemplateFromFile("newProject").evaluate();
  //render("newProject");
}

function loadProject(){
//   return HtmlService.createTemplateFromFile("projects").evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME)
// .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return render("projects");
}







