
function getData(){
  var data = getProjects();
  Logger.log(data);
  return data;
}

//Note: the project sheet must have the same name as the project

function getProjects(){
  var main_folder = DriveApp.getFoldersByName("logging_app").next();
  var project_folders = main_folder.getFolders();

  var project_array = [];
  while(project_folders.hasNext()){
    var projects = project_folders.next();

    //get name and date
    var project_name = projects.getName();
    var date = new Date(projects.getDateCreated());
    var formattedDate = Utilities.formatDate(date, "GMT", "EEE, d MMM yyyy HH:mm:ss");
  
    //get sheet url
    var files = projects.getFilesByName(project_name);
    var sheet = files.next();
    var sheet_url = sheet.getUrl();
    
    var subarray=[project_name, formattedDate, sheet_url];
    project_array.push(subarray);

  }
  return project_array
}






