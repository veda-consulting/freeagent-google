function onInstall(){
  onOpen();
}

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Update TimeSheets", functionName: "freeagentEntryPoint"}];
  ss.addMenu("Veda", menuEntries);
   
}


function requestAndHandleData() {
  var response = getData();
  var responseCode = response.getResponseCode();
  if (responseCode < 300) {
    //Browser.msgBox(response.getContentText());
    writeDataToSpreadsheet(JSON.parse(response.getContentText()));
    
  } else if (responseCode == 401) { 
    // Try to reauthenticate using the refresh token. 
    // If that fails, try to log in.
    //Browser.msgBox("Error 401 - attempting to reauthenticate");
    if (refreshToken()) { requestAndHandleData(); } 
    else { login(); }
    
  } else { 
    // Some unexpected error - print it out
    Browser.msgBox("Error " + responseCode + ": " + response.getContentText());
  }
}

function freeagentEntryPoint(){
  if(isTokenValid()){
    HTMLToOutput = '<html><h1>Timesheets Updated</h1></html>';
    var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
    status = authInfo.getAuthorizationStatus();
    url = authInfo.getAuthorizationUrl();
    //adddatatosheet(status);
    //adddatatosheet(url);
    getData();
    //tabsData();
    //HTMLToOutput += uploadData();
    SpreadsheetApp.getActiveSpreadsheet().show(HtmlService.createHtmlOutput(HTMLToOutput)); 
  }
  else {//we are starting from scratch or resetting
    HTMLToOutput = "<html><h1>You need to login</h1><a href='"+getURLForAuthorization()+"'>click here to start</a><br>Re-open this window when you return.</html>";
    SpreadsheetApp.getActiveSpreadsheet().show(HtmlService.createHtmlOutput(HTMLToOutput));    
  }
}



function doGet(e) {
  var HTMLToOutput;
  if(e.parameters.code){//if we get "code" as a parameter in, then this is a callback. we can make this more explicit
    getAndStoreAccessToken(e.parameters.code);
    HTMLToOutput = '<html><h1>Finished with oAuth</h1>You can close this window.</html>';
  }
  return HtmlService.createHtmlOutput(HTMLToOutput);
}

//do meaningful freeagent access here
function getData(){
  return runSOQL('SELECT+name+from+Account');
}

var uniqueUserArray = [],      
    uniqueTaskArray = [],
    uniqueProjectArray = [],
    uniqueContactArray = [];

function runSOQL(soql){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetTimeZone = ss.getSpreadsheetTimeZone(); 
  var sheets = ss.getSheets();
  var sheet = ss.getActiveSheet();  
  
   var defaultSheet = sheets[0];
  
  // Push Veda Consulting into the contacts Array
  processContactArray.push(['Veda Consulting']);
  
  Logger.log("processContactArray = " + processContactArray );
  
  //days = Browser.inputBox('Enter the number of days to  fetch the timesheets', 'Default '+days+' days', Browser.Buttons.OK_CANCEL);
  var fromdate = Utilities.formatDate(subDaysFromDate(new Date(),days), "GMT", "yyyy-MM-dd");
  var titles = [];
  // Push the Titles on the spreadsheet
  titles.push(["Task Name", "Date", "Hours", "Developer", "", "Client", "Project", "Comment", "is Billable?", "Rate", "Bill Period", "Status"]);
  
  Logger.log('fromdate' + fromdate);
  //get URL to fetch all timeslips
  // Eg : https://api.sandbox.freeagent.com/v2/timeslips?page=1&per_page=100&from_date=2015-07-15;
  var getTimeslipsURL = apiurl+'timeslips?page=1&per_page='+perpage+'&from_date='+fromdate;
  
  var getPagesUrl = UrlFetchApp.fetch(getTimeslipsURL,getUrlFetchPOSTOptions()).getHeaders().Link;
  
  var lastPageNo = 1;
  
  if(getPagesUrl!==""){
    
    var spliturl = getPagesUrl.substring(getPagesUrl.lastIndexOf('?') + 1);
  
    Logger.log("spliturl" + spliturl);
    lastPageNo = spliturl.split('page=')[1].split('&')[0];
    
    Logger.log("lastPageNo" + lastPageNo);
  
  } 
  
  var pageNo, 
      timeslipsArray= [],
      totalArray = [];
  
  for(v=1; v<=lastPageNo; v++){
    
    pageNo = v;        
  
    //get URL to fetch all timeslips by page
    // Eg : https://api.sandbox.freeagent.com/v2/timeslips?page=1&per_page=75>
    var getTimeslipsPerPageURL = apiurl+'timeslips?page='+pageNo+'&per_page='+perpage+'&from_date='+fromdate;
    
    var timeslipsResponse = UrlFetchApp.fetch(getTimeslipsPerPageURL,getUrlFetchPOSTOptions()).getContentText();
    
     //parse JSON data and get  timeslips Array
    timeslipsArray[v] = JSON.parse(timeslipsResponse).timeslips;
    
   // Browser.msgBox("Number of timeslips in page "+pageNo+" = "+timeslipsArray[v].length)
 
    totalArray = totalArray.concat(timeslipsArray[v]);

  }  
  Logger.log("total length" + totalArray.length)  ;

  for (c = 0; c < processContactArray.length; c++){
    
    var orgRows =[],
      orgName;
    
    orgName = processContactArray[c];  
    
    var rows = [],
        timeslip;  
    
    // Loop to get all timeslip data and push data to rows
    for (i = 0; i < totalArray.length; i++) {
      timeslip = totalArray[i];
      
      var userurl = timeslip.user;   
      var taskurl = timeslip.task;    
      var projecturl = timeslip.project;
      var contacturl;
      
      // If a new userurl found call the getUserDetails function    
      if(uniqueUserArray.indexOf(userurl) == -1){
        getUserDetails(userurl);
      }
      
      // Loop through unique USER array and assign userdetails
      for(j = 0; j < uniqueUserArray.length; j++){      
        if(userurl == uniqueUserArray[j].url){
          developerFirst = uniqueUserArray[j].first_name;
          developerLast = uniqueUserArray[j].last_name;
        }
      }
      
      // If a new taskurl found call the getTaskDetails function    
      if(uniqueTaskArray.indexOf(taskurl) == -1){
        getTaskDetails(taskurl);
      }
      
      // Loop through unique TASK array and assign taskdetails
      for(k = 0; k < uniqueTaskArray.length; k++){      
        if(taskurl == uniqueTaskArray[k].url){              
          taskName = uniqueTaskArray[k].name;
          isBillable = uniqueTaskArray[k].is_billable;
          billingRate = uniqueTaskArray[k].billing_rate;
          billingPeriod = uniqueTaskArray[k].billing_period;
          status = uniqueTaskArray[k].status;
        }
      }
      
      // If a new projecturl found call the getProjectDetails function    
      if(uniqueProjectArray.indexOf(projecturl) == -1){
        getProjectDetails(projecturl);
      }
      
      // Loop through unique PROJECT array and assign project details
      for(m = 0; m < uniqueProjectArray.length; m++){      
        if(projecturl == uniqueProjectArray[m].url){              
          projectName = uniqueProjectArray[m].name;
          contacturl = uniqueProjectArray[m].contact;
        }
      }
      
      // If a new contacturl found call the getContactDetails function    
      if(uniqueContactArray.indexOf(contacturl) == -1){
        getContactDetails(contacturl);
      }
      
      
      // Loop through unique CONTACT array and assign contact details
      for(n = 0; n < uniqueContactArray.length; n++){      
        if(contacturl == uniqueContactArray[n].url){              
          contactName = uniqueContactArray[n].organisation_name;
        }
        
      }
        
      if(orgName == contactName){
        orgRows.push([taskName, timeslip.dated_on, timeslip.hours, developerFirst, developerLast, contactName, projectName, timeslip.comment, isBillable, billingRate, billingPeriod, status]);
      }

      // create a row for each timeslip
      rows.push([taskName, timeslip.dated_on, timeslip.hours, developerFirst, developerLast, contactName, projectName, timeslip.comment, isBillable, billingRate, billingPeriod, status]);    
            
      
    }     
    // End of Loop to get all timeslip data and push data to rows
    
    createFiles(orgName, rows, orgRows, titles);
   
  }
  // End of Loop - Process contacts
  
  defaultSheet.clear({contentsOnly: true});
 
 
  //titleRange = defaultSheet.getRange(1, 1, titles.length, 12);
  //titleRange.setValues(titles);
  
  // Write data on the spreadsheet  
  //Browser.msgBox("Number of rows updated = "+ rows.length, Browser.Buttons.OK_CANCEL);
  //dataRange = defaultSheet.getRange(2, 1, rows.length, 12);
  //dataRange.setValues(rows);
  
}

function createFiles(orgName, rows, orgRows, titles){
  
  // get Folder or Create Folder named 'Veda Time Sheets'
  var folders = DriveApp.getFolders();
  var getFolder = false;
  while (folders.hasNext()) {
    var folder = folders.next();
     if(folder.getName() == 'Veda Timesheets') {
       var folderId = folder.getId();
       getFolder = true;
       break;
     }
  }
  if (!getFolder) {
    var folder   = DriveApp.createFolder('Veda Timesheets');
    var folderId = folder.getId();
  }
  
  var folderWithId = DriveApp.getFolderById(folderId);
  var files = folderWithId.getFiles();      
  var getFile = false;
  while (files.hasNext()) {    
    var file = files.next();
    if(file.getName() == orgName) {
      var ssNew = SpreadsheetApp.open(file);
      getFile = true;
      break;
    }
  }
  
  if (!getFile) {
    var ssNew = SpreadsheetApp.create(orgName);
    var spreadFile = DriveApp.getFileById(ssNew.getId());
    DriveApp.getFolderById(folderId).addFile(spreadFile);    
  }
  
  //var ssNew = SpreadsheetApp.getActiveSpreadsheet();
  //var sheets = ss.getSheets();  
  var sheets = ssNew.getSheets();
  var sheet = ssNew.getActiveSheet();
  
  var defaultSheet = sheets[0];
  
  if(!(sheets[1] && sheets[1].getName() == 'User Hours')){
    var userHourSheet   = ssNew.insertSheet('User Hours', 1);
  }
  
  if(!(sheets[2] && sheets[2].getName() == 'Task Hours')){
    var taskHourSheet   = ssNew.insertSheet('Task Hours', 2);
  }
  
  if(!(sheets[3] && sheets[3].getName() == 'Client Hours')){
    var clientHourSheet   = ssNew.insertSheet('Client Hours', 3);
  }
  
  defaultSheet  = ssNew.getSheets()[0];
  userHourSheet = ssNew.getSheets()[1];
  taskHourSheet = ssNew.getSheets()[2];
  clientHourSheet = ssNew.getSheets()[3]; 
  
  sheet.clear();
  
  titleRange = sheet.getRange(1, 1, titles.length, 12);
  titleRange.setValues(titles);
  
  //styling the titles
  textFormat(sheet, titleRange);
  
  if(orgName == 'Veda Consulting'){
    dataRange = sheet.getRange(2, 1, rows.length, 12);
    dataRange.setValues(rows);
    //styling the data
    dataRange.setVerticalAlignment("top");
    // format hours into two decimal
    hoursRange = sheet.getRange(2, 3, rows.length, 1);
    twoDecimal(hoursRange);
    //apply colors to cells
    colorme = true;
    tabsData(defaultSheet, userHourSheet, taskHourSheet, clientHourSheet, colorme);
  }
  else{
    dataRange = sheet.getRange(2, 1, orgRows.length, 12);
    dataRange.setValues(orgRows);
    //styling the data
    dataRange.setVerticalAlignment("top");
    // format hours into two decimal
    hoursRange = sheet.getRange(2, 3, rows.length, 1);
    twoDecimal(hoursRange);
    tabsData(defaultSheet, userHourSheet, taskHourSheet, clientHourSheet);
  }
  
}


function getUserDetails(userurl){
  
  var userResponse = UrlFetchApp.fetch(userurl, getUrlFetchPOSTOptions()).getContentText();
  var userArray = JSON.parse(userResponse).user;
  
  // developer details
  var developerFirst = userArray.first_name;      
  var developerLast = userArray.last_name;
  var url = userArray.url;
  
  // add fetched data to a temporary array
  var arrayToPush = {url: url, first_name : developerFirst, last_name : developerLast};
  // push array into the unique user array
  uniqueUserArray.push(userurl, arrayToPush);
  Logger.log("I came here to fetch user :" + userurl);
  return uniqueUserArray;  
}

function getTaskDetails(taskurl){
  
  var taskResponse = UrlFetchApp.fetch(taskurl, getUrlFetchPOSTOptions()).getContentText();
  var taskArray = JSON.parse(taskResponse).task;
  
  // task details
  var taskName = taskArray.name;      
  var isBillable = taskArray.is_billable;
  var billingRate = taskArray.billing_rate;
  var billingPeriod = taskArray.billing_period;
  var status = taskArray.status;
  var url = taskArray.url;
  
  // add fetched data to a temporary array
  var arrayToPush = {url: url, name : taskName, is_billable : isBillable, billing_rate : billingRate, billing_period : billingPeriod, status : status};
  // push array into the unique task array
  uniqueTaskArray.push(taskurl, arrayToPush);
  Logger.log("I came here to fetch Task :" + taskurl);
  return uniqueTaskArray;  
}

function getProjectDetails(projecturl){
  
  var projectResponse = UrlFetchApp.fetch(projecturl, getUrlFetchPOSTOptions()).getContentText();
  var projectArray = JSON.parse(projectResponse).project;
  
  // project details
  var projectName = projectArray.name;      
  var contactUrl = projectArray.contact;
  var url = projectArray.url;
  
  // add fetched data to a temporary array
  var arrayToPush = {url: url, name : projectName, contact : contactUrl};
  // push array into the unique project array
  uniqueProjectArray.push(projecturl, arrayToPush);
  Logger.log("I came here to fetch Project :" + projecturl);
  return uniqueProjectArray;  
}

function getContactDetails(contacturl){
  
  var contactResponse = UrlFetchApp.fetch(contacturl, getUrlFetchPOSTOptions()).getContentText();
  var contactArray = JSON.parse(contactResponse).contact;
  
  // contact details
  var organisationName = contactArray.organisation_name;
  var url = contactArray.url;
  
  // add fetched data to a temporary array
  var arrayToPush = {url: url, organisation_name : organisationName};
  // push array into the unique contact array
  uniqueContactArray.push(contacturl, arrayToPush);
  Logger.log("I came here to fetch Contact :" + contacturl);
  
  return uniqueContactArray;  
}


function tabsData(defaultSheet, userHourSheet, taskHourSheet, clientHourSheet, colorme) {
  
  userHourSheet.clear();
  taskHourSheet.clear();
  clientHourSheet.clear();
  
  // last row number with data
  numRow = defaultSheet.getLastRow();
  
  // Date Column
  dateColumn = "Sheet1!B2:B";
  
  //User Column
  userColumn = "Sheet1!D2:D";
  
  //hours Column
  hourColumn = "Sheet1!C2:C";
  
  //tasks column
  taskColumn = "Sheet1!A2:A";
  
  //client column
  clientColumn = "Sheet1!F2:F"; 
  
  // week Column
  weekColumn = "'User Hours'!A2:A";  
  
  userHours(defaultSheet, userHourSheet, colorme);
  taskHours(taskHourSheet);
  clientHours(clientHourSheet);
  
}


function userHours(defaultSheet, userHourSheet, colorme) {
  
  // Get range to print week number 
  var getWeekRange = userHourSheet.getRange("A2:A"+numRow);
  
  // Print weeknumbers
  getWeekRange.setFormula("=weekNum("+dateColumn+")");
  
  //print unique user names
  userHourSheet.getRange("B2").setFormula("=UNIQUE("+userColumn+")");  
  
  //print unique week numbers
  userHourSheet.getRange("C2").setFormula("=UNIQUE("+weekColumn+")");
  
  // Get all the unique user names
  var usernamesArray = userHourSheet.getRange("B2:B").getValues();  
  // Remove empty value from the user names array
  usernamesArray = usernamesArray.filter(emptyElement);
  
  // Get all the unique week numbers
  var weeknumbersArray = userHourSheet.getRange("C2:C").getValues();
  // Remove empty value from the week number array
  weeknumbersArray = weeknumbersArray.filter(emptyElement);
  
 // sheet2.getRange(8, 9).setFormula("=TRANSPOSE(F2:F)");
  
  // print title for Week Number column
  userHourSheet.getRange(1, 3).setValue("Week No.");
    
  var columnNum,
      username;
  
  for(i=0; i<usernamesArray.length; i++){
    
    username = usernamesArray[i];
    
    columnNum = (4+i);
    
    var rowNum,
        weekNum;
    
    userHourSheet.getRange(1, columnNum).setValue(username);    
    
    for(j=0; j<weeknumbersArray.length; j++){
      
      // Print Title for Users Total hours By Week
      userHourSheet.getRange(1, columnNum+1).setValue("Total Hours");
      
      rowNum = (2+j);
      var uniqueWeekCellNum = j+2;
      
      var uniqueUserCellNum = i+2;
      
        //print total hours for each user
      userHourSheet.getRange(rowNum, columnNum).setFormula("=SUMIFS(("+hourColumn+"),"+weekColumn+",C"+uniqueWeekCellNum+",("+userColumn+"),B"+uniqueUserCellNum+")");
                                                          // =SUMIFS(Sheet1!C2:C,C2:C,C2,B2:B,B2)
      
      //print total hours for each week
      userHourSheet.getRange(rowNum, columnNum+1).setFormula("=SUMIFS(("+hourColumn+"),"+weekColumn+",C"+uniqueWeekCellNum+")");
      
      // Print Start date for Users Total hours By Week
      userHourSheet.getRange(1, columnNum+2).setValue("Start Date");
      userHourSheet.getRange(rowNum, columnNum+2).setFormula("=DATE(2015,1,-2)-WEEKDAY(DATE(2015,1,3))+C"+uniqueWeekCellNum+"*7");
    }                                                      
  }
 
  userHourSheet.hideColumns(1,2);
  
  var rowRange = weeknumbersArray.length;
  var ColRange = usernamesArray.length;
  
  if(colorme == true){
    var range = userHourSheet.getRange(2, 4, rowRange, ColRange);
    colorize(range);
  }
  // format hours into two decimal
  hoursRange = userHourSheet.getRange(2, 4, rowRange, ColRange+1);
  twoDecimal(hoursRange);  
}


function taskHours( taskHourSheet) {
  
  // Get range to print week number 
  var getWeekRange = taskHourSheet.getRange("A2:A"+numRow);
  
  // Print weeknumbers
  getWeekRange.setFormula("=weekNum("+dateColumn+")");
  
  //print unique task names
  taskHourSheet.getRange("B2").setFormula("=UNIQUE("+taskColumn+")");  
  
  //print unique week numbers
  taskHourSheet.getRange("C2").setFormula("=UNIQUE("+weekColumn+")");
  
  // Get all the unique task names
  var tasknamesArray = taskHourSheet.getRange("B2:B").getValues();  
  // Remove empty value from the user names array
  tasknamesArray = tasknamesArray.filter(emptyElement);
  
  // Get all the unique week numbers
  var weeknumbersArray = taskHourSheet.getRange("C2:C").getValues();
  // Remove empty value from the week number array
  weeknumbersArray = weeknumbersArray.filter(emptyElement);
 
  // print title for Week Number column
  taskHourSheet.getRange(1, 3).setValue("Week No.");
    
  var columnNum,
      taskname;
  
  for(i=0; i<tasknamesArray.length; i++){
    
    taskname = tasknamesArray[i];
    
    columnNum = (4+i);
    
    var rowNum,
        weekNum;
    
    taskHourSheet.getRange(1, columnNum).setValue(taskname);    
    
    for(j=0; j<weeknumbersArray.length; j++){
      
       // Print Title for Tasks Total hours By Week
      taskHourSheet.getRange(1, columnNum+1).setValue("Total Hours");
      
      rowNum = (2+j);
      var uniqueWeekCellNum = j+2;
      
      var uniqueTaskCellNum = i+2;
        
      //var lastColumn = sheet.getLastColumn();
      
      // Print total hours by week
      //sheet3.getRange(rowNum, lastColumn).setFormula("=SUMIFS((Sheet1!"+hourColumn+"),"+ weekColumn+ ",C" + uniqueWeekCellNum+ ")");
      
        //print total hours for each task
      taskHourSheet.getRange(rowNum, columnNum).setFormula("=SUMIFS(("+hourColumn+"),"+weekColumn+",C"+uniqueWeekCellNum+",("+taskColumn+"),B"+uniqueTaskCellNum+")");
                                                          // =SUMIFS(Sheet1!C2:C,C2:C,C2,B2:B,B2)
      
      //print total hours for each week
      taskHourSheet.getRange(rowNum, columnNum+1).setFormula("=SUMIFS(("+hourColumn+"),"+weekColumn+",C"+uniqueWeekCellNum+")");
      
      // Print Start date for Users Total hours By Week
      taskHourSheet.getRange(1, columnNum+2).setValue("Start Date");
      taskHourSheet.getRange(rowNum, columnNum+2).setFormula("=DATE(2015,1,-2)-WEEKDAY(DATE(2015,1,3))+C"+uniqueWeekCellNum+"*7");
      
    }   
  }
  taskHourSheet.hideColumns(1,2);
  
  var rowRange = weeknumbersArray.length;
  var ColRange = tasknamesArray.length;
  
  // format hours into two decimal
  hoursRange = taskHourSheet.getRange(2, 4, rowRange, ColRange+1);
  twoDecimal(hoursRange);
}


function clientHours(clientHourSheet) {
  
  // Get range to print week number 
  var getWeekRange = clientHourSheet.getRange("A2:A"+numRow);
  
  // Print weeknumbers
  getWeekRange.setFormula("=weekNum("+dateColumn+")");
  
  //print unique task names
  clientHourSheet.getRange("B2").setFormula("=UNIQUE("+clientColumn+")");  
  
  //print unique week numbers
  clientHourSheet.getRange("C2").setFormula("=UNIQUE("+weekColumn+")");
  
  // Get all the unique task names
  var clientnamesArray = clientHourSheet.getRange("B2:B").getValues();  
  // Remove empty value from the user names array
  clientnamesArray = clientnamesArray.filter(emptyElement);
  
  // Get all the unique week numbers
  var weeknumbersArray = clientHourSheet.getRange("C2:C").getValues();
  // Remove empty value from the week number array
  weeknumbersArray = weeknumbersArray.filter(emptyElement);
 
  // print title for Week Number column
  clientHourSheet.getRange(1, 3).setValue("Week No.");
    
  var columnNum,
      clientname;
  
  for(i=0; i<clientnamesArray.length; i++){
    
    clientname = clientnamesArray[i];
    
    columnNum = (4+i);
    
    var rowNum,
        weekNum;
    
    clientHourSheet.getRange(1, columnNum).setValue(clientname);    
    
    for(j=0; j<weeknumbersArray.length; j++){
      
      // Print Title for Tasks Total hours By Week
      clientHourSheet.getRange(1, columnNum+1).setValue("Total Hours");
      
      rowNum = (2+j);
      var uniqueWeekCellNum = j+2;
      
      var uniqueClientCellNum = i+2;
      
        //print total hours for each client
      clientHourSheet.getRange(rowNum, columnNum).setFormula("=SUMIFS(("+hourColumn+"),"+weekColumn+",C"+uniqueWeekCellNum+",("+clientColumn+"),B"+uniqueClientCellNum+")");
                                                          // =SUMIFS(Sheet1!C2:C,C2:C,C2,B2:B,B2)
      
      //print total hours for each week
      clientHourSheet.getRange(rowNum, columnNum+1).setFormula("=SUMIFS(("+hourColumn+"),"+weekColumn+",C"+uniqueWeekCellNum+")");
      
      // Print Start date for Users Total hours By Week
      clientHourSheet.getRange(1, columnNum+2).setValue("Start Date");
      clientHourSheet.getRange(rowNum, columnNum+2).setFormula("=DATE(2015,1,-2)-WEEKDAY(DATE(2015,1,3))+C"+uniqueWeekCellNum+"*7");
            
    }   
  }
  clientHourSheet.hideColumns(1,2);
  
  var rowRange = weeknumbersArray.length;
  var ColRange = clientnamesArray.length;
  
  // format hours into two decimal
  hoursRange = clientHourSheet.getRange(2, 4, rowRange, ColRange+1);
  twoDecimal(hoursRange);
  
}

// remove empty values from an array
function emptyElement(element) {
  //Removes nulls and blank element
  if (element == null || element == ''){
    return false;
  }
  else{
    return true;
  }
}


// Make hours two decimal
function twoDecimal(hoursRange){
  hoursRange.setNumberFormat("0.00");
}


function uploadData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var contactDataRange = ss.getDataRange();
  var contactObjects = getRowsData(sheet, contactDataRange,1);
  var runningLog = '<br>Uploaded following:<br><br>';
  for(var i = 1;i<contactObjects.length;i++){
    var payload=  Utilities.jsonStringify(
      {"FirstName" : contactObjects[i].firstname,
       "LastName" : contactObjects[i].lastname,
       "Email" : contactObjects[i].email,
       "Phone" : contactObjects[i].phone
      }
    );
    Logger.log('trying ' + payload);
    var getDataURL = UserProperties.getProperty(baseURLPropertyName) + '/services/data/v26.0/sobjects/Contact/';
    runningLog += UrlFetchApp.fetch(getDataURL,getUrlFetchPOSTOptions(payload)).getContentText() + '<br>';  
  }
  return runningLog;
}


//this is the URL where they'll authorize with freeagent.com
//may need to add a "scope" param here. like &scope=full for freeagent
function getURLForAuthorization(){
  return AUTHORIZE_URL + '?response_type=code&client_id='+CLIENT_ID+'&redirect_uri='+REDIRECT_URL
}

function getAndStoreAccessToken(code){
  //var nextURL = TOKEN_URL + '?client_id='+CLIENT_ID+'&client_secret='+CLIENT_SECRET+'&grant_type=authorization_code&redirect_uri='+REDIRECT_URL+'&code=' + code;
  var nextURL = TOKEN_URL;
  var payload = 'grant_type=authorization_code&redirect_uri='+REDIRECT_URL+'&code=' + code;


  var response = UrlFetchApp.fetch(nextURL, getUrlFetchPOSTAuthOptions(payload)).getContentText();   
  var tokenResponse = JSON.parse(response);

  //freeagent requires you to call against the instance URL that is against the token (eg. https://dev.freeagent.com/)
  UserProperties.setProperty(baseURLPropertyName, tokenResponse.instance_url);
  //store the token for later retrival
  UserProperties.setProperty(tokenPropertyName, tokenResponse.access_token);
}


function getUrlFetchOptions() {
  var token = UserProperties.getProperty(tokenPropertyName);
  return {
    "contentType" : "application/json",
    "headers" : {
      "Authorization" : "Bearer " + token,
      "Accept" : "application/json"
    }
  };
}

function getUrlFetchPOSTAuthOptions(payload){
  var encoded = Utilities.base64Encode(CLIENT_ID+':'+CLIENT_SECRET);
  return {
    "method": "post",
    "accept" : "application/json",
    "contentType" : "application/x-www-form-urlencoded",
    "payload" : payload,
    "headers" : {
      "Authorization" : "Basic " + encoded
    }
  }
}

function getUrlFetchPOSTOptions(payload){
  var token = UserProperties.getProperty(tokenPropertyName);
  return {
    "method": "GET",
    "contentType" : "application/json",
    "payload" : payload,
    "headers" : {
      "Authorization" : "Bearer " + token
    }
  }
}

function isTokenValid() {
  var token = UserProperties.getProperty(tokenPropertyName);
  //adddatatosheet(token, 'token');
  if(!token){ //if its empty or undefined
    return false;
  }
  return true; //naive check
}

// colorize cells
function colorize(range) {
  var values = range.getValues();
    //Logger.log('values:' + values);
  
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  for (var i = 1; i <= numRows; i++) {
    for (var j = 1; j <= numCols; j++) {
          
      var currentValue = range.getCell(i,j).getValue();
      if(currentValue!== ""){
        if(currentValue>=minGreen){
         range.getCell(i,j).setBackgroundRGB(101,250,29);
        }
        else if(currentValue<maxAmber && currentValue >= minAmber){
         range.getCell(i,j).setBackgroundRGB(250,219,15);
        }
        else if(currentValue<maxBlue && currentValue >= minBlue){
         range.getCell(i,j).setBackgroundRGB(38,100,131);
        }
        else if (currentValue<maxRed && currentValue >= minRed){
         range.getCell(i,j).setBackgroundRGB(219,44,55);
        }
      }
    }
  }     
}


function textFormat(sheet, titleRange){
  titleRange.setFontSize(titleFontSize);
  titleRange.setFontFamily(titleFontFamily);
  titleRange.setFontWeight(titleFontWeight);
  titleRange.setFontColor(titleFontColor);
  titleRange.setHorizontalAlignment(titleHorizontalAlignment);
  titleRange.setVerticalAlignment(titleVerticalAlignment);
  sheet.setRowHeight(1, titleRowHeight);  
  sheet.getRange('D1:E1').merge();
  sheet.setFrozenRows(1);
}