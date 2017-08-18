// Calendar Add Web App

/*
* The intention of this script is to create a web app that allows users to create calendar events based on the school timetable (6-day cycle, etc.)
*
* A spreadsheet should be used in connection with this script to set defaults and settings
*
* - achiu@ais.edu.hk 8 Aug 2014
*/

// -- GLOBALS --
// Spreadsheet details - for settings and DB details stored in a spreadsheet
var mySheetID = "1X4TIWl6MxgoptgX-7FuANYjxzydokSmab3OZbY8rz6E"; // Calendar Settings Spreadsheet 2014
var sheetTeacherCourses = "TeacherCourses" //Name of sheet with details of the courses each teacher teaches


//Get User
var thisUser = Session.getActiveUser().getEmail();
//var thisUser = 'jwade@ais.edu.hk';

//High School Calendar
//var referenceCalendar = 'ais.edu.hk_l4fben0lb5jblvd8q6b0501ijc@group.calendar.google.com';


// -- FUNCTIONS --
//
// -----
//  doGet - main function for web app
// -----
function doGet(){
  var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
  
  var myDoc = 'index';  
  return HtmlService.createTemplateFromFile(myDoc).evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}
  
//-----
// loadGInfo - loads up at beginning after html has been loaded
//-----
function loadGInfo(){
  var mySheet = SpreadsheetApp.openById(mySheetID);

  //get user courses
  var teacherCourseSheet = mySheet.getSheetByName(sheetTeacherCourses);
  var myCoursesDb = getRowsMatching(teacherCourseSheet.getRange(1, 1, teacherCourseSheet.getLastRow(), teacherCourseSheet.getLastColumn()).getValues(),0,thisUser);
  
  var myCourses = {};
  for( var course = 0; course < myCoursesDb.length; course++){
    var myPeriodCode = myCoursesDb[course][4];
    if (myPeriodCode != ""){ myCourses[myPeriodCode] = {subject: myCoursesDb[course][1], location:myCoursesDb[course][5]};}
  }
  
  //Get available timetable templates
  var tempSheets = mySheet.getSheets();
  var templateNames = new Array();
  for (var i = 0; i < tempSheets.length; i++){//cycle through all sheets. 
    var thisName = tempSheets[i].getName();
    var thisTemplateActive = false;

    if (thisName.substring(0,9) == "Timetable") { //Any sheet starting with "Timetable", capture table details into JS object
      var thisTable = tempSheets[i].getDataRange().getValues();
      var thisNumOfDays = thisTable[0].length - 4; //get number of days in this template
      var thisRefCalId = thisTable[0][0]; //get the refCalendar ID.
      var thisDays = new Array();      
      for (var day=4; day < thisNumOfDays + 4; day++){ //for each day
        var dayActive = false;
        var thisRows = new Array();         
        for( var r=1; r < thisTable.length; r++){ //for each row
          var subject = "";
          var location = "";
          var active = false;
          var shareWith = "";
          var thisPCode = thisTable[r][day];
          
          if (thisPCode in myCourses){ //If the period code is in myCourses period code
            subject = myCourses[thisPCode].subject;
            location = myCourses[thisPCode].location;
            active = true;
            dayActive = true;
            thisTemplateActive = true;
          }
          
          thisRows.push({
                          title: thisTable[r][1],
                          startTime: thisTable[r][2],
                          endTime: thisTable[r][3],
                          periodCode: thisPCode,
                          subject: subject,
                          location: location,
                          active: active,
                          shareWith: shareWith
                          });
        } //Basics for each row
        thisDays.push({
                        dayName: thisTable[0][day],
                        rows: thisRows,
                        active: dayActive});
      }
      templateNames.push({name: thisName.substring(9,thisName.length),
                           templateActive: thisTemplateActive,
                           refCalId: thisRefCalId,
                           days: thisDays
                           });      
    }
  }

  //Get Script Properties for Start and End Dates
  var sP = PropertiesService.getScriptProperties();
  var thisStartDate = sP.getProperty('yrStartDate');
  var thisEndDate = sP.getProperty('yrEndDate');
  
  //Get user's calendars
  var myCalendars = [];
  var myCals = CalendarApp.getAllOwnedCalendars();
  var myDefCal = CalendarApp.getDefaultCalendar().getId();
  for(var c = 0; c < myCals.length; c++){
   myCalendars.push({name: myCals[c].getName(), id: myCals[c].getId()});
  }
  myCalendars.push({name: "-- CREATE A NEW CALENDAR --", id:0});
  
  
  return {fromDate: thisStartDate, toDate: thisEndDate, templates: templateNames, myCalendars: myCalendars, myDefCal: myDefCal };
}


//-----
// loadme - loads up at beginning after loading image has been loaded
//-----
function loadme(e){
  var app = UiApp.getActiveApplication();

  //Create Main Page
  app.add(app.createHTML('<h1>AIS Google Timetable Creator</h1>'));
  
  var instructionsStr = '<p>Use this app to create your timetable to your google calendar.</p>';
  instructionsStr += '<p>Set the range of dates you wish to create events from / to using Start Date and End Date</p>';
  instructionsStr += '<p>It is recommended that you first create a calendar in Google calendars to try this out. Refresh this page once you have created a calendar to see it show up in the calendar list below</p>';
  
  app.add(app.createHTML(instructionsStr)); //Instructions blurb
  
  
  var BigVPage = app.createVerticalPanel().setId("BigVPage"); //Will hold all widgets to return on form on the page
  
  

  var mySheet = SpreadsheetApp.openById(mySheetID);
  
  //Get available timetable templates
  var tempSheets = mySheet.getSheets();
  var templateNames = new Array();
  var templatePicker = app.createListBox().setId('templatePicker').setName('templatePicker').addChangeHandler(app.createServerHandler('changeTemplate').addCallbackElement(BigVPage)).addChangeHandler(loadingHandler(app));
  var templatePanel = app.createHorizontalPanel();
  templatePanel.add(app.createHTML('<h3>Template</h3>'));
  templatePanel.add(templatePicker);
  app.add(templatePanel);
  app.add(BigVPage);
  for (var i = 0; i < tempSheets.length; i++){
    var thisName = tempSheets[i].getName();

    if (thisName.substring(0,9) == "Timetable") {
      templatePicker.addItem(thisName.substring(9,thisName.length),thisName);
      templateNames.push([thisName.substring(9,thisName.length),thisName]);      
    }
  }
  

  //BigVPage HERE
  createBigVPage(app, mySheet, templateNames[0][1], BigVPage);
  
  //app.add(BigVPage);
  
  app.add(app.createHTML('<p><i>created by achiu@ais.edu.hk - Aug 2014...</i></p>')); //leave contact detail
  
  loadingFinished(app); //Hide loading image
  return app;  
}


//------
// function createBigVPage
//------
function createBigVPage(app, mySheet, templateSheet, BigVPage){
 
  
  //Get Courses of current user
  var teacherCourseSheet = mySheet.getSheetByName(sheetTeacherCourses);
  var myCourses = getRowsMatching(teacherCourseSheet.getRange(1, 1, teacherCourseSheet.getLastRow(), teacherCourseSheet.getLastColumn()).getValues(),0,thisUser);
  Logger.log(myCourses);
  
  //Get Script Properties for Start and End Dates
  var sP = PropertiesService.getScriptProperties();
  var thisStartDate = sP.getProperty('yrStartDate');
  var thisEndDate = sP.getProperty('yrEndDate');
  
  //Start & End Dates
  var seDatesGrid = app.createGrid(2,2);
  var startDateLbl = app.createLabel("Start Date");
  var endDateLbl = app.createLabel("End Date");
  var startDateTxt = app.createDateBox().setId("startDateTxt").setName("startDateTxt").setValue(new Date(thisStartDate));
  var endDateTxt = app.createDateBox().setId("endDateTxt").setName("endDateTxt").setValue(new Date(thisEndDate));
  seDatesGrid.setWidget(0,0,startDateLbl);
  seDatesGrid.setWidget(0,1,startDateTxt);
  seDatesGrid.setWidget(1,0,endDateLbl);
  seDatesGrid.setWidget(1,1,endDateTxt);
  
  BigVPage.add(seDatesGrid);
  
  
  //create buttons for each template timetable
  //var templateSheet = templateNames[0][1]
  
  //Timetable container
  var myTabPanel = app.createTabPanel().setId("myTabPanel");
 
  
  //get timetable info from sheet
  var myTimetableDataRange = mySheet.getSheetByName(templateSheet).getDataRange();
  var myTimetableData = myTimetableDataRange.getValues();
  var numOfDays = myTimetableDataRange.getLastColumn() - 4;
  var numOfTimes = myTimetableDataRange.getLastRow();
  
  BigVPage.add(app.createHidden('totalDays').setValue(numOfDays));
  BigVPage.add(app.createHidden('totalTimes').setValue(numOfTimes));
  BigVPage.add(app.createHidden('refCal').setValue(myTimetableData[0][0]));
  Logger.log([numOfDays,numOfTimes, myTimetableData[0][0]]);
  Logger.log(myTimetableData);

  

  
  //Create tabPanel
    //For each column day
      
  for (var day = 1; day <= numOfDays; day++){
    var dayLabel = myTimetableData[0][3+day];
    BigVPage.add(app.createHidden('d'+day).setValue(dayLabel).setId('d'+day)); //to let the handler pick up the name of the Day to search.
    var dayGrid = app.createGrid(numOfTimes+1, 8); //Create grid to put in panel
    //for each row
    dayGrid.setWidget(0, 0, app.createHTML('<p>Period</p>'));
    dayGrid.setWidget(0, 1, app.createHTML('<p>Start</p>'));
    dayGrid.setWidget(0, 2, app.createHTML('<p>End</p>'));
    dayGrid.setWidget(0, 3, app.createHTML('<p>Active</p>'));
    dayGrid.setWidget(0, 4, app.createHTML('<p>Subject</p>'));
    dayGrid.setWidget(0, 5, app.createHTML('<p>Location</p>'));
    //dayGrid.setWidget(0, 6, app.createHTML('<p>Status</p>'));
    for (var timeslot = 1; timeslot < numOfTimes; timeslot++){
      
        //add start & end time text boxes, label
      dayGrid.setWidget(timeslot, 0, app.createLabel(myTimetableData[timeslot][1])); //Label
      dayGrid.setWidget(timeslot, 1, app.createTextBox().setValue(myTimetableData[timeslot][2]).setId('st'+'d'+day+'t'+timeslot).setName('st'+'d'+day+'t'+timeslot).setWidth(timeboxWidth));  //Start time
      dayGrid.setWidget(timeslot, 2, app.createTextBox().setValue(myTimetableData[timeslot][3]).setId('et'+'d'+day+'t'+timeslot).setName('et'+'d'+day+'t'+timeslot).setWidth(timeboxWidth)); //End time

      
      //Subject
      var mySbj = app.createTextBox().setId('sbj'+'d'+day+'t'+timeslot).setName('sbj'+'d'+day+'t'+timeslot);
      var courseFound = false;
      var findCourse = getRowsMatching(myCourses, 4, myTimetableData[timeslot][day+3]);
      Logger.log(findCourse);
      if(findCourse.length > 0){
        courseFound = true;
        mySbj.setValue(findCourse[0][1]);
      }                
      
      dayGrid.setWidget(timeslot, 4, mySbj );//Title
      dayGrid.setWidget(timeslot, 5, app.createTextBox().setId('loc'+'d'+day+'t'+timeslot).setName('loc'+'d'+day+'t'+timeslot));//Location
  
      dayGrid.setWidget(timeslot, 3,  app.createCheckBox().setId('active'+'d'+day+'t'+timeslot).setName('active'+'d'+day+'t'+timeslot).setValue(courseFound));//Active
     
      /* Cannot set Free busy data using Calendar service?
      //Free Busy
      var freeBusy = app.createListBox().addItem('Free').addItem('Busy');                  
      //if (courseFound) freeBusy.setValue(1, 'Busy');
      dayGrid.setWidget(timeslot, 6, freeBusy); //Free Busy
      
        //create textbox & fill with default if course taught
        //create busy / free and set to busy if textbox is filled
      */
      
      myTabPanel.add(dayGrid, dayLabel);
    }

  }
  myTabPanel.selectTab(0);
  //Add tabs to page  
  BigVPage.add(myTabPanel);

  //Choice for recurring or individual events
  var pickRecurring = app.createCheckBox('Create Recurring Events').setId('Recurring').setName('Recurring');
  BigVPage.add(pickRecurring);
  
  //Get Calendar list and create a Calender dropdown list box.
  var pickCal = app.createListBox().setName('targetCal').setId('targetCal');
  var myCals = CalendarApp.getAllOwnedCalendars();
  for(var c = 0; c < myCals.length; c++){
   pickCal.addItem(myCals[c].getName(), myCals[c].getId());
  }
  pickCal.setItemSelected(0, true);
  
  var calPanel = app.createHorizontalPanel();
  calPanel.add(app.createHTML('<p>Add to calendar:</p>'));
  calPanel.add(pickCal);
  
  BigVPage.add(calPanel);
  
  //Event Handler and Button
  var goDoItH = app.createServerHandler('addAllEvents').addCallbackElement(BigVPage);
  var goDoIt = app.createButton('Add to my Calendar', goDoItH)
  var disabler = app.createClientHandler().forTargets([goDoIt]).setEnabled(false);
  BigVPage.add(goDoIt.setId('goDoItBtn').addClickHandler(loadingHandler(app)).addClickHandler(disabler));

    //Event Handler and Button for spreadsheet
  var goDoItH2 = app.createServerHandler('addAllEventsToSS').addCallbackElement(BigVPage);
  var goDoIt2 = app.createButton('Create Spreadsheet with My Events', goDoItH2)
  var disabler2 = app.createClientHandler().forTargets([goDoIt2]).setEnabled(false);
  BigVPage.add(goDoIt2.setId('goDoItBtn2').addClickHandler(loadingHandler(app)).addClickHandler(disabler2));

}

//-----
// loadingFinished - hides the loading image
//-----
function loadingFinished(app) {

  var image = app.getElementById('loadImage');
  image.setVisible(false);
  var btn = app.getElementById('goDoItBtn');
  btn.setEnabled(true);

}

//-----
// loadingHandler - creates and returns a handler for the loading image to appear
//-----
function loadingHandler(app){
  var image = app.getElementById('loadImage');
  var lHandler = app.createClientHandler().forTargets([image]).setVisible(true);
  return lHandler;
}


// -----
// function saveToCal - adds all the events to the user's calendar
// -----
function saveToCal(e) {
  var finalSheetID = "";
  //check target calendar. If 0 then make new calendar
  var targetCalId = e.targetCal;
  if (targetCalId == 0){ //if chose to create a new calendar, create one and get id.
    targetCalId = CalendarApp.createCalendar('New Calendar by AIS Google Timetable Creator').getId();
  }
  var targetCalendar = CalendarApp.getCalendarById(targetCalId);
  var startDate = new Date(e.startDate);
  var endDate = new Date(e.endDate);
  
  var myCals = [["Cal ID", "Title", "Start", "End", "Options", "event ID"]];

  for (var i=0; i < e.days.length; i++){ //for each day to search
    var refCal = CalendarApp.getCalendarById(e.days[i].refCal);
    var dayLabel = e.days[i].search
    var foundEvents = refCal.getEvents(startDate, endDate, {search: dayLabel}); //search for events on the ref calendar
    for (var evt = 0; evt < foundEvents.length; evt++){ //for each event found
     for (var s=0; s < e.days[i].events.length; s++){           
       var thisStartDate = new Date(foundEvents[evt].getStartTime());
       thisStartDate.setHours(e.days[i].events[s].startTimeHours);
       thisStartDate.setMinutes(e.days[i].events[s].startTimeMins);
       var thisEndDate = new Date (foundEvents[evt].getStartTime());
       thisEndDate.setHours(e.days[i].events[s].endTimeHours);
       thisEndDate.setMinutes(e.days[i].events[s].endTimeMins);
       myCals.push([targetCalId, e.days[i].events[s].title, thisStartDate, thisEndDate, e.days[i].events[s].options, ""]);
       

     }
    }

  }
  
  var myNewSs = SpreadsheetApp.create('Calendar Events by AIS Google Timetable Creator');
  finalSheetID = myNewSs.getId();
  var myRecordSheet = myNewSs.insertSheet('Created Events');
  myNewSs.getSheets()[0].getRange(1,1,myCals.length, myCals[0].length).setValues(myCals);
  var savedRows = [];
  for (var row = 1; row < myCals.length; row++){
    if((row%18 == 0) && (savedRows.length > 0)) {
      try{
        myRecordSheet.getRange(myRecordSheet.getLastRow()+1, 1, savedRows.length, savedRows[0].length).setValues(savedRows); //write the current progress
        savedRows = [];
      }catch(err)  {
        Logger.log(err);
        Logger.log(savedRows);
        return err
      }
      Utilities.sleep(1900);
    }
    //Logger.log([e.days[i].events[s].title, thisStartDate, thisEndDate, e.days[i].events[s].options]);
    var myEventID = targetCalendar.createEvent(myCals[row][1], myCals[row][2], myCals[row][3], myCals[row][4]).getId();
    savedRows.push([myCals[row][0],myCals[row][1],myCals[row][2],myCals[row][3],myCals[row][4], myEventID] );
  }
  myRecordSheet.getRange(myRecordSheet.getLastRow(), 1, savedRows.length, savedRows[0].length).setValues(savedRows); //write the last records
  
  return finalSheetID;
}

// -----
// function addAllEvents - adds all the events to the user's calendar
// -----
function addAllEvents(e) {

  var app = UiApp.getActiveApplication();
  
  var numOfDays = e.parameter.totalDays;
  var numOfTimes = e.parameter.totalTimes - 1;
  var startDate = e.parameter.startDateTxt;
  var endDate = e.parameter.endDateTxt;
  
  var referenceCalendar = e.parameter.refCal;
  var Recurring = e.parameter.Recurring;
  
  var refCal = CalendarApp.getCalendarById(referenceCalendar);
  var targetCalendar = CalendarApp.getCalendarById(e.parameter.targetCal);
  
  Logger.log(targetCalendar.getName());
  
  //Collect all active event details 
  
  for (var d = 1; d <= numOfDays; d++){
    var dayLabel = e.parameter['d'+d.toString()];
    Logger.log(dayLabel);
    
    var foundEvents = refCal.getEvents(startDate, endDate, {search: dayLabel});
    
    Logger.log(foundEvents.length);
    
    for (var slots = 1; slots <= numOfTimes; slots++){
      var tail = 'd' + d + 't' + slots;
      //Logger.log(tail);
      var thisActive = e.parameter['active'+tail];
      if (thisActive == 'true'){
        var mySbj = e.parameter['sbj'+tail];
        var myLoc = e.parameter['loc'+tail];
        var myStart = new Date(new Date().toDateString() + ' ' + e.parameter['st'+tail]);
        var myEnd = new Date(new Date().toDateString() + ' ' + e.parameter['et'+tail]);
        
       
        if (Recurring == 'true'){
         var myRecurrence = CalendarApp.newRecurrence()
         //Create recurrence from all 2nd occurrence onwards 
         for (var evt = 0; evt < foundEvents.length; evt++){
           // Create new Date variables based on any given day A - F
           myRecurrence.addDate(new Date(foundEvents[evt].getStartTime().toDateString()));
         }
            var thisStartDate = new Date(foundEvents[0].getStartTime());
            thisStartDate.setHours(myStart.getHours());
            thisStartDate.setMinutes(myStart.getMinutes());
            var thisEndDate = new Date (foundEvents[0].getStartTime());
            thisEndDate.setHours(myEnd.getHours());
            thisEndDate.setMinutes(myEnd.getMinutes());
            targetCalendar.createEventSeries(mySbj, thisStartDate, thisEndDate, myRecurrence);
        } else {
        //Recurring: NO (Create Individual Events)
       
          for (var evt = 0; evt < foundEvents.length; evt++){
            
            var thisStartDate = new Date(foundEvents[evt].getStartTime());
            thisStartDate.setHours(myStart.getHours());
            thisStartDate.setMinutes(myStart.getMinutes());
            var thisEndDate = new Date (foundEvents[evt].getStartTime());
            thisEndDate.setHours(myEnd.getHours());
            thisEndDate.setMinutes(myEnd.getMinutes());
            if(evt%15 == 0) Utilities.sleep(2500);
            targetCalendar.createEvent(mySbj, thisStartDate, thisEndDate, {location: myLoc});
            Logger.log([thisActive, mySbj, myLoc, thisStartDate, thisEndDate]);
          }
        }
          
 
      }
    }
  }
    
  loadingFinished(app); //Hide loading image
  return app;
}


function changeTemplate(e){
  var app = UiApp.getActiveApplication();
  
  var templateSheet = e.parameter.templatePicker;
  var BigVPage = app.getElementById('BigVPage');
  
  var mySheet = SpreadsheetApp.openById(mySheetID);
  
  BigVPage.clear();
  
  createBigVPage(app, mySheet, templateSheet, BigVPage);
  loadingFinished(app); //Hide loading image
  
  return app;
}

function myClickHandler(e) {
  var app = UiApp.getActiveApplication();

  var label = app.getElementById('statusLabel');
  label.setVisible(true);

  app.close();
  return app;
}

// function addAllEventsToSS - adds all the events to a new Spreadsheet
// -----
function addAllEventsToSS(e) {

  var app = UiApp.getActiveApplication();
  
  var numOfDays = e.parameter.totalDays;
  var numOfTimes = e.parameter.totalTimes - 1;
  var startDate = e.parameter.startDateTxt;
  var endDate = e.parameter.endDateTxt;
  
  var referenceCalendar = e.parameter.refCal;
  var Recurring = e.parameter.Recurring;
  
  var refCal = CalendarApp.getCalendarById(referenceCalendar);
  //var targetCalendar = CalendarApp.getCalendarById(e.parameter.targetCal);
  var targetCalendar = SpreadsheetApp.create('New Calendar Events Output');
 
  var targetCalendarInput = new Array();
  targetCalendarInput.push(["Day", "Month", "Hour", "Minutes", "Subject", "year"]);
//  targetCalendarInput.push(["Subject2", "Timestart", "TimeEnd", "Location"]);

//    Logger.log(targetCalendarInput);
 // targetCalendar.getActiveSheet().getRange(1, 1, targetCalendarInput.length, 4).setValues(targetCalendarInput);
  

  
  //Collect all active event details 
  
  for (var d = 1; d <= numOfDays; d++){
    var dayLabel = e.parameter['d'+d.toString()];
    Logger.log(dayLabel);
    
    var foundEvents = refCal.getEvents(startDate, endDate, {search: dayLabel});
    
    Logger.log(foundEvents.length);
    
    for (var slots = 1; slots <= numOfTimes; slots++){
      var tail = 'd' + d + 't' + slots;
      //Logger.log(tail);
      var thisActive = e.parameter['active'+tail];
      if (thisActive == 'true'){
        var mySbj = e.parameter['sbj'+tail];
        var myLoc = e.parameter['loc'+tail];
        var myStart = new Date(new Date().toDateString() + ' ' + e.parameter['st'+tail]);
        var myEnd = new Date(new Date().toDateString() + ' ' + e.parameter['et'+tail]);
        

       
        if (Recurring == 'true'){
         var myRecurrence = CalendarApp.newRecurrence()
         //Create recurrence from all 2nd occurrence onwards 
         for (var evt = 0; evt < foundEvents.length; evt++){
           // Create new Date variables based on any given day A - F
           myRecurrence.addDate(new Date(foundEvents[evt].getStartTime().toDateString()));
         }
            var thisStartDate = new Date(foundEvents[0].getStartTime());
            thisStartDate.setHours(myStart.getHours());
            thisStartDate.setMinutes(myStart.getMinutes());
            var thisEndDate = new Date (foundEvents[0].getStartTime());
            thisEndDate.setHours(myEnd.getHours());
            thisEndDate.setMinutes(myEnd.getMinutes());
            targetCalendarInput.push([mySbj, thisStartDate, thisEndDate, myRecurrence]);
        } else {
        //Recurring: NO (Create Individual Events)
       
          for (var evt = 0; evt < foundEvents.length; evt++){
            
            var thisStartDate = new Date(foundEvents[evt].getStartTime());
            thisStartDate.setHours(myStart.getHours());
            thisStartDate.setMinutes(myStart.getMinutes());
            /*
            var thisEndDate = new Date (foundEvents[evt].getStartTime());
            thisEndDate.setHours(myEnd.getHours());
            thisEndDate.setMinutes(myEnd.getMinutes());
            */
            
            targetCalendarInput.push([thisStartDate.getDate(), thisStartDate.getMonth(), thisStartDate.getHours(), thisStartDate.getMinutes(), mySbj, thisStartDate.getYear()]);
            //Logger.log([thisActive, mySbj, myLoc, thisStartDate, thisEndDate]);
          }
        }
          
 
      }
    }
  }
  Logger.clear();
  Logger.log(targetCalendarInput);
 targetCalendar.getActiveSheet().getRange(1, 1, targetCalendarInput.length,6).setValues(targetCalendarInput); 
  loadingFinished(app); //Hide loading image
  return app;
}


//getRowsMatching takes a data list and searches the sortIndex for all values that match valueToFind, returning the rows that match this value

function getRowsMatching(myDataList, sortIndex, valueToFind){
  
  var foundList = new Array();  
  myDataList.sort(function(a, b){ //Sort the items by studentID
    var x = a[sortIndex];
    var y = b[sortIndex];
    return (x < y ? -1 : (x > y ? 1 : 0));});
  var cdr = 0;
  var found = false; 
  while ( cdr < myDataList.length){
    if (myDataList[cdr][sortIndex] == valueToFind) {
      found=true;
      foundList.push(myDataList[cdr])
    }
    else if (found){
      return foundList;
    }
    cdr++;
  }
  return foundList;
  
}


// -----
// include - include files
// -----
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}



function testTTD(){
  var now = new Date();
  var time = "13:50";
  var ttd = timeToDate(time,now);
  Logger.log(ttd);
}


// timeToDate
function timeToDate(tStr,ddate) {
	var now = ddate;
  now.setHours(tStr.substr(0,tStr.indexOf(":")));
  now.setMinutes(tStr.substr(tStr.indexOf(":")+1));
  now.setSeconds(0);
  return now;

}