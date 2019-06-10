function scheduleShifts() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();
  var calendarId = spreadsheet.getRange("AN3").getValue();
  var eventCal = CalendarApp.getCalendarById(calendarId);
  
  var event = spreadsheet.getRange("Q4:Q366").getValues();
  var startTime = spreadsheet.getRange("AK4:AK366").getValues();
  var endTime = spreadsheet.getRange("AM4:AM366").getValues();
  
  var clear = eventCal.getEvents(new Date('01/01/2019 00:00:00'), new Date('31/12/2019 23:59:59'));
  Logger.log('Number of events: ' + clear.length);
  var arrayLength = clear.length;
  for (var i=0; i<arrayLength; i++) {
    clear[i].deleteEvent();
  }
  
  for (x=0; x<event.length; x++) {
    if (event[x] != '') {
      eventCal.createEvent(event[x], new Date(startTime[x]), new Date(endTime[x]))
    }
  }
}

function clearCalendar() {
  var clear = eventCal.getEvents(new Date('01/01/2019 00:00:00'), new Date('31/12/2019 23:59:59'));
  Logger.log('Number of events: ' + clear.length);
  var arrayLength = clear.length;
  for (var i=0; i<arrayLength; i++) {
    clear[i].deleteEvent();
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Calendar Sync")
  .addItem("Sync Shifts To Calendar", 'scheduleShifts')
  .addItem("Clear Calendar", 'clearCalendar')
  .addToUi();
}
