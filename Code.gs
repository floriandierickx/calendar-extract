// Notes documentation:
// function export_gcal_to_gsheet is ran every hour

function export_gcal_to_gsheet(){

// variables
 
var mycal = "1si69sarlsi2k2orlnij0423jk@group.calendar.google.com";
var cal = CalendarApp.getCalendarById(mycal);
var events = cal.getEvents(new Date("February 1, 2020 00:00:00 CST"), new Date("December 30, 2025 23:59:59 CST"), {search: ''});
console.log({message: 'Number of events', initialData: events});

//var splitEventId = events.getId().split('@');  
//var eventURL = "https://www.google.com/calendar/event?eid=" + Utilities.base64Encode(splitEventId[i] + " " + mycal).toString().replace('=','');
//var eventURL = "https://www.google.com/calendar/event?eid=" + Utilities.base64Encode(events.getId().split('@') + " " + mycal).toString().replace('=','');
  
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Import from Google iCal");
var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event", "iCalUID"]]
var range = sheet.getRange(1,1,1,15);
range.setValues(header);

// for-loop
 
for (var i=0;i<events.length;i++) {

var eventURL = "https://www.google.com/calendar/event?eid=" + Utilities.base64Encode((events[i].getId().split('@') + " " + mycal).toString());
var row=i+2;
var myformula_placeholder = '';
var details=[[mycal,events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent(), eventURL]];
var range=sheet.getRange(row,1,1,15);
range.setValues(details);
var cell=sheet.getRange(row,7);
cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
cell.setNumberFormat('.00');

}};




function addressToPosition() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Import from Google iCal");
  // sheet.getRange("R491:V1000").clearContent();
  var cells = sheet.getRange("Q491:V1000");
  // var cells = sheet.getActiveRange();
  
  // Must have selected 3 columns (Address, Lat, Lng).
  // Must have selected at least 1 row.

  if (cells.getNumColumns() != (6)) {
    Logger.log("Must select at least 3 columns: Address, Lat, Lng columns.");
    return;
  }
  
  var addressColumn = 1;
  var addressRow;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  var cntColumn = addressColumn + 3;
  var ctyColumn = addressColumn + 4;
  var adrColumn = addressColumn + 5;
  
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  var location;
  
  for (addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow) {
    var address = cells.getCell(addressRow, addressColumn).getValue();
    
    // Geocode the address and plug the lat, lng pair into the 
    // 2nd and 3rd elements of the current range row.
    location = geocoder.geocode(address);
   
    // Only change cells if geocoder seems to have gotten a 
    // valid response.
    if (location.status == 'OK') {
      lat = location["results"][0]["geometry"]["location"]["lat"];
      lng = location["results"][0]["geometry"]["location"]["lng"];
      
      addresscomponents = location["results"][0]["address_components"];
      
      cnt = extractFromAdress(addresscomponents, "country");
      cty = extractFromAdress(addresscomponents, "locality");
      adr = location["results"][0]["formatted_address"];
      
      
      // cnt = location["results"][0]["geometry"]["location"]["lng"];
      // cty = 
      // adr = location["results"][0]["geometry"]["location"]["street_address"];
      
      cells.getCell(addressRow, latColumn).setValue(lat);
      cells.getCell(addressRow, lngColumn).setValue(lng);
      cells.getCell(addressRow, cntColumn).setValue(cnt);
      cells.getCell(addressRow, ctyColumn).setValue(cty);
      cells.getCell(addressRow, adrColumn).setValue(adr);
    }
  }
};

function extractFromAdress(components, type){
    for (var i=0; i<components.length; i++)
        for (var j=0; j<components[i].types.length; j++)
            if (components[i].types[j]==type) return components[i].long_name;
    return "";
}

function extractFromAdress_old(components, type){
    for (var i=0; i<components; i++)
        for (var j=0; j<components[i].types; j++)
            if (components[i].types[j]==type) return components[i].long_name;
    return "";
}





// OLD:

// commented versions

function export_gcal_to_gsheet_old(){

//
// Export Google Calendar Events to a Google Spreadsheet
//
// This code retrieves events between 2 dates for the specified calendar.
// It logs the results in the current spreadsheet starting at cell A2 listing the events,
// dates/times, etc and even calculates event duration (via creating formulas in the spreadsheet) and formats the values.
//
// I do re-write the spreadsheet header in Row 1 with every run, as I found it faster to delete then entire sheet content,
// change my parameters, and re-run my exports versus trying to save the header row manually...so be sure if you change
// any code, you keep the header in agreement for readability!
//
// 1. Please modify the value for mycal to be YOUR calendar email address or one visible on your MY Calendars section of your Google Calendar
// 2. Please modify the values for events to be the date/time range you want and any search parameters to find or omit calendar entires
// Note: Events can be easily filtered out/deleted once exported from the calendar
// 
// Reference Websites:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event
//

var mycal = "1si69sarlsi2k2orlnij0423jk@group.calendar.google.com";
var cal = CalendarApp.getCalendarById(mycal);

// Optional variations on getEvents
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"));
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"), {search: 'word1'});
// 
// Explanation of how the search section works (as it is NOT quite like most things Google) as part of the getEvents function:
//    {search: 'word1'}              Search for events with word1
//    {search: '-word1'}             Search for events without word1
//    {search: 'word1 word2'}        Search for events with word2 ONLY
//    {search: 'word1-word2'}        Search for events with ????
//    {search: 'word1 -word2'}       Search for events without word2
//    {search: 'word1+word2'}        Search for events with word1 AND word2
//    {search: 'word1+-word2'}       Search for events with word1 AND without word2
//
var events = cal.getEvents(new Date("February 1, 2020 00:00:00 CST"), new Date("December 30, 2025 23:59:59 CST"), {search: ''});
  
// var startDate = Date("February 1, 2020 00:00:00 CST");
// var maxDate = "December 30, 2025 23:59:59 CST"; 
  
// extra events part for Google Calendar PAI
// var items = Calendar.Events.list(mycal, {
//    "orderBy" : "startTime",
//    "timeMin" : "2020-02-01T00:00:00. 000Z",
//    "timeMax" : "2025-12-30T23:59:59. 000Z"});
  
// , items[i].htmlLink()
// , "iCalUID"

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Import from Google iCal");
// Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
// sheet.clearContents();  

// Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
// of the getRange entry below
var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event", "iCalUID"]]
var range = sheet.getRange(1,1,1,15);
range.setValues(header);

  
// Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
for (var i=0;i<events.length;i++) {
var row=i+2;
var myformula_placeholder = '';
// Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
// NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
var details=[[mycal,events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent(), events[i].getId()]];
var range=sheet.getRange(row,1,1,15);
range.setValues(details);

// Writing formulas from scripts requires that you write the formulas separate from non-formulas
// Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
var cell=sheet.getRange(row,7);
cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
cell.setNumberFormat('.00');

}
}





// Not used:



function getGeocodingRegion() {
  return PropertiesService.getDocumentProperties().getProperty('GEOCODING_REGION') || 'us';
}


function setGeocodingRegion(region) {
  PropertiesService.getDocumentProperties().setProperty('GEOCODING_REGION', region);
  updateMenu();
}
function promptForGeocodingRegion() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Set the Geocoding Country Code (currently: ' + getGeocodingRegion() + ')',
    'Enter the 2-letter country code (ccTLD) that you would like ' +
    'the Google geocoder to search first for results. ' +
    'For example: Use \'uk\' for the United Kingdom, \'us\' for the United States, etc. ' +
    'For more country codes, see: https://en.wikipedia.org/wiki/Country_code_top-level_domain',
    ui.ButtonSet.OK_CANCEL
  );
  // Process the user's response.
  if (result.getSelectedButton() == ui.Button.OK) {
    setGeocodingRegion(result.getResponseText());
  }
}

// CHANGE LOCATION SCRIPT

function goecode_sheets_events() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Import from Google iCal");
  var range = sheet.getDataRange();
  var cells = range.getValues();
 
  var latitudes = [];
  var longitudes = [];
  var formatted_addresses = [];
 
  for (var i = 0; i < cells.length; i++) {
    var address = cells[i][0];
    var geocoder = Maps.newGeocoder().geocode(address);
    var res = geocoder.results[0];
    var lat = lng = 0;
    var formatted_address = '';
    if (res) {
      lat = res.geometry.location.lat;
      lng = res.geometry.location.lng;
      formatted_address = res.formatted_address;
    }
     
    latitudes.push([lat]);
    longitudes.push([lng]);
    formatted_addresses.push([formatted_address]);
    Utilities.sleep(1000);
  }
 
  sheet.getRange('O2')
  .offset(0, 0, latitudes.length)
  .setValues(latitudes);
  sheet.getRange('P2')
  .offset(0, 0, longitudes.length)
  .setValues(longitudes);
  sheet.getRange('P3')
  .offset(0, 0, formatted_addresses.length)
  .setValues(formatted_addresses);
}



function positionToAddress() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cells = sheet.getActiveRange();
  
  // Must have selected 3 columns (Address, Lat, Lng).
  // Must have selected at least 1 row.

  if (cells.getNumColumns() != 3) {
    Logger.log("Must select at least 3 columns: Address, Lat, Lng columns.");
    return;
  }

  var addressColumn = 1;
  var addressRow;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  var location;
  
  for (addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow) {
    var lat = cells.getCell(addressRow, latColumn).getValue();
    var lng = cells.getCell(addressRow, lngColumn).getValue();
    
    // Geocode the lat, lng pair to an address.
    location = geocoder.reverseGeocode(lat, lng);
   
    // Only change cells if geocoder seems to have gotten a 
    // valid response.
    Logger.log(location.status);
    if (location.status == 'OK') {
      var address = location["results"][0]["formatted_address"];

      cells.getCell(addressRow, addressColumn).setValue(address);
    }
  }  
};

function generateMenu() {
  // var setGeocodingRegionMenuItem = 'Set Geocoding Region (Currently: ' + getGeocodingRegion() + ')';
  
  // {
  //   name: setGeocodingRegionMenuItem,
  //   functionName: "promptForGeocodingRegion"
  // },
  
  var entries = [{
    name: "Geocode Selected Cells (Address to   Lat, Long)",
    functionName: "addressToPosition"
  },
  {
    name: "Geocode Selected Cells (Address from Lat, Long)",
    functionName: "positionToAddress"
  }];
  
  return entries;
}

function updateMenu() {
  SpreadsheetApp.getActiveSpreadsheet().updateMenu('Geocode', generateMenu())
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 *
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('Geocode', generateMenu());
  // SpreadsheetApp.getActiveSpreadsheet().addMenu('Region',  generateRegionMenu());
  // SpreadsheetApp.getUi()
  //   .createMenu();
};













// Notes documentation:
// function export_gcal_to_gsheet is ran every hour

function export_gcal_to_gsheet_test(){

// variables
 
var mycal = "1si69sarlsi2k2orlnij0423jk@group.calendar.google.com";
var cal = CalendarApp.getCalendarById(mycal);
var events = cal.getEvents(new Date("February 1, 2020 00:00:00 CST"), new Date("December 30, 2025 23:59:59 CST"), {search: ''});
console.log({message: 'Events', initialData: events});

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Import from Google iCal");
var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event", "iCalUID","Event Location","Event Location-nonzero","Latitude","Longitude","Country","City","Address"]]
console.log({message: 'Header', initialData: header});
var range = sheet.getRange(1,1,1,22);
range.setValues(header);

// for-loop
 
for (var i=0;i<events.length;i++) {
  
var row=i+1;
var myformula_placeholder = '';

//loop through calendar events and get values
var details=[[mycal,events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent(), events[i].getId(),'=D'+(i+2),'=if(N'+(i+2)+'<>"",if(P'+(i+2)+'<>"",P'+(i+2)+',"/"),"")']];
console.log({message: 'Details', initialData: details});
var range=sheet.getRange(row+1,1,1,17);
range.setValues(details);

// var addresscomponents = location["results"]["address_components"]
// console.log({message: 'Address components', initialData: addresscomponents});
var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion())
var locationz = geocoder.geocode([events[i].getLocation()]);
console.log({message: 'Location', initialData: locationz, });
console.log({message: 'Location status', initialData: locationz.status, });

var cells = sheet.getRange("Q2:V1000");
var addressColumn = 1;
var latColumn = addressColumn + 1;
var lngColumn = addressColumn + 2;
var cntColumn = addressColumn + 3;
var ctyColumn = addressColumn + 4;
var adrColumn = addressColumn + 5;

if (locationz.status == 'OK') {

  
var lat = locationz["results"][i]["geometry"]["location"]["lat"];
var lng = locationz["results"][i]["geometry"]["location"]["lng"];
cells.getCell(row, latColumn).setValue(lat);
cells.getCell(row, lngColumn).setValue(lng);

var addresscomponents = locationz["results"][i]["address_components"];
var cnt = extractFromAdress(addresscomponents, "country");
var cty = extractFromAdress(addresscomponents, "locality");
cells.getCell(row, cntColumn).setValue(cnt);
cells.getCell(row, ctyColumn).setValue(cty);
      
var adr = locationz["results"][i]["formatted_address"];
cells.getCell(row, adrColumn).setValue(adr);
    
  }
if (locationz.status == null) {
  {cells.getCell(row, latColumn).setValue("/");
cells.getCell(row, lngColumn).setValue("/");
  cells.getCell(row, cntColumn).setValue("/");
cells.getCell(row, ctyColumn).setValue("/");
  cells.getCell(row, adrColumn).setValue("/");}

}
}
var cell=sheet.getRange(row,7);
cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
cell.setNumberFormat('.00');
}

 //time-formatting


//  if (location.status == 'OK') {
//var details_geocode=[[location["results"][i]["geometry"]["location"]["lat"],location["results"][i]["geometry"]["location"]["lng"],extractFromAdress(addresscomponents[i], "country"),extractFromAdress(addresscomponents[i], "locality"),location["results"][i]["formatted_address"]]];
//console.log({message: 'Geocode details', initialData: details_geocode});
//var range_geocode = sheet.getRange(row,18,1,5);
//range_geocode.setValues(details_geocode);
//}
  

  

  





function addressToPosition_old() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Import from Google iCal");
  sheet.getRange("R2:V1000").clearContent();
  var cells = sheet.getRange("Q2:V1000");
  // var cells = sheet.getActiveRange();
  
  // Must have selected 3 columns (Address, Lat, Lng).
  // Must have selected at least 1 row.

  if (cells.getNumColumns() != (6)) {
    Logger.log("Must select at least 3 columns: Address, Lat, Lng columns.");
    return;
  }
  
  var addressColumn = 1;
  var addressRow;
  
  var latColumn = addressColumn + 1;
  var lngColumn = addressColumn + 2;
  var cntColumn = addressColumn + 3;
  var ctyColumn = addressColumn + 4;
  var adrColumn = addressColumn + 5;
  
  var geocoder = Maps.newGeocoder().setRegion(getGeocodingRegion());
  var location;
  
  for (addressRow = 1; addressRow <= cells.getNumRows(); ++addressRow) {
    var address = cells.getCell(addressRow, addressColumn).getValue();
    
    // Geocode the address and plug the lat, lng pair into the 
    // 2nd and 3rd elements of the current range row.
    location = geocoder.geocode(address);
   
    // Only change cells if geocoder seems to have gotten a 
    // valid response.
    if (location.status == 'OK') {
      lat = location["results"][0]["geometry"]["location"]["lat"];
      lng = location["results"][0]["geometry"]["location"]["lng"];
      
      addresscomponents = location["results"][0]["address_components"];
      
      cnt = extractFromAdress(addresscomponents, "country");
      cty = extractFromAdress(addresscomponents, "locality");
      adr = location["results"][0]["formatted_address"];
      
      
      // cnt = location["results"][0]["geometry"]["location"]["lng"];
      // cty = 
      // adr = location["results"][0]["geometry"]["location"]["street_address"];
      
      cells.getCell(addressRow, latColumn).setValue(lat);
      cells.getCell(addressRow, lngColumn).setValue(lng);
      cells.getCell(addressRow, cntColumn).setValue(cnt);
      cells.getCell(addressRow, ctyColumn).setValue(cty);
      cells.getCell(addressRow, adrColumn).setValue(adr);
    }
  }
};