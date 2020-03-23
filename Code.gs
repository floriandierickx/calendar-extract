//*** Extract Google Calendar Events and add them to the Google Sheet ***

function export_gcal_to_gsheet(){

// variables
 
var mycal = "1si69sarlsi2k2orlnij0423jk@group.calendar.google.com";
var cal = CalendarApp.getCalendarById(mycal);
var events = cal.getEvents(new Date("February 1, 2020 00:00:00 CST"), new Date("December 30, 2025 23:59:59 CST"), {search: ''});
console.log({message: 'Number of events', initialData: events});

var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Raw Import from Google iCal");
var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event", "iCalUID"]]
var range = sheet.getRange(1,1,1,15);
range.setValues(header);

// for-loop to extract google calendar events and add them to the google sheets
 
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


// *** Geocode event location to latitude, longitude, country, city and address ***

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