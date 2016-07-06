/* Management Magic for Google Analytics
*    Inserts Views for a GA property
*    https://developers.google.com/analytics/devguides/config/mgmt/v3/mgmtReference/management/profiles/insert
*
* Copyright Rutger Meekers (rutger@meekers.eu)
***************************************************************************/


/**************************************************************************
* Obtains input from user necessary for updating custom dimensions.
*/
function requestInsertView() {
  // Check that the necessary named range exists.
  if (SpreadsheetApp.getActiveSpreadsheet().getRangeByName("header_row")) {
    
    // Update views from the sheet.
    var insertViewResponse = insertViews();
    
    // Output errors and log successes.
    if (insertViewResponse != "success") {
      Browser.msgBox(insertViewResponse);
    } else {
      Logger.log("insertView response: "+ insertViewResponse)
    }
  }
  
  // If there is no named range (necessary to update values), format the sheet and display instructions to the user
  else {
    var sheet = formatViewSheet(true);
    Browser.msgBox("Enter View values into the sheet provided before requesting to update Views.")
  }
}

/**************************************************************************
* Inserts a view with the settings from the active sheet to a property.
* @return {string} Operation output ("success" or error message)
*/
function insertViews() {
  // set common values
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRows = sheet.getDataRange().getNumRows()-1;
  var dataRange = sheet.getRange(2,1,dataRows,sheet.getMaxColumns());
  var views = dataRange.getValues();
  var maxViews = 25;
  var numviewsInserted = 0;
  var viewsInserted = [];
  
  // Iterate through rows of values in the sheet.
  for (var v = 0; v < views.length; v++) {
    
    // Process values marked for inclusion.
    if (views[v][0]) {
      var webPropertyId = views[v][1];
      var accountId = webPropertyId.match(/UA-(.+)-.*/)[1];
      var resource = {};
      resource.name = views[v][2];
      if (views[v][3]) {resource.botFilteringEnabled = views[v][3];}
      if (views[v][4]) {resource.currency = views[v][4];}
      if (views[v][5]) {resource.eCommerceTracking = views[v][5];}
      if (views[v][6]) {resource.excludeQueryParameters = views[v][6];}
      if (views[v][7]) {resource.siteSearchCategoryParameters = views[v][7];}
      if (views[v][8]) {resource.siteSearchQueryParameters = views[v][8];}
      if (views[v][9]) {resource.stripSiteSearchCategoryParameters = views[v][9];}
      if (views[v][10]) {resource.stripSiteSearchQueryParameters = views[v][10];}
      if (views[v][11]) {resource.timezone = views[v][11];}
      if (views[v][12]) {resource.type = views[v][12];}
      if (views[v][13]) {resource.websiteUrl = views[v][13];}
      var profileId = (views[v][14]) ? views[v][14].toString() : '0';
      
      // Increment the number of views updated and add the property to the array of updated properties if it's not already there.
      numviewsInserted++;
      if (viewsInserted.indexOf(webPropertyId) < 0) viewsInserted.push(webPropertyId);
      
      // If there is no name, return an error to the user.
      if (resource.name == "" || resource.name == undefined) {
        Logger.log("Name for View '"+ webPropertyId +"' cannot be empty");
        return "Name for View '"+ webPropertyId +"' cannot be empty";
      }
      
      // If the index is valid, push the value to Google Analytics.
      else {
        
        // Attempt to get the id for the View in the sheet (the API throws an exception when no View exists for the id).
        try {
          
          // If the id exists, set the necessary values update the View
          if (Analytics.Management.Profiles.get(accountId, webPropertyId, profileId).id) {
            resource.id = profileId;
            
            // Attempt to update the View through the API
            try {
              
              var uv = Analytics.Management.Profiles.update(resource, accountId, webPropertyId, profileId);
              //Logger.log("Update View Response: "+uv);
            
              // Update the sheet values with the returned object
              updateViewValues(accountId, webPropertyId, profileId, v+2);
              
              Logger.log("View updated");
            } catch (e) {
              Logger.log("Failed to update View: "+ e);
              return "Failed to update View: "+ e;
            }
          }
        }
        
        // As noted in the try-block comment above, if no View exists, the API throws an exception
        // if no View exists, catch this exception and set the necessary values to insert the View
        catch (e) {
          // Attempt to insert the View
          try {
            var cv = Analytics.Management.Profiles.insert(resource, accountId, webPropertyId);
            //Logger.log("Create View Response: "+cv);
            
            // Update the sheet values with the returned object
            sheet.getRange(v+2, 15).setValue(cv.id);
            updateViewValues(accountId, webPropertyId, profileId, v+2);
            
            Logger.log("View created");
          } catch (e) {
            Logger.log("Failed to insert View: "+ e);
            return "Failed to insert View: "+ e;
          }
        }
      }
    }        
  }
  Logger.log("View(s) created / updated");
  return "success";
}

function updateViewValues(gAccountId, gWebPropertyId, gProfileId, gRow) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Attempt to get the View
  try {
    var gv = Analytics.Management.Profiles.get(gAccountId, gWebPropertyId, gProfileId);
    Logger.log("Get View Response: "+gv);
    Logger.log("#### gv.botFilteringEnabled: "+gv.botFilteringEnabled);
    Logger.log("#### gv.eCommerceTracking: "+gv.eCommerceTracking);
    
    // Update the sheet values with the returned object
    if (gv.botFilteringEnabled) {sheet.getRange(gRow, 4).setValue(gv.botFilteringEnabled)};
    sheet.getRange(gRow, 5).setValue(gv.currency);
    sheet.getRange(gRow, 6).setValue(gv.eCommerceTracking);
    if (gv.excludeQueryParameters) {sheet.getRange(gRow, 7).setValue(gv.excludeQueryParameters)};
    if (gv.siteSearchCategoryParameters) {sheet.getRange(gRow, 8).setValue(gv.siteSearchCategoryParameters)};
    if (gv.siteSearchQueryParameters) {sheet.getRange(gRow, 9).setValue(gv.siteSearchQueryParameters)};
    if (gv.stripSiteSearchCategoryParameters) {sheet.getRange(gRow, 10).setValue(gv.stripSiteSearchCategoryParameters)};
    if (gv.stripSiteSearchQueryParameters) {sheet.getRange(gRow, 11).setValue(gv.stripSiteSearchQueryParameters)};
    sheet.getRange(gRow, 12).setValue(gv.timezone);
    sheet.getRange(gRow, 13).setValue(gv.type);
    sheet.getRange(gRow, 14).setValue(gv.websiteUrl);
    
    Logger.log("View retrieved");
  } catch (e) {
    Logger.log("Failed to retrieve View: "+ e);
    return "Failed to retrieve View: "+ e;
  }
}

/**************************************************************************
* Creates a view with the settings from the active sheet to a property.
* @return {string} Operation output ('viewCreated' or error message)
*/
function createView(resource, accountId, webPropertyId) {

}

/**************************************************************************
* Updates a view with the settings from the active sheet to a property.
* @return {string} Operation output ('viewUpdated' or error message)
*/
function updateView(resource, accountId, webPropertyId, profileId) {
  
}

/**************************************************************************
* Retrieves the values of an existing view and writes the values to the row.
* @return {string} Operation output ('viewDataRetrieved' or error message)
*/
function retrieveViewValues(accountId, webPropertyId, profileId) {
  
}

/**************************************************************************
* Processes the active sheet.
* @return {string} Operation output ('sheetProcessed' or error message)
*/
function processViewSheet(operation) {
  
}
