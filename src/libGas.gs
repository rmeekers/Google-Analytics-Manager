/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 * @version 1.2
 * @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
 *
 ******************
 * Google Apps Script functions
 ******************
 */


/**
 * Logger wrapper function
 * @customfunction
 *
 * @param {string} text Text to send to the logger.
 */
function log(text) {
    Logger.log(text);
}


/**
 * Add custom edit trigger
 * @customfunction
 */
function addEditTrigger_() {
    ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
}

/**
 * Process the trigger edit event, execute necessary actions for the edited sheet.
 * This is an alternative to the default edit event on the content sheet, which allows extra functionality as UrlFetch.
 * @customfunction
 *
 * @param {event} event - Edit event
 */
function editEventTriggerController(event) {

  try {
    var sheet = event.source.getActiveSheet();

    // do something if required

  } catch(e){

  }
}

/**
 * Wrapper function to call functions in the library that are not accesible for google.script.run
 * @customfunction
 *
 * @param {string} function_name The of the function to call
 * @param {object} parameters Object containing the parameters and their values to pass to the function
 *
 * @return {object} Object containing success or error message and the return value of the called function
 */
function dialogRunFunction(function_name, parameters) {
  //log('dialogRunFunction')
  //log(function_name);
  //log(parameters);

  switch (function_name) {

    // Load the content of the settings dialog

    case 'saveReportDataFromSidebar':

      return saveReportDataFromSidebar_(parameters);

    default:

      return helperDialogRunFunctionUnknown_(function_name);
  }
}

/**
 * Build the basic return object for google.script.run calls
 * @customfunction
 *
 * @return {object} 'Dialog result object' to contain the success flag, (error/warning) message and the result.
 */
function helperDialogRunFunctionGetReturnObject_() {

  return {
    success: false,
    message: '',
    result: null
  }
}

/**
 * Returns an appropriate result object when an non-defined function is called
 * @customfunction
 *
 * @param {string} function_name The of the function to call
 *
 * @return {object} Dialog result object
 */
function helperDialogRunFunctionUnknown_(function_name) {
  var result = helperDialogRunFunctionGetReturnObject_();
  result.message = 'Unknown function' + ': ' + function_name;

  return result;
}
