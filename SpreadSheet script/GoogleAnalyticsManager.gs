/** Google Analytics Manager
* Manage your Google Analytics account in batch via a Google Sheet
*
* @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
* @author Rutger Meekers [rutger@meekers.eu]
* @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
*
******************
* Library base functions
******************
*
* Disclaimer:
* 1. Do not edit the code below, as this will break the application
* 2. The identifier of the 'Google Analytics Manager' must remain 'GoogleAnalyticsManager', or this will break the application.
*/


/**
* Repair the required custom triggers.
*/
function repairTriggers() {
  GoogleAnalyticsManager.initTriggers();
}

/**
* Wrapper for google.scripts.run to call library functions
*/
function runFunction(function_name, parameters) {
  return GoogleAnalyticsManager.dialogRunFunction(function_name, parameters);
}

/**
* Run basic document setup.
*/
function runSetup() {
  GoogleAnalyticsManager.setupDocument();
}

/**
* Load app settings, build menu, etc
*
* @param {event} event Open event
*/
function onOpen(event) {
  GoogleAnalyticsManager.initMenu(event);
}

/**
* Process sheet onEdit events with custom trigger
*
* @param {event} event Edit event
*/
function onEditTrigger(event) {
  GoogleAnalyticsManager.editEventTriggerController(event);
}
