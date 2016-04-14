/** Google Analytics Manager
 * Initializes stuff
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 */

/*
 * Configuration
 */

var colors = {
    blue: '#4d90fe',
    green: '#0fc357',
    purple: '#c27ba0',
    yellow: '#e7fe2b',
    grey: '#666666',
    white: '#ffffff',
    red: '#e06666'
};

/**
 * onOpen function runs on application open
 * @param {*} e
 */
function onOpen(e) {

  // Create the menu
  try {
    SpreadsheetApp.getUi()
       .createMenu('GA Manager')
       .addSubMenu(SpreadsheetApp.getUi().createMenu('Views')
           .addItem('List Views', 'listViews')
           .addItem('Create / Update Views', 'processViews')
           )
       .addSeparator()
       .addToUi();

  } catch (e) {
    Browser.msgBox(e.message);
  }
}

function initSheet(config) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(config.name) || spreadsheet.insertSheet(config.name);
}

/**
 * onInstall runs when the script is installed
 * @param {*} e
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Process the views
 */
function processViews() {

}

/**
 * Create new views
 * @param {object} viewData
 * @param {string} accountId
 * @param {string} webPropertyId
 */
function createViews(viewData, accountId, webPropertyId) {

}

/**
 * Update existing views
 * @param {object} viewData
 * @param {string} accountId
 * @param {string} webPropertyId
 * @param {string} profileId
 */
function updateViews(viewData, accountId, webPropertyId, profileId) {

}

/**
 * List existing views
 * @param {string} accountId
 * @param {string} webPropertyId
 */
function listViews(accountId, webPropertyId) {

}
