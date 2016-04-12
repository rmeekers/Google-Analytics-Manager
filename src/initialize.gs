/** Google Analytics Manager
 * Initializes stuff
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 */


/**
 * onOpen function runs on application open
 * @param {*} e
 */
function onOpen(e) {

  // Create the addon menu
  try {
    SpreadsheetApp.getUi()
       .createMenu('GA Manager')
       .addSubMenu(SpreadsheetApp.getUi().createMenu('Views')
           .addItem('One Submenu Item', 'mySecondFunction')
           .addItem('Another Submenu Item', 'myThirdFunction')
           )
       .addSeparator()
       .addToUi();

  } catch (e) {
    Browser.msgBox(e.message);
  }
}

/**
 * onInstall runs when the script is installed
 * @param {*} e
 */
function onInstall(e) {
  onOpen(e);
}
