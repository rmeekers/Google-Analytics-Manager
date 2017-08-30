/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 * @version 1.2
 * @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
 *
 ******************
 * UI functions
 ******************
 */

 /**
  * Define the spreadsheet UI
  */
 var ui = SpreadsheetApp.getUi();

 /**
  * Show the sidebar
  */
 function showSidebar() {
     var ui = HtmlService
         .createTemplateFromFile('index')
         .evaluate()
         .setTitle('GA Manager')
         .setSandboxMode(HtmlService.SandboxMode.IFRAME);

     registerGoogleAnalyticsHit('event', 'showSidebar', 'Click', 'Menu');

     return SpreadsheetApp
         .getUi()
         .showSidebar(ui);
 }
