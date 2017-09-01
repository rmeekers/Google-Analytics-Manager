var doc = SpreadsheetApp.getActiveSpreadsheet();
var s_sites = doc.getSheetByName('Websites');

var CLIENT_ID = '718847306519-orks55oek1ithoaduqh7jnuej52tjmuq.apps.googleusercontent.com';
var CLIENT_SECRET = 'EFVr7vNrGOIc0ehBDZlEGuGT';

/**
 * Create OAuth2 service for Google Search Console
 *
 * @return {service} - OAuth2 service
 */
function googleSearchConsoleService() {
  return OAuth2.createService('searchconsole')
      .setAuthorizationBaseUrl(settings.googleOAuth2.authorizationBaseUrl)
      .setTokenUrl(settings.googleOAuth2.tokenUrl)
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)
      .setCallbackFunction('GoogleAnalyticsManager.authCallback')
      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())
      .setScope(settings.searchConsoleApi.oauth2Scope)
      // Below are Google-specific OAuth2 parameters.
      // Sets the login hint, which will prevent the account chooser screen
      // from being shown to users logged in with multiple accounts.
      .setParam('login_hint', Session.getActiveUser().getEmail())
      // Requests offline access.
      .setParam('access_type', 'offline')
      // Forces the approval prompt every time. This is useful for testing,
      // but not desirable in a production application.
      //.setParam('approval_prompt', 'force');
}


/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  var service = googleSearchConsoleService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success!');
  } else {
    return HtmlService.createHtmlOutput('Denied');
  }
}


/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  var service = googleSearchConsoleService();
  service.reset();
}


function getSite() {
  var service = googleSearchConsoleService();

  if (service.hasAccess()) {

    var headers = {
      'Authorization': 'Bearer ' + service.getAccessToken()
    };

    var options = {
      'headers': headers,
      'method': 'GET',
      'muteHttpExceptions': true
    };

    //var maxRows = doc.getSheetByName('Websites').getLastRow();
    //var urls = doc.getSheetByName('Websites').getRange(2, 1, maxRows);

    var siteUrl = doc.getActiveCell().getValue();
    var urlClean = encodeURIComponent(siteUrl);

    var response = UrlFetchApp.fetch(settings.searchConsoleApi.apiUrl + urlClean, options);
    //TODO: do something useful instead of logging :)
    log(getSearchConsoleSiteAccessLevel(response));

  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s', authorizationUrl);
    Browser.msgBox('Open the following URL and re-run the script: ' + authorizationUrl)
  }
}


function getSearchConsoleSiteAccessLevel(response) {
    var result;
    var responseContentText = JSON.parse(response.getContentText());
    var responseCode = response.getResponseCode();

    if(responseCode != '200') {
      result = 'No Access';
  } else if(responseContentText.siteUrl.indexOf(decodeURIComponent(urlClean)) > -1){
      result = responseContentText.permissionLevel;
    } else {
      result = 'No Access';
    }

    return result;
}

/**
* Authorize with Google Search Console
*/
function authorizeGoogleSearchConsole() {
    var service = googleSearchConsoleService();
    if (service.hasAccess()) {
    var url = 'https://www.googleapis.com/webmasters/v3/sites';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });
    var result = JSON.parse(response.getContentText());
    Logger.log(JSON.stringify(result, null, 2));
  } else {
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log('Open the following URL and re-run the script: %s', authorizationUrl);
    Browser.msgBox('Open the following URL and re-run the script: ' + authorizationUrl)
  }
}


/**
 * List Search Console Sites
 */
function listSearchConsoleSites() {
    var callApi = api['searchConsoleSites'];

    callApi.init('generateReport', function() {
        listData(function(data) {
            // Verify if data is present
            if (data[0] === undefined || !data[0].length) {
                throw new Error('No data found for searchConsoleSites in your account.');
            }
            log('callApi.config: ' + callApi.config);
            sheet
            .init('initSheet', 'GAM: ' + callApi.name, callApi.config, data)
            .buildData();
        });
    });
}
