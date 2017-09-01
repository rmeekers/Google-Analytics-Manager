/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 * @version 1.5
 * @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
 *
 ******************
 * Google Search Console API functions
 ******************
 */

function setGoogleOauthClientId() {
    var ui = SpreadsheetApp.getUi();
    var scriptProperties = PropertiesService.getScriptProperties();
    var existing = scriptProperties.getProperty(settings.googleOAuth2.clientIdScriptPropertyKey);

    if (existing) {
        var response = ui.prompt('Update existing Client ID?', existing, ui.ButtonSet.YES_NO);
    } else {
        var response = ui.prompt('Please enter your Client ID', ui.ButtonSet.OK);
    }

    // Process the user's response
    if (response.getSelectedButton() == ui.Button.YES || response.getSelectedButton() == ui.Button.OK) {
        try {
            scriptProperties.setProperty(settings.googleOAuth2.clientIdScriptPropertyKey, response.getResponseText());
            helperToast_('Client ID saved',response.getResponseText());
        } catch(e) {
            helperToast_('Client ID not saved', e);
        }
    } else {
        helperToast_('Client ID not saved','');
    }
}

function setGoogleOauthClientSecret() {
    var ui = SpreadsheetApp.getUi();
    var scriptProperties = PropertiesService.getScriptProperties();
    var existing = scriptProperties.getProperty(settings.googleOAuth2.clientSecretScriptPropertyKey);

    if (existing) {
        var response = ui.prompt('Update existing Client Secret?', existing, ui.ButtonSet.YES_NO);
    } else {
        var response = ui.prompt('Please enter your Client Secret', ui.ButtonSet.OK);
    }

    // Process the user's response
    if (response.getSelectedButton() == ui.Button.YES || response.getSelectedButton() == ui.Button.OK) {
        try {
            scriptProperties.setProperty(settings.googleOAuth2.clientSecretScriptPropertyKey, response.getResponseText());
            helperToast_('Client Secret saved',response.getResponseText());
        } catch(e) {
            helperToast_('Client Secret not saved', e);
        }
    } else {
        helperToast_('Client Secret not saved','');
    }
}

function getOauthClientId() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var existing = scriptProperties.getProperty(settings.googleOAuth2.clientIdScriptPropertyKey);
    if (existing) {
        return existing;
    } else {
        helperToast_('Client ID unknown','Please finish configuration first and try again');
        setGoogleOauthClientId();
    }
}

function getOauthClientSecret() {
    var scriptProperties = PropertiesService.getScriptProperties();
    var existing = scriptProperties.getProperty(settings.googleOAuth2.clientSecretScriptPropertyKey);
    if (existing) {
        return existing;
    } else {
        helperToast_('Client Secret unknown','Please finish configuration first and try again');
        setGoogleOauthClientSecret();
    }
}


/**
 * Create OAuth2 service for Google Search Console
 *
 * @return {service} - OAuth2 service
 */
function googleSearchConsoleService() {
  return OAuth2.createService('searchconsole')
      .setAuthorizationBaseUrl(settings.googleOAuth2.authorizationBaseUrl)
      .setTokenUrl(settings.googleOAuth2.tokenUrl)
      .setClientId(getOauthClientId())
      .setClientSecret(getOauthClientSecret())
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
    return HtmlService.createHtmlOutput('Access to the Google Search Console API granted. You can close this window and return to the Google Analytics Manager spreadhseet.');
  } else {
    return HtmlService.createHtmlOutput('Access to the Google Search Console API denied.');
  }
}


/**
 * Reset the authorization state.
 */
function resetGoogleSearchConsoleServiceAuth() {
  var service = googleSearchConsoleService();
  service.reset();
  helperToast_('','Google Search Console: API authorization revoked');
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

    var htmlOutput = HtmlService
         .createHtmlOutput('<a href="' + authorizationUrl + '" target="_blank">' + authorizationUrl + '</p>')
         .setWidth(250)
         .setHeight(100);
     SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Open URL to grant access');

  }
}


/**
 * List Search Console Sites
 */
function listSearchConsoleSites() {
    var callApi = api['searchConsoleSites'];

    callApi.init('generateReport', function() {
        callApi.listData(function(data) {
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
