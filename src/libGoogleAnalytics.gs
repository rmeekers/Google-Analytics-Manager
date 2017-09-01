/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 * @version 1.5
 * @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
 *
 ******************
 * Google Analytics functions
 ******************
 */


/**
 * Helper function to send a hit to Google Analytics
 * @param {string} hitType - event or exception
 * @param {string} screenName - Identifier of the apiType / called function
 * @param {string|boolean} p1 - case event: eventCategory, case exception: is exception fatal? (boolean)
 * @param {string} p2 - case event: eventAction, case exception: exception description
 */
function registerGoogleAnalyticsHit(hitType, screenName, p1, p2) {
    try {
        var payloadData = [];
        payloadData.push(
            ['v', '1'],
            ['tid', 'UA-34001397-11'],
            ['cid', SpreadsheetApp.getActiveSpreadsheet().getId()],
            ['an', settings.applicationName],
            ['av', settings.applicationVersion],
            ['z', Math.floor(Math.random() * 10E7)]
        );

        if (hitType == 'event') {
            payloadData.push(
                ['t', hitType],
                ['cd', screenName],
                ['ec', screenName],
                ['ea', p1],
                ['el', p2]
            );
        } else if (hitType == 'exception') {
            payloadData.push(
                ['t', hitType],
                ['cd', screenName],
                ['exf', p1],
                ['exd', p2]
            );
        } else {
            payloadData.push(
                ['t', 'exception'],
                ['cd', 'registerGoogleAnalyticsHit'],
                ['exf', false],
                ['exd', 'Incorrect hitType used for function registerGoogleAnalyticsHit(): ' + hitType]
            );
        };

        var payload = payloadData.map(function(el) {
            return el.join('=');
        }).join('&');

        var options = {
            'method': 'post',
            'payload': payload
        };

        UrlFetchApp.fetch('https://ssl.google-analytics.com/collect', options);

    } catch (e) {
        registerGoogleAnalyticsHit('exception', 'registerGoogleAnalyticsHit', false, 'registerGoogleAnalyticsHit failed to execute');
    }

    return;
}
