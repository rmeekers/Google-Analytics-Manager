/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 *
 * Global Logger, SpreadsheetApp, HtmlService, Analytics
 */

/**
 * Configuration
 ***************
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
 * Helper functions
 ******************
 */

/**
 * Return the sheet settings
 * @param {array} array
 */
function sheetConfig(array) {
    return {
        names: [array.map(function(element) { return element.name; })],
        colors: [array.map(function(element) { return element.color || colors.blue; })],
        dataValidation: array.map(function(element) {
            if (element.dataValidation) {
                return element.dataValidation;
            }
            else {
                return;
            }
        })
    };
}

/**
 * onOpen function runs on application open
 * @param {*} e
 */
function onOpen(e) {
    return SpreadsheetApp
        .getUi()
        .createAddonMenu()
        .addItem('Create report', 'showSidebar')
        .addToUi();
}

/**
 * onInstall runs when the script is installed
 * @param {*} e
 */
function onInstall(e) {
    onOpen(e);
}

/**
 * Show the showSidebar
 */
function showSidebar() {
    var ui = HtmlService
        .createTemplateFromFile('index')
        .evaluate()
        .setTitle('GA Auditor')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    return SpreadsheetApp
        .getUi()
        .showSidebar(ui);
}

function getReports() {

    return JSON.stringify([{
        'name': 'Views',
        'id': 'views'
    }]);
}

function saveReportDataFromSidebar(data) {
    var parsed = JSON.parse(data);
    return createSheet(parsed.report);
}

/**
 * Main functions
 ****************
 */

/**
 * Sheet Data
 */

var sheet = {
    init: function(config) {
        this.workbook = SpreadsheetApp.getActiveSpreadsheet();
        this.name = config.name;
        this.header = config.header;
        this.headerLength = this.header.names[0].length;
        this.data = config.data;
        this.sheet =
          this.workbook.getSheetByName(this.name) ||
          this.workbook.insertSheet(this.name);

        return this;
    },

    /*
     * Define the number of columns needed in the sheet and add or remove columns
     */
    setNumberOfColumns: function(cb) {
        var maxColumns = this.sheet.getMaxColumns();
        var numberOfColumnsNeeded = this.headerLength;

        if (maxColumns - numberOfColumnsNeeded == 0) {
            this.sheet.deleteColumns(numberOfColumns + 1, maxColumns - numberOfColumnsNeeded);
        }
        else if (maxColumns - numberOfColumnsNeeded < 0) {
            this.sheet.insertColumns(this.sheet.getLastColumn(), numberOfColumnsNeeded - maxColumns);
        }

        cb.call(this);
    },

    buildTitle: function(cb) {
        var rowHeight = 35;
        var titleRow = this.sheet.setRowHeight(1, rowHeight).getRange(1, 1, 1, this.headerLength);
        var titleCell = this.sheet.getRange(1, 1);

        titleRow
            .setBackground(colors.grey)
            .setFontColor(colors.white)
            .setFontSize(12)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setHorizontalAlignment('left')
            .mergeAcross();

        titleCell
            .setValue(this.name);

        cb.call(this);
    },

    buildHeader: function(cb) {
        var rowHeight = 35;
        var headerRow = this.sheet.setRowHeight(2, rowHeight).getRange(2, 1, 1, this.headerLength);

        // add style header row
        headerRow
            .setBackgrounds(this.header.colors)
            .setFontColor(colors.white)
            .setFontSize(12)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setValues(this.header.names);

        // freeze the header row
        this.sheet.setFrozenRows(2);

        cb.call(this);
    },

    buildDataValidation: function(cb) {

        var dataValidationArray = this.header.dataValidation;

        // dataValidation should be set per column, so we have to loop over the dataValidationArray
        for (var i = 0; i < dataValidationArray.length; i++) {

            // Only set dataValidation when the config is present for a column
            if (dataValidationArray[i]) {
                var dataRange = this.sheet.getRange(3, i + 1, this.sheet.getMaxRows(), 1);
                dataRule = SpreadsheetApp.newDataValidation()
                            .requireValueInList(dataValidationArray[i], true)
                            .build();

                // clear existing data
                if (!dataRange.isBlank()) {
                    dataRange.clearContent();
                }

                // add dataValidation to the sheet
                dataRange.setDataValidation(dataRule);
            }
        }

        cb.call(this);
    },

    insertData: function(cb) {
        var dataRange = this.sheet.getRange(3, 1, this.data.length, this.headerLength);
        var allData = this.sheet.getRange(3, 1, this.sheet.getMaxRows(), this.headerLength);

        // clear existing data
        if (!dataRange.isBlank()) {
            allData.clearContent();
        }

        // add data to sheet
        dataRange.setValues(this.data);

        cb.call(this);
    },

    cleanup: function() {
        // auto resize all columns
        this.header.names[0].forEach(function(e, i) {
            this.sheet.autoResizeColumn(i + 1);
        }, this);
    },

    buildSheet: function() {
        this.setNumberOfColumns(function() {
            this.buildTitle(function() {
                this.buildHeader(function() {
                    this.buildDataValidation(function() {
                        this.cleanup();
                    });
                });
            });
        });
      
    },

    buildData: function() {
        this.setNumberOfColumns(function() {
            this.buildTitle(function() {
                this.buildHeader(function() {
                    this.insertData(function() {
                        this.cleanup();
                    });
                });
            });
        });
    }
};

/**
 * API Data
 */
var api = {
  views: {
    init: function(cb) {
      this.header = this.getConfig();
      cb();
      return this;
    },
    account: function(config) {
        this.account = config.account;
        this.accountName = this.account.name;

        return this;
    },
    name: 'Views',
    getConfig: function() {
        var data = [{
          name: 'Include',
          dataValidation: [
            'Yes',
            'No'
          ]
        },{
          name: 'webPropertyId'
        },{
          name: 'name'
        },{
          name: 'botFilteringEnabled',
          dataValidation: [
            'TRUE',
            'FALSE'
          ]
        },{
          name: 'currency',
          dataValidation: [
            'ARS',
            'AUD',
            'BGN',
            'BRL',
            'CAD',
            'CHF',
            'CNY',
            'CZK',
            'DKK',
            'EUR',
            'GBP',
            'HKD',
            'HUF',
            'IDR',
            'INR',
            'JPY',
            'KRW',
            'LTL',
            'MXN',
            'NOK',
            'NZD',
            'PHP',
            'PLN',
            'RUB',
            'SEK',
            'THB',
            'TRY',
            'TWD',
            'USD',
            'VND',
            'ZAR'
          ]
        },{
          name: 'eCommerceTracking',
          dataValidation: [
            'TRUE',
            'FALSE'
          ]
        },{
          name: 'excludeQueryParameters'
        },{
          name: 'siteSearchCategoryParameters'
        },{
          name: 'siteSearchQueryParameters'
        },{
          name: 'stripSiteSearchCategoryParameters',
          dataValidation: [
            'TRUE',
            'FALSE'
          ]
        },{
          name: 'stripSiteSearchQueryParameters',
          dataValidation: [
            'TRUE',
            'FALSE'
          ]
        },{
          name: 'timezone',
          dataValidation: [
            'Africa/Abidjan',
            'Africa/Accra',
            'Africa/Addis_Ababa',
            'Africa/Algiers',
            'Africa/Asmara',
            'Africa/Bamako',
            'Africa/Bangui',
            'Africa/Banjul',
            'Africa/Bissau',
            'Africa/Blantyre',
            'Africa/Brazzaville',
            'Africa/Bujumbura',
            'Africa/Cairo',
            'Africa/Casablanca',
            'Africa/Ceuta',
            'Africa/Conakry',
            'Africa/Dakar',
            'Africa/Dar_es_Salaam',
            'Africa/Djibouti',
            'Africa/Douala',
            'Africa/El_Aaiun',
            'Africa/Freetown',
            'Africa/Gaborone',
            'Africa/Harare',
            'Africa/Johannesburg',
            'Africa/Juba',
            'Africa/Kampala',
            'Africa/Khartoum',
            'Africa/Kigali',
            'Africa/Kinshasa',
            'Africa/Lagos',
            'Africa/Libreville',
            'Africa/Lome',
            'Africa/Luanda',
            'Africa/Lubumbashi',
            'Africa/Lusaka',
            'Africa/Malabo',
            'Africa/Maputo',
            'Africa/Maseru',
            'Africa/Mbabane',
            'Africa/Mogadishu',
            'Africa/Monrovia',
            'Africa/Nairobi',
            'Africa/Ndjamena',
            'Africa/Niamey',
            'Africa/Nouakchott',
            'Africa/Ouagadougou',
            'Africa/Porto-Novo',
            'Africa/Sao_Tome',
            'Africa/Tripoli',
            'Africa/Tunis',
            'Africa/Windhoek',
            'America/Adak',
            'America/Anchorage',
            'America/Anguilla',
            'America/Antigua',
            'America/Araguaina',
            'America/Argentina/Buenos_Aires',
            'America/Argentina/Catamarca',
            'America/Argentina/Cordoba',
            'America/Argentina/Jujuy',
            'America/Argentina/La_Rioja',
            'America/Argentina/Mendoza',
            'America/Argentina/Rio_Gallegos',
            'America/Argentina/Salta',
            'America/Argentina/San_Juan',
            'America/Argentina/San_Luis',
            'America/Argentina/Tucuman',
            'America/Argentina/Ushuaia',
            'America/Aruba',
            'America/Asuncion',
            'America/Atikokan',
            'America/Bahia',
            'America/Bahia_Banderas',
            'America/Barbados',
            'America/Belem',
            'America/Belize',
            'America/Blanc-Sablon',
            'America/Boa_Vista',
            'America/Bogota',
            'America/Boise',
            'America/Cambridge_Bay',
            'America/Campo_Grande',
            'America/Cancun',
            'America/Caracas',
            'America/Cayenne',
            'America/Cayman',
            'America/Chicago',
            'America/Chihuahua',
            'America/Costa_Rica',
            'America/Creston',
            'America/Cuiaba',
            'America/Curacao',
            'America/Danmarkshavn',
            'America/Dawson',
            'America/Dawson_Creek',
            'America/Denver',
            'America/Detroit',
            'America/Dominica',
            'America/Edmonton',
            'America/Eirunepe',
            'America/El_Salvador',
            'America/Fort_Nelson',
            'America/Fortaleza',
            'America/Glace_Bay',
            'America/Godthab',
            'America/Goose_Bay',
            'America/Grand_Turk',
            'America/Grenada',
            'America/Guadeloupe',
            'America/Guatemala',
            'America/Guayaquil',
            'America/Guyana',
            'America/Halifax',
            'America/Havana',
            'America/Hermosillo',
            'America/Indiana/Indianapolis',
            'America/Indiana/Knox',
            'America/Indiana/Marengo',
            'America/Indiana/Petersburg',
            'America/Indiana/Tell_City',
            'America/Indiana/Vevay',
            'America/Indiana/Vincennes',
            'America/Indiana/Winamac',
            'America/Inuvik',
            'America/Iqaluit',
            'America/Jamaica',
            'America/Juneau',
            'America/Kentucky/Louisville',
            'America/Kentucky/Monticello',
            'America/Kralendijk',
            'America/La_Paz',
            'America/Lima',
            'America/Los_Angeles',
            'America/Lower_Princes',
            'America/Maceio',
            'America/Managua',
            'America/Manaus',
            'America/Marigot',
            'America/Martinique',
            'America/Matamoros',
            'America/Mazatlan',
            'America/Menominee',
            'America/Merida',
            'America/Metlakatla',
            'America/Mexico_City',
            'America/Miquelon',
            'America/Moncton',
            'America/Monterrey',
            'America/Montevideo',
            'America/Montserrat',
            'America/Nassau',
            'America/New_York',
            'America/Nipigon',
            'America/Nome',
            'America/Noronha',
            'America/North_Dakota/Beulah',
            'America/North_Dakota/Center',
            'America/North_Dakota/New_Salem',
            'America/Ojinaga',
            'America/Panama',
            'America/Pangnirtung',
            'America/Paramaribo',
            'America/Phoenix',
            'America/Port_of_Spain',
            'America/Port-au-Prince',
            'America/Porto_Velho',
            'America/Puerto_Rico',
            'America/Rainy_River',
            'America/Rankin_Inlet',
            'America/Recife',
            'America/Regina',
            'America/Resolute',
            'America/Rio_Branco',
            'America/Santa_Isabel',
            'America/Santarem',
            'America/Santiago',
            'America/Santo_Domingo',
            'America/Sao_Paulo',
            'America/Scoresbysund',
            'America/Sitka',
            'America/St_Barthelemy',
            'America/St_Johns',
            'America/St_Kitts',
            'America/St_Lucia',
            'America/St_Thomas',
            'America/St_Vincent',
            'America/Swift_Current',
            'America/Tegucigalpa',
            'America/Thule',
            'America/Thunder_Bay',
            'America/Tijuana',
            'America/Toronto',
            'America/Tortola',
            'America/Vancouver',
            'America/Whitehorse',
            'America/Winnipeg',
            'America/Yakutat',
            'America/Yellowknife',
            'Antarctica/Casey',
            'Antarctica/Davis',
            'Antarctica/DumontDUrville',
            'Antarctica/Macquarie',
            'Antarctica/Mawson',
            'Antarctica/McMurdo',
            'Antarctica/Palmer',
            'Antarctica/Rothera',
            'Antarctica/Syowa',
            'Antarctica/Troll',
            'Antarctica/Vostok',
            'Arctic/Longyearbyen',
            'Asia/Aden',
            'Asia/Almaty',
            'Asia/Amman',
            'Asia/Anadyr',
            'Asia/Aqtau',
            'Asia/Aqtobe',
            'Asia/Ashgabat',
            'Asia/Baghdad',
            'Asia/Bahrain',
            'Asia/Baku',
            'Asia/Bangkok',
            'Asia/Beirut',
            'Asia/Bishkek',
            'Asia/Brunei',
            'Asia/Chita',
            'Asia/Choibalsan',
            'Asia/Colombo',
            'Asia/Damascus',
            'Asia/Dhaka',
            'Asia/Dili',
            'Asia/Dubai',
            'Asia/Dushanbe',
            'Asia/Gaza',
            'Asia/Hebron',
            'Asia/Ho_Chi_Minh',
            'Asia/Hong_Kong',
            'Asia/Hovd',
            'Asia/Irkutsk',
            'Asia/Jakarta',
            'Asia/Jayapura',
            'Asia/Jerusalem',
            'Asia/Kabul',
            'Asia/Kamchatka',
            'Asia/Karachi',
            'Asia/Kathmandu',
            'Asia/Khandyga',
            'Asia/Kolkata',
            'Asia/Krasnoyarsk',
            'Asia/Kuala_Lumpur',
            'Asia/Kuching',
            'Asia/Kuwait',
            'Asia/Macau',
            'Asia/Magadan',
            'Asia/Makassar',
            'Asia/Manila',
            'Asia/Muscat',
            'Asia/Nicosia',
            'Asia/Novokuznetsk',
            'Asia/Novosibirsk',
            'Asia/Omsk',
            'Asia/Oral',
            'Asia/Phnom_Penh',
            'Asia/Pontianak',
            'Asia/Pyongyang',
            'Asia/Qatar',
            'Asia/Qyzylorda',
            'Asia/Rangoon',
            'Asia/Riyadh',
            'Asia/Sakhalin',
            'Asia/Samarkand',
            'Asia/Seoul',
            'Asia/Shanghai',
            'Asia/Singapore',
            'Asia/Srednekolymsk',
            'Asia/Taipei',
            'Asia/Tashkent',
            'Asia/Tbilisi',
            'Asia/Tehran',
            'Asia/Thimphu',
            'Asia/Tokyo',
            'Asia/Ulaanbaatar',
            'Asia/Urumqi',
            'Asia/Ust-Nera',
            'Asia/Vientiane',
            'Asia/Vladivostok',
            'Asia/Yakutsk',
            'Asia/Yekaterinburg',
            'Asia/Yerevan',
            'Atlantic/Azores',
            'Atlantic/Bermuda',
            'Atlantic/Canary',
            'Atlantic/Cape_Verde',
            'Atlantic/Faroe',
            'Atlantic/Madeira',
            'Atlantic/Reykjavik',
            'Atlantic/South_Georgia',
            'Atlantic/St_Helena',
            'Atlantic/Stanley',
            'Australia/Adelaide',
            'Australia/Brisbane',
            'Australia/Broken_Hill',
            'Australia/Currie',
            'Australia/Darwin',
            'Australia/Eucla',
            'Australia/Hobart',
            'Australia/Lindeman',
            'Australia/Lord_Howe',
            'Australia/Melbourne',
            'Australia/Perth',
            'Australia/Sydney',
            'Europe/Amsterdam',
            'Europe/Andorra',
            'Europe/Athens',
            'Europe/Belgrade',
            'Europe/Berlin',
            'Europe/Bratislava',
            'Europe/Brussels',
            'Europe/Bucharest',
            'Europe/Budapest',
            'Europe/Busingen',
            'Europe/Chisinau',
            'Europe/Copenhagen',
            'Europe/Dublin',
            'Europe/Gibraltar',
            'Europe/Guernsey',
            'Europe/Helsinki',
            'Europe/Isle_of_Man',
            'Europe/Istanbul',
            'Europe/Jersey',
            'Europe/Kaliningrad',
            'Europe/Kiev',
            'Europe/Lisbon',
            'Europe/Ljubljana',
            'Europe/London',
            'Europe/Luxembourg',
            'Europe/Madrid',
            'Europe/Malta',
            'Europe/Mariehamn',
            'Europe/Minsk',
            'Europe/Monaco',
            'Europe/Moscow',
            'Europe/Oslo',
            'Europe/Paris',
            'Europe/Podgorica',
            'Europe/Prague',
            'Europe/Riga',
            'Europe/Rome',
            'Europe/Samara',
            'Europe/San_Marino',
            'Europe/Sarajevo',
            'Europe/Simferopol',
            'Europe/Skopje',
            'Europe/Sofia',
            'Europe/Stockholm',
            'Europe/Tallinn',
            'Europe/Tirane',
            'Europe/Uzhgorod',
            'Europe/Vaduz',
            'Europe/Vatican',
            'Europe/Vienna',
            'Europe/Vilnius',
            'Europe/Volgograd',
            'Europe/Warsaw',
            'Europe/Zagreb',
            'Europe/Zaporozhye',
            'Europe/Zurich',
            'Indian/Antananarivo',
            'Indian/Chagos',
            'Indian/Christmas',
            'Indian/Cocos',
            'Indian/Comoro',
            'Indian/Kerguelen',
            'Indian/Mahe',
            'Indian/Maldives',
            'Indian/Mauritius',
            'Indian/Mayotte',
            'Indian/Reunion',
            'Pacific/Apia',
            'Pacific/Auckland',
            'Pacific/Bougainville',
            'Pacific/Chatham',
            'Pacific/Chuuk',
            'Pacific/Easter',
            'Pacific/Efate',
            'Pacific/Enderbury',
            'Pacific/Fakaofo',
            'Pacific/Fiji',
            'Pacific/Funafuti',
            'Pacific/Galapagos',
            'Pacific/Gambier',
            'Pacific/Guadalcanal',
            'Pacific/Guam',
            'Pacific/Honolulu',
            'Pacific/Johnston',
            'Pacific/Kiritimati',
            'Pacific/Kosrae',
            'Pacific/Kwajalein',
            'Pacific/Majuro',
            'Pacific/Marquesas',
            'Pacific/Midway',
            'Pacific/Nauru',
            'Pacific/Niue',
            'Pacific/Norfolk',
            'Pacific/Noumea',
            'Pacific/Pago_Pago',
            'Pacific/Palau',
            'Pacific/Pitcairn',
            'Pacific/Pohnpei',
            'Pacific/Port_Moresby',
            'Pacific/Rarotonga',
            'Pacific/Saipan',
            'Pacific/Tahiti',
            'Pacific/Tarawa',
            'Pacific/Tongatapu',
            'Pacific/Wake',
            'Pacific/Wallis']
        },{
          name: 'type',
          dataValidation: [
            'WEB',
            'APP'
          ]
        },{
          name: 'websiteUrl'
        },{
          name: 'viewID'
        }];
        return sheetConfig(data);
    }
  }
};


/**
 * Create sheet, set headers and validation data
 * @param {string} type
 */
function createSheet(type) {
    var setup = api[type];

    setup
        //.init()
        .init(function() {
            sheet
                .init({
                    'name': 'GAM: ' + setup.name,
                    'header': setup.header
                })
                .buildSheet();
        });
}

function testCreateSheet() {
    createSheet('views');
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

/**
 * Retrieve Account Summary from the Google Analytics Management API
 * @return {array}
 */
function getAccountSummary() {
    var items = Analytics.Management.AccountSummaries.list().getItems();

    if (!items) {
        return [];
    }

    return JSON.stringify(items, ['name', 'id', 'webProperties', 'profiles']);
}
