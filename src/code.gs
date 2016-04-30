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
  white: '#ffffff',
  purple: '#c27ba0',
  yellow: '#e7fe2b',
  grey: '#666666',
  valid: '#0fc357',
  invalid: '#e06666'
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
 * Replace 'undefined' in a given array with another value
 * @param {array} array
 * @param {string} value
 */
function replaceUndefinedInArray(array, value) {
    return array.map(function(el) {
        return el === undefined ? value : el;
    });
}

/**
 * onOpen function runs on application open
 * @param {*} e
 */
function onOpen(e) {
    ui = SpreadsheetApp.getUi();
    return ui
        .createMenu('GA Manager')
        .addItem('Create report', 'showSidebar')
        .addSeparator()
        .addSubMenu(ui.createMenu('Advanced')
            .addItem('Insert Properties Sheet', 'createSheetProperties')
            .addItem('Insert Views Sheet', 'createSheetViews'))
        .addToUi();
}

/**
 * Helper function to create a new Properties sheet from the menu
 */
function createSheetProperties() {
    var ui = SpreadsheetApp.getUi();

    var result = ui.alert(
        'Pay attention',
        'If there is already a sheet named \'GAM: Properties\' that sheet will be reinitialized. ' +
        'If you want to start over, please delete the existing sheet first and re-run this function',
            ui.ButtonSet.OK_CANCEL);

    if (result == ui.Button.OK) {
        createSheet('properties');
    }
}

/**
 * Helper function to create a new Views sheet from the menu
 */
function createSheetViews() {
    var ui = SpreadsheetApp.getUi();
    
    var result = ui.alert(
        'Pay attention',
        'If there is already a sheet named \'GAM: Views\' that sheet will be reinitialized. ' +
        'If you want to start over, please delete the existing sheet first and re-run this function',
            ui.ButtonSet.OK_CANCEL);

    if (result == ui.Button.OK) {
        createSheet('views');
    }
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
        .setTitle('GA Manager')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    return SpreadsheetApp
        .getUi()
        .showSidebar(ui);
}

function getReports() {
    var arr = [];
    for (var p in api) {
        var o = {
            name: api[p].name,
            id: p
        };
        arr.push(o);
    }
    return JSON.stringify(arr);
}

function saveReportDataFromSidebar(data) {
    var parsed = JSON.parse(data);
    //return createSheet(parsed.report);
    return generateReport(parsed.ids, parsed.report);
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

        // If there are too many columns, delete additional columns
        if (maxColumns - numberOfColumnsNeeded > 0) {
            this.sheet.deleteColumns(numberOfColumnsNeeded + 1, maxColumns - numberOfColumnsNeeded);
        }
        // If there aren't enough columns, add additional columns
        else if (maxColumns - numberOfColumnsNeeded < 0) {
            this.sheet.insertColumns(this.sheet.getLastColumn(), numberOfColumnsNeeded - maxColumns);
        }

        cb.call(this);
    },

    /**
     * Clear existing the existing sheet
     */
    clearSheet: function(cb) {
        var range = this.sheet.getRange(1, 1, this.sheet.getMaxRows(), this.sheet.getMaxColumns());
        range
            .clearContent()
            .clearDataValidations()
            .clearFormat()
            .clearNote();

        cb.call(this);
    },

    buildHeader: function(cb) {
        var rowHeight = 35;
        var headerRow = this.sheet.setRowHeight(1, rowHeight).getRange(1, 1, 1, this.headerLength);

        // add style header row
        headerRow
            .setBackgrounds(this.header.colors)
            .setFontColor(colors.white)
            .setFontSize(12)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setValues(this.header.names);

        // freeze the header row
        this.sheet.setFrozenRows(1);

        cb.call(this);
    },

    buildDataValidation: function(cb) {

        var dataValidationArray = this.header.dataValidation;

        // dataValidation should be set per column, so we have to loop over the dataValidationArray
        for (var i = 0; i < dataValidationArray.length; i++) {

            // Only set dataValidation when the config is present for a column
            if (dataValidationArray[i]) {
                var dataRange = this.sheet.getRange(2, i + 1, this.sheet.getMaxRows(), 1);
                dataRule = SpreadsheetApp.newDataValidation()
                            .requireValueInList(dataValidationArray[i], true)
                            .build();

                // add dataValidation to the sheet
                dataRange.setDataValidation(dataRule);
            }
        }

        cb.call(this);
    },

    /**
     * Insert data into the sheet
     */
    insertData: function(cb) {
        var dataRange = this.sheet.getRange(2, 1, this.data.length, this.headerLength);

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
            this.buildHeader(function() {
                this.buildDataValidation(function() {
                    this.cleanup();
                });
            });
        });
      
    },

    buildData: function() {
        this.setNumberOfColumns(function() {
            this.clearSheet(function() {
                this.buildHeader(function() {
                    this.buildDataValidation(function() {
                        this.insertData(function() {
                            this.cleanup();
                        });
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
    properties: {
        initSheet: function(cb) {
            this.header = this.getConfig();

            cb();

            return this;
        },
        initData: function(config) {
            this.account = config.account;
            this.accountName = this.account.name;

            this.header = this.getConfig();

            return this;
        },
        name: 'Properties',
        getConfig: function() {
            var data = [
                {
                    name: 'Include',
                    dataValidation: [
                        'Yes',
                        'No'
                    ]
                },{
                    name: 'Account Name'
                },{
                    name: 'Property Name'
                },{
                    name: 'Property ID'
                },{
                    name: 'Industry',
                    dataValidation: [
                        'UNSPECIFIED',
                        'ARTS_AND_ENTERTAINMENT',
                        'AUTOMOTIVE',
                        'BEAUTY_AND_FITNESS',
                        'BOOKS_AND_LITERATURE',
                        'BUSINESS_AND_INDUSTRIAL_MARKETS',
                        'COMPUTERS_AND_ELECTRONICS',
                        'FINANCE',
                        'FOOD_AND_DRINK',
                        'GAMES',
                        'HEALTHCARE',
                        'HOBBIES_AND_LEISURE',
                        'HOME_AND_GARDEN',
                        'INTERNET_AND_TELECOM',
                        'JOBS_AND_EDUCATION',
                        'LAW_AND_GOVERNMENT',
                        'NEWS',
                        'ONLINE_COMMUNITIES',
                        'OTHER',
                        'PEOPLE_AND_SOCIETY',
                        'PETS_AND_ANIMALS',
                        'REAL_ESTATE',
                        'REFERENCE',
                        'SCIENCE',
                        'SHOPPING',
                        'SPORTS',
                        'TRAVEL'
                    ]
                },{
                    name: 'Default View ID'
                },{
                    name: 'Starred',
                    dataValidation: [
                        'TRUE',
                        'FALSE'
                    ]
                },{
                    name: 'Website URL'
                }
            ];
            return sheetConfig(data);
        },
        requestData: function(account, cb) {
            var propertiesList = Analytics.Management.Webproperties.list(account).getItems();
            return cb.call(this, propertiesList);
        },
        getData: function(cb) {
            var results = [];

            this.account.forEach(function(account) {
                this.requestData(account.id, function(propertiesList) {
                    propertiesList.forEach(function(property) {
                        var defaults = [
                            '',
                            account.name,
                            property.name,
                            property.id,
                            property.industryVertical,
                            property.defaultProfileId,
                            property.starred,
                            property.websiteUrl
                        ];
                        defaults = replaceUndefinedInArray(defaults, '');

                        results.push(defaults);
                        }, this);
                });
            }, this);

            cb(results);
        }
    },
    views: {
        initSheet: function(cb) {
          this.header = this.getConfig();

          cb();

          return this;
        },
        initData: function(config) {
            this.account = config.account;
            this.accountName = this.account.name;

            this.header = this.getConfig();

            return this;
        },
        name: 'Views',
        getConfig: function() {
            var data = [
                {
                    name: 'Include',
                    dataValidation: [
                        'Yes',
                        'No'
                    ]
                },{
                    name: 'Account Name'
                },{
                    name: 'Property Name'
                },{
                    name: 'Property ID'
                },{
                    name: 'View Name'
                },{
                    name: 'View ID'
                },{
                    name: 'Bot Filtering Enabled',
                    dataValidation: [
                        'TRUE',
                        'FALSE'
                    ]
                },{
                    name: 'Currency',
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
                    name: 'eCommerce Tracking',
                    dataValidation: [
                        'TRUE',
                        'FALSE'
                    ]
                },{
                    name: 'Exclude Query Params'
                },{
                    name: 'Site Search Category Params'
                },{
                    name: 'Site Search Query Params'
                },{
                    name: 'Strip Site Search Category Params',
                    dataValidation: [
                        'TRUE',
                        'FALSE'
                    ]
                },{
                    name: 'Strip Site Search Query Params',
                    dataValidation: [
                        'TRUE',
                        'FALSE'
                    ]
                },{
                    name: 'Timezone',
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
                        'Pacific/Wallis'
                    ]
                },{
                    name: 'Type',
                    dataValidation: [
                        'WEB',
                        'APP'
                    ]
                },{
                    name: 'Website URL'
                }
            ];
            return sheetConfig(data);
        },
        requestData: function(account, property, cb) {
            var viewsList = Analytics.Management.Profiles.list(account, property).getItems();
            return cb.call(this, viewsList);
        },
        getData: function(cb) {
            var results = [];

            this.account.forEach(function(account) {
                account.webProperties.forEach(function(property) {
                    this.requestData(account.id, property.id, function(viewsList) {
                        viewsList.forEach(function(view) {
                            var defaults = [
                                '',
                                account.name,
                                property.name,
                                property.id,
                                view.name,
                                view.id,
                                view.botFilteringEnabled,
                                view.currency,
                                view.eCommerceTracking,
                                view.excludeQueryParameters,
                                view.siteSearchCategoryParameters,
                                view.siteSearchQueryParameters,
                                view.stripSiteSearchCategoryParameters,
                                view.stripSiteSearchQueryParameters,
                                view.timezone,
                                view.type,
                                view.websiteUrl
                            ];
                            defaults = replaceUndefinedInArray(defaults, '');

                            results.push(defaults);
                        }, this);
                    });
                }, this);
            }, this);

            cb(results);
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
        .initSheet(function() {
            sheet
                .init({
                    'name': 'GAM: ' + setup.name,
                    'header': setup.header
                })
                .buildSheet();
        });
}

/**
 * Generate report from a certain type for a given account
 * @param {object} account
 * @param {string} type
 */
function generateReport(account, type) {
    var setup = api[type];

    setup
        .initData({ account: account })
        .getData(function(data) {

            // Verify if data is present
            if (data[0] === undefined || !data[0].length) {
                throw new Error('No data found for ' + type + ' in ' + account.name + '.');
            }

            sheet
                .init({
                    'name': 'GAM: ' + setup.name,
                    'header': setup.header,
                    'data': data
                })
                .buildData();
        });
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
