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
  primary: '#4CAF50',
  primaryDark: '#388E3C',
  primaryLight: '#C8E6C9',
  primaryText: '#FFFFFF',
  accent: '#FFC107',
  text: '#212121',
  textSecondary: '#727272',
  divider: '#B6B6B6'
};

/**
 * Helper functions
 ******************
 */

/**
 * Return the sheet settings
 * @param {array} array
 * @returns {array}
 */
function createApiSheetColumnConfigArray(array) {
    return {
        names: [array.map(function(element) { return element.name; })],
        colors: [array.map(function(element) { return element.color || colors.primary; })],
        dataValidation: array.map(function(element) {
            if (element.dataValidation) {
                return element.dataValidation;
            }
            else {
                return;
            }
        }),
        regexValidation: array.map(function(element) { return element.regexValidation; })
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
 * Sorts an array of objects
 *
 *  @param {array}      array               - Array of objects
 *  @param {string}     objectParamToSortBy - Name of object parameter to sort by 
 *  @param {boolean}    sortAscending       - (optional) Sort ascending (default) or decending
*/
function sortArrayOfObjectsByParam(array, objectParamToSortBy, sortAscending) {

    // default to true
    if(sortAscending == undefined || sortAscending != false) {
        sortAscending = true;
    }

    if(sortAscending) {
        array.sort(function (a, b) {
            return a[objectParamToSortBy] > b[objectParamToSortBy];
        });
    }
    else {
        array.sort(function (a, b) {
            return a[objectParamToSortBy] < b[objectParamToSortBy];
        });
    }
}

/**
 * Sort a multidimensional array
 *
 * @param {array}   array       - Array or arrays
 * @param {integer} sortIndex   - Array index to sort by
 *
 */
function sortMultidimensionalArray(array, sortIndex) {
    array.sort(sortFunction);

    function sortFunction(array, b) {
        if (array[sortIndex] === b[sortIndex]) {
            return 0;
        }
        else {
            return (array[sortIndex] < b[sortIndex]) ? -1 : 1;
        }
    }
    return array;
}

/**
 * Display a message to the user
 *
 * @param {string} message
 */
function displayMessageToUser(message) {
    SpreadsheetApp.getUi().alert(message);
}

/**
 * onOpen function runs on application open
 * @param {*} e
 */
function onOpen(e) {
    ui = SpreadsheetApp.getUi();
    return ui
        .createMenu('GA Manager')
        .addItem('Audit GA', 'showSidebar')
        .addItem('Insert / Update Data from active sheet to GA', 'insertData')
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
        'If there is already a sheet named \'GAM: Properties\' that sheet will be reinitialized ' +
        '(all content will be cleared), otherwise a new one will be inserted',
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
        'If there is already a sheet named \'GAM: Views\' that sheet will be reinitialized ' +
        '(all content will be cleared), otherwise a new one will be inserted',
            ui.ButtonSet.OK_CANCEL);

    if (result == ui.Button.OK) {
        createSheet('views');
    }
}

/**
 * Helper function to transpose an array
 * @param {array} a
 * @return {array}
 */
function transposeArray(a)
{
  return a[0].map(function (_, c) { return a.map(function (r) { return r[c]; }); });
}

/**
 * Helper function to define the API Type based on the sheet name
 * @param {string} sheetName
 * @return {string}
 */
function getApiTypeBySheetName(sheetName) {
    // Retrieve API type by looping over the API array
    for (var type in api) {
        if (sheetName == 'GAM: ' + api[type].name) {
            return type;
        }
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
    /*
     * Initiate Sheet functions and prepare some basic stuff
     * @param {string} type
     * @param {string} name
     * @param {array} config
     * @param {array} data
     */
    init: function(type, name, config, data) {
        this.workbook = SpreadsheetApp.getActiveSpreadsheet();
        this.name = name;
        this.sheetColumnConfig = config;

        switch(type) {
            case 'initSheet':
                this.sheet =
                  this.workbook.getSheetByName(this.name) ||
                  this.workbook.insertSheet(this.name);
                this.headerLength = config.names[0].length;
                this.data = data;
                break;
            case 'validateData':
                this.sheet = this.workbook.getSheetByName(this.name);
                this.regexValidation = config.regexValidation;
                this.data = transposeArray(data);
                break;
        }

        return this;
    },
    // TODO: do something usefull after validation
    validate: function() {
        for (var row = 0; row < this.data.length; row++) {
            var regex = this.regexValidation[row];
            // Only validate a cell if there is a regexValidation defined
            if (regex) {
                for (var column = 0; column < this.data[row].length; column++) {
                    var string = String(this.data[row][column]);
                    if (string.match(regex)) {
                      Logger.log('Validate OK: ' + string);
                    }
                    else {
                      Logger.log('Validate NOK: ' + string);
                    }
                } 
            }
        }
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
     * Clear the existing sheet
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
        var rowHeight = 30;
        var headerRow = this.sheet.setRowHeight(1, rowHeight).getRange(1, 1, 1, this.headerLength);

        // add style header row
        headerRow
            .setBackgrounds(this.sheetColumnConfig.colors)
            .setFontColor(colors.primaryText)
            .setFontSize(12)
            .setFontWeight('bold')
            .setVerticalAlignment('middle')
            .setValues(this.sheetColumnConfig.names);

        // freeze the header row
        this.sheet.setFrozenRows(1);

        cb.call(this);
    },

    buildDataValidation: function(cb) {

        var dataValidationArray = this.sheetColumnConfig.dataValidation;

        // dataValidation should be set per column, so we have to loop over the dataValidationArray
        for (var i = 0; i < dataValidationArray.length; i++) {

            // Only set dataValidation when the config is present for a column
            if (dataValidationArray[i]) {
                var dataRange = this.sheet.getRange(2, i + 1, this.sheet.getMaxRows(), 1);
                dataRule = SpreadsheetApp.newDataValidation()
                            .requireValueInList(dataValidationArray[i], true)
                            .setAllowInvalid(false)
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
        this.sheetColumnConfig.names[0].forEach(function(e, i) {
            this.sheet.autoResizeColumn(i + 1);
        }, this);
    },

    buildSheet: function() {
        this.setNumberOfColumns(function() {
            this.clearSheet(function() {
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
    },

    validateData: function() {
        this.validate();
    }
};

/**
 * API Data
 */
var api = {
    properties: {
        name: 'Properties',
        init: function(type, cb, options) {
            this.config = this.sheetColumnConfig();

            switch(type) {
                case 'createSheet':
                    break;
                case 'generateReport':
                    this.account = options.accounts;
                    break;
                case 'getConfig':
                    break;
                case 'insertData':
                    break;
            }

            cb();

            return this;
        },
        sheetColumnConfig: function() {
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
                    name: 'Account ID'
                },{
                    name: 'Name',
                    regexValidation: /.*\S.*/
                },{
                    name: 'ID',
                    regexValidation: /(UA|YT|MO)-\d+-\d+/
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
                    ],
                    regexValidation: /.*\S.*/
                },{
                    name: 'Default View ID'
                },{
                    name: 'Starred',
                    dataValidation: [
                        'TRUE',
                        'FALSE'
                    ]
                },{
                    name: 'Website URL',
                    regexValidation: /^(?:(?:https?|ftp):\/\/)(?:\S+(?::\S*)?@)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,}))\.?)(?::\d{2,5})?(?:[/?#]\S*)?$/i 
                }
            ];
            return createApiSheetColumnConfigArray(data);
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
                            account.id,
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
        name: 'Views',
        init: function(type, cb, options) {
            this.config = this.sheetColumnConfig();

            switch(type) {
                case 'createSheet':
                    break;
                case 'generateReport':
                    this.account = options.accounts;
                    break;
                case 'getConfig':
                    break;
                case 'insertData':
                    break;
            }

            cb();

            return this;
        },
        sheetColumnConfig: function() {
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
                    name: 'Account ID'
                },{
                    name: 'Property Name'
                },{
                    name: 'Property ID',
                    regexValidation: /(UA|YT|MO)-\d+-\d+/
                },{
                    name: 'Name'
                },{
                    name: 'ID'
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
                    ],
                    regexValidation: /.*\S.*/
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
                    ],
                    regexValidation: /.*\S.*/
                },{
                    name: 'Type',
                    dataValidation: [
                        'WEB',
                        'APP'
                    ],
                    regexValidation: /.*\S.*/
                },{
                    name: 'Website URL',
                    regexValidation: /^(?:(?:https?|ftp):\/\/)(?:\S+(?::\S*)?@)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,}))\.?)(?::\d{2,5})?(?:[/?#]\S*)?$/i 
                }
            ];
            return createApiSheetColumnConfigArray(data);
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
                                account.id,
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
    },
    filterLinks: {
        name: 'Filter Links',
        init: function(type, cb, options) {
            this.config = this.sheetColumnConfig();

            switch(type) {
                case 'createSheet':
                    break;
                case 'generateReport':
                    this.account = options.accounts;
                    break;
                case 'getConfig':
                    break;
                case 'insertData':
                    break;
            }

            cb();

            return this;
        },
        sheetColumnConfig: function() {
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
                    name: 'Account ID'
                },{
                    name: 'Property Name'
                },{
                    name: 'Property Id',
                    regexValidation: /(UA|YT|MO)-\d+-\d+/
                },{
                    name: 'View Name'
                },{
                    name: 'Name'
                },{
                    name: 'ID',
                    regexValidation: /^[0-9]+$/
                }
            ];
            return createApiSheetColumnConfigArray(data);
        },
        requestData: function(account, property, view, cb) {
            var flList = Analytics.Management.ProfileFilterLinks.list(account, property, view).getItems();
            return cb.call(this, flList);
        },
        getData: function(cb) {
            var results = [];

            this.account.forEach(function(account) {
                account.webProperties.forEach(function(property) {
                    property.profiles.forEach(function(view) {
                        this.requestData(account.id, property.id, view.id, function(flList) {
                            flList.forEach(function(fl) {
                                var defaults = [
                                    '',
                                    account.name,
                                    account.id,
                                    property.name,
                                    property.id,
                                    view.name,
                                    fl.filterRef.name,
                                    fl.filterRef.id
                                ];
                                defaults = replaceUndefinedInArray(defaults, '');

                                results.push(defaults);
                            }, this);
                        });
                    }, this);
                }, this);
            }, this);

            cb(results);
        }
    },
    customDimensions: {
        name: 'Custom Dimensions',
        init: function(type, cb, options) {
            this.config = this.sheetColumnConfig();

            switch(type) {
                case 'createSheet':
                    break;
                case 'generateReport':
                    this.account = options.accounts;
                    break;
                case 'getConfig':
                    break;
                case 'insertData':
                    break;
            }

            cb();

            return this;
        },
        sheetColumnConfig: function() {
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
                    name: 'Account ID'
                },{
                    name: 'Property Name'
                },{
                    name: 'Property ID',
                    regexValidation: /(UA|YT|MO)-\d+-\d+/
                },{
                    name: 'Name',
                    regexValidation: /.*\S.*/
                },{
                    name: 'Index',
                    regexValidation: /^[0-9]{1,3}$/
                },{
                    name: 'Scope',
                    dataValidation: [
                        'HIT',
                        'SESSION',
                        'USER',
                        'PRODUCT'
                    ],
                    regexValidation: /.*\S.*/
                },{
                    name: 'Active',
                    dataValidation: [
                        'TRUE',
                        'FALSE'
                    ],
                    regexValidation: /.*\S.*/
                }
            ];
            return createApiSheetColumnConfigArray(data);
        },
        listApiData: function(data) {
            var account = data[2];
            var property = data[4];
            return = Analytics.Management.CustomDimensions.list(account, property).getItems();
        },
        getApiData: function(account, property, dimensionId, cb) {
            var cdItem = Analytics.Management.CustomDimensions.get(account, property, dimensionId);
            return cdItem;
        },
        insertApiData: function(data) {
            var values = {
                'name': data[5],
                'index': data[6],
                'scope': data[7],
                'active': data[8]
            };
            var account = data[2];
            var property = data[4];

            return Analytics.Management.CustomDimensions.insert(values, account, property);
        },
        updateApiData: function(data) {
            var values = {
                'name': data[5],
                'scope': data[7],
                'active': data[8]
            };
            var account = data[2];
            var property = data[4];
            var customDimension = 'ga:dimension' + data[6];

            return Analytics.Management.CustomDimensions.update(values, account, property, customDimension);
        },
        getData: function(cb) {
            var results = [];

            this.account.forEach(function(account) {
                account.webProperties.forEach(function(property) {
                    this.listApiData(account.id, property.id, function(cdList) {
                        cdList.forEach(function(cd) {
                            var defaults = [
                                '',
                                account.name,
                                account.id,
                                property.name,
                                property.id,
                                cd.name,
                                cd.index,
                                cd.scope,
                                cd.active
                            ];
                            defaults = replaceUndefinedInArray(defaults, '');

                            results.push(defaults);
                        }, this);
                    });
                }, this);
            }, this);

            cb(results);
        },
        prepareInsertData: function(insertDataRange) {

            // Split insertDataRange in data that needs to be inserted and data that needs to be updated
            var insertData = [];
            var updateData = [];

            for (var i = 0; i < insertDataRange.length; i++) {

                var account = insertDataRange[i][2];
                var property = insertDataRange[i][4];
                var customDimensionId = 'ga:dimension' + insertDataRange[i][6];

                try {
                    var existingCD = this.getApiData(account, property, customDimensionId);
                }
                catch(e) {
                    var existingCD = false;
                }
                if(existingCD && existingCD != false) {
                    updateData.push(insertDataRange[i]);
                }
                else if (existingCD == false) {
                    insertData.push(insertDataRange[i]);
                }
            }

            if (insertData.length > 0){
                for (var i = 0; i < insertData.length; i++) {
                    this.insertApiData(insertData[i]);
                }
            }

            if (updateData.length > 0){
                for (var i = 0; i < updateData.length; i++) {
                    this.updateApiData(updateData[i]);
                }
            }
        }
    },
    accountSummaries: {
        name: 'Account Summaries',
        requestAccountSummaryList: function() {
            var items = Analytics.Management.AccountSummaries.list().getItems();

            // Return an empty array if there is no result
            if (!items) {
                return [];
            }

            return items;
        }
    }
};


/**
 * Create sheet, set headers and validation data
 * @param {string} type
 */
function createSheet(type) {
    var callApi = api[type];

    callApi.init('createSheet', function() {
        sheet
            .init('initSheet', 'GAM: ' + callApi.name, callApi.config)
            .buildSheet();
    });
}

/**
 * Generate report from a certain apiType for the given account(s)
 * @param {array} accounts
 * @param {string} apiType
 */
function generateReport(accounts, apiType) {
    var callApi = api[apiType];

    callApi.init('generateReport', function() {
        callApi.getData(function(data) {
            // Verify if data is present
            if (data[0] === undefined || !data[0].length) {
                throw new Error('No data found for ' + type + ' in ' + account.name + '.');
            }

            sheet
                .init('initSheet', 'GAM: ' + callApi.name, callApi.config, data)
                .buildData();
        });
    }, {'accounts': accounts});
}
// TODO: finalize & cleanup function
// TODO: install change detection trigger programatically
function onChangeValidation(event) {
    var activeSheet = event.source.getActiveSheet();
    var sheetName = activeSheet.getName();
    var activeRange = activeSheet.getActiveRange();
    var numRows = activeRange.getNumRows();
    var firstRow = activeRange.getLastRow() - numRows + 1;
    var lastColumn = activeSheet.getLastColumn();
    var rowValues = activeSheet.getRange(firstRow, 1, numRows, lastColumn).getValues();
    var callApi = api[getApiTypeBySheetName(sheetName)];

    callApi.init('getConfig', function() {
        sheet
            .init('validateData', sheetName, callApi.config, rowValues)
            .validateData();
    });
}

/**
 * Insert / update data marked for inclusion from the active sheet to Google Analytics
 */
function insertData() {
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheetName = activeSheet.getName();
    var sheetRange = activeSheet.getDataRange();
    var sheetColumns = sheetRange.getNumColumns();
    var sheetRows = sheetRange.getNumRows();
    // Get the values of the sheet excluding the header row
    var sheetDataRange = activeSheet.getRange(2, 1, sheetRows, sheetColumns).getValues();
    var markedDataRange = [];
    var callApi = api[getApiTypeBySheetName(sheetName)];

    // Iterate over the rows in the sheetDataRange
    for (var r = 0; r < sheetDataRange.length; r++) {
        // Only process rows marked for inclusion
        if (sheetDataRange[r][0] == 'Yes') {
            // Add rows to array markedDataRange
            markedDataRange.push(sheetDataRange[r]);
        }
    }

    if (markedDataRange.length > 0) {
        // TODO: Add try catch + feedback to the user
        callApi.init('insertData', function() {
            callApi.prepareInsertData(markedDataRange);
        });
    }
    else {
        displayMessageToUser('Please mark at least on row for inclusion.');
    }
}
/**
 * Retrieve Account Summary from the Google Analytics Management API
 * @return {array}
 */
function getAccountSummary() {
    var items = api.accountSummaries.requestAccountSummaryList();
    return JSON.stringify(items, ['name', 'id', 'webProperties', 'profiles']);
}
