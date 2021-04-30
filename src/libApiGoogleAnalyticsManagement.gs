/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 * @version 1.5
 * @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
 *
 ******************
 * Google Analytics Management API functions
 ******************
 */


 /**
  * Return the sheet settings
  * @param {array} array
  * @returns {array}
  */
 function createApiSheetColumnConfigArray(array) {
     return {
         names: [array.map(function(element) {
             return element.name;
         })],
         namesInApi: [array.map(function(element) {
             return element.nameInApi;
         })],
         fieldType: [array.map(function(element) {
             return element.fieldType;
         })],
         colors: [array.map(function(element) {
             return element.color || colors.primary;
         })],
         dataValidation: array.map(function(element) {
             if (element.dataValidation) {
                 return element.dataValidation;
             } else {
                 return;
             }
         }),
         regexValidation: array.map(function(element) {
             return element.regexValidation;
         })
     };
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
  * Helper function to retrieve the columnIndex based on the column apiName
  * @param {string} apiType
  * @param {string} nameInApi
  * @return {integer}
  */
 function getApiColumnIndexByName(apiType, nameInApi) {
     var columns = api[apiType].sheetColumnConfig().namesInApi;
     return columns[0].indexOf(nameInApi) + 1;
 }


 /**
  * Helper function to retrieve the columnIndexRange for a given apiType based on the fieldType
  * @param {string} apiType
  * @param {string} fieldType
  * @return {object}
  */
 function getApiColumnIndexRangeByType(apiType, fieldType) {
     var columns = api[apiType].sheetColumnConfig().fieldType;
     var firstPosition = columns[0].indexOf(fieldType) + 1;
     var lastPosition = columns[0].lastIndexOf(fieldType) + 1;
     return [firstPosition, lastPosition];
 }


 /**
  * Get the available reports from the api
  * @return {array} with available reports
  */
 function getReports() {
     var arr = [];
     for (var p in api) {
         var o = {
             name: api[p].name,
             id: p
         };
         if (api[p].availableForAudit == true) {
             arr.push(o);
         }
     }
     return JSON.stringify(arr);
 }


 /**
  * Save the GA data from the selected report
  * @return {string} account id's
  * @return {string} report
  */
 function saveReportDataFromSidebar_(data) {
     var parsed = JSON.parse(data);
     return generateReport(parsed.ids, parsed.report);
 }


 /**
  * API Data
  */
 var api = {
     properties: {
         name: 'Properties',
         availableForAudit: true,
         init: function(type, cb, options) {
             this.config = this.sheetColumnConfig();

             switch (type) {
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
         /**
          * Column and field related config properties
          */
         sheetColumnConfig: function() {
             var data = [{
                 name: 'Include',
                 nameInApi: 'include',
                 fieldType: 'system',
                 dataValidation: [
                     'Yes',
                     'No'
                 ]
             }, {
                 name: 'Account Name',
                 nameInApi: 'accountName',
                 fieldType: 'account',
             }, {
                 name: 'Account ID',
                 nameInApi: 'accountId',
                 fieldType: 'account',
             }, {
                 name: 'Name',
                 nameInApi: 'name',
                 fieldType: 'property',
                 regexValidation: /.*\S.*/
             }, {
                 name: 'ID',
                 nameInApi: 'id',
                 fieldType: 'property',
                 regexValidation: /(UA|YT|MO)-\d+-\d+/
             }, {
                 name: 'Industry',
                 nameInApi: 'industryVertical',
                 fieldType: 'property',
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
             }, {
                 name: 'Default View ID',
                 nameInApi: 'defaultProfileId',
                 fieldType: 'property',
             }, {
                 name: 'Starred',
                 nameInApi: 'starred',
                 fieldType: 'property',
                 dataValidation: [
                     'TRUE',
                     'FALSE'
                 ]
             }, {
                 name: 'Website URL',
                 nameInApi: 'websiteUrl',
                 fieldType: 'property',
                 regexValidation: /^(?:(?:https?|ftp):\/\/)(?:\S+(?::\S*)?@)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,}))\.?)(?::\d{2,5})?(?:[/?#]\S*)?$/i
             }];
             return createApiSheetColumnConfigArray(data);
         },
         listApiData: function(account, cb) {
             //TODO: add try catch
             var propertiesList = Analytics.Management.Webproperties.list(account).getItems();

             if (typeof cb === 'function') {
                 return cb.call(this, propertiesList);
             } else {
                 return propertiesList;
             }
         },
         getApiData: function(account, id, cb) {
             var result;

             try {
                 result = Analytics.Management.Webproperties.get(account, id);
             } catch (e) {
                 result = e;
                 registerGoogleAnalyticsHit('exception', 'properties', false, 'getApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         insertApiData: function(
             accountId,
             name,
             industryVertical,
             starred,
             websiteUrl,
             cb
         ) {
             var values = {};
             if (name) {
                 values.name = name;
             }
             if (industryVertical) {
                 values.industryVertical = industryVertical;
             }
             if (starred) {
                 values.starred = starred;
             }
             if (websiteUrl) {
                 values.websiteUrl = websiteUrl;
             }
             var result = {};

             try {
                 result.call = Analytics.Management.Webproperties.insert(values, accountId);
                 if (isObject(result.call)) {
                     result.status = 'Success';
                     var insertedData = [
                         result.call.name,
                         result.call.id,
                         result.call.industryVertical,
                         result.call.defaultProfileId,
                         result.call.starred,
                         result.call.websiteUrl
                     ];
                     insertedData = replaceUndefinedInArray(insertedData, '');
                     insertedData = replaceNullInArray(insertedData, '');
                     result.insertedData = [insertedData];
                     //TODO: defining the type should be improved (not hardcoded?)
                     result.insertedDataType = 'property';
                     result.message = 'Success: ' + name + ' (' + result.call.id + ') has been inserted';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'properties', false, 'insertApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         updateApiData: function(
             accountId,
             name,
             id,
             industryVertical,
             defaultProfileId,
             starred,
             websiteUrl,
             cb
         ) {
             var values = {};
             if (name) {
                 values.name = name;
             }
             if (industryVertical) {
                 values.industryVertical = industryVertical;
             }
             if (defaultProfileId) {
                 values.defaultProfileId = defaultProfileId;
             }
             if (starred) {
                 values.starred = starred;
             }
             if (websiteUrl) {
                 values.websiteUrl = websiteUrl;
             }

             var result = {};

             try {
                 result.call = Analytics.Management.Webproperties.update(values, accountId, id);
                 if (isObject(result.call)) {
                     result.status = 'Success';
                     result.message = 'Success: ' + name + ' (' + id + ') has been updated';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'properties', false, 'updateApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         getData: function(cb) {
             var results = [];

             this.account.forEach(function(account) {
                 this.listApiData(account.id, function(propertiesList) {
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
         },
         insertData: function(insertData) {

             var accountId = insertData[2];
             var name = insertData[3];
             var id = insertData[4];
             var industryVertical = insertData[5];
             var defaultProfileId = insertData[6];
             var starred = insertData[7];
             var websiteUrl = insertData[8];
             var existingPropertyId = this.getApiData(accountId, id).id;
             var result = {};

             if (!id) {
                 return this.insertApiData(
                     accountId,
                     name,
                     industryVertical,
                     starred,
                     websiteUrl
                 );
             } else if (id == existingPropertyId) {
                 return this.updateApiData(
                     accountId,
                     name,
                     id,
                     industryVertical,
                     defaultProfileId,
                     starred,
                     websiteUrl
                 );
             } else if (id != existingPropertyId) {
                 result.status = 'Fail';
                 result.message = 'Property does not exist. Please verify Account and/or Property ID';
                 return result;
             }
         },
     },
     views: {
         name: 'Views',
         availableForAudit: true,
         init: function(type, cb, options) {
             this.config = this.sheetColumnConfig();

             switch (type) {
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
         /**
          * Column and field related config properties
          */
         sheetColumnConfig: function() {
             var data = [{
                 name: 'Include',
                 nameInApi: 'include',
                 fieldType: 'system',
                 dataValidation: [
                     'Yes',
                     'No'
                 ]
             }, {
                 name: 'Account Name',
                 nameInApi: 'accountName',
                 fieldType: 'account',
             }, {
                 name: 'Account ID',
                 nameInApi: 'accountId',
                 fieldType: 'account',
             }, {
                 name: 'Property Name',
                 nameInApi: 'propertyName',
                 fieldType: 'property',
             }, {
                 name: 'Property ID',
                 nameInApi: 'webPropertyId',
                 fieldType: 'property',
                 regexValidation: /(UA|YT|MO)-\d+-\d+/
             }, {
                 name: 'Name',
                 nameInApi: 'name',
                 fieldType: 'view',
             }, {
                 name: 'ID',
                 nameInApi: 'id',
                 fieldType: 'view',
             }, {
                 name: 'Bot Filtering Enabled',
                 nameInApi: 'botFilteringEnabled',
                 fieldType: 'view',
                 dataValidation: [
                     'TRUE',
                     'FALSE'
                 ]
             }, {
                 name: 'Currency',
                 nameInApi: 'currency',
                 fieldType: 'view',
                 dataValidation: [
                     'AED',
                     'ARS',
                     'AUD',
                     'BGN',
                     'BOB',
                     'BRL',
                     'CAD',
                     'CHF',
                     'CLP',
                     'CNY',
                     'COP',
                     'CZK',
                     'DKK',
                     'EGP',
                     'EUR',
                     'GBP',
                     'HKD',
                     'HRK',
                     'HUF',
                     'IDR',
                     'ILS',
                     'INR',
                     'JPY',
                     'KRW',
                     'LTL',
                     'LVL',
                     'MAD',
                     'MXN',
                     'MYR',
                     'NOK',
                     'NZD',
                     'PEN',
                     'PKR',
                     'PHP',
                     'PLN',
                     'RON',
                     'RSD',
                     'RUB',
                     'SAR',
                     'SEK',
                     'SGD',
                     'THB',
                     'TRY',
                     'TWD',
                     'UAH',
                     'USD',
                     'VEF',
                     'VND',
                     'ZAR'
                 ],
                 regexValidation: /.*\S.*/
             }, {
                 name: 'eCommerce Tracking',
                 nameInApi: 'eCommerceTracking',
                 fieldType: 'view',
                 dataValidation: [
                     'TRUE',
                     'FALSE'
                 ]
             }, {
                 name: 'Exclude Query Params',
                 nameInApi: 'excludeQueryParameters',
                 fieldType: 'view',
             }, {
                 name: 'Site Search Category Params',
                 nameInApi: 'siteSearchCategoryParameters',
                 fieldType: 'view',
             }, {
                 name: 'Site Search Query Params',
                 nameInApi: 'siteSearchQueryParameters',
                 fieldType: 'view',
             }, {
                 name: 'Strip Site Search Category Params',
                 nameInApi: 'stripSiteSearchCategoryParameters',
                 fieldType: 'view',
                 dataValidation: [
                     'TRUE',
                     'FALSE'
                 ]
             }, {
                 name: 'Strip Site Search Query Params',
                 nameInApi: 'stripSiteSearchQueryParameters',
                 fieldType: 'view',
                 dataValidation: [
                     'TRUE',
                     'FALSE'
                 ]
             }, {
                 name: 'Timezone',
                 nameInApi: 'timezone',
                 fieldType: 'view',
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
                     'America/Buenos_Aires',
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
                     'Asia/Calcutta',
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
                     'Etc/GMT',
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
             }, {
                 name: 'Type',
                 nameInApi: 'type',
                 fieldType: 'view',
                 dataValidation: [
                     'WEB',
                     'APP'
                 ],
                 regexValidation: /.*\S.*/
             }, {
                 name: 'Website URL',
                 nameInApi: 'websiteUrl',
                 fieldType: 'view',
                 regexValidation: /^(?:(?:https?|ftp):\/\/)(?:\S+(?::\S*)?@)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,}))\.?)(?::\d{2,5})?(?:[/?#]\S*)?$/i
             }];
             return createApiSheetColumnConfigArray(data);
         },
         listApiData: function(account, property, cb) {
             //TODO: add try catch
             var viewsList = Analytics.Management.Profiles.list(account, property).getItems();

             if (typeof cb === 'function') {
                 return cb.call(this, viewsList);
             } else {
                 return viewsList;
             }
         },
         getApiData: function(account, property, id, cb) {
             var result;

             try {
                 result = Analytics.Management.Profiles.get(account, property, id);
             } catch (e) {
                 result = e;
                 registerGoogleAnalyticsHit('exception', 'views', false, 'getApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         insertApiData: function(
             accountId,
             propertyId,
             name,
             botFilteringEnabled,
             currency,
             eCommerceTracking,
             excludeQueryParameters,
             siteSearchCategoryParameters,
             siteSearchQueryParameters,
             stripSiteSearchCategoryParameters,
             stripSiteSearchQueryParameters,
             timezone,
             type,
             websiteUrl,
             cb
         ) {
             var values = {};
             if (name) {
                 values.name = name;
             }
             if (botFilteringEnabled) {
                 values.botFilteringEnabled = botFilteringEnabled;
             }
             if (currency) {
                 values.currency = currency;
             }
             if (eCommerceTracking) {
                 values.eCommerceTracking = eCommerceTracking;
             }
             if (excludeQueryParameters) {
                 values.excludeQueryParameters = excludeQueryParameters;
             }
             if (siteSearchCategoryParameters) {
                 values.siteSearchCategoryParameters = siteSearchCategoryParameters;
             }
             if (siteSearchQueryParameters) {
                 values.siteSearchQueryParameters = siteSearchQueryParameters;
             }
             if (stripSiteSearchCategoryParameters) {
                 values.stripSiteSearchCategoryParameters = stripSiteSearchCategoryParameters;
             }
             if (stripSiteSearchQueryParameters) {
                 values.stripSiteSearchQueryParameters = stripSiteSearchQueryParameters;
             }
             if (timezone) {
                 values.timezone = timezone;
             }
             if (type) {
                 values.type = type;
             }
             if (websiteUrl) {
                 values.websiteUrl = websiteUrl;
             }
             var result = {};

             try {
                 result.call = Analytics.Management.Profiles.insert(values, accountId, propertyId);
                 if (isObject(result.call)) {
                     result.status = 'Success';
                     var insertedData = [
                         result.call.name,
                         result.call.id,
                         result.call.botFilteringEnabled,
                         result.call.currency,
                         result.call.eCommerceTracking,
                         result.call.excludeQueryParameters,
                         result.call.siteSearchCategoryParameters,
                         result.call.siteSearchQueryParameters,
                         result.call.stripSiteSearchCategoryParameters,
                         result.call.stripSiteSearchQueryParameters,
                         result.call.timezone,
                         result.call.type,
                         result.call.websiteUrl
                     ];
                     insertedData = replaceUndefinedInArray(insertedData, '');
                     insertedData = replaceNullInArray(insertedData, '');
                     result.insertedData = [insertedData];
                     //TODO: defining the type should be improved (not hardcoded?)
                     result.insertedDataType = 'view';
                     result.message = 'Success: ' + name + ' (' + result.call.id + ') from ' + propertyId + ' has been inserted';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'views', false, 'insertApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         updateApiData: function(
             accountId,
             propertyId,
             name,
             viewId,
             botFilteringEnabled,
             currency,
             eCommerceTracking,
             excludeQueryParameters,
             siteSearchCategoryParameters,
             siteSearchQueryParameters,
             stripSiteSearchCategoryParameters,
             stripSiteSearchQueryParameters,
             timezone,
             type,
             websiteUrl,
             cb
         ) {
             var values = {};
             if (name) {
                 values.name = name;
             }
             if (botFilteringEnabled) {
                 values.botFilteringEnabled = botFilteringEnabled;
             }
             if (currency) {
                 values.currency = currency;
             }
             if (eCommerceTracking) {
                 values.eCommerceTracking = eCommerceTracking;
             }
             if (excludeQueryParameters) {
                 values.excludeQueryParameters = excludeQueryParameters;
             }
             if (siteSearchCategoryParameters) {
                 values.siteSearchCategoryParameters = siteSearchCategoryParameters;
             }
             if (siteSearchQueryParameters) {
                 values.siteSearchQueryParameters = siteSearchQueryParameters;
             }
             if (stripSiteSearchCategoryParameters) {
                 values.stripSiteSearchCategoryParameters = stripSiteSearchCategoryParameters;
             }
             if (stripSiteSearchQueryParameters) {
                 values.stripSiteSearchQueryParameters = stripSiteSearchQueryParameters;
             }
             if (timezone) {
                 values.timezone = timezone;
             }
             if (type) {
                 values.type = type;
             }
             if (websiteUrl) {
                 values.websiteUrl = websiteUrl;
             }

             var result = {};

             try {
                 result.call = Analytics.Management.Profiles.update(values, accountId, propertyId, viewId);
                 if (isObject(result.call)) {
                     result.status = 'Success';
                     result.message = 'Success: ' + name + ' (' + viewId + ') from ' + propertyId + ' has been updated';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'views', false, 'updateApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         getData: function(cb) {
             var results = [];

             this.account.forEach(function(account) {
                 try {
                     account.webProperties.forEach(function(property) {
                         this.listApiData(account.id, property.id, function(viewsList) {
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
                 } catch (e) {
                     registerGoogleAnalyticsHit('exception', 'views', false, 'getData failed to execute: ' + e);
                 }
             }, this);

             cb(results);
         },
         insertData: function(insertData) {

             var accountId = insertData[2];
             var propertyId = insertData[4];
             var name = insertData[5];
             var viewId = insertData[6];
             var botFilteringEnabled = insertData[7];
             var currency = insertData[8];
             var eCommerceTracking = insertData[9];
             var excludeQueryParameters = insertData[10];
             var siteSearchCategoryParameters = insertData[11];
             var siteSearchQueryParameters = insertData[12];
             var stripSiteSearchCategoryParameters = insertData[13];
             var stripSiteSearchQueryParameters = insertData[14];
             var timezone = insertData[15];
             var type = insertData[16];
             var websiteUrl = insertData[17];
             var existingViewId = this.getApiData(accountId, propertyId, viewId).id;
             var result = {};

             if (!viewId) {
                 return this.insertApiData(
                     accountId,
                     propertyId,
                     name,
                     botFilteringEnabled,
                     currency,
                     eCommerceTracking,
                     excludeQueryParameters,
                     siteSearchCategoryParameters,
                     siteSearchQueryParameters,
                     stripSiteSearchCategoryParameters,
                     stripSiteSearchQueryParameters,
                     timezone,
                     type,
                     websiteUrl
                 );
             } else if (viewId == existingViewId) {
                 return this.updateApiData(
                     accountId,
                     propertyId,
                     name,
                     viewId,
                     botFilteringEnabled,
                     currency,
                     eCommerceTracking,
                     excludeQueryParameters,
                     siteSearchCategoryParameters,
                     siteSearchQueryParameters,
                     stripSiteSearchCategoryParameters,
                     stripSiteSearchQueryParameters,
                     timezone,
                     type,
                     websiteUrl
                 );
             } else if (viewId != existingViewId) {
                 result.status = 'Fail';
                 result.message = 'View does not exist. Please verify Account, Property and/or View ID';
                 return result;
             }
         },
     },
     filterLinks: {
         name: 'Filter Links',
         availableForAudit: true,
         init: function(type, cb, options) {
             this.config = this.sheetColumnConfig();

             switch (type) {
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
         /**
          * Column and field related config properties
          */
         sheetColumnConfig: function() {
             var data = [{
                 name: 'Include',
                 nameInApi: 'include',
                 fieldType: 'system',
                 dataValidation: [
                     'Yes',
                     'No'
                 ]
             }, {
                 name: 'Account Name',
                 nameInApi: 'accountName',
                 fieldType: 'account',
             }, {
                 name: 'Account ID',
                 nameInApi: 'profileRefAccountId',
                 fieldType: 'account',
             }, {
                 name: 'Property Name',
                 nameInApi: 'profileRefWebPropertyName',
                 fieldType: 'property',
             }, {
                 name: 'Property ID',
                 nameInApi: 'profileRefWebPropertyId',
                 fieldType: 'property',
                 regexValidation: /(UA|YT|MO)-\d+-\d+/
             }, {
                 name: 'View Name',
                 nameInApi: 'profileRefName',
                 fieldType: 'view',
             }, {
                 name: 'View ID',
                 nameInApi: 'profileRefId',
                 fieldType: 'view',
             }, {
                 name: 'Filter Link ID',
                 nameInApi: 'id',
                 fieldType: 'filterLink',
                 regexValidation: /^[0-9]+$/
             }, {
                 name: 'Filter Link Rank',
                 nameInApi: 'rank',
                 fieldType: 'filterLink',
                 // Allow integer 0-255
                 regexValidation: /^([0-9]{1}|[1-9]{1,2}|1[0-9]{2}|2[0-4][0-9]|25[0-5])$/
             }, {
                 name: 'Filter Name',
                 nameInApi: 'filterRef.name',
                 fieldType: 'filterLink',
             }, {
                 name: 'Filter ID',
                 nameInApi: 'filterRef.id',
                 fieldType: 'filterLink',
                 regexValidation: /^[0-9]+$/
             }];
             return createApiSheetColumnConfigArray(data);
         },
         listApiData: function(account, property, view, cb) {
             //TODO: add try catch
             var flList = Analytics.Management.ProfileFilterLinks.list(account, property, view).getItems();

             if (typeof cb === 'function') {
                 return cb.call(this, flList);
             } else {
                 return flList;
             }
         },
         getApiData: function(account, property, view, id, cb) {
             var result;

             try {
                 result = Analytics.Management.ProfileFilterLinks.get(account, property, view, id);
             } catch (e) {
                 result = e;
                 registerGoogleAnalyticsHit('exception', 'filterLinks', false, 'getApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         insertApiData: function(
             accountId,
             propertyId,
             viewId,
             filterLinkFilterRefId,
             filterLinkRank,
             cb
         ) {
             var values = {};
             values.filterRef = {};
             if (filterLinkFilterRefId) {
                 values.filterRef.id = filterLinkFilterRefId;
             }
             if (filterLinkRank) {
                 values.rank = filterLinkRank;
             }
             var result = {};

             try {
                 result.call = Analytics.Management.ProfileFilterLinks.insert(values, accountId, propertyId, viewId);
                 if (isObject(result.call)) {
                     result.status = 'Success';
                     var insertedData = [
                         result.call.id,
                         result.call.rank,
                         result.call.filterRef.name,
                         result.call.filterRef.id
                     ];
                     insertedData = replaceUndefinedInArray(insertedData, '');
                     insertedData = replaceNullInArray(insertedData, '');
                     result.insertedData = [insertedData];
                     //TODO: defining the type should be improved (not hardcoded?)
                     result.insertedDataType = 'filterLink';
                     result.message = 'Success: ' + result.call.id + ' for ' + viewId + ' has been inserted';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'filterLinks', false, 'insertApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         updateApiData: function(
             accountId,
             propertyId,
             viewId,
             filterLinkId,
             filterLinkRank,
             cb
         ) {
             var values = {};
             if (filterLinkRank) {
                 values.rank = filterLinkRank;
             }

             var result = {};

             try {
                 result.call = Analytics.Management.ProfileFilterLinks.update(values, accountId, propertyId, viewId, filterLinkId);
                 if (isObject(result.call)) {
                     result.status = 'Success';
                     result.message = 'Success: ' + filterLinkId + ' from ' + viewId + ' has been updated';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'filterLinks', false, 'updateApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         getData: function(cb) {
             var results = [];

             this.account.forEach(function(account) {
                 try {
                     account.webProperties.forEach(function(property) {
                         try {
                             property.profiles.forEach(function(view) {
                                 this.listApiData(account.id, property.id, view.id, function(flList) {
                                     flList.forEach(function(fl) {
                                         var defaults = [
                                             '',
                                             account.name,
                                             account.id,
                                             property.name,
                                             property.id,
                                             view.name,
                                             view.id,
                                             fl.id,
                                             fl.rank,
                                             fl.filterRef.name,
                                             fl.filterRef.id
                                         ];
                                         defaults = replaceUndefinedInArray(defaults, '');

                                         results.push(defaults);
                                     }, this);
                                 });
                             }, this);
                         } catch (e) {
                         registerGoogleAnalyticsHit('exception', 'filterLinks', false, 'getData failed to execute: ' + e);
                     }
                     }, this);
                 } catch (e) {
                 registerGoogleAnalyticsHit('exception', 'filterLinks', false, 'getData failed to execute: ' + e);}
             }, this);

             cb(results);
         },
         insertData: function(insertData) {

             var accountId = insertData[2];
             var propertyId = insertData[4];
             var viewId = insertData[6];
             var filterLinkId = insertData[7];
             var filterLinkRank = insertData[8]
             var filterLinkFilterRefName = insertData[9];
             var filterLinkFilterRefId = insertData[10];
             var existingFilterLinkId = this.getApiData(accountId, propertyId, viewId, filterLinkId).id;
             var result = {};

             if (!filterLinkId) {
                 return this.insertApiData(
                     accountId,
                     propertyId,
                     viewId,
                     filterLinkFilterRefId,
                     filterLinkRank
                 );
             } else if (filterLinkId == existingFilterLinkId) {
                 return this.updateApiData(
                     accountId,
                     propertyId,
                     viewId,
                     filterLinkId,
                     filterLinkRank
                 );
             } else if (filterLinkId != existingFilterLinkId) {
                 result.status = 'Fail';
                 result.message = 'FilterLink does not exist. Please verify Account, Property, View and/or FilterLink ID';
                 return result;
             }
         },
     },
     customDimensions: {
         name: 'Custom Dimensions',
         availableForAudit: true,
         init: function(type, cb, options) {
             this.config = this.sheetColumnConfig();

             switch (type) {
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
         /**
          * Column and field related config properties
          */
         sheetColumnConfig: function() {
             var data = [{
                 name: 'Include',
                 nameInApi: 'include',
                 fieldType: 'system',
                 dataValidation: [
                     'Yes',
                     'No'
                 ]
             }, {
                 name: 'Account Name',
                 nameInApi: 'accountName',
                 fieldType: 'account',
             }, {
                 name: 'Account ID',
                 nameInApi: 'accountId',
                 fieldType: 'account',
             }, {
                 name: 'Property Name',
                 nameInApi: 'webPropertyName',
                 fieldType: 'property',
             }, {
                 name: 'Property ID',
                 nameInApi: 'webPropertyId',
                 fieldType: 'property',
                 regexValidation: /(UA|YT|MO)-\d+-\d+/
             }, {
                 name: 'Name',
                 nameInApi: 'name',
                 fieldType: 'customDimension',
                 regexValidation: /.*\S.*/
             }, {
                 name: 'Index',
                 nameInApi: 'index',
                 fieldType: 'customDimension',
                 regexValidation: /^[0-9]{1,3}$/
             }, {
                 name: 'Scope',
                 nameInApi: 'scope',
                 fieldType: 'customDimension',
                 dataValidation: [
                     'HIT',
                     'SESSION',
                     'USER',
                     'PRODUCT'
                 ],
                 regexValidation: /.*\S.*/
             }, {
                 name: 'active',
                 nameInApi: 'include',
                 fieldType: 'customDimension',
                 dataValidation: [
                     'TRUE',
                     'FALSE'
                 ],
                 regexValidation: /.*\S.*/
             }];
             return createApiSheetColumnConfigArray(data);
         },
         listApiData: function(account, property, cb) {
             //TODO: add try catch
             var cdList = Analytics.Management.CustomDimensions.list(account, property).getItems();

             if (typeof cb === 'function') {
                 return cb.call(this, cdList);
             } else {
                 return cdList;
             }
         },
         getApiData: function(account, property, index, cb) {
             var result;

             try {
                 result = Analytics.Management.CustomDimensions.get(account, property, index);
             } catch (e) {
                 result = e;
                 registerGoogleAnalyticsHit('exception', 'customDimensions', false, 'getApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         insertApiData: function(account, property, name, index, scope, active, cb) {
             var values = {
                 'name': name,
                 'index': index,
                 'scope': scope,
                 'active': active
             };
             var result = {};

             try {
                 result.call = Analytics.Management.CustomDimensions.insert(values, account, property);
                 if (isObject(result.call)) {
                     result.status = 'Success';
                     var insertedData = [
                         result.call.name,
                         result.call.index,
                         result.call.scope,
                         result.call.active,
                     ];
                     insertedData = replaceUndefinedInArray(insertedData, '');
                     insertedData = replaceNullInArray(insertedData, '');
                     result.insertedData = [insertedData];
                     //TODO: defining the type should be improved (not hardcoded?)
                     result.insertedDataType = 'customDimension';
                     result.message = 'Success: ' + result.call.index + ' from ' + result.call.webPropertyId + ' has been inserted.';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'customDimensions', false, 'insertApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         updateApiData: function(account, property, name, index, scope, active, cb) {
             var values = {
                 'name': name,
                 'scope': scope,
                 'active': active
             };
             var result = {};
             try {
                 result.call = Analytics.Management.CustomDimensions.update(values, account, property, index);
                 if (isObject(result.call)) {
                     result.message = 'Success: ' + result.call.index + ' from ' + result.call.webPropertyId + ' has been updated';
                 }
             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'customDimensions', false, 'updateApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         getData: function(cb) {
             var results = [];

             this.account.forEach(function(account) {
                 try {
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
                 } catch (e) {
                 registerGoogleAnalyticsHit('exception', 'customDimensions', false, 'getData failed to execute: ' + e);
             }
             }, this);

             cb(results);
         },
         insertData: function(insertData) {

             var account = insertData[2];
             var property = insertData[4];
             var name = insertData[5];
             var index = insertData[6];
             var scope = insertData[7];
             var active = insertData[8];
             var existingData = this.getApiData(account, property, index);

             if (existingData && existingData !== false) {
                 return this.updateApiData(account, property, name, index, scope, active);
             } else {
                 return this.insertApiData(account, property, name, index, scope, active);
             }

         },
     },
     accountSummaries: {
         name: 'Account Summaries',
         availableForAudit: false,
         requestAccountSummaryList: function() {
             var items = Analytics.Management.AccountSummaries.list().getItems();

             // Return an empty array if there is no result
             if (!items) {
                 return [];
             }

             return items;
         }
     },
     searchConsoleSites: {
         name: 'Search Console Sites',
         availableForAudit: false,
         init: function(type, cb, options) {
             this.config = this.sheetColumnConfig();

             switch (type) {
                 case 'createSheet':
                     break;
                 case 'generateReport':
                     break;
                 case 'getConfig':
                     break;
                 case 'insertData':
                     break;
             }

             cb();

             return this;
         },
         /**
          * Column and field related config properties
          */
         sheetColumnConfig: function() {
             var data = [{
                 name: 'Include',
                 nameInApi: 'include',
                 fieldType: 'system',
                 dataValidation: [
                     'Yes',
                     'No'
                 ]
             }, {
                 name: 'Site URL',
                 nameInApi: 'siteUrl',
                 fieldType: 'site',
             }, {
                 name: 'Permission Level',
                 nameInApi: 'permissionLevel',
                 fieldType: 'site',
             }];
             return createApiSheetColumnConfigArray(data);
         },
         listApiData: function(cb) {
             var service = googleSearchConsoleService();

             var apiURL = 'https://www.googleapis.com/webmasters/v3/sites/';

             var headers = {
                 'Authorization': 'Bearer ' + googleSearchConsoleService().getAccessToken()
             };

             var options = {
                 'headers': headers,
                 'method' : 'GET',
                 'muteHttpExceptions': true
             };

             var apiResult = UrlFetchApp.fetch(apiURL, options);
             var items = JSON.parse(apiResult).siteEntry;

             var result = [];

             for (var i = 0; i < items.length; i++) {
                 result.push(['', items[i].siteUrl, items[i].permissionLevel]);
             }

             //TODO: add try catch

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         addApiData: function(siteUrl, cb) {
             var service = googleSearchConsoleService();
             var result = {};

             try {
                var apiURL = 'https://www.googleapis.com/webmasters/v3/sites/';

                var headers = {
                  'Authorization': 'Bearer ' + googleSearchConsoleService().getAccessToken()
                };

                var options = {
                  'headers': headers,
                  'method' : 'PUT',
                  'muteHttpExceptions': true
                };

                var urlClean = encodeURIComponent(siteUrl);

                result.call = UrlFetchApp.fetch(apiURL + urlClean, options);

                if (result.call.getResponseCode() == 204) {
                    result.status = 'Success';
                    result.insertedDataType = 'searchConsoleSite';
                    result.message = 'Success: ' + siteUrl + ' has been added';
                } else if (JSON.parse(result.call).error.code == '400') {
                    var resultObject = JSON.parse(result.call);
                    result.status = 'Fail';
                    result.message = 'Error ' + resultObject.error.code + ': ' + resultObject.error.message;
                    registerGoogleAnalyticsHit('exception', 'searchConsoleSites', false, 'addApiData failed to execute: ' + result.message);
                } else {
                    return false;
                }

             } catch (e) {
                 result.status = 'Fail';
                 result.message = e;
                 registerGoogleAnalyticsHit('exception', 'searchConsoleSites', false, 'addApiData failed to execute: ' + e);
             }

             if (typeof cb === 'function') {
                 return cb.call(this, result);
             } else {
                 return result;
             }
         },
         listData: function(cb) {
             var results;

             this.listApiData(function(siteList) {
                 results = siteList;
             });

             cb(results);
         },
         insertData: function(insertData) {

             var siteUrl = insertData[1];

             return this.addApiData(siteUrl);

         },
     },
 };



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
                 throw new Error('No data found for ' + apiType + ' in ' + accounts[0].name + '.');
             }

             sheet
                 .init('initSheet', 'GAM: ' + callApi.name, callApi.config, data)
                 .buildData();
         });
     }, {
         'accounts': accounts
     });
 }


 /**
  * Validate the changed data
  * @param {string} event
  */
 function onChangeValidation(event) {
     // TODO: finalize & cleanup function
     // TODO: install change detection trigger programatically
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
  * @return Displays a message to the user
  */
 function insertData() {
     var resultMessages = [];
     var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
     var sheetName = activeSheet.getName();
     var sheetRange = activeSheet.getDataRange();
     var sheetColumns = sheetRange.getNumColumns();
     var sheetRows = sheetRange.getNumRows();
     // Get the values of the sheet excluding the header row
     var sheetDataRange = activeSheet.getRange(2, 1, sheetRows, sheetColumns).getValues();
     var markedDataRange = [];
     var apiType = getApiTypeBySheetName(sheetName);
     var callApi = api[apiType];
     var result;

     registerGoogleAnalyticsHit('event', apiType, 'Click', 'Menu');

     // Iterate over the rows in the sheetDataRange
     sheetDataRange.forEach(function(rowArray, rowId) {
         // Define the real row ID in order to provide feedback to the user
         var realRowId = rowId + 2;

         // Only process rows marked for inclusion
         if (rowArray[0] == 'Yes') {

             callApi.init('insertData', function() {
                 result = callApi.insertData(rowArray);
                 resultMessages.push('\nRow ' + realRowId + ': ' + result.message);

                 // Write the specified data to the sheet
                 if (result.dataToUpdate) {
                     var dataToUpdate = result.dataToUpdate;

                     for (var key in dataToUpdate) {
                         var columnId = getApiColumnIndexByName(apiType, key);
                         var dataCell = activeSheet.getRange(realRowId, columnId, 1, 1);
                         dataCell.setValue(dataToUpdate[key]);
                     }
                 }

                 // Write insertedData back to the sheet
                 if (result.insertedData) {
                     var colRange = getApiColumnIndexRangeByType(apiType, result.insertedDataType);
                     var dataRange = activeSheet.getRange(realRowId, colRange[0], 1, colRange[1] - colRange[0] + 1);
                     dataRange.setValues(result.insertedData);
                 }

             });

         }
     });

     if (resultMessages.length > 0) {
         registerGoogleAnalyticsHit('event', apiType, 'insertData', resultMessages);
         ui.alert('Results', resultMessages, ui.ButtonSet.OK);
     } else {
         registerGoogleAnalyticsHit('event', apiType, 'insertData', 'Error: Please mark at least on row for inclusion.');
         ui.alert('Error', 'Please mark at least on row for inclusion.', ui.ButtonSet.OK);
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
