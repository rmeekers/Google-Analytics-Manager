/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 * @version 1.2
 * @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
 *
 ******************
 * Spreadsheet functions
 ******************
 */


/**
 * Helper function to create a new Properties sheet from the menu
 */
function createSheetProperties() {

    registerGoogleAnalyticsHit('event', 'createSheetProperties', 'Click', 'Menu');

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

    registerGoogleAnalyticsHit('event', 'createSheetViews', 'Click', 'Menu');

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
 * Sheet / Application related functions
 ******************
 */

/**
 * onInstall runs when the script is installed
 * @param {*} e
 */
function onInstall(e) {
    onOpen(e);
}

/**
 * Install custom triggers if not already present
 * @customfunction
 */
function initTriggers() {

  var triggers = ScriptApp.getProjectTriggers();

  // If no custom triggers are present, we can safely add ours
  if(!triggers.length) {
    addEditTrigger_();

    return;
  }

  // Check present triggers for the appropriate handler and source combinations
  var add_edit = true;

  for(var i = 0; i < triggers.length; i++) {

    var trigger = triggers[i];
    var handler = trigger.getHandlerFunction();
    var source = trigger.getTriggerSource();

    switch (handler) {

      case 'onEditTrigger':

        if(source == 'SPREADSHEETS') {
          add_edit = false;
        }
        break;

      default:
    }
  }

  // Add missing triggers
  if(add_edit) {
    addEditTrigger_();
  }
}

/**
 * Main setup: add the required sheet, titles, data validation, style, etc.
 * @customfunction
 */
function setupDocument() {

    // Initialize custom triggers
    initTriggers();

    helperToast_('','Setup complete');
}


/**
 * Initialize menu items
 * (needs to be public)
 * @customfunction
 */
function loadMenu_() {

    ui = SpreadsheetApp.getUi();

    //TODO: add setupran identifier to avoid infinite loop
    //registerGoogleAnalyticsHit('event', 'onOpen', 'Open', 'Spreadsheet');

    return ui
        .createMenu('GA Manager')
        .addItem('Audit GA', 'GoogleAnalyticsManager.showSidebar')
        .addItem('Insert / Update Data from active sheet to GA', 'GoogleAnalyticsManager.insertData')
        .addSeparator()
        .addSubMenu(ui.createMenu('Advanced')
            .addItem('Insert Properties Sheet', 'GoogleAnalyticsManager.createSheetProperties')
            .addItem('Insert Views Sheet', 'GoogleAnalyticsManager.createSheetViews'))
        .addToUi();
}


/**
 * Main function to start the application.
 * Triggered by onOpen event, to be set in the document using the library.
 *
 * (needs to be public)
 * @customfunction
 */
function initMenu(event) {
  log('init app');

  // Load the application menu
  loadMenu_();
}


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
        log('sheet.init');
        this.workbook = SpreadsheetApp.getActiveSpreadsheet();
        this.name = name;
        this.sheetColumnConfig = config;

        switch (type) {
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
    /**
     * Validate data in the sheet
     */
    validate: function() {
        // TODO: do something usefull after validation
        // FIXME: return correct row number
        var results = [];

        for (var column = 0; column < this.data.length; column++) {
            var regex = this.regexValidation[column];
            // Only validate a cell if there is a regexValidation defined
            if (regex) {
                for (var row = 0; row < this.data[column].length; row++) {
                    var string = String(this.data[column][row]);
                    if (string.match(regex)) {
                        Logger.log('Validate OK: ' + string);
                    } else {
                        Logger.log('Validate NOK: ' + string);
                        results.push('\nError: data validation failed for value ' + string + 'on row ' + row + ' column ' + column);
                    }
                }
            }
        }

        /*
        if (results.length > 0) {
            ui.alert('Results', results, ui.ButtonSet.OK);
        }
        */
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

    /**
     * Build the sheet header
     */
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

    /**
     * Set the sheet dataValidation
     */
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

    /**
     * Cleanup the sheet
     */
    cleanup: function() {
        // auto resize all columns
        this.sheetColumnConfig.names[0].forEach(function(e, i) {
            this.sheet.autoResizeColumn(i + 1);
        }, this);
    },

    /**
     * Build the sheet structure
     */
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

    /**
     * Build the sheet structure and insert data
     */
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
