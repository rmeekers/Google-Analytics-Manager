/** Google Analytics Manager
 * Manage your Google Analytics account in batch via a Google Sheet
 *
 * @license GNU LESSER GENERAL PUBLIC LICENSE Version 3
 * @author Rutger Meekers [rutger@meekers.eu]
 * @version 1.2
 * @see {@link http://rutger.meekers.eu/Google-Analytics-Manager/ Project Page}
 *
 ******************
 * Helper functions
 ******************
 */


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
 * Replace 'null' in a given array with another value
 * @param {array} array
 * @param {string} value
 */
function replaceNullInArray(array, value) {
    return array.map(function(el) {
        return el === null ? value : el;
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
    if (sortAscending === undefined || sortAscending !== false) {
        sortAscending = true;
    }

    if (sortAscending) {
        array.sort(function(a, b) {
            return a[objectParamToSortBy] > b[objectParamToSortBy];
        });
    } else {
        array.sort(function(a, b) {
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
        } else {
            return (array[sortIndex] < b[sortIndex]) ? -1 : 1;
        }
    }
    return array;
}


/**
 * Test if a given parameter is an object
 *
 * @param obj - Object to test
 * @return boolean
 */
function isObject(obj) {
    return obj === Object(obj);
}


/**
 * Helper function to transpose an array
 * @param {array} a
 * @return {array}
 */
function transposeArray(a) {
    return a[0].map(function(_, c) {
        return a.map(function(r) {
            return r[c];
        });
    });
}


/**
 * Toast wrapper function
 * @customfunction
 *
 * @param {string} title Message titel
 * @param {string} message Message text
 */
function helperToast_(title, text) {
    SpreadsheetApp.getActiveSpreadsheet().toast(text, title, 3);
}
