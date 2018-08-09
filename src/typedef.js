/**
 * @external {ZipObject} http://stuk.github.io/jszip/documentation/api_zipobject.html
 */

/**
 * The data of the sheet.
 *
 * @typedef {Object} SheetData
 * @property {String} name Sheet name.
 * @property {Object} sheet Data obtained by converting the XML of the sheet to the JavaScript Object.
 * @property {Object} strings Data obtained by converting the XML of the shared strings to the JavaScript Object.
 */

/**
 * Range of an output sheets.
 *
 * @typedef {Object} Range
 * @property {Number} begin Begin of an output sheets.
 * @property {Number} end End of an output sheets.
 */

/**
 * Commad line options.
 *
 * @typedef {Object} CLIOptions
 * @property {Boolean} help Mode to display the help text.
 * @property {Boolean} version Mode to display the version number.
 * @property {String} input Path of the XLSX file.
 * @property {Range} input Range of an output sheets.
 * @property {Boolean} count Outputs the number of sheet.
 */
