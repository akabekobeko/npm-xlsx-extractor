import XlsxUtil from './xlsx-util.js';

/**
 * Defines the file path in the XLSX.
 * @type {Object}
 */
const FilePaths = {
  WorkBook: 'xl/workbook.xml',
  SharedStrings: 'xl/sharedStrings.xml',
  SheetBase: 'xl/worksheets/sheet'
};

/**
 * The maximum number of sheets ( Excel 97 ).
 * @type {Number}
 */
const MaxSheets = 256;

/**
 * Extract the colums/rows from XLSX file.
 */
export default class XlsxExtractor {
  /**
   * Initialize instance.
   *
   * @param {String} path XLSX file path.
   *
   * @throws {Error} The specified XLSX file is invalid.
   */
  constructor( path ) {
    const zip = XlsxUtil.unzip( path );
    if( !( zip ) ) {
      throw new Error( 'Failed to expand the XLSX file.' );
    }

    const count = this._getSheetsCount( zip );
    if( count === 0 ) {
      throw new Error( 'Sheets was not found in XLSX file.' );
    }

    this._path  = path;
    this._count = count;
  }

  /**
   * Gets the number of sheets.
   *
   * @return {Number} The number of sheets.
   */
  get count() {
    return this._count;
  }

  /**
   * Extract a sheet.
   *
   * @param {Number} index Index of sheet. Range of from 1 to XlsxExtractor.count.
   *
   * @return {Promise} Extract task.
   */
  extract( index ) {
    return Promise
    .resolve()
    .then( () => {
      if( index < 1 || this._count < index ) {
        throw new Error( 'The index is out of range: ' + index );
      }

      return this._parseXML( index );
    } )
    .then( ( xmls ) => {
      return this._extract( index, xmls );
    } );
  }

  /**
   * Extract a sheet.
   *
   * @param {SheetData} data Sheet data.
   * @param {SheetData} data Sheet data.
   *
   * @return {Promise} Extract task.
   */
  _extract( id, data ) {
    return new Promise( ( resolve ) => {
      const cells   = XlsxUtil.getCells( data.sheet.worksheet.sheetData[ 0 ].row );
      const size    = XlsxUtil.getSheetSize( data.sheet, cells );
      const rows    = ( size.row.max - size.row.min ) + 1;
      const cols    = ( size.col.max - size.col.min ) + 1;
      const sheet   = XlsxUtil.createEmptyCells( rows, cols );
      const strings = data.strings.sst.si;

      cells.forEach( ( cell ) => {
        let value = cell.value;
        if( cell.type === 's' ) {
          const index  = parseInt( value, 10 );
          value = XlsxUtil.valueFromStrings( strings[ index ] );
        }

        sheet[ cell.row - size.row.min ][ cell.col - size.col.min ] = value;
      } );

      resolve( {
        id:    id,
        name:  data.name,
        sheet: sheet
      } );
    } );
  }

  /**
   * Gets the number of sheets.
   *
   * @param {ZipObject} zip Zip object.
   *
   * @return {Number} The number of sheets.
   */
  _getSheetsCount( zip ) {
    let count = 0;
    for( let i = 1; i < MaxSheets; ++i ) {
      const path = FilePaths.SheetBase + i + '.xml';
      if( !( zip.files[ path ] ) ) {
        break;
      }

      ++count;
    }

    return count;
  }

  /**
   * Parses the XML.
   *
   * @param {Number} index Index of sheet. Range of from 1 to XlsxExtractor.count.
   *
   * @return {Promise} Parse task.
   */
  _parseXML( index ) {
    const zip    = XlsxUtil.unzip( this._path );
    const result = {};

    return Promise
    .resolve()
    .then( () => {
      const xml = zip.files[ FilePaths.WorkBook ].asText();
      return XlsxUtil.parseXML( xml );
    } )
    .then( ( root ) => {
      // Get a sheet name
      if( root && root.workbook && root.workbook.sheets && 0 < root.workbook.sheets.length && root.workbook.sheets[ 0 ].sheet ) {
        root.workbook.sheets[ 0 ].sheet.some( ( sheet ) => {
          const id = Number( sheet.$.sheetId );
          if( id === index ) {
            result.name = ( sheet.$.name || '' );
            return true;
          }

          return false;
        } );
      }

      const xml = zip.files[ FilePaths.SheetBase + index + '.xml' ].asText();
      return XlsxUtil.parseXML( xml );
    } )
    .then( ( sheet ) => {
      result.sheet = sheet;

      if( zip.files[ FilePaths.SharedStrings ] ) {
        const xml = zip.files[ FilePaths.SharedStrings ].asText();
        return XlsxUtil.parseXML( xml );
      }

      return Promise.resolve();
    } )
    .then( ( strings ) => {
      if( strings ) {
        result.strings = strings;
      }

      return result;
    } );
  }
}
