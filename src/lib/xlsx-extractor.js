import XlsxUtil from './xlsx-util.js';

/**
 * Defines the file path in the XLSX.
 * @type {Object}
 */
const FilePaths = {
  SheetBase: 'xl/worksheets/sheet',
  SharedStrings: 'xl/sharedStrings.xml'
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
      if( index < 0 && this._count <= index ) {
        throw new Error( 'The index is out of range.' );
      }

      return this._parseXML( index );
    } )
    .then( ( xmls ) => {
      return this._extract( xmls );
    } );
  }

  /**
   * Extract a sheet.
   *
   * @param {Object} xmls [description]
   *
   * @return {Promise} Extract task.
   */
  _extract( xmls ) {
    return new Promise( ( resolve ) => {
      const cells   = XlsxUtil.getCells( xmls.sheet.worksheet.sheetData[ 0 ].row );
      const size    = XlsxUtil.getSheetSize( xmls.sheet, cells );
      const rows    = ( size.row.max - size.row.min ) + 1;
      const cols    = ( size.col.max - size.col.min ) + 1;
      const data    = XlsxUtil.createEmptyCells( rows, cols );
      const strings = xmls.strings.sst.si;

      cells.forEach( ( cell ) => {
        let value = cell.value;
        if( cell.type === 's' ) {
          const index  = parseInt( value, 10 );
          value = XlsxUtil.valueFromStrings( strings[ index ] );
        }

        data[ cell.row - size.row.min ][ cell.col - size.col.min ] = value;
      } );

      resolve( data );
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
    const zip   = XlsxUtil.unzip( this._path );
    const tasks = [ XlsxUtil.parseXML( zip.files[ FilePaths.SheetBase + index + '.xml' ] ) ];
    if( zip.files[ FilePaths.SharedStrings ] ) {
      tasks.push( XlsxUtil.parseXML( zip.files[ FilePaths.SharedStrings ] ) );
    }

    return Promise
    .all( tasks )
    .then( ( xmls ) => {
      const result = {};
      if( xmls.length < 2 ) {
        result.sheet = xmls[ 0 ];
      } else if( xmls[ 0 ].worksheet ) {
        result.sheet   = xmls[ 0 ];
        result.strings = xmls[ 1 ];
      } else {
        result.sheet   = xmls[ 1 ];
        result.strings = xmls[ 0 ];
      }

      return result;
    } );
  }
}
