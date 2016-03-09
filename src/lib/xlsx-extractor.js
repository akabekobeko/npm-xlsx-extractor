import Fs from 'fs';
import Path from 'path';
import Zip from 'node-zip';
import XmlParser from 'xml2js';

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
    const zip = XlsxExtractor._unzip( path );
    if( !( zip ) ) {
      throw new Error( 'Failed to expand the XLSX file.' );
    }

    const count = XlsxExtractor._getSheetsCount( zip );
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

      return XlsxExtractor._parseXML( index );
    } )
    .then( ( xmls ) => {
      return XlsxExtractor._extract( xmls );
    } );
  }

  /**
   * Create a empty cells.
   *
   * @param {Number} rows Rows count.
   * @param {Number} cols Columns count.
   *
   * @return {Array.<Array.<String>>} Cells.
   */
  static _createEmptyCells( rows, cols ) {
    const arr = [];
    for( let i = 0; i < rows; ++i ) {
      const row = [];
      for( let j = 0; j < cols; ++j ) {
        row.push( '' );
      }

      arr.push( row );
    }

    return arr;
  }

  /**
   * Extract a sheet.
   *
   * @param {Object} xmls [description]
   *
   * @return {Promise} Extract task.
   */
  static _extract( xmls ) {
    return new Promise( ( resolve ) => {
      const cells   = XlsxExtractor._getCells( xmls.sheet.worksheet.sheetData[ 0 ].row );
      const size    = XlsxExtractor._calcSheetSize( xmls.sheet, cells );
      const rows    = ( size.row.max - size.row.min ) + 1;
      const cols    = ( size.col.max - size.col.min ) + 1;
      const data    = XlsxExtractor._createEmptyCells( rows, cols );
      const strings = xmls.strings.sst.si;

      cells.forEach( ( cell ) => {
        let value = cell.value;
        if( cell.type === 's' ) {
          const index  = parseInt( value, 10 );
          value = XlsxExtractor._valueFromStrings( strings[ index ] );
        }

        data[ cell.row - size.row.min ][ cell.col - size.col.min ] = value;
      } );

      resolve( data );
    } );
  }

  /**
   * Get a cells from a rows.
   *
   * @param {Array.<Object>} rows Rows。
   *
   * @return {Array.<Object>} Cells。
   */
  static _getCells( rows ) {
    const cells = [];
    rows
    .filter( ( row ) => {
      return ( row.c && 0 < row.c.length );
    } )
    .forEach( ( row ) => {
      row.c.forEach( ( cell ) => {
        const position = XlsxExtractor._getPosition( cell.$.r );
        cells.push( {
          row:   position.row,
          col:   position.col,
          type:  ( cell.$.t ? cell.$.t : '' ),
          value: ( cell.v && 0 < cell.v.length ? cell.v[ 0 ] : '' )
        } );
      } );
    } );

    return cells;
  }

  /**
   * Calculate the size of the sheet.
   *
   * @param {Object}        sheet Sheet data.
   * @param {Array.<Array>} cells Cells.
   *
   * @return {Object} Size
   */
  static _calcSheetSize( sheet, cells ) {
    // Get the there if size is defined
    if( sheet.worksheet.dimension && 0 <= sheet.worksheet.dimension.length ) {
      const range = sheet.worksheet.dimension[ 0 ].$.ref.split( ':' );
      if( range.length === 2 ) {
        const min = XlsxExtractor._getPosition( range[ 0 ] );
        const max = XlsxExtractor._getPosition( range[ 1 ] );

        return {
          row: { min: min.row, max: max.row },
          col: { min: min.col, max: max.col }
        };
      }
    }

    const ascend = ( a, b ) => { return a - b; };
    const rows   = cells.map( ( cell ) => { return cell.row; } ).sort( ascend );
    const cols   = cells.map( ( cell ) => { return cell.col; } ).sort( ascend );

    return {
      row: { min: rows[ 0 ], max: rows[ rows.length - 1 ] },
      col: { min: cols[ 0 ], max: cols[ cols.length - 1 ] }
    };
  }

  /**
   * Get the coordinates of the cell.
   *
   * @param {String} text Position text. Such as "A1" and "U109".
   *
   * @return {Object} Position.
   */
  static _getPosition( text ) {
    const units = text.split( /([0-9]+)/ ); // 'A1' -> [ A, 1 ]
    if( units.length < 2 ) {
      return { row: 0, col: 0 };
    }

    return {
      row: parseInt( units[ 1 ], 10 ),
      col: XlsxExtractor._indexOfColumn( units[ 0 ] )
    };
  }

  /**
   * Gets the number of sheets.
   *
   * @param {ZipObject} zip Zip object.
   *
   * @return {Number} The number of sheets.
   */
  static _getSheetsCount( zip ) {
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
   * Gets the index of the column.
   *
   * @param {String} text Column text. Such as A" and "AA".
   *
   * @return {Number} Column number.
   */
  static _indexOfColumn( text ) {
    const letters = [ '', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' ];
    const col     = text.trim().split( '' );

    let num = 0;
    for( let i = 0, max = col.length; i < max; ++i ) {
      num *= 26;
      num += letters.indexOf( col[ i ] );
    }

    return num;
  }

  /**
   * To analyze the "r" element of XML.
   *
   * @param {Array.<Object>} r "r" elements.
   *
   * @return {String} Parse result.
   */
  static _parseR( r ) {
    let value = '';
    r.forEach( ( obj ) => {
      if( obj.t ) {
        value += XlsxExtractor._parseT( obj.t );
      }
    } );

    return value;
  }

  /**
   * To analyze the "t" element of XML.
   *
   * @param {Array.<Object>} t "t" elements.
   *
   * @return {String} Parse result.
   */
  static _parseT( t ) {
    let value = '';
    t.forEach( ( obj ) => {
      switch( typeof obj ) {
        case 'string':
          value += obj;
          break;

        //  The value of xml:space="preserve" is stored in the underscore
        case 'object':
          if( obj._ && typeof obj._ === 'string' ) {
            value += obj._;
          }
          break;

        default:
          break;
      }
    } );

    return value;
  }

  /**
   * Parses the XML.
   *
   * @param {Number} index Index of sheet. Range of from 1 to XlsxExtractor.count.
   *
   * @return {Promise} Parse task.
   */
  static _parseXML( index ) {
    const task = ( xml ) => {
      return new Promise( ( resolve, reject ) => {
        XmlParser.parseString( xml, ( err, obj ) => {
          return ( err ? reject( err ) : resolve( obj ) );
        } );
      } );
    };

    const zip   = XlsxExtractor._unzip( this._path );
    const tasks = [ task( zip.files[ FilePaths.SheetBase + index + '.xml' ] ) ];
    if( zip.files[ FilePaths.SharedStrings ] ) {
      tasks.push( task( zip.files[ FilePaths.SharedStrings ] ) );
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

  /**
   * Extract a zip file.
   *
   * @param {String} path Zip file path.
   *
   * @return {ZipObject} If success zip object, otherwise null.
   */
  static _unzip( path ) {
    try {
      const file = Fs.readFileSync( Path.resolve( path ) );
      return Zip( file );
    } catch( err ) {
      return null;
    }
  }

  /**
   * Get a value from the cell strings.
   *
   * @param {Object} str Cell strings.
   *
   * @return {String} Value.
   */
  static _valueFromStrings( str ) {
    let   value = '';
    const keys  = Object.keys( str );

    keys.forEach( ( key ) => {
      switch( key ) {
        case 't':
          value += XlsxExtractor._parseT( str[ key ] );
          break;

        case 'r':
          value += XlsxExtractor._parseR( str[ key ] );
          break;

        default:
          break;
      }
    } );

    return value;
  }
}
