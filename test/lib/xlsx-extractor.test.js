const assert = require( 'power-assert' );
const XlsxExtractor = require( '../../src/lib/xlsx-extractor.js' );

/** @test {XlsxExtractor} */
describe( 'XlsxExtractor', () => {
  /** @test {XlsxExtractor#constructor} */
  describe( 'constructor', () => {
    it( 'Invalid XLSX', () => {
      assert.throws( () => {
        const extractor = new XlsxExtractor();
        assert( !( extractor ) );
      } );
    } );
  } );

  /** @test {XlsxExtractor#extract} */
  describe( 'extract', () => {
    const sampleXML = './test/data/sample.xlsx';

    it( 'Count', () => {
      const extractor = new XlsxExtractor( sampleXML );
      assert( extractor.count === 2 );
    } );

    it( 'Index out of range: lower', () => {
      const extractor = new XlsxExtractor( sampleXML );
      return extractor
      .extract( 0 )
      .then( null, ( err ) => {
        assert( err );
      } );
    } );

    it( 'Index out of range: upper', () => {
      const extractor = new XlsxExtractor( sampleXML );
      return extractor
      .extract( 5 )
      .then( null, ( err ) => {
        assert( err );
      } );
    } );

    it( 'Sheet: 1', () => {
      const extractor = new XlsxExtractor( sampleXML );
      return extractor
      .extract( 1 )
      .then( ( result ) => {
        assert( result.id === 1 );
        assert( result.name === 'Sample Sheet' );
        assert( result.cells.length === 10 );
        assert( result.cells[ 0 ].length === 17 );
      } );
    } );

    it( 'Sheet: 2', () => {
      const extractor = new XlsxExtractor( sampleXML );
      return extractor
      .extract( 2 )
      .then( ( result ) => {
        assert( result.id === 2 );
        assert( result.name === 'Example Sheet' );
        assert( result.cells.length === 7 );
        assert( result.cells[ 0 ].length === 8 );
      } );
    } );
  } );
} );
