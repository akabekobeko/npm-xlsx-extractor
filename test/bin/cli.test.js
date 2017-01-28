const assert = require( 'power-assert' );
const Path = require( 'path' );
const CLI = require( '../../src/bin/cli.js' ).CLI;
const HelpText = require( '../../src/bin/cli.js' ).HelpText;
const Options = require( '../../src/bin/cli.js' ).Options;
const Package = require( '../../package.json' );
const StdOutMock = require( '../mock/stdout.mock.js' );

/** @test {CLI} */
describe( 'CLI', () => {
  /** @test {CLI#printHelp} */
  describe( 'printHelp', () => {
    it( 'Print', () => {
      const mock = new StdOutMock();
      CLI.printHelp( mock );
      assert( mock.text === HelpText );
    } );
  } );

  /** @test {CLI#printVersion} */
  describe( 'printVersion', () => {
    it( 'Print', () => {
      const mock = new StdOutMock();
      CLI.printVersion( mock );

      const expected = 'v' + Package.version + '\n';
      assert( mock.text === expected );
    } );
  } );

  /** @test {CLI#parseArgv} */
  describe( 'parseArgv', () => {
    it( 'Help', () => {
      let options = CLI.parseArgv( [] );
      assert( options.help );

      options = CLI.parseArgv( [ Options.help[ 0 ] ] );
      assert( options.help );

      options = CLI.parseArgv( [ Options.help[ 1 ] ] );
      assert( options.help );
    } );

    it( 'Version', () => {
      let options = CLI.parseArgv( [ Options.version[ 0 ] ] );
      assert( options.version );

      options = CLI.parseArgv( [ Options.version[ 1 ] ] );
      assert( options.version );
    } );

    it( 'Input', () => {
      const input    = './test/data/sample.xlsx';
      const expected = Path.resolve( input );
      let options = CLI.parseArgv( [ Options.input[ 0 ], input ] );
      assert( options.input === expected );

      options = CLI.parseArgv( [ Options.input[ 1 ], input ] );
      assert( options.input === expected );

      options = CLI.parseArgv( [ Options.input[ 0 ] ] );
      assert( options.input !== expected );

      options = CLI.parseArgv( [ Options.input[ 1 ], Options.help[ 0 ] ] );
      assert( options.input !== expected );
    } );

    it( 'Range', () => {
      let options = CLI.parseArgv();
      options = CLI.parseArgv( [ Options.range[ 0 ] ] );
      assert( options.range.begin === 0 );
      assert( options.range.end   === 0 );

      options = CLI.parseArgv( [ Options.range[ 1 ], '5' ] );
      assert( options.range.begin === 5 );
      assert( options.range.end   === 5 );

      options = CLI.parseArgv( [ Options.range[ 1 ], '1-4' ] );
      assert( options.range.begin === 1 );
      assert( options.range.end   === 4 );
    } );

    it( 'Count', () => {
      let options = CLI.parseArgv();
      assert( !( options.count ) );

      options = CLI.parseArgv( [ Options.count[ 0 ] ] );
      assert( options.count );

      options = CLI.parseArgv( [ Options.count[ 1 ] ] );
      assert( options.count );
    } );
  } );
} );
