import Path from 'path'

/**
 * Help text.
 * @type {String}
 */
export const HelpText =
`
Usage: xlsx-extractor [OPTIONS]

  Extract the colums/rows from XLSX file.

  Options:
    -h, --help    Display this text.

    -v, --version Display the version number.

    -i, --input   Path of the XLSX file.

    -r, --range   Range of sheets to be output.
                  Specify the numeric value of "N" or "N-N".
                  When omitted will output all of the sheet.

    -c, --count   Outputs the number of sheet.
                  This option overrides the -r and --range.

  Examples:
    $ xlsx-extractor -i sample.xlsx
    $ xlsx-extractor -i sample.xlsx -c
    $ xlsx-extractor -i sample.xlsx -r 3
    $ xlsx-extractor -i sample.xlsx -r 1-5

  See also:
    https://github.com/akabekobeko/npm-xlsx-extractor/issues
`

/**
 * CLI options.
 * @type {Object}
 */
export const Options = {
  help: [ '-h', '--help' ],
  version: [ '-v', '--version' ],
  input: [ '-i', '--input' ],
  range: [ '-r', '--range' ],
  count: [ '-c', '--count' ]
}

/**
 * Provides a command line interface.
 */
export default class CLI {
  /**
   * Parse for the command line argumens.
   *
   * @param {Array.<String>} argv Arguments of the command line.
   *
   * @return {CLIOptions} Parse results.
   */
  static parseArgv (argv) {
    if (!(argv && 0 < argv.length)) {
      return { help: true }
    }

    switch (argv[ 0 ]) {
      case Options.help[ 0 ]:
      case Options.help[ 1 ]:
        return { help: true }

      case Options.version[ 0 ]:
      case Options.version[ 1 ]:
        return { version: true }

      default:
        return CLI._parseArgv(argv)
    }
  }

  /**
   * Print a help text.
   *
   * @param {WritableStream} stdout Standard output.
   */
  static printHelp (stdout) {
    stdout.write(HelpText)
  }

  /**
   * Print a version number.
   *
   * @param {WritableStream} stdout Standard output.
   */
  static printVersion (stdout) {
    const read = (path) => {
      try {
        return require(path).version
      } catch (err) {
        return null
      }
    }

    const version = read('../package.json') || read('../../package.json')
    stdout.write('v' + version + '\n')
  }

  /**
   * Check that it is an option value.
   *
   * @param {String} value Value.
   *
   * @return {Boolean} If the option of the value "true".
   */
  static _isValue (value) {
    const keys = Object.keys(Options)
    return !(keys.some((key) => value === Options[ key ][ 0 ] || value === Options[ key ][ 1 ]))
  }

  /**
   * Parse for the command line argumens.
   *
   * @param {Array.<String>} argv Arguments of the command line.
   *
   * @return {CLIOptions} Parse results.
   */
  static _parseArgv (argv) {
    const options = {}
    let   value   = null

    argv.forEach((arg, index) => {
      switch (arg) {
        case Options.input[ 0 ]:
        case Options.input[ 1 ]:
          value = CLI._parseArgValue(argv, index)
          if (value) {
            options.input = Path.resolve(value)
          }
          break

        case Options.range[ 0 ]:
        case Options.range[ 1 ]:
          value = CLI._parseArgValue(argv, index)
          options.range = CLI._parseRange(value)
          break

        case Options.count[ 0 ]:
        case Options.count[ 1 ]:
          options.count = true
          break

        default:
          break
      }
    })

    if (options.count) {
      if (options.range) {
        options.range = undefined
      }
    } else if (!(options.range)) {
      options.range = { begin: 0, end: 0 }
    }

    return options
  }

  /**
   * Parse for option value.
   *
   * @param {Array.<String>} argv  Arguments of the command line.
   * @param {Number}         index Index of argumens.
   *
   * @return {String} Its contents if the option value, otherwise null.
   */
  static _parseArgValue (argv, index) {
    if (!(index + 1 < argv.length)) {
      return null
    }

    const value = argv[ index + 1 ]
    return (CLI._isValue(value) ? value : null)
  }

  /**
   * Parse for the output option.
   *
   * @param {String} arg Option.
   *
   * @return {Range} Range.
   */
  static _parseRange (arg) {
    const result = { begin: 0, end: 0 }
    if (typeof arg !== 'string') {
      return result
    }

    const range  = arg.split('-')
    if (1 < range.length) {
      result.begin = Number(range[ 0 ])
      result.end   = Number(range[ 1 ])
    } else {
      // Single mode
      result.begin = Number(range[ 0 ])
      result.end   = Number(range[ 0 ])
    }

    return result
  }
}
