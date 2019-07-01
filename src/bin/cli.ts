import Path from 'path'
import { getSheetCount, extractAll, extractRange } from '../lib/index'

/** Commad line options. */
type CLIOption = {
  /** Mode to display the help text. */
  help?: boolean
  /** Mode to display the version number. */
  version?: boolean
  /** Path of the XLSX file. */
  input?: string
  /** Range of an output sheets. */
  range: {
    begin: number
    end: number
  }
  /** Outputs the number of sheet. */
  showCount: boolean
}

/**
 * Help text.
 * @type {String}
 */
const HelpText = `
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
  help: { name: '--help', shortName: '-h' },
  version: { name: '--version', shortName: '-v' },
  input: { name: '--input', shortName: '-i' },
  range: { name: '--range', shortName: '-r' },
  count: { name: '--count', shortName: '-c' }
}

/**
 * Check that it is an option value.
 * @param value Value.
 * @return If the option of the value `true`.
 */
const isValue = (value: string) => {
  switch (value) {
    case Options.help.name:
    case Options.help.shortName:
    case Options.version.name:
    case Options.version.shortName:
    case Options.input.name:
    case Options.input.shortName:
    case Options.range.name:
    case Options.range.shortName:
    case Options.count.name:
    case Options.count.shortName:
      return false
    default:
      return true
  }
}

/**
 * Parse for option value.
 * @param argv Arguments of the command line.
 * @param index Index of argumens.
 * @return Its contents if the option value, otherwise null.
 */
const parseArgValue = (argv: string[], index: number) => {
  if (!(index + 1 < argv.length)) {
    return null
  }

  const value = argv[index + 1]
  return isValue(value) ? value : null
}

/**
 * Parse for the output option.
 * @param arg Argument of command line options.
 * @return Range.
 */
const parseRange = (arg: string) => {
  const result = { begin: 0, end: 0 }
  const range = arg.split('-')

  if (1 < range.length) {
    result.begin = Number(range[0])
    result.end = Number(range[1])
  } else {
    // Single mode
    result.begin = Number(range[0])
    result.end = Number(range[0])
  }

  return result
}

/**
 * Parse for the command line argumens.
 * @param argv Arguments of the command line.
 * @return Parse results.
 */
export const parseArgv = (argv: string[]) => {
  const options: CLIOption = { showCount: false, range: { begin: 0, end: 0 } }
  let value = null

  argv.forEach((arg, index) => {
    switch (arg) {
      case Options.input.name:
      case Options.input.shortName:
        value = parseArgValue(argv, index)
        if (value) {
          options.input = Path.resolve(value)
        }
        break

      case Options.range.name:
      case Options.range.shortName:
        value = parseArgValue(argv, index)
        if (value) {
          options.range = parseRange(value)
        }
        break

      case Options.count.name:
      case Options.count.shortName:
        options.showCount = true
        break

      default:
        break
    }
  })

  return options
}

/**
 * Execute an extract.
 * @param filePath Path of XML files.
 * @param range Output range.
 * @return Sheets.
 */
const extractXLSX = (filePath: string, begin: number, end: number) => {
  if (begin === 0 && end === 0) {
    return extractAll(filePath)
  }

  return extractRange(filePath, begin, end)
}

/**
 * Print a help text.
 * @param stdout Standard output.
 */
const printHelp = (stdout: NodeJS.WritableStream) => {
  stdout.write(HelpText)
}

/**
 * Print a version number.
 * @param stdout Standard output.
 */
const printVersion = (stdout: NodeJS.WritableStream) => {
  const read = (path: string) => {
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
 * Entry point of the CLI.
 * @param argv Arguments of the command line.
 * @param stdout Standard output.
 * @return Asynchronous task.
 */
const CLI = (argv: string[], stdout: NodeJS.WritableStream) => {
  return new Promise((resolve, reject) => {
    const options = parseArgv(argv)
    if (options.help) {
      printHelp(stdout)
      return resolve()
    }

    if (options.version) {
      printVersion(stdout)
      return resolve()
    }

    if (!options.input) {
      return reject(
        new Error(
          '"-i" or "--input" has not been specified. This parameter is required.'
        )
      )
    }

    if (options.showCount) {
      const count = getSheetCount(options.input)
      stdout.write(count + '\n')
      return resolve()
    }

    return extractXLSX(
      options.input,
      options.range.begin,
      options.range.end
    ).then((results) => {
      const sheets = results.sort((a, b) => a.id - b.id)
      stdout.write(JSON.stringify(sheets, null, '  ') + '\n')
    })
  })
}

export default CLI
