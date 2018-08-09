import Path from 'path'
import XlsxExtractor from '../lib/index.js'

/**
 * Help text.
 * @type {String}
 */
const HelpText =
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
const Options = {
  help: { name: '--help', shortName: '-h' },
  version: { name: '--version', shortName: '-v' },
  input: { name: '--input', shortName: '-i' },
  range: { name: '--range', shortName: '-r' },
  count: { name: '--count', shortName: '-c' }
}

/**
 * Check that it is an option value.
 *
 * @param {String} value Value.
 *
 * @return {Boolean} If the option of the value "true".
 */
const isValue = (value) => {
  const keys = Object.keys(Options)
  return !(keys.some((key) => value === Options[key].name || value === Options[key].shortName))
}

/**
 * Parse for option value.
 *
 * @param {String[]} argv Arguments of the command line.
 * @param {Number} index Index of argumens.
 *
 * @return {String} Its contents if the option value, otherwise null.
 */
const parseArgValue =  (argv, index) => {
  if (!(index + 1 < argv.length)) {
    return null
  }

  const value = argv[index + 1]
  return (isValue(value) ? value : null)
}

/**
 * Parse for the output option.
 *
 * @param {String} arg Option.
 *
 * @return {Range} Range.
 */
const parseRange = (arg) => {
  const result = { begin: 0, end: 0 }
  if (typeof arg !== 'string') {
    return result
  }

  const range  = arg.split('-')
  if (1 < range.length) {
    result.begin = Number(range[0])
    result.end   = Number(range[1])
  } else {
    // Single mode
    result.begin = Number(range[0])
    result.end   = Number(range[0])
  }

  return result
}

/**
 * Parse for the command line argumens.
 *
 * @param {String[]} argv Arguments of the command line.
 *
 * @return {CLIOptions} Parse results.
 */
const parseArgv = (argv = []) => {
  const options = {}
  let   value   = null

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
        options.range = parseRange(value)
        break

      case Options.count.name:
      case Options.count.shortName:
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
 * Create a extract tasks.
 *
 * @param {XlsxExtractor} extractor Extractor.
 * @param {Object} range Output range.
 *
 * @return {Promise[]} Tasks.
 */
const createExtractTasks = (extractor, range) => {
  let begin = 1
  let end   = extractor.count
  if (!(range.begin === 0 && range.end === 0)) {
    begin = range.begin
    end   = range.end
  }

  const tasks = []
  for (let i = begin; i <= end; ++i) {
    tasks.push(extractor.extract(i))
  }

  return tasks
}

/**
 * Print a help text.
 *
 * @param {WritableStream} stdout Standard output.
 */
const printHelp = (stdout) => {
  stdout.write(HelpText)
}

/**
 * Print a version number.
 *
 * @param {WritableStream} stdout Standard output.
 */
const printVersion = (stdout) => {
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
 * Entry point of the CLI.
 *
 * @param {String[]} argv Arguments of the command line.
 * @param {WritableStream} stdout Standard output.
 *
 * @return {Promise} Asynchronous task.
 */
const CLI = (argv, stdout) => {
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

    if (!(options.input)) {
      return reject(new Error('"-i" or "--input" has not been specified. This parameter is required.'))
    }

    const extractor = new XlsxExtractor(options.input)
    if (options.count) {
      stdout.write(extractor.count + '\n')
      return resolve()
    }

    const tasks = createExtractTasks(extractor, options.range)
    return Promise
      .all(tasks)
      .then((results) => {
        const sheets = results.sort((a, b) => a.id - b.id)
        stdout.write(JSON.stringify(sheets, null, '  ') + '\n')
      })
  })
}

export default CLI
