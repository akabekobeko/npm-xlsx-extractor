#!/usr/bin/env node

import Commander from 'commander'
import { extract, extractAll, extractRange, getSheetCount } from '../lib/index'

/** Options of command line interface. */
type CLIOptions = {
  /** Path of the XLSX file. */
  input: string
  /**
   * Range of sheets to be output.
   * If `begin` and `end` are `0`, all sheets will be output, and the same number greater than `1` will be output as a single sheet.
   */
  range: {
    begin: number
    end: number
  }
  /** `true` to output only the number of sheets */
  count?: boolean
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
 * @return Parse results.
 */
const parseArgv = (): CLIOptions => {
  Commander.usage('xlsx-extractor [options]')
    .description('Extract the colums/rows from XLSX file.')
    .option('-i, --input <File>', 'Path of the XLSX file')
    .option(
      '-r, --range <Range>',
      'Range of sheets to be output. Specify the numeric value of "N" or "N-N".',
      parseRange
    )
    .option(
      '-c, --count',
      'Outputs the number of sheet. This option overrides --range.'
    )
    .version(require('../../package.json').version, '-v, --version')

  Commander.on('--help', () => {
    console.log(`
Examples:
  $ xlsx-extractor -i sample.xlsx
  $ xlsx-extractor -i sample.xlsx -c
  $ xlsx-extractor -i sample.xlsx -r 3
  $ xlsx-extractor -i sample.xlsx -r 1-5

See also:
  https://github.com/akabekobeko/npm-xlsx-extractor/issues`)
  })

  // Print help and exit if there are no arguments
  if (process.argv.length < 3) {
    Commander.help()
  }

  Commander.parse(process.argv)
  const opts = Commander.opts()
  return {
    input: opts.input || '',
    range: opts.range || { begin: 0, end: 0 }
  }
}

/**
 * Entry point of command line interface.
 */
const main = () => {
  const options = parseArgv()
  if (options.count) {
    console.log(getSheetCount(options.input))
  } else if (options.range.begin === 0 && options.range.end === 0) {
    extractAll(options.input)
      .then((sheets) => {
        console.log(JSON.stringify(sheets, null, '  '))
      })
      .catch((err) => {
        console.log(err)
      })
  } else if (options.range.begin !== options.range.end) {
    extractRange(options.input, options.range.begin, options.range.end)
      .then((sheets) => {
        console.log(JSON.stringify(sheets, null, '  '))
      })
      .catch((err) => {
        console.log(err)
      })
  } else {
    extract(options.input, options.range.begin)
      .then((sheet) => {
        console.log(JSON.stringify(sheet, null, '  '))
      })
      .catch((err) => {
        console.log(err)
      })
  }
}

main()
