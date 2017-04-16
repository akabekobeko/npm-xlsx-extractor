#!/usr/bin/env node

import CLI from './cli.js'
import XlsxExtractor from '../lib/index.js'

/**
 * Create a extract tasks.
 *
 * @param {XlsxExtractor} extractor Extractor.
 * @param {Object}        range     Output range.
 *
 * @return {Array.<Promise>} Tasks.
 */
function createTasks (extractor, range) {
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
 * Entry point of the CLI.
 *
 * @param {Array.<String>} argv   Arguments of the command line.
 * @param {WritableStream} stdout Standard output.
 *
 * @return {Promise} Promise object.
 */
function main (argv, stdout) {
  return new Promise((resolve, reject) => {
    const options = CLI.parseArgv(argv)
    if (options.help) {
      CLI.printHelp(stdout)
      return resolve()
    }

    if (options.version) {
      CLI.printVersion(stdout)
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

    const tasks = createTasks(extractor, options.range)
    return Promise
    .all(tasks)
    .then((results) => {
      const sheets = results.sort((a, b) => a.id - b.id)
      stdout.write(JSON.stringify(sheets, null, '  ') + '\n')
    })
  })
}

main(process.argv.slice(2), process.stdout)
.then()
.catch((err) => {
  console.error(err)
})
