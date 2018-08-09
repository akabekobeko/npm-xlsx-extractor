import assert from 'assert'
import Path from 'path'
import Rewire from 'rewire'

/** @test {CLI} */
describe('CLI', () => {
  const Module = Rewire('./cli.js')

  /** @test {CLI#parseArgv} */
  describe('parseArgv', () => {
    const parseArgv = Module.__get__('parseArgv')
    const Options = Module.__get__('Options')

    it('Input', () => {
      const input = './test/data/sample.xlsx'
      const expected = Path.resolve(input)
      let options = parseArgv([Options.input.name, input])
      assert(options.input === expected)

      options = parseArgv([Options.input.shortName, input])
      assert(options.input === expected)

      options = parseArgv([Options.input.name])
      assert(options.input !== expected)

      options = parseArgv([Options.input.shortName, Options.help.name])
      assert(options.input !== expected)
    })

    it('Range', () => {
      let options = parseArgv()
      options = parseArgv([Options.range.name])
      assert(options.range.begin === 0)
      assert(options.range.end   === 0)

      options = parseArgv([Options.range.shortName, '5'])
      assert(options.range.begin === 5)
      assert(options.range.end   === 5)

      options = parseArgv([Options.range.shortName, '1-4'])
      assert(options.range.begin === 1)
      assert(options.range.end   === 4)
    })

    it('Count', () => {
      let options = parseArgv()
      assert(!(options.count))

      options = parseArgv([Options.count.name])
      assert(options.count)

      options = parseArgv([Options.count.shortName])
      assert(options.count)
    })
  })
})
