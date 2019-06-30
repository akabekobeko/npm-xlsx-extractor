import Path from 'path'
import { Options, parseArgv } from './cli'

describe('CLI', () => {
  describe('parseArgv', () => {
    it('Input', () => {
      const input = './test/data/sample.xlsx'
      const expected = Path.resolve(input)
      let options = parseArgv([Options.input.name, input])
      expect(options.input).toBe(expected)

      options = parseArgv([Options.input.shortName, input])
      expect(options.input).toBe(expected)

      options = parseArgv([Options.input.name])
      expect(options.input).not.toBe(expected)

      options = parseArgv([Options.input.shortName, Options.help.name])
      expect(options.input).not.toBe(expected)
    })

    it('Range', () => {
      let options = parseArgv([Options.range.name])
      expect(options.range.begin).toBe(0)
      expect(options.range.end).toBe(0)

      options = parseArgv([Options.range.shortName, '5'])
      expect(options.range.begin).toBe(5)
      expect(options.range.end).toBe(5)

      options = parseArgv([Options.range.shortName, '1-4'])
      expect(options.range.begin).toBe(1)
      expect(options.range.end).toBe(4)
    })

    it('Count', () => {
      let options = parseArgv([Options.count.name])
      expect(options.showCount).toBeTruthy()

      options = parseArgv([Options.count.shortName])
      expect(options.showCount).toBeTruthy()
    })
  })
})
