import { getSheetCount, extract, extractRange } from './index'

describe('XlsxExtractor', () => {
  const sampleXML = './examples/sample.xlsx'

  describe('extract', () => {
    it('Count', () => {
      expect(getSheetCount(sampleXML)).toBe(2)
    })

    it('Out of range: lower', () => {
      return extract(sampleXML, 0).then(null, (err) => {
        expect(err).not.toBeNull()
      })
    })

    it('Out of range: upper', () => {
      return extract(sampleXML, 5).then(null, (err) => {
        expect(err).not.toBeNull()
      })
    })

    it('Sheet: 1', () => {
      return extract(sampleXML, 1).then((result) => {
        expect(result.id).toBe(1)
        expect(result.name).toBe('Sample Sheet')
        expect(result.cells.length).toBe(10)
        expect(result.cells[0].length).toBe(17)
      })
    })

    it('Sheet: 2', () => {
      return extract(sampleXML, 2).then((result) => {
        expect(result.id).toBe(2)
        expect(result.name).toBe('Example Sheet')
        expect(result.cells.length).toBe(7)
        expect(result.cells[0].length).toBe(8)
      })
    })
  })

  describe('extractRange', () => {
    it('Range', () => {
      return extractRange(sampleXML, { begin: 1, end: 2 }).then((results) => {
        expect(results.length).toBe(2)

        let sheet = results[0]
        expect(sheet.id).toBe(1)
        expect(sheet.name).toBe('Sample Sheet')
        expect(sheet.cells.length).toBe(10)
        expect(sheet.cells[0].length).toBe(17)

        sheet = results[1]
        expect(sheet.id).toBe(2)
        expect(sheet.name).toBe('Example Sheet')
        expect(sheet.cells.length).toBe(7)
        expect(sheet.cells[0].length).toBe(8)
      })
    })

    it('All', () => {
      return extractRange(sampleXML).then((results) => {
        expect(results.length).toBe(2)

        let sheet = results[0]
        expect(sheet.id).toBe(1)
        expect(sheet.name).toBe('Sample Sheet')
        expect(sheet.cells.length).toBe(10)
        expect(sheet.cells[0].length).toBe(17)

        sheet = results[1]
        expect(sheet.id).toBe(2)
        expect(sheet.name).toBe('Example Sheet')
        expect(sheet.cells.length).toBe(7)
        expect(sheet.cells[0].length).toBe(8)
      })
    })
  })
})
