import assert from 'assert'
import { getSheetCount, extract, extractAll, extractRange } from './index'

describe('XlsxExtractor', () => {
  const sampleXML = './examples/sample.xlsx'

  describe('extract', () => {
    it('Count', () => {
      assert.strictEqual(getSheetCount(sampleXML), 2)
    })

    it('Out of range: lower', () => {
      return extract(sampleXML, 0).then(null, (err) => {
        assert.notStrictEqual(err, null)
      })
    })

    it('Out of range: upper', () => {
      return extract(sampleXML, 5).then(null, (err) => {
        assert.notStrictEqual(err, null)
      })
    })

    it('Sheet: 1', () => {
      return extract(sampleXML, 1).then((sheet) => {
        assert.strictEqual(sheet.id, 1)
        assert.strictEqual(sheet.name, 'Sample Sheet')
        assert.strictEqual(sheet.cells.length, 10)
        assert.strictEqual(sheet.cells[0].length, 17)
      })
    })

    it('Sheet: 2', () => {
      return extract(sampleXML, 2).then((sheet) => {
        assert.strictEqual(sheet.id, 2)
        assert.strictEqual(sheet.name, 'Example Sheet')
        assert.strictEqual(sheet.cells.length, 7)
        assert.strictEqual(sheet.cells[0].length, 8)
      })
    })
  })

  describe('extractRange', () => {
    it('Range', () => {
      return extractRange(sampleXML, 1, 2).then((sheets) => {
        assert.strictEqual(sheets.length, 2)

        let sheet = sheets[0]
        assert.strictEqual(sheet.id, 1)
        assert.strictEqual(sheet.name, 'Sample Sheet')
        assert.strictEqual(sheet.cells.length, 10)
        assert.strictEqual(sheet.cells[0].length, 17)

        sheet = sheets[1]
        assert.strictEqual(sheet.id, 2)
        assert.strictEqual(sheet.name, 'Example Sheet')
        assert.strictEqual(sheet.cells.length, 7)
        assert.strictEqual(sheet.cells[0].length, 8)
      })
    })
  })

  describe('extractAll', () => {
    it('All', () => {
      return extractAll(sampleXML).then((sheets) => {
        assert.strictEqual(sheets.length, 2)

        let sheet = sheets[0]
        assert.strictEqual(sheet.id, 1)
        assert.strictEqual(sheet.name, 'Sample Sheet')
        assert.strictEqual(sheet.cells.length, 10)
        assert.strictEqual(sheet.cells[0].length, 17)

        sheet = sheets[1]
        assert.strictEqual(sheet.id, 2)
        assert.strictEqual(sheet.name, 'Example Sheet')
        assert.strictEqual(sheet.cells.length, 7)
        assert.strictEqual(sheet.cells[0].length, 8)
      })
    })
  })
})
