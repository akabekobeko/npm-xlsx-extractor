import assert from 'assert'
import XlsxExtractor from './xlsx-extractor.js'

/** @test {XlsxExtractor} */
describe('XlsxExtractor', () => {
  /** @test {XlsxExtractor#constructor} */
  describe('constructor', () => {
    it('Invalid XLSX', () => {
      assert.throws(() => {
        const extractor = new XlsxExtractor()
        assert(!(extractor))
      })
    })
  })

  /** @test {XlsxExtractor#extract} */
  describe('extract', () => {
    const sampleXML = './examples/sample.xlsx'

    it('Count', () => {
      const extractor = new XlsxExtractor(sampleXML)
      assert(extractor.count === 2)
    })

    it('Out of range: lower', () => {
      const extractor = new XlsxExtractor(sampleXML)
      return extractor
        .extract(0)
        .then(null, (err) => {
          assert(err)
        })
    })

    it('Out of range: upper', () => {
      const extractor = new XlsxExtractor(sampleXML)
      return extractor
        .extract(5)
        .then(null, (err) => {
          assert(err)
        })
    })

    it('Sheet: 1', () => {
      const extractor = new XlsxExtractor(sampleXML)
      return extractor
        .extract(1)
        .then((result) => {
          assert(result.id === 1)
          assert(result.name === 'Sample Sheet')
          assert(result.cells.length === 10)
          assert(result.cells[0].length === 17)
        })
    })

    it('Sheet: 2', () => {
      const extractor = new XlsxExtractor(sampleXML)
      return extractor
        .extract(2)
        .then((result) => {
          assert(result.id === 2)
          assert(result.name === 'Example Sheet')
          assert(result.cells.length === 7)
          assert(result.cells[0].length === 8)
        })
    })
  })

  /** @test {XlsxExtractor#extractAll} */
  describe('extractAll', () => {
    it('Extract all', () => {
      const extractor = new XlsxExtractor('./examples/sample.xlsx')
      return extractor
        .extractAll()
        .then((results) => {
          assert(results.length === 2)

          let sheet = results[0]
          assert(sheet.id === 1)
          assert(sheet.name === 'Sample Sheet')
          assert(sheet.cells.length === 10)
          assert(sheet.cells[0].length === 17)

          sheet = results[1]
          assert(sheet.id === 2)
          assert(sheet.name === 'Example Sheet')
          assert(sheet.cells.length === 7)
          assert(sheet.cells[0].length === 8)
        })
    })
  })
})
