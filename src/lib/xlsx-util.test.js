import assert from 'assert'
import XlsxUtil from './xlsx-util.js'

/** @test {XlsxUtil} */
describe('XlsxUtil', () => {
  /** @test {XlsxUtil#createEmptyCells} */
  describe('createEmptyCells', () => {
    it('Create empty cells', () => {
      const rows  = 10
      const cols  = 5
      const cells = XlsxUtil.createEmptyCells(rows, cols)
      assert(cells.length === rows)
      assert(cells[0].length === cols)
    })
  })

  /** @test {XlsxUtil#getSheetSize} */
  describe('getSheetSize', () => {
    const cells = [
      {row: 1, col: 20},
      {row: 2, col: 11},
      {row: 3, col: 47}
    ]

    it('From dimension', () => {
      const sheet  = {worksheet: {dimension: [{$: {ref: 'D1:E23'}}]}}
      const actual = XlsxUtil.getSheetSize(sheet, cells)
      assert(actual.row.min === 1)
      assert(actual.row.max === 23)
      assert(actual.col.min === 4)
      assert(actual.col.max === 5)
    })

    it('Calculate from cells', () => {
      const actual = XlsxUtil.getSheetSize(null, cells)
      assert(actual.row.min === 1)
      assert(actual.row.max === 3)
      assert(actual.col.min === 11)
      assert(actual.col.max === 47)
    })
  })

  /** @test {XlsxUtil#numOfColumn} */
  describe('numOfColumn', () => {
    it('X', () => {
      const actual = XlsxUtil.numOfColumn('X')
      assert(actual === 24)
    })

    it('AB', () => {
      const actual = XlsxUtil.numOfColumn('AB')
      assert(actual === 28)
    })

    it('ZZ', () => {
      const actual = XlsxUtil.numOfColumn('ZZ')
      assert(actual === 702)
    })

    it('Non numeric', () => {
      const actual = XlsxUtil.numOfColumn('7')
      assert(actual === -1)
    })
  })
})
