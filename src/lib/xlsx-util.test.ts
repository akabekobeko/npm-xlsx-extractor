import assert from 'assert'
import { createEmptyCells, getSheetSize, numOfColumn } from './xlsx-util'

describe('XlsxUtil', () => {
  describe('createEmptyCells', () => {
    it('Create empty cells', () => {
      const rows = 10
      const cols = 5
      const cells = createEmptyCells(rows, cols)
      assert.strictEqual(cells.length, rows)
      assert.strictEqual(cells[0].length, cols)
    })
  })

  describe('getSheetSize', () => {
    const cells = [
      { row: 1, col: 20 },
      { row: 2, col: 11 },
      { row: 3, col: 47 }
    ]

    it('From dimension', () => {
      const sheet = { worksheet: { dimension: [{ $: { ref: 'D1:E23' } }] } }
      const size = getSheetSize(sheet, cells)
      assert.strictEqual(size.row.min, 1)
      assert.strictEqual(size.row.max, 23)
      assert.strictEqual(size.col.min, 4)
      assert.strictEqual(size.col.max, 5)
    })

    it('Calculate from cells', () => {
      const size = getSheetSize(null, cells)
      assert.strictEqual(size.row.min, 1)
      assert.strictEqual(size.row.max, 3)
      assert.strictEqual(size.col.min, 11)
      assert.strictEqual(size.col.max, 47)
    })
  })

  describe('numOfColumn', () => {
    it('X', () => {
      const count = numOfColumn('X')
      assert.strictEqual(count, 24)
    })

    it('AB', () => {
      const count = numOfColumn('AB')
      assert.strictEqual(count, 28)
    })

    it('ZZ', () => {
      const count = numOfColumn('ZZ')
      assert.strictEqual(count, 702)
    })

    it('Non numeric', () => {
      const count = numOfColumn('7')
      assert.strictEqual(count, -1)
    })
  })
})
