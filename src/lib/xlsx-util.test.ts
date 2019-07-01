import { createEmptyCells, getSheetSize, numOfColumn } from './xlsx-util'

describe('XlsxUtil', () => {
  describe('createEmptyCells', () => {
    it('Create empty cells', () => {
      const rows = 10
      const cols = 5
      const cells = createEmptyCells(rows, cols)
      expect(cells.length).toBe(rows)
      expect(cells[0].length).toBe(cols)
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
      const actual = getSheetSize(sheet, cells)
      expect(actual.row.min).toBe(1)
      expect(actual.row.max).toBe(23)
      expect(actual.col.min).toBe(4)
      expect(actual.col.max).toBe(5)
    })

    it('Calculate from cells', () => {
      const actual = getSheetSize(null, cells)
      expect(actual.row.min).toBe(1)
      expect(actual.row.max).toBe(3)
      expect(actual.col.min).toBe(11)
      expect(actual.col.max).toBe(47)
    })
  })

  describe('numOfColumn', () => {
    it('X', () => {
      const actual = numOfColumn('X')
      expect(actual).toBe(24)
    })

    it('AB', () => {
      const actual = numOfColumn('AB')
      expect(actual).toBe(28)
    })

    it('ZZ', () => {
      const actual = numOfColumn('ZZ')
      expect(actual).toBe(702)
    })

    it('Non numeric', () => {
      const actual = numOfColumn('7')
      expect(actual).toBe(-1)
    })
  })
})
