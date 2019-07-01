import {
  unzip,
  getCells,
  getSheetSize,
  createEmptyCells,
  valueFromStrings,
  getSheetInnerCount,
  getSheetData
} from './xlsx-util'
import { ZipObject } from './xlsx-util'

/** Sheet values. */
export type Sheet = {
  /** Number of the extract sheets. */
  id: number
  /** Name of the sheet. */
  name: string
  /** Cells of the sheet. Empty cell is stored is `""`. */
  cells: string[][]
}

/**
 * Get a sheet.
 * @param zip Extract data of XLSX (Zip) file.
 * @param index Index of sheet. Range of from 1 to XlsxExtractor.count.
 * @returns Sheet.
 */
const getSheet = async (zip: ZipObject, index: number): Promise<Sheet> => {
  const data = await getSheetData(zip, index)
  const cells = getCells(data.sheet.worksheet.sheetData[0].row)
  const size = getSheetSize(data.sheet, cells)
  const rows = size.row.max - size.row.min + 1
  const cols = size.col.max - size.col.min + 1
  const sheet = createEmptyCells(rows, cols)
  const strings = !data.strings ? {} : data.strings.sst.si

  cells.forEach((cell) => {
    let value = cell.value
    if (cell.type === 's') {
      const index = parseInt(value, 10)
      value = valueFromStrings(strings[index])
    }

    let row = cell.row - size.row.min
    row = row >= 0 ? row : size.row.min

    let col = cell.col - size.col.min
    col = col >= 0 ? col : size.col.min

    sheet[row][col] = value
  })

  return {
    id: index,
    name: data.name,
    cells: sheet
  }
}

/**
 * Extract a specified range of sheets.
 * Using `ZipObject` efficiently handles counting and getting multiple sheets without opening the file every time.
 * @param filePath Path of XMLS file.
 * @param range Range of sheets. The default is all.
 * @returns Sheets.
 */
const extractRangeInner = async (
  zip: ZipObject,
  begin: number,
  end: number
): Promise<Sheet[]> => {
  const sheets: Sheet[] = []
  const count = getSheetInnerCount(zip)
  if (begin < 1 || count < end) {
    throw new Error(
      `Index out of range. begin = ${begin} (min = 1), end = ${end} (max = ${count})`
    )
  }

  for (let i = begin; i <= end; ++i) {
    sheets.push(await getSheet(zip, i))
  }

  return sheets
}

/**
 * Extract and get an index of sheets.
 * @param filePath Path of XMLS file.
 * @returns Index of sheets (1 - Sheet count).
 * @throws Failed to expand the XLSX file.
 */
export const getSheetCount = (filePath: string): number => {
  // Hide structure in ZIP file to "Inner" side
  return getSheetInnerCount(unzip(filePath))
}

/**
 * Extract a sheet.
 * @param filePath Path of XMLS file.
 * @param index Index of sheet. Range of from 1 to XlsxExtractor.count.
 * @return Sheet.
 */
export const extract = (filePath: string, index: number): Promise<Sheet> => {
  return getSheet(unzip(filePath), index)
}

/**
 * Extract and get a specified range of sheets.
 * @param filePath Path of XMLS file.
 * @param begin Begin index (1 - Sheet count).
 * @param end End index (1 - Sheet count).
 * @returns Sheets.
 */
export const extractRange = async (
  filePath: string,
  begin: number,
  end: number
): Promise<Sheet[]> => {
  return extractRangeInner(unzip(filePath), begin, end)
}

/**
 * Extract and get specified all of sheets.
 * @param filePath Path of XMLS file.
 * @returns Sheets.
 */
export const extractAll = (filePath: string): Promise<Sheet[]> => {
  const zip = unzip(filePath)
  return extractRangeInner(zip, 1, getSheetInnerCount(zip))
}
