import {
  unzip,
  getCells,
  getSheetSize,
  createEmptyCells,
  valueFromStrings,
  parseXML
} from './xlsx-util'
import { ZipObject } from './xlsx-util'

/**
 * Defines the file path in the XLSX.
 * @type {Object}
 */
const FilePaths = {
  WorkBook: 'xl/workbook.xml',
  SharedStrings: 'xl/sharedStrings.xml',
  SheetBase: 'xl/worksheets/sheet'
}

/**
 * The maximum number of sheets (Excel 97).
 * @type {Number}
 */
const MaxSheets = 256

type SheetData = {
  /** Sheet name. */
  name: string
  /** Data obtained by converting the XML of the sheet to the JavaScript Object. */
  sheet: any
  /** Data obtained by converting the XML of the shared strings to the JavaScript Object. */
  strings?: any
}

export type Sheet = {
  /** Number of the extract sheets. */
  id: number
  /** Name of the sheet. */
  name: string
  /** Cells of the sheet. Empty cell is stored is `""`. */
  cells: string[][]
}

/** Index range of sheets. */
export type SheetRange = {
  /** Begin index of sheets. */
  begin: number
  /** End index of sheets. */
  end: number
}

/**
 * Get a sheet name.
 * @param zip Extract data of XLSX (Zip) file.
 * @param index Index of sheet. Range of from 1 to XlsxExtractor.count.
 * @returns Sheet name.
 */
const getSheetName = async (zip: ZipObject, index: number) => {
  const root = await parseXML(zip.files[FilePaths.WorkBook].asText())
  let name = ''
  if (
    root &&
    root.workbook &&
    root.workbook.sheets &&
    0 < root.workbook.sheets.length &&
    root.workbook.sheets[0].sheet
  ) {
    root.workbook.sheets[0].sheet.some((sheet: any) => {
      const id = Number(sheet.$.sheetId)
      if (id === index) {
        name = sheet.$.name || ''
        return true
      }

      return false
    })
  }

  return name
}

/**
 * Get a sheet data.
 * @param zip Extract data of XLSX (Zip) file.
 * @param index Index of sheet. Range of from 1 to XlsxExtractor.count.
 * @returns Sheet data.
 */
const getSheetData = async (zip: ZipObject, index: number) => {
  const data: SheetData = {
    name: '',
    sheet: {}
  }

  data.name = await getSheetName(zip, index)
  data.sheet = await parseXML(
    zip.files[FilePaths.SheetBase + index + '.xml'].asText()
  )

  if (zip.files[FilePaths.SharedStrings]) {
    data.strings = await parseXML(zip.files[FilePaths.SharedStrings].asText())
  }

  return data
}

/**
 * Get a sheet.
 * @param zip Extract data of XLSX (Zip) file.
 * @param index Index of sheet. Range of from 1 to XlsxExtractor.count.
 * @returns Sheet.
 */
const getSheet = async (zip: ZipObject, index: number) => {
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
 * Gets the number of sheets.
 * @param filePath Path of XMLS file.
 * @returns Number of sheets
 * @throws Failed to expand the XLSX file.
 */
export const getSheetCount = (filePath: string) => {
  const zip = unzip(filePath)
  if (!zip) {
    throw new Error('Failed to expand the XLSX file.')
  }

  let count = 0
  for (let i = 1; i < MaxSheets; ++i) {
    const path = FilePaths.SheetBase + i + '.xml'
    if (!zip.files[path]) {
      break
    }

    ++count
  }

  return count
}

/**
 * Extract a sheet.
 * @param filePath Path of XMLS file.
 * @param index Index of sheet. Range of from 1 to XlsxExtractor.count.
 * @return Sheet.
 */
export const extract = async (
  filePath: string,
  index: number
): Promise<Sheet> => {
  const zip = unzip(filePath)
  if (!zip) {
    throw new Error('Failed to extract ZIP file')
  }

  return getSheet(zip, index)
}

/**
 * Extract a specified range of sheets.
 * @param filePath Path of XMLS file.
 * @param range Range of sheets. The default is all.
 * @returns Sheets.
 */
export const extractRange = async (filePath: string, range?: SheetRange) => {
  const zip = unzip(filePath)
  if (!zip) {
    throw new Error('Failed to extract ZIP file')
  }

  const sheets: Sheet[] = []
  const count = await getSheetCount(filePath)
  const target = range
    ? { begin: range.begin, end: Math.min(range.end, count) }
    : { begin: 1, end: count }

  for (let i = target.begin; i <= target.end; ++i) {
    sheets.push(await getSheet(zip, i))
  }

  return sheets
}
