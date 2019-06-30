import Fs from 'fs'
import Path from 'path'
import XmlParser from 'xml2js'
const Zip = require('node-zip')

/**
 * This represents an entry in the zip file. If the entry comes from an existing archive previously loaded, the content will be automatically decompressed/converted first.
 * @see https://stuk.github.io/jszip/documentation/api_zipobject.html
 */
export type ZipObject = {
  /** the absolute path of the file. */
  name: string
  /** `true` if this is a directory. */
  dir: boolean
  /** the last modification date. */
  date: Date
  /** the comment for this file. */
  comment: string
  /** The UNIX permissions of the file, if any. 16 bits number. */
  unixPermissions: number
  /** 	The DOS permissions of the file, if any. 6 bits number. */
  dosPermissions: number
  /** the options of the file. The available options. */
  options: {
    compression: (
      name: string,
      data:
        | string
        | ArrayBuffer
        | Uint8Array
        | Buffer
        | Blob
        | Promise<any>
        | WritableStream
    ) => void
  }
  /** Files. */
  files: any
}

/** It is a cell in a sheet. */
export type Cell = {
  /** Row position. */
  row: number
  /** Column position. */
  col: number
  /** Type.. */
  type: string
  /** Value string. */
  value: string
}

/**
 * Create a empty cells.
 * @param rows Rows count.
 * @param cols Columns count.
 * @return Cells.
 */
export const createEmptyCells = (rows: number, cols: number) => {
  const arr = []
  for (let i = 0; i < rows; ++i) {
    const row = []
    for (let j = 0; j < cols; ++j) {
      row.push('')
    }

    arr.push(row)
  }

  return arr
}

/**
 * Get a cells from a rows.
 * @param {rows Rows.
 * @return Cells.
 */
export const getCells = (rows: any[]) => {
  const cells: Cell[] = []
  rows
    .filter((row) => {
      return row.c && 0 < row.c.length
    })
    .forEach((row) => {
      row.c.forEach((cell: any) => {
        const position = getPosition(cell.$.r)
        cells.push({
          row: position.row,
          col: position.col,
          type: cell.$.t ? cell.$.t : '',
          value: cell.v && 0 < cell.v.length ? cell.v[0] : ''
        })
      })
    })

  return cells
}

/**
 * Get the coordinates of the cell.
 * @param text Position text. Such as "A1" and "U109".
 * @return Position.
 */
export const getPosition = (text: string) => {
  // 'A1' -> [A, 1]
  const units = text.split(/([0-9]+)/)
  if (units.length < 2) {
    return { row: 0, col: 0 }
  }

  return {
    row: parseInt(units[1], 10),
    col: numOfColumn(units[0])
  }
}

/**
 * Get the size of the sheet.
 * @param sheet Sheet data.
 * @param cells Cells.
 * @return Size.
 */
export const getSheetSize = (sheet: any, cells: any[]) => {
  // Get the there if size is defined
  if (
    sheet &&
    sheet.worksheet &&
    sheet.worksheet.dimension &&
    0 <= sheet.worksheet.dimension.length
  ) {
    const range = sheet.worksheet.dimension[0].$.ref.split(':')
    if (range.length === 2) {
      const min = getPosition(range[0])
      const max = getPosition(range[1])

      return {
        row: { min: min.row, max: max.row },
        col: { min: min.col, max: max.col }
      }
    }
  }

  const ascend = (a: number, b: number) => a - b
  const rows = cells.map((cell) => cell.row).sort(ascend)
  const cols = cells.map((cell) => cell.col).sort(ascend)

  return {
    row: { min: rows[0], max: rows[rows.length - 1] },
    col: { min: cols[0], max: cols[cols.length - 1] }
  }
}

/**
 * Convert the column text to number.
 * @param text Column text, such as A" and "AA".
 * @return Column number, otherwise -1.
 */
export const numOfColumn = (text: string) => {
  const letters = [
    '',
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    'G',
    'H',
    'I',
    'J',
    'K',
    'L',
    'M',
    'N',
    'O',
    'P',
    'Q',
    'R',
    'S',
    'T',
    'U',
    'V',
    'W',
    'X',
    'Y',
    'Z'
  ]
  const col = text.trim().split('')

  let num = 0
  for (let i = 0, max = col.length; i < max; ++i) {
    num *= 26
    num += letters.indexOf(col[i])
  }

  return num
}

/**
 * Parse the `r` element of XML.
 * @param r `r` elements.
 * @return Parse result.
 */
export const parseR = (r: any[]) => {
  let value = ''
  r.forEach((obj) => {
    if (obj.t) {
      value += parseT(obj.t)
    }
  })

  return value
}

/**
 * Parse the `t` element of XML.
 * @param t `t` elements.
 * @return Parse result.
 */
export const parseT = (t: any[]) => {
  let value = ''
  t.forEach((obj) => {
    switch (typeof obj) {
      case 'string':
        value += obj
        break

      //  The value of xml:space="preserve" is stored in the underscore
      case 'object':
        if (obj._ && typeof obj._ === 'string') {
          value += obj._
        }
        break

      default:
        break
    }
  })

  return value
}

/**
 * Parse the XML text.
 * @param xml XML text.
 * @return XML parse task.
 */
export const parseXML = (xml: string): Promise<any> => {
  return new Promise((resolve, reject) => {
    XmlParser.parseString(xml, (err, obj) => {
      return err ? reject(err) : resolve(obj)
    })
  })
}

/**
 * Extract a zip file.
 * @param path Zip file path.
 * @return If success zip object, otherwise null.
 */
export const unzip = (path: string): ZipObject | null => {
  try {
    const file = Fs.readFileSync(Path.resolve(path))
    return Zip(file)
  } catch (err) {
    return null
  }
}

/**
 * Get a value from the cell strings.
 *
 * @param str Cell strings.
 *
 * @return Value.
 */
export const valueFromStrings = (str: any) => {
  let value = ''
  const keys = Object.keys(str)

  keys.forEach((key) => {
    switch (key) {
      case 't':
        value += parseT(str[key])
        break

      case 'r':
        value += parseR(str[key])
        break

      default:
        break
    }
  })

  return value
}
