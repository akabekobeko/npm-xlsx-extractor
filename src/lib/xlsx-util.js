import Fs from 'fs'
import Path from 'path'
import XmlParser from 'xml2js'
import Zip from 'node-zip'

/**
 * Provides utility methods for XLSX.
 */
export default class XlsxUtil {
  /**
   * Create a empty cells.
   *
   * @param {Number} rows Rows count.
   * @param {Number} cols Columns count.
   *
   * @return {Array.<String[]>} Cells.
   */
  static createEmptyCells (rows, cols) {
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
   *
   * @param {Object[]} rows Rows.
   *
   * @return {Object[]} Cells.
   */
  static getCells (rows) {
    const cells = []
    rows
      .filter((row) => {
        return (row.c && 0 < row.c.length)
      })
      .forEach((row) => {
        row.c.forEach((cell) => {
          const position = XlsxUtil.getPosition(cell.$.r)
          cells.push({
            row: position.row,
            col: position.col,
            type: (cell.$.t ? cell.$.t : ''),
            value: (cell.v && 0 < cell.v.length ? cell.v[0] : '')
          })
        })
      })

    return cells
  }

  /**
   * Get the coordinates of the cell.
   *
   * @param {String} text Position text. Such as "A1" and "U109".
   *
   * @return {Object} Position.
   */
  static getPosition (text) {
    // 'A1' -> [A, 1]
    const units = text.split(/([0-9]+)/)
    if (units.length < 2) {
      return { row: 0, col: 0 }
    }

    return {
      row: parseInt(units[1], 10),
      col: XlsxUtil.numOfColumn(units[0])
    }
  }

  /**
   * Get the size of the sheet.
   *
   * @param {Object} sheet Sheet data.
   * @param {Object[]} cells Cells.
   *
   * @return {Object} Size
   */
  static getSheetSize (sheet, cells) {
    // Get the there if size is defined
    if (sheet && sheet.worksheet && sheet.worksheet.dimension && 0 <= sheet.worksheet.dimension.length) {
      const range = sheet.worksheet.dimension[0].$.ref.split(':')
      if (range.length === 2) {
        const min = XlsxUtil.getPosition(range[0])
        const max = XlsxUtil.getPosition(range[1])

        return {
          row: { min: min.row, max: max.row },
          col: { min: min.col, max: max.col }
        }
      }
    }

    const ascend = (a, b) => { return a - b }
    const rows = cells.map((cell) => { return cell.row }).sort(ascend)
    const cols = cells.map((cell) => { return cell.col }).sort(ascend)

    return {
      row: { min: rows[0], max: rows[rows.length - 1] },
      col: { min: cols[0], max: cols[cols.length - 1] }
    }
  }

  /**
   * Convert the column text to number.
   *
   * @param {String} text Column text, such as A" and "AA".
   *
   * @return {Number} Column number, otherwise -1.
   */
  static numOfColumn (text) {
    const letters = ['', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    const col     = text.trim().split('')

    let num = 0
    for (let i = 0, max = col.length; i < max; ++i) {
      num *= 26
      num += letters.indexOf(col[i])
    }

    return num
  }

  /**
   * Parse the "r" element of XML.
   *
   * @param {Array.<Object>} r "r" elements.
   *
   * @return {String} Parse result.
   */
  static parseR (r) {
    let value = ''
    r.forEach((obj) => {
      if (obj.t) {
        value += XlsxUtil.parseT(obj.t)
      }
    })

    return value
  }

  /**
   * Parse the "t" element of XML.
   *
   * @param {Object[]} t "t" elements.
   *
   * @return {String} Parse result.
   */
  static parseT (t) {
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
   *
   * @param {String} xml XML text.
   *
   * @return {Promise} XML parse task.
   */
  static parseXML (xml) {
    return new Promise((resolve, reject) => {
      XmlParser.parseString(xml, (err, obj) => {
        return (err ? reject(err) : resolve(obj))
      })
    })
  }

  /**
   * Extract a zip file.
   *
   * @param {String} path Zip file path.
   *
   * @return {ZipObject} If success zip object, otherwise null.
   */
  static unzip (path) {
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
   * @param {Object} str Cell strings.
   *
   * @return {String} Value.
   */
  static valueFromStrings (str) {
    let value = ''
    const keys  = Object.keys(str)

    keys.forEach((key) => {
      switch (key) {
        case 't':
          value += XlsxUtil.parseT(str[key])
          break

        case 'r':
          value += XlsxUtil.parseR(str[key])
          break

        default:
          break
      }
    })

    return value
  }
}
