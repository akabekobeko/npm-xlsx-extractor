import {
  getSheetCount,
  extract,
  extractAll,
  extractRange
} from 'xlsx-extractor'

// Get sheets count
console.log(getSheetCount('./sample.xlsx'))

// Single
extract('./sample.xlsx', 1)
  .then((sheet) => {
    console.log(sheet)
  })
  .catch((err) => {
    console.log(err)
  })

// Range
extractRange('./sample.xlsx', 1, 2)
  .then((sheets) => {
    console.log(sheets)
  })
  .catch((err) => {
    console.log(err)
  })

// All
extractAll('./sample.xlsx')
  .then((sheets) => {
    console.log(sheets)
  })
  .catch((err) => {
    console.log(err)
  })
