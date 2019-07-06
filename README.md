# xlsx-extractor

[![Support Node of LTS](https://img.shields.io/badge/node-LTS-brightgreen.svg)](https://nodejs.org/)
[![npm version](https://badge.fury.io/js/xlsx-extractor.svg)](https://badge.fury.io/js/xlsx-extractor)
[![Build Status](https://travis-ci.org/akabekobeko/npm-xlsx-extractor.svg?branch=master)](https://travis-ci.org/akabekobeko/npm-xlsx-extractor)
[![code style: prettier](https://img.shields.io/badge/code_style-prettier-ff69b4.svg)](https://github.com/prettier/prettier)

Extract the colums/rows from XLSX file. The cells of the sheet parsed by this tool will be filled with the largest columns and rows.

For example, parsing a columns:`3` rows:`3` sheet:

```json
{
  "id": 2,
  "name": "Example Sheet",
  "cells": [
    ["", "a", ""],
    ["", "", "b"],
    ["c", "", ""]
  ]
}
```

Don't trim empty cells. Therefore, it is convenient for processing programmatically while maintaining cell coordinates.


## Installation

```
$ npm install xlsx-extractor
```

## Node.js API

How to use in Node.js.

### `getSheetCount(filePath)`

Extract and get the number of sheets.

- `filePath`: `string` - Path of the XLSX file.
- **Returns**: `number` - Number of sheets.

```js
const xlsx = require('xlsx-extractor');

const count = xlsx.getSheetCount('./sample.xlsx')
console.log(count);
```

### `extract(filePath, index)`

Extract and get an index of sheets.

- `filePath`: `string` - Path of the XLSX file.
- `index`: `number` - Index of sheets (1 - Sheet count).
- **Returns**: `Promise<Sheet>` - Value of sheet.

```js
const xlsx = require('xlsx-extractor');

xlsx.extract('./sample.xlsx', 1)
  .then((sheet) => {
    console.log(sheet)
  })
  .catch((err) => {
    console.log(err)
  });
```

### `extractRange(filePath, begin, end)`

Extract and get a specified range of sheets.

- `filePath`: `string` - Path of the XLSX file.
- `begin`: `number` - Begin index (1 - Sheet count).
- `end`: `number` - End index (1 - Sheet count).
- **Returns**: `Promise<Sheet[]>` - Value of sheets.

```js
const xlsx = require('xlsx-extractor');

xlsx.extractRange('./sample.xlsx', 1, 2)
  .then((sheets) => {
    console.log(sheets)
  })
  .catch((err) => {
    console.log(err)
  });
```

### `extractAll(filePath)`

Extract and get specified all of sheets.

- `filePath`: `string` - Path of the XLSX file.
- **Returns**: `Promise<Sheet[]>` - Value of sheets.

```js
const xlsx = require('xlsx-extractor');

xlsx.extractAll('./sample.xlsx')
  .then((sheets) => {
    console.log(sheets)
  })
  .catch((err) => {
    console.log(err)
  });
```

### `Sheet`

Value of sheet.

- `id`: `number` - Index of the sheets.
- `name`: `string` - Name of the sheet.
- `cells`: `string[][]` - Cells of the sheet. Empty cell is stored is `""`.

## CLI

```
Usage:  xlsx-extractor [options]

Extract the colums/rows from XLSX file.

Options:
  -i, --input [File]        Path of the XLSX file
  -r, --range [N] or [N-N]  Range of sheets to be output. Specify the numeric value of "N" or "N-N".
  -c, --count               Outputs the number of sheet. This option overrides --range.
  -v, --version             output the version number
  -h, --help                output usage information

Examples:
  $ xlsx-extractor -i sample.xlsx
  $ xlsx-extractor -i sample.xlsx -c
  $ xlsx-extractor -i sample.xlsx -r 3
  $ xlsx-extractor -i sample.xlsx -r 1-5

See also:
  https://github.com/akabekobeko/npm-xlsx-extractor/issues
```

## ChangeLog

* [CHANGELOG](CHANGELOG.md)

## License

* [MIT](LICENSE.txt)
