# xlsx-extractor

[![Support Node of LTS](https://img.shields.io/badge/node-LTS-brightgreen.svg)](https://nodejs.org/)
[![npm version](https://badge.fury.io/js/xlsx-extractor.svg)](https://badge.fury.io/js/xlsx-extractor)
[![Build Status](https://travis-ci.org/akabekobeko/npm-xlsx-extractor.svg?branch=master)](https://travis-ci.org/akabekobeko/npm-xlsx-extractor)
[![Document](https://doc.esdoc.org/github.com/akabekobeko/npm-xlsx-extractor/badge.svg?t=0)](https://doc.esdoc.org/github.com/akabekobeko/npm-xlsx-extractor)
[![Standard - JavaScript Style Guide](https://img.shields.io/badge/code_style-standard-brightgreen.svg)](http://standardjs.com/)

Extract the colums/rows from XLSX file.

## Installation

```
$ npm install xlsx-extractor
```

## Usage

### Node API

```js
const XlsxExtractor = require('xlsx-extractor');

const extractor = new XlsxExtractor('./sample.xlsx');
const tasks     = [];
for (let i = 1, max = extractor.count; i <= max; ++i) {
  tasks.push(extractor.extract(i));
}

Promise
.all(tasks)
.then((results) => {
  console.log(JSON.stringify(results, null, '  ') + '\n');
} )
.catch((err) => {
  console.error(err);
});
```

**constructor**

`new XlsxExtractor( path )`

| Name | Type | Description |
|:--------|:--|:--|
| `path` | `String` | Path of the XLSX file. |

**count**

`XlsxExtractor.count` is a number of sheets.

**extract**

`XlsxExtractor.extract( index )` is promisify.

| Name | Type | Description |
|:--------|:--|:--|
| `index` | `Number` | Number of the extract sheet. |

result:

| Name | Type | Description |
|:--------|:--|:--|
| `id` | `Number` | Number of the extract sheets. |
| `name` | `String` | Name of the sheet. |
| `cells` | `Array<Array.<String>>` | Cells of the sheet. Empty cell is stored is `""`. |

### CLI

```
Usage: xlsx-extractor [OPTIONS]

  Extract the colums/rows from XLSX file.

  Options:
    -h, --help    Display this text.

    -v, --version Display the version number.

    -i, --input   Path of the XLSX file.

    -r, --range   Range of sheets to be output.
                  Specify the numeric value of "N" or "N-N".
                  When omitted will output all of the sheet.

    -c, --count   Outputs the number of sheet.
                  This option overrides the -r and --range.

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
