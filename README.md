XLSXParser
==========

The XLSXParser can parsing **Microsoft Excel** and **Google Spreadsheet**.

## Installation
```
npm install xlsx-parser
```

## Usage
### Quick start
#### excel
example sheet.xlsx

|   | A | B | C | D |
|:-:|--:|--:|--:|--:|
| 1 | key | name | number | other |
| 2 | aa1 | cat | 1 | |
| 3 | bb1 | dog | 2 | sleep |
| 4 | cc1 | mouse | 3 | run |
Sheet1

```
var xlsxParser = require('xlsx-parser');
var opts = { sheets: [ 1 ] };
elsxPaser.excel.parse('./sheet.xlsx', opts, function(err, result) {
    console.log(result);
    /* output
        [{
            num: 1,         // sheet number
            name: 'Sheet1', // sheet name
            cells: [
                { cell: 'A1', column: 'A', row: 1, value: 'key' },
                { cell: 'B1', column: 'B', row: 1, value: 'name' },
                { cell: 'C1', column: 'C', row: 1, value: 'number' },
                { cell: 'D1', column: 'D', row: 1, value: 'other' },
                { cell: 'A2', column: 'A', row: 2, value: 'aa1' },
                { cell: 'B2', column: 'B', row: 2, value: 'cat' },
                { cell: 'C2', column: 'C', row: 2, value: '1' },
                { cell: 'A3', column: 'A', row: 3, value: 'bb1' },
                { cell: 'B3', column: 'B', row: 3, value: 'dog' },
                { cell: 'C3', column: 'C', row: 3, value: '2' },
                { cell: 'D3', column: 'D', row: 3, value: 'sleep' },
                { cell: 'A4', column: 'A', row: 4, value: 'cc1' },
                { cell: 'B4', column: 'B', row: 4, value: 'mouse' },
                { cell: 'C4', column: 'C', row: 4, value: '3' },
                { cell: 'D4', column: 'D', row: 4, value: 'run' }
            ]
        }]
    */
});
```

#### spreadsheet
example spreadsheet

|   | A | B | C | D |
|:-:|--:|--:|--:|--:|
| 1 | key | name | number | other |
| 2 | aa1 | cat | 1 | |
| 3 | bb1 | dog | 2 | sleep |
| 4 | cc1 | mouse | 3 | run |
Sheet1

```
var xlsxParser = require('xlsx-parser');

// setup
xlsxPaeser.spreadsheet.setup({
    api: {
        client_id: 'YOUR CLIENT ID HERE',
        client_secret: 'YOUR CLIENT SECRET HERE',
        redirect_url: 'http://localhost',
        token_path: './dist/token.json'
    }
});

// spreadsheet id
var SPREADSHEET_KEY = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo';
// worksheet ids
var opts = { sheets: [ 'od6' ] };
elsxPaser.spreadsheet.parse(SPREADSHEET_KEY, opts, function(err, result) {
    console.log(result);
    /* output
        [{
            id: 'od6',      // sheet id
            name: 'Sheet1', // sheet name
            cells: [
                { cell: 'A1', column: 'A', row: 1, value: 'key' },
                { cell: 'B1', column: 'B', row: 1, value: 'name' },
                { cell: 'C1', column: 'C', row: 1, value: 'number' },
                { cell: 'D1', column: 'D', row: 1, value: 'other' },
                { cell: 'A2', column: 'A', row: 2, value: 'aa1' },
                { cell: 'B2', column: 'B', row: 2, value: 'cat' },
                { cell: 'C2', column: 'C', row: 2, value: '1' },
                { cell: 'A3', column: 'A', row: 3, value: 'bb1' },
                { cell: 'B3', column: 'B', row: 3, value: 'dog' },
                { cell: 'C3', column: 'C', row: 3, value: '2' },
                { cell: 'D3', column: 'D', row: 3, value: 'sleep' },
                { cell: 'A4', column: 'A', row: 4, value: 'cc1' },
                { cell: 'B4', column: 'B', row: 4, value: 'mouse' },
                { cell: 'C4', column: 'C', row: 4, value: '3' },
                { cell: 'D4', column: 'D', row: 4, value: 'run' }
            ]
        }]
    */
});

```

## Contribution
1. Fork it ( https://github.com/iyu/xlsx-parser/fork )
2. Create a feature branch
3. Commit your changes
4. Rebase your local changes against the master branch
5. Run test suite with the `npm test; npm run-script jshint` command and confirm that it passes
5. Create new Pull Request
