var path = require('path');

var assert = require('power-assert'),
    sinon = require('sinon');

var xlsxParser = require('../lib');

describe('lib/index.js', function() {
    
    it('xlsxParser.excel', function() {
        assert.ok(xlsxParser.excel);
        assert.ok(xlsxParser.ExcelParser);
        assert.deepEqual(xlsxParser.excel, new xlsxParser.ExcelParser());
    });

    it('xlsxParser.spreadsheet', function() {
        assert.ok(xlsxParser.spreadsheet);
        assert.ok(xlsxParser.SpreadsheetParser);
        assert.deepEqual(xlsxParser.spreadsheet, new xlsxParser.SpreadsheetParser());
    });
});

describe('lib/excel/parse.js', function() {
    var excelParser = xlsxParser.excel;

    describe('ExcelParser#parse()', function() {
        it('file: test/data/excel.xlsx, sheet: 1', function(done) {
            excelParser.parse(path.join(__dirname, './data/excel.xlsx'), { sheets: [ 1 ] }, function(err, result) {
                assert(err == null);
                assert(result);
                assert(result.length === 1);
                assert(result[0].num === 1);
                assert(result[0].name === 'Sheet1');
                assert.deepEqual(result[0].cells, [
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
                ]);
                done();
            });
        });
    });
});

describe('lib/spreadsheet/parse.js', function() {
    var spreadsheetParser = xlsxParser.spreadsheet,
        SPREADSHEET_KEY = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo',
        WORKSHEET_KEYS = [ 'od6' ];

    before(function() {
        var api = require('../lib/spreadsheet/api'),
            opts = require('./data/spreadsheet.opts'),
            mock = require('./data/spreadsheet.mock'),
            sandbox = sinon.sandbox.create();

        spreadsheetParser.setup({ api: opts.installed });
        sandbox.stub(api, 'getWorksheet').yields(null, mock.getWorksheet);
        var stubGetCells = sandbox.stub(api, 'getCells');
        for (var key in mock.getCells) {
            stubGetCells.withArgs(SPREADSHEET_KEY, key).yields(null, mock.getCells[key]);
        }
    });

    describe('SpreadsheetParser#parse()', function() {
        it('link: https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo, sheet: 1', function(done) {
            spreadsheetParser.parse(SPREADSHEET_KEY, { sheets: WORKSHEET_KEYS }, function(err, result) {
                assert(err == null);
                assert(result);
                assert(result.length === 1);
                assert(result[0].id === 'od6');
                assert(result[0].name === 'Test1');
                assert.deepEqual(result[0].cells[0], { cell: 'A1', column: 'A', row: 1, value: '{ \"name\": \"Test1\" }' });
                assert.deepEqual(result[0].cells[1], { cell: 'A2', column: 'A', row: 2, value: '_id' });
                assert.deepEqual(result[0].cells[2], { cell: 'B2', column: 'B', row: 2, value: 'str' });
                assert.deepEqual(result[0].cells[3], { cell: 'C2', column: 'C', row: 2, value: 'num:number' });
                assert.deepEqual(result[0].cells[4], { cell: 'D2', column: 'D', row: 2, value: 'date:date' });
                assert.deepEqual(result[0].cells[5], { cell: 'E2', column: 'E', row: 2, value: 'bool:boolean' });
                assert.deepEqual(result[0].cells[6], { cell: 'F2', column: 'F', row: 2, value: 'obj.type1' });
                assert.deepEqual(result[0].cells[7], { cell: 'G2', column: 'G', row: 2, value: 'obj.type2' });
                assert.deepEqual(result[0].cells[8], { cell: 'H2', column: 'H', row: 2, value: '#arr:number' });
                assert.deepEqual(result[0].cells[9], { cell: 'I2', column: 'I', row: 2, value: '#lists.code' });
                assert.deepEqual(result[0].cells[10], { cell: 'J2', column: 'J', row: 2, value: '#lists.bool:boolean' });
                assert.deepEqual(result[0].cells[11], { cell: 'A4', column: 'A', row: 4, value: 'test1_1' });
                assert.deepEqual(result[0].cells[12], { cell: 'B4', column: 'B', row: 4, value: 'testtest' });
                assert.deepEqual(result[0].cells[13], { cell: 'C4', column: 'C', row: 4, value: '1' });
                assert.deepEqual(result[0].cells[14], { cell: 'D4', column: 'D', row: 4, value: '2014/08/01 10:00:00' });
                done();
            });
        });
    });
});
