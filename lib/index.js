/**
 * @fileOverview main script
 * @name index.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 */
'use strict';

var ExcelParser = require('./excel/parser'),
    SpreadsheetParser = require('./spreadsheet/parser');

module.exports = {
    excel: new ExcelParser(),
    ExcelParser: ExcelParser,
    spreadsheet: new SpreadsheetParser(),
    SpreadsheetParser: SpreadsheetParser
};
