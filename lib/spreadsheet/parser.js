/**
 * @fileOverview spreadsheet parser
 * @name parser.js
 * @author Yuhei Aihara <aihara_yuhei@cyberagent.co.jp>
 */
'use strict';

var fs = require('fs'),
    path = require('path');

var _ = require('lodash'),
    async = require('async');

var api = require('./api'),
    logger = require('../logger');

function SpreadsheetParser() {
    this.api = {
        client_id: undefined,
        client_secret: undefined,
        redirect_url: 'http://localhost',
        token_path: './dist/token.json'
    };
}

module.exports = SpreadsheetParser;

/**
 * API function
 */
SpreadsheetParser.prototype.generateAuthUrl = function() { return api.generateAuthUrl.apply(api, arguments); };
SpreadsheetParser.prototype.getAccessToken = function() { return api.getAccessToken.apply(api, arguments); };
SpreadsheetParser.prototype.refreshAccessToken = function() { return api.refreshAccessToken.apply(api, arguments); };

/**
 * setup
 * @param {Object} options
 * @example
 * var opts = {
 *     api: {
 *         client_id: 'YOUR CLIENT ID HERE',
 *         client_secret: 'YOUR CLIENT SECRET HERE',
 *         redirect_url: 'http://localhost',
 *         token_path: './dist/token.json'
 *     },
 * }
 */
SpreadsheetParser.prototype.setup = function(opts) {
    _.extend(this.api, opts && opts.api || {});

    var tokenDir = path.dirname(this.api.token_path);
    if (!fs.existsSync(tokenDir)) {
        fs.mkdirSync(tokenDir);
    }
    api.setup.call(api, this.api);
};

/**
 * get worksheet info in spreadsheet
 * @param {String} key spreadsheetKey
 * @param {Function} callback
 * @example
 * > url = 'https://docs.google.com/spreadsheets/d/1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo/edit#gid=0'
 * key = '1YXVzaaxqkPKsr-excIOXScnTQC7y_DKrUKs0ukzSIgo'
 */
SpreadsheetParser.prototype.getWorksheet = function(key, callback) {
    api.getWorksheet(key, function(err, result) {
        if (err) {
            return callback(err);
        }
        var entry = result && result.feed && result.feed.entry;
        if (!entry) {
            return callback(new Error('not found entry.'), result);
        }

        var list = [];
        for (var i = 0; i < entry.length; i++) {
            var worksheet = entry[i];
            list.push({
                id: worksheet.id && worksheet.id.$t && worksheet.id.$t.replace(/.+\/([^\/]+)$/, '$1'),
                updated: worksheet.updated && worksheet.updated.$t,
                title: worksheet.title && worksheet.title.$t
            });
        }
        callback(null, list);
    });
};

/**
 * get worksheets data
 * @param {String} key spreadsheetKey
 * @param {Object} opts
 * @param {Function} callback
 * @see #getWorksheet arguments.key and returns list.id
 */
SpreadsheetParser.prototype.parse = function(key, opts, callback) {
    var worksheetIds = opts.sheets;
    async.map(worksheetIds, function(worksheetId, next) {
        api.getCells(key, worksheetId, function(err, result) {
            if (err) {
                return next(err);
            }

            var sheetName = result && result.feed && result.feed.title.$t;
            var entry = result && result.feed && result.feed.entry;
            if (!entry) {
                logger.error('not found entry.', worksheetId);
                return next();
            }

            var cells = _.map(entry, function(cell) {
                var title = cell.title.$t.match(/^(\D+)(\d+)$/);
                return {
                    cell: cell.title.$t,
                    column: title[1],
                    row: parseInt(title[2], 10),
                    value: cell.content.$t
                };
            });

            next(null, {
                id: worksheetId,
                name: sheetName,
                cells: cells
            });
        });
    }, function(err, result) {
        if (err) {
            return callback(err);
        }

        callback(null, _.compact(result));
    });
};
