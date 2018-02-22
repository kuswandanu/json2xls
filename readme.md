json2xls
========

utility to convert json to a excel file, based on [Node-Excel-Export](https://github.com/functionscope/Node-Excel-Export)

About this fork
------

this fork based on [Node.JS Excel-Export](https://github.com/andreyan-andreev/node-excel-export)

Usage
------

Use to save as file:

    var json2xls = require('json2xls');
    var fs = require('fs');

    var dataset = [
        { customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown', confirmed: false },
        { customer_name: 'HP', status_id: 0, note: 'some note', confirmed: true },
        { customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown', confirmed: false }
    ];

    var xls = json2xls(dataset);

    fs.writeFileSync(__dirname + '/output.xlsx', xls, 'binary');

Or use as restify or express middleware. It adds a convenience `xls` method to the response object to immediately output an excel as download.

    var json2xls = require('../lib/json2xls');
    var restify = require('restify');

    var server = restify.createServer();

    var dataset = [
        { customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown', confirmed: false },
        { customer_name: 'HP', status_id: 0, note: 'some note', confirmed: true },
        { customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown', confirmed: false }
    ];

    server.use(json2xls.middleware);

    server.get('/', function (req, res, next) {
        res.xls('data.xlsx', dataset);
        return next();
    });

    server.listen(8080, function () {
        console.log('%s listening at %s', server.name, server.url);
    });

Options
-------

Add styles

    var json2xls = require('json2xls');
    var fs = require('fs');

    var dataset = [
        { customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown', confirmed: false },
        { customer_name: 'HP', status_id: 0, note: 'some note', confirmed: true },
        { customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown', confirmed: false }
    ];

    var styles = {
        headerDark: {
            fill: {
                fgColor: {
                    rgb: 'FF000000'
                }
            },
            font: {
                color: {
                    rgb: 'FFFFFFFF'
                },
                sz: 14,
                bold: true,
                underline: true
            }
        },
        cellPink: {
            fill: {
                fgColor: {
                    rgb: 'FFFFCCFF'
                }
            }
        },
        cellGreen: {
            fill: {
                fgColor: {
                    rgb: 'FF00FF00'
                }
            }
        }
    };

    var options = {};
    options.specification = {
        customer_name: {
            displayName: 'Customer',
            headerStyle: styles.headerDark,
            cellStyle: function (value, row) {
                return (row.status_id == 1) ? styles.cellGreen : { fill: { fgColor: { rgb: 'FFFF0000' } } };
            },
            width: 120
        },
        status_id: {
            displayName: 'Status',
            headerStyle: styles.headerDark,
            cellFormat: function (value, row) {
                return (value == 1) ? 'Active' : 'Inactive';
            },
            width: '10'
        },
        note: {
            displayName: 'Description',
            headerStyle: styles.headerDark,
            cellStyle: styles.cellPink,
            width: 220
        }
    };

    var xls = json2xls(dataset, options);

    fs.writeFileSync(__dirname + '/output.xlsx', xls, 'binary');

