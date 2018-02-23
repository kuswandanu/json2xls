json2xls
========

utility to convert json to a excel file, based on [Node-Excel-Export](https://github.com/functionscope/Node-Excel-Export)

About this fork
------

this fork based on [Node.JS Excel-Export](https://github.com/andreyan-andreev/node-excel-export)

Usage
------

Use to save as file:

```javascript
var json2xls = require('json2xls');
var fs = require('fs');

var dataset = [
    { customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown', confirmed: false },
    { customer_name: 'HP', status_id: 0, note: 'some note', confirmed: true },
    { customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown', confirmed: false }
];

var xls = json2xls(dataset);

fs.writeFileSync(__dirname + '/output.xlsx', xls, 'binary');
```

Or use as restify or express middleware. It adds a convenience `xls` method to the response object to immediately output an excel as download.

```javascript
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
```

Options
-------

Add sheet name, styles, header, and merge cell

```javascript
var json2xls = require('../lib/json2xls');
var fs = require('fs');

var dataset = [
    { customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown', confirmed: false },
    { customer_name: 'HP', status_id: 0, note: 'some note', confirmed: true },
    { customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown', confirmed: false }
];

const styles = {
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

options.name = "Report";

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
    },
    misc: {
        displayName: 'Misc',
        headerStyle: styles.headerDark,
        cellFormat: function (value, row) {
            return (value || 'null data');
        },
        width: 150
    },
    confirmed: {
        displayName: 'Confirmed',
        headerStyle: styles.headerDark,
        cellStyle: function (value, row) {
            return (value ? styles.cellGreen : styles.cellPink);
        },
        width: 220
    }
};

options.heading = [
    [
        { value: 'a1', style: styles.headerDark },
        { value: 'b1', style: styles.headerDark },
        { value: 'c1', style: styles.headerDark },
        { value: 'd1', style: styles.headerDark },
        { value: 'e1', style: styles.headerDark }
    ],
    ['a2', 'b2', 'c2', 'd2', 'e2']
];

options.merges = [
    { start: { row: 1, column: 1 }, end: { row: 1, column: 2 } },
    { start: { row: 1, column: 3 }, end: { row: 1, column: 4 } },
    { start: { row: 2, column: 1 }, end: { row: 2, column: 2 } },
    { start: { row: 2, column: 4 }, end: { row: 2, column: 5 } },
    { start: { row: 3, column: 4 }, end: { row: 4, column: 4 } },
    { start: { row: 5, column: 4 }, end: { row: 6, column: 4 } }
];

var xls = json2xls(dataset, options);

fs.writeFileSync(__dirname + '/output.xlsx', xls, 'binary');
```

Cells Style
------
* Check [here](https://github.com/protobi/js-xlsx#cell-styles), for more styling

