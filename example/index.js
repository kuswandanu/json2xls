var json2xls = require('../lib/json2xls');
var data = require('../spec/arrayData.json');
var fs = require('fs');


var dataset = [
    { customer_name: 'IBM', status_id: 1, note: 'some note', misc: 'not shown' },
    { customer_name: 'HP', status_id: 0, note: 'some note' },
    { customer_name: 'MS', status_id: 0, note: 'some note', misc: 'not shown' }
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
// var xls = json2xls(data, {});


fs.writeFileSync(__dirname + '/output.xlsx', xls, 'binary');