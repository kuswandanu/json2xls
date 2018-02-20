var nodeExcel = require('node-excel-export');

var transform = function (json, config) {
    var conf = transform.prepareJson(json, config);
    var result = nodeExcel.buildExport(conf);
    return result;
};

//prepare json to be in the correct format for excel-export
transform.prepareJson = function (json, config) {
    var conf = config || {};
    var jsonArr = [].concat(json);
    var fields = conf.fields || Object.keys(jsonArr[0] || {});
    if (!(fields instanceof Array)) {
        fields = Object.keys(fields);
    }
    var name = conf.name || 'Sheet';
    var heading = conf.heading || [];
    var merges = conf.merges || [];
    var specification = conf.specification || fields.reduce(function (o, key) {
        return Object.assign(o, {
            [key]: {
                displayName: key,
                headerStyle: { font: { bold: true } },
                width: 100
            }
        });
    }, {});
    return [
        {
            name: name,
            heading: heading,
            merges: merges,
            specification: specification,
            data: jsonArr
        }
    ];
};

transform.middleware = function (req, res, next) {
    res.xls = function (fn, data, config) {
        var xls = transform(data, config);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats');
        res.setHeader("Content-Disposition", "attachment; filename=" + fn);
        res.end(xls, 'binary');
    };
    next();
};

module.exports = transform;
