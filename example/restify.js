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
