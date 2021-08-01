exports.openZip = openZip;

var fs = require("fs");

var promises = require("./promises");
var zipfile = require("./zipfile");

exports.openZip = openZip;

function openZip(options) {
    return promises.resolve(zipfile.openArrayBuffer(options));
}
