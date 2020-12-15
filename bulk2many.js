const json2xls = require('json2xls-xml')({ pretty: false });
const _ = require('lodash');
const csv = require("csvtojson");
const jsonfile = require('jsonfile');
const path = require('path');

const fs = require('fs');


var createFolder = function (loc) {
    if (!fs.existsSync(loc)) {
        fs.mkdirSync(loc);
    }
}

var walk = function (dir) {
    var results = [];
    var list = fs.readdirSync(dir);
    list.forEach(function (file) {
        file = dir + '/' + file;
        var stat = fs.statSync(file);
        if (stat && stat.isDirectory()) {
            /* Recurse into a subdirectory */
            results = results.concat(walk(file));
        } else {
            /* Is a file */
            results.push(file);
        }
    });
    return results;
}


//Load Config File
try {
    var config = jsonfile.readFileSync("config.json");
    var INPUT_FOLDER = process.argv[2] || config.CSV_FOLDER;
    var OUTPUT_FOLDER = process.argv[3] || config.XLS_FOLDER;
    var OUTPUT_FILE = process.argv[4] || config.XLS_FILENAME;
    var IGNORE_NON_CSV = config.IGNORE_NON_CSV;

    createFolder(INPUT_FOLDER);
    createFolder(OUTPUT_FOLDER);
} catch (e) {
    console.log("[error] " + new Date().toGMTString() + " : Config/Params Invalid");
    return process.exit(-1);
}

let csv_files = walk(INPUT_FOLDER);

(async () => {

    for (csv_file of csv_files) {
        var json_data = {};

        if (IGNORE_NON_CSV && csv_file.search(/\.csv$/igs) < 0) {
            console.log("[skipping] " + new Date().toGMTString() + " : " + csv_file);
            continue;
        }
        console.log("[processing] " + new Date().toGMTString() + " : " + csv_file);
        let jsonArray = await csv().fromFile(csv_file);
        //Using Sheet Name as File Name
        let sheet_name = csv_file.split("/").pop().split(".")[0].split("-")[0].trim();

        //Memory Caution Here; Very HUGE CSV may Break Here
        json_data[sheet_name] = jsonArray;

        //Convert JSON to XLS
        console.log("[converting] " + new Date().toGMTString() + " : " + csv_file);
        let xls_data = json2xls(json_data);

        // Bulk Writing to XLS
        console.log("[writing] " + new Date().toGMTString() + " : " + csv_file);
        fs.writeFileSync(path.join(OUTPUT_FOLDER, sheet_name + ".xls"), xls_data);
    }
})();

