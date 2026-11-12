// this module lets the specs require lib modules by path relative to lib/
// (it used to switch to the transpiled dist/es5 build, which was dropped in v5)

const libs = {
    exceljs: require("../../lib/exceljs.nodejs"),
};
const basePath = "../../lib/";

module.exports = function verquire(path) {
    if (!libs[path]) {
        libs[path] = require(basePath + path);
    }
    return libs[path];
};
