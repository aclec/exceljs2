// this module lets the specs require lib modules by path relative to lib/ (compiled from src/ by `bun run build:lib`).
// A module compiled from `export default X` is returned as X, like the CommonJS `module.exports = X` it replaced;
// modules with named exports only are returned as is.

const libs = {
    exceljs: require("../../lib/exceljs.nodejs"),
};
const basePath = "../../lib/";

module.exports = function verquire(path) {
    if (!libs[path]) {
        const mod = require(basePath + path);
        libs[path] = mod && mod.__esModule && "default" in mod ? mod.default : mod;
    }
    return libs[path];
};
