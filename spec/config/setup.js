const chai = require("chai");

const chaiXml = require("chai-xml");
const chaiDatetime = require("chai-datetime");
const dirtyChaiMod = require("dirty-chai");
const dirtyChai = dirtyChaiMod.default || dirtyChaiMod;
global.verquire = require("../utils/verquire");

global.expect = chai.expect;

chai.use(chaiXml);
chai.use(chaiDatetime);
chai.use(dirtyChai);
