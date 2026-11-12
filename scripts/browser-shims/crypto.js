"use strict";

// Browser replacement for the subset of node:crypto used by lib/utils/encryptor.js
// (getHashes, createHash(...).update().digest(), randomBytes). Only bundled into dist/;
// Node and Bun keep using the native node:crypto module.
const { md5, sha1 } = require("@noble/hashes/legacy.js");
const { sha224, sha256, sha384, sha512 } = require("@noble/hashes/sha2.js");
const { Buffer } = require("buffer");

const algorithms = { md5, sha1, sha224, sha256, sha384, sha512 };
// node:crypto accepts both 'sha512' and 'sha-512' (case-insensitive)
const normalise = (name) => String(name).toLowerCase().replace(/^sha-/, "sha");

module.exports = {
    getHashes() {
        const names = Object.keys(algorithms);
        return names.concat(names.filter((n) => n.startsWith("sha")).map((n) => `sha-${n.slice(3)}`));
    },
    createHash(name) {
        const algorithm = algorithms[normalise(name)];
        if (!algorithm) {
            throw new Error(`Digest method not supported: ${name}`);
        }
        const hash = algorithm.create();
        return {
            update(data) {
                hash.update(typeof data === "string" ? Buffer.from(data, "utf8") : data);
                return this;
            },
            digest(encoding) {
                const result = Buffer.from(hash.digest());
                return encoding ? result.toString(encoding) : result;
            },
        };
    },
    randomBytes(size) {
        return Buffer.from(globalThis.crypto.getRandomValues(new Uint8Array(size)));
    },
};
