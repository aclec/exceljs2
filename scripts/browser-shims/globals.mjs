// Injected by esbuild into the browser bundle: provides the Node globals that lib/ and its
// dependencies (readable-stream, jszip) expect.
import {Buffer} from 'buffer';
import process from 'process/browser.js';

process.browser = true;

export {Buffer, process};
