'use strict';

// Runs the dist/ bundles in headless Google Chrome (driven by playwright-core).
// - uses the locally installed Chrome (channel 'chrome') or CHROME_PATH; never downloads a browser
// - uses a fresh temporary profile that is deleted afterwards (never the user's profile)
// - serves the bundles from 127.0.0.1 only
// Run with `bun run build && bun run test:browser`.

const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const {chromium} = require('playwright-core');
const Encryptor = require('../../lib/utils/encryptor');

const DIST = path.resolve(__dirname, '../../dist');
const BUNDLES = ['exceljs.js', 'exceljs.min.js', 'exceljs.bare.js', 'exceljs.bare.min.js'];

function startServer() {
  const files = new Map(BUNDLES.map(name => [`/dist/${name}`, path.join(DIST, name)]));
  const server = http.createServer((req, res) => {
    const {pathname, searchParams} = new URL(req.url, 'http://127.0.0.1');
    if (pathname === '/') {
      const bundle = searchParams.get('bundle');
      if (!BUNDLES.includes(bundle)) {
        res.writeHead(404).end();
        return;
      }
      res.writeHead(200, {'Content-Type': 'text/html; charset=utf-8'});
      res.end(`<!doctype html><html><head><link rel="icon" href="data:,"><script src="/dist/${bundle}"></script></head><body></body></html>`);
      return;
    }
    if (files.has(pathname)) {
      res.writeHead(200, {'Content-Type': 'text/javascript; charset=utf-8'});
      fs.createReadStream(files.get(pathname)).pipe(res);
      return;
    }
    res.writeHead(404).end();
  });
  return new Promise((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => resolve(server));
  });
}

async function launchChrome(userDataDir) {
  const options = {headless: true};
  if (process.env.CHROME_PATH) {
    options.executablePath = process.env.CHROME_PATH;
  } else {
    options.channel = 'chrome';
  }
  try {
    return await chromium.launchPersistentContext(userDataDir, options);
  } catch (error) {
    throw new Error(
      'Could not start Google Chrome for the browser tests. Install Chrome stable, or point ' +
        `CHROME_PATH at a Chrome/Chromium executable. Browsers are never downloaded.\n${error.message}`,
      {cause: error},
    );
  }
}

describe('ExcelJS in the browser', function() {
  this.timeout(60000);

  let server;
  let baseUrl;
  let userDataDir;
  let context;

  before(async () => {
    const missing = BUNDLES.filter(name => !fs.existsSync(path.join(DIST, name)));
    if (missing.length) {
      throw new Error(`Missing ${missing.join(', ')} in dist/: run \`bun run build\` first`);
    }
    server = await startServer();
    baseUrl = `http://127.0.0.1:${server.address().port}`;
    userDataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-chrome-'));
    context = await launchChrome(userDataDir);
  });

  after(async () => {
    if (context) await context.close();
    if (server) await new Promise(resolve => server.close(resolve));
    if (userDataDir) fs.rmSync(userDataDir, {recursive: true, force: true});
  });

  BUNDLES.forEach(bundle => {
    describe(bundle, () => {
      let page;
      let errors;

      beforeEach(async () => {
        errors = [];
        page = await context.newPage();
        page.on('pageerror', error => errors.push(error.message));
        page.on('console', message => {
          if (message.type() === 'error') errors.push(message.text());
        });
        await page.goto(`${baseUrl}/?bundle=${bundle}`);
      });

      afterEach(async () => {
        await page.close();
        expect(errors).to.deep.equal([]);
      });

      it('exposes the ExcelJS global', async () => {
        const result = await page.evaluate(() => ({
          workbook: typeof ExcelJS.Workbook,
          valueType: ExcelJS.ValueType.Number,
        }));
        expect(result).to.deep.equal({workbook: 'function', valueType: 2});
      });

      it('reads and writes xlsx via binary buffer', async () => {
        const result = await page.evaluate(async () => {
          const wb = new ExcelJS.Workbook();
          const ws = wb.addWorksheet('blort');
          ws.getCell('A1').value = 'Hello, World!';
          ws.getCell('A2').value = 7;
          ws.getCell('A3').value = new Date(Date.UTC(2020, 0, 2));
          const buffer = await wb.xlsx.writeBuffer();

          const wb2 = new ExcelJS.Workbook();
          await wb2.xlsx.load(buffer);
          const ws2 = wb2.getWorksheet('blort');
          return [ws2.getCell('A1').value, ws2.getCell('A2').value, ws2.getCell('A3').value.toISOString()];
        });
        expect(result).to.deep.equal(['Hello, World!', 7, '2020-01-02T00:00:00.000Z']);
      });

      it('reads and writes xlsx via base64 buffer', async () => {
        const result = await page.evaluate(async () => {
          const options = {base64: true};
          const wb = new ExcelJS.Workbook();
          const ws = wb.addWorksheet('blort');
          ws.getCell('A1').value = 'Hello, World!';
          ws.getCell('A2').value = 7;
          const buffer = await wb.xlsx.writeBuffer(options);

          const wb2 = new ExcelJS.Workbook();
          await wb2.xlsx.load(buffer.toString('base64'), options);
          const ws2 = wb2.getWorksheet('blort');
          return [ws2.getCell('A1').value, ws2.getCell('A2').value];
        });
        expect(result).to.deep.equal(['Hello, World!', 7]);
      });

      it('writes csv via buffer', async () => {
        const result = await page.evaluate(async () => {
          const wb = new ExcelJS.Workbook();
          const ws = wb.addWorksheet('blort');
          ws.getCell('A1').value = 'Hello, World!';
          ws.getCell('B1').value = 'What time is it?';
          ws.getCell('A2').value = 7;
          ws.getCell('B2').value = '12pm';
          ws.getCell('C2').value = new Date(Date.UTC(2020, 0, 2));
          const buffer = await wb.csv.writeBuffer({dateFormat: 'YYYY-MM-DD', dateUTC: true});
          return buffer.toString();
        });
        expect(result).to.equal('"Hello, World!",What time is it?\n7,12pm,2020-01-02');
      });

      it('generates conditional formatting ids without crypto.randomUUID', async () => {
        const result = await page.evaluate(async () => {
          // randomUUID is missing in insecure contexts (plain http on a non-local host)
          Object.defineProperty(globalThis.crypto, 'randomUUID', {value: undefined});
          const wb = new ExcelJS.Workbook();
          const ws = wb.addWorksheet('cf');
          ws.getCell('A1').value = 1;
          ws.getCell('A2').value = 5;
          ws.addConditionalFormatting({
            ref: 'A1:A2',
            rules: [{type: 'dataBar', cfvo: [{type: 'min'}, {type: 'max'}], color: {argb: 'FF638EC6'}}],
          });
          const buffer = await wb.xlsx.writeBuffer();
          const wb2 = new ExcelJS.Workbook();
          await wb2.xlsx.load(buffer);
          return wb2.getWorksheet('cf').conditionalFormattings[0].rules[0].x14Id;
        });
        expect(result).to.match(/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
      });

      it('protects a sheet with the same SHA-512 hash as node:crypto', async () => {
        const result = await page.evaluate(async () => {
          const wb = new ExcelJS.Workbook();
          const ws = wb.addWorksheet('p');
          await ws.protect('s3cret é€', {spinCount: 1000});
          return ws.sheetProtection;
        });
        expect(result.algorithmName).to.equal('SHA-512');
        expect(result.spinCount).to.equal(1000);
        expect(Buffer.from(result.saltValue, 'base64')).to.have.length(16);
        expect(result.hashValue).to.equal(
          Encryptor.convertPasswordToHash('s3cret é€', 'SHA512', result.saltValue, 1000),
        );
      });
    });
  });
});
