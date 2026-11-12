// Streaming reader (ExcelJS.stream.xlsx.WorkbookReader) robustness tests.
// Fixtures are built in-test with ExcelJS + JSZip so they stay reviewable.
const fs = require('fs');
const {Readable} = require('stream');
const JSZip = require('jszip');

const {baseWorkbookBuffer} = require('../utils/xlsx-tamper');

const ExcelJS = verquire('exceljs');

// feed the buffer in small chunks so that chunk boundaries fall everywhere
function chunkedStream(buffer, size = 1024) {
  return Readable.from(
    (function*() {
      for (let i = 0; i < buffer.length; i += size) {
        yield buffer.subarray(i, i + size);
      }
    })(),
    {objectMode: false}
  );
}

async function readAll(input, options) {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(input, options);
  const sheets = [];
  for await (const worksheet of reader) {
    const rows = [];
    for await (const row of worksheet) {
      rows.push({number: row.number, values: row.values});
    }
    sheets.push({id: worksheet.id, name: worksheet.name, worksheet, rows});
  }
  return sheets;
}

// rebuild the zip with its parts sorted by rank(path) (stable: equal ranks keep their order)
async function reorderZip(buffer, rank) {
  const source = await JSZip.loadAsync(buffer);
  const names = Object.keys(source.files).filter(name => !source.files[name].dir);
  const ordered = names
    .map((name, index) => ({name, index, rank: rank(name)}))
    .sort((a, b) => a.rank - b.rank || a.index - b.index)
    .map(({name}) => name);
  const target = new JSZip();
  for (const name of ordered) {
    // eslint-disable-next-line no-await-in-loop
    target.file(name, await source.file(name).async('nodebuffer'));
  }
  return target.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'});
}

const isWorksheet = name => /^xl\/worksheets\/sheet\d+\.xml$/.test(name);

// records every temp dir the reader creates (spy on the lookup done at call time)
function trackTempDirs() {
  const created = [];
  const original = fs.promises.mkdtemp;
  fs.promises.mkdtemp = async (...args) => {
    const dir = await original.apply(fs.promises, args);
    created.push(dir);
    return dir;
  };
  return {
    created,
    restore() {
      fs.promises.mkdtemp = original;
    },
  };
}

async function bigSheetFirstBuffer() {
  const buffer = await baseWorkbookBuffer((wb, ws) => {
    for (let i = 2; i <= 3000; i++) {
      ws.getCell(`A${i}`).value = `row ${i} ${(i * 7919) % 10007}`;
      ws.getCell(`B${i}`).value = i;
    }
    wb.addWorksheet('Second').getCell('A1').value = 'second';
  });
  return reorderZip(buffer, name => (isWorksheet(name) ? 0 : 1));
}

describe('WorkbookReader (stream)', () => {
  describe('temp files for worksheets read before shared strings', () => {
    let tracker;
    beforeEach(() => {
      tracker = trackTempDirs();
    });
    afterEach(() => {
      tracker.restore();
    });

    it('removes the temp dir when the consumer stops early', async () => {
      const buffer = await bigSheetFirstBuffer();
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer));
      let firstRow;
      // eslint-disable-next-line no-unreachable-loop
      for await (const worksheet of reader) {
        // eslint-disable-next-line no-unreachable-loop
        for await (const row of worksheet) {
          firstRow = row.values;
          break;
        }
        break;
      }
      expect(firstRow[1]).to.equal('hello');
      expect(tracker.created).to.have.length(1);
      expect(fs.existsSync(tracker.created[0])).to.equal(false);
    });

    it('closes the spooled worksheet files the consumer skips', async () => {
      const opened = [];
      const original = fs.createReadStream;
      fs.createReadStream = (...args) => {
        const stream = original.apply(fs, args);
        opened.push(stream);
        return stream;
      };
      try {
        const buffer = await bigSheetFirstBuffer();
        const names = [];
        const reader = new ExcelJS.stream.xlsx.WorkbookReader(chunkedStream(buffer));
        for await (const worksheet of reader) {
          names.push(worksheet.name);
        }
        expect(names).to.deep.equal(['Sheet1', 'Second']);
      } finally {
        fs.createReadStream = original;
      }
      expect(opened).to.have.length(2);
      opened.forEach(stream => expect(stream.destroyed).to.equal(true));
      expect(fs.existsSync(tracker.created[0])).to.equal(false);
    });

    it('rejects and removes the temp dir on a truncated file', async function() {
      this.timeout(5000);
      const buffer = await bigSheetFirstBuffer();
      const truncated = buffer.subarray(0, Math.floor(buffer.length * 0.9));
      let error;
      try {
        await readAll(chunkedStream(truncated));
      } catch (err) {
        error = err;
      }
      expect(error).to.be.an.instanceOf(Error);
      expect(tracker.created).to.have.length(1);
      expect(fs.existsSync(tracker.created[0])).to.equal(false);
    });

    it('rejects and removes the temp dir on corrupt compressed data', async function() {
      this.timeout(5000);
      const buffer = Buffer.from(await bigSheetFirstBuffer());
      // flip bytes in the middle of the first (deflated) worksheet entry
      for (let i = 2000; i < 2400; i++) {
        buffer[i] ^= 0x55;
      }
      let error;
      try {
        await readAll(chunkedStream(buffer));
      } catch (err) {
        error = err;
      }
      expect(error).to.be.an.instanceOf(Error);
      expect(tracker.created).to.have.length(1);
      expect(fs.existsSync(tracker.created[0])).to.equal(false);
    });
  });

  it('rejects when the input file does not exist', async () => {
    let error;
    try {
      await readAll('./spec/out/does-not-exist.xlsx');
    } catch (err) {
      error = err;
    }
    expect(error).to.be.an.instanceOf(Error);
    expect(error.code).to.equal('ENOENT');
  });
});
