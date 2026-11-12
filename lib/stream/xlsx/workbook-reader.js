const fs = require('fs');
const {EventEmitter} = require('events');
const {Readable, pipeline} = require('stream');
const {pipeline: pipelinePromise} = require('stream/promises');
const unzip = require('unzipper');
const os = require('os');
const path = require('path');
const iterateStream = require('../../utils/iterate-stream');
const parseSax = require('../../utils/parse-sax');

const StyleManager = require('../../xlsx/xform/style/styles-xform');
const WorkbookXform = require('../../xlsx/xform/book/workbook-xform');
const RelationshipsXform = require('../../xlsx/xform/core/relationships-xform');

const WorksheetReader = require('./worksheet-reader');
const HyperlinkReader = require('./hyperlink-reader');

// workbook.xml.rels targets are relative to xl/ or absolute from the package root (openpyxl)
function workbookPartPath(target) {
  if (typeof target !== 'string') {
    return undefined;
  }
  return target.startsWith('/') ? target.slice(1) : `xl/${target}`;
}

class WorkbookReader extends EventEmitter {
  constructor(input, options = {}) {
    super();

    this.input = input;

    this.options = {
      worksheets: 'emit',
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      styles: 'ignore',
      entries: 'ignore',
      ...options,
    };

    this.styles = new StyleManager();
    this.styles.init();
  }

  _getStream(input) {
    if (input instanceof Readable) {
      return input;
    }
    if (typeof input === 'string') {
      return fs.createReadStream(input);
    }
    throw new Error(`Could not recognise input: ${input}`);
  }

  async read(input, options) {
    try {
      for await (const {eventType, value} of this.parse(input, options)) {
        switch (eventType) {
          case 'shared-strings':
            this.emit(eventType, value);
            break;
          case 'worksheet':
            this.emit(eventType, value);
            await value.read();
            break;
          case 'hyperlinks':
            this.emit(eventType, value);
            break;
        }
      }
      this.emit('end');
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator]() {
    for await (const {eventType, value} of this.parse()) {
      if (eventType === 'worksheet') {
        yield value;
      }
    }
  }

  async *parse(input, options) {
    if (options) this.options = options;
    const stream = (this.stream = this._getStream(input || this.input));
    const zip = unzip.Parse({forceStream: true});
    // pipeline (unlike pipe) forwards input errors to zip, which rejects the iteration below,
    // and destroys the input when iteration stops early (zip is then destroyed by for await)
    pipeline(stream, zip, () => {});

    // worksheets, deferred for parsing after shared strings reading
    const waitingWorkSheets = [];
    let tmpDir = null;
    try {
      for await (const entry of zip) {
        let match;
        let sheetNo;
        switch (entry.path) {
          case '_rels/.rels':
            break;
          case 'xl/_rels/workbook.xml.rels':
            await this._parseRels(entry);
            break;
          case 'xl/workbook.xml':
            await this._parseWorkbook(entry);
            break;
          case 'xl/sharedStrings.xml':
            yield* this._parseSharedStrings(entry);
            break;
          case 'xl/styles.xml':
            await this._parseStyles(entry);
            break;
          default:
            if (entry.path.match(/xl\/worksheets\/sheet\d+[.]xml/)) {
              match = entry.path.match(/xl\/worksheets\/sheet(\d+)[.]xml/);
              sheetNo = match[1];
              if (this._worksheetDependenciesParsed()) {
                yield* this._parseWorksheet(iterateStream(entry), sheetNo);
              } else {
                // defer worksheet until the parts it depends on are known: spool to a temp file
                if (!tmpDir) {
                  tmpDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'exceljs-'));
                }
                const tmpPath = path.join(tmpDir, `${waitingWorkSheets.length}.xml`);
                waitingWorkSheets.push({sheetNo, path: tmpPath});
                // rejects on entry errors too (truncated or corrupt zip data)
                await pipelinePromise(entry, fs.createWriteStream(tmpPath));
              }
            } else if (entry.path.match(/xl\/worksheets\/_rels\/sheet\d+[.]xml.rels/)) {
              match = entry.path.match(/xl\/worksheets\/_rels\/sheet(\d+)[.]xml.rels/);
              sheetNo = match[1];
              yield* this._parseHyperlinks(iterateStream(entry), sheetNo);
            }
            break;
        }
        entry.autodrain();
      }

      for (const {sheetNo, path: tmpPath} of waitingWorkSheets) {
        const sheetStream = fs.createReadStream(tmpPath);
        try {
          yield* this._parseWorksheet(sheetStream, sheetNo);
        } finally {
          // the consumer may skip (or stop reading) a worksheet: release its file handle
          sheetStream.destroy();
        }
      }
    } finally {
      // also runs when the consumer stops iterating early (break / throw)
      if (tmpDir) {
        await fs.promises.rm(tmpDir, {recursive: true, force: true});
      }
    }
  }

  // shared strings, sheet names (workbook rels + workbook) and, when cached, styles
  _worksheetDependenciesParsed() {
    const stylesReady = this.options.styles !== 'cache' || this._stylesParsed;
    return Boolean(this.sharedStrings && this.workbookRels && this.model && stylesReady);
  }

  _emitEntry(payload) {
    if (this.options.entries === 'emit') {
      this.emit('entry', payload);
    }
  }

  async _parseRels(entry) {
    const xform = new RelationshipsXform();
    this.workbookRels = await xform.parseStream(iterateStream(entry));
  }

  async _parseWorkbook(entry) {
    this._emitEntry({type: 'workbook'});

    const workbook = new WorkbookXform();
    await workbook.parseStream(iterateStream(entry));

    this.properties = workbook.map.workbookPr;
    this.model = workbook.model;
  }

  async *_parseSharedStrings(entry) {
    this._emitEntry({type: 'shared-strings'});
    switch (this.options.sharedStrings) {
      case 'cache':
        this.sharedStrings = [];
        break;
      case 'emit':
        break;
      default:
        return;
    }

    let text = null;
    let richText = [];
    let index = 0;
    let font = null;
    for await (const events of parseSax(iterateStream(entry))) {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          const node = value;
          switch (node.name) {
            case 'b':
              font = font || {};
              font.bold = true;
              break;
            case 'charset':
              font = font || {};
              font.charset = parseInt(node.attributes.charset, 10);
              break;
            case 'color':
              font = font || {};
              font.color = {};
              if (node.attributes.rgb) {
                font.color.argb = node.attributes.argb;
              }
              if (node.attributes.val) {
                font.color.argb = node.attributes.val;
              }
              if (node.attributes.theme) {
                font.color.theme = node.attributes.theme;
              }
              break;
            case 'family':
              font = font || {};
              font.family = parseInt(node.attributes.val, 10);
              break;
            case 'i':
              font = font || {};
              font.italic = true;
              break;
            case 'outline':
              font = font || {};
              font.outline = true;
              break;
            case 'rFont':
              font = font || {};
              font.name = node.value;
              break;
            case 'si':
              font = null;
              richText = [];
              text = null;
              break;
            case 'sz':
              font = font || {};
              font.size = parseInt(node.attributes.val, 10);
              break;
            case 'strike':
              break;
            case 't':
              text = null;
              break;
            case 'u':
              font = font || {};
              font.underline = true;
              break;
            case 'vertAlign':
              font = font || {};
              font.vertAlign = node.attributes.val;
              break;
          }
        } else if (eventType === 'text') {
          text = text ? text + value : value;
        } else if (eventType === 'closetag') {
          const node = value;
          switch (node.name) {
            case 'r':
              richText.push({
                font,
                text,
              });

              font = null;
              text = null;
              break;
            case 'si':
              if (this.options.sharedStrings === 'cache') {
                this.sharedStrings.push(richText.length ? {richText} : text);
              } else if (this.options.sharedStrings === 'emit') {
                yield {index: index++, text: richText.length ? {richText} : text};
              }

              richText = [];
              font = null;
              text = null;
              break;
          }
        }
      }
    }
  }

  async _parseStyles(entry) {
    this._emitEntry({type: 'styles'});
    if (this.options.styles === 'cache') {
      this.styles = new StyleManager();
      await this.styles.parseStream(iterateStream(entry));
      this._stylesParsed = true;
    }
  }

  *_parseWorksheet(iterator, sheetNo) {
    this._emitEntry({type: 'worksheet', id: sheetNo});
    const worksheetReader = new WorksheetReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });

    const sheetPath = `xl/worksheets/sheet${sheetNo}.xml`;
    const matchingRel = (this.workbookRels || []).find(rel => workbookPartPath(rel.Target) === sheetPath);
    const sheets = (this.model && this.model.sheets) || [];
    const matchingSheet = matchingRel && sheets.find(sheet => sheet.rId === matchingRel.Id);
    if (matchingSheet) {
      worksheetReader.id = matchingSheet.id;
      worksheetReader.name = matchingSheet.name;
      worksheetReader.state = matchingSheet.state;
    }
    if (this.options.worksheets === 'emit') {
      yield {eventType: 'worksheet', value: worksheetReader};
    }
  }

  *_parseHyperlinks(iterator, sheetNo) {
    this._emitEntry({type: 'hyperlinks', id: sheetNo});
    const hyperlinksReader = new HyperlinkReader({
      workbook: this,
      id: sheetNo,
      iterator,
      options: this.options,
    });
    if (this.options.hyperlinks === 'emit') {
      yield {eventType: 'hyperlinks', value: hyperlinksReader};
    }
  }
}

// for reference - these are the valid values for options
WorkbookReader.Options = {
  worksheets: ['emit', 'ignore'],
  sharedStrings: ['cache', 'emit', 'ignore'],
  hyperlinks: ['cache', 'emit', 'ignore'],
  styles: ['cache', 'ignore'],
  entries: ['emit', 'ignore'],
};

module.exports = WorkbookReader;
