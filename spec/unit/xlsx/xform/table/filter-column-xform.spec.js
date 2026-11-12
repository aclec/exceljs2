const {Readable} = require('stream');
const testXformHelper = require('../test-xform-helper');

const parseSax = verquire('utils/parse-sax');

const FilterColumnXform = verquire('xlsx/xform/table/filter-column-xform');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: true},
    preparedModel: {colId: '0', filterButton: true},
    xml: '<filterColumn colId="0" hiddenButton="0" />',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 0},
  },
  {
    title: 'hidden filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: false},
    preparedModel: {colId: '1', filterButton: false},
    xml: '<filterColumn colId="1" hiddenButton="1" />',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 1},
  },
  {
    title: 'with custom filter',
    create() {
      return new FilterColumnXform();
    },
    initialModel: {filterButton: false, customFilters: [{val: '*brandywine*'}]},
    preparedModel: {
      colId: '0',
      filterButton: false,
      customFilters: [{val: '*brandywine*'}],
    },
    xml:
      '<filterColumn colId="0" hiddenButton="1"><customFilters><customFilter val="*brandywine*"/></customFilters></filterColumn>',
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['prepare', 'render', 'renderIn', 'parse'],
    options: {index: 0},
  },
];

describe('FilterColumnXform', () => {
  testXformHelper(expectations);

  it('does not carry custom filters over to the next column', async () => {
    const xform = new FilterColumnXform();
    const parse = xml => xform.parse(parseSax(Readable.from([xml])));
    await parse(
      '<filterColumn colId="0" hiddenButton="0"><customFilters><customFilter val="a"/></customFilters></filterColumn>'
    );
    expect(await parse('<filterColumn colId="1" hiddenButton="0"/>')).to.deep.equal({
      filterButton: true,
    });
  });

  it('skips the whole subtree of an element it does not model', async () => {
    const xform = new FilterColumnXform();
    const model = await xform.parse(
      parseSax(
        Readable.from([
          '<filterColumn colId="0" hiddenButton="0"><extLst><ext uri="{0}">' +
            '<customFilters><customFilter val="no"/></customFilters></ext></extLst></filterColumn>',
        ])
      )
    );
    expect(model).to.deep.equal({filterButton: true});
  });
});
