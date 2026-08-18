const fs = require('fs');

const testXformHelper = require('../test-xform-helper');

const WorkbookXform = verquire('xlsx/xform/book/workbook-xform');

const expectations = [
  {
    title: 'book.1',
    create() {
      return new WorkbookXform();
    },
    preparedModel: require('./data/book.1.1.json'),
    xml: fs.readFileSync(`${__dirname}/data/book.1.2.xml`).toString().replace(/\r\n/g, '\n'),
    parsedModel: require('./data/book.1.3.json'),
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'book.2 - no properties',
    create() {
      return new WorkbookXform();
    },
    xml: fs.readFileSync(`${__dirname}/data/book.2.2.xml`).toString().replace(/\r\n/g, '\n'),
    parsedModel: require('./data/book.2.3.json'),
    tests: ['parse'],
  },
];

describe('WorkbookXform', () => {
  testXformHelper(expectations);

  // XLS-861: reconcile() must not crash on an `_xlnm.Print_Area` defined name
  // that carries no ref (empty `ranges`). Excel and third-party writers leave
  // this artifact behind when a print area is cleared; the streaming reader
  // already tolerates it, but the full loader (workbook.xlsx.load) used to throw
  // `TypeError: Cannot read properties of undefined (reading 'match')` in
  // colCache.decodeEx(undefined). Reconcile should treat it as "no print area".
  describe('reconcile print areas', () => {
    function modelWith(definedNames) {
      const worksheet = {};
      return {
        model: {
          workbookRels: [{Id: 'rId1', Target: 'worksheets/sheet1.xml'}],
          sheets: [{rId: 'rId1', name: 'Sheet1', id: 1, state: 'visible'}],
          worksheetHash: {'xl/worksheets/sheet1.xml': worksheet},
          definedNames,
          media: [],
        },
        worksheet,
      };
    }

    it('tolerates an empty _xlnm.Print_Area (no ref) instead of throwing', () => {
      const xform = new WorkbookXform();
      const {model, worksheet} = modelWith([
        {name: '_xlnm.Print_Area', localSheetId: 0, ranges: []},
      ]);
      expect(() => xform.reconcile(model)).to.not.throw();
      // Treated as "no print area": pageSetup.printArea is never set.
      expect(worksheet.pageSetup && worksheet.pageSetup.printArea).to.be.undefined();
    });

    it('still reconciles a valid _xlnm.Print_Area into pageSetup.printArea', () => {
      const xform = new WorkbookXform();
      const {model, worksheet} = modelWith([
        {name: '_xlnm.Print_Area', localSheetId: 0, ranges: ['Sheet1!$A$1:$B$2']},
      ]);
      expect(() => xform.reconcile(model)).to.not.throw();
      expect(worksheet.pageSetup.printArea).to.equal('A1:B2');
    });
  });
});
