const {Readable} = require('stream');

const ExcelJS = verquire('exceljs');

// XLS-214 — the streaming WorkbookReader dropped the cached ERROR token on a FORMULA cell.
//
// A cell like `<c t="e"><f>1/0</f><v>#DIV/0!</v></c>` carries a formula AND a cached error
// result. The streaming reader's formula branch only special-cased t="str"; everything else
// (including t="e") fell through to `parseFloat(c.v.text)`, so `parseFloat('#DIV/0!')` → NaN and
// the token was LOST. The bare-error `case 'e'` lower down only covers non-formula error cells.
//
// The oracle is the non-streaming `Workbook.load()`, which returns `result:{error:'#DIV/0!'}` for
// the same cell. This test writes such a cell with the (non-streaming) writer, reads it back BOTH
// ways, and asserts the streaming reader now matches the oracle. Revert the fork fix (the t="e"
// case in worksheet-reader.js) and the streaming assertion goes RED (result becomes NaN).
describe('github issues: XLS-214 streaming reader preserves a formula cell cached error token', () => {
  let buffer;
  let oracleValue;
  let streamedValue;

  before(async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = {formula: '1/0', result: {error: '#DIV/0!'}};
    buffer = await wb.xlsx.writeBuffer();

    // Oracle: the non-streaming loader.
    const oracleWb = new ExcelJS.Workbook();
    await oracleWb.xlsx.load(buffer);
    oracleValue = oracleWb.getWorksheet('S').getCell('A1').value;

    // Subject: the streaming WorkbookReader path.
    await new Promise((resolve, reject) => {
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(Readable.from(buffer), {
        worksheets: 'emit',
        sharedStrings: 'cache',
        styles: 'cache',
        hyperlinks: 'ignore',
        entries: 'ignore',
      });
      reader.on('worksheet', worksheet =>
        worksheet.on('row', row => {
          if (row.number === 1) {
            streamedValue = row.getCell(1).value;
          }
        })
      );
      reader.on('end', resolve);
      reader.on('error', reject);
      reader.read();
    });
  });

  it('the non-streaming oracle returns result:{error} (the target shape)', () => {
    expect(oracleValue).to.deep.equal({formula: '1/0', result: {error: '#DIV/0!'}});
  });

  it('the streaming reader preserves the error token (not NaN) — matches the oracle', () => {
    expect(streamedValue).to.deep.equal(oracleValue);
    expect(streamedValue.result).to.deep.equal({error: '#DIV/0!'});
  });
});
