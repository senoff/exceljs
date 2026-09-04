const fs = require('fs');
const os = require('os');
const path = require('path');
const JSZip = require('jszip');

const ExcelJS = verquire('exceljs');

// XLS-352 — teach the fork to pin every zip entry's mtime itself, so the server's
// `PinnedWorkbookWriter` subclass (which reached past the published surface to wrap `zip.append`
// from inside the base constructor) can be deleted.
//
// The defect: `date` is a PER-ENTRY zip option (archiver's `data.date` / JSZip's file `date`),
// defaulting to `new Date()` at append time. A module-level `zip: {date}` is a no-op — archiver
// never forwards it to an entry header. So each entry's local-file-header mtime carries the wall
// clock at the DOS format's two-second granularity, and two builds are byte-identical only when
// they happen to land in the same two-second bucket (XLS-350's coin-flip).
//
// The fix adds one option, `zipEntryDate`, threaded to EVERY `zip.append` on BOTH writer paths:
//   - streaming: `stream.xlsx.WorkbookWriter` — via the `_append` choke point.
//   - buffered:  `Workbook.xlsx.writeBuffer/writeFile` — via `ZipWriter.append` (JSZip `.file`).
//
// This is pinned three ways per path, because each alone is a false green:
//   1. Byte-identity across a real DOS tick — the determinism property itself.
//   2. Every entry's parsed mtime equals the pin — proves the date reached the entry headers, not
//      just that two runs happened to match.
//   3. POSITIVE CONTROL: WITHOUT the option the output differs from the pinned output AND the
//      entries carry the wall clock (a recent year), so the option is demonstrably what pins them.
//      Revert the fork fix and arms 1+2 stay green vacuously; this arm is what reddens.

// A fixed pin on an even two-second boundary (DOS granularity) and at midday, so a CI timezone
// offset shifts the hour but never the calendar date the readback asserts.
const PIN = new Date(2000, 5, 15, 12, 0, 0);

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

let tmpCounter = 0;
function tmpFile() {
  return path.join(os.tmpdir(), `xls352-stream-${process.pid}-${Date.now()}-${tmpCounter++}.xlsx`);
}

// Build a workbook through the STREAMING writer and return its bytes. `extraOptions` carries (or
// omits) `zipEntryDate`. `created`/`modified` are pinned too so docProps/core.xml — entry CONTENT,
// the other half of byte-identity — does not itself carry the wall clock and mask the entry-mtime
// property under test.
async function buildStreaming(extraOptions) {
  const file = tmpFile();
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({
    stream: fs.createWriteStream(file, {flags: 'w'}),
    useStyles: true,
    useSharedStrings: false,
    ...extraOptions,
  });
  wb.created = PIN;
  wb.modified = PIN;
  const ws = wb.addWorksheet('S');
  ws.addRow(['hello', 42]).commit();
  ws.commit();
  await wb.commit();
  const buffer = fs.readFileSync(file);
  fs.unlinkSync(file);
  return buffer;
}

// Build the same workbook through the BUFFERED writer (`writeBuffer`) and return its bytes.
async function buildBuffered(extraOptions) {
  const wb = new ExcelJS.Workbook();
  wb.created = PIN;
  wb.modified = PIN;
  const ws = wb.addWorksheet('S');
  ws.addRow(['hello', 42]);
  return wb.xlsx.writeBuffer(extraOptions);
}

// Every entry, INCLUDING the folder entries JSZip auto-creates — those default to `new Date()` and
// are exactly what made the buffered writer non-deterministic until the finalize-time pin.
async function entryDates(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return Object.values(zip.files).map(entry => entry.date);
}

describe('github issues: XLS-352 fork pins every zip entry mtime (durable fix for XLS-350)', function() {
  // Two full builds plus a real >2s delay per determinism arm.
  this.timeout(30000);

  describe('streaming writer (stream.xlsx.WorkbookWriter)', () => {
    it('two runs separated by more than one DOS tick are byte-identical when zipEntryDate is pinned', async () => {
      const first = await buildStreaming({zipEntryDate: PIN});
      await sleep(2500);
      const second = await buildStreaming({zipEntryDate: PIN});
      expect(first.equals(second)).to.equal(true);
    });

    it('stamps every zip entry mtime with zipEntryDate', async () => {
      const dates = await entryDates(await buildStreaming({zipEntryDate: PIN}));
      expect(dates.length).to.be.greaterThan(3);
      dates.forEach(date => {
        expect(date.getFullYear()).to.equal(2000);
        expect(date.getMonth()).to.equal(5);
        expect(date.getDate()).to.equal(15);
      });
    });

    it('POSITIVE CONTROL: without zipEntryDate the bytes differ and entries carry the wall clock', async () => {
      const pinned = await buildStreaming({zipEntryDate: PIN});
      const unpinned = await buildStreaming({});
      expect(pinned.equals(unpinned)).to.equal(false);
      const dates = await entryDates(unpinned);
      expect(dates.some(date => date.getFullYear() >= 2020)).to.equal(true);
    });
  });

  describe('buffered writer (Workbook.xlsx.writeBuffer)', () => {
    it('two runs separated by more than one DOS tick are byte-identical when zipEntryDate is pinned', async () => {
      const first = await buildBuffered({zipEntryDate: PIN});
      await sleep(2500);
      const second = await buildBuffered({zipEntryDate: PIN});
      expect(first.equals(second)).to.equal(true);
    });

    it('stamps every zip entry mtime with zipEntryDate', async () => {
      const dates = await entryDates(await buildBuffered({zipEntryDate: PIN}));
      expect(dates.length).to.be.greaterThan(3);
      dates.forEach(date => {
        expect(date.getFullYear()).to.equal(2000);
        expect(date.getMonth()).to.equal(5);
        expect(date.getDate()).to.equal(15);
      });
    });

    it('POSITIVE CONTROL: without zipEntryDate the bytes differ and entries carry the wall clock', async () => {
      const pinned = await buildBuffered({zipEntryDate: PIN});
      const unpinned = await buildBuffered({});
      expect(pinned.equals(unpinned)).to.equal(false);
      const dates = await entryDates(unpinned);
      expect(dates.some(date => date.getFullYear() >= 2020)).to.equal(true);
    });
  });
});
