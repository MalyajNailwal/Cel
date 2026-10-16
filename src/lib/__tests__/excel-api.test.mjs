import './fake-office.mjs';
import { getCtx } from './fake-office.mjs';
// excel-api.ts is bundled to a sibling api.mjs before this suite runs:
//   esbuild src/lib/excel-api.ts --bundle --format=esm --platform=browser \
//     --outfile=src/lib/__tests__/api.mjs
import * as API from './api.mjs';

const DATA = [
  ['Region','Product','Qty','Sales'],
  ['North','Widget A',10,1200],
  ['South','Widget B',4,800],
  ['North','Gadget X',7,2400],
  ['East','Gadget Y',2,600],
];

let ctx;
async function ensureCtx() {
  if (!getCtx()) await API.getWorkbookStructure();
  ctx = getCtx();
  return ctx;
}

async function seed() {
  await ensureCtx();
  const s = getCtx().sheets.Sheet1.sheet;
  for (const k of Object.keys(s.grid)) delete s.grid[k];
  for (let r=0;r<DATA.length;r++)
    for (let c=0;c<DATA[r].length;c++)
      s.grid[`${String.fromCharCode(65+c)}${r+1}`] = DATA[r][c];
  s.calls.length = 0;
}

// Every API call spins a fresh context; always assert against the newest one.
function sheet1() { return getCtx().sheets.Sheet1.sheet; }

const tests = [];
const t = (name, fn) => tests.push({name, fn});

t('getWorkbookStructure', async () => {
  const r = await API.getWorkbookStructure();
  if (!r.sheets.find(s=>s.name==='Sheet1')) throw new Error('Sheet1 missing');
});

t('getSelectedRange', async () => {
  const r = await API.getSelectedRange();
  if (!r) throw new Error('null');
});

t('getRange', async () => {
  const r = await API.getRange('A1:D5','Sheet1');
  if (r.values.length !== 5) throw new Error('bad rows');
});

t('getSheetData', async () => {
  const r = await API.getSheetData('Sheet1', 3);
  if (r.values.length !== 3) throw new Error('bad rows: '+r.values.length);
});

t('setValues', async () => {
  await API.setValues('A1', [['H1','H2']], 'Sheet1');
  const ctx = getCtx();
  if (getCtx().sheets.Sheet1.sheet.grid.A1 !== 'H1') throw new Error('not written');
});

t('setFormulas', async () => {
  await API.setFormulas('A1', [['=SUM(1,2)']], 'Sheet1');
});

t('apply_format incl borders', async () => {
  await API.applyFormat({ address:'A1:D1', sheetName:'Sheet1', bold:true, fillColor:'#FFFF00',
    borders:{ all:true, color:'#000000', style:'Continuous', weight:'Thin' } });
});

t('insertRows / deleteRows', async () => {
  await API.insertRows('A2', 1, 'Sheet1');
  await API.deleteRows('A2', 1, 'Sheet1');
});

t('insertColumns / deleteColumns', async () => {
  await API.insertColumns('A1', 1, 'Sheet1');
  await API.deleteColumns('A1', 1, 'Sheet1');
});

t('addWorksheet', async () => {
  const n = await API.addWorksheet('TestSheet');
  if (!n) throw new Error('no name');
});

t('createTable', async () => {
  await API.createTable('A1:D5','MyTable','Sheet1');
});

t('sortRange', async () => {
  await API.sortRange('A1:D5', 0, true, 'Sheet1');
});

t('autoFill', async () => {
  await API.autoFill('A1:A5','E1:E5','Sheet1');
});

function STORE_CALLS(){ return (globalThis.__storeCalls || []); }

t('createChart + options', async () => {
  const r = await API.createChart('columnStacked100','A1:D5','Sheet1','Sales by Region',
    {left:10,top:10,width:400,height:300},
    { showLegend:true, legendPosition:'Right', showDataLabels:true,
      xAxisTitle:'Region', yAxisTitle:'Sales', gapWidth:50, overlap:100, seriesColor:'#217346' });
  const calls = STORE_CALLS();
  const joined = STORE_CALLS().join(' | ');
  for (const want of ['chart:columnStacked100','title:Sales by Region','xAxis:Region','yAxis:Sales','seriesColor:#217346','gap:50','overlap:100'])
    if (!joined.includes(want)) throw new Error(`missing ${want} in ${joined}`);
});

t('createChartFromTwoColumns', async () => {
  await API.createChartFromTwoColumns('pie','A1:A5','D1:D5','Sheet1','Share',
    undefined, { showDataLabels:true });
});

t('conditional_format cellValue', async () => {
  await API.conditionalFormat('D2:D5', { type:'cellValue', operator:'GreaterThan', formula1:'1000',
    format:{ fillColor:'#FFFF00' } }, 'Sheet1');
});

t('conditional_format colorScale', async () => {
  await API.conditionalFormat('D2:D5', { type:'colorScale', colors:['#F8696B','#FFEB84','#63BE7B'] }, 'Sheet1');
});

t('conditional_format dataBar', async () => {
  await API.conditionalFormat('D2:D5', { type:'dataBar' }, 'Sheet1');
});

t('conditional_format duplicateValues', async () => {
  await API.conditionalFormat('A2:A5', { type:'duplicateValues' }, 'Sheet1');
});

t('conditional_format topBottom', async () => {
  await API.conditionalFormat('D2:D5', { type:'topBottom', formula1:'3', operator:'top' }, 'Sheet1');
});

t('conditional_format rejects unknown type', async () => {
  let threw=false;
  try { await API.conditionalFormat('A1', { type:'nonsense' }, 'Sheet1'); } catch { threw=true; }
  if (!threw) throw new Error('should have thrown');
});

t('find_replace', async () => {
  const n = await API.findAndReplace('A1:D5','North','NORTH',false,'Sheet1');
  if (typeof n !== 'number') throw new Error('not a number');
});

t('mergeCells / unmergeCells', async () => {
  await API.mergeCells('A1:D1','Sheet1');
  await API.unmergeCells('A1:D1','Sheet1');
});

t('addDropdown', async () => {
  await API.addDropdown('B2:B5',['Yes','No'],'Sheet1');
});

t('copy_range', async () => {
  await API.copyRange('A1:D5','A1','Sheet1','Sheet2');
});

t('auto_fit_columns', async () => {
  await API.autoFitColumns(undefined,'Sheet1');
});

t('set_row_height', async () => {
  await API.setRowHeight('1:1', 24, 'Sheet1');
});

t('set_column_width', async () => {
  await API.setColumnWidth('A:A', 120, 'Sheet1');
});

t('freeze_panes', async () => {
  await API.freezePanes(1, 0, 'Sheet1');
});

t('add_comment', async () => {
  await API.addComment('A1','check this','Sheet1','Analyst');
});

t('add_hyperlink', async () => {
  await API.addHyperlink('A1','https://example.com','Example','Sheet1');
});

// textToColumns writes the split array onto a range anchored at the source
// cell, so G1 keeps field 1 and H1 gets field 2 (overwriting whatever was there).
t('text_to_columns writes first field back and second alongside', async () => {
  sheet1().grid.G1 = 'Ada Lovelace,1815';
  await API.textToColumns('G1', ',', 'Sheet1');
  const s = sheet1();
  if (s.grid.G1 !== 'Ada Lovelace') throw new Error(`G1=${JSON.stringify(s.grid.G1)}`);
  if (s.grid.H1 !== '1815') throw new Error(`H1=${JSON.stringify(s.grid.H1)}`);
});

t('text_to_columns 3 fields', async () => {
  sheet1().grid.G1 = 'Smith;John;42';
  await API.textToColumns('G1', ';', 'Sheet1');
  const s = sheet1();
  if (s.grid.G1 !== 'Smith' || s.grid.H1 !== 'John' || s.grid.I1 !== '42') {
    throw new Error(`G1=${s.grid.G1} H1=${s.grid.H1} I1=${s.grid.I1}`);
  }
});

t('text_to_columns pads short rows with null', async () => {
  sheet1().grid.G1 = 'a,b,c';
  sheet1().grid.G2 = 'd';
  await API.textToColumns('G1:G2', ',', 'Sheet1');
  const s = sheet1();
  if (s.grid.H2 !== null || s.grid.I2 !== null) throw new Error(`row2 not padded: H2=${s.grid.H2} I2=${s.grid.I2}`);
});

t('remove_duplicates', async () => {
  const rows = [['Dup','X',1,1],['Dup','X',1,1],['Dup','X',1,1]];
  for (let r=0;r<3;r++)
    for (let c=0;c<4;c++)
      sheet1().grid[`${String.fromCharCode(74+c)}${r+1}`] = rows[r][c];
  const removed = await API.removeDuplicates('J1:M3', 'Sheet1');
  if (removed !== 2) throw new Error('expected 2 removed, got '+removed);
  const s = sheet1();
  if (s.grid.J3 !== null) throw new Error('duplicate row not cleared: J3='+s.grid.J3);
});

t('create_pivot_table', async () => {
  // seed() writes to Sheet1.sheet.grid but pivot reads via Range.values, which
  // reads the same grid. Verify what pivot actually sees.
  const ctx = getCtx();
  const probe = ctx.sheets.Sheet1.getRange('A1:D1');
  const seen = probe.values;
  if (seen[0][0] !== 'Region') throw new Error('harness seed broken: '+JSON.stringify(seen));
  const out = await API.createPivotTable('A1:D5','H1',['Region'],['Sales'],undefined,'Sheet1');
  if (!out.includes('Pivot table')) throw new Error('bad output: '+out);
  if (ctx.sheets.Sheet1.sheet.grid.H1 !== 'Region') throw new Error('header not written: '+JSON.stringify(ctx.sheets.Sheet1.sheet.grid.H1));
  // North: 1200+2400=3600, South: 800, East: 600
  const sums = [ctx.sheets.Sheet1.sheet.grid.I1, ctx.sheets.Sheet1.sheet.grid.I2, ctx.sheets.Sheet1.sheet.grid.I3];
  if (sums.includes(undefined)) throw new Error('sums missing: '+JSON.stringify(sums));
});

t('create_pivot_table with column field', async () => {
  const out = await API.createPivotTable('A1:D5','N1',['Region'],['Sales'],['Product'],'Sheet1');
  if (!out.includes('Pivot table')) throw new Error('bad output: '+out);
});

t('create_pivot_table rejects unknown field', async () => {
  let threw=false;
  try { await API.createPivotTable('A1:D5','T1',['Nope'],['Sales'],undefined,'Sheet1'); } catch { threw=true; }
  if (!threw) throw new Error('should reject unknown field');
});

t('setValues to missing sheet fails clearly', async () => {
  let msg='';
  try { await API.setValues('A1',[['x']],'GhostSheet'); } catch(e){ msg=String(e.message||e) }
  if (!msg.includes('not found')) throw new Error('bad error: '+msg);
});

let pass=0, fail=0;
for (const {name, fn} of tests) {
  await seed();
  try { await fn(); console.log(`ok   ${name}`); pass++; }
  catch (e) { console.log(`FAIL ${name}: ${e.message}`); fail++; }
}
console.log(`\n${pass}/${tests.length} passed, ${fail} failed`);
process.exit(fail?1:0);
