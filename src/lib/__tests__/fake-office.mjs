// Minimal Office.js fake: enough to drive every excel-api.ts action.
export class FakeExcel {
  static async run(fn) { return fn(new Ctx()); }
}
class Range {
  constructor(sheet, addr){ this.sheet=sheet; this.addr=addr; }
  load(){return this} sync(){return this}
  get isNullObject(){ return this.addr===null }
  get rowCount(){ const {s,e}=split(this.addr); return e.row-s.row+1 }
  get columnCount(){ const {s,e}=split(this.addr); return ci(e.col)-ci(s.col)+1 }
  get values(){
    if (globalThis.__trace) console.log(`  READ values ${this.addr} ->`, JSON.stringify(this.sheet._values(this.addr)).slice(0,90));
    return this.sheet._values(this.addr);
  }
  set values(v){
    if (globalThis.__trace) console.log(`  WRITE values ${this.addr} <-`, JSON.stringify(v).slice(0,90), 'caller:', (new Error().stack.split('\n')[2]||'').trim().slice(0,70));
    this.sheet._set(this.addr, v);
  }
  get formulas(){ return this.sheet._values(this.addr) }
  set formulas(v){ this.sheet._set(this.addr, v) }
  get isNullObject(){ return this.addr === null || this.addr === undefined }
  get numberFormat(){ return this._nf ?? this.sheet._nf }
  set numberFormat(v){ this._nf=v }
  getResizedRange(dr,dc){ const {s,e}=split(this.addr); return new Range(this.sheet, `${s.col}${s.row}:${indexToCol(ci(s.col)+dc)}${e.row+dr}`) }
  getRow(i){ const {s,e}=split(this.addr); return new Range(this.sheet, `${s.col}${i+1}:${e.col}${i+1}`) }
  getEntireRow(){ const {s,e}=split(this.addr); return new Range(this.sheet, `${s.col}1:${e.col}${e.row}`) }
  getEntireColumn(){ const {s,e}=split(this.addr); return new Range(this.sheet, `${s.col}1:${e.col}1048576`) }
  getColumn(i){ const {s,e}=split(this.addr); return new Range(this.sheet, `${indexToCol(ci(s.col)+i)}${s.row}:${indexToCol(ci(s.col)+i)}${e.row}`) }
  getRowsAbove(n){ const {s,e}=split(this.addr); return new Range(this.sheet, `${s.col}${Math.max(1,s.row-n)}:${e.col}${Math.max(1,s.row-1)}`) }
  getColumnsBefore(n){ const {s,e}=split(this.addr); return new Range(this.sheet, `${indexToCol(Math.max(0,ci(s.col)-n))}1:${indexToCol(Math.max(0,ci(s.col)-1))}${e.row}`) }
  insert(shift, count){ const n=count||1;
    this.sheet.calls.push(`insert:${shift}:${this.addr}:${n}`);
    this.sheet.shift(this.addr, shift==='Down'||shift==='Right' ? 1 : -1, n, shift==='Down'?'row':shift==='Right'?'col':'row');
  }
  delete(shift, count){ const n=count||1;
    this.sheet.calls.push(`delete:${shift}:${this.addr}:${n}`);
    this.sheet.shift(this.addr, shift==='Up'||shift==='Left' ? -1 : 1, n, (shift==='Up'||shift==='Down')?'row':'col');
  }
  merge(){ this.sheet.calls.push(`merge:${this.addr}`) }
  unmerge(){ this.sheet.calls.push(`unmerge:${this.addr}`) }
  select(){ this.sheet.calls.push(`select:${this.addr}`) }
  clear(){ this.sheet.calls.push(`clear:${this.addr}`) }
  copyFrom(src, t){ this.sheet.calls.push(`copyFrom:${src.addr||'null'}->${this.addr}:${t}`) }
  autoFill(t){ this.sheet.calls.push(`autoFill:${this.addr}->${t.addr}`) }
  get format(){ const s=this.sheet; return new Proxy({}, { get:(_,p)=>{
      if(p==='font') return new Proxy({}, {get:(_,q)=>{
        if(q==='bold'||q==='italic') return undefined
        return { setSolidColor(c){s.calls.push(`fontColor:${c}`)} }
      }});
      if(p==='fill') return { get color(){return undefined}, set color(c){s.calls.push(`fill:${c}`)} };
      if(p==='borders') return { getItem:()=>({ set style(v){}, set color(v){}, set weight(v){} }) };
      if(p==='autofitColumns'||p==='autofitRows') return ()=>s.calls.push('autofit');
      if(p==='rowHeight'||p==='columnWidth') return undefined;
      return undefined;
    }, set:(_,p,v)=>{ s.calls.push(`${p}:${JSON.stringify(v)}`); return true } }) }
  get numberFormat(){ return undefined }
  get dataValidation(){ const s=this.sheet; return { set rule(v){ s.calls.push(`validation:${JSON.stringify(v)}`) } } }
  get conditionalFormats(){ const s=this.sheet; return { add(t){
    s.calls.push(`cf:${t}`);
    const makeFill = () => ({ set color(c){ s.calls.push(`cfFill:${c}`) } });
    const makeFont = () => ({ set color(c){ s.calls.push(`cfFont:${c}`) }, set bold(b){ s.calls.push(`cfBold:${b}`) }, set italic(v){ s.calls.push(`cfItalic:${v}`) }, set name(v){ s.calls.push(`cfName:${v}`) }, set size(v){ s.calls.push(`cfSize:${v}`) } });
    const fmt = {
      get fill(){ return makeFill() },
      get font(){ return makeFont() },
      set font(v){ s.calls.push(`cfFont:${JSON.stringify(v)}`) },
      set bold(v){ s.calls.push(`cfBold:${v}`) },
      set color(v){ s.calls.push(`cfColor:${v}`) },
      set italic(v){ s.calls.push(`cfItalic:${v}`) }, set name(v){ s.calls.push(`cfName:${v}`) }, set size(v){ s.calls.push(`cfSize:${v}`) },
    };
    const o = {};
    o.cellValue = { format: fmt, set rule(v){ s.calls.push(`cfRule:${JSON.stringify(v)}`) } };
    o.topBottom = { format: fmt, set rule(v){ s.calls.push(`topBottom:${JSON.stringify(v)}`) } };
    o.presetCriteria = { format: fmt, set rule(v){ s.calls.push(`cfRule:${v.criterion}`) } };
    o.set = v => s.calls.push(`cfSet:${JSON.stringify(v).slice(0,90)}`);
    return o; } } }
  get hyperlink(){ return undefined }
  set hyperlink(v){ this.sheet.calls.push(`hyperlink:${v.address}`) }
  get text(){ return undefined }
  addNote(t){ this.sheet.calls.push(`note:${t}`) }
  get sort(){ return { apply:()=>this.sheet.calls.push(`sort:${this.addr}`) } }
  get tables(){ return { add:(a,n)=>this.sheet.calls.push(`table:${a}:${n}`) } }
  insert(){ this.sheet.calls.push(`insertRows:${this.addr}`) }
  delete(){ this.sheet.calls.push(`deleteRows:${this.addr}`) }
  getEntireRange(){ return this }
  get name(){ return this.sheet.name }
  get worksheet(){ return this.sheet.ws }
}
function split(a){
  // whole row  1:5   -> A..XFD, rows 1..5
  let m = a.match(/^(\d+):(\d+)$/);
  if (m) return { s:{col:'A', row:+m[1]}, e:{col:'XFD', row:+m[2]} };
  // whole column A:A -> rows 1..1048576
  m = a.match(/^([A-Z]+):([A-Z]+)$/);
  if (m) return { s:{col:m[1], row:1}, e:{col:m[2], row:1048576} };
  // A1:C4, A1 (single cell)
  m = a.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (m) return { s:{col:m[1], row:+m[2]}, e:{col:m[3]||m[1], row:m[4]?+m[4]:+m[2]} };
  // A1:7  /  A1:7 already handled; A1:E1 handled above.
  // "A1:D"   -> row implied by left cell
  m = a.match(/^([A-Z]+)(\d+):([A-Z]+)$/);
  if (m) return { s:{col:m[1], row:+m[2]}, e:{col:m[3], row:+m[2]} };
  // "A1:4"   -> col implied by left cell
  m = a.match(/^([A-Z]+)(\d+):(\d+)$/);
  if (m) return { s:{col:m[1], row:+m[2]}, e:{col:m[1], row:+m[3]} };
  // "A:4"    -> whole column, row-bounded
  m = a.match(/^([A-Z]+):(\d+)$/);
  if (m) return { s:{col:m[1], row:1}, e:{col:m[1], row:+m[2]} };
  // "4:E"    -> whole row, col-bounded
  m = a.match(/^(\d+):([A-Z]+)$/);
  if (m) return { s:{col:'A', row:+m[1]}, e:{col:m[2], row:+m[1]} };
  // "A"      -> whole column
  m = a.match(/^([A-Z]+)$/);
  if (m) return { s:{col:m[1], row:1}, e:{col:m[1], row:1048576} };
  // "4"      -> whole row
  m = a.match(/^(\d+)$/);
  if (m) return { s:{col:'A', row:+m[1]}, e:{col:'XFD', row:+m[1]} };
  throw new Error('bad address: '+a);
}
function ci(l){let n=0;for(const c of l)n=n*26+(c.charCodeAt(0)-64);return n-1}
function indexToCol(i){let n=i+1,o='';while(n>0){const r=(n-1)%26;o=String.fromCharCode(65+r)+o;n=Math.floor((n-1)/26)}return o}
class Sheet {
  constructor(name, ws){
    if (!STORE.grids[name]) STORE.grids[name] = {};
    this.ws=ws; this.name=name; this.grid=STORE.grids[name];
    this.calls=[]; this._nf='General'; this.staged=null;
  }
  _values(addr){ const {s,e}=split(addr); const out=[];
    for(let r=s.row;r<=e.row;r++){const row=[];for(let c=ci(s.col);c<=ci(e.col);c++)row.push(this._read(`${indexToCol(c)}${r}`));out.push(row)}
    return out }
  _set(addr,v){ const {s,e}=split(addr);
    this.staged = this.staged || {};
    for(let r=s.row;r<=e.row;r++){for(let c=ci(s.col);c<=ci(e.col);c++){this.staged[`${indexToCol(c)}${r}`]=v[r-s.row]?.[c-ci(s.col)]??null}}}
  _read(k){ const st=this.staged; if(st && k in st) return st[k]; return this.grid[k]??null }
  getRange(a){ return new Range(this,a) }
  getRangeOrNullObject(){ return new Range(this,'A1:C4') }
  getUsedRangeOrNullObject(){ return new Range(this,'A1:D5') }
  shift(addr, dir, n, axis){ const {s:st,e:en}=split(addr);
    const colKey = k => k.replace(/\d+/,'');
    const rowKey = k => k.replace(/[A-Z]+/,'');
    const keys = Object.keys(this.grid);
    const out = {};
    for (const k of keys) {
      const m = k.match(/^([A-Z]+)(\d+)$/);
      let col = ci(m[1]), row = +m[2];
      if (axis==='row') { if (row >= st.row) row += dir*n; }
      else { if (col >= ci(st.col)) col += dir*n; }
      if (row < 1 || col < 0) continue;
      out[`${indexToCol(col)}${row}`] = this.grid[k];
    }
    // Mutate in place so every Sheet instance keeps the same grid reference.
    for (const k of Object.keys(this.grid)) delete this.grid[k];
    Object.assign(this.grid, out);
    this.staged = null;
  }
  get charts(){ const s=this; return { add:(t,r,by)=>{ s.calls.push(`chart:${t}:${r&&r.addr}:${by}`)
    const c={ legend:{set visible(v){s.calls.push(`legend:${v}`)},set position(p){s.calls.push(`legendPos:${p}`)}},
      title:{set text(t){s.calls.push(`title:${t}`)},set visible(v){}},
      axes:{ categoryAxis:{ title:{ set text(t){ s.calls.push(`xAxis:${t}`) }, set visible(v){ s.calls.push(`xAxisVis:${v}`) } } },
             valueAxis:{ title:{ set text(t){ s.calls.push(`yAxis:${t}`) }, set visible(v){ s.calls.push(`yAxisVis:${v}`) } } } },
      series:{ items:[{ hasDataLabels:false,
                        format:{ fill:{ setSolidColor(c){ s.calls.push(`seriesColor:${c}`) } } },
                        set gapWidth(v){ s.calls.push(`gap:${v}`) },
                        set overlap(v){ s.calls.push(`overlap:${v}`) } }] },
      set dataLabels(v){ s.calls.push(`dataLabels:${JSON.stringify(v)}`) },
      get dataLabels(){ return { showValue:false } },
      set left(v){ s.calls.push(`left:${v}`) },
      set top(v){ s.calls.push(`top:${v}`) },
      set width(v){ s.calls.push(`width:${v}`) },
      set height(v){ s.calls.push(`height:${v}`) },
      load(){ return c },
      get name(){ return 'Chart 1' } };
    return c } } }
  get tables(){ const s=this; return { add:(a,n,opts)=>s.calls.push(`table:${a}:${n}`) } }
  get view(){ return {} }
}
class Ws {
  constructor(name){ this.name=name; this.sheet=new Sheet(name, this); this.calls=this.sheet.calls; this.position=0; this.visibility='visible' }
  load(){ return this }
  getRange(a){ return this.sheet.getRange(a) }
  getUsedRangeOrNullObject(){ return this.sheet.getUsedRangeOrNullObject() }
  set position(v){ this._position=v }
  get position(){ return this._position }
  set visibility(v){ this._visibility=v }
  get visibility(){ return this._visibility }
  get columns(){ return { items: [] } }
  get worksheet(){ return this }
  get charts(){ return this.sheet.charts }
  get tables(){ return this.sheet.tables }
  get view(){ return {} }
  delete(){ this.calls.push('deleteSheet') }
}
class Collection {
  constructor(owner){ this.owner=owner; this.items=[] }
  load(){ return this }
  get count(){ return this.items.length }
}
const STORE = { grids: {}, names: ['Sheet1','Sheet2'], charts: [], calls: [] };

class Ctx {
  constructor(){
    this.sheets={};
    for (const n of STORE.names) this.sheets[n]=new Ws(n);
    const self=this;
    this.workbook={
      worksheets:Object.assign(new Collection(this), {
        getItem:(n)=>{ if(!self.sheets[n]) throw new Error(`Sheet "${n}" not found`); return self.sheets[n] },
        getActiveWorksheet(){ return self.sheets.Sheet1 },
        add:(name)=>{ const n=name||`Sheet${STORE.names.length+1}`;
          if(!STORE.names.includes(n)) STORE.names.push(n);
          const w=new Ws(n);
          Object.keys(self.sheets).forEach((k,i)=>{ self.sheets[k].position=i });
          w.position=STORE.names.indexOf(n); self.sheets[n]=w;
          self.workbook.worksheets.items=STORE.names.map(x=>self.sheets[x]);
          return w; } }),
      tables:Object.assign(new Collection(this), { add:()=>({}) }),
      names:Object.assign(new Collection(this), {}),
      getSelectedRange(){ return self.sheets.Sheet1.getRange('A1:D5') } };
    this.workbook.worksheets.items = Object.values(this.sheets);
    Object.values(this.sheets).forEach(w=>{ w.context={workbook:this.workbook, application:{suspendApiCalculationUntilNextSync(){}}}});
  }
  async sync(){ return undefined }
}
globalThis.__ctx = null;
// Load populates props; run() persists loaded values onto the fake after the
// callback, so a get inside the same run() still sees the pre-sync value.
class LoadStore { constructor(obj){ this.obj=obj; this.data={} } }
globalThis.Excel = {
  run: async (fn)=>{
    const c=new Ctx(); globalThis.__ctx=c;
    const r=await fn(c);
    // commit staged writes so they survive the next run()
    for (const ws of Object.values(c.sheets)) {
      for (const [k,v] of Object.entries(ws.sheet.staged||{})) ws.sheet.grid[k]=v;
      ws.sheet.staged=null;
      for (const c of ws.calls) { STORE.calls.push(`${ws.name}::${c}`); }
      globalThis.__storeCalls = STORE.calls;
      ws.calls.splice(0);
    }
    globalThis.__storeGrids = STORE.grids;
    globalThis.__storeCalls = STORE.calls;
    return r;
  },
  ChartType: new Proxy({}, { get:(_,k)=>k }),
  ChartSeriesBy: { auto:'auto', columns:'columns', rows:'rows' },
  RangeCopyType: { all:'all', formulas:'formulas', values:'values', formats:'formats' },
  ClearApplyTo: { all:'all', contents:'contents', formats:'formats' },
  ConditionalFormatType: { cellValue:'CellValue', colorScale:'ColorScale', dataBar:'DataBar', iconSet:'IconSet', topBottom:'TopBottom', presetCriteria:'PresetCriteria', containsText:'ContainsText' },
  ConditionalFormatColorCriterionType: { lowestValue:'LowestValue', highestValue:'HighestValue', percentile:'Percentile', number:'Number' },
  ConditionalFormatPresetCriterion: { blanks:'Blanks', duplicateValues:'DuplicateValues', uniqueValues:'UniqueValues', aboveAverage:'AboveAverage' },
  ConditionalTopBottomCriterionType: { topItems:'TopItems', topPercent:'TopPercent', bottomItems:'BottomItems', bottomPercent:'BottomPercent' },
  ConditionalFormatValueObjectType: { min:'min', max:'max', percentile:'percentile' },
  ChartLegendPosition: { bottom:'Bottom', right:'Right' },
  NumericValue: {}, // placeholder for any other enum
};
export function getCtx(){ return globalThis.__ctx }
