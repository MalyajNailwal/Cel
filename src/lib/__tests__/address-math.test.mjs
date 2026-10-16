// Mirrors the helpers in src/App.tsx. Run: node src/lib/__tests__/address-math.test.mjs
function colToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
function indexToCol(index) {
  let n = index + 1, out = '';
  while (n > 0) { const rem = (n - 1) % 26; out = String.fromCharCode(65 + rem) + out; n = Math.floor((n - 1) / 26); }
  return out;
}
function columnLetter(cols) {
  let n = cols, out = '';
  while (n > 0) { const rem = (n - 1) % 26; out = String.fromCharCode(65 + rem) + out; n = Math.floor((n - 1) / 26); }
  return out || 'A';
}
let fail = 0;
const eq = (name, got, want) => {
  const ok = got === want;
  if (!ok) { fail++; console.log(`FAIL ${name}: got ${got} want ${want}`); }
  else console.log(`ok   ${name} = ${got}`);
};
// columnLetter must be correct across the single->double letter boundary
eq('cols 1', columnLetter(1), 'A');
eq('cols 26', columnLetter(26), 'Z');
eq('cols 27', columnLetter(27), 'AA');
eq('cols 52', columnLetter(52), 'AZ');
eq('cols 53', columnLetter(53), 'BA');
eq('cols 702', columnLetter(702), 'ZZ');
eq('cols 703', columnLetter(703), 'AAA');
// round trip
eq('roundtrip AZ', indexToCol(colToIndex('AZ')), 'AZ');
eq('roundtrip XFD', indexToCol(colToIndex('XFD')), 'XFD');
eq('colToIndex A', colToIndex('A'), 0);
eq('colToIndex Z', colToIndex('Z'), 25);
eq('colToIndex AA', colToIndex('AA'), 26);

// formula clash detection
function clashes(base, values, formulaCells) {
  const set = new Set(formulaCells);
  const m = base.replace(/\$/g, '').match(/^([A-Z]+)(\d+)/);
  if (!m) return [];
  const startCol = colToIndex(m[1]);
  const startRow = parseInt(m[2], 10);
  const out = [];
  for (let r = 0; r < values.length; r++)
    for (let c = 0; c < (values[r]?.length || 0); c++)
      if (set.has(`${indexToCol(startCol + c)}${startRow + r}`)) out.push(`${indexToCol(startCol + c)}${startRow + r}`);
  return out;
}
// D3 sits inside a 2x2 write starting at D2 -> must be caught
eq('clash inside write', JSON.stringify(clashes('D2', [[1,2],[3,4]], ['D3'])), '["D3"]');
// A1 is outside D2:E3 -> no clash
eq('clash outside write', JSON.stringify(clashes('D2', [[1,2],[3,4]], ['A1'])), '[]');
eq('clash hit', JSON.stringify(clashes('C1', [[1,2,3]], ['C1'])), '["C1"]');
eq('clash across rows', JSON.stringify(clashes('A1', [[1],[2]], ['A2'])), '["A2"]');
eq('clash at AA', JSON.stringify(clashes('AA1', [[1]], ['AA1'])), '["AA1"]');
console.log(fail === 0 ? '\nALL PASS' : `\n${fail} FAILED`);
process.exit(fail ? 1 : 0);
