#!/usr/bin/env node
// XML-level regression check for DOCX table widths.
// Usage: node scripts/verify-table-xml.js <file.docx>
// Exits 0 on PASS, 1 on FAIL, 2 on error.
//
// Validates the fix for: docs/10_分析/20260527_DOCX表格列宽导致QuickLook预览挤压.md

const fs = require('fs');
const { execSync } = require('child_process');

const docx = process.argv[2];
if (!docx) {
    console.error('Usage: node scripts/verify-table-xml.js <file.docx>');
    process.exit(2);
}
if (!fs.existsSync(docx)) {
    console.error(`File not found: ${docx}`);
    process.exit(2);
}

let xml;
try {
    xml = execSync(`unzip -p "${docx}" word/document.xml`, { encoding: 'utf8' });
} catch (err) {
    console.error(`Failed to unzip: ${err.message}`);
    process.exit(2);
}

const checks = [];

// Check 1: no literal w:w="100%"
checks.push({
    name: 'No literal w:w="100%"',
    pass: !/w:w="100%"/.test(xml),
});

// Check 2: every tblW uses legal type (attr-order-independent)
const tblWMatches = [...xml.matchAll(/<w:tblW\b([^/>]*)/g)];
const tblWBad = tblWMatches.filter(m => {
    const attrs = m[1];
    const hasType = /w:type="([^"]+)"/.exec(attrs);
    const hasW = /w:w="([^"]+)"/.exec(attrs);
    if (!hasType) return true;
    const type = hasType[1];
    const wVal = hasW ? hasW[1] : '';
    if (type === 'dxa') return !/^\d+$/.test(wVal);
    if (type === 'auto') return false;
    if (type === 'pct') return !/^\d+$/.test(wVal);
    return true;
});
checks.push({
    name: `All tblW use legal type (dxa/auto/pct-integer): ${tblWMatches.length} tables`,
    pass: tblWBad.length === 0,
    detail: tblWBad.length > 0 ? `bad: ${tblWBad.map(b => b[0]).join(', ')}` : null,
});

// Check 3: no tblGrid has all gridCols = 100
const tableBlocks = xml.split('<w:tbl>').slice(1);
const allTblGridBad = tableBlocks.map((block, idx) => {
    const tblGridMatch = block.match(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/);
    if (!tblGridMatch) return null;
    const gridCols = [...tblGridMatch[1].matchAll(/w:w="(\d+)"/g)].map(m => parseInt(m[1], 10));
    const allHundred = gridCols.length > 0 && gridCols.every(w => w === 100);
    return allHundred ? { idx, gridCols } : null;
}).filter(Boolean);
checks.push({
    name: `No table has all gridCol = 100 twips (${tableBlocks.length} tables checked)`,
    pass: allTblGridBad.length === 0,
    detail: allTblGridBad.length > 0 ? `bad tables: ${JSON.stringify(allTblGridBad)}` : null,
});

// Check 4: every table has tblLayout w:type="fixed"
const tablesWithoutFixed = tableBlocks.filter(block => !/<w:tblLayout\s+w:type="fixed"/.test(block));
checks.push({
    name: `All tables have tblLayout=fixed (${tableBlocks.length} tables)`,
    pass: tablesWithoutFixed.length === 0,
    detail: tablesWithoutFixed.length > 0 ? `${tablesWithoutFixed.length} tables missing tblLayout=fixed` : null,
});

// Check 5: gridCol total width sane (sum > 5000 twips ~= 8.8cm)
const tinyGridCols = [];
tableBlocks.forEach((block, idx) => {
    const tblGridMatch = block.match(/<w:tblGrid>([\s\S]*?)<\/w:tblGrid>/);
    if (!tblGridMatch) return;
    const gridCols = [...tblGridMatch[1].matchAll(/w:w="(\d+)"/g)].map(m => parseInt(m[1], 10));
    const sum = gridCols.reduce((a, b) => a + b, 0);
    if (sum < 5000) tinyGridCols.push({ idx, sum, gridCols });
});
checks.push({
    name: 'tblGrid total width sane (sum > 5000 twips ~= 8.8cm)',
    pass: tinyGridCols.length === 0,
    detail: tinyGridCols.length > 0 ? `tables with narrow total: ${JSON.stringify(tinyGridCols)}` : null,
});

// Report
console.log(`\n=== XML Table Width Check for ${docx} ===\n`);
let allPass = true;
checks.forEach(c => {
    const icon = c.pass ? '✅' : '❌';
    console.log(`${icon} ${c.name}`);
    if (!c.pass && c.detail) console.log(`   ${c.detail}`);
    if (!c.pass) allPass = false;
});

if (allPass) {
    console.log(`\n✅ ALL CHECKS PASSED\n`);
    process.exit(0);
} else {
    console.log(`\n❌ SOME CHECKS FAILED\n`);
    process.exit(1);
}
