#!/usr/bin/env node
/**
 * MV Family Office -- automatic data refresh
 *
 * Looks for the most recently modified .xlsm/.xlsx file under data/, parses it with the
 * same logic the app uses in the browser (parse-core.js), and writes the result into
 * index.html by replacing the EMBEDDED_STATE marker block. This is what GitHub Actions
 * runs automatically whenever a new Excel file is pushed into data/.
 *
 * It does NOT touch the USERS array -- user management still happens from the app's
 * Admin panel + "Descargar archivo actualizado".
 */
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const MVParse = require('./parse-core.js');

const REPO_ROOT = path.join(__dirname, '..');
const DATA_DIR = path.join(REPO_ROOT, 'data');
const INDEX_PATH = path.join(REPO_ROOT, 'index.html');

function findLatestExcel(dir){
  if (!fs.existsSync(dir)) return null;
  const files = fs.readdirSync(dir)
    .filter(f => /\.(xlsm|xlsx)$/i.test(f))
    .map(f => {
      const full = path.join(dir, f);
      return { full, mtime: fs.statSync(full).mtimeMs };
    })
    .sort((a, b) => b.mtime - a.mtime);
  return files.length ? files[0].full : null;
}

function substituteMarker(html, startMarker, endMarker, replacement){
  const startIdx = html.indexOf(startMarker);
  const endIdx = html.indexOf(endMarker);
  if (startIdx === -1 || endIdx === -1 || endIdx < startIdx) return null;
  const before = html.slice(0, startIdx + startMarker.length);
  const after = html.slice(endIdx);
  return before + replacement + after;
}

function main(){
  const excelPath = findLatestExcel(DATA_DIR);
  if (!excelPath) {
    console.log('No .xlsm/.xlsx file found under data/. Nothing to do.');
    process.exit(0);
  }
  console.log('Processing:', excelPath);

  const wb = XLSX.readFile(excelPath, { cellDates: true });
  const state = MVParse.parseAll(wb);
  console.log('Parsed OK. netWorth =', state.dashboard && state.dashboard.netWorth, '| asOf =', state.meta && state.meta.asOf);

  if (!fs.existsSync(INDEX_PATH)) {
    console.error('ERROR: index.html not found at repo root. Publish the MV Family Office app there first.');
    process.exit(1);
  }
  let html = fs.readFileSync(INDEX_PATH, 'utf8');

  const stateJs = JSON.stringify(state);
  const newHtml = substituteMarker(
    html, '/*__DATA_START__*/', '/*__DATA_END__*/',
    '\n  const EMBEDDED_STATE = ' + stateJs + ';\n  '
  );
  if (newHtml === null) {
    console.error('ERROR: could not find the /*__DATA_START__*/ ... /*__DATA_END__*/ markers in index.html.');
    console.error('Make sure index.html is the latest MV Family Office build (it must contain those markers).');
    process.exit(1);
  }

  fs.writeFileSync(INDEX_PATH, newHtml, 'utf8');
  console.log('index.html updated with the new snapshot.');
}

main();
