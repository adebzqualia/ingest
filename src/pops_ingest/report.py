"""Self-contained audit report for extracted and clean tables."""

from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any


def _json_for_script(value: Any) -> str:
    """Serialize JSON without permitting an embedded HTML/script boundary."""

    return (
        json.dumps(value, ensure_ascii=False)
        .replace("&", "\\u0026")
        .replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("\u2028", "\\u2028")
        .replace("\u2029", "\\u2029")
    )


def build_html_report(report_data: dict[str, Any], destination: Path) -> None:
    """Write a dependency-free clean-table and raw-audit inspector."""

    title = html.escape(
        str(report_data.get("source", {}).get("filename", "POPS workbook")),
        quote=True,
    )
    data = _json_for_script(report_data)
    document = _REPORT_TEMPLATE.replace("__REPORT_TITLE__", title, 1).replace(
        "__REPORT_DATA__", data, 1
    )
    destination.write_text(document, encoding="utf-8")


_REPORT_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; img-src data:; connect-src 'none'; base-uri 'none'; form-action 'none'">
  <title>POPS extraction · __REPORT_TITLE__</title>
  <style>
    :root {
      --ink:#10233f; --muted:#607089; --line:#d9e2ec; --soft:#f4f7fb;
      --navy:#0b315e; --green:#08775f; --cyan:#087f9a; --amber:#b66a00;
      --red:#b42318; --white:#fff; --label:#f7fafc;
    }
    * { box-sizing:border-box }
    [hidden] { display:none !important }
    body { margin:0; color:var(--ink); background:#eef3f8; font:14px/1.45 Inter,Segoe UI,Arial,sans-serif }
    header { padding:20px clamp(18px,3vw,30px); color:white; background:linear-gradient(115deg,var(--navy),#075b75) }
    header h1 { margin:0 0 5px; font-size:clamp(20px,2vw,24px) }
    header p { margin:0; opacity:.86; overflow-wrap:anywhere }
    .layout { display:grid; grid-template-columns:310px minmax(0,1fr); min-height:calc(100vh - 88px) }
    aside { min-width:0; padding:16px; background:white; border-right:1px solid var(--line); overflow:auto }
    main { min-width:0; padding:20px; overflow:hidden }
    .summary { display:grid; grid-template-columns:repeat(auto-fit,minmax(125px,1fr)); gap:10px; margin-bottom:16px }
    .stat,.panel { background:white; border:1px solid var(--line); border-radius:10px; box-shadow:0 1px 2px #17324d10 }
    .stat { padding:11px 12px }
    .stat b { display:block; font-size:20px }
    .stat span { color:var(--muted) }
    .panel { min-width:0; padding:16px; margin-bottom:14px }
    .panel h2 { margin:0; font-size:20px }
    .table-heading { display:flex; align-items:flex-start; justify-content:space-between; gap:12px; flex-wrap:wrap }
    .table-heading p { margin:5px 0 0; color:var(--muted) }
    .view-badge { flex:none; display:inline-flex; align-items:center; gap:5px; padding:4px 9px; border-radius:999px;
      color:#075b4b; background:#e4f4ef; font-size:12px; font-weight:700 }
    .nav-summary { display:flex; justify-content:space-between; gap:8px; margin-bottom:12px; color:var(--muted); font-size:12px }
    .nav-summary strong { color:var(--ink) }
    .sheet-title { margin:17px 0 7px; font-size:12px; letter-spacing:.06em; text-transform:uppercase; color:var(--muted) }
    button { font:inherit }
    button.table-card { width:100%; text-align:left; border:1px solid var(--line); background:var(--soft); border-radius:8px;
      margin:0 0 8px; padding:10px; cursor:pointer; color:var(--ink) }
    button.table-card:hover,button.table-card.active { border-color:var(--cyan); background:#e8f7fa }
    button.table-card:focus-visible,.toolbar button:focus-visible,.diagnostic-toggle:focus-visible { outline:3px solid #48a7bb66; outline-offset:1px }
    .card-title { display:block; font-weight:650; overflow-wrap:anywhere }
    .card-meta { display:block; margin-top:3px; color:var(--muted); font-size:12px; overflow-wrap:anywhere }
    .badge { display:inline-block; padding:2px 7px; border-radius:999px; font-size:11px; font-weight:700; color:white; background:var(--green) }
    .badge.uncertain { background:var(--amber) }
    .badge.warning { background:var(--red) }
    .diagnostic-toggle { display:flex; width:100%; align-items:center; justify-content:space-between; gap:8px; margin:16px 0 7px;
      padding:8px 10px; border:1px solid var(--line); border-radius:8px; color:var(--muted); background:white; cursor:pointer }
    .diagnostic-toggle:hover { color:var(--ink); background:var(--soft) }
    .diagnostic-count { min-width:24px; padding:1px 7px; border-radius:999px; color:white; background:var(--muted); text-align:center; font-size:11px; font-weight:700 }
    .diagnostics { padding:7px; border:1px dashed #b9c5d2; border-radius:8px; background:#fbfcfe }
    .diagnostics-note { margin:2px 2px 9px; color:var(--muted); font-size:12px }
    .diagnostic-card { background:white !important }
    .toolbar { display:flex; align-items:center; gap:12px; flex-wrap:wrap; margin:13px 0 12px }
    .control-group { display:inline-flex; align-items:center; gap:4px; padding:3px; border:1px solid var(--line); border-radius:8px; background:var(--soft) }
    .control-label { margin:0 3px 0 5px; color:var(--muted); font-size:12px; font-weight:650 }
    .toolbar button { border:1px solid transparent; background:transparent; border-radius:6px; padding:6px 10px; cursor:pointer; color:var(--ink) }
    .toolbar button:hover:not(:disabled) { background:white }
    .toolbar button.active { color:white; border-color:var(--navy); background:var(--navy) }
    .toolbar button:disabled { color:#9aa6b5; cursor:not-allowed }
    .grid-wrap { width:100%; max-height:calc(100vh - 305px); min-height:160px; overflow:auto; overscroll-behavior:contain;
      border:1px solid var(--line); border-radius:8px; background:white; scrollbar-gutter:stable }
    table { border-collapse:separate; border-spacing:0; min-width:100% }
    th,td { min-width:96px; max-width:300px; padding:7px 9px; border-right:1px solid var(--line); border-bottom:1px solid var(--line);
      white-space:pre-wrap; overflow-wrap:anywhere; vertical-align:top }
    thead th { position:sticky; top:0; z-index:3; color:white; background:var(--navy); font-weight:650; text-align:left;
      box-shadow:0 1px 0 var(--line) }
    tbody tr:nth-child(even) td { background:#fbfcfe }
    tr.section td,tr.section th { font-weight:700; background:#e4f2ed }
    tr.total td,tr.total th { font-weight:700; border-top:2px solid #8091a5 }
    td.formula { background:#eef6ff }
    td.blank,th.blank { color:#a0aabd }
    .raw-table { width:max-content }
    .rownum { position:sticky; left:0; min-width:56px; width:56px; max-width:56px; z-index:2; color:var(--muted);
      background:#f8fafc !important; text-align:right }
    thead .rownum { z-index:5; color:white; background:#082645 !important }
    .clean-table { width:max-content; table-layout:auto }
    .clean-table th,.clean-table td { min-width:112px }
    .clean-table thead th:first-child { left:0; z-index:6; min-width:clamp(210px,28vw,390px); max-width:440px; background:#063f53 }
    .clean-table .clean-label { position:sticky; left:0; z-index:2; min-width:clamp(210px,28vw,390px); max-width:440px;
      color:var(--ink); background:var(--label); text-align:left; font-weight:650; box-shadow:1px 0 0 var(--line) }
    .clean-table tbody tr:nth-child(even) .clean-label { background:#eef4f8 }
    .clean-table tr.section .clean-label { background:#dceee8 }
    .clean-table tr.total .clean-label { background:#edf1f5 }
    .table-note { margin:9px 0 0; color:var(--muted); font-size:12px }
    .warnings { margin:0; padding-left:20px }
    .warnings li { margin:5px 0 }
    code { background:#eff3f7; padding:1px 4px; border-radius:4px; overflow-wrap:anywhere }
    a { color:#075b75 }
    .empty { color:var(--muted); padding:30px; text-align:center }
    .sr-only { position:absolute; width:1px; height:1px; padding:0; margin:-1px; overflow:hidden; clip:rect(0,0,0,0); white-space:nowrap; border:0 }
    @media (max-width:850px) {
      .layout { grid-template-columns:1fr }
      aside { max-height:42vh; border-right:0; border-bottom:1px solid var(--line) }
      main { padding:12px }
      .grid-wrap { max-height:62vh }
      .panel { padding:12px }
      .clean-table thead th:first-child,.clean-table .clean-label { min-width:190px; max-width:270px }
    }
    @media (prefers-reduced-motion:reduce) { * { scroll-behavior:auto !important } }
  </style>
</head>
<body>
<header><h1>POPS workbook extraction</h1><p id="source-line"></p></header>
<div class="layout">
  <aside><div id="navigation"></div></aside>
  <main><div id="summary" class="summary"></div><div id="content"></div></main>
</div>
<script id="report-data" type="application/json">__REPORT_DATA__</script>
<script>
'use strict';
const DATA=JSON.parse(document.getElementById('report-data').textContent);
const $=selector=>document.querySelector(selector);
const state={active:null,entry:null,mode:'value',view:'clean',showExcluded:false};
const esc=value=>String(value??'').replace(/[&<>"']/g,char=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
const own=(object,key)=>object!==null&&typeof object==='object'&&Object.prototype.hasOwnProperty.call(object,key);
const pick=(object,keys)=>{for(const key of keys){if(own(object,key)&&object[key]!==undefined)return object[key];}return undefined;};
const safeRole=value=>{const role=String(value||'').toLowerCase();return ['section','total','subtotal'].includes(role)?role:'';};
const friendlyStatus=value=>String(value||'not ready').replace(/[_-]+/g,' ');
const tableLookup=id=>own(DATA.table_lookup||{},id)?DATA.table_lookup[id]:null;
const cleanReady=table=>Boolean(table&&table.clean&&table.clean.status==='ready');

const source=DATA.source||{};
const sourceBits=[source.filename||'POPS workbook'];
if(typeof source.sha256==='string'&&source.sha256)sourceBits.push(`SHA-256 ${source.sha256.slice(0,16)}…`);
if(DATA.schema_version)sourceBits.push(`schema ${DATA.schema_version}`);
$('#source-line').textContent=sourceBits.join(' · ');

const entries=[];
(DATA.sheets||[]).forEach(sheet=>(sheet.tables||[]).forEach(table=>entries.push({sheet,table})));
const readyEntries=entries.filter(item=>item.table.clean_status==='ready');
const excludedEntries=entries.filter(item=>item.table.clean_status!=='ready');
const totals=DATA.totals||{};
$('#summary').innerHTML=[
  ['Sheets',totals.sheets],['Clean tables',readyEntries.length],['Excluded regions',excludedEntries.length],
  ['Cells',totals.cells],['Formulas',totals.formulas],['Warnings',warningCount()]
].map(([label,value])=>`<div class="stat"><b>${esc(value??0)}</b><span>${esc(label)}</span></div>`).join('');

function tableCard(item,diagnostic=false){
  const table=item.table;
  const clean=table.clean_status==='ready';
  const cardTitle=clean?(table.clean_title||'Clean table'):(table.title||table.range||'Excluded region');
  const confidence=Math.round((Number(table.confidence)||0)*100);
  const badge=clean
    ? `<span class="badge ${table.uncertain?'uncertain':''}">${confidence}%</span>`
    : `<span class="badge warning">${esc(friendlyStatus(table.clean_status))}</span>`;
  const sheetMeta=diagnostic?`${esc(item.sheet.name)} · `:'';
  return `<button type="button" class="table-card ${diagnostic?'diagnostic-card':''} ${state.entry===item?'active':''}" data-id="${esc(table.table_id)}"><span class="card-title">${esc(cardTitle)}</span><span class="card-meta">${sheetMeta}${esc(table.range)} · ${badge}</span></button>`;
}

function renderNavigation(){
  const parts=[`<div class="nav-summary"><strong>${readyEntries.length} clean table${readyEntries.length===1?'':'s'}</strong><span>review view</span></div>`];
  (DATA.sheets||[]).forEach(sheet=>{
    const sheetEntries=readyEntries.filter(item=>item.sheet===sheet);
    if(!sheetEntries.length)return;
    parts.push(`<div class="sheet-title">${esc(sheet.name)} <small>(${esc(sheet.state)})</small></div>`);
    sheetEntries.forEach(item=>parts.push(tableCard(item)));
  });
  if(!readyEntries.length)parts.push('<p class="empty">No clean tables are ready. Open diagnostics to inspect raw regions.</p>');
  parts.push(`<button type="button" id="diagnostic-toggle" class="diagnostic-toggle" aria-expanded="${state.showExcluded}"><span>${state.showExcluded?'Hide':'Show'} excluded regions</span><span class="diagnostic-count">${excludedEntries.length}</span></button>`);
  parts.push(`<div class="diagnostics" ${state.showExcluded?'':'hidden'}><p class="diagnostics-note">Excluded or uncertain regions remain available as raw audit views.</p>${excludedEntries.length?excludedEntries.map(item=>tableCard(item,true)).join(''):'<p class="diagnostics-note">No excluded regions.</p>'}</div>`);
  $('#navigation').innerHTML=parts.join('');
  document.querySelectorAll('.table-card').forEach(button=>button.addEventListener('click',()=>showTable(button.dataset.id)));
  $('#diagnostic-toggle').addEventListener('click',()=>{state.showExcluded=!state.showExcluded;renderNavigation();});
}

function findEntry(id){return entries.find(item=>String(item.table.table_id)===String(id))||null;}
function showTable(id){
  const table=tableLookup(id);
  if(!table)return;
  state.active=table;
  state.entry=findEntry(id);
  state.view=cleanReady(table)?'clean':'raw';
  renderNavigation();
  render();
}

function formulaText(value){
  if(value===null||value===undefined)return '';
  if(typeof value!=='object')return value;
  const nested=pick(value,['exact','resolved','raw','formula_text','text','formula']);
  return nested===undefined?value:nested;
}

function displayScalar(value){
  if(value===null||value===undefined)return '';
  if(Array.isArray(value))return value.map(displayScalar).join(' / ');
  if(typeof value==='object'){
    const preferred=pick(value,['display','text','label','value','exact','resolved','raw']);
    if(preferred!==undefined&&preferred!==value)return displayScalar(preferred);
    try{return JSON.stringify(value);}catch(_error){return String(value);}
  }
  return value;
}

function valueForMode(cell,mode){
  if(cell===null||cell===undefined||typeof cell!=='object'||Array.isArray(cell))return mode==='value'?cell:'';
  if(mode==='formula')return formulaText(pick(cell,['formula','exact_formula','formula_text','raw_formula']));
  if(mode==='cached')return pick(cell,['cached','cached_value','cache_value','cache']);
  const value=pick(cell,['value','display','text','clean_value','effective_value']);
  if(value!==undefined&&value!==null)return value;
  const literal=pick(cell,['literal','literal_value']);
  if(literal!==undefined&&literal!==null)return literal;
  return formulaText(pick(cell,['formula','exact_formula','formula_text','raw_formula']));
}

function coordinateFor(cell){
  if(cell===null||typeof cell!=='object'||Array.isArray(cell))return '';
  return pick(cell,['coordinate','source_cell','cell','ref'])||'';
}

function renderedCell(cell,mode){
  const raw=valueForMode(cell,mode);
  return {value:displayScalar(raw),coordinate:coordinateFor(cell),hasFormula:Boolean(cell&&typeof cell==='object'&&pick(cell,['formula','exact_formula','formula_text','raw_formula']))};
}

function cleanGrid(){
  const clean=state.active.clean||{};
  const rows=Array.isArray(clean.rows)?clean.rows:[];
  const suppliedColumns=Array.isArray(clean.columns)?clean.columns:[];
  const widest=rows.reduce((width,row)=>Math.max(width,Array.isArray(row&&row.cells)?row.cells.length:0),0);
  const width=Math.max(suppliedColumns.length,widest);
  const columns=Array.from({length:width},(_unused,index)=>suppliedColumns[index]||{});
  const heads=columns.map((column,index)=>{
    const label=own(column,'label')?column.label:`Column ${index+1}`;
    return `<th scope="col" title="${esc(label)}">${esc(label)}</th>`;
  }).join('');
  const body=rows.map((row,rowIndex)=>{
    const sourceCells=Array.isArray(row&&row.cells)?row.cells:[];
    const cells=columns.map((_column,index)=>sourceCells[index]);
    const rendered=cells.map((cell,index)=>{
      const role=String(columns[index]&&columns[index].role||'').toLowerCase();
      const structural=['row_group','row_label','dimension'].includes(role);
      return renderedCell(cell,structural?'value':state.mode);
    });
    const htmlCells=rendered.map((cell,index)=>{
      const blank=cell.value===null||cell.value===undefined||cell.value==='';
      const classes=`${index===0?'clean-label ':''}${cell.hasFormula?'formula ':''}${blank?'blank':''}`.trim();
      const tag=index===0?'th':'td';
      const scope=index===0?' scope="row"':'';
      return `<${tag}${scope} class="${classes}" title="${esc(cell.coordinate)}">${esc(cell.value)}</${tag}>`;
    }).join('');
    return `<tr class="${safeRole(row&&row.role)}" data-row-index="${rowIndex+1}">${htmlCells}</tr>`;
  }).join('');
  const empty=`<tr><td colspan="${Math.max(1,width)}" class="empty">No clean rows available.</td></tr>`;
  return `<div class="grid-wrap clean-grid"><table class="clean-table"><caption class="sr-only">Clean structured table</caption><thead><tr>${heads||'<th scope="col">Clean table</th>'}</tr></thead><tbody>${body||empty}</tbody></table></div>`;
}

function rawGrid(){
  const columns=Array.isArray(state.active.columns)?state.active.columns:[];
  const preview=Array.isArray(state.active.preview)?state.active.preview:[];
  const heads=columns.map(column=>`<th scope="col" title="${esc((column.header_path||[]).join(' / '))}">${esc(column.letter)}<br>${esc(column.label||column.key||'')}</th>`).join('');
  const rows=preview.map(row=>{
    const sourceCells=Array.isArray(row&&row.cells)?row.cells:[];
    const cells=sourceCells.map(cell=>{
      const rendered=renderedCell(cell,state.mode);
      const blank=rendered.value===null||rendered.value===undefined||rendered.value==='';
      const classes=`${rendered.hasFormula?'formula ':''}${blank?'blank':''}`.trim();
      return `<td class="${classes}" title="${esc(rendered.coordinate)}">${esc(rendered.value)}</td>`;
    }).join('');
    return `<tr class="${safeRole(row&&row.role)}"><th scope="row" class="rownum">${esc(row&&row.row)}</th>${cells}</tr>`;
  }).join('');
  const empty=`<tr><td colspan="${Math.max(1,columns.length+1)}" class="empty">No raw preview rows available.</td></tr>`;
  return `<div class="grid-wrap raw-grid"><table class="raw-table"><caption class="sr-only">Raw workbook audit table</caption><thead><tr><th scope="col" class="rownum">Row</th>${heads}</tr></thead><tbody>${rows||empty}</tbody></table></div>`;
}

function aggregateWarnings(warnings){
  const grouped=new Map();
  warnings.forEach(item=>{
    const code=String(item&&item.code||'WARNING');
    const current=grouped.get(code);
    if(current){current.count+=1;return;}
    grouped.set(code,{code,count:1,message:String(item&&item.message||''),severity:String(item&&item.severity||'warning')});
  });
  return Array.from(grouped.values()).sort((left,right)=>right.count-left.count||left.code.localeCompare(right.code));
}

function warningCount(){
  const supplied=Number(DATA.warning_total);
  return Number.isFinite(supplied)&&supplied>=0?supplied:(DATA.warnings||[]).length;
}

function warningGroups(){
  if(Array.isArray(DATA.warning_summary)){
    return DATA.warning_summary.map(item=>({
      code:String(item&&item.code||'WARNING'),
      count:Math.max(0,Number(item&&item.count)||0),
      message:String(item&&item.message||''),
      severity:String(item&&item.severity||'warning')
    })).sort((left,right)=>right.count-left.count||left.code.localeCompare(right.code));
  }
  return aggregateWarnings(DATA.warnings||[]);
}

function overview(){
  const totalWarnings=warningCount();
  const groups=warningGroups();
  const warningBody=groups.length
    ? `<p class="table-note">${totalWarnings} warning record${totalWarnings===1?'':'s'} grouped into ${groups.length} code${groups.length===1?'':'s'}. The manifest retains every occurrence.</p><ul class="warnings">${groups.map(item=>`<li><code>${esc(item.code)}</code> <strong>×${item.count}</strong> ${esc(item.message)}</li>`).join('')}</ul>`
    : (totalWarnings?`<p>${totalWarnings} warning record${totalWarnings===1?'':'s'}. See the manifest for complete diagnostics.</p>`:'<p>None.</p>');
  return `<div class="panel"><h2>Extraction overview</h2><p>Choose a clean table on the left. Clean tables are shown first; raw coordinates, formulas, and workbook caches remain one click away.</p><p><a href="manifest.json">Manifest</a> · <a href="schema.json">Schema</a> · <a href="records_long.jsonl">Long JSONL</a> · <a href="records_long.csv">Review-safe CSV</a></p></div><div class="panel"><h3>Warnings</h3>${warningBody}</div>`;
}

function render(){
  if(!state.active){$('#content').innerHTML=overview();return;}
  const hasClean=cleanReady(state.active);
  if(state.view==='clean'&&!hasClean)state.view='raw';
  const entry=state.entry&&state.entry.table;
  const cleanTitle=entry&&entry.clean_title?entry.clean_title:'Clean table';
  const title=state.view==='clean'?cleanTitle:(state.active.title||state.active.range||'Raw region');
  const viewControls=['clean','raw'].map(view=>`<button type="button" data-view="${view}" class="${state.view===view?'active':''}" aria-pressed="${state.view===view}" ${view==='clean'&&!hasClean?'disabled title="No clean table is available for this region"':''}>${view==='clean'?'Clean':'Raw'}</button>`).join('');
  const modeControls=['value','formula','cached'].map(mode=>`<button type="button" data-mode="${mode}" class="${state.mode===mode?'active':''}" aria-pressed="${state.mode===mode}">${mode[0].toUpperCase()+mode.slice(1)}</button>`).join('');
  const grid=state.view==='clean'?cleanGrid():rawGrid();
  const note=state.view==='clean'
    ? (state.active.clean&&pick(state.active.clean,['preview_note','note']))
    : state.active.preview_note;
  const meta=`<code>${esc(state.active.sheet)}!${esc(state.active.range)}</code>${state.view==='raw'?` · ${esc(state.active.layout_kind||'raw audit')} · ${Math.round((Number(state.active.confidence)||0)*100)}% confidence`:' · structured clean view'}`;
  $('#content').innerHTML=`<div class="panel"><div class="table-heading"><div><h2>${esc(title)}</h2><p>${meta}</p></div><span class="view-badge">${state.view==='clean'?'Clean table':'Raw audit'}</span></div><div class="toolbar"><div class="control-group"><span class="control-label">View</span>${viewControls}</div><div class="control-group"><span class="control-label">Show</span>${modeControls}</div></div>${grid}${note?`<p class="table-note">${esc(note)}</p>`:''}</div>`;
  document.querySelectorAll('[data-view]').forEach(button=>button.addEventListener('click',()=>{if(button.disabled)return;state.view=button.dataset.view;render();}));
  document.querySelectorAll('[data-mode]').forEach(button=>button.addEventListener('click',()=>{state.mode=button.dataset.mode;render();}));
}

const firstReady=readyEntries.length?readyEntries[0].table.table_id:null;
if(firstReady&&tableLookup(firstReady))showTable(firstReady);
else{renderNavigation();render();}
</script>
</body>
</html>
"""
