"""Self-contained audit report for extracted tables."""

from __future__ import annotations

import html
import json
from pathlib import Path
from typing import Any


def _json_for_script(value: Any) -> str:
    # Prevent a workbook value containing </script> from breaking out of the
    # embedded JSON block.
    return json.dumps(value, ensure_ascii=False).replace("</", "<\\/")


def build_html_report(report_data: dict[str, Any], destination: Path) -> None:
    """Write a dependency-free HTML table inspector."""

    title = html.escape(str(report_data.get("source", {}).get("filename", "POPS workbook")))
    data = _json_for_script(report_data)
    document = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>POPS extraction · {title}</title>
  <style>
    :root {{ --ink:#10233f; --muted:#607089; --line:#d9e2ec; --soft:#f4f7fb;
      --navy:#0b315e; --green:#08775f; --cyan:#087f9a; --amber:#b66a00; --red:#b42318; }}
    * {{ box-sizing:border-box }}
    body {{ margin:0; color:var(--ink); background:#eef3f8; font:14px/1.45 Inter,Segoe UI,Arial,sans-serif }}
    header {{ padding:22px 28px; color:white; background:linear-gradient(115deg,var(--navy),#075b75) }}
    header h1 {{ margin:0 0 5px; font-size:23px }} header p {{ margin:0; opacity:.85 }}
    .layout {{ display:grid; grid-template-columns:320px minmax(0,1fr); min-height:calc(100vh - 92px) }}
    aside {{ padding:18px; background:white; border-right:1px solid var(--line); overflow:auto }}
    main {{ min-width:0; padding:20px; overflow:auto }}
    .summary {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(145px,1fr)); gap:10px; margin-bottom:16px }}
    .stat,.panel {{ background:white; border:1px solid var(--line); border-radius:10px; box-shadow:0 1px 2px #17324d10 }}
    .stat {{ padding:12px }} .stat b {{ display:block; font-size:21px }} .stat span {{ color:var(--muted) }}
    .panel {{ padding:16px; margin-bottom:14px }}
    .sheet-title {{ margin:18px 0 7px; font-size:12px; letter-spacing:.06em; text-transform:uppercase; color:var(--muted) }}
    button.table-card {{ width:100%; text-align:left; border:1px solid var(--line); background:var(--soft); border-radius:8px;
      margin:0 0 8px; padding:10px; cursor:pointer; color:var(--ink) }}
    button.table-card:hover,button.table-card.active {{ border-color:var(--cyan); background:#e8f7fa }}
    .card-title {{ display:block; font-weight:650 }} .card-meta {{ color:var(--muted); font-size:12px }}
    .badge {{ display:inline-block; padding:2px 7px; border-radius:999px; font-size:11px; font-weight:700; color:white; background:var(--green) }}
    .badge.uncertain {{ background:var(--amber) }} .badge.warning {{ background:var(--red) }}
    .toolbar {{ display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin:8px 0 12px }}
    .toolbar button {{ border:1px solid var(--line); background:white; border-radius:6px; padding:6px 10px; cursor:pointer }}
    .toolbar button.active {{ color:white; border-color:var(--navy); background:var(--navy) }}
    .grid-wrap {{ max-height:65vh; overflow:auto; border:1px solid var(--line); border-radius:8px; background:white }}
    table {{ border-collapse:separate; border-spacing:0; min-width:100% }}
    th,td {{ min-width:92px; max-width:260px; padding:6px 8px; border-right:1px solid var(--line); border-bottom:1px solid var(--line);
      white-space:pre-wrap; overflow-wrap:anywhere; vertical-align:top }}
    th {{ position:sticky; top:0; z-index:2; color:white; background:var(--navy); font-weight:650 }}
    .rownum {{ position:sticky; left:0; min-width:54px; width:54px; z-index:1; color:var(--muted); background:#f8fafc; text-align:right }}
    th.rownum {{ z-index:3; color:white; background:#082645 }}
    tr.section td {{ font-weight:700; background:#e4f2ed }} tr.total td {{ font-weight:700; border-top:2px solid #8091a5 }}
    td.formula {{ background:#eef6ff }} td.blank {{ color:#a0aabd }}
    .warnings {{ margin:0; padding-left:20px }} .warnings li {{ margin:5px 0 }}
    code {{ background:#eff3f7; padding:1px 4px; border-radius:4px }}
    a {{ color:#075b75 }} .empty {{ color:var(--muted); padding:35px; text-align:center }}
    @media (max-width:850px) {{ .layout {{ grid-template-columns:1fr }} aside {{ border-right:0; border-bottom:1px solid var(--line) }} }}
  </style>
</head>
<body>
<header><h1>POPS workbook extraction</h1><p id="source-line"></p></header>
<div class="layout">
  <aside><div id="navigation"></div></aside>
  <main><div id="summary" class="summary"></div><div id="content"></div></main>
</div>
<script id="report-data" type="application/json">{data}</script>
<script>
const DATA=JSON.parse(document.getElementById('report-data').textContent);
const $=s=>document.querySelector(s); let active=null, mode='value';
const esc=v=>String(v??'').replace(/[&<>\"']/g,c=>({{'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',"'":'&#39;'}}[c]));
$('#source-line').textContent=`${{DATA.source.filename}} · SHA-256 ${{DATA.source.sha256.slice(0,16)}}… · schema ${{DATA.schema_version}}`;
const totals=DATA.totals||{{}};
$('#summary').innerHTML=[['Sheets',totals.sheets],['Tables',totals.tables],['Cells',totals.cells],['Formulas',totals.formulas],['Warnings',(DATA.warnings||[]).length]].map(([k,v])=>`<div class="stat"><b>${{v??0}}</b><span>${{k}}</span></div>`).join('');
const nav=[]; (DATA.sheets||[]).forEach(s=>{{ nav.push(`<div class="sheet-title">${{esc(s.name)}} <small>(${{esc(s.state)}})</small></div>`);
 (s.tables||[]).forEach(t=>nav.push(`<button class="table-card" data-id="${{esc(t.table_id)}}"><span class="card-title">${{esc(t.title||t.range)}}</span><span class="card-meta">${{esc(t.range)}} · <span class="badge ${{t.uncertain?'uncertain':''}}">${{Math.round((t.confidence||0)*100)}}%</span></span></button>`)); }});
$('#navigation').innerHTML=nav.join('')||'<p class="empty">No accepted tables.</p>';
document.querySelectorAll('.table-card').forEach(b=>b.addEventListener('click',()=>showTable(b.dataset.id)));
function showTable(id){{ active=DATA.table_lookup[id]; document.querySelectorAll('.table-card').forEach(b=>b.classList.toggle('active',b.dataset.id===id)); render(); }}
 function render(){{ if(!active){{ const w=DATA.warnings||[]; $('#content').innerHTML=`<div class="panel"><h2>Extraction overview</h2><p>Choose a table on the left. The grid retains source coordinates and lets you switch between effective values, exact formulas, and workbook caches.</p><p><a href="manifest.json">Manifest</a> · <a href="schema.json">Schema</a> · <a href="records_long.jsonl">Long JSONL</a> · <a href="records_long.csv">Review-safe CSV</a></p></div><div class="panel"><h3>Warnings</h3>${{w.length?`<ul class="warnings">${{w.map(x=>`<li><code>${{esc(x.code)}}</code> ${{esc(x.message)}}</li>`).join('')}}</ul>`:'<p>None.</p>'}}</div>`; return; }}
 const modes=['value','formula','cached']; const controls=modes.map(m=>`<button data-mode="${{m}}" class="${{mode===m?'active':''}}">${{m}}</button>`).join('');
 const rows=(active.preview||[]).map(r=>`<tr class="${{esc(r.role||'')}}"><td class="rownum">${{r.row}}</td>${{r.cells.map(c=>{{let v=c[mode]; if(mode==='formula'&&!v)v=''; const cls=(c.formula?'formula ':'')+((v===null||v==='')?'blank':''); return `<td class="${{cls}}" title="${{esc(c.coordinate)}}">${{esc(v)}}</td>`;}}).join('')}}</tr>`).join('');
 const heads=(active.columns||[]).map(c=>`<th title="${{esc((c.header_path||[]).join(' / '))}}">${{esc(c.letter)}}<br>${{esc(c.label||c.key||'')}}</th>`).join('');
 $('#content').innerHTML=`<div class="panel"><h2>${{esc(active.title||active.range)}}</h2><p><code>${{esc(active.sheet)}}!${{esc(active.range)}}</code> · ${{esc(active.layout_kind)}} · ${{Math.round((active.confidence||0)*100)}}% confidence</p><div class="toolbar"><span>Show:</span>${{controls}}</div><div class="grid-wrap"><table><thead><tr><th class="rownum">Row</th>${{heads}}</tr></thead><tbody>${{rows}}</tbody></table></div><p>${{esc(active.preview_note||'')}}</p></div>`;
 document.querySelectorAll('[data-mode]').forEach(b=>b.addEventListener('click',()=>{{mode=b.dataset.mode;render();}})); }}
render();
</script>
</body></html>"""
    destination.write_text(document, encoding="utf-8")
