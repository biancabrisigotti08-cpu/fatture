"""
Estrattore Fatture Web - Flask Backend
"""
import zipfile
import io
import os
import json
import urllib.parse
import hmac
from flask import Flask, request, send_file, jsonify, render_template_string, Response
import crediti
import estrazione
import excel
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max upload
PRICING_URL = os.environ.get("PRICING_URL", "#prezzi")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "")
crediti.inizializza()

# ─── HTML Template ────────────────────────────────────────────────────────────
HTML = '''<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Estrattore Fatture</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Space+Grotesk:wght@500;600;700&display=swap" rel="stylesheet"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
 --bg:#0a0a0f;
 --surface:#111118;
 --surface2:#1a1a24;
 --border:#ffffff0f;
 --border2:#ffffff18;
 --accent:#ff6b35;
 --accent2:#ff9a6c;
 --green:#00e5a0;
 --red:#ff4d6d;
 --text:#f0f0f5;
 --muted:#7070a0;
 --dim:#3a3a55;
}
body{
 background:var(--bg);
 color:var(--text);
 font-family:'Inter',sans-serif;
 min-height:100vh;
 display:flex;
 flex-direction:column;
 align-items:center;
 padding:0 0 80px;
}
::-webkit-scrollbar{width:4px}
::-webkit-scrollbar-thumb{background:var(--accent);border-radius:2px}
/* HERO */
.hero{
 width:100%;
 background:linear-gradient(135deg,#0a0a0f 0%,#12101e 50%,#0f1218 100%);
 border-bottom:1px solid var(--border2);
 padding:60px 24px 48px;
 text-align:center;
 position:relative;
 overflow:hidden;
}
.hero::before{
 content:'';
 position:absolute;
 top:-120px;left:50%;transform:translateX(-50%);
 width:600px;height:300px;
 background:radial-gradient(ellipse,rgba(255,107,53,0.15) 0%,transparent 70%);
 pointer-events:none;
}
.hero-badge{
 display:inline-flex;align-items:center;gap:8px;
 background:rgba(255,107,53,0.1);
 border:1px solid rgba(255,107,53,0.3);
 color:var(--accent2);
 font-size:12px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;
 padding:6px 16px;border-radius:100px;
 margin-bottom:24px;
}
.hero-badge::before{content:'●';font-size:8px;color:var(--accent)}
h1{
 font-family:'Space Grotesk',sans-serif;
 font-size:clamp(32px,5vw,56px);
 font-weight:700;
 color:#fff;
 letter-spacing:-1.5px;
 line-height:1.1;
 margin-bottom:12px;
}
h1 span{
 background:linear-gradient(90deg,var(--accent),var(--accent2));
 -webkit-background-clip:text;-webkit-text-fill-color:transparent;
}
.hero-sub{
 font-size:16px;color:var(--muted);font-weight:400;
 max-width:480px;margin:0 auto;line-height:1.6;
}
.hero-features{
 list-style:none;display:flex;flex-direction:column;gap:10px;
 max-width:520px;margin:32px auto 0;text-align:left;
}
.hero-features li{
 font-size:14px;color:var(--text);line-height:1.5;
 display:flex;gap:10px;align-items:baseline;
}
.hero-features li::before{content:'✓';color:var(--accent);font-weight:700;flex-shrink:0}
.hero-features strong{color:#fff;font-weight:600}
/* MAIN */
.main{width:100%;max-width:680px;padding:40px 24px 0;display:flex;flex-direction:column;gap:20px}
/* CARD */
.card{
 background:var(--surface);
 border:1px solid var(--border);
 border-radius:16px;
 overflow:hidden;
}
.card-header{
 padding:16px 20px;
 border-bottom:1px solid var(--border);
 display:flex;align-items:center;gap:10px;
}
.card-icon{
 width:32px;height:32px;border-radius:8px;
 background:rgba(255,107,53,0.1);
 display:flex;align-items:center;justify-content:center;
 font-size:16px;
}
.card-title{font-size:14px;font-weight:600;color:#fff}
.card-sub{font-size:12px;color:var(--muted);margin-top:1px}
/* UPLOAD */
.drop-zone{
 padding:32px 24px;
 display:flex;flex-direction:column;align-items:center;gap:10px;
 cursor:pointer;
 transition:.2s;
 border-bottom:1px solid var(--border);
 background:transparent;
}
.drop-zone:hover,.drop-zone.over{background:rgba(255,107,53,0.04)}
.drop-icon-wrap{
 width:56px;height:56px;border-radius:16px;
 background:linear-gradient(135deg,rgba(255,107,53,0.15),rgba(255,107,53,0.05));
 border:1px solid rgba(255,107,53,0.2);
 display:flex;align-items:center;justify-content:center;
 font-size:24px;
 transition:.2s;
}
.drop-zone:hover .drop-icon-wrap{
 background:linear-gradient(135deg,rgba(255,107,53,0.25),rgba(255,107,53,0.1));
 transform:translateY(-2px);
}
.drop-text{font-size:14px;font-weight:500;color:var(--text)}
.drop-sub{font-size:12px;color:var(--muted)}
.or-row{
 display:flex;align-items:center;gap:12px;
 padding:0 24px;
 font-size:11px;color:var(--dim);letter-spacing:1px;text-transform:uppercase;
}
.or-row::before,.or-row::after{content:'';flex:1;height:1px;background:var(--border)}
.folder-btn{
 padding:16px 24px;
 display:flex;align-items:center;gap:12px;
 cursor:pointer;transition:.2s;
 background:transparent;font-family:'Inter',sans-serif;
 color:var(--muted);font-size:13px;width:100%;
}
.folder-btn:hover{background:rgba(255,255,255,0.03);color:var(--text)}
.folder-icon{
 width:36px;height:36px;border-radius:10px;
 background:var(--surface2);border:1px solid var(--border);
 display:flex;align-items:center;justify-content:center;font-size:18px;
 flex-shrink:0;
}
.folder-text{text-align:left}
.folder-label{font-weight:500;color:var(--text);font-size:13px}
.folder-desc{font-size:11px;color:var(--muted);margin-top:2px}
/* FILE LIST */
.file-panel{background:var(--surface);border:1px solid var(--border);border-radius:16px;overflow:hidden;display:none}
.file-panel.show{display:block}
.file-panel-header{
 padding:12px 16px;border-bottom:1px solid var(--border);
 display:flex;justify-content:space-between;align-items:center;
 background:var(--surface2);
}
.file-count{font-size:12px;font-weight:600;color:var(--accent)}
.clear-btn{
 background:transparent;border:1px solid var(--border);color:var(--muted);
 font-size:11px;padding:4px 10px;border-radius:6px;cursor:pointer;font-family:'Inter',sans-serif;
 transition:.15s;
}
.clear-btn:hover{color:var(--red);border-color:var(--red)}
.file-list{max-height:160px;overflow-y:auto;padding:4px 0}
.file-row{
 display:flex;justify-content:space-between;align-items:center;
 padding:8px 16px;transition:.15s;
}
.file-row:hover{background:rgba(255,255,255,0.02)}
.file-name{font-size:12px;color:var(--muted);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
.file-name::before{content:'📄 '}
.rm-btn{background:transparent;border:none;color:var(--dim);cursor:pointer;font-size:13px;padding:2px 6px;border-radius:4px;transition:.15s}
.rm-btn:hover{color:var(--red);background:rgba(255,77,109,0.1)}
/* RUN BUTTON */
.run-btn{
 width:100%;
 background:linear-gradient(135deg,var(--accent),#ff8c5a);
 color:#fff;
 font-family:'Space Grotesk',sans-serif;font-weight:600;font-size:15px;
 border:none;border-radius:12px;padding:16px;
 cursor:pointer;transition:.2s;
 box-shadow:0 4px 20px rgba(255,107,53,0.3);
 position:relative;overflow:hidden;
}
.run-btn::before{
 content:'';position:absolute;top:0;left:-100%;width:100%;height:100%;
 background:linear-gradient(90deg,transparent,rgba(255,255,255,0.1),transparent);
 transition:.5s;
}
.run-btn:hover:not(:disabled)::before{left:100%}
.run-btn:hover:not(:disabled){transform:translateY(-1px);box-shadow:0 8px 30px rgba(255,107,53,0.4)}
.run-btn:disabled{opacity:0.4;cursor:not-allowed;transform:none}
/* PROGRESS */
.progress-card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:20px;display:none}
.progress-card.show{display:block}
.progress-label{font-size:12px;color:var(--muted);margin-bottom:10px;display:flex;justify-content:space-between}
.progress-track{width:100%;height:4px;background:var(--surface2);border-radius:2px;overflow:hidden}
.progress-bar{height:100%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:2px;width:0%;transition:width .4s ease}
/* LOG */
.log-card{background:var(--surface);border:1px solid var(--border);border-radius:16px;overflow:hidden;display:none}
.log-card.show{display:block}
.log-header{padding:12px 16px;border-bottom:1px solid var(--border);background:var(--surface2);font-size:12px;font-weight:600;color:var(--muted);letter-spacing:.5px;text-transform:uppercase}
.log-body{padding:16px;max-height:220px;overflow-y:auto;display:flex;flex-direction:column;gap:4px}
.log-line{font-size:12px;font-family:'SF Mono','Fira Code',monospace;line-height:1.7;display:flex;gap:8px;align-items:baseline}
.log-line::before{content:'›';color:var(--dim);flex-shrink:0}
.log-info{color:#8888bb}
.log-ok{color:var(--green)}
.log-err{color:var(--red)}
.log-zip{color:var(--accent)}
/* DONE */
.done-card{
 background:linear-gradient(135deg,rgba(0,229,160,0.08),rgba(0,229,160,0.03));
 border:1px solid rgba(0,229,160,0.2);
 border-radius:16px;padding:24px;
 display:none;text-align:center;
}
.done-card.show{display:block}
.done-icon{font-size:36px;margin-bottom:12px}
.done-title{font-family:'Space Grotesk',sans-serif;font-size:18px;font-weight:600;color:var(--green);margin-bottom:6px}
.done-sub{font-size:13px;color:var(--muted)}
.done-stats{display:flex;gap:20px;justify-content:center;margin-top:16px;flex-wrap:wrap}
.done-stat{
 background:rgba(0,229,160,0.08);border:1px solid rgba(0,229,160,0.15);
 border-radius:8px;padding:10px 20px;text-align:center;
}
.done-stat-val{font-family:'Space Grotesk',sans-serif;font-size:20px;font-weight:700;color:var(--green)}
.done-stat-label{font-size:11px;color:var(--muted);margin-top:2px}
/* FOOTER */
.footer{
 margin-top:40px;text-align:center;font-size:11px;color:var(--dim);
 padding:0 24px;
}
.footer a{color:var(--muted);text-decoration:none}
/* CREDITI */
.quota-body{padding:16px 20px;display:flex;flex-direction:column;gap:12px}
.quota-line{font-size:13px;color:var(--muted);line-height:1.5}
.quota-line strong{color:var(--text)}
.quota-ok{color:var(--green)!important}
.quota-err{color:var(--red)!important}
.code-row{display:flex;gap:8px;flex-wrap:wrap}
.code-input{flex:1;min-width:180px;background:var(--surface2);border:1px solid var(--border2);color:var(--text);
 border-radius:10px;padding:10px 12px;font-family:'SF Mono','Fira Code',monospace;font-size:13px;text-transform:uppercase}
.code-btn{background:var(--surface2);border:1px solid rgba(255,107,53,0.4);color:var(--accent2);border-radius:10px;
 padding:10px 16px;font-family:'Inter',sans-serif;font-size:13px;font-weight:600;cursor:pointer}
.code-btn:hover{background:rgba(255,107,53,0.1)}
.buy-link{color:var(--accent2);font-weight:600}
.paywall{background:rgba(255,107,53,0.08);border:1px solid rgba(255,107,53,0.3);border-radius:16px;padding:20px;
 display:none;font-size:14px;line-height:1.6}
.paywall.show{display:block}
.paywall a{color:var(--accent2);font-weight:600}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.pulsing{animation:pulse 1.2s infinite}
@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}
.fade-in{animation:fadeIn .3s ease}
</style>
</head>
<body>
<!-- HERO -->
<div class="hero">
<div class="hero-badge">Estrattore Fatture</div>
<h1>Da fattura a <span>Excel</span><br>in pochi secondi</h1>
<p class="hero-sub">Carica le tue fatture: il sito legge i dati e li mette in un file Excel pronto da usare.</p>
<ul class="hero-features">
<li><span><strong>Formati supportati:</strong> PDF (anche scansioni), XML FatturaPA, XML firmati .p7m e ZIP</span></li>
<li><span>Estrazione dei dati di qualsiasi fattura su Excel: fornitore, partita IVA, numero, data, imponibile, IVA, totale e righe di dettaglio</span></li>
<li><span>Foglio separato con i doppioni, per eliminare le righe ripetute che creano ridondanza</span></li>
</ul>
</div>
<!-- MAIN -->
<div class="main">
<!-- UPLOAD CARD -->
<div class="card">
<div class="card-header">
<div class="card-icon">📂</div>
<div>
<div class="card-title">Carica i tuoi file</div>
<div class="card-sub">PDF, XML, P7M o ZIP · anche più file insieme</div>
</div>
</div>
<label class="drop-zone" id="dropZone">
<input type="file" id="fileInput" accept=".xml,.zip,.pdf,.p7m" multiple style="display:none"/>
<div class="drop-icon-wrap">⬇</div>
<div class="drop-text">Trascina qui i file oppure clicca</div>
<div class="drop-sub">PDF · XML · P7M · ZIP</div>
</label>
<div class="or-row">oppure</div>
<label class="folder-btn">
<input type="file" id="folderInput" style="display:none" webkitdirectory directory/>
<div class="folder-icon">🗂</div>
<div class="folder-text">
<div class="folder-label">Carica una cartella intera</div>
<div class="folder-desc">Prende automaticamente tutti i PDF, XML, P7M e ZIP al suo interno</div>
</div>
</label>
</div>
<!-- FILE LIST -->
<div class="file-panel" id="filePanel">
<div class="file-panel-header">
<span class="file-count" id="fpCount">0 file selezionati</span>
<button class="clear-btn" onclick="clearAll()">✕ Svuota tutto</button>
</div>
<div class="file-list" id="fileList"></div>
</div>
<!-- CREDITI -->
<div class="card">
<div class="card-header">
<div class="card-icon">🎟</div>
<div>
<div class="card-title">I tuoi crediti</div>
<div class="card-sub">XML sempre gratis · {{ free_pdf }} PDF gratis al mese</div>
</div>
</div>
<div class="quota-body">
<div class="quota-line" id="quotaText">Caricamento…</div>
<div class="code-row">
<input class="code-input" id="codeInput" placeholder="FE-XXXX-XXXX-XXXX" autocomplete="off"/>
<button class="code-btn" onclick="saveCode()">Usa codice</button>
</div>
<div class="quota-line">Ti servono più PDF? <a class="buy-link" href="{{ pricing_url }}" target="_blank" rel="noopener">Vedi i pacchetti</a></div>
</div>
</div>
<!-- PAYWALL -->
<div class="paywall" id="paywall"></div>
<!-- RUN -->
<button class="run-btn" id="runBtn" onclick="handleRun()" disabled>
   ▶ &nbsp;Avvia Estrazione
</button>
<!-- PROGRESS -->
<div class="progress-card" id="progressWrap">
<div class="progress-label">
<span>Elaborazione in corso...</span>
<span id="progressPct">0%</span>
</div>
<div class="progress-track"><div class="progress-bar" id="progressBar"></div></div>
</div>
<!-- LOG -->
<div class="log-card" id="logCard">
<div class="log-header">Log di sistema</div>
<div class="log-body" id="logBox"></div>
</div>
<!-- DONE -->
<div class="done-card" id="doneBanner">
<div class="done-icon">✅</div>
<div class="done-title">Estrazione completata!</div>
<div class="done-sub">Il file <strong>estrazione_fatture.xlsx</strong> è stato scaricato</div>
<div class="done-stats" id="doneStats"></div>
</div>
<div class="footer">
   I file vengono elaborati sul server e non vengono salvati. &nbsp;·&nbsp;
<a href="#">estrattore-fatture.onrender.com</a>
</div>
</div>
<script>
let selectedFiles = [];
const PRICING_URL = {{ pricing_url|tojson }};
function getCode(){ try{ return localStorage.getItem('fe_codice')||''; }catch(e){ return window._feCode||''; } }
function setCode(c){ try{ localStorage.setItem('fe_codice',c); }catch(e){ window._feCode=c; } }
async function refreshQuota(){
 const el=document.getElementById('quotaText');
 try{
   const r=await fetch('/quota',{headers:{'X-Codice':getCode()}});
   const q=await r.json();
   const parts=[];
   parts.push('PDF gratis rimasti questo mese: <strong>'+q.gratis_rimasti+' di '+q.gratis_mese+'</strong>');
   el.className='quota-line';
   if(q.codice){
     const c=q.codice;
     if(!c.valido){
       parts.push('<span class="quota-err">Codice '+(c.scaduto?'scaduto':'esaurito')+'</span>');
     }else if(c.illimitato){
       parts.push('<span class="quota-ok">Codice attivo: PDF illimitati</span>');
     }else{
       parts.push('<span class="quota-ok">Codice attivo: '+c.rimasti+' PDF rimasti'+(c.scadenza?' (scade il '+c.scadenza.split('-').reverse().join('/')+')':'')+'</span>');
     }
   }else if(getCode()){
     parts.push('<span class="quota-err">Codice non riconosciuto</span>');
   }
   el.innerHTML=parts.join('<br>');
 }catch(e){ el.textContent='Impossibile leggere i crediti in questo momento.'; }
}
function saveCode(){
 const c=document.getElementById('codeInput').value.trim().toUpperCase();
 setCode(c); refreshQuota();
}
document.getElementById('codeInput').value=getCode();
refreshQuota();
function showPaywall(msg){
 const p=document.getElementById('paywall');
 p.textContent='';
 const t=document.createElement('div'); t.textContent=msg; p.appendChild(t);
 const a=document.createElement('a'); a.href=PRICING_URL; a.target='_blank'; a.rel='noopener';
 a.textContent='Acquista un pacchetto PDF →'; p.appendChild(document.createElement('br')); p.appendChild(a);
 p.classList.add('show');
}
function addFiles(newFiles) {
 const valid = Array.from(newFiles).filter(f =>
   ['xml','zip','pdf','p7m'].some(ext => f.name.toLowerCase().endsWith('.'+ext))
 );
 const existing = new Set(selectedFiles.map(f => f.name+f.size));
 valid.forEach(f => { if(!existing.has(f.name+f.size)) selectedFiles.push(f); });
 render();
}
function removeFile(i){ selectedFiles.splice(i,1); render(); }
function clearAll(){
 selectedFiles=[];render();
 document.getElementById('logCard').classList.remove('show');
 document.getElementById('doneBanner').classList.remove('show');
}
function render(){
 const panel=document.getElementById('filePanel');
 const list=document.getElementById('fileList');
 const count=document.getElementById('fpCount');
 const btn=document.getElementById('runBtn');
 if(!selectedFiles.length){panel.classList.remove('show');btn.disabled=true;return;}
 panel.classList.add('show');btn.disabled=false;
 count.textContent=selectedFiles.length+' file selezionati';
 list.innerHTML=selectedFiles.map((f,i)=>`
<div class="file-row fade-in">
<span class="file-name">${f.name}</span>
<button class="rm-btn" onclick="removeFile(${i})">✕</button>
</div>`).join('');
}
document.getElementById('fileInput').onchange=e=>addFiles(e.target.files);
document.getElementById('folderInput').onchange=e=>addFiles(e.target.files);
const dz=document.getElementById('dropZone');
dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('over');});
dz.addEventListener('dragleave',()=>dz.classList.remove('over'));
dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('over');addFiles(e.dataTransfer.files);});
dz.addEventListener('click',()=>document.getElementById('fileInput').click());
function log(msg,type='info'){
 const box=document.getElementById('logBox');
 document.getElementById('logCard').classList.add('show');
 const d=document.createElement('div');
 d.className='log-line log-'+type+' fade-in';
 d.textContent=msg;
 box.appendChild(d);
 box.scrollTop=box.scrollHeight;
}
function setProgress(pct){
 document.getElementById('progressBar').style.width=pct+'%';
 document.getElementById('progressPct').textContent=pct+'%';
}
async function handleRun(){
 if(!selectedFiles.length) return;
 const btn=document.getElementById('runBtn');
 btn.disabled=true;btn.classList.add('pulsing');
 btn.innerHTML='⏳ &nbsp;Elaborazione in corso…';
 document.getElementById('logBox').innerHTML='';
 document.getElementById('logCard').classList.add('show');
 document.getElementById('doneBanner').classList.remove('show');
 document.getElementById('progressWrap').classList.add('show');
 setProgress(20);
 const fd=new FormData();
 selectedFiles.forEach(f=>fd.append('files',f));
 try{
   log('Invio '+selectedFiles.length+' file al server…','info');
   document.getElementById('paywall').classList.remove('show');
   const resp=await fetch('/process',{method:'POST',body:fd,headers:{'X-Codice':getCode()}});
   setProgress(85);
   if(!resp.ok){
     let err={};
     try{ err=await resp.json(); }catch(e){}
     if(resp.status===402){
       log('Crediti PDF insufficienti','err');
       showPaywall(err.error||'Crediti PDF esauriti.');
     }else{
       log('Errore: '+(err.error||resp.statusText),'err');
     }
     setProgress(0);
     return;
   }
   const fatture=resp.headers.get('X-Rows-Fatture')||'?';
   const dups=resp.headers.get('X-Rows-Duplicati')||'?';
   log('Elaborazione completata con successo','ok');
   log('Fatture lette: '+fatture,'ok');
   log('Fatture doppie spostate nel foglio Duplicati: '+dups,'ok');
   try{ JSON.parse(decodeURIComponent(resp.headers.get('X-Avvisi')||'%5B%5D')).forEach(a=>log('Attenzione: '+a,'err')); }catch(e){}
   setProgress(100);
   const blob=await resp.blob();
   const url=URL.createObjectURL(blob);
   const a=document.createElement('a');
   a.href=url;a.download='estrazione_fatture.xlsx';
   document.body.appendChild(a);a.click();
   document.body.removeChild(a);URL.revokeObjectURL(url);
   const banner=document.getElementById('doneBanner');
   document.getElementById('doneStats').innerHTML=`
<div class="done-stat"><div class="done-stat-val">${fatture}</div><div class="done-stat-label">Fatture lette</div></div>
<div class="done-stat"><div class="done-stat-val">${dups}</div><div class="done-stat-label">Doppioni trovati</div></div>`;
   banner.classList.add('show','fade-in');
   refreshQuota();
 }catch(e){
   log('Errore di rete: '+e.message,'err');
 }finally{
   btn.disabled=false;btn.classList.remove('pulsing');
   btn.innerHTML='▶ &nbsp;Avvia Estrazione';
 }
}
</script>
</body>
</html>'''

ADMIN_HTML = '''<!DOCTYPE html>
<html lang="it"><head><meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Admin codici</title>
<style>
body{font-family:system-ui,sans-serif;background:#0a0a0f;color:#f0f0f5;max-width:760px;margin:0 auto;padding:24px 16px}
h1{font-size:22px;margin-bottom:16px}h2{font-size:16px;margin:28px 0 10px;color:#ff9a6c}
form{display:grid;gap:10px;background:#111118;border:1px solid #ffffff18;border-radius:12px;padding:16px}
label{font-size:13px;color:#a0a0c0;display:grid;gap:4px}
input,select{background:#1a1a24;color:#fff;border:1px solid #ffffff22;border-radius:8px;padding:9px;font-size:14px}
button{background:#ff6b35;color:#fff;border:none;border-radius:8px;padding:11px;font-weight:600;font-size:14px;cursor:pointer}
.nuovo{background:rgba(0,229,160,.1);border:1px solid rgba(0,229,160,.3);border-radius:12px;padding:16px;margin-bottom:16px}
.nuovo code{font-size:22px;color:#00e5a0;user-select:all}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{text-align:left;padding:7px 6px;border-bottom:1px solid #ffffff12}
th{color:#7070a0;font-weight:500}td code{color:#ff9a6c}
.wrap{overflow-x:auto}
</style></head><body>
<h1>Codici di sblocco</h1>
{% if nuovo %}<div class="nuovo">Codice creato, copialo e mandalo al cliente:<br><code>{{ nuovo }}</code></div>{% endif %}
<form method="post">
<label>Tipo
<select name="tipo">
<option value="pacchetto">Pacchetto PDF (nessuna scadenza)</option>
<option value="abbonamento">Abbonamento (scade dopo i giorni indicati)</option>
<option value="illimitato">Illimitato (uso personale)</option>
</select></label>
<label>PDF inclusi<input name="crediti" type="number" min="1" value="100"/></label>
<label>Giorni di validità (solo abbonamento)<input name="giorni" type="number" min="1" value="31"/></label>
<label>Nota (es. nome cliente o ID transazione PayPal)<input name="nota" maxlength="200"/></label>
<button type="submit">Crea codice</button>
</form>
<h2>Ultimi codici</h2>
<div class="wrap"><table>
<tr><th>Codice</th><th>PDF</th><th>Usati</th><th>Scadenza</th><th>Nota</th><th>Creato</th></tr>
{% for c in codici %}<tr><td><code>{{ c[0] }}</code></td>
<td>{{ "illimitati" if c[1] == -1 else c[1] }}</td><td>{{ c[2] }}</td>
<td>{{ c[3] or "—" }}</td><td>{{ c[4] or "" }}</td><td>{{ c[5][:16].replace("T", " ") }}</td></tr>
{% else %}<tr><td colspan="6">Nessun codice ancora.</td></tr>{% endfor %}
</table></div>
</body></html>'''

def _admin_ok():
   auth = request.authorization
   return bool(ADMIN_PASSWORD) and auth is not None and \
       hmac.compare_digest((auth.password or "").encode(), ADMIN_PASSWORD.encode())

@app.route('/admin', methods=['GET', 'POST'])
def admin():
   if not ADMIN_PASSWORD:
       return "Pagina admin disattivata: imposta ADMIN_PASSWORD su Render.", 404
   if not _admin_ok():
       return Response("Accesso riservato", 401,
                       {"WWW-Authenticate": 'Basic realm="Admin codici"'})
   nuovo = None
   if request.method == 'POST':
       tipo = request.form.get('tipo', 'pacchetto')
       nota = request.form.get('nota', '')[:200]
       try:
           n = max(1, int(request.form.get('crediti', '100')))
           giorni = max(1, int(request.form.get('giorni', '31')))
       except ValueError:
           return "Valori non validi", 400
       if tipo == 'illimitato':
           nuovo = crediti.crea_codice(crediti.ILLIMITATO, None, nota)
       elif tipo == 'abbonamento':
           nuovo = crediti.crea_codice(n, giorni, nota)
       else:
           nuovo = crediti.crea_codice(n, None, nota)
   return render_template_string(ADMIN_HTML, nuovo=nuovo, codici=crediti.elenco_codici())

@app.route('/')
def index():
   return render_template_string(HTML, free_pdf=crediti.FREE_PDF_MONTH, pricing_url=PRICING_URL)

def conta_pdf(name, data, depth=0):
   """Conta i PDF in un file caricato, compresi quelli dentro gli ZIP."""
   if name.endswith('.pdf'):
       return 1
   if name.endswith('.zip') and depth < 5:
       try:
           with zipfile.ZipFile(io.BytesIO(data)) as zf:
               totale = 0
               for n in zf.namelist():
                   low = n.lower()
                   if low.startswith('__macosx'):
                       continue
                   if low.endswith('.pdf'):
                       totale += 1
                   elif low.endswith('.zip'):
                       totale += conta_pdf(low, zf.read(n), depth + 1)
               return totale
       except zipfile.BadZipFile:
           return 0
   return 0

@app.route('/quota')
def quota():
   client = crediti.client_id(request)
   info = crediti.info_codice(request.headers.get('X-Codice', ''))
   return jsonify({
       "gratis_mese": crediti.FREE_PDF_MONTH,
       "gratis_rimasti": crediti.gratis_rimasti(client),
       "codice": info,
   })

@app.route('/process', methods=['POST'])
def process():
   files = request.files.getlist('files')
   if not files:
       return jsonify({"error": "Nessun file ricevuto"}), 400
   caricati = [(f.filename.lower(), f.read()) for f in files]
   n_pdf = sum(conta_pdf(name, data) for name, data in caricati)
   client = crediti.client_id(request)
   codice = request.headers.get('X-Codice', '')
   if n_pdf:
       disp = crediti.disponibili(client, codice)
       if disp is not None and n_pdf > disp:
           return jsonify({
               "error": f"Hai caricato {n_pdf} PDF ma te ne restano {disp}. "
                        f"Gli XML sono sempre gratis; per i PDF oltre la quota serve un codice.",
               "quota_esaurita": True,
               "pdf_richiesti": n_pdf,
               "pdf_disponibili": disp,
               "pricing_url": PRICING_URL,
           }), 402
   fatture, avvisi = [], []
   for name, data in caricati:
       fatture += estrazione.leggi_file(name, data, avvisi)
   for a in avvisi:
       print("Avviso:", a)
   if not fatture:
       dettaglio = ("; ".join(avvisi[:5])) if avvisi else "formati non riconosciuti"
       return jsonify({"error": f"Nessuna fattura letta dai file caricati ({dettaglio})"}), 422
   crediti.addebita(client, codice, n_pdf)
   excel_bytes, n_fatture, n_dups = excel.crea_excel(fatture)
   response = send_file(
       excel_bytes,
       mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
       as_attachment=True,
       download_name='estrazione_fatture.xlsx'
   )
   response.headers['X-Rows-Fatture']   = str(n_fatture)
   response.headers['X-Rows-Duplicati'] = str(n_dups)
   response.headers['X-Avvisi'] = urllib.parse.quote(json.dumps(avvisi[:20], ensure_ascii=False))
   response.headers['Access-Control-Expose-Headers'] = 'X-Rows-Fatture, X-Rows-Duplicati, X-Avvisi'
   return response

if __name__ == '__main__':
   port = int(os.environ.get('PORT', 5000))
   app.run(host='0.0.0.0', port=port)
