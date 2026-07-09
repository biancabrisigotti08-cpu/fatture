"""
Estrattore Fatture Web - Flask Backend
"""
import zipfile
import io
import re
import os
import xml.etree.ElementTree as ET
from pathlib import Path
from flask import Flask, request, send_file, jsonify, render_template_string
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from pypdf import PdfReader
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500MB max upload
HEADER_BG   = "FF6B35"
HEADER_FONT = "FFFFFF"
IMPORTO_BOLLO = 2.00

# ─── Parser XML FatturaPA ──────────────────────────────────────────────────────
def find_text(element, *tags):
   for tag in tags:
       found = element.find('.//' + tag)
       if found is not None and found.text:
           return found.text.strip()
       for child in element.iter():
           local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
           if local == tag and child.text:
               return child.text.strip()
   return ""

def parse_xml_fattura(xml_bytes):
   try:
       root = ET.fromstring(xml_bytes)
   except ET.ParseError as e:
       raise ValueError(f"XML non valido: {e}")
   cedente_block = None
   cessionario_block = None
   for child in root.iter():
       local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
       if local == "CedentePrestatore" and cedente_block is None:
           cedente_block = child
       if local == "CessionarioCommittente" and cessionario_block is None:
           cessionario_block = child
   cedente     = find_text(cedente_block, "Denominazione")     if cedente_block     else ""
   cessionario = find_text(cessionario_block, "Denominazione") if cessionario_block else ""
   numero_documento = find_text(root, "Numero")
   data_documento   = find_text(root, "Data")
   righe = []
   for linea in root.iter():
       local = linea.tag.split('}')[-1] if '}' in linea.tag else linea.tag
       if local != "DettaglioLinee":
           continue
       descrizione_raw = find_text(linea, "Descrizione")
       prezzo_totale   = find_text(linea, "PrezzoTotale")
       try:
           if abs(float(prezzo_totale)) == IMPORTO_BOLLO:
               continue
       except (ValueError, TypeError):
           pass
       telaio = ""
       descrizione = ""
       if descrizione_raw:
           m = re.match(r'^(\S+)\s+RMK\S+\s+(.+)$', descrizione_raw.strip(), re.IGNORECASE)
           if m:
               telaio      = m.group(1).strip()
               descrizione = m.group(2).strip()
           else:
               parts = descrizione_raw.strip().split(None, 1)
               telaio      = parts[0] if parts else ""
               descrizione = parts[1] if len(parts) > 1 else ""
       righe.append({
           "telaio": telaio,
           "descrizione": descrizione,
           "prezzo_totale": prezzo_totale,
       })
   return {
       "cedente": cedente,
       "cessionario": cessionario,
       "numero_documento": numero_documento,
       "data_documento": data_documento,
       "righe": righe,
   }

# ─── Parser PDF — formato generico (PSA + Romana Diesel + altri) ──────────────
def parse_pdf_fattura(pdf_bytes):
   text = ""
   reader = PdfReader(io.BytesIO(pdf_bytes))
   for page in reader.pages:
       t = page.extract_text()
       if t:
           text += t + "\n"
   lines = text.splitlines()
   # Cedente
   cedente = ""
   for i, l in enumerate(lines):
       if re.search(r'cedente|prestatore|fornitore', l, re.IGNORECASE):
           for j in range(i+1, min(i+6, len(lines))):
               nl = lines[j].strip()
               if nl and not re.match(r'^[A-Z\s/()]{6,}$', nl):
                   cedente = nl
                   break
           if cedente:
               break
   if not cedente:
       for l in lines[:10]:
           nl = l.strip()
           if nl and len(nl) > 3:
               cedente = nl
               break
   # Cessionario
   cessionario = ""
   for i, l in enumerate(lines):
       if re.search(r'cessionario|committente|spett', l, re.IGNORECASE):
           for j in range(i+1, min(i+6, len(lines))):
               nl = lines[j].strip()
               if nl and not re.match(r'^[A-Z\s/()]{6,}$', nl):
                   cessionario = nl
                   break
           if cessionario:
               break
   # Numero documento
   numero_documento = ""
   # Formato VW: "TD01 (fattura) 000866019 16980504532 30-03-2026"
   m = re.search(r'TD0\d\s*\([^)]+\)\s*(\d{6,})', text, re.IGNORECASE)
   if m:
       numero_documento = m.group(1).strip()
   else:
       # Formato VW alternativo: "NUMERO DOCUMENTO ART. 73 NUMERO DOCUMENTO\n...000866019"
       m = re.search(r'NUMERO\s+DOCUMENTO(?:\s+ART[.\s]+\d+)?\s*[\n\r]+\s*(?:ART[.\s]+\d+\s*[\n\r]+\s*)?(\d{6,})', text, re.IGNORECASE)
       if m:
           numero_documento = m.group(1).strip()
   if not numero_documento:
       # Formato PSA: "NUMERO DOCUMENTO\n1181358498"
       m = re.search(r'NUMERO\s+DOCUMENTO\s*[\n\r]+\s*(\S+)', text, re.IGNORECASE)
       if m:
           val = m.group(1).strip()
           if not re.match(r'^[A-Z]+$', val, re.IGNORECASE):
               numero_documento = val
   if not numero_documento:
       # Formato Romana Diesel: "Numero\nG000617"
       m = re.search(r'\bNumero\b\s*[\n\r]+\s*([A-Z0-9]+)', text, re.IGNORECASE)
       if m:
           numero_documento = m.group(1).strip()
   if not numero_documento:
       m = re.search(r'\bNumero\b\s+([A-Z0-9]{4,})', text, re.IGNORECASE)
       if m:
           numero_documento = m.group(1).strip()
   # Data documento
   data_documento = ""
   m = re.search(r'DATA\s+DOCUMENTO\s*[\n\r\s]+(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})', text, re.IGNORECASE)
   if m:
       data_documento = m.group(1).strip()
   else:
       m = re.search(r'\bdata\b\s*[\n\r]+\s*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})', text, re.IGNORECASE)
       if m:
           data_documento = m.group(1).strip()
       else:
           m = re.search(r'(\d{1,2}[-/]\d{2}[-/]\d{4})', text)
           if m:
               data_documento = m.group(1).strip()
   is_psa_format     = bool(re.search(r'RMK\w+', text))
   is_romana_format  = bool(re.search(r'RIF\.TARGA', text, re.IGNORECASE))
   is_vw_format      = bool(re.search(r'Tipo dato:TELAIO', text, re.IGNORECASE))
   righe = []
   if is_vw_format:
       # ── Parser Volkswagen ──
       # Estrae telaio, targa, prezzi e descrizioni dal testo
       telaio_pattern = re.compile(r'Tipo\s+dato:TELAIO[\n\r]+Rif\.\s*testo:(\S+)', re.IGNORECASE)
       targa_pattern  = re.compile(r'Tipo\s+dato:TARGA[\n\r]+Rif\.\s*testo:(\S+)', re.IGNORECASE)
       prezzo_pattern = re.compile(r'[\d,.]+ N[12T] ([\d]{1,3}(?:[.,]\d{3})*[.,]\d{2})', re.IGNORECASE)
       desc_pattern   = re.compile(r'([A-Z][A-Z\s\-]+(?:DANNI|PENALE|PERIZIA|ESUBERO|FORFAIT|TAGLIANDO|PLAFOND)[A-Z\s\-]*)', re.IGNORECASE)
       telai  = [m.group(1) for m in telaio_pattern.finditer(text)]
       targhe = [m.group(1) for m in targa_pattern.finditer(text)]
       prezzi = [m.group(1).replace('.','').replace(',','.') for m in prezzo_pattern.finditer(text)]
       descs  = [m.group(1).strip() for m in desc_pattern.finditer(text)]
       # Abbina targa e telaio per indice.
       # Parser basato su posizione nel testo — gestisce pagina a capo
       # Ogni blocco è: N. riga → ADDEBITO... → Tipo dato:TELAIO → Rif. testo:XXX → Tipo dato:TARGA → Rif. testo:XXX → prezzo N1 prezzo
       # Anche quando il prezzo finisce sulla pagina successiva
       telaio_matches = [(m.group(1), m.start()) for m in telaio_pattern.finditer(text)]
       targa_matches  = [(m.group(1), m.start()) for m in targa_pattern.finditer(text)]
       prezzo_matches = [(m.group(1).replace('.','').replace(',','.'), m.start()) for m in prezzo_pattern.finditer(text)]
       desc_matches   = [(m.group(1).strip(), m.start()) for m in desc_pattern.finditer(text)]
       # Costruisci blocchi basati sui telai (uno per riga fattura)
       for t_idx, (telaio_val, t_pos) in enumerate(telaio_matches):
           # Prossima posizione di telaio (per delimitare il blocco)
           next_t_pos = telaio_matches[t_idx + 1][1] if t_idx + 1 < len(telaio_matches) else len(text)
           # Targa: cerca dopo il telaio corrente e prima del prossimo telaio
           targa_val = ""
           for tg_val, tg_pos in targa_matches:
               if t_pos < tg_pos < next_t_pos:
                   targa_val = tg_val
                   break
           # Prezzo: cerca dopo il telaio corrente, anche oltre il prossimo telaio (pagina a capo)
           # ma prima del telaio successivo + 2000 caratteri
           prezzo_val = ""
           search_end = next_t_pos + 2000
           for p_val, p_pos in prezzo_matches:
               if t_pos < p_pos < search_end:
                   try:
                       if abs(float(p_val)) != IMPORTO_BOLLO:
                           prezzo_val = p_val
                           break
                   except (ValueError, TypeError):
                       pass
           # Descrizione: cerca prima del telaio corrente (nella stessa riga)
           desc_val = ""
           for d_val, d_pos in reversed(desc_matches):
               if d_pos < t_pos:
                   desc_val = d_val
                   break
           # Pulisci descrizione: prendi solo la parte prima del trattino
           if desc_val and '-' in desc_val:
               desc_val = desc_val.split('-')[0].strip()
           if telaio_val or prezzo_val:
               righe.append({
                   "targa":         targa_val,
                   "telaio":        telaio_val,
                   "descrizione":   desc_val,
                   "prezzo_totale": prezzo_val,
               })
   elif is_psa_format:
       row_start_re = re.compile(r'^\s*(\d+)\s*$')
       desc_re = re.compile(r'([A-Z0-9]{8,})\s+(RMK\S+)\s+(.*)', re.IGNORECASE)
       i = 0
       while i < len(lines):
           line = lines[i].strip()
           if row_start_re.match(line):
               block_lines = []
               j = i + 1
               while j < len(lines):
                   next_line = lines[j].strip()
                   if row_start_re.match(next_line) and next_line != line:
                       break
                   block_lines.append(next_line)
                   j += 1
               telaio = ""
               descrizione = ""
               for bl in block_lines:
                   md = desc_re.search(bl)
                   if md:
                       telaio = md.group(1).strip()
                       descrizione = md.group(3).strip()
                       break
               prezzo_totale = ""
               for bl in reversed(block_lines):
                   if re.search(r'\bN\d\b', bl):
                       nums = re.findall(r'[\d]+[.,][\d]+', bl)
                       if nums:
                           prezzo_totale = nums[-1].replace(',', '.')
                       break
               try:
                   if prezzo_totale and abs(float(prezzo_totale)) == IMPORTO_BOLLO:
                       i = j
                       continue
               except (ValueError, TypeError):
                   pass
               if telaio or prezzo_totale:
                   righe.append({"targa": "", "telaio": telaio, "descrizione": descrizione, "prezzo_totale": prezzo_totale})
               i = j
           else:
               i += 1
   elif is_romana_format:
       targa_pattern = re.compile(r'RIF\.TARGA\s+([A-Z0-9]+)', re.IGNORECASE)
       current_targa = ""
       for i, line in enumerate(lines):
           mt = targa_pattern.search(line)
           if mt:
               current_targa = mt.group(1).strip()
               continue
           line_stripped = line.strip()
           if not line_stripped:
               continue
           if re.search(r'bollo', line_stripped, re.IGNORECASE):
               continue
           if re.search(r'imponibile|totale fattura|pagamento|p\.i\.|partita|sede|tel|fax|bonifico|iva|aliq|q\.t[\xc3\xa0a]|prezzo unitario|importo netto|descrizione', line_stripped, re.IGNORECASE):
               continue
           m = re.match(r'^([A-Z][A-Z\s]+?)\s+([\d]{1,3}(?:\.\d{3})*,\d{2})\s*(?:\d+)?$', line_stripped)
           if m and current_targa:
               desc = m.group(1).strip()
               importo = m.group(2).replace('.', '').replace(',', '.')
               try:
                   val = float(importo)
                   if val == IMPORTO_BOLLO:
                       continue
               except ValueError:
                   continue
               righe.append({"targa": current_targa, "telaio": "", "descrizione": desc, "prezzo_totale": importo})
   else:
       for line in lines:
           line_stripped = line.strip()
           if re.search(r'bollo', line_stripped, re.IGNORECASE):
               continue
           m = re.match(r'^(.+?)\s+([\d]{1,3}(?:\.\d{3})*,\d{2})\s*(?:\d+)?$', line_stripped)
           if m:
               desc = m.group(1).strip()
               importo = m.group(2).replace('.', '').replace(',', '.')
               try:
                   val = float(importo)
                   if val == IMPORTO_BOLLO:
                       continue
               except ValueError:
                   continue
               if len(desc) > 2:
                   righe.append({"targa": "", "telaio": "", "descrizione": desc, "prezzo_totale": importo})
   return {"cedente": cedente, "cessionario": cessionario,
           "numero_documento": numero_documento, "data_documento": data_documento, "righe": righe}

def process_pdf_bytes(pdf_bytes, all_rows):
   try:
       # Log testo grezzo per debug
       text_preview = ""
       reader = PdfReader(io.BytesIO(pdf_bytes))
       for page in reader.pages:
           t = page.extract_text()
           if t:
               text_preview += t + "\n"
       print("=== PDF TEXT (primi 1500 char) ===")
       print(repr(text_preview[:1500]))
       print("==================================")
       fattura = parse_pdf_fattura(pdf_bytes)
       righe   = fattura.pop("righe", [])
       print(f"Righe estratte: {len(righe)}")
       for r in righe:
           all_rows.append({**fattura, **r})
       return len(righe)
   except Exception as e:
       print(f"Errore PDF: {e}")
       return 0

def process_xml_bytes(xml_bytes, all_rows):
   try:
       fattura = parse_xml_fattura(xml_bytes)
       righe   = fattura.pop("righe", [])
       for r in righe:
           all_rows.append({**fattura, **r})
       return len(righe)
   except Exception as e:
       return 0

def process_zip_bytes(zip_bytes, all_rows, depth=0):
   try:
       with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
           entries = zf.namelist()
           xml_entries = [n for n in entries if n.lower().endswith('.xml')
                         and not n.startswith('__MACOSX')
                         and not n.lower().endswith('signature.xml')]
           pdf_entries = [n for n in entries if n.lower().endswith('.pdf')
                         and not n.startswith('__MACOSX')]
           zip_entries = [n for n in entries if n.lower().endswith('.zip')
                         and not n.startswith('__MACOSX')]
           for xml_name in xml_entries:
               process_xml_bytes(zf.read(xml_name), all_rows)
           for pdf_name in pdf_entries:
               process_pdf_bytes(zf.read(pdf_name), all_rows)
           for zip_name in zip_entries:
               process_zip_bytes(zf.read(zip_name), all_rows, depth + 1)
   except zipfile.BadZipFile:
       pass

# ─── Excel builder ────────────────────────────────────────────────────────────
def valida_targa(targa: str) -> str:
   """Valida e restituisce la targa se valida, altrimenti stringa vuota.
   Formato valido: 2 lettere + 3 cifre + 2 lettere (es. AA000AA), 7 caratteri totali."""
   t = targa.strip().upper()
   if len(t) != 7:
       return ""
   if re.match(r'^[A-Z]{2}\d{3}[A-Z]{2}$', t):
       return t
   return ""

def valida_telaio(telaio: str) -> str:
   """Valida e restituisce il telaio se valido, altrimenti stringa vuota.
   Formato valido: 17 caratteri, inizia con lettere, finisce con numeri."""
   t = telaio.strip().upper()
   if len(t) != 17:
       return ""
   # Inizia con almeno 2 lettere e finisce con almeno 4 cifre
   if re.match(r'^[A-Z]{2,}.*\d{4,}$', t):
       return t
   return ""

def valida_numero_documento(numero: str) -> str:
   """Restituisce il numero documento se valido (numerico o alfanumerico tipo G000617)."""
   n = numero.strip()
   # Accetta numeri puri (000866019) o alfanumerici (G000617)
   # Esclude parole come "DOCUMENTO", "ART", ecc.
   if re.match(r'^[A-Z0-9]{4,}$', n, re.IGNORECASE) and not re.match(r'^[A-Z]+$', n, re.IGNORECASE):
       return n
   return ""

def to_float(val):
   try:
       return float(str(val).replace(',', '.'))
   except (ValueError, AttributeError):
       return None

def get_categoria(descrizione: str) -> str:
   """Estrae la categoria dalla descrizione, solo parole chiave specifiche."""
   d = descrizione.lower()
   if "forfait" in d:
       return "Forfait"
   if "over plafond" in d:
       return "Over Plafond"
   if "km eccedenti" in d or "esubero km" in d:
       return "Km Eccedenti"
   if "eam" in d:
       return "EAM"
   if "tagliando" in d:
       return "Tagliando"
   if "perizia" in d:
       return "Perizia"
   return ""

def build_row(row):
   descrizione = row.get("descrizione", "")
   targa  = valida_targa(row.get("targa", ""))
   telaio = valida_telaio(row.get("telaio", ""))
   return [
       row.get("cedente", ""),
       row.get("cessionario", ""),
       valida_numero_documento(row.get("numero_documento", "")),
       row.get("data_documento", ""),
       targa,
       telaio,
       to_float(row.get("prezzo_totale", "")),
       descrizione,
       get_categoria(descrizione),
   ]

def write_header(ws, headers, col_widths):
   header_fill = PatternFill("solid", fgColor=HEADER_BG)
   header_font = Font(bold=True, color=HEADER_FONT)
   for c, (h, w) in enumerate(zip(headers, col_widths), start=1):
       cell = ws.cell(row=1, column=c, value=h)
       cell.fill = header_fill
       cell.font = header_font
       cell.alignment = Alignment(horizontal="center", vertical="center")
       ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
   ws.row_dimensions[1].height = 20
   ws.freeze_panes = "A2"

def build_excel(all_rows):
   headers    = ["Cedente", "Cessionario", "N. Documento", "Data Documento",
                 "Targa", "Telaio", "Prezzo Totale (€)",
                 "Descrizione (Forfait/Over Plafond/ETM/EAM/KM…)",
                 "Descrizione"]
   col_widths = [30, 30, 18, 16, 14, 22, 18, 45, 18]
   CATEGORIE = ["EAM", "Forfait", "Km Eccedenti", "Over Plafond", "Tagliando"]
   # Filtra BOLLO
   filtered = [r for r in all_rows
               if "bollo" not in r.get("descrizione", "").lower()
               and "bollo" not in r.get("telaio", "").lower()]
   # Duplicati
   seen, uniq, dups = {}, [], []
   for r in filtered:
       key = (r.get("cedente",""), r.get("cessionario",""),
              r.get("numero_documento",""), r.get("data_documento",""),
              r.get("telaio",""), r.get("prezzo_totale",""), r.get("descrizione",""))
       if key in seen:
           dups.append(r)
       else:
           seen[key] = True
           uniq.append(r)
   wb = openpyxl.Workbook()
   # ── Foglio 1: Fatture ──
   ws1 = wb.active
   ws1.title = "Fatture"
   write_header(ws1, headers, col_widths)
   for row in uniq:
       ws1.append(build_row(row))
   # ── Foglio 2: Duplicati ──
   ws2 = wb.create_sheet(title="Duplicati")
   write_header(ws2, headers, col_widths)
   for row in dups:
       ws2.append(build_row(row))
   out = io.BytesIO()
   wb.save(out)
   out.seek(0)
   return out, len(uniq), len(dups)

# ─── HTML Template ────────────────────────────────────────────────────────────
HTML = '''<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Estrattore Fatture</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@700;800&display=swap" rel="stylesheet"/>
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{--orange:#ff6b35;--bg:#0e0e0e;--bg2:#131313;--border:#2a2a2a;--text:#e8e8e8;--green:#5aff9a;--red:#ff5a5a}
body{background:var(--bg);color:var(--text);font-family:"DM Mono",monospace;min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:48px 24px 64px;gap:24px}
::-webkit-scrollbar{width:6px}::-webkit-scrollbar-thumb{background:var(--orange);border-radius:3px}
.badge{background:var(--orange);color:var(--bg);font-family:"Syne",sans-serif;font-weight:700;font-size:11px;letter-spacing:2px;padding:4px 12px;border-radius:2px;margin-bottom:12px;display:inline-block}
h1{font-family:"Syne",sans-serif;font-weight:800;font-size:clamp(32px,5vw,58px);color:#fff;letter-spacing:-1px;line-height:1}
.sub{margin-top:10px;font-size:13px;color:#888}
.header{text-align:center;margin-bottom:8px}
.upload-area{width:100%;max-width:620px;display:flex;flex-direction:column}
.drop-zone{border:1.5px dashed var(--border);border-radius:8px 8px 0 0;padding:32px 24px;cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:8px;transition:.2s;background:rgba(255,255,255,.02);text-align:center}
.drop-zone:hover,.drop-zone.over{border-color:var(--orange);background:rgba(255,107,53,.06)}
.drop-icon{font-size:30px;opacity:.5}
.drop-text{font-size:14px;color:#ccc}
.drop-sub{font-size:11px;color:#444}
.or-div{background:#1a1a1a;padding:8px;text-align:center;font-size:11px;color:#444;letter-spacing:1px;border-left:1.5px solid var(--border);border-right:1.5px solid var(--border)}
.folder-btn{border:1.5px dashed var(--border);border-top:none;border-radius:0 0 8px 8px;padding:20px 24px;cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:4px;transition:.2s;background:rgba(255,255,255,.02);color:#aaa;font-size:14px;font-family:"DM Mono",monospace}
.folder-btn:hover{border-color:var(--orange);background:rgba(255,107,53,.06);color:var(--orange)}
.folder-sub{font-size:11px;color:#444}
.file-panel{width:100%;max-width:620px;background:var(--bg2);border:1px solid var(--border);border-radius:8px;overflow:hidden;display:none}
.file-panel.show{display:block}
.fp-header{display:flex;justify-content:space-between;align-items:center;padding:10px 16px;border-bottom:1px solid #1e1e1e;background:#161616}
.fp-count{font-size:12px;color:var(--orange)}
.clear-btn{background:transparent;border:1px solid var(--border);color:#555;font-size:11px;padding:3px 10px;border-radius:4px;cursor:pointer;font-family:"DM Mono",monospace}
.clear-btn:hover{color:var(--red);border-color:var(--red)}
.file-list{max-height:160px;overflow-y:auto;padding:6px 0}
.file-row{display:flex;justify-content:space-between;align-items:center;padding:4px 16px;font-size:12px;color:#888}
.rm-btn{background:transparent;border:none;color:#444;cursor:pointer;font-size:12px;padding:2px 6px}
.rm-btn:hover{color:var(--red)}
.run-btn{background:var(--orange);color:var(--bg);font-family:"Syne",sans-serif;font-weight:700;font-size:15px;border:none;border-radius:6px;padding:14px 48px;cursor:pointer;transition:.2s;box-shadow:0 4px 16px rgba(255,107,53,.25)}
.run-btn:hover:not(:disabled){background:#ff8c5a;transform:translateY(-1px);box-shadow:0 6px 24px rgba(255,107,53,.4)}
.run-btn:disabled{opacity:.4;cursor:not-allowed}
.progress-wrap{width:100%;max-width:620px;display:none;flex-direction:column;gap:6px}
.progress-wrap.show{display:flex}
.progress-track{width:100%;height:6px;background:#222;border-radius:3px;overflow:hidden}
.progress-bar{height:100%;background:linear-gradient(90deg,var(--orange),#ffaa70);border-radius:3px;width:0%;transition:width .3s}
.log-box{width:100%;max-width:620px;background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:20px;display:none;flex-direction:column;gap:5px;max-height:240px;overflow-y:auto}
.log-box.show{display:flex}
.log{font-size:12px;line-height:1.7}
.log.ok{color:var(--green)}.log.err{color:var(--red)}.log.zip{color:var(--orange)}.log.info{color:#ccc}
.done-banner{width:100%;max-width:620px;background:rgba(90,255,154,.08);border:1px solid rgba(90,255,154,.3);color:var(--green);border-radius:8px;padding:16px 24px;font-size:14px;text-align:center;display:none}
.done-banner.show{display:block}
.footer{font-size:11px;color:#333;text-align:center;margin-top:8px}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.pulsing{animation:pulse 1.2s infinite}
</style>
</head>
<body>
<div class="header">
<div class="badge">ESTRATTORE FATTURE</div>
<h1>XML / ZIP → Excel</h1>
<p class="sub">Carica XML, PDF o ZIP · Scarica l'Excel</p>
</div>
<div class="upload-area">
<label class="drop-zone" id="dropZone">
<input type="file" id="fileInput" accept=".xml,.zip,.pdf" multiple style="display:none"/>
<div class="drop-icon">⬇</div>
<div class="drop-text">Trascina qui XML, PDF o ZIP oppure clicca</div>
<div class="drop-sub">Puoi caricare più file contemporaneamente</div>
</label>
<div class="or-div">OPPURE</div>
<label class="folder-btn">
<input type="file" id="folderInput" style="display:none" webkitdirectory directory/>
   📁 &nbsp;Carica una cartella intera
<div class="folder-sub">Prende tutti gli XML e ZIP al suo interno</div>
</label>
</div>
<div class="file-panel" id="filePanel">
<div class="fp-header">
<span class="fp-count" id="fpCount">0 file selezionati</span>
<button class="clear-btn" onclick="clearAll()">✕ Svuota tutto</button>
</div>
<div class="file-list" id="fileList"></div>
</div>
<button class="run-btn" id="runBtn" onclick="handleRun()" disabled>▶ &nbsp;Avvia Estrazione</button>
<div class="progress-wrap" id="progressWrap">
<div class="progress-track"><div class="progress-bar" id="progressBar"></div></div>
</div>
<div class="log-box" id="logBox"></div>
<div class="done-banner" id="doneBanner"></div>
<div class="footer">I file vengono elaborati sul server e non vengono salvati.</div>
<script>
let selectedFiles = [];
function addFiles(newFiles) {
 const valid = Array.from(newFiles).filter(f =>
   f.name.toLowerCase().endsWith('.xml') || f.name.toLowerCase().endsWith('.zip') || f.name.toLowerCase().endsWith('.pdf')
 );
 const existing = new Set(selectedFiles.map(f => f.name + f.size));
 valid.forEach(f => { if (!existing.has(f.name + f.size)) selectedFiles.push(f); });
 render();
}
function removeFile(i) { selectedFiles.splice(i,1); render(); }
function clearAll() {
 selectedFiles = [];
 render();
 document.getElementById('logBox').classList.remove('show');
 document.getElementById('doneBanner').classList.remove('show');
}
function render() {
 const panel = document.getElementById('filePanel');
 const list  = document.getElementById('fileList');
 const count = document.getElementById('fpCount');
 const btn   = document.getElementById('runBtn');
 if (!selectedFiles.length) { panel.classList.remove('show'); btn.disabled=true; return; }
 panel.classList.add('show');
 btn.disabled = false;
 count.textContent = selectedFiles.length + ' file selezionati';
 list.innerHTML = selectedFiles.map((f,i) =>
   `<div class="file-row"><span>📄 ${f.name}</span><button class="rm-btn" onclick="removeFile(${i})">✕</button></div>`
 ).join('');
}
document.getElementById('fileInput').onchange = e => addFiles(e.target.files);
document.getElementById('folderInput').onchange = e => addFiles(e.target.files);
const dz = document.getElementById('dropZone');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('over'));
dz.addEventListener('drop', e => { e.preventDefault(); dz.classList.remove('over'); addFiles(e.dataTransfer.files); });
dz.addEventListener('click', () => document.getElementById('fileInput').click());
function log(msg, type='info') {
 const box = document.getElementById('logBox');
 box.classList.add('show');
 const d = document.createElement('div');
 d.className = 'log ' + type;
 d.textContent = msg;
 box.appendChild(d);
 box.scrollTop = box.scrollHeight;
}
async function handleRun() {
 if (!selectedFiles.length) return;
 const btn = document.getElementById('runBtn');
 btn.disabled = true;
 btn.classList.add('pulsing');
 btn.textContent = '⏳  Elaborazione in corso…';
 document.getElementById('logBox').innerHTML = '';
 document.getElementById('logBox').classList.add('show');
 document.getElementById('doneBanner').classList.remove('show');
 document.getElementById('progressWrap').classList.add('show');
 document.getElementById('progressBar').style.width = '30%';
 const fd = new FormData();
 selectedFiles.forEach(f => fd.append('files', f));
 try {
   log('📤 Invio ' + selectedFiles.length + ' file al server…', 'info');
   const resp = await fetch('/process', { method: 'POST', body: fd });
   document.getElementById('progressBar').style.width = '80%';
   if (!resp.ok) {
     const err = await resp.json();
     log('❌ Errore: ' + (err.error || resp.statusText), 'err');
     return;
   }
   // Leggi stats dagli header
   const fatture  = resp.headers.get('X-Rows-Fatture')  || '?';
   const dups     = resp.headers.get('X-Rows-Duplicati') || '?';
   log('✅ Elaborazione completata!', 'ok');
   log('   Foglio Fatture:   ' + fatture + ' righe', 'ok');
   log('   Foglio Duplicati: ' + dups + ' righe', 'ok');
   document.getElementById('progressBar').style.width = '100%';
   // Download automatico
   const blob = await resp.blob();
   const url  = URL.createObjectURL(blob);
   const a    = document.createElement('a');
   a.href = url; a.download = 'estrazione_fatture.xlsx';
   document.body.appendChild(a); a.click();
   document.body.removeChild(a);
   URL.revokeObjectURL(url);
   const banner = document.getElementById('doneBanner');
   banner.innerHTML = '✅ &nbsp;<strong>estrazione_fatture.xlsx</strong> scaricato! &nbsp;· &nbsp;' + fatture + ' fatture &nbsp;· &nbsp;' + dups + ' duplicati';
   banner.classList.add('show');
 } catch(e) {
   log('❌ Errore di rete: ' + e.message, 'err');
 } finally {
   btn.disabled = false;
   btn.classList.remove('pulsing');
   btn.textContent = '▶  Avvia Estrazione';
 }
}
</script>
</body>
</html>'''

@app.route('/debug', methods=['POST'])
def debug():
   """Endpoint di debug: restituisce il testo grezzo estratto dal PDF."""
   files = request.files.getlist('files')
   if not files:
       return jsonify({"error": "Nessun file"}), 400
   f = files[0]
   data = f.read()
   text = ""
   try:
       with pdfplumber.open(io.BytesIO(data)) as pdf:
           for page in pdf.pages:
               t = page.extract_text()
               if t:
                   text += t + "\n"
   except Exception as e:
       return jsonify({"error": str(e)}), 500
   return jsonify({"text": text[:3000]})

@app.route('/')
def index():
   return render_template_string(HTML)

@app.route('/process', methods=['POST'])
def process():
   files = request.files.getlist('files')
   if not files:
       return jsonify({"error": "Nessun file ricevuto"}), 400
   all_rows = []
   for f in files:
       name = f.filename.lower()
       data = f.read()
       if name.endswith('.xml'):
           process_xml_bytes(data, all_rows)
       elif name.endswith('.pdf'):
           process_pdf_bytes(data, all_rows)
       elif name.endswith('.zip'):
           process_zip_bytes(data, all_rows)
   if not all_rows:
       return jsonify({"error": "Nessun dato estratto dai file caricati"}), 422
   excel_bytes, n_fatture, n_dups = build_excel(all_rows)
   response = send_file(
       excel_bytes,
       mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
       as_attachment=True,
       download_name='estrazione_fatture.xlsx'
   )
   response.headers['X-Rows-Fatture']   = str(n_fatture)
   response.headers['X-Rows-Duplicati'] = str(n_dups)
   response.headers['Access-Control-Expose-Headers'] = 'X-Rows-Fatture, X-Rows-Duplicati'
   return response

if __name__ == '__main__':
   port = int(os.environ.get('PORT', 5000))
   app.run(host='0.0.0.0', port=port)
