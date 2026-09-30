"""
Estrazione dati da fatture di qualsiasi tipo: XML FatturaPA, XML firmati
(.p7m), PDF (con Gemini, più lettore locale di riserva) e ZIP.

Ogni fattura diventa un dizionario con lo stesso schema, qualunque sia il
formato di partenza:
    file, letto_con, tipo, numero, data (AAAA-MM-GG), valuta,
    fornitore, piva_fornitore, cliente, piva_cliente,
    imponibile, iva, totale,
    righe: [{descrizione, quantita, prezzo_unitario, aliquota_iva,
             importo, targa, telaio}]
"""
import base64
import io
import json
import os
import re
import urllib.error
import urllib.request
import zipfile
import xml.etree.ElementTree as ET

from pypdf import PdfReader

TIPI_DOCUMENTO = {
    "TD01": "Fattura", "TD02": "Acconto su fattura", "TD03": "Acconto su parcella",
    "TD04": "Nota di credito", "TD05": "Nota di debito", "TD06": "Parcella",
    "TD16": "Integrazione reverse charge", "TD17": "Autofattura servizi estero",
    "TD18": "Integrazione beni UE", "TD19": "Integrazione beni art.17",
    "TD20": "Autofattura regolarizzazione", "TD21": "Autofattura splafonamento",
    "TD22": "Estrazione deposito IVA", "TD23": "Estrazione deposito IVA con versamento",
    "TD24": "Fattura differita", "TD25": "Fattura differita triangolare",
    "TD26": "Cessione beni ammortizzabili", "TD27": "Autoconsumo / omaggio",
    "TD28": "Acquisti da San Marino", "TD29": "Omessa o irregolare fatturazione",
}

RE_TARGA = re.compile(r'(?<![A-Z0-9])([A-Z]{2}\s?\d{3}\s?[A-Z]{2})(?![A-Z0-9])')
RE_TELAIO = re.compile(r'(?<![A-Z0-9])([A-HJ-NPR-Z0-9]{17})(?![A-Z0-9])')


# ─── Utilità ────────────────────────────────────────────────────────────────
def numero(val):
    """Converte '1.234,56', '1234.56', '-12,5' ecc. in float (None se vuoto)."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace('€', '').replace('EUR', '').replace(' ', '')
    if not s:
        return None
    if ',' in s and '.' in s:
        if s.rfind(',') > s.rfind('.'):
            s = s.replace('.', '').replace(',', '.')
        else:
            s = s.replace(',', '')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def data_iso(val):
    """Porta una data a AAAA-MM-GG (accetta anche gg/mm/aaaa e gg-mm-aa)."""
    if not val:
        return ""
    s = str(val).strip()
    m = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})', s)
    if m:
        return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
    m = re.match(r'^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})$', s)
    if m:
        anno = int(m.group(3))
        if anno < 100:
            anno += 2000
        return f"{anno:04d}-{int(m.group(2)):02d}-{int(m.group(1)):02d}"
    return s


def targa_valida(t):
    t = re.sub(r'\s', '', (t or '').upper())
    return t if re.fullmatch(r'[A-Z]{2}\d{3}[A-Z]{2}', t) else ""


def telaio_valido(t):
    t = (t or '').strip().upper()
    if re.fullmatch(r'[A-HJ-NPR-Z0-9]{17}', t) and re.search(r'\d', t) and re.search(r'[A-Z]', t):
        return t
    return ""


def cerca_veicolo(testo):
    """Trova targa e telaio dentro una descrizione, se ci sono."""
    testo = (testo or '').upper()
    targa = ""
    telaio = ""
    for m in RE_TELAIO.finditer(testo):
        if telaio_valido(m.group(1)):
            telaio = m.group(1)
            break
    for m in RE_TARGA.finditer(testo):
        t = targa_valida(m.group(1))
        if t and t not in telaio:
            targa = t
            break
    return targa, telaio


def fattura_vuota(nome_file, letto_con):
    return {
        "file": nome_file, "letto_con": letto_con, "tipo": "", "numero": "",
        "data": "", "valuta": "EUR", "fornitore": "", "piva_fornitore": "",
        "cliente": "", "piva_cliente": "", "imponibile": None, "iva": None,
        "totale": None, "righe": [],
    }


# ─── XML FatturaPA ──────────────────────────────────────────────────────────
def _senza_namespace(root):
    for el in root.iter():
        if isinstance(el.tag, str) and '}' in el.tag:
            el.tag = el.tag.split('}', 1)[1]
    return root


def _testo(el, percorso):
    if el is None:
        return ""
    trovato = el.find(percorso)
    return trovato.text.strip() if trovato is not None and trovato.text else ""


def _soggetto(blocco):
    """Nome e partita IVA (o codice fiscale) di cedente o cessionario."""
    if blocco is None:
        return "", ""
    ana = blocco.find('DatiAnagrafici')
    nome = _testo(ana, 'Anagrafica/Denominazione')
    if not nome:
        nome = " ".join(x for x in (_testo(ana, 'Anagrafica/Nome'),
                                    _testo(ana, 'Anagrafica/Cognome')) if x)
    paese = _testo(ana, 'IdFiscaleIVA/IdPaese')
    codice = _testo(ana, 'IdFiscaleIVA/IdCodice')
    if codice:
        piva = codice if paese in ("", "IT") else f"{paese}{codice}"
    else:
        piva = _testo(ana, 'CodiceFiscale')
    return nome, piva


def _pulisci_xml(dati):
    """Toglie BOM e spazzatura prima della dichiarazione XML."""
    if dati.startswith(b'\xef\xbb\xbf'):
        dati = dati[3:]
    inizio = dati.find(b'<')
    return dati[inizio:] if inizio > 0 else dati


def leggi_xml(dati, nome_file="", letto_con="XML"):
    """Restituisce una fattura per ogni corpo (lotto) presente nel file."""
    try:
        root = ET.fromstring(_pulisci_xml(dati))
    except ET.ParseError as e:
        raise ValueError(f"XML non valido: {e}")
    root = _senza_namespace(root)
    if root.tag != 'FatturaElettronica':
        trovato = root.find('.//FatturaElettronica')
        if trovato is None and root.find('.//FatturaElettronicaBody') is None:
            raise ValueError("Non è una fattura elettronica")
        root = trovato if trovato is not None else root

    header = root.find('.//FatturaElettronicaHeader')
    fornitore, piva_f = _soggetto(header.find('CedentePrestatore') if header is not None else None)
    cliente, piva_c = _soggetto(header.find('CessionarioCommittente') if header is not None else None)

    fatture = []
    for body in root.iter('FatturaElettronicaBody'):
        f = fattura_vuota(nome_file, letto_con)
        f.update(fornitore=fornitore, piva_fornitore=piva_f,
                 cliente=cliente, piva_cliente=piva_c)
        doc = body.find('DatiGenerali/DatiGeneraliDocumento')
        codice_tipo = _testo(doc, 'TipoDocumento')
        f["tipo"] = TIPI_DOCUMENTO.get(codice_tipo, codice_tipo)
        f["numero"] = _testo(doc, 'Numero')
        f["data"] = data_iso(_testo(doc, 'Data'))
        f["valuta"] = _testo(doc, 'Divisa') or "EUR"

        for linea in body.iter('DettaglioLinee'):
            descr = _testo(linea, 'Descrizione')
            if 'bollo' in descr.lower() and numero(_testo(linea, 'PrezzoTotale')) == 2.0:
                continue   # bollo virtuale riaddebitato: non è una voce della spesa
            targa, telaio = cerca_veicolo(descr)
            for altro in linea.iter('AltriDatiGestionali'):
                tipo_dato = _testo(altro, 'TipoDato').upper()
                rif = _testo(altro, 'RiferimentoTesto')
                if 'TARGA' in tipo_dato and targa_valida(rif):
                    targa = targa_valida(rif)
                elif 'TELAIO' in tipo_dato and telaio_valido(rif):
                    telaio = telaio_valido(rif)
            f["righe"].append({
                "descrizione": descr,
                "quantita": numero(_testo(linea, 'Quantita')),
                "prezzo_unitario": numero(_testo(linea, 'PrezzoUnitario')),
                "aliquota_iva": numero(_testo(linea, 'AliquotaIVA')),
                "importo": numero(_testo(linea, 'PrezzoTotale')),
                "targa": targa, "telaio": telaio,
            })

        riepiloghi = list(body.iter('DatiRiepilogo'))
        if riepiloghi:
            f["imponibile"] = round(sum(numero(_testo(r, 'ImponibileImporto')) or 0 for r in riepiloghi), 2)
            f["iva"] = round(sum(numero(_testo(r, 'Imposta')) or 0 for r in riepiloghi), 2)
        totale = numero(_testo(doc, 'ImportoTotaleDocumento'))
        if totale is None and f["imponibile"] is not None:
            totale = round(f["imponibile"] + (f["iva"] or 0), 2)
        f["totale"] = totale
        fatture.append(f)
    return fatture


# ─── XML firmati (.p7m) ─────────────────────────────────────────────────────
def _leggi_tlv(buf, pos):
    """Legge un elemento DER/BER: (tag, costruito, inizio_contenuto, fine, fine_elemento)."""
    tag = buf[pos]
    pos += 1
    if tag & 0x1F == 0x1F:           # tag su più byte
        while buf[pos] & 0x80:
            pos += 1
        pos += 1
    lung = buf[pos]
    pos += 1
    if lung == 0x80:                 # lunghezza indefinita
        return tag, True, pos, None, None
    if lung & 0x80:
        n = lung & 0x7F
        lung = int.from_bytes(buf[pos:pos + n], 'big')
        pos += n
    return tag, bool(tag & 0x20), pos, pos + lung, pos + lung


def _octet_strings(buf, pos, fine, risultati, profondita=0):
    """Raccoglie il contenuto di tutte le OCTET STRING (anche spezzate a pezzi)."""
    if profondita > 40:
        return pos
    while pos < fine:
        if buf[pos] == 0 and pos + 1 < len(buf) and buf[pos + 1] == 0:
            return pos + 2           # fine di un elemento a lunghezza indefinita
        tag, costruito, inizio, fine_cont, fine_el = _leggi_tlv(buf, pos)
        if fine_cont is None:        # indefinita
            if tag == 0x24:          # OCTET STRING costruita
                pezzi = []
                pos = _octet_strings(buf, inizio, len(buf), pezzi, profondita + 1)
                risultati.append(b"".join(pezzi))
            else:
                pos = _octet_strings(buf, inizio, len(buf), risultati, profondita + 1)
            continue
        if tag == 0x04:
            risultati.append(buf[inizio:fine_cont])
        elif tag == 0x24:
            pezzi = []
            _octet_strings(buf, inizio, fine_cont, pezzi, profondita + 1)
            risultati.append(b"".join(pezzi))
        elif costruito:
            _octet_strings(buf, inizio, fine_cont, risultati, profondita + 1)
        pos = fine_el
    return pos


def estrai_xml_da_p7m(dati):
    """Estrae l'XML della fattura da un file firmato .p7m (DER o base64)."""
    grezzo = dati.strip()
    if not grezzo.startswith(b'\x30'):
        try:
            testo = re.sub(rb'-----[^-]+-----', b'', grezzo)
            grezzo = base64.b64decode(re.sub(rb'\s', b'', testo), validate=False)
        except Exception:
            pass
    trovati = []
    try:
        _octet_strings(grezzo, 0, len(grezzo), trovati)
    except (IndexError, ValueError):
        pass
    for pezzo in sorted(trovati, key=len, reverse=True):
        if b'FatturaElettronica' in pezzo:
            return pezzo
    # Ultima risorsa: ritaglia l'XML direttamente dai byte
    m = re.search(rb'<\?xml.*?</[\w:]*FatturaElettronica>', grezzo, re.S)
    if m:
        return m.group(0)
    raise ValueError("Impossibile leggere il file .p7m")


# ─── PDF ────────────────────────────────────────────────────────────────────
def xml_allegato_nel_pdf(pdf_bytes):
    """Alcuni PDF contengono l'XML della fattura come allegato."""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        root = reader.trailer['/Root']
        if '/Names' not in root or '/EmbeddedFiles' not in root['/Names']:
            return None
        nomi = root['/Names']['/EmbeddedFiles'].get('/Names', [])
        for i in range(0, len(nomi), 2):
            spec = nomi[i + 1].get_object()
            if '/EF' in spec:
                contenuto = spec['/EF']['/F'].get_object().get_data()
                if b'FatturaElettronica' in contenuto:
                    return contenuto
    except Exception:
        pass
    return None


GEMINI_PROMPT = """Sei un esperto contabile. Nel PDF c'è una o più fatture (o note di credito), di qualsiasi fornitore, layout o lingua: possono essere scansioni, foto o PDF generati da gestionali.
Estrai i dati e rispondi SOLO con JSON valido in questo formato:
{"fatture": [{
  "tipo": "Fattura | Nota di credito | Parcella | Ricevuta | altro",
  "numero": "numero del documento così come scritto",
  "data": "AAAA-MM-GG",
  "valuta": "EUR",
  "fornitore": "ragione sociale o nome di chi emette",
  "piva_fornitore": "partita IVA o VAT number di chi emette, senza spazi",
  "cliente": "ragione sociale o nome del destinatario",
  "piva_cliente": "partita IVA o codice fiscale del destinatario",
  "imponibile": 0.00,
  "iva": 0.00,
  "totale": 0.00,
  "righe": [{
    "descrizione": "testo della riga",
    "quantita": 1,
    "prezzo_unitario": 0.00,
    "aliquota_iva": 22,
    "importo": 0.00,
    "targa": "solo se la riga riguarda un veicolo, formato AA000AA, altrimenti vuoto",
    "telaio": "solo se presente, 17 caratteri, altrimenti vuoto"
  }]
}]}
Regole:
- Tutti gli importi sono numeri con il punto come separatore decimale, senza simbolo di valuta.
- Se un dato non c'è scrivi "" per i testi e null per i numeri. Non inventare nulla.
- Una riga continua su più pagine è UNA sola riga: unisci descrizione, targa e telaio.
- Non includere come righe i totali, i riepiloghi IVA, il bollo virtuale da 2,00 euro o le scadenze di pagamento.
- Se il PDF contiene più fatture, restituisci un elemento per ciascuna."""


def leggi_pdf_con_gemini(pdf_bytes, nome_file=""):
    chiave = os.environ.get("GEMINI_API_KEY", "").strip()
    if not chiave:
        raise RuntimeError("GEMINI_API_KEY non impostata")
    modello = os.environ.get("GEMINI_MODEL", "gemini-3.6-flash").strip()
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{modello}:generateContent"
    payload = {
        "contents": [{"parts": [
            {"text": GEMINI_PROMPT},
            {"inline_data": {"mime_type": "application/pdf",
                             "data": base64.b64encode(pdf_bytes).decode()}},
        ]}],
        "generationConfig": {"temperature": 0, "maxOutputTokens": 32768,
                             "responseMimeType": "application/json"},
    }
    req = urllib.request.Request(
        url, data=json.dumps(payload).encode(), method="POST",
        headers={"Content-Type": "application/json", "x-goog-api-key": chiave})
    try:
        with urllib.request.urlopen(req, timeout=150) as resp:
            risposta = json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Gemini HTTP {e.code}: {e.read().decode('utf-8', 'replace')[:300]}")
    return interpreta_risposta_gemini(risposta, nome_file)


def interpreta_risposta_gemini(risposta, nome_file=""):
    parti = risposta['candidates'][0]['content']['parts']
    testo = "".join(p.get('text', '') for p in parti if not p.get('thought'))
    testo = re.sub(r'```(?:json)?', '', testo).strip()
    dati = json.loads(testo)
    if isinstance(dati, dict) and 'fatture' not in dati:
        dati = {"fatture": [dati]}
    if isinstance(dati, list):
        dati = {"fatture": dati}
    fatture = []
    for g in dati.get("fatture", []):
        f = fattura_vuota(nome_file, "PDF (AI)")
        for campo in ("tipo", "numero", "fornitore", "piva_fornitore", "cliente", "piva_cliente"):
            f[campo] = str(g.get(campo) or "").strip()
        f["piva_fornitore"] = f["piva_fornitore"].replace(" ", "")
        f["piva_cliente"] = f["piva_cliente"].replace(" ", "")
        f["data"] = data_iso(g.get("data"))
        f["valuta"] = (g.get("valuta") or "EUR").strip() or "EUR"
        for campo in ("imponibile", "iva", "totale"):
            f[campo] = numero(g.get(campo))
        for r in g.get("righe") or []:
            descr = str(r.get("descrizione") or "").strip()
            if 'bollo' in descr.lower() and (numero(r.get("importo")) or 0) == 2.0:
                continue
            targa = targa_valida(r.get("targa")) or ""
            telaio = telaio_valido(r.get("telaio")) or ""
            if not (targa and telaio):
                t2, v2 = cerca_veicolo(descr)
                targa, telaio = targa or t2, telaio or v2
            f["righe"].append({
                "descrizione": descr,
                "quantita": numero(r.get("quantita")),
                "prezzo_unitario": numero(r.get("prezzo_unitario")),
                "aliquota_iva": numero(r.get("aliquota_iva")),
                "importo": numero(r.get("importo")),
                "targa": targa, "telaio": telaio,
            })
        fatture.append(f)
    return fatture


def leggi_pdf_locale(pdf_bytes, nome_file=""):
    """Riserva senza AI: usa il vecchio lettore per i formati conosciuti."""
    from pdf_locale import parse_pdf_fattura
    vecchio = parse_pdf_fattura(pdf_bytes)
    f = fattura_vuota(nome_file, "PDF (lettura base)")
    f.update(numero=vecchio.get("numero_documento", ""),
             data=data_iso(vecchio.get("data_documento", "")),
             fornitore=vecchio.get("cedente", ""), cliente=vecchio.get("cessionario", ""))
    for r in vecchio.get("righe", []):
        f["righe"].append({
            "descrizione": r.get("descrizione", ""), "quantita": None,
            "prezzo_unitario": None, "aliquota_iva": None,
            "importo": numero(r.get("prezzo_totale")),
            "targa": targa_valida(r.get("targa", "")), "telaio": telaio_valido(r.get("telaio", "")),
        })
    if f["righe"]:
        f["imponibile"] = round(sum(r["importo"] or 0 for r in f["righe"]), 2)
    return [f] if (f["righe"] or f["numero"]) else []


def leggi_pdf(pdf_bytes, nome_file="", log=print):
    xml = xml_allegato_nel_pdf(pdf_bytes)
    if xml:
        try:
            return leggi_xml(xml, nome_file, "PDF con XML allegato")
        except ValueError:
            pass
    try:
        return leggi_pdf_con_gemini(pdf_bytes, nome_file)
    except Exception as e:
        log(f"Gemini non disponibile per {nome_file}: {e}")
    try:
        return leggi_pdf_locale(pdf_bytes, nome_file)
    except Exception as e:
        log(f"Lettura locale fallita per {nome_file}: {e}")
        return []


# ─── Smistamento per tipo di file ───────────────────────────────────────────
def leggi_file(nome, dati, errori, profondita=0, log=print):
    """Legge un file di qualsiasi tipo supportato e restituisce le fatture trovate."""
    basso = nome.lower()
    breve = os.path.basename(nome)
    try:
        if basso.endswith('.p7m'):
            return leggi_xml(estrai_xml_da_p7m(dati), breve, "XML firmato (.p7m)")
        if basso.endswith('.xml'):
            if basso.endswith(('signature.xml', 'metadata.xml')) or basso.startswith('__macosx') \
                    or re.search(r'_(mt|rc|ns|mc|ne|dt|at)_\d+\.xml$', basso):
                return []   # ricevute e metadati dello SdI, non fatture
            return leggi_xml(dati, breve)
        if basso.endswith('.pdf'):
            trovate = leggi_pdf(dati, breve, log)
            if not trovate:
                errori.append(f"{breve}: nessun dato riconosciuto")
            return trovate
        if basso.endswith('.zip') and profondita < 5:
            fatture = []
            with zipfile.ZipFile(io.BytesIO(dati)) as zf:
                for interno in zf.namelist():
                    if interno.endswith('/') or '__MACOSX' in interno:
                        continue
                    fatture += leggi_file(interno, zf.read(interno), errori, profondita + 1, log)
            return fatture
    except zipfile.BadZipFile:
        errori.append(f"{breve}: ZIP danneggiato")
    except ValueError as e:
        if profondita > 0 and str(e).startswith("Non è una fattura"):
            return []   # es. file di metadati SdI dentro gli ZIP del Cassetto fiscale
        errori.append(f"{breve}: {e}")
    except Exception as e:
        errori.append(f"{breve}: errore imprevisto ({e})")
    return []
