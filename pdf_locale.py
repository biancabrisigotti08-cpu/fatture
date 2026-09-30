"""
Lettore PDF locale (senza AI): usato solo se Gemini non è disponibile.
Riconosce i formati già noti (Volkswagen, PSA, Romana Diesel) e prova un
lettore semplice per gli altri.
"""
import io
import re
from pypdf import PdfReader

IMPORTO_BOLLO = 2.00

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
        # ── Parser Volkswagen — usa numeri di riga come separatori ──
        # Il testo reale dai log mostra che prima di ogni "ADDEBITO PENALE PER"
        # c'è un numero di riga (es. "30\nADDEBITO PENALE PER DANNI -\n...")
        # Usiamo il pattern "\nN.\n" o "\nN \n" come separatore di blocco.
        # Questo funziona anche quando il blocco va a capo pagina perché
        # il numero di riga appare sempre prima della descrizione.
        import pdfplumber
        # Estrai testo pagina per pagina con pdfplumber (più preciso sui \n)
        full_pages_text = []
        try:
            with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
                for page in pdf.pages:
                    t = page.extract_text(layout=True)
                    if t:
                        full_pages_text.append(t)
        except Exception:
            full_pages_text = [text]
        # Unisci tutto il testo delle pagine con un marcatore di pagina
        combined = '\n'.join(full_pages_text)
        # Cerca tutti i numeri di riga nel testo
        # Pattern: numero intero su riga propria (o con spazi) seguito da ADDEBITO
        # Dal PDF reale: "30 ADDEBITO PENALE PER DANNI -"
        # oppure su righe separate: "30\nADDEBITO PENALE PER DANNI"
        row_pattern = re.compile(
            r'(\d+)\s+ADDEBITO\s+PENALE\s+PER',
            re.IGNORECASE
        )
        matches = list(row_pattern.finditer(combined))
        print(f"=== Blocchi VW trovati: {len(matches)} ===")
        for idx, match in enumerate(matches):
            block_start = match.start()
            block_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(combined)
            blocco_text = combined[block_start:block_end]
            # Descrizione
            desc_upper = blocco_text.upper()
            if 'ELEMENTI TECNICI' in desc_upper:
                desc_val = 'Addebito Penale per Elementi Tecnici Mancanti'
            elif 'ECCEDENZA' in desc_upper or 'CHILOMETRICA' in desc_upper:
                desc_val = 'Addebito Penale per Eccedenza Chilometrica'
            elif 'DANNI' in desc_upper:
                desc_val = 'Addebito Penale per Danni'
            else:
                desc_val = 'Addebito Penale'
            # Telaio — cerca in tutto il blocco
            m_telaio = re.search(
                r'Tipo\s*dato:\s*TELAIO\s*[\n\r]+\s*Rif\.\s*testo:\s*(\S+)',
                blocco_text, re.IGNORECASE
            )
            if not m_telaio:
                m_telaio = re.search(r'Rif\.\s*testo:\s*([A-Z0-9]{17})\b', blocco_text, re.IGNORECASE)
            telaio_val = m_telaio.group(1).strip() if m_telaio else ""
            # Targa — cerca in tutto il blocco
            m_targa = re.search(
                r'Tipo\s*dato:\s*TARGA\s*[\n\r]+\s*Rif\.\s*testo:\s*(\S+)',
                blocco_text, re.IGNORECASE
            )
            if not m_targa:
                m_targa = re.search(r'(?<!\w)([A-Z]{2}\d{3}[A-Z]{2})(?!\w)', blocco_text, re.IGNORECASE)
            targa_val = m_targa.group(1).strip() if m_targa else ""
            # Prezzo
            m_prezzo = re.search(
                r'[\d,.]+\s+N[12T]\s+([\d]{1,3}(?:[.,]\d{3})*[.,]\d{2})',
                blocco_text, re.IGNORECASE
            )
            prezzo_val = ""
            if m_prezzo:
                prezzo_val = m_prezzo.group(1).replace('.', '').replace(',', '.')
            try:
                if prezzo_val and abs(float(prezzo_val)) == IMPORTO_BOLLO:
                    continue
            except (ValueError, TypeError):
                pass
            if not prezzo_val:
                continue
            print(f"  Riga {match.group(1)}: telaio={telaio_val} targa={targa_val} prezzo={prezzo_val}")
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
