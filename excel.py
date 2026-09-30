"""
Costruzione del file Excel: fogli Fatture, Righe e Duplicati.
"""
import io
import re
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

COLORE_INTESTAZIONE = "FF6B35"
FORMATO_EURO = '#,##0.00'
FORMATO_DATA = 'DD/MM/YYYY'


def categoria(descrizione):
    """Categorie usate per le fatture dei veicoli (colonna mostrata solo se serve)."""
    d = (descrizione or "").lower()
    for chiave, nome in (("forfait", "Forfait"), ("over plafond", "Over Plafond"),
                         ("km eccedenti", "Km Eccedenti"), ("esubero km", "Km Eccedenti"),
                         ("eccedenza chilometrica", "Km Eccedenti"), ("tagliando", "Tagliando"),
                         ("perizia", "Perizia"), ("elementi tecnici mancanti", "ETM"),
                         ("penale per danni", "Danni")):
        if chiave in d:
            return nome
    if re.search(r'\beam\b', d):
        return "EAM"
    return ""


def _data(val):
    try:
        a, m, g = (int(x) for x in val.split('-'))
        return date(a, m, g)
    except Exception:
        return val or None


def _chiave(f):
    """Due fatture sono la stessa se coincidono fornitore, numero, data e totale."""
    numero = re.sub(r'\s', '', (f.get("numero") or "")).upper()
    if not numero:
        return None
    fornitore = (f.get("piva_fornitore") or f.get("fornitore") or "").upper().replace(" ", "")
    totale = round(f["totale"], 2) if isinstance(f.get("totale"), float) else None
    return (fornitore, numero, f.get("data") or "", totale)


def _intestazione(ws, colonne):
    riempimento = PatternFill("solid", fgColor=COLORE_INTESTAZIONE)
    for c, (titolo, larghezza, _) in enumerate(colonne, start=1):
        cella = ws.cell(row=1, column=c, value=titolo)
        cella.fill = riempimento
        cella.font = Font(bold=True, color="FFFFFF")
        cella.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c)].width = larghezza
    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"


def _scrivi(ws, colonne, righe):
    for valori in righe:
        ws.append(valori)
    for c, (_, _, formato) in enumerate(colonne, start=1):
        if formato:
            for (cella,) in ws.iter_rows(min_row=2, min_col=c, max_col=c):
                cella.number_format = formato
    if ws.max_row > 1:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(colonne))}{ws.max_row}"


COLONNE_FATTURE = [
    ("Fornitore", 32, None), ("P.IVA fornitore", 16, None), ("Cliente", 30, None),
    ("P.IVA cliente", 16, None), ("Tipo documento", 16, None), ("Numero", 16, None),
    ("Data", 12, FORMATO_DATA), ("Imponibile (€)", 15, FORMATO_EURO),
    ("IVA (€)", 13, FORMATO_EURO), ("Totale (€)", 15, FORMATO_EURO), ("Valuta", 8, None),
    ("N. righe", 9, None), ("File", 30, None), ("Letto da", 18, None),
]


def _riga_fattura(f):
    return [f["fornitore"], f["piva_fornitore"], f["cliente"], f["piva_cliente"],
            f["tipo"], f["numero"], _data(f["data"]), f["imponibile"], f["iva"],
            f["totale"], f["valuta"], len(f["righe"]), f["file"], f["letto_con"]]


def crea_excel(fatture):
    """Restituisce (BytesIO, n_fatture_uniche, n_duplicati)."""
    uniche, duplicati, visti = [], [], {}
    for f in fatture:
        k = _chiave(f)
        if k and k in visti:
            duplicati.append((f, visti[k]))
        else:
            if k:
                visti[k] = f["file"]
            uniche.append(f)

    wb = openpyxl.Workbook()

    # Foglio 1: una riga per fattura
    ws = wb.active
    ws.title = "Fatture"
    _intestazione(ws, COLONNE_FATTURE)
    _scrivi(ws, COLONNE_FATTURE, [_riga_fattura(f) for f in uniche])

    # Foglio 2: una riga per ogni voce delle fatture
    righe = [(f, r) for f in uniche for r in f["righe"]]
    con_targa = any(r["targa"] for _, r in righe)
    con_telaio = any(r["telaio"] for _, r in righe)
    con_categoria = any(categoria(r["descrizione"]) for _, r in righe)
    colonne = [("Fornitore", 30, None), ("Numero", 14, None), ("Data", 12, FORMATO_DATA),
               ("Descrizione", 50, None), ("Quantità", 10, '#,##0.##'),
               ("Prezzo unitario (€)", 16, FORMATO_EURO), ("Aliquota IVA %", 12, '0.##'),
               ("Importo (€)", 15, FORMATO_EURO)]
    if con_targa:
        colonne.append(("Targa", 11, None))
    if con_telaio:
        colonne.append(("Telaio", 21, None))
    if con_categoria:
        colonne.append(("Categoria", 16, None))
    ws2 = wb.create_sheet("Righe")
    _intestazione(ws2, colonne)
    dati_righe = []
    for f, r in righe:
        valori = [f["fornitore"], f["numero"], _data(f["data"]), r["descrizione"],
                  r["quantita"], r["prezzo_unitario"], r["aliquota_iva"], r["importo"]]
        if con_targa:
            valori.append(r["targa"])
        if con_telaio:
            valori.append(r["telaio"])
        if con_categoria:
            valori.append(categoria(r["descrizione"]))
        dati_righe.append(valori)
    _scrivi(ws2, colonne, dati_righe)

    # Foglio 3: fatture caricate più di una volta
    colonne_dup = COLONNE_FATTURE + [("Già presente nel file", 30, None)]
    ws3 = wb.create_sheet("Duplicati")
    _intestazione(ws3, colonne_dup)
    _scrivi(ws3, colonne_dup, [_riga_fattura(f) + [primo] for f, primo in duplicati])

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out, len(uniche), len(duplicati)
