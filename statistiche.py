"""
Statistiche di utilizzo, anonime: per ogni visita o estrazione salva solo
giorno, impronta del visitatore (hash dell'IP, mai l'IP), esito e numeri
(file, PDF, fatture). Nessun contenuto delle fatture viene registrato.
Le righe più vecchie di 13 mesi vengono cancellate automaticamente.
"""
import random
from datetime import date, datetime, timedelta

from crediti import _esegui

ESITI = ("visita", "ok", "quota", "vuoto")
BOT = ("bot", "crawl", "spider", "slurp", "preview", "monitor", "uptime", "curl", "python", "headless")


def inizializza():
    _esegui("""CREATE TABLE IF NOT EXISTS utilizzi (
        quando  TEXT NOT NULL,
        giorno  TEXT NOT NULL,
        client  TEXT NOT NULL,
        esito   TEXT NOT NULL,
        tipo    TEXT NOT NULL,
        file    INTEGER NOT NULL DEFAULT 0,
        pdf     INTEGER NOT NULL DEFAULT 0,
        fatture INTEGER NOT NULL DEFAULT 0)""")


def e_un_bot(user_agent):
    ua = (user_agent or "").lower()
    return not ua or any(b in ua for b in BOT)


def registra(client, esito, tipo="gratis", file=0, pdf=0, fatture=0):
    """Non deve mai far fallire il sito: in caso di errore non registra nulla."""
    try:
        oggi = date.today()
        _esegui("INSERT INTO utilizzi (quando, giorno, client, esito, tipo, file, pdf, fatture) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (datetime.now().isoformat(timespec="seconds"), oggi.isoformat(),
                 client, esito, tipo, int(file), int(pdf), int(fatture)))
        if random.random() < 0.02:
            limite = (oggi - timedelta(days=400)).isoformat()
            _esegui("DELETE FROM utilizzi WHERE giorno < ?", (limite,))
    except Exception as e:
        print("Statistiche non registrate:", e)


def _vuoto():
    return {"visite": 0, "visitatori": set(), "estrazioni": 0, "utenti": set(),
            "fatture": 0, "pdf": 0, "bloccati": set(), "blocchi": 0, "vuote": 0}


def _somma(acc, r):
    giorno, client, esito, tipo, file, pdf, fatture = r
    if esito == "visita":
        acc["visite"] += 1
        acc["visitatori"].add(client)
    elif esito == "ok":
        acc["estrazioni"] += 1
        acc["utenti"].add(client)
        acc["fatture"] += fatture
        acc["pdf"] += pdf
    elif esito == "quota":
        acc["blocchi"] += 1
        acc["bloccati"].add(client)
    elif esito == "vuoto":
        acc["vuote"] += 1


def _numeri(acc):
    return {"visite": acc["visite"], "visitatori": len(acc["visitatori"]),
            "estrazioni": acc["estrazioni"], "utenti": len(acc["utenti"]),
            "fatture": acc["fatture"], "pdf": acc["pdf"],
            "bloccati": len(acc["bloccati"]), "blocchi": acc["blocchi"], "vuote": acc["vuote"]}


def riepilogo(giorni_tabella=14):
    """Totali per oggi / 7 / 30 giorni e dettaglio giornaliero (escluso l'uso personale)."""
    oggi = date.today()
    inizio = (oggi - timedelta(days=29)).isoformat()
    try:
        righe = _esegui("SELECT giorno, client, esito, tipo, file, pdf, fatture FROM utilizzi "
                        "WHERE giorno >= ?", (inizio,), tutti=True) or []
    except Exception as e:
        print("Statistiche non disponibili:", e)
        righe = []
    periodi = {"Oggi": 0, "Ultimi 7 giorni": 6, "Ultimi 30 giorni": 29}
    acc_periodi = {k: _vuoto() for k in periodi}
    acc_giorni = {(oggi - timedelta(days=i)).isoformat(): _vuoto() for i in range(giorni_tabella)}
    personale = {"estrazioni": 0, "fatture": 0}
    for r in righe:
        giorno, tipo, esito = r[0], r[3], r[2]
        if tipo == "illimitato":
            if esito == "ok":
                personale["estrazioni"] += 1
                personale["fatture"] += r[6]
            continue
        delta = (oggi - date.fromisoformat(giorno)).days
        for nome, max_giorni in periodi.items():
            if delta <= max_giorni:
                _somma(acc_periodi[nome], r)
        if giorno in acc_giorni:
            _somma(acc_giorni[giorno], r)
    return {
        "periodi": [(nome, _numeri(acc)) for nome, acc in acc_periodi.items()],
        "giorni": [(g, _numeri(acc)) for g, acc in acc_giorni.items()],
        "personale": personale,
    }
