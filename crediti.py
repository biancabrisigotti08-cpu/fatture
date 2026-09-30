"""
Crediti: quota gratuita mensile e codici di sblocco a pagamento.

Come funziona
- Gli XML sono sempre gratis e illimitati.
- Ogni PDF consuma 1 credito (la lettura dei PDF usa l'AI, che costa).
- Ogni visitatore ha FREE_PDF_MONTH PDF gratis al mese.
- Oltre la quota serve un codice. I codici si creano dalla pagina /admin
  e si mandano al cliente dopo il pagamento PayPal.

Variabili d'ambiente
- DATABASE_URL    Postgres (es. Neon, piano gratuito). Se manca usa SQLite
                  locale, che su Render free si azzera a ogni riavvio.
- FREE_PDF_MONTH  PDF gratis al mese per visitatore (default 10).
- ADMIN_PASSWORD  password della pagina /admin (se manca, /admin è spenta).
- HASH_SALT       stringa casuale: gli indirizzi IP non vengono salvati in
                  chiaro ma solo come impronta (hash).
"""
import hashlib
import os
import secrets
import sqlite3
from datetime import date, datetime, timedelta

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
FREE_PDF_MONTH = int(os.environ.get("FREE_PDF_MONTH", "10"))
HASH_SALT = os.environ.get("HASH_SALT", "cambia-questo-valore")
SQLITE_PATH = os.environ.get("SQLITE_PATH", "crediti.db")

USE_PG = DATABASE_URL.startswith(("postgres://", "postgresql://"))
ILLIMITATO = -1  # crediti = -1 -> codice senza limiti (per uso personale)

_ALFABETO = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # niente 0/O, 1/I


# ─── Accesso al database (Postgres o SQLite) ────────────────────────────────
def _connetti():
    if USE_PG:
        import psycopg
        return psycopg.connect(DATABASE_URL, autocommit=True)
    conn = sqlite3.connect(SQLITE_PATH)
    conn.isolation_level = None  # autocommit
    return conn


def _sql(query):
    return query.replace("?", "%s") if USE_PG else query


def _esegui(query, params=(), uno=False, tutti=False):
    conn = _connetti()
    try:
        cur = conn.cursor()
        cur.execute(_sql(query), params)
        if uno:
            return cur.fetchone()
        if tutti:
            return cur.fetchall()
        return cur.rowcount
    finally:
        conn.close()


def inizializza():
    _esegui("""CREATE TABLE IF NOT EXISTS uso_gratuito (
        client TEXT NOT NULL,
        mese   TEXT NOT NULL,
        pdf    INTEGER NOT NULL DEFAULT 0,
        PRIMARY KEY (client, mese))""")
    _esegui("""CREATE TABLE IF NOT EXISTS codici (
        codice   TEXT PRIMARY KEY,
        crediti  INTEGER NOT NULL,
        usati    INTEGER NOT NULL DEFAULT 0,
        scadenza TEXT,
        nota     TEXT,
        creato   TEXT NOT NULL)""")


# ─── Identità del visitatore (solo impronta dell'IP, mai l'IP in chiaro) ────
def client_id(request):
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "")
    ip = ip.split(",")[0].strip()
    return hashlib.sha256((HASH_SALT + ip).encode()).hexdigest()[:32]


def _mese():
    return date.today().strftime("%Y-%m")


# ─── Quota gratuita ─────────────────────────────────────────────────────────
def gratis_usati(client):
    row = _esegui("SELECT pdf FROM uso_gratuito WHERE client=? AND mese=?",
                  (client, _mese()), uno=True)
    return row[0] if row else 0


def gratis_rimasti(client):
    return max(0, FREE_PDF_MONTH - gratis_usati(client))


def _consuma_gratis(client, n):
    mese = _mese()
    _esegui("INSERT INTO uso_gratuito (client, mese, pdf) VALUES (?, ?, 0) "
            "ON CONFLICT (client, mese) DO NOTHING", (client, mese))
    _esegui("UPDATE uso_gratuito SET pdf = pdf + ? WHERE client=? AND mese=?",
            (n, client, mese))


# ─── Codici ─────────────────────────────────────────────────────────────────
def normalizza(codice):
    return (codice or "").strip().upper().replace(" ", "")


def info_codice(codice):
    """Stato di un codice, o None se non esiste."""
    codice = normalizza(codice)
    if not codice:
        return None
    row = _esegui("SELECT crediti, usati, scadenza FROM codici WHERE codice=?",
                  (codice,), uno=True)
    if not row:
        return None
    crediti, usati, scadenza = row
    scaduto = bool(scadenza) and date.fromisoformat(scadenza) < date.today()
    illimitato = crediti == ILLIMITATO
    rimasti = None if illimitato else max(0, crediti - usati)
    return {
        "codice": codice,
        "illimitato": illimitato,
        "rimasti": rimasti,
        "usati": usati,
        "scadenza": scadenza,
        "valido": not scaduto and (illimitato or rimasti > 0),
        "scaduto": scaduto,
    }


def crea_codice(crediti, giorni=None, nota=""):
    blocchi = ["".join(secrets.choice(_ALFABETO) for _ in range(4)) for _ in range(3)]
    codice = "FE-" + "-".join(blocchi)
    scadenza = (date.today() + timedelta(days=giorni)).isoformat() if giorni else None
    _esegui("INSERT INTO codici (codice, crediti, usati, scadenza, nota, creato) "
            "VALUES (?, ?, 0, ?, ?, ?)",
            (codice, crediti, scadenza, nota, datetime.now().isoformat(timespec="seconds")))
    return codice


def elenco_codici(limite=100):
    return _esegui("SELECT codice, crediti, usati, scadenza, nota, creato FROM codici "
                   "ORDER BY creato DESC LIMIT ?", (limite,), tutti=True)


# ─── Controllo e addebito ───────────────────────────────────────────────────
def disponibili(client, codice):
    """Quanti PDF può elaborare adesso questo visitatore (None = illimitati)."""
    info = info_codice(codice)
    if info and info["valido"] and info["illimitato"]:
        return None
    extra = info["rimasti"] if info and info["valido"] else 0
    return gratis_rimasti(client) + extra


def addebita(client, codice, n):
    """Scala n PDF: prima dalla quota gratuita, poi dal codice."""
    if n <= 0:
        return
    info = info_codice(codice)
    if info and info["valido"] and info["illimitato"]:
        _esegui("UPDATE codici SET usati = usati + ? WHERE codice=?", (n, info["codice"]))
        return
    dal_gratis = min(n, gratis_rimasti(client))
    if dal_gratis:
        _consuma_gratis(client, dal_gratis)
    resto = n - dal_gratis
    if resto and info and info["valido"]:
        _esegui("UPDATE codici SET usati = usati + ? WHERE codice=?", (resto, info["codice"]))
