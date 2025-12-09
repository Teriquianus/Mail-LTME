# -*- coding: utf-8 -*-
import os
import re
import shutil
import tempfile
from datetime import datetime
import pandas as pd
import unicodedata

# COM (Word/Outlook)
import win32com.client as win32
from win32com.client import constants as c

# =========================
# KONFIG (Pfad & Konto)
# =========================
PATH_USTVA_BWA = r"C:\Users\AlexanderHaller\Unternehmenskompass GmbH\Unternehmenskompass - CRM\LTME\Automation\UStVA + BWA (Mandanten).docx"
PATH_FEEDBACK_SB = r"C:\Users\AlexanderHaller\Unternehmenskompass GmbH\Unternehmenskompass - CRM\LTME\Automation\Feedback (Selbstbucher).docx"
PATH_EXCEL = r"C:\Users\AlexanderHaller\Unternehmenskompass GmbH\Unternehmenskompass - CRM\LTME\LTME Working.xlsm"
FEEDBACK_ROOT = r"C:\Users\AlexanderHaller\Unternehmenskompass GmbH\Unternehmenskompass - CRM\LTME"

SHEET_NAME = "Vorlage Mail"
SMTP_INFO = "info@ltme-consulting.de"

# Spaltenindex (1-basiert in Excel, hier 0-basiert nach pandas)
COL_MANDANT = 0
COL_TYP = 2
COL_INTERVALL = 3
COL_VORNAME = 4
COL_EMAIL = 5
COL_ZEITRAUM = 6
COL_ZAHLLAST = 7
COL_FLAG_FEEDBACK = 8
COL_FLAG_USTVA = 9

EURO = "\u20AC"

# =========================
# Hilfsfunktionen
# =========================
def to_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    s = str(v).strip().upper()
    return s in {"WAHR", "TRUE", "JA", "X", "1"}

def format_zahllast(v) -> str:
    try:
        x = float(v)
        if x < 0:
            return f"- {abs(x):,.2f} {EURO}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"{x:,.2f} {EURO}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

def html_encode(s: str) -> str:
    return (s.replace("&","&amp;")
              .replace("<","&lt;")
              .replace(">","&gt;")
              .replace('"',"&quot;")
              .replace("'","&#39;"))

def read_text_utf8(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()

def ensure_utf8_meta(html: str) -> str:
    if re.search(r"charset\s*=", html, flags=re.I):
        return html
    return html.replace("<head>", '<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">', 1)

# ---- Addison-kompatibler Zeitraum-Parser ----
def _expand_two_digit_year(jj: str) -> str:
    jj = (jj or "").strip()
    if len(jj) == 2 and jj.isdigit():
        return f"20{jj}"  # simple Regel: 25 -> 2025
    return jj

def _norm_token(s: str) -> str:
    # trim, lower, NBSP→space, Diakritika weg, Mehrfachspaces raus
    s = (s or "").replace("\u00A0", " ").strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s

_MONTHS = {
    # Deutsch kurz/lang + übliche Ersatzschreibungen
    "jan": 1, "januar": 1,
    "feb": 2, "februar": 2,
    "mar": 3, "maerz": 3, "mrz": 3, "marz": 3, "maer": 3,
    "apr": 4, "april": 4,
    "mai": 5,
    "jun": 6, "juni": 6,
    "jul": 7, "juli": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "okt": 10, "oktober": 10,
    "nov": 11, "november": 11,
    "dez": 12, "dezember": 12,
}

def _year4(y: str) -> str:
    y = y.strip()
    if len(y) == 2 and y.isdigit():
        return f"20{y}"
    return y

def _quarter_months(q: int):
    start = (q-1)*3 + 1
    return [start, start+1, start+2]

def months_for_search(tf: str) -> list[str]:
    """Erkennt Addison-Formate wie "JAN 25" oder "II 2025" und liefert YYYY-MM Strings."""

    if not tf:
        return []

    try:
        import pandas as _pd
    except Exception:
        _pd = None

    if isinstance(tf, datetime) or (_pd is not None and isinstance(tf, _pd.Timestamp)):
        return [f"{tf.year}-{int(tf.month):02d}"]

    s_raw = str(tf).strip()
    s = _norm_token(s_raw)

    # Römische Quartale: "II 2025"
    m_q = re.match(r"^(i{1,3}|iv)\s*(\d{2,4})$", s)
    if m_q:
        q = {"i": 1, "ii": 2, "iii": 3, "iv": 4}[m_q.group(1)]
        year = _expand_two_digit_year(m_q.group(2))
        return [f"{year}-{m:02d}" for m in _quarter_months(q)] if len(year) == 4 else []

    # Monatsnamen: "jan 25", "maerz 2025", "mrz 25"
    m_m = re.match(r"^([a-zäöü]+)\s*(\d{2,4})$", s)
    if m_m:
        mon_key = m_m.group(1)
        year = _expand_two_digit_year(m_m.group(2))
        if mon_key in _MONTHS and len(year) == 4 and year.isdigit():
            mm = _MONTHS[mon_key]
            return [f"{year}-{mm:02d}"]

    return []


def run_months_for_search_selftest() -> None:
    """Schnelle Regressionstests für die Addison-Zeitraum-Erkennung."""

    cases = {
        "Jan 25": ["2025-01"],
        "JAN 25": ["2025-01"],
        "Mär 25": ["2025-03"],
        "Maerz 2025": ["2025-03"],
        "III 2025": ["2025-07", "2025-08", "2025-09"],
        "II 2025": ["2025-04", "2025-05", "2025-06"],
    }
    for k, exp in cases.items():
        got = months_for_search(k)
        assert got == exp, f"{k} -> {got} != {exp}"
    print("[OK] months_for_search Selftest")

def display_timeframe(tf: str) -> str:
    keys = months_for_search(tf)
    if not keys:
        return tf
    if len(keys) == 1:
        y, m = keys[0].split("-")
        month_names = ["Januar","Februar","März","April","Mai","Juni","Juli","August","September","Oktober","November","Dezember"]
        return f"{month_names[int(m)-1]} {y}"
    # Quartalsanzeige aus dem ersten Monat ableiten
    y, m1 = keys[0].split("-")
    q = ((int(m1)-1)//3)+1
    return f"{q}. Quartal {y}"

def build_nested_feedback_html(content: str) -> str:
    # Erwartet "BELEGE"/"BANK" Überschriften und "- " Bulletpoints; tolerant ohne Überschriften.
    content = content.replace("\r\n\r\n", "\r\n")
    lines = [ln.strip() for ln in content.splitlines()]
    section = ""
    belege, bank = [], []
    for t in lines:
        if not t:
            continue
        tu = t.upper().rstrip(":")
        if tu in ("BELEGE", "BANK"):
            section = tu
            continue
        if t.startswith("- "):
            t = t[2:]
        if not section:
            section = "BELEGE"
        if section == "BELEGE":
            belege.append(html_encode(t))
        else:
            bank.append(html_encode(t))
    if not belege and not bank:
        for t in lines:
            if not t:
                continue
            if t.startswith("- "):
                t = t[2:]
            belege.append(html_encode(t))
    def joinlis(arr):
        return "".join(f"<li>{x}</li>" for x in arr if x.strip())
    html = "<ul>"
    if belege:
        html += f"<li><u>Belege</u><ul>{joinlis(belege)}</ul></li>"
    if bank:
        html += f"<li><u>Bank</u><ul>{joinlis(bank)}</ul></li>"
    html += "</ul>"
    return html

def normalize_mandant(v) -> str:
    """Excel-Zahlen wie 10010.0 -> '10010'; Strings wie '10010 ' -> '10010'."""
    s = str(v).strip()
    # 10010.0 / 10010.000
    m = re.match(r"^(\d+)(?:\.0+)?$", s)
    if m:
        return m.group(1)
    # Falls echt floatig oder mit Komma/Punkt formatiert:
    try:
        f = float(s.replace(" ", "").replace(",", "."))
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    # Fallback: führenden Ziffernblock nehmen (z.B. "10010 - Kunde")
    m = re.match(r"^(\d+)", s)
    return m.group(1) if m else s

def build_feedback_block(mandant: str, timeframe: str) -> str:
    mandant = normalize_mandant(mandant)
    base = None
    for name in os.listdir(FEEDBACK_ROOT):
        # exakte führende Zahl extrahieren und vergleichen
        m = re.match(r"^(\d+)", name.strip())
        if m and m.group(1) == mandant:
            p = os.path.join(FEEDBACK_ROOT, name)
            if os.path.isdir(p):
                base = p
                break

    if not base:
        print(f"[WARN] Mandantenordner nicht gefunden für '{mandant}' unter {FEEDBACK_ROOT}")
        return ""

    # Einheitliche Zeitraumschlüssel (YYYY-MM, ggf. 3 Stück bei Quartal)
    keys = months_for_search(timeframe)
    if not keys:
        print(f"[WARN] Zeitraum nicht erkannt: '{timeframe}'")
        return ""

    out = []
    for key in keys:
        fname = f"Feedback FiBu {key}.txt"
        fpath = os.path.join(base, fname)
        if os.path.isfile(fpath):
            print(f"[HIT] {fname}")
            content = read_text_utf8(fpath)
            html_list = build_nested_feedback_html(content)
            header = html_encode(os.path.splitext(fname)[0])
            out.append(f"<b><u>{header}</u></b>{html_list}<br>")
        else:
            print(f"[MISS] {fname} nicht gefunden in {base}")

    # Alles gesammelt → Schriftgröße 11pt für gesamten Block
    if out:
        return f'<div style="font-size:11pt;">{"".join(out)}</div>'
    return ""

# =========================
# Word → HTML (UTF-8)
# =========================
def word_fill_to_html(word_app, template_path: str, placeholders: dict, tmpdir: str) -> str:
    # Vorlage öffnen und in neues Doc übernehmen
    tpl = word_app.Documents.Open(FileName=template_path, ReadOnly=True)
    doc = word_app.Documents.Add()
    doc.Content.FormattedText = tpl.Content.FormattedText

    # Platzhalter ersetzen
    find = doc.Content.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()
    for ph, val in placeholders.items():
        find.Execute(FindText=ph, ReplaceWith=val, Replace=c.wdReplaceAll)

    # HTML (nicht gefiltert!) + UTF-8 + PNG erlauben
    out_html = os.path.join(tmpdir, f"mail_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.htm")
    doc.WebOptions.Encoding = 65001        # UTF-8
    doc.WebOptions.AllowPNG = True         # bessere Grafikqualität
    # WICHTIG: Voll-HTML behalten, damit Styles drin bleiben
    doc.SaveAs2(FileName=out_html, FileFormat=c.wdFormatHTML)  # = 8

    tpl.Close(False)
    doc.Close(False)
    return out_html

# =========================
# Outlook Helpers
# =========================
def get_account(ns, smtp: str):
    target = smtp.strip().lower()
    for ac in ns.Accounts:
        try:
            if str(ac.SmtpAddress).strip().lower() == target:
                return ac
        except Exception:
            pass
    return None


def create_draft_mail(outlook_app, account, email: str, subject: str, html: str, drafts_folder) -> None:
    """Erzeugt einen HTML-Entwurf und verschiebt ihn in den Drafts-Ordner."""

    mail = outlook_app.CreateItem(0)  # olMailItem
    mail.BodyFormat = 2               # olFormatHTML
    mail.HTMLBody = html
    mail.To = email
    mail.Subject = subject
    mail.SendUsingAccount = account
    mail.Save()
    mail.Move(drafts_folder)

def main():
    # Vorab: Pfade prüfen
    for p in (PATH_USTVA_BWA, PATH_FEEDBACK_SB, PATH_EXCEL, FEEDBACK_ROOT):
        if not os.path.exists(p):
            raise FileNotFoundError(f"Pfad nicht gefunden: {p}")

    run_months_for_search_selftest()

    # Excel laden
    df = pd.read_excel(PATH_EXCEL, sheet_name=SHEET_NAME, header=0)

    # Mandantenspalte-Name holen
    COL_MANDANT_NAME = df.columns[COL_MANDANT]

    # Leere Mandantennummern filtern und Copy ziehen (wichtig!)
    df = df.loc[df[COL_MANDANT_NAME].notna()].copy()

    # Spalte endgültig auf Text umstellen, dann normalisieren
    df[COL_MANDANT_NAME] = df[COL_MANDANT_NAME].astype('string')
    df[COL_MANDANT_NAME] = df[COL_MANDANT_NAME].map(normalize_mandant).astype('string')


    # COM-Apps
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    outlook = win32.gencache.EnsureDispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")

    acct = get_account(ns, SMTP_INFO)
    if acct is None:
        raise RuntimeError(f"Outlook-Konto '{SMTP_INFO}' nicht gefunden.")
    drafts = acct.DeliveryStore.GetDefaultFolder(16)  # olFolderDrafts

    tmpdir = tempfile.mkdtemp(prefix="ltme_")

    count_fb = 0
    count_u = 0
    summary_lines = []

    try:
        for idx, row in df.iterrows():
            # Sichtbarkeitslogik aus Excel (falls nötig): hier ignoriert — wir nehmen alle nichtleeren
            mandant = row.iloc[COL_MANDANT]  # bereits normalisiert -> '10010'
            typ       = str(row.iloc[COL_TYP]).strip().upper() if not pd.isna(row.iloc[COL_TYP]) else ""
            intervall = "" if pd.isna(row.iloc[COL_INTERVALL]) else str(row.iloc[COL_INTERVALL])
            vorname   = "" if pd.isna(row.iloc[COL_VORNAME]) else str(row.iloc[COL_VORNAME])
            email     = "" if pd.isna(row.iloc[COL_EMAIL]) else str(row.iloc[COL_EMAIL])
            raw_zeitraum = row.iloc[COL_ZEITRAUM]
            zeitraum = "" if pd.isna(raw_zeitraum) else raw_zeitraum
            # zeitraum  = "" if pd.isna(row.iloc[COL_ZEITRAUM]) else str(row.iloc[COL_ZEITRAUM])
            zahllast  = format_zahllast(row.iloc[COL_ZAHLLAST])
            did_fb = False
            did_ust = False

            has_feedback = to_bool(row.iloc[COL_FLAG_FEEDBACK])
            has_ustva    = to_bool(row.iloc[COL_FLAG_USTVA])
            if not (has_feedback or has_ustva):
                continue

            # Platzhalter
            placeholders = {
                "{{Vorname}}": vorname,
                "{{Email}}": email,
                "{{Zeitraum}}": display_timeframe(zeitraum),
                "{{Zahllast}}": zahllast,
                "{{UStVA-Intervall}}": intervall,
                "{{Feedback}}": "{{Feedback}}"
            }

            # -----------------------
            # FEEDBACK
            # -----------------------
            if has_feedback:
                html_path = word_fill_to_html(word, PATH_FEEDBACK_SB, placeholders, tmpdir)
                html = read_text_utf8(html_path)
                block = build_feedback_block(mandant, zeitraum)
                html = html.replace("{{Feedback}}", block)
                html = ensure_utf8_meta(html)
                subject = f"Feedback Finanzbuchhaltung f\u00FCr {display_timeframe(zeitraum)}"
                create_draft_mail(outlook, acct, email, subject, html, drafts)
                did_fb = True
                count_fb += 1

            # -----------------------
            # USTVA/BWA
            # -----------------------
            if has_ustva:
                html_path = word_fill_to_html(word, PATH_USTVA_BWA, placeholders, tmpdir)
                html = read_text_utf8(html_path)
                html = html.replace("{{Feedback}}", "")  # falls Platzhalter existiert
                html = ensure_utf8_meta(html)
                subject = f"UStVA- und BWA-Ergebnis f\u00FCr {display_timeframe(zeitraum)}"
                create_draft_mail(outlook, acct, email, subject, html, drafts)
                did_ust = True
                count_u += 1

            parts = []
            if did_fb:
                parts.append("Feedback")
            if did_ust:
                parts.append("UStVA/BWA")
            if parts:  # nur wenn überhaupt etwas erstellt wurde
                summary_lines.append(f"- Mandant {mandant}, Zeitraum {display_timeframe(zeitraum)}: {', '.join(parts)}")

        # ANSI-Codes für Unterstreichung
        UNDERLINE = "\033[4m"
        RESET = "\033[0m"
        print(f"\n{UNDERLINE}Erstellt: {count_fb} Feedback-Entwürfe, {count_u} UStVA/BWA-Entwürfe:{RESET}")
        for line in summary_lines:
            print(line)
        print()

    finally:
        # Aufräumen
        try:
            word.Quit(SaveChanges=False)
        except Exception:
            pass
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            pass
        # Excel-Datei nach dem Durchlauf wieder öffnen
        try:
            os.startfile(PATH_EXCEL)  # öffnet die .xlsm mit Excel
        except Exception as e:
            print(f"[WARN] Excel konnte nicht geöffnet werden: {e}")

if __name__ == "__main__":
    main()