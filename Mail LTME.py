# -*- coding: utf-8 -*-
import os
import re
import shutil
import tempfile
import time
from datetime import datetime
from pathlib import Path
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
PATH_OPOS_TEMPLATE = r"C:\Users\AlexanderHaller\Unternehmenskompass GmbH\Unternehmenskompass - CRM\LTME\Automation\Offene Bankbewegungen.docx"
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
COL_FLAG_OPOS = 10

# Für OPOS-Zeitraum-Erkennung in Dateinamen
_MONTH_NAMES_DE = {
    "januar": 1, "februar": 2, "märz": 3, "maerz": 3, "april": 4, "mai": 5, "juni": 6,
    "juli": 7, "august": 8, "september": 9, "oktober": 10, "november": 11, "dezember": 12,
}

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


def format_count(count: int, singular: str, plural: str) -> str:
    """Einfache Singular/Plural-Helfer für Ausgaben."""

    return f"{count} {singular if count == 1 else plural}"

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


def find_mandant_folder(mandant: str) -> str | None:
    """Sucht Mandantenordner im FEEDBACK_ROOT; gibt Pfad oder None zurück."""

    mandant = normalize_mandant(mandant)
    for name in os.listdir(FEEDBACK_ROOT):
        m = re.match(r"^(\d+)", name.strip())
        if m and m.group(1) == mandant:
            p = os.path.join(FEEDBACK_ROOT, name)
            if os.path.isdir(p):
                return p
    return None


def opos_period_to_months(text: str) -> set[str]:
    """Parst den Zeitraum aus einem OPOS-Dateinamen-Fragment und liefert YYYY-MM-Werte."""

    if not text:
        return set()

    raw = text.strip()

    # (1) Jahresangabe allein: 2025
    m_year = re.fullmatch(r"(\d{4})", raw)
    if m_year:
        y = m_year.group(1)
        return {f"{y}-{m:02d}" for m in range(1, 13)}

    # (2) Quartal: "1. Quartal 2025" / "4. Quartal 2025"
    m_q = re.fullmatch(r"([1-4])\.\s*quartal\s+(\d{4})", raw, flags=re.IGNORECASE)
    if m_q:
        q = int(m_q.group(1))
        y = m_q.group(2)
        start = (q - 1) * 3 + 1
        return {f"{y}-{m:02d}" for m in (start, start + 1, start + 2)}

    # (3) Monat ausgeschrieben: "Januar 2025" ... "Dezember 2025"
    m_mon = re.fullmatch(r"([A-Za-zÄÖÜäöüß]+)\s+(\d{4})", raw)
    if m_mon:
        name = m_mon.group(1).lower()
        name = name.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
        y = m_mon.group(2)
        if name in _MONTH_NAMES_DE:
            mnum = _MONTH_NAMES_DE[name]
            return {f"{y}-{mnum:02d}"}

    return set()

def build_feedback_block(mandant: str, timeframe: str) -> str:
    mandant = normalize_mandant(mandant)
    base = find_mandant_folder(mandant)

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


def rename_sent_suffix(path: str, ts: str) -> None:
    """Hängt _sent_<ts> vor die .png-Endung (Firefox-Suffixe bleiben erhalten)."""

    p = Path(path)
    new_name = f"{p.stem}_sent_{ts}{p.suffix}"
    new_path = p.with_name(new_name)
    os.rename(p, new_path)

def main():
    # Vorab: Pfade prüfen
    for p in (PATH_USTVA_BWA, PATH_FEEDBACK_SB, PATH_OPOS_TEMPLATE, PATH_EXCEL, FEEDBACK_ROOT):
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
    count_opos = 0
    summary_lines = []
    summary_lines_opos = []

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
            did_opos = False
            opos_attach_count = 0
            opos_pngs: list[str] = []

            has_feedback = to_bool(row.iloc[COL_FLAG_FEEDBACK])
            has_ustva    = to_bool(row.iloc[COL_FLAG_USTVA])
            has_opos     = to_bool(row.iloc[COL_FLAG_OPOS])
            if not (has_feedback or has_ustva or has_opos):
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

            # -----------------------
            # OPOS (Grundgerüst, kein Mailversand)
            # -----------------------
            if has_opos:
                mail_months = set(months_for_search(zeitraum))
                if not mail_months:
                    print(f"[WARN] Zeitraum für OPOS nicht erkannt: '{zeitraum}'")
                else:
                    folder = find_mandant_folder(mandant)
                    if not folder:
                        print(f"[WARN] Mandantenordner nicht gefunden für Mandant {mandant}")
                    else:
                        # PNGs im Mandantenordner sammeln (keine Unterordner)
                        candidates = [p for p in os.listdir(folder) if p.lower().endswith(".png")]
                        for fname in candidates:
                            if "_sent_" in fname.lower():
                                continue
                            # Zeitraum-Teil aus Dateiname extrahieren: Segment vor evtl. " (n)" / "_sent" / ".png"
                            stem = Path(fname).stem
                            stem = re.sub(r"\s*\(\d+\)$", "", stem)  # Firefox-Suffix entfernen
                            stem = re.sub(r"_sent.*$", "", stem, flags=re.IGNORECASE)
                            period = stem.rsplit("_", 1)[-1] if "_" in stem else stem
                            png_months = opos_period_to_months(period)
                            if not png_months:
                                print(f"[WARN] Zeitraum im Dateinamen nicht erkannt: '{fname}'")
                                continue
                            if png_months.issubset(mail_months):
                                opos_pngs.append(os.path.join(folder, fname))

                        if opos_pngs:
                            # OPOS-Mail erstellen (ohne Versand)
                            html_path = word_fill_to_html(word, PATH_OPOS_TEMPLATE, placeholders, tmpdir)
                            html = read_text_utf8(html_path)
                            html = ensure_utf8_meta(html)

                            mail = outlook.CreateItem(0)
                            mail.BodyFormat = 2
                            mail.HTMLBody = html
                            mail.To = email
                            mail.Subject = f"Offene Bankbewegungen f\u00FCr {display_timeframe(zeitraum)}"
                            mail.SendUsingAccount = acct

                            attached_success: list[Path] = []
                            for fpath in opos_pngs:
                                try:
                                    mail.Attachments.Add(fpath)
                                    attached_success.append(Path(fpath))
                                except Exception:
                                    print(f"[WARN] Fehler beim Anhängen von '{os.path.basename(fpath)}' an OPOS-Mail für Mandant {mandant}")

                            if attached_success:
                                mail.Save()
                                mail.Move(drafts)

                                sent_ts = datetime.now().strftime("%Y%m%d%H%M%S")
                                for p in attached_success:
                                    try:
                                        rename_sent_suffix(str(p), sent_ts)
                                    except Exception:
                                        print(f"[WARN] Fehler beim Umbenennen von '{p.name}'")

                                did_opos = True
                                count_opos += 1
                                opos_attach_count = len(attached_success)
                            else:
                                print(f"[WARN] Keine Anhänge für OPOS übernommen (Mandant {mandant}, Zeitraum {display_timeframe(zeitraum)})")
                        else:
                            print(f"[WARN] Keine passenden PNGs für OPOS gefunden (Mandant {mandant}, Zeitraum {display_timeframe(zeitraum)})")

            parts = []
            if did_fb:
                parts.append("Feedback")
            if did_ust:
                parts.append("UStVA/BWA")
            if did_opos:
                parts.append(f"OPOS mit {format_count(opos_attach_count, 'Anhang', 'Anhänge')}")
            if parts:
                summary_lines.append(f"- Für Mandant {mandant} im Zeitraum {display_timeframe(zeitraum)}: {', '.join(parts)}")

        # ANSI-Codes für Unterstreichung
        UNDERLINE = "\033[4m"
        RESET = "\033[0m"
        summary_counts = ", ".join([
            format_count(count_fb, "Feedback-Entwurf", "Feedback-Entwürfe"),
            format_count(count_u, "UStVA/BWA-Entwurf", "UStVA/BWA-Entwürfe"),
            format_count(count_opos, "OPOS-Mail", "OPOS-Mails"),
        ])
        print(f"\n{UNDERLINE}Erstellt: {summary_counts}:{RESET}")
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
    print("[INFO] Fenster schließt sich selbst in 30 Sekunden...")
    time.sleep(30)  # Zeit zum Lesen der Zusammenfassung