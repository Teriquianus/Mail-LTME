# -*- coding: utf-8 -*-
from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd

import win32com.client as win32
from win32com.client import constants as c

# =========================
# KONFIG (Pfad & Konto)
# =========================
SHEET_NAME = "Übersicht"
TABLE_NAME = "Übersicht"
# 0-basierte Indizes entsprechend des Tabellenaufbaus
COL_MAILANS = 5
COL_EMAIL = 13
COL_FLAG_RUNDMAIL = 14

BASE_DIR = Path(__file__).resolve().parent
# Pfad zur Rundmail-Vorlage (hier anpassen, falls sich die Datei verschiebt)
PATH_RUNDMAIL_TEMPLATE = BASE_DIR / "Rundmail.docx"
PATH_EXCEL = r"C:\Users\AlexanderHaller\Unternehmenskompass GmbH\Unternehmenskompass - CRM\LTME\LTME Working.xlsm"
PATH_RUNDMAIL_ATTACHMENTS = Path(r"C:\Users\AlexanderHaller\Unternehmenskompass GmbH\Unternehmenskompass - CRM\LTME\Automation\Anhang Rundmail")
PATH_RUNDMAIL_ARCHIVE = PATH_RUNDMAIL_ATTACHMENTS / "Archiv"
SMTP_INFO = "info@ltme-consulting.de"

_SENT_FLAG_RE = re.compile(r"_sent_on_\d{4}_\d{2}_\d{2}($|_)", flags=re.IGNORECASE)

# =========================
# Hilfsfunktionen
# =========================
def to_bool_like(value) -> bool:
    if isinstance(value, bool):
        return value
    normalized = str(value).strip().upper()
    return normalized in {"WAHR", "TRUE", "JA", "X", "1"}


def format_count(count: int, singular: str, plural: str) -> str:
    return f"{count} {singular if count == 1 else plural}"


def read_text_utf8(path: str) -> str:
    with open(path, "r", encoding="utf-8") as file:
        return file.read()


def ensure_utf8_meta(html: str) -> str:
    if "charset" in html.lower():
        return html
    return html.replace(
        "<head>",
        '<head><meta http-equiv="Content-Type" content="text/html; charset=utf-8">',
        1,
    )


def pop_heading_text(doc) -> str | None:
    heading_names = {"ÜBERSCHRIFT 1", "HEADING 1"}
    for para in doc.Paragraphs:
        style_name = getattr(para.Style, "NameLocal", None)
        if not style_name:
            style_name = getattr(para.Style, "Name", "")
        style_name = (style_name or "").strip().upper()
        if style_name in heading_names:
            raw_text = para.Range.Text
            para.Range.Delete()
            return raw_text.replace("\r", "").replace("\n", "").strip()
    return None


def collect_recipients(df: pd.DataFrame) -> list[dict[str, str]]:
    recipients: list[dict[str, str]] = []
    max_idx = max(COL_MAILANS, COL_EMAIL, COL_FLAG_RUNDMAIL)
    if df.shape[1] <= max_idx:
        raise ValueError(
            f"Tabelle '{SHEET_NAME}' hat zu wenige Spalten (erwartet >= {max_idx + 1})."
        )

    for _, row in df.iterrows():
        if not to_bool_like(row.iloc[COL_FLAG_RUNDMAIL]):
            continue

        email = row.iloc[COL_EMAIL]
        if pd.isna(email) or not str(email).strip():
            continue

        mailansprache = row.iloc[COL_MAILANS]
        recipients.append(
            {
                "email": str(email).strip(),
                "mailansprache": "" if pd.isna(mailansprache) else str(mailansprache).strip(),
            }
        )

    return recipients


def gather_rundmail_attachments() -> tuple[list[Path], list[Path]]:
    attachments: list[Path] = []
    flagged: list[Path] = []
    if not PATH_RUNDMAIL_ATTACHMENTS.is_dir():
        return attachments, flagged

    for entry in sorted(PATH_RUNDMAIL_ATTACHMENTS.iterdir()):
        if entry.is_dir():
            continue
        if _SENT_FLAG_RE.search(entry.stem):
            flagged.append(entry)
            continue
        attachments.append(entry)
    return attachments, flagged


def archive_attachments(files: list[Path], date_tag: str) -> None:
    if not files:
        return
    PATH_RUNDMAIL_ARCHIVE.mkdir(parents=True, exist_ok=True)
    for src in files:
        target = PATH_RUNDMAIL_ARCHIVE / f"{src.stem}_sent_on_{date_tag}{src.suffix}"
        suffix_counter = 1
        while target.exists():
            target = PATH_RUNDMAIL_ARCHIVE / f"{src.stem}_sent_on_{date_tag}_{suffix_counter}{src.suffix}"
            suffix_counter += 1
        src.rename(target)


# =========================
# Word → HTML (UTF-8)
# =========================
def render_personalized_html(word_app, placeholders: dict[str, str], tmpdir: str) -> str:
    tpl = word_app.Documents.Open(FileName=str(PATH_RUNDMAIL_TEMPLATE), ReadOnly=True)
    doc = word_app.Documents.Add()
    doc.Content.FormattedText = tpl.Content.FormattedText
    pop_heading_text(doc)

    finder = doc.Content.Find
    finder.ClearFormatting()
    finder.Replacement.ClearFormatting()
    for token, value in placeholders.items():
        finder.Execute(FindText=token, ReplaceWith=value, Replace=c.wdReplaceAll)

    html_path = os.path.join(tmpdir, f"rundmail_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.htm")
    doc.WebOptions.Encoding = 65001
    doc.WebOptions.AllowPNG = True
    doc.SaveAs2(FileName=html_path, FileFormat=c.wdFormatHTML)

    tpl.Close(False)
    doc.Close(False)

    html = read_text_utf8(html_path)
    return ensure_utf8_meta(html)


def extract_subject_from_template(word_app) -> str:
    tpl = word_app.Documents.Open(FileName=str(PATH_RUNDMAIL_TEMPLATE), ReadOnly=True)
    doc = word_app.Documents.Add()
    doc.Content.FormattedText = tpl.Content.FormattedText
    subject = pop_heading_text(doc)
    doc.Close(False)
    tpl.Close(False)
    return subject or "Rundmail"


# =========================
# Outlook Helpers
# =========================
def get_account(ns, smtp: str):
    target = smtp.strip().lower()
    for account in ns.Accounts:
        try:
            if str(account.SmtpAddress).strip().lower() == target:
                return account
        except Exception:  # pragma: no cover
            pass
    return None


# =========================
# Excel-Handling
# =========================
def read_overview_table(excel_path: Path) -> pd.DataFrame:
    excel_app = win32.gencache.EnsureDispatch("Excel.Application")
    excel_app.Visible = False
    wb = None
    try:
        wb = excel_app.Workbooks.Open(str(excel_path), ReadOnly=True)
        sheet = wb.Sheets(SHEET_NAME)
        table = sheet.ListObjects(TABLE_NAME)
        headers = [str(col.Name).strip() if col.Name else "" for col in table.ListColumns]

        data_rows = []
        data_range = table.DataBodyRange
        if data_range is not None:
            raw = data_range.Value
            if isinstance(raw, tuple) and raw:
                if isinstance(raw[0], tuple):
                    data_rows = [list(row) for row in raw]
                else:
                    data_rows = [list(raw)]
            elif raw is not None:
                data_rows = [[raw]]

        df = pd.DataFrame(data_rows, columns=headers)
    finally:
        if wb is not None:
            wb.Close(False)
        try:
            excel_app.Quit()
        except Exception:
            pass

    return df


# =========================
# Hauptlogik (main)
# =========================
def main() -> None:
    excel_path = Path(PATH_EXCEL)
    # print(f"[INFO] Verwende fest codierten Excel-Pfad: {excel_path}")

    attachments, flagged = gather_rundmail_attachments()
    if not PATH_RUNDMAIL_ATTACHMENTS.exists():
        print(f"[INFO] Rundmail-Anhang-Ordner nicht gefunden: {PATH_RUNDMAIL_ATTACHMENTS}")
    else:
        if attachments:
            names = ", ".join(p.name for p in attachments)
            print(f"[INFO] Rundmail-Anhänge gefunden ({len(attachments)}): {names}")
        else:
            print(f"[INFO] Rundmail-Anhang-Ordner ist leer: {PATH_RUNDMAIL_ATTACHMENTS}")
        if flagged:
            names = ", ".join(p.name for p in flagged)
            print(f"[WARN] Flagged-Anhänge übersprungen (nicht im Archiv): {names}")

    if not excel_path.is_file():
        print(f"[WARN] Excel-Datei nicht gefunden: {excel_path}")
        sys.exit(1)

    try:
        df = read_overview_table(excel_path)
    except Exception as exc:  # pragma: no cover - Fehlermeldung
        print(f"[WARN] Tabellenbereich konnte nicht geladen werden: {exc}")
        sys.exit(1)

    if df.empty:
        print(f"[WARN] Tabelle '{SHEET_NAME}' enthält keine Zeilen.")

    try:
        recipients = collect_recipients(df)
    except ValueError as exc:
        print(f"[WARN] {exc}")
        sys.exit(1)

    count = len(recipients)
    print()
    print(f"[INFO] Gefilterte Rundmail-Empfänger: {count}")
    for idx, entry in enumerate(recipients, start=1):
        ansprache = entry["mailansprache"] or "<leer>"
        print(f"{idx}. {entry['email']} – {ansprache}")
    print()

    if count == 0:
        print("[INFO] Keine Rundmail-Entwürfe nötig. Script beendet.")
        return

    if not PATH_RUNDMAIL_TEMPLATE.is_file():
        print(f"[WARN] Rundmail-Vorlage nicht gefunden: {PATH_RUNDMAIL_TEMPLATE}")
        sys.exit(1)

    word_app = None
    outlook_app = None
    tmpdir = tempfile.mkdtemp(prefix="rundmail_")
    created = 0

    try:
        word_app = win32.gencache.EnsureDispatch("Word.Application")
        word_app.Visible = False
        subject = extract_subject_from_template(word_app)

        outlook_app = win32.gencache.EnsureDispatch("Outlook.Application")
        namespace = outlook_app.GetNamespace("MAPI")
        account = get_account(namespace, SMTP_INFO)
        if account is None:
            print(f"[WARN] Outlook-Account '{SMTP_INFO}' nicht gefunden.")
            sys.exit(1)

        drafts = account.DeliveryStore.GetDefaultFolder(16)

        for recipient in recipients:
            placeholders = {
                "{{Mailansprache}}": recipient["mailansprache"],
                "{{Vorname}}": recipient["mailansprache"],
                "{{Email}}": recipient["email"],
            }
            html = render_personalized_html(word_app, placeholders, tmpdir)
            mail_item = outlook_app.CreateItem(0)
            mail_item.BodyFormat = 2
            mail_item.HTMLBody = html
            mail_item.To = recipient["email"]
            mail_item.Subject = subject
            mail_item.SendUsingAccount = account
            for attachment in attachments:
                mail_item.Attachments.Add(str(attachment))
            mail_item.Save()
            mail_item.Move(drafts)
            created += 1

        print(f"[INFO] Rundmail-Entwürfe erstellt: {format_count(created, 'Entwurf', 'Entwürfe')}")
        try:
            if attachments:
                try:
                    archive_attachments(attachments, datetime.now().strftime("%Y_%m_%d"))
                except Exception as exc:
                    print(f"[WARN] Rundmail-Anhänge konnten nicht archiviert werden: {exc}")
            os.startfile(str(excel_path))
        except Exception as exc:
            print(f"[WARN] Excel konnte nicht erneut geöffnet werden: {exc}")

    except Exception as exc:  # pragma: no cover - COM-Fehler
        print(f"[WARN] Rundmail konnte nicht vollständig ausgeführt werden: {exc}")
        sys.exit(1)
    finally:
        if word_app is not None:
            try:
                word_app.Quit(False)
            except Exception:
                pass
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
