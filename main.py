#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import json
import logging
import pdfplumber
import re
from datetime import datetime

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ---------------- CONFIG ----------------
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_FILE = "token.json"
USER_SHEETS_FILE = "user_sheets.json"

# Ensure user_sheets storage exists
if not os.path.exists(USER_SHEETS_FILE):
    with open(USER_SHEETS_FILE, "w", encoding="utf-8") as f:
        json.dump({}, f, ensure_ascii=False)

# -------------- Google OAuth / Service --------------
def get_google_service(interactive=True):
    """
    Returns Google Sheets service object.
    If token.json is missing or expired, starts InstalledAppFlow (interactive).
    interactive=False will not prompt for code and will return None if no token.
    """
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception as e:
            logging.warning("Failed loading token file: %s", e)
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                logging.warning("Failed to refresh token: %s", e)
        else:
            if not interactive:
                logging.error("No valid credentials and interactive=False")
                return None
            # Need user consent
            client_id = os.environ.get("GOOGLE_CLIENT_ID")
            client_secret = os.environ.get("GOOGLE_CLIENT_SECRET")
            if not client_id or not client_secret:
                logging.error("GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET not set in env")
                return None

            flow = InstalledAppFlow.from_client_config(
                {
                    "installed": {
                        "client_id": client_id,
                        "client_secret": client_secret,
                        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token",
                    }
                },
                SCOPES,
            )

            auth_url, _ = flow.authorization_url(prompt="consent")
            # Show URL in stdout/logs
            print("\n--- GOOGLE AUTH REQUIRED ---\n")
            print("Open this URL in a browser and authorize the app:\n")
            print(auth_url)
            print("\nAfter allowing access, copy the authorization code and paste it here.\n")
            code = input("Authorization code: ").strip()
            flow.fetch_token(code=code)
            creds = flow.credentials

        # Save token
        try:
            with open(TOKEN_FILE, "w", encoding="utf-8") as token_out:
                token_out.write(creds.to_json())
        except Exception as e:
            logging.warning("Could not save token file: %s", e)

    # Build service
    try:
        service = build("sheets", "v4", credentials=creds)
        return service
    except HttpError as e:
        logging.error("Google API error: %s", e)
        return None


# -------------- User sheet mapping --------------
def load_user_sheets():
    try:
        with open(USER_SHEETS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_user_sheets(data):
    with open(USER_SHEETS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_user_sheet_id(user_id: int):
    data = load_user_sheets()
    return data.get(str(user_id))


def set_user_sheet_id(user_id: int, sheet_id: str):
    data = load_user_sheets()
    data[str(user_id)] = sheet_id
    save_user_sheets(data)


# -------------- Google Sheets helpers --------------
def get_sheet_names(service, spreadsheet_id):
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]


def create_patient_sheet(service, spreadsheet_id, sheet_name):
    body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {"title": sheet_name, "gridProperties": {"rowCount": 2000, "columnCount": 50}}
                }
            }
        ]
    }
    service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
    # Initialize header row if empty: A1="Показатель", B1="Референс"
    try:
        service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1:B1",
            valueInputOption="USER_ENTERED",
            body={"values": [["Показатель", "Референс"]]},
        ).execute()
    except Exception as e:
        logging.debug("Failed to initialize header: %s", e)


def ensure_patient_sheet(service, spreadsheet_id, sheet_name):
    try:
        sheets = get_sheet_names(service, spreadsheet_id)
    except Exception as e:
        logging.error("Cannot get sheet names: %s", e)
        raise
    if sheet_name not in sheets:
        create_patient_sheet(service, spreadsheet_id, sheet_name)


def read_column_a(service, spreadsheet_id, sheet_name):
    res = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A:A").execute()
    values = res.get("values", [])
    return [r[0] for r in values] if values else []


def append_rows(service, spreadsheet_id, sheet_name, rows):
    if not rows:
        return
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:A",
        valueInputOption="USER_ENTERED",
        body={"values": rows},
    ).execute()


def ensure_rows_for_analytes(service, spreadsheet_id, sheet_name, analytes):
    existing = read_column_a(service, spreadsheet_id, sheet_name)
    missing = [a for a in analytes if a not in existing]
    if missing:
        rows = [[m, ""] for m in missing]
        append_rows(service, spreadsheet_id, sheet_name, rows)


def column_number_to_letter(n):
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result


def get_next_date_column(service, spreadsheet_id, sheet_name, date_str):
    # Read header row
    res = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=f"{sheet_name}!1:1").execute()
    header = res.get("values", [[]])[0]
    # header may be empty
    if date_str in header:
        idx = header.index(date_str) + 1
        return column_number_to_letter(idx)
    # append at the end
    idx = len(header) + 1
    col_letter = column_number_to_letter(idx)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{col_letter}1",
        valueInputOption="USER_ENTERED",
        body={"values": [[date_str]]},
    ).execute()
    return col_letter


def get_row_for_analyte(service, spreadsheet_id, sheet_name, analyte):
    col_a = read_column_a(service, spreadsheet_id, sheet_name)
    for i, name in enumerate(col_a, start=1):
        if name.strip().lower() == analyte.strip().lower():
            return i
    return None


def write_values(service, spreadsheet_id, sheet_name, col_letter, values_dict):
    """
    Writes values_dict: {analyte_name: value} into cells at column col_letter.
    """
    # For each analyte find its row and update single cell
    for analyte, value in values_dict.items():
        row = get_row_for_analyte(service, spreadsheet_id, sheet_name, analyte)
        if row is None:
            # If row missing, append at bottom
            append_rows(service, spreadsheet_id, sheet_name, [[analyte, ""]])
            col_a = read_column_a(service, spreadsheet_id, sheet_name)
            row = len(col_a)
        cell = f"{col_letter}{row}"
        try:
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!{cell}",
                valueInputOption="USER_ENTERED",
                body={"values": [[str(value)]]},
            ).execute()
        except Exception as e:
            logging.warning("Failed to write %s -> %s at %s: %s", analyte, value, cell, e)


# -------------- PDF parsing (heuristic) --------------
def parse_pdf(file_path):
    """Return (full_name, date_str, analytes_dict)
    analytes_dict: { 'Гемоглобин': {'value':'155','ref':'135–169'}, ... }
    """
    text_parts = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                text_parts.append(page_text)
    except Exception as e:
        logging.error("pdfplumber failed: %s", e)
        # fallback: try reading file bytes and attempt OCR? (not implemented here)
        raise

    joined = "\n".join([p for p in text_parts if p])

    # Normalize
    lines = [ln.strip() for ln in joined.splitlines() if ln.strip()]
    joined_text = "\n".join(lines)

    # FIO detection: try several patterns
    fio = None
    m = re.search(r"Фамилия[:\s]*([А-ЯЁ][а-яё\-]+)", joined_text, re.IGNORECASE)
    if m:
        surname = m.group(1).strip()
        # try to find full name nearby
        m2 = re.search(rf"Фамилия[:\s]*{re.escape(surname)}[^\n]*\n.*?([\wЁёА-Яа-я]+\s+[\wЁёА-Яа-я]+)", joined_text, re.IGNORECASE)
        fio = (surname + " " + m2.group(1).strip()) if m2 else surname
    if not fio:
        # fallback to pattern "ФИО: Иванов Иван Иванович"
        m = re.search(r"ФИО[:\s]+([\wЁёА-Яа-я]+\s+[\wЁёА-Яа-я]+\s+[\wЁёА-Яа-я]+)", joined_text)
        if m:
            fio = m.group(1).strip()
    if not fio:
        # take first line with 2-3 cyrillic words
        for ln in lines[:12]:
            parts = re.findall(r"[А-ЯЁ][а-яё]+", ln)
            if len(parts) >= 2:
                fio = ln.strip()
                break
    full_name = fio if fio else "Пациент"

    # Date detection (prefer sample date)
    date_match = re.search(r"Дата взятия образца[:\s]*([0-3]?\d[.\-/][01]?\d[.\-/]\d{4})", joined_text)
    if not date_match:
        date_match = re.search(r"(\d{2}[.\-/]\d{2}[.\-/]\d{4})", joined_text)
    date_str = date_match.group(1) if date_match else datetime.now().strftime("%d.%m.%Y")

    # Extract analytes: lines that contain a name, a numeric value and a reference or range
    analytes = {}
    # Loosen matching: name then number then possible units then ref
    pattern = re.compile(
        r"^([А-ЯЁа-яA-Za-z0-9\s\-\(\)\/%µμ]+?)\s+([<>]?\d+[.,]?\d*)\s*(?:[^\d\n]{0,8})\s*([\d.,<>]+–[\d.,<>]+|<\d+|>?\d+|отрицательно|не обнаружено)?",
        re.IGNORECASE,
    )

    for ln in lines:
        m = pattern.match(ln)
        if m:
            name = m.group(1).strip()
            val = m.group(2).replace(",", ".").strip()
            ref = m.group(3) or ""
            ref = ref.strip().replace(" ", "")
            analytes[name] = {"value": val, "ref": ref}

    # If analytes empty, try to find "Name ... value ... range" inside text by tokens
    if not analytes:
        tokens = re.findall(r"([А-ЯЁа-яA-Za-z\-\s]{3,}?)[:\s]{1,3}([0-9]+(?:[.,][0-9]+)?)\s*([^\d\n]*)", joined_text)
        for t in tokens:
            name = t[0].strip()
            val = t[1].replace(",", ".")
            analytes[name] = {"value": val, "ref": ""}

    return full_name, date_str, analytes


# -------------- Telegram handlers --------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 Привет! Отправь PDF с анализами. \n"
        "Перед использованием укажи таблицу: /set_sheet <SPREADSHEET_ID> \n"
        "Пример SPREADSHEET_ID: это длинная строка в URL после /d/"
    )


async def set_sheet(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not context.args or len(context.args) < 1:
        await update.message.reply_text("Использование: /set_sheet <SPREADSHEET_ID>")
        return
    sheet_id = context.args[0].strip()
    set_user_sheet_id(user_id, sheet_id)
    await update.message.reply_text(f"✔ Таблица сохранена: {sheet_id}")


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    sheet_id = get_user_sheet_id(user_id)
    if not sheet_id:
        await update.message.reply_text("Вы не указали Google Sheet. Сделайте: /set_sheet <ID>")
        return

    doc = update.message.document
    if not doc:
        await update.message.reply_text("Пожалуйста, отправьте PDF-файл (document).")
        return
    if not doc.file_name.lower().endswith(".pdf"):
        await update.message.reply_text("Пожалуйста, отправьте файл в формате PDF.")
        return

    await update.message.reply_text("📥 Скачиваю файл...")
    file = await doc.get_file()
    bio = io.BytesIO()
    await file.download_to_memory(out=bio)
    pdf_bytes = bio.getvalue()

    tmp_path = f"/tmp/{doc.file_unique_id}.pdf"
    try:
        with open(tmp_path, "wb") as f:
            f.write(pdf_bytes)
    except Exception as e:
        logging.error("Failed write tmp file: %s", e)
        await update.message.reply_text("Ошибка записи временного файла.")
        return

    await update.message.reply_text("🔎 Парсю PDF...")
    try:
        full_name, date_str, analytes = parse_pdf(tmp_path)
    except Exception as e:
        logging.exception("Парсер упал: %s", e)
        await update.message.reply_text("Ошибка при разборе PDF.")
        return

    await update.message.reply_text(f"Пациент: {full_name}\nДата: {date_str}\nПоказателей: {len(analytes)}")

    # Get Google service (interactive if needed)
    service = get_google_service(interactive=True)
    if not service:
        await update.message.reply_text("Ошибка авторизации Google.")
        return

    # Ensure sheet
    try:
        ensure_patient_sheet(service, sheet_id, full_name)
    except Exception as e:
        logging.exception("Ошибка проверки/создания листа: %s", e)
        await update.message.reply_text("Ошибка работы с Google Sheets (создание листа).")
        return

    # Ensure rows
    try:
        ensure_rows_for_analytes(service, sheet_id, full_name, analytes.keys())
    except Exception as e:
        logging.exception("Ошибка при создании строк анализов: %s", e)
        await update.message.reply_text("Ошибка при создании строк анализов.")
        return

    # Date column
    try:
        col_letter = get_next_date_column(service, sheet_id, full_name, date_str)
    except Exception as e:
        logging.exception("Ошибка при получении колонки даты: %s", e)
        await update.message.reply_text("Ошибка при создании колонки с датой.")
        return

    # Prepare values dict mapping analyte names (exact match) -> values
    values_for_write = {}
    # We attempt to match analytes keys to sheet row names case-insensitively
    sheet_rows = read_column_a(service, sheet_id, full_name)
    lower_map = {r.strip().lower(): r for r in sheet_rows}
    for name, info in analytes.items():
        key = name.strip().lower()
        target_name = lower_map.get(key, None)
        if target_name:
            values_for_write[target_name] = info["value"]
        else:
            # fallback: write under original name (it will append)
            values_for_write[name] = info["value"]

    # Write values
    try:
        write_values(service, sheet_id, full_name, col_letter, values_for_write)
    except Exception as e:
        logging.exception("Ошибка при записи значений: %s", e)
        await update.message.reply_text("Ошибка при записи значений в таблицу.")
        return

    await update.message.reply_text("✅ Данные записаны в Google Sheet.")


# -------------- Run bot --------------
def main():
    token = os.environ.get("TELEGRAM_BOT_TOKEN") or os.environ.get("BOT_TOKEN")
    if not token:
        print("ERROR: TELEGRAM_BOT_TOKEN or BOT_TOKEN not set in environment")
        return

    app = ApplicationBuilder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("set_sheet", set_sheet))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_pdf))

    print("Bot started")
    app.run_polling(stop_signals=None)


if __name__ == "__main__":
    main()
