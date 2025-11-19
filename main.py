import os
import json
import logging
import pdfplumber
import re
from datetime import datetime

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from telegram import Update, InputFile
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ========================
#       CONFIG
# ========================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

# Google OAuth scopes
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# File that will store OAuth token in Railway
TOKEN_FILE = "token.json"

# File where we store per-user Google Sheet IDs
USER_SHEETS_FILE = "user_sheets.json"


# Make sure file exists
if not os.path.exists(USER_SHEETS_FILE):
    with open(USER_SHEETS_FILE, "w", encoding="utf-8") as f:
        f.write("{}")
# ======================================
#     GOOGLE OAUTH – AUTHENTICATION
# ======================================

def get_google_service():
    """
    Initializes Google Sheets API service.
    Uses OAuth installed app flow.
    Railway will keep token.json between restarts.
    """
    creds = None

    # Load token if exists
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    # If no token or token expired → require login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # OAuth via InstalledAppFlow
            flow = InstalledAppFlow.from_client_config(
                {
                    "installed": {
                        "client_id": os.environ.get("GOOGLE_CLIENT_ID"),
                        "client_secret": os.environ.get("GOOGLE_CLIENT_SECRET"),
                        "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"],
                        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                        "token_uri": "https://oauth2.googleapis.com/token"
                    }
                },
                SCOPES
            )

            # Generate URL for user
            auth_url, _ = flow.authorization_url(prompt="consent")

            print("\n================ GOOGLE AUTH REQUIRED ================\n")
            print("👉 Open this URL in a browser:")
            print(auth_url)
            print("\nThen paste the authorization code below.\n")
            code = input("Authorization code: ")

            flow.fetch_token(code=code)
            creds = flow.credentials

        # Save token
        with open(TOKEN_FILE, "w", encoding="utf-8") as token:
# ======================================
#          USER SHEET MANAGEMENT
# ======================================

def load_user_sheets():
    with open(USER_SHEETS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


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


# ======================================
#      GOOGLE SHEET: CREATE / WRITE
# ======================================

def create_patient_sheet(service, spreadsheet_id, sheet_name):
    """
    Creates sheet inside the spreadsheet for a patient (full FIO)
    """
    requests = [
        {
            "addSheet": {
                "properties": {
                    "title": sheet_name
                }
            }
        }
    ]

    body = {"requests": requests}
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body=body
    ).execute()


def get_sheet_names(service, spreadsheet_id):
    """
    Returns list of sheet names in spreadsheet
    """
    meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return [s["properties"]["title"] for s in meta["sheets"]]


def ensure_patient_sheet(service, spreadsheet_id, sheet_name):
    """
    Ensures sheet with patient name exists.
    If not — create.
    """
    sheets = get_sheet_names(service, spreadsheet_id)
    if sheet_name not in sheets:
        create_patient_sheet(service, spreadsheet_id, sheet_name)


# ======================================
#        WRITE ANALYSES TO SHEET
# ======================================

def ensure_rows_for_analytes(service, spreadsheet_id, sheet_name, analytes):
    """
    Ensures rows exist for each analyte.
    Row format:
        A: analyte name
        B: reference (min–max)
        C..: dates (values)
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!A:A"
    ).execute()

    existing = [row[0] for row in result.get("values", [])]

    missing = [a for a in analytes if a not in existing]

    if missing:
        body = {
            "values": [[m, ""]] for m in missing  # create row with empty reference
        }
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A:A",
            valueInputOption="USER_ENTERED",
            body={"values": [[m, ""] for m in missing]}
        ).execute()


def write_values(service, spreadsheet_id, sheet_name, date_col, values_dict):
    """
    Writes analyte values into the date column.
    values_dict = {analyte: value}
    """
    write_data = []
    for analyte, value in values_dict.items():
        write_data.append({
            "range": f"{sheet_name}!{date_col}:{date_col}",
            "values": [[str(value)]]
        })

    service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": write_data
        }
    ).execute()


def get_next_date_column(service, spreadsheet_id, sheet_name, date_str):
    """
    Ensures date column exists. Returns column letter.
    Sheet format:
        A = analyte
        B = reference
        C.. = dates
    """
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!1:1"
    ).execute()

    header = result.get("values", [[]])[0]

    if date_str in header:
        col_index = header.index(date_str) + 1
        return column_number_to_letter(col_index)

    # append date column
    col_index = len(header) + 1
    col_letter = column_number_to_letter(col_index)

    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{sheet_name}!{col_letter}1",
        valueInputOption="USER_ENTERED",
        body={"values": [[date_str]]}
    ).execute()

    return col_letter


def column_number_to_letter(n):
    """Convert 1-indexed number to Excel column letter."""
    result = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        result = chr(65 + r) + result
    return result
# ======================================
#             PDF PARSER
# ======================================

def parse_pdf(file_path):
    """
    Parses a PDF with medical analyses and returns:
    - full_name
    - date
    - analytes: { analyte_name: {"value":..., "ref":...} }
    """

    with pdfplumber.open(file_path) as pdf:
        text = "\n".join([page.extract_text() or "" for page in pdf.pages])

    # Normalize
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    joined = "\n".join(lines)

    # ==============================
    # Extract FIO
    # ==============================
    fio_pattern = r"ФИО[:\s]+([\wЁёА-Яа-я]+\s+[\wЁёА-Яа-я]+\s+[\wЁёА-Яа-я]+)"
    fio_match = re.search(fio_pattern, joined)
    full_name = fio_match.group(1).strip() if fio_match else "Пациент"

    # ==============================
    # Extract date
    # ==============================
    date_pattern = r"(\d{2}\.\d{2}\.\d{4})"
    date_match = re.search(date_pattern, joined)
    date_str = date_match.group(1) if date_match else datetime.now().strftime("%d.%m.%Y")

    # ==============================
    # Extract analytes
    # ==============================
    analytes = {}

    # Pattern like:
    # Гемоглобин 155 г/л 135–169
    analyte_pattern = re.compile(
        r"([A-Za-zА-Яа-яЁё\s\-()%+]+?)\s+([\d.,<>]+)\s+[^\s]+\s+([\d.,<>]+–[\d.,<>]+|<\d+|>?\d+|отрицательно|не обнаружено)",
        re.IGNORECASE
    )

    for match in analyte_pattern.finditer(joined):
        analyte = match.group(1).strip()
        value = match.group(2).strip()
        ref = match.group(3).strip()

        # Normalize reference
        ref = ref.replace(" ", "")
        ref = ref.replace("отрицательно", "отрицательно")
        ref = ref.replace("необнаружено", "не обнаружено")

        analytes[analyte] = {
            "value": value,
            "ref": ref
        }

    return full_name, date_str, analytes
# ======================================
#            TELEGRAM BOT
# ======================================

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

    await update.message.reply_text(
        "👋 Привет! Я бот для анализа медицинских PDF.\n\n"
        "Отправь мне PDF с анализами — я распознаю:\n"
        "• ФИО пациента\n"
        "• Дату анализа\n"
        "• Все показатели\n"
        "• Референсные значения\n\n"
        "И загружу всё в твою Google-таблицу.\n\n"
        "Чтобы установить таблицу, введи:\n"
        "/set_sheet <GoogleSheetID>\n\n"
    )


async def set_sheet(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """User installs Google Sheet ID."""
    user_id = update.effective_user.id

    if len(context.args) != 1:
        await update.message.reply_text("Использование:\n/set_sheet <ID таблицы>")
        return

    sheet_id = context.args[0].strip()

    set_user_sheet_id(user_id, sheet_id)

    await update.message.reply_text(
        f"✔ Google-таблица установлена!\nID: {sheet_id}"
    )


async def handle_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles PDF upload."""
    user_id = update.effective_user.id
    sheet_id = get_user_sheet_id(user_id)

    if not sheet_id:
        await update.message.reply_text(
            "❗ Вы ещё не указали Google-таблицу.\n"
            "Сделайте это командой:\n/set_sheet <ID>"
        )
        return

    # Download file
    pdf_file = await update.message.document.get_file()
    file_path = f"/tmp/{pdf_file.file_unique_id}.pdf"
    await pdf_file.download_to_drive(file_path)

    await update.message.reply_text("📄 Обрабатываю PDF…")

    # Parse PDF
    full_name, date_str, analytes = parse_pdf(file_path)

    # Google service
    service = get_google_service()
    if not service:
        await update.message.reply_text("Ошибка Google OAuth.")
        return

    # Patient sheet
    ensure_patient_sheet(service, sheet_id, full_name)

    # Ensure analytes rows
    analyte_names = list(analytes.keys())
    ensure_rows_for_analytes(service, sheet_id, full_name, analyte_names)

    # Column for date
    date_col = get_next_date_column(service, sheet_id, full_name, date_str)

    # Values → dict {analyte → value}
    values_dict = {k: v["value"] for k in analytes.values()}

    # Write values
    write_values(service, sheet_id, full_name, date_col, values_dict)

    await update.message.reply_text(
        f"✔ Готово!\n"
        f"Пациент: *{full_name}*\n"
        f"Дата: *{date_str}*\n\n"
        f"Анализы успешно добавлены в Google-таблицу.",
        parse_mode="Markdown"
    )
