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
