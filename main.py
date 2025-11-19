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
