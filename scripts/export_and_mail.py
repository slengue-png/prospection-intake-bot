import os
import json
import requests
import pandas as pd
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook

EXPORT_BASE_URL = os.environ.get("EXPORT_BASE_URL")
EXPORT_TOKEN = os.environ.get("EXPORT_TOKEN")
BREVO_API_KEY = os.environ.get("BREVO_API_KEY")
MODE = os.environ.get("MODE")
RUN_DATE = os.environ.get("RUN_DATE")
AGENCY = os.environ.get("AGENCY")
INITIALS = os.environ.get("INITIALS")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")

TODAY = RUN_DATE if RUN_DATE else datetime.now().strftime("%Y-%m-%d")

MAIL_MAP = {
    "GR": {
        "JL": "sam9999@free.fr",
        "CZ": "slengue@gmail.com",
        "manager": "sam9999@free.fr",
    },
    "VR": {
        "JB": "sam9999@free.fr",
        "LB": "slengue@gmail.com",
        "manager": "sam9999@free.fr",
    },
    "GRS": {
        "PV": "sam9999@free.fr",
        "ST": "slengue@gmail.com",
        "manager": "sam9999@free.fr",
    },
    "SLS": {
        "AC": "sam9999@free.fr",
        "manager": "sam9999@free.fr",
    }
}

OWNER_EMAIL = "samuel.lengue@ras-interim.fr"

def fetch_data():
    url = f"{EXPORT_BASE_URL}/dump?date={TODAY}&kind=prospects"
    headers = {"X-Export-Token": EXPORT_TOKEN}
    r = requests.get(url, headers=headers)
    lines = r.text.strip().split("\n")
    return [json.loads(l) for l in lines if l.strip()]

def normalize_naf(value):
    return "".join(c for c in str(value).upper() if c.isalnum())

def send_mail(to_email, subject, body, attachment_path=None):
    url = "https://api.brevo.com/v3/smtp/email"
    headers = {
        "api-key": BREVO_API_KEY,
        "content-type": "application/json"
    }

    data = {
        "sender": {"name": "Prospection Bot", "email": OWNER_EMAIL},
        "to": [{"email": to_email}],
        "subject": subject,
        "htmlContent": body,
    }

    if attachment_path:
        with open(attachment_path, "rb") as f:
            import base64
            encoded = base64.b64encode(f.read()).decode()
        data["attachment"] = [{
            "name": os.path.basename(attachment_path),
            "content": encoded
        }]

    requests.post(url, headers=headers, json=data)

def generate_excel(records, filename):
    df = pd.DataFrame(records)
    df["NAF"] = df["naf"].apply(normalize_naf)
    df.to_excel(filename, index=False)

def download_card(file_id, dest_path):
    file_info = requests.get(
        f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile",
        json={"file_id": file_id}
    ).json()

    file_path = file_info["result"]["file_path"]

    file_url = f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_path}"
    r = requests.get(file_url)

    with open(dest_path, "wb") as f:
        f.write(r.content)

def session_mode():
    data = fetch_data()
    user_records = [
        d for d in data
        if d.get("agency") == AGENCY and d.get("initials") == INITIALS
    ]

    if not user_records:
        return

    filename = f"{AGENCY}_{INITIALS}_{TODAY}.xlsx"
    generate_excel(user_records, filename)

    recipient = MAIL_MAP.get(AGENCY, {}).get(INITIALS)
    if not recipient:
        return

    body = f"<h3>Prospection {INITIALS} - {TODAY}</h3><p>{len(user_records)} actions</p>"

    send_mail(recipient, f"Prospection {INITIALS} - {TODAY}", body, filename)

def daily_mode():
    data = fetch_data()

    grouped = defaultdict(list)
    for d in data:
        grouped[d.get("agency")].append(d)

    for agency, records in grouped.items():
        filename = f"{agency}_daily_{TODAY}.xlsx"
        generate_excel(records, filename)

        manager_email = MAIL_MAP.get(agency, {}).get("manager")
        if manager_email:
            send_mail(
                manager_email,
                f"Récap agence {agency} - {TODAY}",
                f"{len(records)} actions",
                filename
            )

    # Global summary
    total = len(data)
    filename = f"GLOBAL_{TODAY}.xlsx"
    generate_excel(data, filename)

    send_mail(
        OWNER_EMAIL,
        f"Récap global - {TODAY}",
        f"{total} actions au total",
        filename
    )

if MODE == "session":
    session_mode()

elif MODE == "daily":
    daily_mode()