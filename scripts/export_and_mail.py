import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime
from openpyxl import Workbook, load_workbook
import yaml
from pathlib import Path

# =========================
# CONFIG
# =========================

IMPORT_COLUMNS = [
    "DATE",
    "AGENCE",
    "COMMERCIAL",
    "SESSION",
    "ENTREPRISE",
    "VILLE",
    "PROSPECT",
    "CLIENT",
    "COMMANDE",
    "RESUME"
]

REPORTS_DIR = "reports"


# =========================
# HELPERS
# =========================

def load_config():
    with open("config.yml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def fetch_jsonl(url, export_token):
    req = urllib.request.Request(url)
    req.add_header("X-Export-Token", export_token)

    with urllib.request.urlopen(req, timeout=45) as resp:
        raw = resp.read().decode("utf-8", errors="replace").strip()

    if not raw:
        return []

    return [json.loads(line) for line in raw.splitlines() if line.strip()]


# =========================
# CUMUL MENSUEL
# =========================

def get_month_file(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    filename = f"{dt.year}-{dt.month:02d}_CUMUL.xlsx"
    return os.path.join(REPORTS_DIR, filename)


def ensure_month_file(path):
    if not os.path.exists(REPORTS_DIR):
        os.makedirs(REPORTS_DIR)

    if not os.path.exists(path):
        wb = Workbook()
        ws = wb.active
        ws.title = "JOURNAL"
        ws.append(IMPORT_COLUMNS)
        wb.save(path)


def append_rows_month_file(path, prospects, date_str):
    wb = load_workbook(path)
    ws = wb["JOURNAL"]

    for p in prospects:
        ws.append([
            date_str,
            p.get("agency",""),
            p.get("commercial",""),
            p.get("session",""),
            p.get("name",""),
            p.get("city",""),
            1,
            1 if p.get("client") else 0,
            1 if p.get("commande") else 0,
            p.get("resume","")
        ])

    wb.save(path)


# =========================
# STATS
# =========================

def compute_stats(prospects):
    stats = {}

    for p in prospects:
        ag = p.get("agency","")
        com = p.get("commercial","")

        if ag not in stats:
            stats[ag] = {"prospects":0,"clients":0,"commandes":0}

        stats[ag]["prospects"] += 1
        if p.get("client"):
            stats[ag]["clients"] += 1
        if p.get("commande"):
            stats[ag]["commandes"] += 1

    return stats


def format_stats_mail(stats, date):
    lines = [f"Résumé global prospection — {date}", ""]
    for ag, s in stats.items():
        tx = (s["commandes"]/s["prospects"]*100) if s["prospects"] else 0
        lines.append(
            f"{ag}: Prospects {s['prospects']} | "
            f"Clients {s['clients']} | "
            f"Commandes {s['commandes']} | "
            f"Tx {tx:.1f}%"
        )
    return "\n".join(lines)


# =========================
# BREVO
# =========================

def send_email_brevo(subject, body, to_list):
    api_key = os.environ["BREVO_API_KEY"]
    sender = os.environ["BREVO_SENDER_EMAIL"]

    payload = {
        "sender": {"email": sender, "name": "Prospection Bot"},
        "to": [{"email": x} for x in to_list],
        "subject": subject,
        "textContent": body,
    }

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://api.brevo.com/v3/smtp/email",
        data=data,
        method="POST"
    )
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", api_key)

    with urllib.request.urlopen(req, timeout=30) as resp:
        print("BREVO:", resp.status)


# =========================
# MAIN
# =========================

def main():
    cfg = load_config()

    worker = cfg["worker_base_url"].rstrip("/")
    export_token = os.environ["EXPORT_TOKEN"]
    date = datetime.utcnow().strftime("%Y-%m-%d")

    url = f"{worker}/dump?date={date}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)

    print("Prospects récupérés:", len(prospects))

    # ---- CUMUL MENSUEL
    month_file = get_month_file(date)
    ensure_month_file(month_file)
    append_rows_month_file(month_file, prospects, date)

    # ---- STATS JOUR
    stats = compute_stats(prospects)
    mail_body = format_stats_mail(stats, date)

    # ---- MAIL GLOBAL
    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date}",
        mail_body,
        cfg["global_to"]
    )


if __name__ == "__main__":
    main()
