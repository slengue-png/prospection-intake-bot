import os, io, json, smtplib, ssl
from datetime import datetime
from email.message import EmailMessage
from openpyxl import Workbook
import yaml
import urllib.request
import urllib.error

IMPORT_COLUMNS = [
    "NOM","ADRESSE","CODE POSTAL","VILLE","TELEPHONE","TELEPHONE 2","MAIL",
    "SIRET","NAF","SITE WEB","Contact: civilité","Contact : prénom","Contact : nom",
    "RESUME ENTRETIEN","COMMANDE"
]

BROWSER_HEADERS = {
    # ⚠️ important : Cloudflare peut bloquer Python-urllib -> on mime Chrome
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "close",
}

def load_config(path="config.yml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def fetch_jsonl(url: str, export_token: str) -> list[dict]:
    print(f"[FETCH] {url}")

    req = urllib.request.Request(url)

    # ✅ headers anti-1010
    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)

    # ✅ auth header pour /dump
    req.add_header("X-Export-Token", export_token)

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8", errors="replace").strip()
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTPError {e.code} on {url} -> {body[:300]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"URLError on {url} -> {e.reason}")
    except Exception as e:
        raise RuntimeError(f"Fetch error on {url} -> {e}")

    if not raw:
        return []

    rows = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            rows.append(json.loads(line))
        except:
            pass
    return rows

def build_excel(records: list[dict]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "IMPORT"
    ws.append(IMPORT_COLUMNS)

    for r in records:
        ws.append([
            r.get("name",""),
            "", "",  # ADRESSE, CODE POSTAL (complétés plus tard)
            r.get("city",""),
            r.get("phone",""),
            "",      # TELEPHONE 2
            r.get("email",""),
            "", "", "",  # SIRET, NAF, SITE WEB
            "", "", r.get("interlocuteur",""),
            r.get("resume",""),
            r.get("commande",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

def send_email_gmail(subject, body, to_list, attachments):
    gmail_user = os.environ["GMAIL_USER"]
    gmail_app_password = os.environ["GMAIL_APP_PASSWORD"]

    msg = EmailMessage()
    msg["From"] = gmail_user
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject
    msg.set_content(body)

    for filename, content_bytes in attachments:
        msg.add_attachment(
            content_bytes,
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
        )

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as smtp:
        smtp.login(gmail_user, gmail_app_password)
        smtp.send_message(msg)

def main():
    cfg = load_config("config.yml")
    worker = (cfg["worker_base_url"] or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError(f"config.yml worker_base_url must start with https:// (got: {worker})")

    export_token = os.environ["EXPORT_TOKEN"]
    date = os.environ.get("RUN_DATE") or datetime.utcnow().strftime("%Y-%m-%d")

    prospects_url = f"{worker}/dump?date={date}&kind=prospects"
    prospects = fetch_jsonl(prospects_url, export_token)

    agency_records = {ag: [] for ag in cfg["agencies"].keys()}
    for p in prospects:
        ag = p.get("agency","")
        if ag in agency_records:
            agency_records[ag].append(p)

    # Emails agences (si adresses réelles, sinon skip)
    for ag, recs in agency_records.items():
        to_list = cfg["agencies"][ag]["daily_to"]
        if any("exemple.com" in x for x in to_list):
            print(f"[SKIP EMAIL] {ag} daily_to still example.com")
            continue

        excel_ag = build_excel(recs)
        send_email_gmail(
            f"[PROSPECTION] {ag} — Export import — {date}",
            f"Export prospection inconnus — agence {ag} — {date}\nFiches: {len(recs)}\n",
            to_list,
            [(f"{date}_{ag}_AGENCE_IMPORT.xlsx", excel_ag)]
        )

    # Email global (stats)
    lines = [f"Résumé global prospection — {date}", ""]
    total_fiches = 0
    for ag in cfg["agencies"].keys():
        n = len(agency_records[ag])
        total_fiches += n
        cmd = sum(1 for r in agency_records[ag] if (r.get("commande") or "").strip())
        lines.append(f"{ag}: {n} fiches, {cmd} commandes")
    lines.append("")
    lines.append(f"TOTAL GLOBAL: {total_fiches} fiches")

    to_global = cfg["global_to"]
    if any("exemple.com" in x for x in to_global):
        print("[SKIP EMAIL] global_to still example.com")
        return

    send_email_gmail(
        f"[PROSPECTION] GLOBAL — Résumé — {date}",
        "\n".join(lines),
        to_global,
        []
    )

if __name__ == "__main__":
    main()
