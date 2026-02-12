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

def load_config(path="config.yml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def fetch_jsonl(url: str):
    print(f"[FETCH] {url}")
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
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
        try:
            rows.append(json.loads(line))
        except:
            pass
    return rows

def build_excel(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "IMPORT"
    ws.append(IMPORT_COLUMNS)

    for r in records:
        ws.append([
            r.get("name",""),
            "",
            "",
            r.get("city",""),
            r.get("phone",""),
            "",
            r.get("email",""),
            "",
            "",
            "",
            "",
            "",
            r.get("interlocuteur",""),
            r.get("resume",""),
            r.get("commande",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

def send_email(subject, body, to_list, attachments):
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

    token = os.environ["EXPORT_TOKEN"]
    date = datetime.utcnow().strftime("%Y-%m-%d")

    prospects_url = f"{worker}/dump?token={token}&date={date}&kind=prospects"
    prospects = fetch_jsonl(prospects_url)

    agency_records = {ag: [] for ag in cfg["agencies"].keys()}
    for p in prospects:
        ag = p.get("agency","")
        if ag in agency_records:
            agency_records[ag].append(p)

    for ag, recs in agency_records.items():
        excel_file = build_excel(recs)
        to_list = cfg["agencies"][ag]["daily_to"]

        if any("exemple.com" in x for x in to_list):
            print(f"[SKIP EMAIL] {ag} daily_to still example.com")
            continue

        send_email(
            f"[PROSPECTION] {ag} — {date}",
            f"Fiches: {len(recs)}",
            to_list,
            [(f"{date}_{ag}.xlsx", excel_file)]
        )

    total = sum(len(v) for v in agency_records.values())
    body_global = f"Résumé global {date}\n\nTOTAL: {total} fiches"

    if any("exemple.com" in x for x in cfg["global_to"]):
        print("[SKIP EMAIL] global_to still example.com")
        return

    send_email(
        f"[PROSPECTION] GLOBAL — {date}",
        body_global,
        cfg["global_to"],
        []
    )

if __name__ == "__main__":
    main()
