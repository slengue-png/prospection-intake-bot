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

    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)

    # auth header
    req.add_header("X-Export-Token", export_token)

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8", errors="replace").strip()
            print(f"[HTTP] {resp.status}")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"[HTTPERROR] code={e.code}")
        print("[HTTPERROR_BODY_300]", body[:300].replace("\n", " ") )
        raise RuntimeError(f"HTTPError {e.code} on {url}")
    except urllib.error.URLError as e:
        print("[URLERROR]", getattr(e, "reason", str(e)))
        raise RuntimeError(f"URLError on {url} -> {getattr(e,'reason',str(e))}")
    except Exception as e:
        print("[FETCH_ERROR]", str(e))
        raise

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
            # si la ligne n'est pas du JSON, on ignore
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
            "", "",
            r.get("city",""),
            r.get("phone",""),
            "",
            r.get("email",""),
            "", "", "",
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
    print(f"[OK] prospects rows={len(prospects)}")

    # Pas d'envoi mail tant que exemple.com (comme tu voulais)
    # Le but ici = valider l'export sans 1010.

if __name__ == "__main__":
    main()
