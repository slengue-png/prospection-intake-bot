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
    req.add_header("X-Export-Token", export_token)

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            raw = resp.read().decode("utf-8", errors="replace").strip()
            print(f"[HTTP] {resp.status}")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTPError {e.code} on {url} -> {body[:300]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"URLError on {url} -> {getattr(e,'reason',str(e))}")

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
            r.get("address",""),
            r.get("postal_code",""),
            r.get("city",""),
            r.get("phone",""),
            r.get("phone2",""),
            r.get("email",""),
            r.get("siret",""),
            r.get("naf",""),
            r.get("website",""),
            r.get("contact_civility",""),
            r.get("contact_firstname",""),
            r.get("contact_lastname","") or r.get("interlocuteur",""),
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

    # Regroupement par agence
    agency_records = {ag: [] for ag in cfg["agencies"].keys()}
    for p in prospects:
        ag = (p.get("agency") or "").strip()
        if ag in agency_records:
            agency_records[ag].append(p)

    # 1) Envoi Excel "import logiciel" par agence
    for ag, recs in agency_records.items():
        to_list = cfg["agencies"][ag]["daily_to"]
        if not to_list or any("exemple.com" in x for x in to_list):
            print(f"[SKIP EMAIL] {ag} daily_to not set or still example.com")
            continue

        excel_ag = build_excel(recs)
        subject = f"[PROSPECTION] {ag} — Import — {date}"
        body = f"Export prospection inconnus — {ag} — {date}\nFiches: {len(recs)}"
        send_email_gmail(subject, body, to_list, [(f"{date}_{ag}_IMPORT.xlsx", excel_ag)])
        print(f"[SENT] agency {ag} -> {len(to_list)} recipients")

    # 2) Mail global stats (uniquement toi)
    lines = [f"Résumé global prospection — {date}", ""]
    total_fiches = 0
    for ag in cfg["agencies"].keys():
        n = len(agency_records[ag])
        total_fiches += n
        cmd = sum(1 for r in agency_records[ag] if (r.get("commande") or "").strip())
        # (par personne viendra à l'étape suivante avec /whoami)
        lines.append(f"{ag}: {n} fiches, {cmd} commandes")
    lines.append("")
    lines.append(f"TOTAL GLOBAL: {total_fiches} fiches")

    to_global = cfg.get("global_to", [])
    if to_global and not any("exemple.com" in x for x in to_global):
        send_email_gmail(
            f"[PROSPECTION] GLOBAL — Résumé — {date}",
            "\n".join(lines),
            to_global,
            []
        )
        print("[SENT] global summary")
    else:
        print("[SKIP EMAIL] global_to not set or still example.com")

if __name__ == "__main__":
    main()
