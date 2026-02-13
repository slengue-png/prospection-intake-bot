import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime
from openpyxl import Workbook
import yaml

IMPORT_COLUMNS = [
    "NOM","ADRESSE","CODE POSTAL","VILLE","TELEPHONE","TELEPHONE 2","MAIL",
    "SIRET","NAF","SITE WEB","Contact: civilité","Contact : prénom","Contact : nom",
    "RESUME ENTRETIEN","COMMANDE"
]

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "*/*",
    "Connection": "close",
}

def load_config():
    with open("config.yml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def normalize_emails(x):
    if x is None:
        return []
    if isinstance(x, list):
        raw = x
    elif isinstance(x, str):
        raw = [p.strip() for p in x.split(",")]
    else:
        raw = []
    return [e.strip() for e in raw if e.strip()]

def fetch_jsonl(url, export_token):
    print(f"[FETCH] {url}")

    req = urllib.request.Request(url)
    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)
    req.add_header("X-Export-Token", export_token)

    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read().decode("utf-8", errors="replace").strip()
        print("[HTTP]", resp.status)

    if not raw:
        return []
    return [json.loads(line) for line in raw.splitlines() if line.strip()]

def build_excel(records):
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
            r.get("contact_lastname",""),
            r.get("resume",""),
            r.get("commande",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

def send_email_brevo(subject, body, to_list, attachments):
    api_key = os.environ.get("BREVO_API_KEY", "")
    if not api_key:
        raise RuntimeError("BREVO_API_KEY missing")

    to_list = normalize_emails(to_list)
    if not to_list:
        return

    sender_email = os.environ.get("BREVO_SENDER_EMAIL")

    payload = {
        "sender": {"email": sender_email, "name": "Prospection Bot"},
        "to": [{"email": x} for x in to_list],
        "subject": subject,
        "textContent": body,
    }

    if attachments:
        payload["attachment"] = []
        for filename, content_bytes in attachments:
            payload["attachment"].append({
                "name": filename,
                "content": base64.b64encode(content_bytes).decode("utf-8")
            })

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://api.brevo.com/v3/smtp/email",
        data=data,
        method="POST"
    )
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", api_key)

    with urllib.request.urlopen(req, timeout=30) as resp:
        print("[BREVO] status:", resp.status)

def main():
    cfg = load_config()
    worker = cfg["worker_base_url"].rstrip("/")
    export_token = os.environ.get("EXPORT_TOKEN", "")
    date = os.environ.get("RUN_DATE") or datetime.utcnow().strftime("%Y-%m-%d")

    url = f"{worker}/dump?date={date}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(prospects))

    agency_records = {ag: [] for ag in cfg.get("agencies", {}).keys()}
    for p in prospects:
        ag = p.get("agency","")
        if ag in agency_records:
            agency_records[ag].append(p)

    # 📩 Emails agences (avec Excel)
    for ag, recs in agency_records.items():
        to_list = cfg["agencies"][ag].get("daily_to", [])
        excel = build_excel(recs)
        send_email_brevo(
            f"[PROSPECTION] {ag} — {date}",
            f"Export agence {ag}\nFiches: {len(recs)}",
            to_list,
            [(f"{date}_{ag}.xlsx", excel)]
        )

    # 📊 Email GLOBAL détaillé
    lines = []
    lines.append(f"RÉSUMÉ GLOBAL — {date}")
    lines.append("")

    total_prospects = 0
    total_clients = 0
    total_commandes = 0

    for ag, recs in agency_records.items():
        prospects_count = len(recs)
        clients_count = sum(1 for r in recs if r.get("siret","").strip())
        commandes_count = sum(1 for r in recs if (r.get("commande") or "").strip())

        total_prospects += prospects_count
        total_clients += clients_count
        total_commandes += commandes_count

        lines.append(f"{ag}")
        lines.append(f"  Prospects : {prospects_count}")
        lines.append(f"  Clients   : {clients_count}")
        lines.append(f"  Commandes : {commandes_count}")
        lines.append("")

    lines.append("TOTAL GLOBAL")
    lines.append(f"  Prospects : {total_prospects}")
    lines.append(f"  Clients   : {total_clients}")
    lines.append(f"  Commandes : {total_commandes}")

    global_to = cfg.get("global_to", [])
    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date}",
        "\n".join(lines),
        global_to,
        []
    )

if __name__ == "__main__":
    main()
