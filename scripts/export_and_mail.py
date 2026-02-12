import os
import io
import json
import base64
import urllib.request
from datetime import datetime
from openpyxl import Workbook
import yaml

IMPORT_COLUMNS = [
    "NOM","ADRESSE","CODE POSTAL","VILLE","TELEPHONE","TELEPHONE 2","MAIL",
    "SIRET","NAF","SITE WEB","Contact: civilité","Contact : prénom","Contact : nom",
    "RESUME ENTRETIEN","COMMANDE"
]

def load_config():
    with open("config.yml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def fetch_jsonl(url, token):
    req = urllib.request.Request(url)
    req.add_header("X-Export-Token", token)
    with urllib.request.urlopen(req, timeout=30) as resp:
        raw = resp.read().decode("utf-8").strip()
    if not raw:
        return []
    return [json.loads(x) for x in raw.splitlines() if x.strip()]

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
            r.get("contact_lastname","") or r.get("interlocuteur",""),
            r.get("resume",""),
            r.get("commande",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

def send_email_brevo(subject, body, to_list, attachments):
    api_key = os.environ["BREVO_API_KEY"]

    payload = {
        "sender": {
            "email": to_list[0],
            "name": "Prospection Bot"
        },
        "to": [{"email": x} for x in to_list],
        "subject": subject,
        "textContent": body,
        "attachment": []
    }

    for filename, content_bytes in attachments:
        b64_content = base64.b64encode(content_bytes).decode("utf-8")
        payload["attachment"].append({
            "name": filename,
            "content": b64_content
        })

    data = json.dumps(payload).encode("utf-8")

    req = urllib.request.Request(
        "https://api.brevo.com/v3/smtp/email",
        data=data,
        method="POST"
    )
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", api_key)

    with urllib.request.urlopen(req) as resp:
        print("[BREVO] status:", resp.status)

def main():
    cfg = load_config()
    worker = cfg["worker_base_url"].rstrip("/")
    token = os.environ["EXPORT_TOKEN"]
    date = datetime.utcnow().strftime("%Y-%m-%d")

    url = f"{worker}/dump?date={date}&kind=prospects"
    prospects = fetch_jsonl(url, token)
    print("Prospects:", len(prospects))

    agency_records = {ag: [] for ag in cfg["agencies"].keys()}
    for p in prospects:
        ag = p.get("agency","")
        if ag in agency_records:
            agency_records[ag].append(p)

    # Mails agences
    for ag, recs in agency_records.items():
        excel = build_excel(recs)
        send_email_brevo(
            f"[PROSPECTION] {ag} — {date}",
            f"Export agence {ag}\nFiches: {len(recs)}",
            cfg["agencies"][ag]["daily_to"],
            [(f"{date}_{ag}.xlsx", excel)]
        )

    # Mail global
    lines = [f"Résumé global {date}",""]
    total = 0
    for ag in agency_records:
        n = len(agency_records[ag])
        total += n
        cmd = sum(1 for r in agency_records[ag] if r.get("commande"))
        lines.append(f"{ag}: {n} fiches / {cmd} commandes")

    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date}",
        "\n".join(lines),
        cfg["global_to"],
        []
    )

if __name__ == "__main__":
    main()
