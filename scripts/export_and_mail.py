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
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "fr-FR,fr;q=0.9",
    "Connection": "close",
}

def load_config():
    with open("config.yml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def normalize_emails(x):
    """
    Accepte:
      - liste ["a@b.com", ...]
      - string "a@b.com"
      - string "a@b.com,b@c.com"
    Retourne une liste nettoyée.
    """
    if x is None:
        return []
    if isinstance(x, list):
        raw = x
    elif isinstance(x, str):
        raw = [p.strip() for p in x.split(",")]
    else:
        raw = []
    out = []
    for e in raw:
        e = (e or "").strip()
        if not e:
            continue
        out.append(e)
    return out

def fetch_jsonl(url, export_token):
    print(f"[FETCH] {url}")
    print(f"[DEBUG] EXPORT_TOKEN present? {'YES' if export_token else 'NO'}")

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
            r.get("contact_lastname","") or r.get("interlocuteur",""),
            r.get("resume",""),
            r.get("commande",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

def send_email_brevo(subject, body, to_list, attachments):
    api_key = os.environ.get("BREVO_API_KEY", "")
    if not api_key:
        raise RuntimeError("BREVO_API_KEY missing in GitHub Secrets")

    to_list = normalize_emails(to_list)
    if not to_list:
        print(f"[BREVO][SKIP] no recipients for subject: {subject}")
        return

    # sender: pour Brevo, mieux = une adresse vérifiée dans Brevo
    sender_email = os.environ.get("BREVO_SENDER_EMAIL") or to_list[0]

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
    req = urllib.request.Request("https://api.brevo.com/v3/smtp/email", data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", api_key)

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
            _ = resp.read()
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        print("[BREVO][HTTPERROR]", e.code, err[:400].replace("\n", " "))
        raise

def main():
    cfg = load_config()
    worker = cfg["worker_base_url"].rstrip("/")
    export_token = os.environ.get("EXPORT_TOKEN", "")
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    date = os.environ.get("RUN_DATE") or datetime.utcnow().strftime("%Y-%m-%d")

    url = f"{worker}/dump?date={date}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(prospects))

    agency_records = {ag: [] for ag in cfg.get("agencies", {}).keys()}
    for p in prospects:
        ag = p.get("agency","")
        if ag in agency_records:
            agency_records[ag].append(p)

    # Emails agences
    for ag, recs in agency_records.items():
        to_list = cfg["agencies"][ag].get("daily_to", [])
        excel = build_excel(recs)
        send_email_brevo(
            f"[PROSPECTION] {ag} — {date}",
            f"Export agence {ag}\nFiches: {len(recs)}",
            to_list,
            [(f"{date}_{ag}.xlsx", excel)]
        )

    # Email global (toi uniquement)
    lines = [f"Résumé global {date}", ""]
    total = 0
    for ag in agency_records:
        n = len(agency_records[ag])
        total += n
        cmd = sum(1 for r in agency_records[ag] if (r.get("commande") or "").strip())
        lines.append(f"{ag}: {n} fiches / {cmd} commandes")
    lines.append("")
    lines.append(f"TOTAL: {total} fiches")

    global_to = cfg.get("global_to", [])
    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date}",
        "\n".join(lines),
        global_to,
        []
    )

if __name__ == "__main__":
    main()
