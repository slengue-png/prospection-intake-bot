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
        if e:
            out.append(e)
    return out

def fetch_jsonl(url, export_token):
    print(f"[FETCH] {url}")
    req = urllib.request.Request(url)
    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)
    req.add_header("X-Export-Token", export_token)

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            raw = resp.read().decode("utf-8", errors="replace").strip()
            print("[HTTP]", resp.status)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTPError {e.code} on {url} -> {body[:300]}")
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

def build_excel_import(records):
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
    api_key = os.environ.get("BREVO_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("BREVO_API_KEY missing in GitHub Secrets")

    to_list = normalize_emails(to_list)
    if not to_list:
        print(f"[BREVO][SKIP] no recipients for subject: {subject}")
        return

    sender_email = os.environ.get("BREVO_SENDER_EMAIL", "").strip()
    if not sender_email:
        raise RuntimeError("BREVO_SENDER_EMAIL missing in GitHub Secrets (must be verified in Brevo)")

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

    with urllib.request.urlopen(req, timeout=45) as resp:
        print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
        _ = resp.read()

def pct_1dec(num, den):
    if den <= 0:
        return "0,0%"
    return f"{(num/den)*100:.1f}%".replace(".", ",")

def main():
    cfg = load_config()

    worker = (cfg.get("worker_base_url","") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError("config.yml: worker_base_url must start with https://")

    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    date = (os.environ.get("RUN_DATE") or "").strip()
    if not date:
        date = datetime.utcnow().strftime("%Y-%m-%d")

    # ---- fetch prospects + stats
    prospects = fetch_jsonl(f"{worker}/dump?date={date}&kind=prospects", export_token)
    stats = fetch_jsonl(f"{worker}/dump?date={date}&kind=stats", export_token)
    print("[OK] prospects rows=", len(prospects))
    print("[OK] stats rows=", len(stats))

    agencies = cfg.get("agencies", {})
    agency_records = {ag: [] for ag in agencies.keys()}
    for p in prospects:
        ag = (p.get("agency","") or "").upper()
        if ag in agency_records:
            agency_records[ag].append(p)

    # stats by agency + by initials
    # We aggregate session_totals only
    stats_ag = {ag: {"prospects":0, "clients":0, "commandes":0, "by_ini":{}} for ag in agencies.keys()}
    for s in stats:
        if s.get("type") != "session_totals":
            continue
        ag = (s.get("agency","") or "").upper()
        ini = (s.get("initials","") or "").upper() or "??"
        if ag not in stats_ag:
            continue
        stats_ag[ag]["prospects"] += int(s.get("total_prospects",0) or 0)
        stats_ag[ag]["clients"] += int(s.get("total_clients",0) or 0)
        stats_ag[ag]["commandes"] += int(s.get("total_commandes",0) or 0)
        if ini not in stats_ag[ag]["by_ini"]:
            stats_ag[ag]["by_ini"][ini] = {"prospects":0, "clients":0, "commandes":0}
        stats_ag[ag]["by_ini"][ini]["prospects"] += int(s.get("total_prospects",0) or 0)
        stats_ag[ag]["by_ini"][ini]["clients"] += int(s.get("total_clients",0) or 0)
        stats_ag[ag]["by_ini"][ini]["commandes"] += int(s.get("total_commandes",0) or 0)

    # ---- Emails agences (Excel import)
    for ag, recs in agency_records.items():
        to_list = agencies[ag].get("daily_to", [])
        excel = build_excel_import(recs)

        # corps simple
        cmd_count = sum(1 for r in recs if (r.get("commande") or "").strip())
        body = (
            f"Export prospection inconnus — agence {ag} — {date}\n"
            f"Fiches: {len(recs)}\n"
            f"Commandes (dans fiches): {cmd_count}\n"
        )

        send_email_brevo(
            f"[PROSPECTION] {ag} — Export import — {date}",
            body,
            to_list,
            [(f"{date}_{ag}_IMPORT.xlsx", excel)]
        )

    # ---- Email global (toi uniquement) : prospects + clients + commandes + taux transfo par commercial
    global_to = cfg.get("global_to", [])
    lines = [f"Résumé global prospection — {date}", ""]

    for ag in agencies.keys():
        agP = stats_ag[ag]["prospects"]
        agC = stats_ag[ag]["clients"]
        agK = stats_ag[ag]["commandes"]
        lines.append(f"{ag}: Prospects {agP} | Clients {agC} | Commandes {agK} | Tx {pct_1dec(agK, agP)}")

        by_ini = stats_ag[ag]["by_ini"]
        if by_ini:
            for ini in sorted(by_ini.keys()):
                p = by_ini[ini]["prospects"]
                c = by_ini[ini]["clients"]
                k = by_ini[ini]["commandes"]
                lines.append(f"  - {ini}: Prospects {p} | Clients {c} | Commandes {k} | Tx {pct_1dec(k,p)}")
        else:
            lines.append("  - (pas de totaux session saisis)")

        lines.append("")

    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date}",
        "\n".join(lines),
        global_to,
        []
    )

if __name__ == "__main__":
    main()
