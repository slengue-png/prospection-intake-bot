import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime, timezone
from openpyxl import Workbook
import yaml

# =========================
# COLONNES IMPORT (logiciel)
# + AJOUT : DIRIGEANT
# =========================
IMPORT_COLUMNS = [
    "NOM",
    "ADRESSE",
    "CODE POSTAL",
    "VILLE",
    "TELEPHONE",
    "TELEPHONE 2",
    "MAIL",
    "SIRET",
    "NAF",
    "SITE WEB",
    "Contact: civilité",
    "Contact : prénom",
    "Contact : nom",
    "DIRIGEANT",
    "RESUME ENTRETIEN",
    "COMMANDE",
]

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "fr-FR,fr;q=0.9",
    "Connection": "close",
}

def load_config(path="config.yml"):
    with open(path, "r", encoding="utf-8") as f:
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
        if e:
            out.append(e)
    return out

def pct_1d(num, den):
    if den <= 0:
        return "0,0%"
    val = (num / den) * 100.0
    # 1 décimale, format FR
    return f"{val:.1f}%".replace(".", ",")

def fetch_jsonl(url, export_token):
    print(f"[FETCH] {url}")
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    req = urllib.request.Request(url)
    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)

    # Auth Worker /dump
    req.add_header("X-Export-Token", export_token)

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            raw = resp.read().decode("utf-8", errors="replace").strip()
            print("[HTTP]", resp.status)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTPError {e.code} on {url} -> {body[:300]}")
    except urllib.error.URLError as e:
        raise RuntimeError(f"URLError on {url} -> {e.reason}")

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

def build_excel(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "IMPORT"
    ws.append(IMPORT_COLUMNS)

    for r in records:
        ws.append([
            r.get("name", ""),
            r.get("address", ""),
            r.get("postal_code", ""),
            r.get("city", ""),
            r.get("phone", ""),
            r.get("phone2", ""),
            r.get("email", ""),
            r.get("siret", ""),
            r.get("naf", ""),
            r.get("website", ""),
            r.get("contact_civility", ""),
            r.get("contact_firstname", ""),
            r.get("contact_lastname", "") or r.get("interlocuteur", ""),
            r.get("dirigeant", ""),
            r.get("resume", ""),
            r.get("commande", ""),
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

    sender_email = os.environ.get("BREVO_SENDER_EMAIL", "").strip()
    if not sender_email:
        raise RuntimeError("BREVO_SENDER_EMAIL missing in GitHub Secrets (ex: prospection@pilopro.fr)")

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
                "content": base64.b64encode(content_bytes).decode("utf-8"),
            })

    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request("https://api.brevo.com/v3/smtp/email", data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", api_key)

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
            _ = resp.read()
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"BREVO HTTPError {e.code}: {err[:400].replace(chr(10),' ')}")

def main():
    cfg = load_config("config.yml")
    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError(f"config.yml worker_base_url must start with https:// (got: {worker})")

    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    # Date de run : si workflow_dispatch avec input, sinon date UTC du jour
    run_date = os.environ.get("RUN_DATE", "").strip()
    if run_date:
        date = run_date
    else:
        date = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    # 1) récup prospects
    url = f"{worker}/dump?date={date}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(prospects))

    agencies_cfg = cfg.get("agencies", {})
    agency_codes = list(agencies_cfg.keys())

    # group by agency
    agency_records = {ag: [] for ag in agency_codes}
    for p in prospects:
        ag = (p.get("agency") or "").strip().upper()
        if ag in agency_records:
            agency_records[ag].append(p)

    # =========================
    # EMAIL AGENCES (Excel import)
    # =========================
    for ag, recs in agency_records.items():
        to_list = agencies_cfg.get(ag, {}).get("daily_to", [])
        excel = build_excel(recs)
        send_email_brevo(
            subject=f"[PROSPECTION] {ag} — Export import — {date}",
            body=f"Export prospection inconnus — agence {ag}\nDate: {date}\nProspects: {len(recs)}\n",
            to_list=to_list,
            attachments=[(f"{date}_{ag}_IMPORT.xlsx", excel)],
        )

    # =========================
    # EMAIL GLOBAL (toi seul)
    # Prospects / Clients / Commandes + Tx transfo
    # + détail par commercial (initials)
    # =========================
    lines = [f"RÉSUMÉ GLOBAL PROSPECTION — {date}", ""]

    total_prospects = 0
    total_clients = 0  # pas encore saisi -> 0 pour l’instant
    total_commandes = 0

    for ag in agency_codes:
        recs = agency_records[ag]
        prospects_n = len(recs)
        clients_n = 0
        commandes_n = sum(1 for r in recs if (r.get("commande") or "").strip())

        total_prospects += prospects_n
        total_clients += clients_n
        total_commandes += commandes_n

        tx_ag = pct_1d(commandes_n, prospects_n)

        lines.append(f"{ag} — Prospects: {prospects_n} | Clients: {clients_n} | Commandes: {commandes_n} | Tx: {tx_ag}")

        # détail par commercial (initials)
        by_init = {}
        for r in recs:
            ini = (r.get("initials") or "NA").strip().upper()
            by_init.setdefault(ini, []).append(r)

        if by_init:
            for ini in sorted(by_init.keys()):
                rr = by_init[ini]
                p_n = len(rr)
                c_n = 0
                cmd_n = sum(1 for x in rr if (x.get("commande") or "").strip())
                tx = pct_1d(cmd_n, p_n)
                lines.append(f"  - {ini}: Prospects {p_n} | Clients {c_n} | Commandes {cmd_n} | Tx {tx}")
        lines.append("")

    tx_global = pct_1d(total_commandes, total_prospects)
    lines.append(f"TOTAL — Prospects: {total_prospects} | Clients: {total_clients} | Commandes: {total_commandes} | Tx: {tx_global}")

    global_to = cfg.get("global_to", [])
    send_email_brevo(
        subject=f"[PROSPECTION] GLOBAL — Résumé — {date}",
        body="\n".join(lines),
        to_list=global_to,
        attachments=[],
    )

if __name__ == "__main__":
    main()
