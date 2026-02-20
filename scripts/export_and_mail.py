import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime, date
from pathlib import Path
import re

import yaml
from openpyxl import Workbook

# =========================
# EXCEL EXPORT (FORMAT AGENCES)
# =========================
EXPORT_COLUMNS = [
    "NOM ENTREPRISE",
    "ADRESSE",
    "CODE POSTAL",
    "VILLE",
    "TELEPHONE",
    "MOBILE",
    "MAIL",
    "SIRET",
    "NAF",
    "SITE WEB",
    "DIRIGEANT 1 : civilité",
    "DIRIGEANT 1 : prénom",
    "DIRIGEANT 1 : nom",
    "ENTRETIEN",
    "COMMANDE",
    "NOMBRE DE VISITES CLIENTS",
    "NOMBRE DE VISITES PROSPECTS",
]

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0",
    "Accept": "*/*",
    "Accept-Language": "fr-FR,fr;q=0.9",
    "Connection": "close",
}

def load_config(path="config.yml"):
    with open(path, "r", encoding="utf-8") as f:
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
    return [e.strip() for e in raw if (e or "").strip()]

def fetch_jsonl(url, export_token):
    print(f"[FETCH] {url}")
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

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

def parse_run_date(s: str) -> date:
    s = (s or "").strip()
    if not s:
        return datetime.utcnow().date()

    # YYYY-MM-DD
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return datetime.strptime(s, "%Y-%m-%d").date()

    # DD/MM/YYYY
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", s):
        return datetime.strptime(s, "%d/%m/%Y").date()

    raise ValueError(f"RUN_DATE invalide: '{s}'. Formats acceptés: YYYY-MM-DD ou DD/MM/YYYY")

def split_dirigeant(full: str):
    full = (full or "").strip()
    if not full:
        return ("", "", "")
    parts = full.split()
    if len(parts) == 1:
        return ("", "", parts[0])
    return ("", parts[0], " ".join(parts[1:]))

def build_excel(records, visits_clients="", visits_prospects=""):
    wb = Workbook()
    ws = wb.active
    ws.title = "EXPORT"
    ws.append(EXPORT_COLUMNS)

    for r in records:
        civ, prenom, nom = split_dirigeant(r.get("dirigeant", ""))
        ws.append([
            r.get("name", ""),
            r.get("address", ""),
            r.get("postal_code", ""),
            r.get("city", ""),
            r.get("phone", ""),
            r.get("mobile", ""),
            r.get("email", ""),
            r.get("siret", ""),
            r.get("naf", ""),
            r.get("website", ""),
            civ,
            prenom,
            nom,
            r.get("resume", ""),
            r.get("commande", ""),
            "",  # visites clients -> pas par ligne
            "",  # visites prospects -> pas par ligne
        ])

    # on ajoute une dernière ligne "TOTAL VISITES" lisible
    ws.append([
        "TOTAL VISITES",
        "", "", "", "", "", "", "", "", "", "", "", "",
        "",
        "",
        str(visits_clients),
        str(visits_prospects),
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
        raise RuntimeError("BREVO_SENDER_EMAIL missing in GitHub Secrets")

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

    with urllib.request.urlopen(req, timeout=30) as resp:
        print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
        _ = resp.read()

def main():
    cfg = load_config()
    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError("config.yml: worker_base_url must start with https://")

    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    run_date_str = (os.environ.get("RUN_DATE") or "").strip()
    d = parse_run_date(run_date_str)
    date_str = d.strftime("%Y-%m-%d")

    agencies_cfg = cfg.get("agencies", {})
    global_to = cfg.get("global_to", [])

    # ===== dump prospects du jour =====
    prospects_url = f"{worker}/dump?date={date_str}&kind=prospects"
    prospects = fetch_jsonl(prospects_url, export_token)
    print("[OK] prospects rows=", len(prospects))

    # ===== dump closures du jour =====
    closures_url = f"{worker}/dump?date={date_str}&kind=closures"
    closures = fetch_jsonl(closures_url, export_token)
    print("[OK] closures rows=", len(closures))

    # closures index : (agency|initials) -> (clients, prospects)
    close_map = {}
    for c in closures:
        ag = (c.get("agency") or "").strip().upper()
        ini = (c.get("initials") or "").strip().upper()
        key = f"{ag}|{ini}"
        close_map[key] = {
            "clients": int(c.get("visits_clients") or 0),
            "prospects": int(c.get("visits_prospects") or 0),
        }

    # regroupe par agence
    agency_records = {ag: [] for ag in agencies_cfg.keys()}
    for p in prospects:
        ag = (p.get("agency") or "").strip().upper()
        if ag in agency_records:
            agency_records[ag].append(p)

    # ===== emails agences (excel + stats) =====
    for ag, recs in agency_records.items():
        to_list = agencies_cfg.get(ag, {}).get("daily_to", [])

        # totals visites = somme des clôtures des users de l'agence
        clients_total = 0
        prospects_total = 0
        for k, v in close_map.items():
            if k.startswith(ag + "|"):
                clients_total += v["clients"]
                prospects_total += v["prospects"]

        excel = build_excel(recs, visits_clients=clients_total, visits_prospects=prospects_total)

        send_email_brevo(
            f"[PROSPECTION] {ag} — {date_str}",
            f"Export agence {ag}\n"
            f"Prospects saisis: {len(recs)}\n"
            f"Visites clients (clôture): {clients_total}\n"
            f"Visites prospects (clôture): {prospects_total}\n",
            to_list,
            [(f"{date_str}_{ag}_EXPORT.xlsx", excel)],
        )

    # ===== email global (toi) =====
    lines = [f"Récap global prospection — {date_str}", ""]
    total_prospects = 0
    total_vc = 0
    total_vp = 0

    for ag in agency_records.keys():
        recs = agency_records[ag]
        total_prospects += len(recs)

        clients_total = 0
        prospects_total = 0
        for k, v in close_map.items():
            if k.startswith(ag + "|"):
                clients_total += v["clients"]
                prospects_total += v["prospects"]
        total_vc += clients_total
        total_vp += prospects_total

        lines.append(f"{ag} — Prospects saisis: {len(recs)} | Visites clients: {clients_total} | Visites prospects: {prospects_total}")

        # détail par commercial (clôtures)
        for k, v in sorted(close_map.items()):
            if k.startswith(ag + "|"):
                ini = k.split("|")[1]
                lines.append(f"  - {ini}: visites clients {v['clients']} | visites prospects {v['prospects']}")
        lines.append("")

    lines.append(f"TOTAL — Prospects saisis: {total_prospects} | Visites clients: {total_vc} | Visites prospects: {total_vp}")

    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date_str}",
        "\n".join(lines),
        global_to,
        [],
    )

if __name__ == "__main__":
    main()