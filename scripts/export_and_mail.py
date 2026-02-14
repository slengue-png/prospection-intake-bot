import os, io, json, base64, urllib.request, urllib.error
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

BREVO_API_URL = "https://api.brevo.com/v3/smtp/email"


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
    req = urllib.request.Request(url)
    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)
    # token via query already, but keep header (harmless)
    req.add_header("X-Export-Token", export_token)

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            raw = resp.read().decode("utf-8", errors="replace").strip()
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTPError {e.code} on {url} -> {body[:300]}")

    if not raw:
        return []
    rows = []
    for line in raw.splitlines():
        line = line.strip()
        if not line:
            continue
        rows.append(json.loads(line))
    return rows


def build_import_excel(records):
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
            r.get("contact_lastname","") or "",
            r.get("resume",""),
            r.get("commande",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def build_monthly_excel(month, prospect_rows, stats_rows):
    """
    Excel 'mensuel cumulé' (pour toi uniquement) :
    - onglet PROSPECTS (toutes entrées du mois)
    - onglet STATS (compteurs session)
    """
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "PROSPECTS"
    ws1.append(["date","agency","user_id","name","city","phone","email","resume","commande","photo_file_id"])

    for r in prospect_rows:
        ws1.append([
            r.get("date",""),
            r.get("agency",""),
            str(r.get("user_id","")),
            r.get("name",""),
            r.get("city",""),
            r.get("phone",""),
            r.get("email",""),
            r.get("resume",""),
            r.get("commande",""),
            r.get("card_photo_file_id",""),
        ])

    ws2 = wb.create_sheet("STATS")
    ws2.append(["date","slot","agency","user_id","prospects","clients","commandes"])

    for r in stats_rows:
        ws2.append([
            r.get("date",""),
            r.get("slot",""),
            r.get("agency",""),
            str(r.get("user_id","")),
            int(r.get("prospects",0)),
            int(r.get("clients",0)),
            int(r.get("commandes",0)),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def send_email_brevo(subject, body, to_list, attachments):
    api_key = os.environ.get("BREVO_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("BREVO_API_KEY missing")

    sender_email = os.environ.get("BREVO_SENDER_EMAIL", "").strip()
    if not sender_email:
        raise RuntimeError("BREVO_SENDER_EMAIL missing")

    to_list = normalize_emails(to_list)
    if not to_list:
        return

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
    req = urllib.request.Request(BREVO_API_URL, data=data, method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("api-key", api_key)

    with urllib.request.urlopen(req, timeout=45) as resp:
        _ = resp.read()
        # status 201 expected


def main():
    cfg = load_config()
    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing")

    date = os.environ.get("RUN_DATE") or datetime.utcnow().strftime("%Y-%m-%d")
    month = date[:7]
    token_q = export_token  # token in query param for worker

    # --- day dumps ---
    prospects_day = fetch_jsonl(f"{worker}/dump?token={token_q}&date={date}&kind=prospects", export_token)
    stats_day = fetch_jsonl(f"{worker}/dump?token={token_q}&date={date}&kind=stats", export_token)

    # --- month dumps (for cumulative) ---
    prospects_month = fetch_jsonl(f"{worker}/dump_month?token={token_q}&month={month}&kind=prospects", export_token)
    stats_month = fetch_jsonl(f"{worker}/dump_month?token={token_q}&month={month}&kind=stats", export_token)

    agencies_cfg = cfg.get("agencies", {})
    agency_prospects = {ag: [] for ag in agencies_cfg.keys()}
    for p in prospects_day:
        ag = p.get("agency","")
        if ag in agency_prospects:
            agency_prospects[ag].append(p)

    # --- EMAIL AGENCES : import excel (prospects only) ---
    for ag, recs in agency_prospects.items():
        to_list = agencies_cfg.get(ag, {}).get("daily_to", [])
        excel = build_import_excel(recs)
        send_email_brevo(
            subject=f"[PROSPECTION] {ag} — {date}",
            body=f"Export import agence {ag}\nFiches: {len(recs)}",
            to_list=to_list,
            attachments=[(f"{date}_{ag}.xlsx", excel)],
        )

    # --- PERFORMANCE (stats session) ---
    # Aggregations per agency & per user
    perf_ag = {ag: {"prospects": 0, "clients": 0, "commandes": 0} for ag in agencies_cfg.keys()}
    perf_user = {}  # (ag,user_id) -> counts

    for s in stats_day:
        ag = s.get("agency","")
        uid = str(s.get("user_id",""))
        if ag not in perf_ag:
            continue
        p = int(s.get("prospects", 0))
        c = int(s.get("clients", 0))
        k = int(s.get("commandes", 0))

        perf_ag[ag]["prospects"] += p
        perf_ag[ag]["clients"] += c
        perf_ag[ag]["commandes"] += k

        key = (ag, uid)
        if key not in perf_user:
            perf_user[key] = {"prospects": 0, "clients": 0, "commandes": 0}
        perf_user[key]["prospects"] += p
        perf_user[key]["clients"] += c
        perf_user[key]["commandes"] += k

    # Rankings (by commandes, then clients, then prospects)
    def rank_items_for_ag(ag):
        items = []
        for (a, uid), v in perf_user.items():
            if a != ag:
                continue
            items.append((uid, v["prospects"], v["clients"], v["commandes"]))
        items.sort(key=lambda x: (x[3], x[2], x[1]), reverse=True)
        return items

    # --- GLOBAL mail to you only (with monthly cumulative attachment) ---
    lines = []
    lines.append(f"RÉSUMÉ GLOBAL — {date}")
    lines.append("")
    total_p = total_c = total_k = 0

    for ag in agencies_cfg.keys():
        ap = perf_ag[ag]["prospects"]
        ac = perf_ag[ag]["clients"]
        ak = perf_ag[ag]["commandes"]
        total_p += ap
        total_c += ac
        total_k += ak

        lines.append(f"{ag} — Prospects: {ap} | Clients: {ac} | Commandes: {ak}")
        top = rank_items_for_ag(ag)[:5]
        if top:
            lines.append("  Classement (Top 5) :")
            for i, (uid, p, c, k) in enumerate(top, start=1):
                lines.append(f"   {i}. user_id {uid} — Cmd {k} | Clients {c} | Prospects {p}")
        lines.append("")

    lines.append(f"TOTAL — Prospects: {total_p} | Clients: {total_c} | Commandes: {total_k}")
    lines.append("")
    lines.append("Note: le classement est basé sur les COMPTEURS SESSION (bouton 🧾).")

    monthly_xlsx = build_monthly_excel(month, prospects_month, stats_month)

    send_email_brevo(
        subject=f"[PROSPECTION] GLOBAL — {date}",
        body="\n".join(lines),
        to_list=cfg.get("global_to", []),
        attachments=[(f"{month}_CUMUL_PROSPECTION.xlsx", monthly_xlsx)],
    )


if __name__ == "__main__":
    main()
