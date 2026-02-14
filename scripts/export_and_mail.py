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
# Colonnes Excel import
# =========================
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

# =========================
# Utils
# =========================
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

def pct_1d(num, den):
    if not den:
        return "0,0%"
    v = (num / den) * 100.0
    return f"{v:.1f}%".replace(".", ",")

def parse_ts_any(x):
    """
    Accepte:
    - unix ms (int)
    - unix sec (int)
    - iso string
    Retourne datetime UTC (naive->UTC)
    """
    if x is None:
        return None
    try:
        if isinstance(x, (int, float)):
            # heuristique ms vs sec
            if x > 10_000_000_000:
                return datetime.fromtimestamp(x / 1000, tz=timezone.utc)
            return datetime.fromtimestamp(x, tz=timezone.utc)
        if isinstance(x, str):
            s = x.strip()
            if not s:
                return None
            # support "2026-02-14T10:15:00Z" etc
            if s.endswith("Z"):
                s = s[:-1] + "+00:00"
            return datetime.fromisoformat(s).astimezone(timezone.utc)
    except Exception:
        return None
    return None

def am_pm_bucket(dt_utc):
    """
    On bucket en 'AM'/'PM' sur base UTC+1 (France hiver).
    Simple et stable pour ton usage (si besoin on rend configurable).
    """
    if not dt_utc:
        return "?"
    # France hiver = UTC+1
    hour_fr = (dt_utc.hour + 1) % 24
    return "AM" if hour_fr < 13 else "PM"

# =========================
# Fetch Worker
# =========================
def fetch_jsonl(url, export_token):
    print(f"[FETCH] {url}")

    req = urllib.request.Request(url)
    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)

    # Header d'auth export
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
        except Exception:
            pass
    return rows

# =========================
# Excel
# =========================
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

# =========================
# Brevo Send
# =========================
def send_email_brevo(subject, body, to_list, attachments):
    api_key = os.environ.get("BREVO_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("BREVO_API_KEY missing in GitHub Secrets")

    to_list = normalize_emails(to_list)
    if not to_list:
        print(f"[BREVO][SKIP] no recipients for subject: {subject}")
        return

    sender_email = (os.environ.get("BREVO_SENDER_EMAIL") or "").strip()
    if not sender_email:
        # fallback (mais évite en prod)
        sender_email = to_list[0]

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

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            _ = resp.read()
            print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        print("[BREVO][HTTPERROR]", e.code, err[:400].replace("\n", " "))
        raise

# =========================
# Stats
# =========================
def is_client_record(r):
    # Si ton worker met is_client true/false => parfait
    if isinstance(r.get("is_client"), bool):
        return r["is_client"]
    # fallback: si le champ "type" vaut "client"
    t = (r.get("type") or "").strip().lower()
    return t == "client"

def has_commande(r):
    return bool((r.get("commande") or "").strip())

def get_initials(r):
    # si ton worker le stocke => top
    ini = (r.get("user_initials") or "").strip().upper()
    if ini:
        return ini
    # fallback (pas idéal)
    return "??"

def build_agency_stats(records):
    prospects = len(records)
    clients = sum(1 for r in records if is_client_record(r))
    commandes = sum(1 for r in records if has_commande(r))
    tx = pct_1d(commandes, prospects)
    return prospects, clients, commandes, tx

def build_user_stats(records):
    """
    retourne dict:
      initials -> dict( prospect, client, cmd, tx, am_cmd/am_prospect, pm_cmd/pm_prospect )
    """
    out = {}
    for r in records:
        ini = get_initials(r)
        out.setdefault(ini, {
            "prospects": 0, "clients": 0, "commandes": 0,
            "am_p": 0, "am_c": 0, "pm_p": 0, "pm_c": 0
        })
        out[ini]["prospects"] += 1
        if is_client_record(r):
            out[ini]["clients"] += 1
        if has_commande(r):
            out[ini]["commandes"] += 1

        dt = parse_ts_any(r.get("ts") or r.get("created_at") or r.get("timestamp"))
        bucket = am_pm_bucket(dt)
        if bucket == "AM":
            out[ini]["am_p"] += 1
            if has_commande(r):
                out[ini]["am_c"] += 1
        elif bucket == "PM":
            out[ini]["pm_p"] += 1
            if has_commande(r):
                out[ini]["pm_c"] += 1
    return out

# =========================
# Main
# =========================
def main():
    cfg = load_config()
    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError(f"config.yml worker_base_url must start with https:// (got: {worker})")

    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    run_date = (os.environ.get("RUN_DATE") or "").strip()
    if run_date:
        date = run_date
    else:
        date = datetime.utcnow().strftime("%Y-%m-%d")

    # 1) Fetch daily records
    url = f"{worker}/dump?date={date}&kind=prospects"
    all_records = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(all_records))

    agencies_cfg = cfg.get("agencies", {})
    agency_records = {ag: [] for ag in agencies_cfg.keys()}

    for r in all_records:
        ag = (r.get("agency") or "").strip()
        if ag in agency_records:
            agency_records[ag].append(r)

    # 2) Emails agences (1 fichier import)
    for ag, recs in agency_records.items():
        to_list = agencies_cfg[ag].get("daily_to", [])
        excel = build_excel_import(recs)

        p, c, cmd, tx = build_agency_stats(recs)
        body = (
            f"Export prospection — agence {ag} — {date}\n"
            f"Prospects: {p}\n"
            f"Clients: {c}\n"
            f"Commandes: {cmd}\n"
            f"Tx: {tx}\n"
        )

        send_email_brevo(
            f"[PROSPECTION] {ag} — Export import — {date}",
            body,
            to_list,
            [(f"{date}_{ag}_IMPORT.xlsx", excel)]
        )

    # 3) Mail global (toi uniquement) + classement
    global_to = cfg.get("global_to", [])

    lines = [f"Résumé global prospection — {date}", ""]

    # Stats par agence
    for ag in agencies_cfg.keys():
        recs = agency_records[ag]
        p, c, cmd, tx = build_agency_stats(recs)
        lines.append(f"{ag}: Prospects {p} | Clients {c} | Commandes {cmd} | Tx {tx}")

        # Détail matin/aprem par user
        user_stats = build_user_stats(recs)
        if user_stats:
            # tri par commandes desc puis tx desc
            ranking = []
            for ini, d in user_stats.items():
                tx_user = (d["commandes"] / d["prospects"] * 100.0) if d["prospects"] else 0.0
                ranking.append((d["commandes"], tx_user, ini, d))
            ranking.sort(reverse=True)

            lines.append("  Classement commerciaux (agence) :")
            for i, (cmds, tx_user, ini, d) in enumerate(ranking, start=1):
                tx_str = f"{tx_user:.1f}%".replace(".", ",")
                tx_am = pct_1d(d["am_c"], d["am_p"])
                tx_pm = pct_1d(d["pm_c"], d["pm_p"])
                lines.append(
                    f"   {i}. {ini} — P:{d['prospects']} C:{d['clients']} Cmd:{d['commandes']} Tx:{tx_str} | AM {tx_am} | PM {tx_pm}"
                )
        else:
            lines.append("  (pas d'initiales commerciaux / pas de stats par utilisateur)")

        lines.append("")

    # Totaux globaux toutes agences
    all_p = len(all_records)
    all_c = sum(1 for r in all_records if is_client_record(r))
    all_cmd = sum(1 for r in all_records if has_commande(r))
    lines.append(f"TOTAL GLOBAL: Prospects {all_p} | Clients {all_c} | Commandes {all_cmd} | Tx {pct_1d(all_cmd, all_p)}")

    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date}",
        "\n".join(lines),
        global_to,
        []
    )

if __name__ == "__main__":
    main()
