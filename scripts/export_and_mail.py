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
    "SIRET","NAF","SITE WEB","DIRIGEANT",
    "Contact: civilité","Contact : prénom","Contact : nom",
    "RESUME ENTRETIEN","COMMANDE"
]

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "fr-FR,fr;q=0.9",
    "Connection": "close",
}

PARASITE_WORDS = {
    "urgent","besoin","profil","poste","postes","recherche","cherche","pour","de","des","du","d","et","avec","a","à",
    "svp","stp","merci","asap"
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
    req = urllib.request.Request(url)
    for k, v in BROWSER_HEADERS.items():
        req.add_header(k, v)
    req.add_header("X-Export-Token", export_token)

    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            raw = resp.read().decode("utf-8", errors="replace").strip()
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
            r.get("dirigeant",""),
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

    sender_email = os.environ.get("BREVO_SENDER_EMAIL", "").strip()
    if not sender_email:
        sender_email = to_list[0]  # fallback

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
        _ = resp.read()
        print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)

def parse_commandes_count(raw: str) -> int:
    """
    Règles demandées :
      - "1 manutentionnaire" -> 1
      - "1" -> 1
      - "1 manutentionnaire 1 chauffeur" -> 2
      - "manutentionnaire chauffeur" -> 2 (métiers distincts)
      - "2 manutentionnaires" -> 1
    Logique :
      - On split en tokens
      - On retire parasites
      - On compte les "métiers distincts" détectés
      - Si on ne détecte aucun mot-métier mais il y a un chiffre => 1
      - Si texte non vide mais aucun chiffre => 1
    """
    if not raw:
        return 0
    s = str(raw).strip().lower()
    if not s:
        return 0

    tokens = [t for t in s.replace(",", " ").replace(";", " ").split() if t]
    if not tokens:
        return 0

    nums = [t for t in tokens if t.isdigit()]
    words = [t for t in tokens if not t.isdigit()]

    # enlève parasites
    words = [w for w in words if w not in PARASITE_WORDS]

    # normalise pluriels simples
    norm = []
    for w in words:
        if w.endswith("s") and len(w) > 3:
            norm.append(w[:-1])
        else:
            norm.append(w)

    distinct_jobs = []
    for w in norm:
        # ignore mots trop courts
        if len(w) < 3:
            continue
        distinct_jobs.append(w)

    # dédoublonne
    distinct_jobs = list(dict.fromkeys(distinct_jobs))

    if distinct_jobs:
        # "manutentionnaire chauffeur" => 2
        return len(distinct_jobs)

    # si que des chiffres -> 1
    if nums:
        return 1

    # sinon texte non vide -> 1
    return 1

def month_key(date_yyyy_mm_dd: str) -> str:
    return date_yyyy_mm_dd[:7]  # YYYY-MM

def load_jsonl(path):
    if not os.path.exists(path):
        return []
    out = []
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except:
                pass
    return out

def save_jsonl(path, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")

def build_excel_global_stats(stats_rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "STATS"
    ws.append(["DATE","AGENCE","INITIALS","PERIOD","PROSPECTS","CLIENTS"])

    for r in stats_rows:
        ws.append([
            r.get("date",""),
            r.get("agency",""),
            r.get("initials",""),
            r.get("period",""),
            r.get("prospects",""),
            r.get("clients",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

def pct_1dec(num, den):
    if den <= 0:
        return "0,0%"
    v = (num / den) * 100.0
    return f"{v:.1f}%".replace(".", ",")

def main():
    cfg = load_config()
    worker = cfg["worker_base_url"].rstrip("/")
    export_token = os.environ.get("EXPORT_TOKEN", "")
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    date = os.environ.get("RUN_DATE") or datetime.utcnow().strftime("%Y-%m-%d")

    # ---- fetch prospects + stats
    prospects = fetch_jsonl(f"{worker}/dump?date={date}&kind=prospects", export_token)
    stats_day = fetch_jsonl(f"{worker}/dump?date={date}&kind=stats", export_token)

    # ---- group prospects by agency
    agency_records = {ag: [] for ag in cfg.get("agencies", {}).keys()}
    for p in prospects:
        ag = (p.get("agency","") or "").strip()
        if ag in agency_records:
            agency_records[ag].append(p)

    # ---- send per agency import excel
    for ag, recs in agency_records.items():
        to_list = cfg["agencies"][ag].get("daily_to", [])
        excel = build_excel_import(recs)
        send_email_brevo(
            f"[PROSPECTION] {ag} — IMPORT — {date}",
            f"Export agence {ag}\nFiches: {len(recs)}",
            to_list,
            [(f"{date}_{ag}_IMPORT.xlsx", excel)]
        )

    # ---- compute commandes count per person & agency from prospects
    # map: (agency, initials) -> {fiches, commandes}
    perf = {}
    for p in prospects:
        ag = (p.get("agency","") or "").strip()
        ini = (p.get("initials","") or "").strip()
        if not ag or not ini:
            continue
        key = (ag, ini)
        if key not in perf:
            perf[key] = {"fiches": 0, "commandes": 0}
        perf[key]["fiches"] += 1
        perf[key]["commandes"] += parse_commandes_count(p.get("commande",""))

    # ---- stats prospects/clients = depuis /stats (saisi Telegram)
    # map: (agency, initials) -> {prospects, clients, periods...}
    pc = {}
    for s in stats_day:
        ag = (s.get("agency","") or "").strip()
        ini = (s.get("initials","") or "").strip()
        if not ag or not ini:
            continue
        key = (ag, ini)
        if key not in pc:
            pc[key] = {"prospects": 0, "clients": 0, "am": {"p":0,"c":0}, "pm": {"p":0,"c":0}, "day": {"p":0,"c":0}}
        def to_int(x):
            try:
                return int(str(x).strip())
            except:
                return 0
        pnum = to_int(s.get("prospects",""))
        cnum = to_int(s.get("clients",""))
        pc[key]["prospects"] += pnum
        pc[key]["clients"] += cnum
        period = (s.get("period","") or "").upper()
        if period == "AM":
            pc[key]["am"]["p"] += pnum
            pc[key]["am"]["c"] += cnum
        elif period == "PM":
            pc[key]["pm"]["p"] += pnum
            pc[key]["pm"]["c"] += cnum
        else:
            pc[key]["day"]["p"] += pnum
            pc[key]["day"]["c"] += cnum

    # ---- Build global mail body (toi uniquement)
    lines = [f"Résumé global — {date}", ""]
    total_fiches = 0
    total_cmd = 0

    agencies = list(cfg.get("agencies", {}).keys())
    agencies.sort()

    for ag in agencies:
        lines.append(f"=== {ag} ===")
        # personnes de l’agence
        persons = sorted([ini for (a, ini) in perf.keys() if a == ag])
        # include those with stats but no fiches
        persons2 = sorted([ini for (a, ini) in pc.keys() if a == ag and ini not in persons])
        persons += persons2

        ag_f = 0
        ag_cmd = 0
        ag_p = 0
        ag_c = 0

        for ini in persons:
            f = perf.get((ag, ini), {}).get("fiches", 0)
            cmd = perf.get((ag, ini), {}).get("commandes", 0)
            pnum = pc.get((ag, ini), {}).get("prospects", 0)
            cnum = pc.get((ag, ini), {}).get("clients", 0)

            ag_f += f
            ag_cmd += cmd
            ag_p += pnum
            ag_c += cnum

            tx = pct_1dec(cmd, pnum if pnum > 0 else f)  # priorité prospects saisis, sinon fiches
            lines.append(f"- {ini}: fiches={f}, prospects={pnum}, clients={cnum}, commandes={cmd}, tx={tx}")

        lines.append(f"TOTAL {ag}: fiches={ag_f}, prospects={ag_p}, clients={ag_c}, commandes={ag_cmd}, tx={pct_1dec(ag_cmd, ag_p if ag_p>0 else ag_f)}")
        lines.append("")

        total_fiches += ag_f
        total_cmd += ag_cmd

    lines.append(f"TOTAL GLOBAL: fiches={total_fiches}, commandes={total_cmd}")

    # ---- CUMUL MENSUEL dans repo
    mkey = month_key(date)
    cumul_path = f"data/cumul_{mkey}.jsonl"
    cumul_rows = load_jsonl(cumul_path)

    # on ajoute stats_day (dédup basique)
    existing_keys = set((r.get("date",""), r.get("agency",""), r.get("initials",""), r.get("period","")) for r in cumul_rows)
    for s in stats_day:
        k = (s.get("date",""), s.get("agency",""), s.get("initials",""), s.get("period",""))
        if k in existing_keys:
            continue
        cumul_rows.append(s)
        existing_keys.add(k)

    save_jsonl(cumul_path, cumul_rows)

    # Excel cumul mensuel (toi uniquement)
    excel_stats = build_excel_global_stats(cumul_rows)

    # ---- send global to you
    global_to = cfg.get("global_to", [])
    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date}",
        "\n".join(lines),
        global_to,
        [(f"{mkey}_CUMUL_STATS.xlsx", excel_stats)]
    )

if __name__ == "__main__":
    main()
