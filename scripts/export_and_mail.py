import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime, timezone
from openpyxl import Workbook
import yaml
import calendar

IMPORT_COLUMNS = [
    "NOM","ADRESSE","CODE POSTAL","VILLE","TELEPHONE","TELEPHONE 2","MAIL",
    "SIRET","NAF","SITE WEB","Contact: civilité","Contact : prénom","Contact : nom",
    "RESUME ENTRETIEN","COMMANDE"
]

# ✅ colonne en plus (comme demandé)
IMPORT_COLUMNS_WITH_DIRIGEANT = IMPORT_COLUMNS + ["DIRIGEANT"]

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

# --------------------------
# ✅ Comptage commandes (TA logique)
# --------------------------
STOPWORDS = {
    "de","des","du","la","le","les","un","une","et","a","à","en","pour","sur",
    "urgent","urgence","svp","stp","merci","besoin","recherche","recherchons",
    "profil","profils","poste","postes","mission","interim","intérim","cdi","cdd","cd",
    "x"
}

ALIASES = {
    # rapprochements (sans pondération)
    "conducteur":"chauffeur",
    "conducteurs":"chauffeur",
    "chauffeurs":"chauffeur",
    "chauffeur":"chauffeur",
    "manutentionnaires":"manutentionnaire",
    "manutentionnaire":"manutentionnaire",
    "caristes":"cariste",
    "cariste":"cariste",
    "preparateurs":"preparateur",
    "préparateurs":"preparateur",
    "preparateur":"preparateur",
    "préparateur":"preparateur",
    "mecaniciens":"mecanicien",
    "mécaniciens":"mecanicien",
    "mecanicien":"mecanicien",
    "mécanicien":"mecanicien",
}

def _clean_tokens(s: str):
    s = (s or "").strip().lower()
    if not s:
        return [], []
    # normalise séparateurs
    for ch in [",", ";", ":", "/", "\\", "|", "+", "(", ")", "[", "]", "{", "}", ".", "!", "?", "\n", "\r", "\t"]:
        s = s.replace(ch, " ")
    s = " ".join(s.split())
    parts = s.split(" ")
    nums = []
    words = []
    for p in parts:
        if not p:
            continue
        if p.isdigit():
            nums.append(p)
            continue
        w = p.strip()
        if w in STOPWORDS:
            continue
        # retire un pluriel simple
        if len(w) > 3 and w.endswith("s"):
            w = w[:-1]
        w = ALIASES.get(w, w)
        if w in STOPWORDS:
            continue
        words.append(w)
    return nums, words

def count_commandes(commande_text: str) -> int:
    """
    Règles demandées:
    - "1 manutentionnaire" => 1
    - "1" => 1
    - "1 manutentionnaire 1 chauffeur" => 2
    - "manutentionnaire chauffeur" => 2 (2 qualifications diff)
    - "2 manutentionnaires" => 1
    """
    t = (commande_text or "").strip()
    if not t:
        return 0

    nums, words = _clean_tokens(t)

    # cas "1" seul (ou "2" seul etc.) => 1 commande
    if nums and not words:
        return 1

    # si on a des nombres: chaque occurrence de nombre = 1 commande
    # ("2 manutentionnaires" => 1, "1 manutentionnaire 1 chauffeur" => 2)
    if nums:
        return len(nums)

    # sinon: on compte les qualifications distinctes (après nettoyage + aliases)
    if words:
        distinct = []
        for w in words:
            if w not in distinct:
                distinct.append(w)
        return len(distinct)

    return 0

# --------------------------
# Excel (import agence)
# --------------------------
def build_excel(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "IMPORT"
    ws.append(IMPORT_COLUMNS_WITH_DIRIGEANT)

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
            # ✅ NAF déjà normalisé côté worker (ex: 4669C)
            r.get("naf",""),
            r.get("website",""),
            r.get("contact_civility",""),
            r.get("contact_firstname",""),
            r.get("contact_lastname","") or r.get("interlocuteur",""),
            r.get("resume",""),
            r.get("commande",""),
            # ✅ nouveau champ
            r.get("dirigeant",""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# --------------------------
# Brevo email
# --------------------------
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
        # fallback (mais idéalement une adresse vérifiée)
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

# --------------------------
# Stats
# --------------------------
def pct_1_decimal(n, d):
    if d <= 0:
        return "0,0%"
    v = (n / d) * 100.0
    return f"{v:.1f}%".replace(".", ",")

def month_key(date_str):
    # "2026-02-14" -> ("2026-02", year, month)
    y, m, _ = date_str.split("-")
    return f"{y}-{m}", int(y), int(m)

def main():
    cfg = load_config()
    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    # date du run (si vide => aujourd'hui UTC)
    run_date = (os.environ.get("RUN_DATE") or "").strip()
    if run_date:
        date = run_date
    else:
        date = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    # 1) fetch prospects
    url = f"{worker}/dump?date={date}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(prospects))

    agencies = cfg.get("agencies", {})
    agency_records = {ag: [] for ag in agencies.keys()}

    for p in prospects:
        ag = (p.get("agency") or "").strip()
        if ag in agency_records:
            agency_records[ag].append(p)

    # 2) email agences (excel import)
    for ag, recs in agency_records.items():
        to_list = agencies[ag].get("daily_to", [])
        excel = build_excel(recs)
        send_email_brevo(
            f"[PROSPECTION] {ag} — Export import — {date}",
            f"Export prospection inconnus — agence {ag} — {date}\nFiches: {len(recs)}\n",
            to_list,
            [(f"{date}_{ag}_IMPORT.xlsx", excel)]
        )

    # 3) email global (toi uniquement): prospects + clients + commandes + tx + classement
    # 👉 Clients: pour l’instant = 0 (tant qu’on n’a pas activé un export “stats” côté worker)
    # Prospects = nb de fiches prospects
    # Commandes = comptage via la saisie "COMMANDE"
    global_to = cfg.get("global_to", [])

    # stats par agence/personne
    per_agency = {}
    per_person = {}  # (ag, initials) -> dict
    total_prospects = 0
    total_clients = 0  # placeholder tant qu’on n’a pas “stats clients”
    total_commandes = 0

    for ag, recs in agency_records.items():
        p_count = len(recs)
        c_count = 0
        cmd_count = sum(count_commandes(r.get("commande","")) for r in recs)

        per_agency[ag] = {"prospects": p_count, "clients": c_count, "commandes": cmd_count}
        total_prospects += p_count
        total_clients += c_count
        total_commandes += cmd_count

        # par commercial
        for r in recs:
            initials = (r.get("initials") or "NA").strip().upper()
            key = (ag, initials)
            if key not in per_person:
                per_person[key] = {"agency": ag, "initials": initials, "prospects": 0, "clients": 0, "commandes": 0}
            per_person[key]["prospects"] += 1
            per_person[key]["commandes"] += count_commandes(r.get("commande",""))

    # classement: commandes desc, puis tx desc
    ranking = list(per_person.values())
    def sort_key(x):
        tx = (x["commandes"] / x["prospects"]) if x["prospects"] else 0.0
        return (-x["commandes"], -tx, x["agency"], x["initials"])
    ranking.sort(key=sort_key)

    # bloc résumé global
    lines = []
    lines.append(f"Résumé global prospection — {date}")
    lines.append("")
    lines.append("=== Par agence ===")
    for ag in agencies.keys():
        a = per_agency.get(ag, {"prospects":0,"clients":0,"commandes":0})
        tx = pct_1_decimal(a["commandes"], a["prospects"])
        lines.append(f"{ag}: Prospects {a['prospects']} | Clients {a['clients']} | Commandes {a['commandes']} | Tx {tx}")
    lines.append("")
    lines.append(f"TOTAL: Prospects {total_prospects} | Clients {total_clients} | Commandes {total_commandes} | Tx {pct_1_decimal(total_commandes, total_prospects)}")
    lines.append("")
    lines.append("=== Classement commerciaux (commandes puis tx) ===")
    if not ranking:
        lines.append("(aucune donnée)")
    else:
        for i, r in enumerate(ranking, 1):
            tx = pct_1_decimal(r["commandes"], r["prospects"])
            lines.append(f"{i}. {r['agency']} / {r['initials']} — Prospects {r['prospects']} | Commandes {r['commandes']} | Tx {tx}")

    # bilans 15 jours (1–15 / 16–fin) basé sur la date
    mk, y, m = month_key(date)
    day_num = int(date.split("-")[2])
    period = "01-15" if day_num <= 15 else "16-fin"
    lines.append("")
    lines.append(f"=== Bilan période {mk} ({period}) ===")
    lines.append("⚠️ Pour un cumul exact multi-jours, il faut un fichier cumul (on le mettra à l’étape “cumul mensuel”).")
    lines.append("Pour l’instant, ce bloc reflète la journée courante (pas l’addition des jours).")

    send_email_brevo(
        f"[PROSPECTION] GLOBAL — Résumé — {date}",
        "\n".join(lines),
        global_to,
        []
    )

if __name__ == "__main__":
    main()
