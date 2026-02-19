import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime, date
from openpyxl import Workbook
import yaml
import re
from pathlib import Path
from typing import Dict, List, Tuple, Any

# =========================
# EXCEL EXPORT (AGENCES)
# =========================
# ✅ Mise à jour selon ton dernier tableau (on garde tes champs + on ajoute visites)
EXPORT_COLUMNS = [
    "NOM",                      # raison sociale
    "ADRESSE",                  # rue
    "CODE POSTAL",
    "VILLE",
    "TELEPHONE",                # fixe (ou principal)
    "MOBILE",                   # mobile
    "MAIL",
    "SIRET",
    "NAF",
    "SITE WEB",
    "DIRIGEANT 1 : civilité",
    "DIRIGEANT 1 : prénom",
    "DIRIGEANT 1 : nom",
    "COMMERCE",                 # (si tu veux le commercial/initiales ici)
    "ENTRETIEN",
    "COMMANDE",
    "NOMBRE DE VISITES CLIENTS",
    "NOMBRE DE VISITES PROSPECTS",
]

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "fr-FR,fr;q=0.9",
    "Connection": "close",
}

# =========================
# CONFIG
# =========================
def load_config(path="config.yml"):
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}

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

def safe_int(x, default=0) -> int:
    try:
        return int(str(x).strip())
    except:
        return default

# =========================
# FETCH /dump (Worker)
# =========================
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

# =========================
# COMMANDES -> comptage intelligent
# =========================
PARASITE_WORDS = {
    "besoin","poste","postes","profil","profils","recherche","recherchons",
    "urgent","urgente","asap","svp","stp","merci","cdi","cdd","interim",
    "intérim","mission","missions","pour","de","des","du","la","le","les",
    "a","à","au","aux","et","ou","sur","en","d","l","un","une",
}

def _stem_fr(word: str) -> str:
    w = word.strip().lower()
    w = re.sub(r"[^a-zàâäéèêëîïôöùûüç0-9\-]", "", w)
    if not w:
        return ""
    for suf in ["es", "s", "x"]:
        if w.endswith(suf) and len(w) > 3:
            w = w[: -len(suf)]
            break
    return w

def count_commandes(commande_text: str) -> int:
    if not commande_text:
        return 0
    s = str(commande_text).strip()
    if not s:
        return 0

    if re.fullmatch(r"\d+", s):
        return 1

    s = s.replace("\n", " ")
    s = re.sub(r"[;,/|]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()

    groups = re.findall(r"\b\d+\s+([A-Za-zÀ-ÖØ-öø-ÿ\-]+)", s)
    if groups:
        uniq = set()
        for g in groups:
            w = _stem_fr(g)
            if w and w not in PARASITE_WORDS:
                uniq.add(w)
        return max(1, len(uniq)) if uniq else 1

    words = re.split(r"\s+", s)
    uniq = set()
    for w in words:
        ww = _stem_fr(w)
        if not ww:
            continue
        if ww in PARASITE_WORDS:
            continue
        if ww.isdigit():
            continue
        uniq.add(ww)

    return len(uniq) if uniq else 1

# =========================
# EXCEL
# =========================
def build_excel(records: List[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "IMPORT"
    ws.append(EXPORT_COLUMNS)

    for r in records:
        # ✅ si interlocuteur vide => dirigeant
        dirigeant_civ = r.get("dirigeant_civility", "") or r.get("dirigeant1_civility", "") or ""
        dirigeant_fn  = r.get("dirigeant_firstname", "") or r.get("dirigeant1_firstname", "") or ""
        dirigeant_ln  = r.get("dirigeant_lastname", "") or r.get("dirigeant1_lastname", "") or ""

        interlocuteur = (r.get("interlocuteur") or "").strip()
        if not interlocuteur:
            # on synthétise un "interlocuteur" à partir du dirigeant si besoin côté logique bot
            pass

        ws.append([
            r.get("name", ""),
            r.get("address", ""),
            r.get("postal_code", ""),
            r.get("city", ""),
            r.get("phone", ""),          # fixe
            r.get("mobile", "") or r.get("phone2", ""),  # mobile
            r.get("email", ""),
            r.get("siret", ""),
            r.get("naf", ""),
            r.get("website", ""),
            dirigeant_civ,
            dirigeant_fn,
            dirigeant_ln,
            r.get("initials", "") or r.get("commerce", ""),   # COMMERCE
            r.get("resume", "") or r.get("entretien", ""),    # ENTRETIEN
            r.get("commande", ""),
            safe_int(r.get("visites_clients", 0), 0),
            safe_int(r.get("visites_prospects", 0), 0),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# =========================
# EMAIL (Brevo API)
# =========================
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

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
            _ = resp.read()
    except urllib.error.HTTPError as e:
        err = e.read().decode("utf-8", errors="replace")
        print("[BREVO][HTTPERROR]", e.code, err[:400].replace("\n", " "))
        raise

# =========================
# CUMUL MENSUEL (repo)
# =========================
def parse_date_yyyy_mm_dd(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()

def last_day_of_month(d: date) -> date:
    if d.month == 12:
        return date(d.year, 12, 31)
    nxt = date(d.year, d.month + 1, 1)
    return nxt.fromordinal(nxt.toordinal() - 1)

def month_key(d: date) -> str:
    return f"{d.year:04d}-{d.month:02d}"

def cumul_path_for(d: date) -> Path:
    Path("data").mkdir(parents=True, exist_ok=True)
    return Path("data") / f"cumul_{month_key(d)}.json"

def load_cumul(path: Path) -> dict:
    if not path.exists():
        return {"month": path.stem.replace("cumul_", ""), "days": {}, "updated_at_utc": ""}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except:
        return {"month": path.stem.replace("cumul_", ""), "days": {}, "updated_at_utc": ""}

def save_cumul(path: Path, data: dict):
    data["updated_at_utc"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def aggregate_day(prospects: List[dict], agencies: dict) -> dict:
    out = {
        "agencies": {},
        "by_user": {},   # key: "AG|INI"
        "totals": {"prospects": 0, "clients": 0, "commandes": 0, "visites_clients": 0, "visites_prospects": 0},
    }

    for ag in agencies.keys():
        out["agencies"][ag] = {"prospects": 0, "clients": 0, "commandes": 0, "visites_clients": 0, "visites_prospects": 0}

    for r in prospects:
        ag = (r.get("agency") or "").strip()
        ini = (r.get("initials") or "NA").strip().upper()

        cmds = count_commandes(r.get("commande", ""))

        vc = safe_int(r.get("visites_clients", 0), 0)
        vp = safe_int(r.get("visites_prospects", 0), 0)

        if ag in out["agencies"]:
            out["agencies"][ag]["prospects"] += 1
            out["agencies"][ag]["commandes"] += cmds
            out["agencies"][ag]["visites_clients"] += vc
            out["agencies"][ag]["visites_prospects"] += vp

        key = f"{ag}|{ini}"
        out["by_user"].setdefault(key, {"agency": ag, "initials": ini, "prospects": 0, "clients": 0, "commandes": 0,
                                        "visites_clients": 0, "visites_prospects": 0})
        out["by_user"][key]["prospects"] += 1
        out["by_user"][key]["commandes"] += cmds
        out["by_user"][key]["visites_clients"] += vc
        out["by_user"][key]["visites_prospects"] += vp

        out["totals"]["prospects"] += 1
        out["totals"]["commandes"] += cmds
        out["totals"]["visites_clients"] += vc
        out["totals"]["visites_prospects"] += vp

    return out

def make_ranking(by_user: dict) -> list[dict]:
    rows = []
    for _, u in by_user.items():
        p = int(u.get("prospects", 0))
        c = int(u.get("clients", 0))
        cmd = int(u.get("commandes", 0))
        tx = (cmd / p * 100.0) if p else 0.0
        rows.append({
            "agency": u.get("agency", ""),
            "initials": u.get("initials", "NA"),
            "prospects": p,
            "clients": c,
            "commandes": cmd,
            "tx": tx,
            "visites_clients": int(u.get("visites_clients", 0)),
            "visites_prospects": int(u.get("visites_prospects", 0)),
        })
    rows.sort(key=lambda r: (r["commandes"], r["tx"], r["prospects"]), reverse=True)
    return rows

# =========================
# UTIL: MODE + payload
# =========================
def get_mode() -> str:
    m = (os.environ.get("MODE") or "").strip()
    return m or "agency_digest"

def get_client_payload() -> dict:
    raw = (os.environ.get("CLIENT_PAYLOAD_JSON") or "").strip()
    if not raw:
        return {}
    try:
        return json.loads(raw)
    except:
        return {}

# =========================
# MAIN
# =========================
def main():
    cfg = load_config()

    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError("config.yml: worker_base_url must start with https://")

    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing in GitHub Secrets")

    mode = get_mode()
    payload = get_client_payload()
    print("[MODE]", mode)
    if payload:
        print("[PAYLOAD] keys:", list(payload.keys()))

    # Date (priorité: payload.date -> RUN_DATE -> today UTC)
    run_date_str = (payload.get("date") or os.environ.get("RUN_DATE") or "").strip()
    if run_date_str:
        d = parse_date_yyyy_mm_dd(run_date_str)
    else:
        d = datetime.utcnow().date()
    date_str = d.strftime("%Y-%m-%d")

    agencies_cfg = cfg.get("agencies", {}) or {}
    global_to = cfg.get("global_to", []) or []

    # ===== 1) close_session : envoi immédiat par commercial =====
    if mode == "close_session":
        # On attend du payload au minimum : agency, initials, email_to (ou on le retrouve via config)
        ag = (payload.get("agency") or "").strip()
        ini = (payload.get("initials") or "").strip().upper()

        if not ag or not ini:
            raise RuntimeError("close_session requires payload {agency, initials}.")
        # Destinataire immédiat = email du commercial (config > payload)
        user_mail = None
        if agencies_cfg.get(ag, {}).get("users", {}).get(ini):
            user_mail = agencies_cfg[ag]["users"][ini].get("email")
        if not user_mail:
            user_mail = payload.get("user_email")

        if not user_mail:
            raise RuntimeError(f"Missing user email for {ag}/{ini}. Add it in config.yml agencies.{ag}.users.{ini}.email")

        # On peut joindre l'export de SA journée uniquement
        url = f"{worker}/dump?date={date_str}&kind=prospects&agency={ag}&initials={ini}"
        prospects = fetch_jsonl(url, export_token)
        print("[OK] prospects rows=", len(prospects))

        # On injecte les compteurs saisis en fin de session dans chaque ligne (pour Excel + cumul)
        visites_clients = safe_int(payload.get("visites_clients", 0), 0)
        visites_prospects = safe_int(payload.get("visites_prospects", 0), 0)
        for r in prospects:
            r["visites_clients"] = visites_clients
            r["visites_prospects"] = visites_prospects

        excel = build_excel(prospects)
        nb_prospects = len(prospects)
        nb_commandes = sum(count_commandes(r.get("commande", "")) for r in prospects)

        subject = f"[PROSPECTION] Clôture session — {ag}/{ini} — {date_str}"
        body = (
            f"Clôture session {ag}/{ini}\n"
            f"Date: {date_str}\n"
            f"Prospects saisis: {nb_prospects}\n"
            f"Commandes: {nb_commandes}\n"
            f"Visites clients (déclaré): {visites_clients}\n"
            f"Visites prospects (déclaré): {visites_prospects}\n"
        )
        send_email_brevo(subject, body, [user_mail], [(f"{date_str}_{ag}_{ini}_IMPORT.xlsx", excel)])

        # Cumul mensuel (on merge le jour, mais au moins on met à jour)
        cumul_file = cumul_path_for(d)
        cumul = load_cumul(cumul_file)

        # On stocke un mini-agrégat "par user" du close_session en plus (sans casser ton cumul)
        day = cumul["days"].get(date_str) or {"by_user_close": {}}
        day.setdefault("by_user_close", {})
        day["by_user_close"][f"{ag}|{ini}"] = {
            "visites_clients": visites_clients,
            "visites_prospects": visites_prospects,
            "updated_at_utc": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        }
        cumul["days"][date_str] = day
        save_cumul(cumul_file, cumul)
        print("[CUMUL] updated (close_session):", str(cumul_file))
        return

    # ===== 2) agency_digest / global_digest : on part du dump global du jour =====
    url = f"{worker}/dump?date={date_str}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(prospects))

    # Regroupement par agence
    agency_records = {ag: [] for ag in agencies_cfg.keys()}
    for p in prospects:
        ag = (p.get("agency") or "").strip()
        if ag in agency_records:
            agency_records[ag].append(p)

    # ===== 2a) agency_digest (17:45) : 1 mail par agence vers le "leader" =====
    if mode == "agency_digest":
        for ag, recs in agency_records.items():
            # ✅ destinataire récap agence : agencies.{AG}.daily_digest_to
            to_list = agencies_cfg.get(ag, {}).get("daily_digest_to", []) or []
            if not to_list:
                print(f"[SKIP] No daily_digest_to configured for agency {ag}")
                continue

            excel = build_excel(recs)

            nb_prospects = len(recs)
            nb_commandes = sum(count_commandes(r.get("commande", "")) for r in recs)

            # stats visites (addition des champs stockés par ligne si présents)
            vc = sum(safe_int(r.get("visites_clients", 0), 0) for r in recs)
            vp = sum(safe_int(r.get("visites_prospects", 0), 0) for r in recs)

            # Détail par commercial
            by_init = {}
            for r in recs:
                ini = (r.get("initials") or "NA").strip().upper()
                by_init.setdefault(ini, []).append(r)

            lines = [f"Récap agence {ag} — {date_str}", ""]
            lines.append(f"TOTAL — Prospects: {nb_prospects} | Commandes: {nb_commandes} | Visites clients: {vc} | Visites prospects: {vp}")
            lines.append("")
            for ini, rr in sorted(by_init.items()):
                pcount = len(rr)
                cmdcount = sum(count_commandes(x.get("commande", "")) for x in rr)
                tx = (cmdcount / pcount * 100.0) if pcount else 0.0
                vc_u = sum(safe_int(x.get("visites_clients", 0), 0) for x in rr)
                vp_u = sum(safe_int(x.get("visites_prospects", 0), 0) for x in rr)
                lines.append(f"- {ini}: {pcount} prospects | {cmdcount} commandes | Tx {tx:.1f}% | VC {vc_u} | VP {vp_u}")

            send_email_brevo(
                f"[PROSPECTION] {ag} — RÉCAP — {date_str}",
                "\n".join(lines),
                to_list,
                [(f"{date_str}_{ag}_IMPORT.xlsx", excel)],
            )
        return

    # ===== 2b) global_digest (17:47) : 1 mail pour toi (global_to) =====
    if mode == "global_digest":
        lines = [f"Récap GLOBAL — {date_str}", ""]
        total_p = 0
        total_cmd = 0
        total_vc = 0
        total_vp = 0

        for ag in agency_records.keys():
            recs = agency_records[ag]
            nb_p = len(recs)
            nb_cmd = sum(count_commandes(r.get("commande", "")) for r in recs)
            vc = sum(safe_int(r.get("visites_clients", 0), 0) for r in recs)
            vp = sum(safe_int(r.get("visites_prospects", 0), 0) for r in recs)

            total_p += nb_p
            total_cmd += nb_cmd
            total_vc += vc
            total_vp += vp

            lines.append(f"{ag} — Prospects: {nb_p} | Commandes: {nb_cmd} | VC: {vc} | VP: {vp}")

            by_init = {}
            for r in recs:
                ini = (r.get("initials") or "NA").strip().upper()
                by_init.setdefault(ini, []).append(r)

            for ini, rr in sorted(by_init.items()):
                pcount = len(rr)
                cmdcount = sum(count_commandes(x.get("commande", "")) for x in rr)
                tx = (cmdcount / pcount * 100.0) if pcount else 0.0
                vc_u = sum(safe_int(x.get("visites_clients", 0), 0) for x in rr)
                vp_u = sum(safe_int(x.get("visites_prospects", 0), 0) for x in rr)
                lines.append(f"  - {ini}: {pcount} prospects | {cmdcount} commandes | Tx {tx:.1f}% | VC {vc_u} | VP {vp_u}")
            lines.append("")

        lines.append(f"TOTAL — Prospects: {total_p} | Commandes: {total_cmd} | VC: {total_vc} | VP: {total_vp}")

        send_email_brevo(
            f"[PROSPECTION] GLOBAL — {date_str}",
            "\n".join(lines),
            global_to,
            [],
        )

        # Cumul mensuel (on garde ta logique)
        cumul_file = cumul_path_for(d)
        cumul = load_cumul(cumul_file)
        day_agg = aggregate_day(prospects, agencies_cfg)
        cumul["days"][date_str] = day_agg
        save_cumul(cumul_file, cumul)
        print("[CUMUL] updated:", str(cumul_file))
        return

    # Fallback
    raise RuntimeError(f"Unknown MODE={mode}")

if __name__ == "__main__":
    main()