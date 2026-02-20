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

# OCR (tesseract via workflow)
import subprocess
import tempfile


# =========================
# EXCEL IMPORT (AGENCES)
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
    "PHOTO_CARTE",
    "INFOS_COMMERCIALES",
]

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122 Safari/537.36",
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


# =========================
# DATE
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
        raise RuntimeError(f"HTTPError {e.code} on {url} -> {body[:500]}")
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
# OCR carte de visite (via Telegram file_id)
# =========================
def tg_api_call(method: str, payload: dict) -> dict:
    token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    if not token:
        return {}
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        f"https://api.telegram.org/bot{token}/{method}",
        data=data,
        method="POST",
        headers={"Content-Type": "application/json", "User-Agent": "prospection-intake-export/1.0"},
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode("utf-8", errors="replace"))
    except Exception:
        return {}


def tg_download_file(file_path: str) -> bytes:
    token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    if not token or not file_path:
        return b""
    url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    req = urllib.request.Request(url, headers={"User-Agent": "prospection-intake-export/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=45) as resp:
            return resp.read()
    except Exception:
        return b""


EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.I)
PHONE_RE = re.compile(r"(?:(?:\+|00)33|0)\s*[1-9](?:[\s.\-]*\d{2}){4}")


def normalize_phone_fr(s: str) -> str:
    s = (s or "").strip()
    s = s.replace("(0)", "")
    s = re.sub(r"[^\d+]", "", s)
    if s.startswith("+33"):
        s = "0" + s[3:]
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 10:
        digits = digits[:10]
    if len(digits) != 10:
        return ""
    return digits


def ocr_extract_text(image_bytes: bytes) -> str:
    if not image_bytes:
        return ""
    # écrit un fichier temporaire
    with tempfile.TemporaryDirectory() as td:
        img_path = os.path.join(td, "card.jpg")
        with open(img_path, "wb") as f:
            f.write(image_bytes)

        # tesseract
        try:
            out = subprocess.check_output(["tesseract", img_path, "stdout", "-l", "fra"], stderr=subprocess.DEVNULL)
            return out.decode("utf-8", errors="replace")
        except Exception:
            return ""


def enrich_from_card(record: dict) -> dict:
    """Si email/phone manquants et photo_file_id présente, tente OCR + extraction."""
    file_id = (record.get("photo_file_id") or "").strip()
    if not file_id:
        return record

    # déjà rempli => on ne casse pas
    cur_email = (record.get("email") or "").strip()
    cur_phone = (record.get("phone") or "").strip()

    # getFile
    js = tg_api_call("getFile", {"file_id": file_id})
    if not js.get("ok"):
        return record
    file_path = (((js.get("result") or {}).get("file_path")) or "").strip()
    if not file_path:
        return record

    img = tg_download_file(file_path)
    txt = ocr_extract_text(img)
    if not txt:
        return record

    # emails
    if not cur_email:
        m = EMAIL_RE.findall(txt)
        if m:
            record["email"] = m[0].lower()

    # phones
    if not cur_phone:
        m2 = PHONE_RE.findall(txt)
        for raw in m2:
            p = normalize_phone_fr(raw)
            if p and (p.startswith("04") or p.startswith("06") or p.startswith("07")):
                record["phone"] = p
                break

    # stocke un extrait utile
    record["commercial_info"] = (record.get("commercial_info") or "")
    if not record["commercial_info"]:
        record["commercial_info"] = "OCR carte: " + " ".join(txt.split())[:200]

    return record


# =========================
# COMMANDES -> comptage intelligent (ton code)
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
            "Oui" if (r.get("photo_file_id") or "").strip() else "Non",
            r.get("commercial_info", ""),
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

    with urllib.request.urlopen(req, timeout=30) as resp:
        _ = resp.read()
        print("[BREVO] status:", resp.status, "| subject:", subject)


# =========================
# CUMUL MENSUEL (repo)
# =========================
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


def aggregate_day(prospects: list[dict], agencies: dict, clients_visits: int, prospects_visits: int) -> dict:
    out = {
        "agencies": {},
        "by_user": {},
        "totals": {"prospects": 0, "clients": int(clients_visits), "commandes": 0, "prospects_visits": int(prospects_visits)},
    }
    for ag in agencies.keys():
        out["agencies"][ag] = {"prospects": 0, "clients": 0, "commandes": 0}

    for r in prospects:
        ag = (r.get("agency") or "").strip()
        ini = (r.get("initials") or "NA").strip().upper()
        cmds = count_commandes(r.get("commande", ""))

        if ag in out["agencies"]:
            out["agencies"][ag]["prospects"] += 1
            out["agencies"][ag]["commandes"] += cmds

        key = f"{ag}|{ini}"
        out["by_user"].setdefault(key, {"agency": ag, "initials": ini, "prospects": 0, "clients": 0, "commandes": 0})
        out["by_user"][key]["prospects"] += 1
        out["by_user"][key]["commandes"] += cmds

        out["totals"]["prospects"] += 1
        out["totals"]["commandes"] += cmds

    return out


# =========================
# MAIN
# =========================
def main():
    cfg = load_config()
    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    if not export_token:
        raise RuntimeError("EXPORT_TOKEN missing")

    # URL worker: secret > config.yml
    worker = (os.environ.get("WORKER_BASE_URL") or "").strip().rstrip("/")
    if not worker:
        worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError("Worker URL missing/invalid. Set WORKER_BASE_URL secret in GitHub.")

    run_date_str = (os.environ.get("RUN_DATE") or "").strip()
    if run_date_str:
        d = parse_date_yyyy_mm_dd(run_date_str)
    else:
        d = datetime.utcnow().date()
    date_str = d.strftime("%Y-%m-%d")

    agencies_cfg = cfg.get("agencies", {})
    global_to = cfg.get("global_to", [])

    # dispatch inputs
    clients_visits = int((os.environ.get("CLIENTS") or "0").strip() or "0")
    prospects_visits = int((os.environ.get("PROSPECTS") or "0").strip() or "0")

    # ===== dump prospects du jour =====
    url = f"{worker}/dump?date={date_str}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(prospects))

    # OCR/enrich depuis carte si dispo
    for r in prospects:
        r = enrich_from_card(r)

    agency_records = {ag: [] for ag in agencies_cfg.keys()}
    for p in prospects:
        ag = p.get("agency", "")
        if ag in agency_records:
            agency_records[ag].append(p)

    # ===== emails agences (excel import) =====
    for ag, recs in agency_records.items():
        to_list = agencies_cfg.get(ag, {}).get("daily_to", [])
        excel = build_excel(recs)

        nb_prospects = len(recs)
        nb_commandes = sum(count_commandes(r.get("commande", "")) for r in recs)

        send_email_brevo(
            f"[PROSPECTION] {ag} — {date_str}",
            f"Export agence {ag}\n"
            f"Prospects saisis: {nb_prospects}\n"
            f"Visites prospects (déclarées): {prospects_visits}\n"
            f"Visites clients (déclarées): {clients_visits}\n"
            f"Commandes: {nb_commandes}\n",
            to_list,
            [(f"{date_str}_{ag}_IMPORT.xlsx", excel)],
        )

    # ===== cumul mensuel =====
    cumul_file = cumul_path_for(d)
    cumul = load_cumul(cumul_file)
    cumul["days"][date_str] = aggregate_day(prospects, agencies_cfg, clients_visits, prospects_visits)
    save_cumul(cumul_file, cumul)
    print("[CUMUL] updated:", str(cumul_file))

    # ===== email global =====
    lines = [
        f"Résumé global prospection — {date_str}",
        "",
        f"Visites clients (déclarées): {clients_visits}",
        f"Visites prospects (déclarées): {prospects_visits}",
        "",
    ]
    total_p = 0
    total_cmd = 0

    for ag in agency_records.keys():
        recs = agency_records[ag]
        nb_p = len(recs)
        nb_cmd = sum(count_commandes(r.get("commande", "")) for r in recs)
        total_p += nb_p
        total_cmd += nb_cmd
        lines.append(f"{ag} — Prospects saisis: {nb_p} | Commandes: {nb_cmd}")

    lines.append("")
    lines.append(f"TOTAL prospects saisis: {total_p} | TOTAL commandes: {total_cmd}")

    send_email_brevo(
        f"[PROSPECTION] GLOBAL — {date_str}",
        "\n".join(lines),
        global_to,
        [],
    )


if __name__ == "__main__":
    main()