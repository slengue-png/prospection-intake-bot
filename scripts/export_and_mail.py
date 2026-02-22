import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime, date
from zoneinfo import ZoneInfo
from openpyxl import Workbook
import yaml
import re
from pathlib import Path
from typing import List, Tuple, Optional

import requests
from PIL import Image
import pytesseract

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
    "INFOS_COMMERCIALES",
    "CARTE_VISITE_FILE_ID",
]

BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept": "*/*",
    "Accept-Language": "fr-FR,fr;q=0.9",
    "Connection": "close",
}

# =========================
# CONFIG
# =========================
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
    out = []
    for e in raw:
        e = (e or "").strip()
        if e:
            out.append(e)
    return out

# =========================
# DATE (Europe/Paris by default)
# =========================
def paris_today() -> date:
    return datetime.now(ZoneInfo("Europe/Paris")).date()

def parse_date_yyyy_mm_dd(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()

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
# TEL / EMAIL PARSING
# =========================
EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
PHONE_RE = re.compile(r"(?:(?:\+|00)33|0)\s*[1-9](?:[\s.\-]*\d{2}){4}")

def normalize_phone(raw: str) -> str:
    if not raw:
        return ""
    s = str(raw).strip()
    s = s.replace("(0)", "")
    s = re.sub(r"[^\d+]", "", s)
    if s.startswith("+33"):
        s = "0" + s[3:]
    digits = re.sub(r"\D", "", s)
    if len(digits) < 10:
        return ""
    return digits[:10]

def phone_allowed(num: str) -> bool:
    return bool(num) and len(num) == 10 and (num.startswith("04") or num.startswith("06") or num.startswith("07"))

def pick_best_phones(nums: List[str]) -> Tuple[str, str]:
    fixed = ""
    mobile = ""
    for n in nums:
        if not phone_allowed(n):
            continue
        if n.startswith("04") and not fixed:
            fixed = n
        if (n.startswith("06") or n.startswith("07")) and not mobile:
            mobile = n
    return fixed, mobile

def extract_emails(text: str) -> List[str]:
    if not text:
        return []
    emails = [e.lower() for e in EMAIL_RE.findall(text)]
    bad = ["noreply", "no-reply", "donotreply", "example", "exemple"]
    out = []
    for e in emails:
        if any(b in e for b in bad):
            continue
        out.append(e)
    seen = set()
    res = []
    for e in out:
        if e not in seen:
            seen.add(e)
            res.append(e)
    return res

def extract_phones(text: str) -> List[str]:
    if not text:
        return []
    found = []
    for m in PHONE_RE.findall(text):
        n = normalize_phone(m)
        if n:
            found.append(n)
    seen = set()
    res = []
    for n in found:
        if n not in seen:
            seen.add(n)
            res.append(n)
    return res

# =========================
# TELEGRAM DOWNLOAD (card photo)
# =========================
def tg_api_get_file(token: str, file_id: str) -> Optional[str]:
    url = f"https://api.telegram.org/bot{token}/getFile"
    r = requests.get(url, params={"file_id": file_id}, timeout=20)
    r.raise_for_status()
    j = r.json()
    if not j.get("ok"):
        return None
    return j["result"].get("file_path")

def tg_download_file(token: str, file_path: str) -> bytes:
    url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content

# =========================
# OCR TESSERACT
# =========================
def ocr_card_image(img_bytes: bytes) -> str:
    with Image.open(io.BytesIO(img_bytes)) as im:
        im = im.convert("RGB")
        text = pytesseract.image_to_string(im, lang="fra")
        return text or ""

def enrich_from_ocr(record: dict, ocr_text: str):
    emails = extract_emails(ocr_text)
    phones = extract_phones(ocr_text)
    fixed, mobile = pick_best_phones(phones)

    if not str(record.get("email", "")).strip() and emails:
        record["email"] = emails[0]

    if not str(record.get("phone", "")).strip() and fixed:
        record["phone"] = fixed

    if not str(record.get("phone2", "")).strip() and mobile:
        record["phone2"] = mobile

# =========================
# FREE EMAIL GRAB (site scrape) - GARDE FOU
# =========================
def fetch_text(url: str, timeout=10) -> str:
    try:
        r = requests.get(url, headers=BROWSER_HEADERS, timeout=timeout, allow_redirects=True)
        if r.status_code != 200:
            return ""
        return r.text or ""
    except:
        return ""

def find_contact_pages(home_html: str, base: str) -> List[str]:
    links = set()
    for m in re.finditer(r'href=["\']([^"\']+)["\']', home_html or "", flags=re.IGNORECASE):
        href = m.group(1).strip()
        if not href or href.startswith("#"):
            continue
        if href.startswith("//"):
            href = "https:" + href
        if href.startswith("/"):
            href = base.rstrip("/") + href
        if not href.startswith(base):
            continue
        low = href.lower()
        if any(k in low for k in ["contact", "mentions", "legal", "rgpd", "privacy"]):
            links.add(href.split("#")[0])
        if len(links) >= 6:
            break
    return list(links)

def enrich_email_from_website(record: dict):
    if record.get("email"):
        return
    website = (record.get("website") or "").strip()
    if not website.startswith("http"):
        return

    base = website.rstrip("/")
    home = fetch_text(base, timeout=10)
    if not home:
        return

    candidates = [
        base,
        base + "/contact",
        base + "/contactez-nous",
        base + "/nous-contacter",
        base + "/mentions-legales",
        base + "/mentions",
    ]
    candidates.extend(find_contact_pages(home, base))

    checked = 0
    for u in candidates:
        if checked >= 4:
            break
        checked += 1
        html = fetch_text(u, timeout=10)
        if not html:
            continue
        emails = extract_emails(html)
        if emails:
            record["email"] = emails[0]
            return

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
        if not ww or ww in PARASITE_WORDS or ww.isdigit():
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
        infos_com = " | ".join([x for x in [
            f"Interlocuteur: {r.get('interlocuteur','')}".strip(),
            f"Entretien: {r.get('resume','')}".strip(),
            f"Commande: {r.get('commande','')}".strip(),
        ] if x and x not in ("Interlocuteur:", "Entretien:", "Commande:")])

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
            infos_com,
            r.get("card_photo_file_id", ""),
        ])

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()

# =========================
# EMAIL (Brevo)
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
        print("[BREVO] status:", resp.status, "| to:", ",".join(to_list), "| subject:", subject)
        _ = resp.read()

# =========================
# CUMUL MENSUEL
# =========================
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
    out = {"agencies": {}, "by_user": {}, "totals": {"prospects": 0, "clients": 0, "commandes": 0}}
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
    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError("config.yml: worker_base_url must start with https://")

    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    telegram_token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    if not telegram_token:
        raise RuntimeError("TELEGRAM_TOKEN missing in GitHub Secrets (needed for card photos download)")

    # ✅ FIX: date Europe/Paris par défaut
    run_date_str = (os.environ.get("RUN_DATE") or "").strip()
    d = parse_date_yyyy_mm_dd(run_date_str) if run_date_str else paris_today()
    date_str = d.strftime("%Y-%m-%d")

    # ✅ anti-confusion Gmail: on met un timestamp d'exécution
    run_ts = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y%m%d-%H%M%S")

    agencies_cfg = cfg.get("agencies", {})

    url = f"{worker}/dump?date={date_str}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] date=", date_str, "| prospects rows=", len(prospects))

    # ===== OCR + enrich =====
    for r in prospects:
        file_id = (r.get("card_photo_file_id") or "").strip()
        if file_id:
            try:
                fp = tg_api_get_file(telegram_token, file_id)
                if fp:
                    img_bytes = tg_download_file(telegram_token, fp)
                    txt = ocr_card_image(img_bytes)
                    enrich_from_ocr(r, txt)
            except Exception as e:
                print("[OCR][WARN]", str(e)[:160])

        try:
            enrich_email_from_website(r)
        except Exception as e:
            print("[WEB][WARN]", str(e)[:160])

    # ===== split par agence =====
    agency_records = {ag: [] for ag in agencies_cfg.keys()}
    for p in prospects:
        ag = (p.get("agency") or "").strip()
        if ag in agency_records:
            agency_records[ag].append(p)

    # ===== emails agences (excel + PJ cartes) =====
    for ag, recs in agency_records.items():
        to_list = agencies_cfg.get(ag, {}).get("daily_to", [])

        excel = build_excel(recs)

        # PJ cartes: toutes les cartes de l'agence
        card_attachments = []
        for i, r in enumerate(recs, 1):
            file_id = (r.get("card_photo_file_id") or "").strip()
            if not file_id:
                continue
            try:
                fp = tg_api_get_file(telegram_token, file_id)
                if not fp:
                    print(f"[CARD][WARN] getFile failed agency={ag} i={i} file_id={file_id[:12]}...")
                    continue
                img = tg_download_file(telegram_token, fp)
                # ✅ nom unique
                card_attachments.append((f"{date_str}_{ag}_{run_ts}_carte_{i}.jpg", img))
            except Exception as e:
                print("[CARD][WARN]", str(e)[:160])

        nb_prospects = len(recs)
        nb_clients = 0
        nb_commandes = sum(count_commandes(r.get("commande", "")) for r in recs)

        attachments = [(f"{date_str}_{ag}_{run_ts}_IMPORT.xlsx", excel)]
        attachments.extend(card_attachments)

        subject = f"[PROSPECTION] {ag} — {date_str} ({run_ts})"
        body = (
            f"Export agence {ag}\n"
            f"Date: {date_str}\n"
            f"Run: {run_ts}\n\n"
            f"Prospects: {nb_prospects}\n"
            f"Clients: {nb_clients}\n"
            f"Commandes: {nb_commandes}\n"
            f"Cartes de visite jointes: {len(card_attachments)}\n"
        )

        print(f"[SEND] {ag} -> prospects={nb_prospects} cards={len(card_attachments)} to={to_list}")

        send_email_brevo(subject, body, to_list, attachments)

    # ===== cumul mensuel =====
    cumul_file = cumul_path_for(d)
    cumul = load_cumul(cumul_file)
    day_agg = aggregate_day(prospects, agencies_cfg)
    cumul["days"][date_str] = day_agg
    save_cumul(cumul_file, cumul)
    print("[CUMUL] updated:", str(cumul_file))

if __name__ == "__main__":
    main()