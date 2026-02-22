import os
import io
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime, date
from pathlib import Path
from typing import List, Tuple, Optional
import re

import yaml
import requests
from openpyxl import Workbook
from PIL import Image
import pytesseract

# =========================
# EXCEL COLUMNS (AGENCES)
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

EMAIL_RE = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
PHONE_RE = re.compile(r"(?:(?:\+|00)33|0)\s*[1-9](?:[\s.\-]*\d{2}){4}")

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
# FETCH /dump (Worker) -> JSONL
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
# EMAIL / PHONE PARSING
# =========================
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

def pick_best_phones(nums: List[str]) -> Tuple[str, str]:
    fixed = ""
    mobile = ""
    for n in nums:
        if not n or len(n) != 10:
            continue
        if n.startswith("04") and not fixed:
            fixed = n
        if (n.startswith("06") or n.startswith("07")) and not mobile:
            mobile = n
    return fixed, mobile

# =========================
# TELEGRAM DOWNLOAD (card photo)
# =========================
def tg_api_get_file(token: str, file_id: str) -> Optional[dict]:
    url = f"https://api.telegram.org/bot{token}/getFile"
    r = requests.get(url, params={"file_id": file_id}, timeout=20)
    r.raise_for_status()
    j = r.json()
    if not j.get("ok"):
        return None
    return j["result"]  # {file_path, ...}

def tg_download_file(token: str, file_path: str) -> bytes:
    url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content

def file_ext_from_path(file_path: str) -> str:
    fp = (file_path or "").lower()
    for ext in [".jpg", ".jpeg", ".png", ".webp", ".pdf"]:
        if fp.endswith(ext):
            return ext
    return ".jpg"

# =========================
# OCR TESSERACT (FREE)
# =========================
def ocr_card_image(img_bytes: bytes) -> str:
    with Image.open(io.BytesIO(img_bytes)) as im:
        im = im.convert("RGB")
        return pytesseract.image_to_string(im, lang="fra") or ""

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
# FREE EMAIL GRAB (site scrape)
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
# EXCEL
# =========================
def build_excel(records):
    wb = Workbook()
    ws = wb.active
    ws.title = "IMPORT"
    ws.append(IMPORT_COLUMNS)

    for r in records:
        interloc = (r.get("interlocuteur") or "").strip()
        dirigeant = (r.get("dirigeant") or "").strip()

        # ✅ règle: si pas d’interlocuteur => contact = dirigeant
        contact_nom = interloc or dirigeant

        infos_com = (r.get("infos_commerciales") or "").strip()
        if not infos_com:
            # fallback
            chunks = []
            if contact_nom: chunks.append(f"Interlocuteur: {contact_nom}")
            if r.get("resume"): chunks.append(f"Entretien: {r.get('resume')}")
            if r.get("commande"): chunks.append(f"Commande: {r.get('commande')}")
            infos_com = " | ".join(chunks)

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
            r.get("contact_lastname", "") or contact_nom,
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
# MAIN
# =========================
def parse_date_yyyy_mm_dd(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()

def main():
    cfg = load_config()
    worker = (cfg.get("worker_base_url") or "").strip().rstrip("/")
    if not worker.startswith("https://"):
        raise RuntimeError("config.yml: worker_base_url must start with https://")

    export_token = os.environ.get("EXPORT_TOKEN", "").strip()
    telegram_token = os.environ.get("TELEGRAM_TOKEN", "").strip()
    if not telegram_token:
        raise RuntimeError("TELEGRAM_TOKEN missing in GitHub Secrets (needed for card photos download)")

    run_date_str = (os.environ.get("RUN_DATE") or "").strip()
    d = parse_date_yyyy_mm_dd(run_date_str) if run_date_str else datetime.utcnow().date()
    date_str = d.strftime("%Y-%m-%d")

    agencies_cfg = cfg.get("agencies", {})

    # dump prospects
    url = f"{worker}/dump?date={date_str}&kind=prospects"
    prospects = fetch_jsonl(url, export_token)
    print("[OK] prospects rows=", len(prospects))

    # OCR + enrich (site scraping)
    for r in prospects:
        file_id = (r.get("card_photo_file_id") or "").strip()
        if file_id:
            try:
                info = tg_api_get_file(telegram_token, file_id)
                if info and info.get("file_path"):
                    img_bytes = tg_download_file(telegram_token, info["file_path"])
                    txt = ocr_card_image(img_bytes)
                    enrich_from_ocr(r, txt)
            except Exception as e:
                print("[OCR][WARN]", str(e)[:160])

        try:
            enrich_email_from_website(r)
        except Exception as e:
            print("[WEB][WARN]", str(e)[:160])

        # règle contact: interlocuteur fallback -> dirigeant
        if not (r.get("interlocuteur") or "").strip() and (r.get("dirigeant") or "").strip():
            r["interlocuteur"] = r["dirigeant"]

    # split by agency
    agency_records = {ag: [] for ag in agencies_cfg.keys()}
    for p in prospects:
        ag = p.get("agency", "")
        if ag in agency_records:
            agency_records[ag].append(p)

    # email agencies (excel + card attachments)
    for ag, recs in agency_records.items():
        to_list = agencies_cfg.get(ag, {}).get("daily_to", [])
        excel = build_excel(recs)

        card_attachments = []
        for i, r in enumerate(recs, 1):
            file_id = (r.get("card_photo_file_id") or "").strip()
            if not file_id:
                continue
            try:
                info = tg_api_get_file(telegram_token, file_id)
                if not info or not info.get("file_path"):
                    continue
                ext = file_ext_from_path(info["file_path"])
                img = tg_download_file(telegram_token, info["file_path"])
                card_attachments.append((f"{date_str}_{ag}_carte_{i}{ext}", img))
            except Exception as e:
                print("[CARD][WARN]", str(e)[:160])

        attachments = [(f"{date_str}_{ag}_IMPORT.xlsx", excel)]
        attachments.extend(card_attachments)

        send_email_brevo(
            f"[PROSPECTION] {ag} — {date_str}",
            f"Export agence {ag}\nProspects: {len(recs)}\nCartes jointes: {len(card_attachments)}\n",
            to_list,
            attachments,
        )

    print("[DONE]")

if __name__ == "__main__":
    main()