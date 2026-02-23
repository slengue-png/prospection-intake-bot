import os
import re
import json
import time
import smtplib
import mimetypes
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from PIL import Image, ImageOps, ImageEnhance
import pytesseract


# =========================
# ENV
# =========================
WORKER_DUMP_URL = os.getenv("WORKER_DUMP_URL", "").strip()
EXPORT_TOKEN = os.getenv("EXPORT_TOKEN", "").strip()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "").strip()

SMTP_HOST = os.getenv("SMTP_HOST", "").strip()
SMTP_PORT = int(os.getenv("SMTP_PORT", "587").strip())
SMTP_USER = os.getenv("SMTP_USER", "").strip()
SMTP_PASS = os.getenv("SMTP_PASS", "").strip()
MAIL_FROM = os.getenv("MAIL_FROM", "").strip()

AGENCY_RECIPIENTS_JSON = os.getenv("AGENCY_RECIPIENTS_JSON", "").strip()

OUT_DIR = Path(os.getenv("OUT_DIR", "out"))
OUT_DIR.mkdir(parents=True, exist_ok=True)

RUN_DATE = os.getenv("RUN_DATE", "").strip()
AGENCY = os.getenv("AGENCY", "").strip().upper()
INITIALS = os.getenv("INITIALS", "").strip().upper()
VISITS_CLIENTS = os.getenv("VISITS_CLIENTS", "").strip()
VISITS_PROSPECTS = os.getenv("VISITS_PROSPECTS", "").strip()
COMMANDES = os.getenv("COMMANDES", "").strip()


# =========================
# Data model
# =========================
@dataclass
class Prospect:
    name: str = ""
    address: str = ""
    postal_code: str = ""
    city: str = ""
    phone: str = ""
    phone2: str = ""
    email: str = ""
    siret: str = ""
    naf: str = ""
    website: str = ""
    contact_civility: str = ""
    contact_firstname: str = ""
    contact_lastname: str = ""
    dirigeant: str = ""
    resume: str = ""
    commande: str = ""
    infos_commerciales: str = ""
    card_file_id: str = ""
    agency: str = ""
    initials: str = ""


# =========================
# Helpers
# =========================
def must(val: str, name: str) -> str:
    if not val:
        raise RuntimeError(f"Missing required env: {name}")
    return val


def normalize_naf(code: str) -> str:
    return re.sub(r"[^0-9A-Z]", "", (code or "").upper())


def clean(s: Any) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip())


def ensure_http_origin(url: str) -> str:
    url = (url or "").strip()
    if not url:
        return ""
    try:
        from urllib.parse import urlparse
        p = urlparse(url)
        if not p.scheme:
            return ""
        return f"{p.scheme}://{p.netloc}"
    except Exception:
        return ""


# =========================
# Fetch prospects from worker dump
# =========================
def fetch_dump_jsonl(date_yyyy_mm_dd: str) -> List[Dict[str, Any]]:
    dump_url = must(WORKER_DUMP_URL, "WORKER_DUMP_URL")
    token = must(EXPORT_TOKEN, "EXPORT_TOKEN")

    r = requests.get(
        dump_url,
        params={"date": date_yyyy_mm_dd, "kind": "prospects"},
        headers={"X-Export-Token": token},
        timeout=60,
    )
    r.raise_for_status()

    lines = [ln.strip() for ln in r.text.splitlines() if ln.strip()]
    out = []
    for ln in lines:
        try:
            out.append(json.loads(ln))
        except Exception:
            # ignore bad line
            continue
    return out


def to_prospect(obj: Dict[str, Any]) -> Prospect:
    return Prospect(
        name=clean(obj.get("name")),
        address=clean(obj.get("address")),
        postal_code=clean(obj.get("postal_code")),
        city=clean(obj.get("city")),
        phone=clean(obj.get("phone")),
        phone2=clean(obj.get("phone2")),
        email=clean(obj.get("email")),
        siret=clean(obj.get("siret")),
        naf=normalize_naf(clean(obj.get("naf"))),
        website=ensure_http_origin(clean(obj.get("website"))),
        contact_civility=clean(obj.get("contact_civility")),
        contact_firstname=clean(obj.get("contact_firstname")),
        contact_lastname=clean(obj.get("contact_lastname")),
        dirigeant=clean(obj.get("dirigeant")),
        resume=clean(obj.get("resume")),
        commande=clean(obj.get("commande")),
        infos_commerciales=clean(obj.get("infos_commerciales")),
        card_file_id=clean(obj.get("card_photo_file_id") or obj.get("card_file_id") or ""),
        agency=clean(obj.get("agency")).upper(),
        initials=clean(obj.get("initials")).upper(),
    )


# =========================
# Telegram file download
# =========================
def telegram_get_file_path(file_id: str) -> Optional[str]:
    if not TELEGRAM_TOKEN or not file_id:
        return None
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=30)
    if r.status_code != 200:
        return None
    j = r.json()
    if not j.get("ok"):
        return None
    return j["result"].get("file_path")


def telegram_download_file(file_id: str, out_path: Path) -> bool:
    fp = telegram_get_file_path(file_id)
    if not fp:
        return False
    file_url = f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{fp}"
    r = requests.get(file_url, timeout=60)
    if r.status_code != 200:
        return False
    out_path.write_bytes(r.content)
    return True


# =========================
# OCR (contact name + email + phone)
# =========================
EMAIL_RE = re.compile(r"[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}", re.I)
PHONE_RE = re.compile(r"(\+33\s?|\b0)([1-9](?:[\s\.-]?\d{2}){4})")


def preprocess_image(img: Image.Image) -> Image.Image:
    # simple & robust preprocessing for business cards
    img = ImageOps.exif_transpose(img)
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = ImageEnhance.Sharpness(img).enhance(1.8)
    # upscale for OCR
    w, h = img.size
    img = img.resize((int(w * 1.8), int(h * 1.8)))
    return img


def ocr_text(image_path: Path) -> str:
    img = Image.open(image_path)
    img = preprocess_image(img)
    # French + English helps names/emails
    return pytesseract.image_to_string(img, lang="fra+eng")


def normalize_fr_phone(raw: str) -> str:
    s = re.sub(r"[^\d+]", "", raw)
    s = s.replace("+33", "0")
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 10:
        digits = digits[:10]
    if len(digits) == 10 and digits.startswith("0"):
        return digits
    return ""


def extract_email(text: str) -> str:
    m = EMAIL_RE.findall(text or "")
    if not m:
        return ""
    # remove obvious placeholders
    for e in m:
        if not re.search(r"(example|test|yourmail)", e, re.I):
            return e.strip()
    return m[0].strip()


def extract_phones(text: str) -> List[str]:
    out = []
    for m in PHONE_RE.finditer(text or ""):
        full = (m.group(0) or "").strip()
        p = normalize_fr_phone(full)
        if p and p not in out:
            out.append(p)
    return out


def looks_like_person(line: str) -> bool:
    line = clean(line)
    if not line or len(line) < 5:
        return False
    bad_words = [
        "sas", "sarl", "eurl", "veolia", "lely", "pissard", "ras", "intérim", "interim",
        "directeur", "directrice", "agence", "service", "www", "http", "@"
    ]
    if any(w in line.lower() for w in bad_words):
        return False

    # heuristic: 2 words, at least one in ALLCAPS or Title
    words = line.split()
    if len(words) < 2:
        return False
    # reject lines mostly numeric
    if sum(ch.isdigit() for ch in line) > 2:
        return False
    return True


def extract_person_name(text: str) -> Tuple[str, str, str]:
    """
    Returns (civility, firstname, lastname)
    Best-effort: detect "Céline MEKHMOUKH" or "MEKHMOUKH Céline"
    """
    lines = [clean(x) for x in (text or "").splitlines() if clean(x)]
    # prioritize top lines
    candidates = lines[:12] + lines[12:25]

    for ln in candidates:
        if not looks_like_person(ln):
            continue
        # remove job title after name
        ln = re.sub(r"(Directeur|Directrice|Responsable|Manager|Chef|Chargé|Charge)\b.*", "", ln, flags=re.I).strip()
        parts = ln.split()
        if len(parts) < 2:
            continue

        # if last name is uppercase (common on cards)
        if parts[-1].isupper() and len(parts[-1]) >= 3:
            firstname = parts[0]
            lastname = " ".join(parts[1:])
            return ("", firstname, lastname)

        # if first token uppercase and next token Title
        if parts[0].isupper() and len(parts[0]) >= 3:
            lastname = parts[0]
            firstname = " ".join(parts[1:])
            return ("", firstname, lastname)

        # fallback: first=first token, last=rest
        return ("", parts[0], " ".join(parts[1:]))

    return ("", "", "")


def enrich_from_card_ocr(prospect: Prospect, image_path: Path) -> Prospect:
    try:
        txt = ocr_text(image_path)
    except Exception:
        return prospect

    # contact name
    civ, fn, ln = extract_person_name(txt)
    if not prospect.contact_firstname and fn:
        prospect.contact_firstname = fn
    if not prospect.contact_lastname and ln:
        prospect.contact_lastname = ln
    if not prospect.contact_civility and civ:
        prospect.contact_civility = civ

    # email
    if not prospect.email:
        em = extract_email(txt)
        if em:
            prospect.email = em

    # phones (prefer mobile 06/07 in phone2 if empty)
    phones = extract_phones(txt)
    if phones:
        mob = next((p for p in phones if p.startswith("06") or p.startswith("07")), "")
        fix = next((p for p in phones if p.startswith("04") or p.startswith("01") or p.startswith("02") or p.startswith("03") or p.startswith("05") or p.startswith("09")), "")
        if not prospect.phone2 and mob:
            prospect.phone2 = mob
        if not prospect.phone and fix:
            prospect.phone = fix

    return prospect


# =========================
# Excel export
# =========================
COLUMNS = [
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


def write_excel(prospects: List[Prospect], out_xlsx: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PROSPECTION"

    ws.append(COLUMNS)

    for p in prospects:
        ws.append([
            p.name,
            p.address,
            p.postal_code,
            p.city,
            p.phone,
            p.phone2,
            p.email,
            p.siret,
            p.naf,
            p.website,
            p.contact_civility,
            p.contact_firstname,
            p.contact_lastname,
            p.dirigeant,
            p.resume,
            p.commande,
            p.infos_commerciales,
            p.card_file_id,
        ])

    # basic sizing
    for col in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 22

    wb.save(out_xlsx)


# =========================
# Email send (agency-specific)
# =========================
def get_recipients_for_agency(agency: str) -> List[str]:
    if not AGENCY_RECIPIENTS_JSON:
        return []
    try:
        m = json.loads(AGENCY_RECIPIENTS_JSON)
        val = m.get(agency) or m.get(agency.upper()) or []
        if isinstance(val, str):
            return [val]
        if isinstance(val, list):
            return [str(x) for x in val if str(x).strip()]
    except Exception:
        return []
    return []


def attach_file(msg: EmailMessage, path: Path) -> None:
    ctype, encoding = mimetypes.guess_type(str(path))
    if ctype is None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    msg.add_attachment(path.read_bytes(), maintype=maintype, subtype=subtype, filename=path.name)


def send_email(subject: str, body: str, recipients: List[str], attachments: List[Path]) -> None:
    if not recipients:
        raise RuntimeError(f"No recipients for agency {AGENCY} (check AGENCY_RECIPIENTS_JSON)")
    must(SMTP_HOST, "SMTP_HOST")
    must(MAIL_FROM, "MAIL_FROM")

    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    for p in attachments:
        if p.exists():
            attach_file(msg, p)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        if SMTP_USER and SMTP_PASS:
            s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)


# =========================
# Main
# =========================
def main() -> None:
    must(RUN_DATE, "RUN_DATE")
    must(AGENCY, "AGENCY")
    must(INITIALS, "INITIALS")

    raw = fetch_dump_jsonl(RUN_DATE)
    prospects_all = [to_prospect(x) for x in raw]

    # Filter by agency (critical!)
    prospects = [p for p in prospects_all if p.agency == AGENCY]

    # Download cards + OCR enrich
    card_dir = OUT_DIR / f"cards_{RUN_DATE}_{AGENCY}"
    card_dir.mkdir(parents=True, exist_ok=True)

    attachments: List[Path] = []

    for i, p in enumerate(prospects, start=1):
        if not p.card_file_id:
            continue

        img_path = card_dir / f"{i:03d}_{p.name[:30].replace(' ', '_')}.jpg"
        ok = telegram_download_file(p.card_file_id, img_path)
        if ok:
            # enrich fields from OCR
            p = enrich_from_card_ocr(p, img_path)
            # keep file for mail
            attachments.append(img_path)

    # Write excel
    xlsx_path = OUT_DIR / f"prospection_{RUN_DATE}_{AGENCY}.xlsx"
    write_excel(prospects, xlsx_path)
    attachments.insert(0, xlsx_path)

    # Compose email
    recipients = get_recipients_for_agency(AGENCY)

    subject = f"[PROSPECTION] {AGENCY} — {RUN_DATE} — {INITIALS}"
    body = (
        f"Export prospection — Agence {AGENCY}\n"
        f"Date: {RUN_DATE}\n"
        f"Initiales: {INITIALS}\n"
        f"Visites clients: {VISITS_CLIENTS or '-'}\n"
        f"Visites prospects: {VISITS_PROSPECTS or '-'}\n"
        f"Commandes: {COMMANDES or '-'}\n\n"
        f"PJ: Excel + cartes de visite (si présentes)\n"
        f"Nb entrées agence {AGENCY}: {len(prospects)}\n"
    )

    send_email(subject, body, recipients, attachments)

    # small debug file for artifacts
    debug = {
        "run_date": RUN_DATE,
        "agency": AGENCY,
        "initials": INITIALS,
        "count": len(prospects),
        "recipients": recipients,
        "attachments": [p.name for p in attachments],
    }
    (OUT_DIR / "debug.json").write_text(json.dumps(debug, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()