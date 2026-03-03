import os
import re
import json
import io
import base64
import zipfile
import datetime as dt
from dataclasses import dataclass
from typing import Dict, List, Any, Optional, Tuple
import requests

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment
from PIL import Image
import pytesseract


# ============================================================
# ENV / CONFIG
# ============================================================

def env_str(name: str, default: str = "") -> str:
    v = os.environ.get(name, default)
    if v is None:
        return default
    return str(v).strip()

def env_int(name: str, default: int) -> int:
    s = env_str(name, "")
    if s == "":
        return default
    try:
        return int(s)
    except Exception:
        return default

def today_paris_ymd() -> str:
    # GitHub runner UTC: on utilise RUN_DATE fourni par workflow sinon UTC date.
    # Le workflow V2 fixe déjà RUN_DATE Europe/Paris.
    return dt.date.today().strftime("%Y-%m-%d")

WORKER_BASE_URL    = env_str("WORKER_BASE_URL")
EXPORT_TOKEN       = env_str("EXPORT_TOKEN")
TELEGRAM_TOKEN     = env_str("TELEGRAM_TOKEN")

BREVO_API_KEY      = env_str("BREVO_API_KEY")
BREVO_SENDER_EMAIL = env_str("BREVO_SENDER_EMAIL", "no-reply@example.com")
BREVO_SENDER_NAME  = env_str("BREVO_SENDER_NAME",  "Prospection Bot")

MAIL_ROUTING_JSON  = env_str("MAIL_ROUTING_JSON", "")  # override possible

GOOGLE_PLACES_API_KEY = env_str("GOOGLE_PLACES_API_KEY", "")
GEMINI_API_KEY        = env_str("GEMINI_API_KEY", "")

SEND_MODE    = env_str("SEND_MODE", "individual")  # individual | agency_manager | admin
RUN_DATE     = env_str("RUN_DATE", today_paris_ymd())
AGENCY       = env_str("AGENCY", "").upper()       # GR|VR|GRS|SLS
INITIALS     = env_str("INITIALS", "").upper()     # JL etc.

MAX_OCR_IMAGES   = env_int("MAX_OCR_IMAGES", 50)    # safe
MAX_PHOTO_IMAGES = env_int("MAX_PHOTO_IMAGES", 15)  # safe

OUT_DIR = env_str("OUT_DIR", ".")  # GitHub artifacts will pickup *.xlsx
MEDIA_DIR = os.path.join(OUT_DIR, "media")
os.makedirs(MEDIA_DIR, exist_ok=True)

VALID_AGENCIES = {"GR", "VR", "GRS", "SLS"}


# ============================================================
# ROUTING (fallback intégré + override secret JSON)
# ============================================================

DEFAULT_ROUTING = {
    "agencies": {
        "GR": {
            "manager": {"initials": "JL", "email": "jennifer.laurens@ras-interim.fr"},
            "commercial": {"initials": "CZ", "email": "celine.zunarelli@ras-interim.fr"},
        },
        "VR": {
            "manager": {"initials": "JB", "email": "jelena.carrasso@ras-interim.fr"},
            "commercial": {"initials": "LB", "email": "laura.berthet@ras-interim.fr"},
        },
        "GRS": {
            "manager": {"initials": "PV", "email": "pauline.vieira@ras-interim.fr"},
            "commercial": {"initials": "ST", "email": "severine.thevenin@ras-interim.fr"},
        },
        "SLS": {
            "manager": {"initials": "AC", "email": "aurelie.curt@ras-interim.fr"},
            "commercial": {"initials": "AC", "email": "aurelie.curt@ras-interim.fr"},
        },
    },
    "admin": {"initials": "SL", "email": "samuel.lengue@ras-interim.fr"},
    # mapping initials->email (facilite individual/media)
    "users": {
        "JL": "jennifer.laurens@ras-interim.fr",
        "CZ": "celine.zunarelli@ras-interim.fr",
        "JB": "jelena.carrasso@ras-interim.fr",
        "LB": "laura.berthet@ras-interim.fr",
        "PV": "pauline.vieira@ras-interim.fr",
        "ST": "severine.thevenin@ras-interim.fr",
        "AC": "aurelie.curt@ras-interim.fr",
        "SL": "samuel.lengue@ras-interim.fr",
    }
}

def load_routing() -> Dict[str, Any]:
    if MAIL_ROUTING_JSON:
        try:
            return json.loads(MAIL_ROUTING_JSON)
        except Exception:
            # fallback
            return DEFAULT_ROUTING
    return DEFAULT_ROUTING

ROUTING = load_routing()


def email_for_initials(initials: str) -> Optional[str]:
    initials = (initials or "").upper().strip()
    if not initials:
        return None
    # 1) routing users
    users = ROUTING.get("users") or {}
    if initials in users:
        return users[initials]
    # 2) scan agencies roles
    agencies = ROUTING.get("agencies") or {}
    for ag, cfg in agencies.items():
        for role in ("manager", "commercial"):
            r = (cfg or {}).get(role) or {}
            if (r.get("initials") or "").upper() == initials and r.get("email"):
                return r["email"]
    return None


# ============================================================
# HTTP HELPERS
# ============================================================

def worker_dump(kind: str, date: str) -> List[Dict[str, Any]]:
    """
    /dump returns JSON lines, one record per line.
    """
    assert WORKER_BASE_URL and EXPORT_TOKEN
    url = f"{WORKER_BASE_URL.rstrip('/')}/dump?date={date}&kind={kind}"
    r = requests.get(url, headers={"X-Export-Token": EXPORT_TOKEN}, timeout=20)
    if r.status_code != 200:
        raise RuntimeError(f"/dump failed {kind} {r.status_code}: {r.text[:200]}")
    lines = [ln.strip() for ln in r.text.splitlines() if ln.strip()]
    out = []
    for ln in lines:
        try:
            out.append(json.loads(ln))
        except Exception:
            # ignore bad line
            pass
    return out


def tg_get_file_url(file_id: str) -> Optional[str]:
    if not file_id:
        return None
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=15)
    j = r.json()
    if not j.get("ok"):
        return None
    fp = j["result"]["file_path"]
    return f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{fp}"


def download_bytes(url: str) -> bytes:
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content


# ============================================================
# PARSING / NORMALISATION
# ============================================================

PHONE_RE = re.compile(r"(\+33|0)\s*[1-9](?:[\s\.-]*\d{2}){4}")
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")

def normalize_phone(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    s2 = re.sub(r"[^\d+]", "", s)
    # keep +33 if present
    return s2

def best_phone(text: str) -> str:
    if not text:
        return ""
    m = PHONE_RE.search(text)
    return normalize_phone(m.group(0)) if m else ""

def best_email(text: str) -> str:
    if not text:
        return ""
    m = EMAIL_RE.search(text)
    return (m.group(0).strip().lower()) if m else ""


# ============================================================
# OPTIONAL: PHONE RETRY (Places) for missing phone
# ============================================================

def places_retry_phone(name: str, city: str) -> str:
    """
    Worker already tries Places; here we retry in Python if phone is empty.
    """
    if not GOOGLE_PLACES_API_KEY or not name:
        return ""

    try:
        # searchText
        r1 = requests.post(
            "https://places.googleapis.com/v1/places:searchText",
            headers={
                "Content-Type": "application/json",
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "places.id",
            },
            json={
                "textQuery": f"{name} {city}".strip(),
                "maxResultCount": 1,
                "languageCode": "fr",
                "regionCode": "FR",
            },
            timeout=15
        )
        j1 = r1.json()
        p = (j1.get("places") or [None])[0]
        if not p or not p.get("id"):
            return ""
        place_id = p["id"]

        r2 = requests.get(
            f"https://places.googleapis.com/v1/places/{place_id}",
            headers={
                "X-Goog-Api-Key": GOOGLE_PLACES_API_KEY,
                "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber",
            },
            timeout=15
        )
        j2 = r2.json()
        ph = j2.get("nationalPhoneNumber") or j2.get("internationalPhoneNumber") or ""
        return normalize_phone(ph)
    except Exception:
        return ""


# ============================================================
# OCR + (optional) GEMINI
# ============================================================

def ocr_image_bytes(img_bytes: bytes) -> str:
    try:
        im = Image.open(io.BytesIO(img_bytes))
        # simple cleanup: convert to RGB
        if im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        txt = pytesseract.image_to_string(im, lang="fra+eng")
        return (txt or "").strip()
    except Exception:
        return ""

def gemini_extract_card(text: str) -> Dict[str, str]:
    """
    Minimal Gemini call (optional).
    If GEMINI_API_KEY missing => fallback regex parsing.
    """
    # fallback
    out = {
        "email": best_email(text),
        "phone": best_phone(text),
        "name": "",
        "company": "",
        "title": ""
    }
    if not GEMINI_API_KEY or not text:
        return out

    # Gemini REST (light) – best effort. If fails, keep fallback.
    try:
        # Google Generative Language API endpoint may vary by model availability.
        # We keep it resilient: if request fails => fallback.
        endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
        prompt = (
            "Tu es un extracteur de carte de visite.\n"
            "A partir du texte OCR ci-dessous, retourne STRICTEMENT un JSON avec les clés:\n"
            "name, company, title, email, phone\n"
            "phone doit être au format FR si possible.\n"
            "Si inconnu, mets une chaîne vide.\n\n"
            f"TEXTE OCR:\n{text}\n"
        )
        r = requests.post(
            endpoint,
            params={"key": GEMINI_API_KEY},
            json={
                "contents": [{"parts": [{"text": prompt}]}],
                "generationConfig": {"temperature": 0.1, "maxOutputTokens": 300},
            },
            timeout=20
        )
        j = r.json()
        cand = (((j.get("candidates") or [None])[0] or {}).get("content") or {}).get("parts") or []
        raw = ""
        for p in cand:
            if isinstance(p, dict) and p.get("text"):
                raw += p["text"]

        raw = raw.strip()
        # extract JSON block
        m = re.search(r"\{.*\}", raw, re.S)
        if not m:
            return out
        data = json.loads(m.group(0))
        out["name"] = (data.get("name") or "").strip()
        out["company"] = (data.get("company") or "").strip()
        out["title"] = (data.get("title") or "").strip()
        out["email"] = (data.get("email") or out["email"] or "").strip().lower()
        out["phone"] = normalize_phone(data.get("phone") or out["phone"] or "")
        return out
    except Exception:
        return out


# ============================================================
# EXCEL BUILDERS
# ============================================================

def autosize(ws):
    for col in range(1, ws.max_column + 1):
        max_len = 10
        for row in range(1, ws.max_row + 1):
            v = ws.cell(row=row, column=col).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v))[:80] if hasattr(len(str(v)), "__call__") else 10)
            max_len = min(max_len, 60)
        ws.column_dimensions[get_column_letter(col)].width = max_len

def make_wb() -> Workbook:
    wb = Workbook()
    # remove default sheet
    wb.remove(wb.active)
    return wb

def add_sheet_table(wb: Workbook, title: str, headers: List[str], rows: List[List[Any]]):
    ws = wb.create_sheet(title=title[:31])
    ws.append(headers)
    for r in rows:
        ws.append(r)

    # style header
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"
    autosize(ws)
    return ws


# ============================================================
# MEDIA ZIP
# ============================================================

def build_media_zip(date: str, agency: str, initials: str,
                    photos: List[Dict[str, Any]], cards: List[Dict[str, Any]]) -> Optional[Tuple[str, bytes]]:
    """
    Construit un zip contenant toutes les images (photos lieux + cartes)
    + un fichier index.csv.
    """
    # filter to user
    photos_u = [p for p in photos if (p.get("agency") == agency and (p.get("user") or "").upper() == initials)]
    cards_u  = [c for c in cards  if (c.get("agency") == agency and (c.get("user") or "").upper() == initials)]

    if not photos_u and not cards_u:
        return None

    # limits
    photos_u = photos_u[:MAX_PHOTO_IMAGES]
    cards_u  = cards_u[:MAX_OCR_IMAGES]

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        # index
        lines = ["type;file;date;agency;initials;city;comment;geo_lat;geo_lon"]
        # photos
        for i, p in enumerate(photos_u, start=1):
            fid = p.get("file_id") or ""
            url = tg_get_file_url(fid)
            if not url:
                continue
            img = download_bytes(url)
            fname = f"photos/{date}_{agency}_{initials}_{i:02d}.jpg"
            z.writestr(fname, img)

            geo = p.get("geo") or {}
            lines.append(
                f"photo;{fname};{date};{agency};{initials};{(p.get('city') or '')};"
                f"{(p.get('comment') or p.get('meeting') or '')};"
                f"{geo.get('lat') or ''};{geo.get('lon') or ''}"
            )

        # cards
        for i, c in enumerate(cards_u, start=1):
            fid = c.get("file_id") or ""
            url = tg_get_file_url(fid)
            if not url:
                continue
            img = download_bytes(url)
            fname = f"cards/{date}_{agency}_{initials}_{i:02d}.jpg"
            z.writestr(fname, img)
            lines.append(
                f"card;{fname};{date};{agency};{initials};;{(c.get('comment') or '')};;"
            )

        z.writestr("index.csv", ("\n".join(lines)).encode("utf-8"))

    zip_name = f"MEDIA_{date}_{agency}_{initials}.zip"
    return zip_name, buf.getvalue()


# ============================================================
# BREVO EMAIL
# ============================================================

def brevo_send_email(to_email: str, subject: str, html: str, attachments: Optional[List[Tuple[str, bytes]]] = None):
    if not BREVO_API_KEY:
        raise RuntimeError("BREVO_API_KEY missing")

    payload = {
        "sender": {"name": BREVO_SENDER_NAME, "email": BREVO_SENDER_EMAIL},
        "to": [{"email": to_email}],
        "subject": subject,
        "htmlContent": html,
    }

    if attachments:
        payload["attachment"] = []
        for filename, content in attachments:
            payload["attachment"].append({
                "name": filename,
                "content": base64.b64encode(content).decode("ascii")
            })

    r = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={
            "accept": "application/json",
            "api-key": BREVO_API_KEY,
            "content-type": "application/json"
        },
        data=json.dumps(payload),
        timeout=30
    )

    if r.status_code >= 300:
        raise RuntimeError(f"Brevo send failed {r.status_code}: {r.text[:300]}")


# ============================================================
# DATA PREP
# ============================================================

def split_by_agency(records: List[Dict[str, Any]]) -> Dict[str, List[Dict[str, Any]]]:
    out: Dict[str, List[Dict[str, Any]]] = {}
    for r in records:
        ag = (r.get("agency") or "").upper()
        if ag not in VALID_AGENCIES:
            continue
        out.setdefault(ag, []).append(r)
    return out

def uniq_users(records: List[Dict[str, Any]]) -> List[Tuple[str, str]]:
    """
    Return unique (agency, initials) combos present in records.
    """
    seen = set()
    out = []
    for r in records:
        ag = (r.get("agency") or "").upper()
        ini = (r.get("initials") or r.get("user") or "").upper()
        if ag in VALID_AGENCIES and ini:
            k = (ag, ini)
            if k not in seen:
                seen.add(k)
                out.append(k)
    return out


# ============================================================
# EXCEL CONTENT
# ============================================================

PROSPECT_HEADERS = [
    "date", "agency", "initials",
    "name", "address", "postal_code", "city",
    "siret", "naf", "dirigeant",
    "interlocuteur", "contact_firstname", "contact_lastname",
    "phone", "phone2", "email", "website",
    "resume", "commande",
]

PHOTO_HEADERS = [
    "date", "agency", "user",
    "city", "comment",
    "geo_lat", "geo_lon",
    "file_id",
]

CARD_HEADERS = [
    "date", "agency", "user",
    "comment",
    "ocr_email", "ocr_phone", "ocr_name", "ocr_company", "ocr_title",
    "file_id",
]

CLOSE_HEADERS = [
    "date", "agency", "initials",
    "visits_clients", "visits_prospects", "commandes",
    "declaratif", "closed_at"
]


def rows_prospects(prospects: List[Dict[str, Any]]) -> List[List[Any]]:
    rows = []
    for p in prospects:
        # retry phone if missing (optional)
        phone = (p.get("phone") or "").strip()
        if not phone:
            phone = places_retry_phone(p.get("name") or p.get("nom") or "", p.get("city") or "")
        row = [
            p.get("date", ""),
            (p.get("agency") or "").upper(),
            (p.get("initials") or "").upper(),
            p.get("name", ""),
            p.get("address", ""),
            p.get("postal_code", ""),
            p.get("city", ""),
            p.get("siret", ""),
            p.get("naf", ""),
            p.get("dirigeant", ""),
            p.get("interlocuteur", ""),
            p.get("contact_firstname", ""),
            p.get("contact_lastname", ""),
            phone,
            p.get("phone2", ""),
            p.get("email", ""),
            p.get("website", ""),
            p.get("resume", ""),
            p.get("commande", ""),
        ]
        rows.append(row)
    return rows


def rows_photos(photos: List[Dict[str, Any]]) -> List[List[Any]]:
    rows = []
    for ph in photos:
        geo = ph.get("geo") or {}
        rows.append([
            ph.get("date", ""),
            (ph.get("agency") or "").upper(),
            (ph.get("user") or "").upper(),
            ph.get("city", ""),
            ph.get("comment") or ph.get("meeting") or "",
            geo.get("lat", ""),
            geo.get("lon", ""),
            ph.get("file_id", ""),
        ])
    return rows


def rows_closes(closes: List[Dict[str, Any]]) -> List[List[Any]]:
    rows = []
    for c in closes:
        rows.append([
            c.get("date", ""),
            (c.get("agency") or "").upper(),
            (c.get("initials") or "").upper(),
            c.get("visits_clients", ""),
            c.get("visits_prospects", ""),
            c.get("commandes", ""),
            c.get("declaratif", True),
            c.get("closed_at", ""),
        ])
    return rows


def rows_cards_with_ocr(cards: List[Dict[str, Any]]) -> List[List[Any]]:
    """
    OCR each card and extract fields.
    """
    rows = []
    for i, c in enumerate(cards[:MAX_OCR_IMAGES], start=1):
        fid = c.get("file_id") or ""
        url = tg_get_file_url(fid)
        ocr_txt = ""
        extracted = {"email": "", "phone": "", "name": "", "company": "", "title": ""}

        if url:
            try:
                img = download_bytes(url)
                ocr_txt = ocr_image_bytes(img)
                extracted = gemini_extract_card(ocr_txt)
            except Exception:
                pass

        rows.append([
            c.get("date", ""),
            (c.get("agency") or "").upper(),
            (c.get("user") or "").upper(),
            c.get("comment", "") or "",
            extracted.get("email", ""),
            extracted.get("phone", ""),
            extracted.get("name", ""),
            extracted.get("company", ""),
            extracted.get("title", ""),
            fid,
        ])
    return rows


# ============================================================
# BUILD FILES
# ============================================================

def build_excel(date: str,
                prospects: List[Dict[str, Any]],
                closes: List[Dict[str, Any]],
                photos: List[Dict[str, Any]],
                cards: List[Dict[str, Any]],
                title_suffix: str) -> str:
    wb = make_wb()

    add_sheet_table(wb, "PROSPECTS", PROSPECT_HEADERS, rows_prospects(prospects))
    add_sheet_table(wb, "CLOSES", CLOSE_HEADERS, rows_closes(closes))

    # MEDIA sheets (metadata)
    add_sheet_table(wb, "PHOTOS", PHOTO_HEADERS, rows_photos(photos))
    add_sheet_table(wb, "CARDS", CARD_HEADERS, rows_cards_with_ocr(cards))

    filename = os.path.join(OUT_DIR, f"PROSPECTION_{date}_{title_suffix}.xlsx")
    wb.save(filename)
    return filename


# ============================================================
# SEND LOGIC
# ============================================================

def send_individual_pack(date: str, agency: str, initials: str,
                         prospects: List[Dict[str, Any]],
                         closes: List[Dict[str, Any]],
                         photos: List[Dict[str, Any]],
                         cards: List[Dict[str, Any]]):
    to_email = email_for_initials(initials)
    if not to_email:
        print(f"[WARN] No email for initials={initials}, skip individual pack.")
        return

    # Filter to this user+agency
    p_u = [p for p in prospects if (p.get("agency") == agency and (p.get("initials") or "").upper() == initials)]
    c_u = [c for c in closes    if (c.get("agency") == agency and (c.get("initials") or "").upper() == initials)]
    ph_u = [ph for ph in photos if (ph.get("agency") == agency and (ph.get("user") or "").upper() == initials)]
    ca_u = [ca for ca in cards  if (ca.get("agency") == agency and (ca.get("user") or "").upper() == initials)]

    xlsx = build_excel(date, p_u, c_u, ph_u, ca_u, f"INDIV_{agency}_{initials}")

    attachments = []
    with open(xlsx, "rb") as f:
        attachments.append((os.path.basename(xlsx), f.read()))

    media = build_media_zip(date, agency, initials, photos, cards)
    if media:
        attachments.append(media)

    subject = f"Prospection {date} — {agency}/{initials} (Excel + médias)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici ton export de prospection du <b>{date}</b> pour <b>{agency}/{initials}</b>.</p>"
        f"<ul>"
        f"<li>Excel récapitulatif</li>"
        f"<li>Médias (zip) : photos lieux + cartes de visite (si présents)</li>"
        f"</ul>"
        f"<p>— Bot Prospection</p>"
    )
    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] individual pack sent to {to_email} ({agency}/{initials})")


def send_agency_manager_pack(date: str, agency: str,
                             prospects: List[Dict[str, Any]],
                             closes: List[Dict[str, Any]]):
    agencies_cfg = ROUTING.get("agencies") or {}
    cfg = agencies_cfg.get(agency) or {}
    manager = (cfg.get("manager") or {})
    to_email = manager.get("email")
    if not to_email:
        print(f"[WARN] No manager email for agency={agency}")
        return

    # Agency consolidated (WITHOUT media attachments)
    p_ag = [p for p in prospects if (p.get("agency") == agency)]
    c_ag = [c for c in closes    if (c.get("agency") == agency)]

    # For managers, we keep media sheets empty to match your current usage
    xlsx = build_excel(date, p_ag, c_ag, [], [], f"AGENCE_{agency}")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    subject = f"Prospection {date} — Agence {agency} (consolidé)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici le consolidé prospection du <b>{date}</b> pour l’agence <b>{agency}</b>.</p>"
        f"<p>(Médias non joints au consolidé agence — envoyés individuellement au preneur.)</p>"
        f"<p>— Bot Prospection</p>"
    )
    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] agency manager pack sent to {to_email} (agency={agency})")


def send_admin_pack(date: str, prospects: List[Dict[str, Any]], closes: List[Dict[str, Any]]):
    admin = ROUTING.get("admin") or {}
    to_email = admin.get("email")
    if not to_email:
        print("[WARN] No admin email configured")
        return

    xlsx = build_excel(date, prospects, closes, [], [], "ADMIN_ALL")

    with open(xlsx, "rb") as f:
        attachments = [(os.path.basename(xlsx), f.read())]

    subject = f"Prospection {date} — ADMIN (toutes agences)"
    html = (
        f"<p>Bonjour,</p>"
        f"<p>Voici le consolidé global du <b>{date}</b> (toutes agences / tous collaborateurs).</p>"
        f"<p>(Médias envoyés individuellement à chaque preneur, y compris SL lorsqu’il aide une agence.)</p>"
        f"<p>— Bot Prospection</p>"
    )
    brevo_send_email(to_email, subject, html, attachments=attachments)
    print(f"[OK] admin pack sent to {to_email}")


# ============================================================
# MAIN
# ============================================================

def main():
    if not WORKER_BASE_URL or not EXPORT_TOKEN or not TELEGRAM_TOKEN:
        raise RuntimeError("Missing WORKER_BASE_URL / EXPORT_TOKEN / TELEGRAM_TOKEN")

    print(f"🚀 export_and_mail.py — mode={SEND_MODE} date={RUN_DATE} agency={AGENCY} initials={INITIALS}")

    prospects = worker_dump("prospects", RUN_DATE)
    closes    = worker_dump("closes",    RUN_DATE)
    photos    = worker_dump("photos",    RUN_DATE)
    cards     = worker_dump("cards",     RUN_DATE)  # NEW V6

    # ========================================================
    # MODE: individual
    # ========================================================
    if SEND_MODE == "individual":
        if AGENCY not in VALID_AGENCIES:
            raise RuntimeError("AGENCY required (GR|VR|GRS|SLS) for mode=individual")
        if not INITIALS:
            raise RuntimeError("INITIALS required for mode=individual")

        send_individual_pack(RUN_DATE, AGENCY, INITIALS, prospects, closes, photos, cards)
        return

    # ========================================================
    # MODE: agency_manager (scheduled 17:45)
    # - send consolidé agence aux managers
    # - send MEDIA packs à tous ceux qui ont pris des médias (photos/cards)
    # ========================================================
    if SEND_MODE == "agency_manager":
        # 1) managers consolidés par agence
        for ag in sorted(VALID_AGENCIES):
            send_agency_manager_pack(RUN_DATE, ag, prospects, closes)

        # 2) MEDIA packs (photos/cards) à chaque preneur (important pour SL quand il aide)
        media_users = set()
        for (ag, ini) in uniq_users(photos) + uniq_users(cards):
            media_users.add((ag, ini))

        for ag, ini in sorted(media_users):
            send_individual_pack(RUN_DATE, ag, ini, prospects, closes, photos, cards)

        return

    # ========================================================
    # MODE: admin (scheduled 17:47)
    # - send consolidé global à SL
    # - send MEDIA packs à tous les preneurs (optionnel, mais utile)
    # ========================================================
    if SEND_MODE == "admin":
        send_admin_pack(RUN_DATE, prospects, closes)

        media_users = set()
        for (ag, ini) in uniq_users(photos) + uniq_users(cards):
            media_users.add((ag, ini))

        for ag, ini in sorted(media_users):
            send_individual_pack(RUN_DATE, ag, ini, prospects, closes, photos, cards)

        return

    raise RuntimeError(f"Unknown SEND_MODE={SEND_MODE}")


if __name__ == "__main__":
    main()