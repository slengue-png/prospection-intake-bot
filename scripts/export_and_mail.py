#!/usr/bin/env python3
"""
export_and_mail.py — VERSION PRO V2 PHOTO + GEMINI VISION (safe add-on)

V1 conservée :
- /dump prospects + closes
- OCR Tesseract cartes de visite (mode entreprise)
- Export Excel depuis Table.2.xlsx (openpyxl)
- Emails Brevo (individual / agency_manager / admin)

V2 ajout :
- /dump photos
- Téléchargement images Telegram côté GitHub
- Gemini Vision (gemini-1.5-flash) optionnel (GEMINI_API_KEY)
- Enrichissement déterministe API gouv + (optionnel) Google Places
- Excel : ajoute SOURCE / CONFIDENCE / STATUS / SUGGESTIONS
- Ne jamais écraser données manuelles si présentes

ENV requis :
- WORKER_BASE_URL
- EXPORT_TOKEN
- TELEGRAM_TOKEN
- BREVO_API_KEY
- BREVO_SENDER_EMAIL
- BREVO_SENDER_NAME
- MAIL_ROUTING_JSON

ENV optionnels :
- GEMINI_API_KEY
- GOOGLE_PLACES_API_KEY   (si tu veux enrichir tel/site côté GitHub pour mode photo)

Inputs workflow :
- SEND_MODE: individual | agency_manager | admin
- RUN_DATE: YYYY-MM-DD (optionnel)
- AGENCY, INITIALS (requis si individual)
- MAX_OCR_IMAGES (défaut 50)
- MAX_PHOTO_IMAGES (défaut 15)
"""

import os
import re
import io
import json
import base64
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple, Any

import requests
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image
import pytesseract


# =========================
# REGEX
# =========================
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)\s*[1-9](?:[\s.\-]*\d{2}){4})")


# =========================
# ENV
# =========================
def must_env(name: str) -> str:
    v = os.getenv(name, "").strip()
    if not v:
        raise RuntimeError(f"Missing required env var: {name}")
    return v


def opt_env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def opt_int(name: str, default: int) -> int:
    s = os.getenv(name, "").strip()
    if not s:
        return default
    try:
        return int(s)
    except Exception:
        return default


# =========================
# GLOBAL CONFIG
# =========================
TEMPLATE_PATH = "Table.2.xlsx"

TEMPLATE_HEADERS = [
    "NOM",
    "RUE",
    "CODE POSTAL",
    "VILLE",
    "Téléphone",
    "Téléphone (Portable)",
    "Mail générique",
    "SIRET",
    "NAF",
    "SITE WEB",
    "INTERLOCUTEUR",
    "DIRIGEANT",
    "RESUME ENTRETIEN",
    "COMMANDE",
    "CARTE DE VISITE",
]

EXTRA_HEADERS = ["AGENCE", "INITIALS"]

PHOTO_HEADERS = ["SOURCE", "CONFIDENCE", "STATUS", "SUGGESTIONS"]

# Gemini config
GEMINI_MODEL = "gemini-1.5-flash"
GEMINI_TEMPERATURE = 0.2
GEMINI_MAX_TOKENS = 512
GEMINI_TIMEOUT_S = 25
GEMINI_MAX_RETRIES = 3


# =========================
# WORKER DUMP
# =========================
def build_dump_url(base: str, run_date: str, kind: str) -> str:
    base = base.rstrip("/")
    return f"{base}/dump?date={run_date}&kind={kind}"


def fetch_dump_jsonl(run_date: str, kind: str) -> List[dict]:
    base = must_env("WORKER_BASE_URL")
    token = must_env("EXPORT_TOKEN")

    url = build_dump_url(base, run_date, kind)
    r = requests.get(url, headers={"X-Export-Token": token}, timeout=60)
    r.raise_for_status()

    out: List[dict] = []
    for ln in r.text.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        try:
            out.append(json.loads(ln))
        except Exception:
            continue
    return out


# =========================
# TELEGRAM DOWNLOAD
# =========================
def telegram_get_file_path(file_id: str) -> Optional[str]:
    token = must_env("TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/bot{token}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=30)
    r.raise_for_status()
    j = r.json()
    if not j.get("ok"):
        return None
    return j["result"].get("file_path")


def telegram_download_file_bytes(file_id: str) -> Optional[bytes]:
    fp = telegram_get_file_path(file_id)
    if not fp:
        return None
    token = must_env("TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/file/bot{token}/{fp}"
    r = requests.get(url, timeout=60)
    if r.ok and r.content:
        return r.content
    return None


def download_card_image_bytes(record: dict) -> Optional[bytes]:
    url = (record.get("card_photo_url") or "").strip()
    if url:
        r = requests.get(url, timeout=60)
        if r.ok and r.content:
            return r.content

    file_id = (record.get("card_photo_file_id") or "").strip()
    if not file_id:
        return None

    return telegram_download_file_bytes(file_id)


# =========================
# OCR (V1 cards)
# =========================
def normalize_phone_fr(raw: str) -> str:
    p = re.sub(r"[^\d+]", "", raw or "")
    if p.startswith("+33"):
        p = "0" + p[3:]
    p = re.sub(r"\D", "", p)
    if len(p) == 10 and p.startswith("0"):
        return p
    return ""


def split_name_simple(full: str) -> Tuple[str, str]:
    s = (full or "").strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return ("", "")
    parts = s.split(" ")
    if len(parts) == 1:
        return ("", parts[0])
    return (parts[0], " ".join(parts[1:]))


def ocr_extract(image_bytes: bytes) -> Dict[str, str]:
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return {"ocr_text": "", "email": "", "mobile": "", "phone": "", "contact_full": ""}

    text = pytesseract.image_to_string(img, lang="fra+eng")
    text = re.sub(r"[ \t]+", " ", text)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    email = ""
    m = EMAIL_RE.search(text)
    if m:
        email = m.group(0).strip()

    phones = PHONE_RE.findall(text)
    normalized = []
    for p in phones:
        np = normalize_phone_fr(p)
        if np:
            normalized.append(np)

    mobile = next((x for x in normalized if x.startswith(("06", "07"))), "")
    phone = next((x for x in normalized if x.startswith("04")), "")
    if not phone:
        phone = next((x for x in normalized if x.startswith(("01", "02", "03", "05", "08", "09"))), "")

    contact_full = ""
    candidates = []
    for ln in lines[:25]:
        if len(ln) < 4 or len(ln) > 60:
            continue
        upper = ln.upper()
        if any(k in upper for k in ["SAS", "SARL", "EURL", "FRANCE", "GROUPE", "WWW", "HTTP", "@"]):
            continue
        if len(ln.split()) < 2:
            continue
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", ln):
            continue
        candidates.append(ln)

    for ln in candidates:
        if re.search(r"\b[A-ZÀ-ÖØ-Ý]{3,}\b", ln):
            contact_full = ln
            break
    if not contact_full and candidates:
        contact_full = candidates[0]

    return {"ocr_text": text, "email": email, "mobile": mobile, "phone": phone, "contact_full": contact_full}


def ensure_contact_fields(record: dict) -> None:
    interloc = (record.get("interlocuteur") or "").strip()
    if interloc:
        if not (record.get("contact_firstname") or "").strip() and not (record.get("contact_lastname") or "").strip():
            fn, ln = split_name_simple(interloc)
            record["contact_firstname"] = fn
            record["contact_lastname"] = ln or interloc
        return

    if (record.get("contact_firstname") or "").strip() or (record.get("contact_lastname") or "").strip():
        return

    dirg = (record.get("dirigeant") or "").strip()
    if dirg:
        fn, ln = split_name_simple(dirg)
        record["contact_firstname"] = fn
        record["contact_lastname"] = ln or dirg


# =========================
# Gemini helpers (strict JSON)
# =========================
def _gemini_enabled() -> bool:
    return bool(opt_env("GEMINI_API_KEY"))


def _extract_json_object(text: str) -> Optional[str]:
    if not text:
        return None
    s = text.strip()
    s = re.sub(r"^```(?:json)?\s*", "", s, flags=re.IGNORECASE).strip()
    s = re.sub(r"\s*```$", "", s).strip()

    start = s.find("{")
    if start < 0:
        return None

    depth = 0
    in_str = False
    esc = False
    for i in range(start, len(s)):
        ch = s[i]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        else:
            if ch == '"':
                in_str = True
                continue
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    return s[start : i + 1]
    return None


def gemini_generate_json_multimodal(system_instruction: str, user_text: str, image_bytes: bytes, mime_type: str, schema_hint: str) -> Optional[dict]:
    api_key = opt_env("GEMINI_API_KEY")
    if not api_key:
        return None

    url = f"https://generativelanguage.googleapis.com/v1beta/models/{GEMINI_MODEL}:generateContent"
    params = {"key": api_key}
    headers = {"Content-Type": "application/json"}

    b64 = base64.b64encode(image_bytes).decode("utf-8")

    payload = {
        "contents": [
            {
                "role": "user",
                "parts": [
                    {"text": "SYSTEM INSTRUCTION:\n" + (system_instruction or "").strip()},
                    {
                        "text": (
                            "\nOUTPUT RULES (STRICT):\n"
                            "- Return ONLY one JSON object.\n"
                            "- No markdown.\n"
                            "- No explanations.\n"
                            "- Use null for unknown.\n"
                            "- Never invent email or phone.\n"
                            "- confidence must be integer 0-100.\n\n"
                            "SCHEMA (HINT):\n"
                            f"{(schema_hint or '').strip()}\n\n"
                            "USER CONTEXT:\n"
                            f"{(user_text or '').strip()}\n"
                        )
                    },
                    {"inline_data": {"mime_type": mime_type, "data": b64}},
                ],
            }
        ],
        "generationConfig": {
            "temperature": GEMINI_TEMPERATURE,
            "maxOutputTokens": GEMINI_MAX_TOKENS,
        },
    }

    last_err = None
    for attempt in range(1, GEMINI_MAX_RETRIES + 1):
        try:
            r = requests.post(url, headers=headers, params=params, json=payload, timeout=GEMINI_TIMEOUT_S)
            if r.status_code in (429, 500, 502, 503, 504):
                last_err = f"HTTP {r.status_code}"
                time.sleep(0.8 * attempt)
                continue
            r.raise_for_status()
            j = r.json()

            text = ""
            candidates = j.get("candidates") or []
            if candidates:
                parts = candidates[0].get("content", {}).get("parts") or []
                if parts:
                    text = parts[0].get("text", "") or ""
            raw = text.strip()
            if not raw:
                return None

            obj_str = _extract_json_object(raw) or raw
            parsed = json.loads(obj_str)
            if isinstance(parsed, dict):
                return parsed
            return None
        except Exception as e:
            last_err = str(e)
            time.sleep(0.8 * attempt)
            continue

    _ = last_err
    return None


def ai_photo_extract(image_bytes: bytes, city: str) -> Optional[dict]:
    """
    Retour EXACTEMENT :
    {
      "kind": "business_card" | "logo" | "signage" | "unknown",
      "company_name_guess": null,
      "text_extracted": null,
      "contact": { ... },
      "confidence": 0
    }
    """
    schema = {
        "kind": "unknown",
        "company_name_guess": None,
        "text_extracted": None,
        "contact": {
            "civility": None,
            "first_name": None,
            "last_name": None,
            "job_title": None,
            "email": None,
            "mobile": None,
            "phone": None,
        },
        "confidence": 0,
    }
    system_instruction = (
        "You analyze an image that may show a business card, a company logo, or a storefront sign.\n"
        "Your job: classify kind, extract a company name guess if visible, and extract contact fields if it is a business card.\n"
        "Rules: never invent email or phone. Use null if missing."
    )
    user_text = f"Confirmed city: {city} (France)."
    schema_hint = json.dumps(schema, ensure_ascii=False)

    # mime guess (simple)
    mime = "image/jpeg"
    res = gemini_generate_json_multimodal(system_instruction, user_text, image_bytes, mime, schema_hint)
    if not res:
        return None

    # sanitize/normalize
    out = {
        "kind": res.get("kind") if res.get("kind") in ("business_card", "logo", "signage", "unknown") else "unknown",
        "company_name_guess": res.get("company_name_guess", None),
        "text_extracted": res.get("text_extracted", None),
        "contact": res.get("contact") if isinstance(res.get("contact"), dict) else {},
        "confidence": res.get("confidence", 0),
    }

    try:
        out["confidence"] = int(out["confidence"])
    except Exception:
        out["confidence"] = 0
    if out["confidence"] < 0:
        out["confidence"] = 0
    if out["confidence"] > 100:
        out["confidence"] = 100

    # contact sub-keys
    c = out["contact"] if isinstance(out["contact"], dict) else {}
    out["contact"] = {
        "civility": (c.get("civility") or None),
        "first_name": (c.get("first_name") or None),
        "last_name": (c.get("last_name") or None),
        "job_title": (c.get("job_title") or None),
        "email": (c.get("email") or None),
        "mobile": (c.get("mobile") or None),
        "phone": (c.get("phone") or None),
    }

    # strip strings
    for k in ("company_name_guess", "text_extracted"):
        if isinstance(out.get(k), str):
            out[k] = out[k].strip() or None
    for k, v in list(out["contact"].items()):
        if isinstance(v, str):
            vv = v.strip()
            out["contact"][k] = vv if vv else None

    return out


# =========================
# API GOUV SEARCH (deterministic enrich for photo)
# =========================
def gov_search_company(q: str, per_page: int = 3) -> List[dict]:
    url = "https://recherche-entreprises.api.gouv.fr/search"
    r = requests.get(url, params={"q": q, "page": 1, "per_page": per_page}, timeout=30)
    r.raise_for_status()
    j = r.json()
    return j.get("results") or []


def _safe_get_siege(result: dict) -> dict:
    s = result.get("siege") or {}
    return s if isinstance(s, dict) else {}


def _extract_name(result: dict) -> str:
    return (result.get("nom_raison_sociale") or result.get("denomination") or "").strip()


def _extract_naf(result: dict) -> str:
    naf = (result.get("activite_principale") or result.get("naf") or "").strip()
    return naf.replace(".", "").replace(" ", "")


def _extract_dirigeant(result: dict) -> str:
    arr = result.get("dirigeants") or result.get("representants") or []
    if not isinstance(arr, list) or not arr:
        return ""
    first = arr[0]
    if isinstance(first, str):
        return first.strip()
    p = first.get("personne") if isinstance(first, dict) else {}
    if not isinstance(p, dict):
        p = first if isinstance(first, dict) else {}
    prenom = (p.get("prenom") or p.get("prenoms") or "").strip()
    nom = (p.get("nom") or p.get("nom_usage") or p.get("nomNaissance") or "").strip()
    full = f"{prenom} {nom}".strip()
    return full


def _score_match(result: dict, name_guess: str, city: str) -> int:
    # scoring simple: name token overlap + city match
    name_guess_l = (name_guess or "").lower()
    city_l = (city or "").lower()
    r_name = _extract_name(result).lower()
    siege = _safe_get_siege(result)
    r_city = (siege.get("libelle_commune") or "").lower()

    score = 0
    if city_l and r_city == city_l:
        score += 30
    if city_l and city_l in r_city:
        score += 15

    # token overlap
    toks = [t for t in re.split(r"[^a-z0-9]+", name_guess_l) if t and len(t) >= 3]
    for t in toks[:6]:
        if t in r_name:
            score += 10

    # prefer longer siret present
    if (siege.get("siret") or ""):
        score += 5

    return score


def enrich_from_gov(name_guess: str, city: str) -> Tuple[Optional[dict], List[str]]:
    q = f"{name_guess} {city}".strip()
    results = gov_search_company(q, per_page=3)
    if not results:
        return None, []

    scored = [(r, _score_match(r, name_guess, city)) for r in results]
    scored.sort(key=lambda x: x[1], reverse=True)

    top = scored[0][0]
    suggestions = [_extract_name(x[0]) for x in scored[:3] if _extract_name(x[0])]

    siege = _safe_get_siege(top)

    payload = {
        "name": _extract_name(top),
        "siret": (siege.get("siret") or "").strip(),
        "naf": _extract_naf(top),
        "address": (siege.get("adresse") or siege.get("libelle_voie") or "").strip(),
        "postal_code": (siege.get("code_postal") or "").strip(),
        "city": (siege.get("libelle_commune") or city).strip(),
        "dirigeant": _extract_dirigeant(top),
        "siren": (top.get("siren") or (siege.get("siret") or "")[:9]).strip(),
    }
    return payload, suggestions


# =========================
# Google Places (optional, GitHub side)
# =========================
def google_places_enrich(name: str, city: str) -> Dict[str, str]:
    api_key = opt_env("GOOGLE_PLACES_API_KEY")
    if not api_key:
        return {"phone": "", "website": ""}

    # Text Search
    ts_url = "https://places.googleapis.com/v1/places:searchText"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": api_key,
        "X-Goog-FieldMask": "places.id,places.websiteUri",
    }
    body = {"textQuery": f"{name} {city}".strip(), "maxResultCount": 1, "languageCode": "fr", "regionCode": "FR"}
    r = requests.post(ts_url, headers=headers, json=body, timeout=25)
    if not r.ok:
        return {"phone": "", "website": ""}

    j = r.json()
    p = (j.get("places") or [None])[0]
    if not p:
        return {"phone": "", "website": ""}

    place_id = p.get("id") or ""
    website = p.get("websiteUri") or ""

    phone = ""
    if place_id:
        det_url = f"https://places.googleapis.com/v1/places/{place_id}"
        det_headers = {
            "X-Goog-Api-Key": api_key,
            "X-Goog-FieldMask": "nationalPhoneNumber,internationalPhoneNumber,websiteUri",
        }
        r2 = requests.get(det_url, headers=det_headers, timeout=25)
        if r2.ok:
            j2 = r2.json()
            phone = j2.get("nationalPhoneNumber") or j2.get("internationalPhoneNumber") or ""
            website = j2.get("websiteUri") or website

    return {"phone": phone or "", "website": website or ""}


# =========================
# Utility: never overwrite manual
# =========================
def fill_if_empty(record: dict, key: str, value: Any) -> None:
    if value is None:
        return
    cur = record.get(key, None)
    if cur is None:
        record[key] = value
        return
    if isinstance(cur, str) and not cur.strip():
        record[key] = value
        return


# =========================
# Excel
# =========================
def build_header_map(ws) -> Dict[str, int]:
    max_col = ws.max_column
    return {str(ws.cell(row=1, column=c).value or "").strip(): c for c in range(1, max_col + 1)}


def ensure_headers(ws, headers_to_ensure: List[str]) -> Dict[str, int]:
    header_map = build_header_map(ws)
    col = ws.max_column
    for h in headers_to_ensure:
        if h not in header_map:
            col += 1
            ws.cell(row=1, column=col).value = h
            header_map[h] = col
    return header_map


def autosize(ws, width: int = 22) -> None:
    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def build_excel_from_template(records: List[dict], out_path: str) -> None:
    if not os.path.exists(TEMPLATE_PATH):
        raise RuntimeError(f"Missing Excel template at repo root: {TEMPLATE_PATH}")

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    header_map = ensure_headers(ws, EXTRA_HEADERS)
    header_map = ensure_headers(ws, PHOTO_HEADERS)

    start_row = ws.max_row + 1

    for idx, r in enumerate(records):
        row = start_row + idx

        ensure_contact_fields(r)

        interloc_excel = (r.get("interlocuteur") or "").strip()
        if not interloc_excel:
            fn = (r.get("contact_firstname") or "").strip()
            ln = (r.get("contact_lastname") or "").strip()
            interloc_excel = f"{fn} {ln}".strip() or (r.get("dirigeant") or "")

        values = {
            "NOM": r.get("name", "") or "",
            "RUE": r.get("address", "") or "",
            "CODE POSTAL": r.get("postal_code", "") or "",
            "VILLE": r.get("city", "") or "",
            "Téléphone": r.get("phone", "") or "",
            "Téléphone (Portable)": r.get("phone2", "") or "",
            "Mail générique": r.get("email", "") or "",
            "SIRET": r.get("siret", "") or "",
            "NAF": (r.get("naf", "") or "").replace(".", "").replace(" ", ""),
            "SITE WEB": r.get("website", "") or "",
            "INTERLOCUTEUR": interloc_excel,
            "DIRIGEANT": r.get("dirigeant", "") or "",
            "RESUME ENTRETIEN": r.get("resume", "") or "",
            "COMMANDE": r.get("commande", "") or "",
            "CARTE DE VISITE": r.get("card_photo_url", "") or r.get("card_photo_file_id", "") or "",
            "AGENCE": r.get("agency", "") or "",
            "INITIALS": r.get("initials", "") or "",
            "SOURCE": r.get("source", "") or "",
            "CONFIDENCE": r.get("confidence", "") or "",
            "STATUS": r.get("status", "") or "",
            "SUGGESTIONS": r.get("suggestions", "") or "",
        }

        for h, v in values.items():
            col = header_map.get(h)
            if col:
                ws.cell(row=row, column=col).value = v

    autosize(ws, 22)
    wb.save(out_path)


# =========================
# Brevo
# =========================
def send_mail_brevo(to_list: List[str], subject: str, html: str, attachments: List[Tuple[str, bytes]]) -> None:
    api_key = must_env("BREVO_API_KEY")
    sender_email = must_env("BREVO_SENDER_EMAIL")
    sender_name = must_env("BREVO_SENDER_NAME")

    payload = {
        "sender": {"email": sender_email, "name": sender_name},
        "to": [{"email": x} for x in to_list],
        "subject": subject,
        "htmlContent": html,
        "attachment": [{"name": name, "content": base64.b64encode(b).decode("utf-8")} for name, b in attachments],
    }

    r = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={"api-key": api_key, "content-type": "application/json"},
        json=payload,
        timeout=60,
    )
    r.raise_for_status()


# =========================
# Stats (closes)
# =========================
def to_int(x) -> int:
    try:
        s = str(x).strip()
        return int(s) if s else 0
    except Exception:
        return 0


def compute_stats(prospects: List[dict], closes: List[dict]) -> dict:
    by_agency: Dict[str, dict] = {}
    by_initials: Dict[str, dict] = {}

    for p in prospects:
        ag = (p.get("agency") or "").upper()
        ini = (p.get("initials") or "").upper()
        if not ag:
            continue

        by_agency.setdefault(ag, {"actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
        by_agency[ag]["actions"] += 1

        by_initials.setdefault(ini, {"agencies": set(), "actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
        by_initials[ini]["actions"] += 1
        by_initials[ini]["agencies"].add(ag)

    for c in closes:
        ag = (c.get("agency") or "").upper()
        ini = (c.get("initials") or "").upper()

        clients = to_int(c.get("visits_clients"))
        prosps = to_int(c.get("visits_prospects"))
        comm = to_int(c.get("commandes"))

        if ag:
            by_agency.setdefault(ag, {"actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
            by_agency[ag]["clients"] += clients
            by_agency[ag]["prospects"] += prosps
            by_agency[ag]["commandes"] += comm

        by_initials.setdefault(ini, {"agencies": set(), "actions": 0, "clients": 0, "prospects": 0, "commandes": 0})
        by_initials[ini]["clients"] += clients
        by_initials[ini]["prospects"] += prosps
        by_initials[ini]["commandes"] += comm
        if ag:
            by_initials[ini]["agencies"].add(ag)

    totals = {
        "actions": sum(v["actions"] for v in by_agency.values()),
        "clients": sum(v["clients"] for v in by_agency.values()),
        "prospects": sum(v["prospects"] for v in by_agency.values()),
        "commandes": sum(v["commandes"] for v in by_agency.values()),
    }

    ranking = sorted([(ini, v["actions"]) for ini, v in by_initials.items() if ini], key=lambda x: x[1], reverse=True)
    for ini in list(by_initials.keys()):
        by_initials[ini]["agencies"] = sorted(list(by_initials[ini]["agencies"]))
    return {"by_agency": by_agency, "by_initials": by_initials, "totals": totals, "ranking": ranking}


def html_table(headers: List[str], rows: List[List[Any]]) -> str:
    th = "".join([f"<th style='border:1px solid #ddd;padding:6px;text-align:left'>{h}</th>" for h in headers])
    trs = []
    for r in rows:
        tds = "".join([f"<td style='border:1px solid #ddd;padding:6px'>{str(x)}</td>" for x in r])
        trs.append(f"<tr>{tds}</tr>")
    return f"<table style='border-collapse:collapse;border:1px solid #ddd;margin:10px 0'><tr>{th}</tr>{''.join(trs)}</table>"


# =========================
# Filters
# =========================
def filter_by_date(items: List[dict], run_date: str) -> List[dict]:
    return [x for x in items if str(x.get("date", "")).strip() == run_date]


def filter_prospects(items: List[dict], run_date: str, mode: str, agency: str, initials: str) -> List[dict]:
    items = filter_by_date(items, run_date)
    if mode == "individual":
        return [x for x in items if (x.get("agency") or "").upper() == agency and (x.get("initials") or "").upper() == initials]
    if mode == "agency_manager":
        return [x for x in items if (x.get("agency") or "").upper() == agency] if agency else items
    return items


def filter_closes(items: List[dict], run_date: str, mode: str, agency: str, initials: str) -> List[dict]:
    items = filter_by_date(items, run_date)
    if mode == "individual":
        return [x for x in items if (x.get("agency") or "").upper() == agency and (x.get("initials") or "").upper() == initials]
    if mode == "agency_manager":
        return [x for x in items if (x.get("agency") or "").upper() == agency] if agency else items
    return items


def filter_photos(items: List[dict], run_date: str, mode: str, agency: str, initials: str) -> List[dict]:
    items = filter_by_date(items, run_date)
    if mode == "individual":
        return [x for x in items if (x.get("agency") or "").upper() == agency and (x.get("user") or "").upper() == initials]
    if mode == "agency_manager":
        return [x for x in items if (x.get("agency") or "").upper() == agency] if agency else items
    return items


# =========================
# Photo -> Prospect line
# =========================
def build_photo_prospect(photo_item: dict, ai: Optional[dict], gov: Optional[dict], suggestions: List[str]) -> dict:
    # Manual priority
    city = (photo_item.get("city") or "").strip()
    meeting = photo_item.get("meeting")
    interloc = photo_item.get("interlocutor")
    phone_manual = photo_item.get("phone_manual")
    mail_manual = photo_item.get("mail_manual")

    out = {
        "date": photo_item.get("date", ""),
        "agency": (photo_item.get("agency") or "").upper(),
        "initials": (photo_item.get("user") or "").upper(),

        "name": "",
        "address": "",
        "postal_code": "",
        "city": city,

        "phone": "",
        "phone2": "",
        "email": "",

        "siret": "",
        "naf": "",
        "website": "",

        "dirigeant": "",
        "interlocuteur": "",
        "contact_civility": "",
        "contact_firstname": "",
        "contact_lastname": "",

        "resume": meeting or "",
        "commande": "",

        "card_photo_file_id": photo_item.get("file_id", ""),
        "card_photo_url": "",

        "source": "photo",
        "confidence": "",
        "status": "A_VERIFIER",
        "suggestions": "",
    }

    if ai:
        out["confidence"] = ai.get("confidence", "")
        cg = (ai.get("company_name_guess") or "").strip() if isinstance(ai.get("company_name_guess"), str) else ""
        # if gov enrich ok, prefer that as name
        if cg:
            out["name"] = cg

        c = ai.get("contact") if isinstance(ai.get("contact"), dict) else {}
        if isinstance(c.get("civility"), str):
            out["contact_civility"] = c.get("civility") or ""
        if isinstance(c.get("first_name"), str):
            out["contact_firstname"] = c.get("first_name") or ""
        if isinstance(c.get("last_name"), str):
            out["contact_lastname"] = c.get("last_name") or ""

        # AI email/phones only if not manual
        if isinstance(c.get("email"), str):
            fill_if_empty(out, "email", c.get("email"))
        if isinstance(c.get("mobile"), str):
            fill_if_empty(out, "phone2", c.get("mobile"))
        if isinstance(c.get("phone"), str):
            fill_if_empty(out, "phone", c.get("phone"))

    # Manual overrides always win
    if mail_manual:
        out["email"] = mail_manual
    if phone_manual:
        # put manual in phone (fixe) if no better, else in phone2? => we keep phone if empty, else phone2
        if not out["phone"]:
            out["phone"] = phone_manual
        elif not out["phone2"]:
            out["phone2"] = phone_manual

    # Interlocutor manual has priority
    if interloc:
        out["interlocuteur"] = interloc
        fn, ln = split_name_simple(interloc)
        fill_if_empty(out, "contact_firstname", fn)
        fill_if_empty(out, "contact_lastname", ln or interloc)

    # Deterministic gov enrich
    if gov:
        out["name"] = gov.get("name") or out["name"]
        out["address"] = gov.get("address") or ""
        out["postal_code"] = gov.get("postal_code") or ""
        out["city"] = gov.get("city") or out["city"]
        out["siret"] = gov.get("siret") or ""
        out["naf"] = gov.get("naf") or ""
        out["dirigeant"] = gov.get("dirigeant") or ""

    # Places enrich optional (if key present and we have name+city)
    if out["name"] and out["city"]:
        pl = google_places_enrich(out["name"], out["city"])
        # only fill if empty
        fill_if_empty(out, "phone", pl.get("phone"))
        fill_if_empty(out, "website", pl.get("website"))

    # status + suggestions
    if gov and out.get("siret"):
        out["status"] = "OK"
    else:
        out["status"] = "A_VERIFIER"
    if suggestions:
        out["suggestions"] = " | ".join([s for s in suggestions if s])

    # ensure contact fallback from dirigeant if still empty
    ensure_contact_fields(out)
    if not out["interlocuteur"]:
        # fallback for display/excel
        fn = (out.get("contact_firstname") or "").strip()
        ln = (out.get("contact_lastname") or "").strip()
        combo = f"{fn} {ln}".strip()
        out["interlocuteur"] = combo or out.get("dirigeant", "")

    return out


# =========================
# MAIN
# =========================
def main() -> None:
    routing = json.loads(must_env("MAIL_ROUTING_JSON"))

    send_mode = opt_env("SEND_MODE", "individual").strip()
    run_date = opt_env("RUN_DATE") or datetime.utcnow().strftime("%Y-%m-%d")

    agency = opt_env("AGENCY", "").upper()
    initials = opt_env("INITIALS", "").upper()

    max_ocr = opt_int("MAX_OCR_IMAGES", 50)
    max_photos = opt_int("MAX_PHOTO_IMAGES", 15)

    prospects_all = fetch_dump_jsonl(run_date, "prospects")
    closes_all = fetch_dump_jsonl(run_date, "closes")

    # NEW: photos
    photos_all = fetch_dump_jsonl(run_date, "photos")

    # --------------------------------
    # INDIVIDUAL
    # --------------------------------
    if send_mode == "individual":
        if not agency or not initials:
            raise RuntimeError("SEND_MODE=individual requires AGENCY and INITIALS")

        prospects = filter_prospects(prospects_all, run_date, "individual", agency, initials)
        closes = filter_closes(closes_all, run_date, "individual", agency, initials)
        photos = filter_photos(photos_all, run_date, "individual", agency, initials)

        # --- V1: OCR card attachments for prospects (unchanged behavior) ---
        card_attachments: List[Tuple[str, bytes]] = []
        ocr_done = 0

        for rec in prospects:
            ensure_contact_fields(rec)

            need_email = not (rec.get("email") or "").strip()
            need_mobile = not (rec.get("phone2") or "").strip()
            need_contact = not ((rec.get("contact_firstname") or "").strip() or (rec.get("contact_lastname") or "").strip())
            need_phone = not (rec.get("phone") or "").strip()

            has_card = (rec.get("card_photo_url") or rec.get("card_photo_file_id") or "").strip() != ""
            if has_card and (need_email or need_mobile or need_contact or need_phone) and ocr_done < max_ocr:
                img = download_card_image_bytes(rec)
                if img:
                    o = ocr_extract(img)
                    if need_email and o.get("email"):
                        fill_if_empty(rec, "email", o["email"])
                    if need_mobile and o.get("mobile"):
                        fill_if_empty(rec, "phone2", o["mobile"])
                    if need_phone and o.get("phone"):
                        fill_if_empty(rec, "phone", o["phone"])
                    if need_contact and o.get("contact_full"):
                        fn, ln = split_name_simple(o["contact_full"])
                        if fn:
                            fill_if_empty(rec, "contact_firstname", fn)
                        if ln:
                            fill_if_empty(rec, "contact_lastname", ln)
                    ocr_done += 1
                    card_attachments.append((f"carte_{agency}_{initials}_{ocr_done}.jpg", img))

            ensure_contact_fields(rec)

        # --- V2 Photo: process photos into prospect-like lines ---
        photo_attachments: List[Tuple[str, bytes]] = []
        photo_lines: List[dict] = []

        processed = 0
        for p in photos:
            if processed >= max_photos:
                break
            file_id = (p.get("file_id") or "").strip()
            city = (p.get("city") or "").strip()
            if not file_id or not city:
                continue

            img_bytes = telegram_download_file_bytes(file_id)
            if not img_bytes:
                continue

            ai = None
            if _gemini_enabled():
                ai = ai_photo_extract(img_bytes, city)

            suggestions: List[str] = []
            gov = None
            if ai and isinstance(ai.get("company_name_guess"), str) and ai.get("company_name_guess").strip():
                gov, suggestions = enrich_from_gov(ai["company_name_guess"].strip(), city)
            else:
                # fallback: if no company guess, no gov enrich (still exported as A_VERIFIER)
                gov, suggestions = (None, [])

            line = build_photo_prospect(p, ai, gov, suggestions)
            photo_lines.append(line)

            # attach original image (for user immediate mail)
            processed += 1
            photo_attachments.append((f"photo_{agency}_{initials}_{processed}.jpg", img_bytes))

        # Merge all for Excel
        merged = prospects + photo_lines

        xlsx_name = f"prospection_{agency}_{initials}_{run_date}.xlsx"
        build_excel_from_template(merged, xlsx_name)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        attachments = [(xlsx_name, excel_bytes)] + card_attachments + photo_attachments

        to_list = routing.get(agency, {}).get(initials)
        if not to_list:
            raise RuntimeError(f"No routing found for {agency}/{initials}")

        clients = sum(to_int(x.get("visits_clients")) for x in closes)
        prosps = sum(to_int(x.get("visits_prospects")) for x in closes)
        comm = sum(to_int(x.get("commandes")) for x in closes)

        html = f"""
        <p>Bonjour,</p>
        <p>Voici ton export <b>{agency}/{initials}</b> du <b>{run_date}</b>.</p>
        <ul>
          <li><b>Prospects (mode entreprise)</b> : {len(prospects)}</li>
          <li><b>Prospects (mode photo)</b> : {len(photo_lines)}</li>
          <li><b>Clients</b> : {clients}</li>
          <li><b>Prospects visités</b> : {prosps}</li>
          <li><b>Commandes</b> : {comm}</li>
        </ul>
        <p>Pièces jointes : Excel + images (cartes/photos).</p>
        """
        subject = f"[PROSPECTION] {agency}/{initials} - {run_date}"
        send_mail_brevo(to_list, subject, html, attachments)

        print(f"OK individual: {agency}/{initials} prospects={len(prospects)} photo={len(photo_lines)} closes={len(closes)}")
        return

    # ---------------------------------------
    # AGENCY MANAGER
    # ---------------------------------------
    if send_mode == "agency_manager":
        agencies = [agency] if agency else [k for k in routing.keys() if k != "_admin"]

        for ag in agencies:
            if not ag:
                continue

            prospects = filter_prospects(prospects_all, run_date, "agency_manager", ag, "")
            closes = filter_closes(closes_all, run_date, "agency_manager", ag, "")
            photos = filter_photos(photos_all, run_date, "agency_manager", ag, "")

            photo_lines: List[dict] = []
            processed = 0
            for p in photos:
                if processed >= 30:  # consolidation agency: cap
                    break
                file_id = (p.get("file_id") or "").strip()
                city = (p.get("city") or "").strip()
                if not file_id or not city:
                    continue

                img_bytes = telegram_download_file_bytes(file_id)
                if not img_bytes:
                    continue

                ai = ai_photo_extract(img_bytes, city) if _gemini_enabled() else None

                suggestions: List[str] = []
                gov = None
                if ai and isinstance(ai.get("company_name_guess"), str) and ai.get("company_name_guess").strip():
                    gov, suggestions = enrich_from_gov(ai["company_name_guess"].strip(), city)

                line = build_photo_prospect(p, ai, gov, suggestions)
                photo_lines.append(line)
                processed += 1

            merged = prospects + photo_lines
            for rec in merged:
                ensure_contact_fields(rec)

            xlsx_name = f"prospection_{ag}_CONSOLIDE_{run_date}.xlsx"
            build_excel_from_template(merged, xlsx_name)
            with open(xlsx_name, "rb") as f:
                excel_bytes = f.read()

            to_list = routing.get(ag, {}).get("_manager")
            if not to_list:
                print(f"SKIP agency {ag}: missing _manager")
                continue

            clients = sum(to_int(x.get("visits_clients")) for x in closes)
            prosps = sum(to_int(x.get("visits_prospects")) for x in closes)
            comm = sum(to_int(x.get("commandes")) for x in closes)

            html = f"""
            <p>Bonjour,</p>
            <p>Récap <b>agence {ag}</b> du <b>{run_date}</b>.</p>
            <ul>
              <li><b>Prospects (mode entreprise)</b> : {len(prospects)}</li>
              <li><b>Prospects (mode photo)</b> : {len(photo_lines)}</li>
              <li><b>Clients</b> : {clients}</li>
              <li><b>Prospects visités</b> : {prosps}</li>
              <li><b>Commandes</b> : {comm}</li>
            </ul>
            <p>Pièce jointe : Excel consolidé agence.</p>
            """
            subject = f"[PROSPECTION] Récap agence {ag} - {run_date}"
            send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
            print(f"OK agency_manager: {ag} prospects={len(prospects)} photo={len(photo_lines)} closes={len(closes)}")

        return

    # ---------------------------------------
    # ADMIN
    # ---------------------------------------
    if send_mode == "admin":
        prospects = filter_by_date(prospects_all, run_date)
        closes = filter_by_date(closes_all, run_date)
        photos = filter_by_date(photos_all, run_date)

        photo_lines: List[dict] = []
        processed = 0
        for p in photos:
            if processed >= 80:  # global cap
                break
            file_id = (p.get("file_id") or "").strip()
            city = (p.get("city") or "").strip()
            if not file_id or not city:
                continue
            img_bytes = telegram_download_file_bytes(file_id)
            if not img_bytes:
                continue

            ai = ai_photo_extract(img_bytes, city) if _gemini_enabled() else None

            suggestions: List[str] = []
            gov = None
            if ai and isinstance(ai.get("company_name_guess"), str) and ai.get("company_name_guess").strip():
                gov, suggestions = enrich_from_gov(ai["company_name_guess"].strip(), city)

            line = build_photo_prospect(p, ai, gov, suggestions)
            photo_lines.append(line)
            processed += 1

        merged = prospects + photo_lines
        for rec in merged:
            ensure_contact_fields(rec)

        xlsx_name = f"prospection_GLOBAL_{run_date}.xlsx"
        build_excel_from_template(merged, xlsx_name)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        stats = compute_stats(merged, closes)
        totals = stats["totals"]

        agency_rows = []
        for ag, v in sorted(stats["by_agency"].items(), key=lambda x: x[0]):
            agency_rows.append([ag, v["actions"], v["clients"], v["prospects"], v["commandes"]])

        initials_rows = []
        for ini, v in sorted(stats["by_initials"].items(), key=lambda x: x[0]):
            if not ini:
                continue
            initials_rows.append([ini, ", ".join(v["agencies"]) if v["agencies"] else "", v["actions"], v["clients"], v["prospects"], v["commandes"]])

        ranking_rows = [[ini, actions] for (ini, actions) in stats["ranking"]]

        html = f"""
        <p>Bonsoir SL,</p>
        <p>Récap GLOBAL du <b>{run_date}</b>.</p>
        <ul>
          <li><b>Actions (prospects saisis)</b> : {totals["actions"]}</li>
          <li><b>Clients</b> : {totals["clients"]}</li>
          <li><b>Prospects visités</b> : {totals["prospects"]}</li>
          <li><b>Commandes</b> : {totals["commandes"]}</li>
          <li><b>Mode photo</b> : {len(photo_lines)} lignes</li>
        </ul>

        <h3>Tableau par agence</h3>
        {html_table(["Agence","Actions","Clients","Prospects","Commandes"], agency_rows)}

        <h3>Tableau par commercial</h3>
        {html_table(["Initiales","Agences","Actions","Clients","Prospects","Commandes"], initials_rows)}

        <h3>Classement (par nombre d’actions)</h3>
        {html_table(["Initiales","Actions"], ranking_rows)}

        <p>Pièce jointe : Excel GLOBAL.</p>
        """

        to_list = routing.get("_admin")
        if not to_list:
            raise RuntimeError("MAIL_ROUTING_JSON missing _admin recipients")

        subject = f"[PROSPECTION] Récap GLOBAL - {run_date}"
        send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
        print(f"OK admin: merged={len(merged)} closes={len(closes)} photo={len(photo_lines)}")
        return

    raise RuntimeError("Invalid SEND_MODE. Use individual|agency_manager|admin")


if __name__ == "__main__":
    main()