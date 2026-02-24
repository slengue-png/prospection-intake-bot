import os
import re
import json
import io
import base64
from datetime import datetime, time as dtime
from typing import Dict, List, Tuple, Optional
from zoneinfo import ZoneInfo

import requests
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from PIL import Image
import pytesseract


# =====================================================
# CONFIG / CONSTANTES
# =====================================================

PARIS = ZoneInfo("Europe/Paris")

EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)\s*[1-9](?:[\s.\-]*\d{2}){4})")

DEFAULT_ROUTING = {
    "GR": {"JL": ["sam9999@free.fr"], "CZ": ["slengue@gmail.com"], "_manager": ["sam9999@free.fr"]},
    "VR": {"JB": ["sam9999@free.fr"], "LB": ["slengue@gmail.com"], "_manager": ["sam9999@free.fr"]},
    "GRS": {"PV": ["sam9999@free.fr"], "ST": ["slengue@gmail.com"], "_manager": ["sam9999@free.fr"]},
    "SLS": {"AC": ["sam9999@free.fr"], "_manager": ["sam9999@free.fr"]},
    "_admin": ["samuel.lengue@ras-interim.fr"],
}

# Colonnes Excel (on met AGENCE + INITIALS en dernier)
EXCEL_COLUMNS = [
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
    "AGENCE",
    "INITIALS",
]


# =====================================================
# HELPERS ENV / STRING
# =====================================================

def must_env(name: str) -> str:
    v = os.getenv(name, "").strip()
    if not v:
        raise RuntimeError(f"Missing required env var: {name}")
    return v

def opt_env(name: str) -> str:
    return os.getenv(name, "").strip()

def safe_filename(s: str) -> str:
    s = re.sub(r"[^\w\-\. ]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:160] if len(s) > 160 else s

def normalize_naf(naf: str) -> str:
    if not naf:
        return ""
    return str(naf).strip().upper().replace(".", "").replace(" ", "")

def iso_to_paris(iso: str) -> Optional[datetime]:
    if not iso:
        return None
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.astimezone(PARIS)
    except Exception:
        return None

def split_name_simple(full: str) -> Tuple[str, str]:
    """
    "Jean Dupont" -> ("Jean", "Dupont")
    "Dupont" -> ("", "Dupont")  (1 mot => NOM)
    """
    s = (full or "").strip()
    s = re.sub(r"\s+", " ", s)
    if not s:
        return ("", "")
    parts = s.split(" ")
    if len(parts) == 1:
        return ("", parts[0])
    return (parts[0], " ".join(parts[1:]))

def load_routing() -> dict:
    s = opt_env("MAIL_ROUTING_JSON")
    if not s:
        return DEFAULT_ROUTING
    try:
        return json.loads(s)
    except Exception:
        return DEFAULT_ROUTING


# =====================================================
# WORKER DUMP
# =====================================================

def build_dump_url(base: str, date: str, kind: str) -> str:
    base = base.strip().rstrip("/")
    return f"{base}/dump?date={date}&kind={kind}"

def fetch_dump_jsonl(date: str, kind: str) -> List[dict]:
    base = must_env("WORKER_BASE_URL")
    token = must_env("EXPORT_TOKEN")
    url = build_dump_url(base, date, kind)
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


# =====================================================
# TELEGRAM FILE DOWNLOAD (fallback si besoin)
# =====================================================

def telegram_get_file_path(file_id: str) -> Optional[str]:
    token = must_env("TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/bot{token}/getFile"
    r = requests.post(url, json={"file_id": file_id}, timeout=30)
    r.raise_for_status()
    j = r.json()
    if not j.get("ok"):
        return None
    return j["result"].get("file_path")

def telegram_download_file(file_path: str) -> bytes:
    token = must_env("TELEGRAM_TOKEN")
    url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.content


# =====================================================
# OCR BUSINESS CARD
# =====================================================

def ocr_extract_contact_fields(image_bytes: bytes) -> Dict[str, str]:
    """
    Retourne: {contact_full, email, mobile, phone}
    """
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return {"contact_full": "", "email": "", "mobile": "", "phone": ""}

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
        p2 = re.sub(r"[^\d+]", "", p)
        if p2.startswith("+33"):
            p2 = "0" + p2[3:]
        p2 = re.sub(r"\D", "", p2)
        if len(p2) == 10 and p2.startswith("0"):
            normalized.append(p2)

    mobile = next((x for x in normalized if x.startswith("06") or x.startswith("07")), "")
    phone = next((x for x in normalized if x.startswith("04")), "")
    if not phone:
        phone = next((x for x in normalized if x.startswith(("01", "02", "03", "05", "09", "08"))), "")

    # Heuristique nom/prénom
    contact_full = ""
    candidates = []
    for ln in lines[:20]:
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

    return {"contact_full": contact_full, "email": email, "mobile": mobile, "phone": phone}


def ensure_contact_fields(r: dict) -> None:
    """
    PRIORITÉ :
    1) Saisie manuelle (interlocuteur) -> split => contact_firstname / contact_lastname
    2) OCR (contact_firstname/contact_lastname déjà remplis ou interlocuteur rempli par OCR)
    3) Dirigeant
    Et règle spéciale : si 1 seul mot, on le met en NOM.
    """
    # 1) Interlocuteur saisi => prime toujours
    interloc = (r.get("interlocuteur") or "").strip()
    if interloc:
        fn, ln = split_name_simple(interloc)
        if fn and ln:
            r["contact_firstname"] = fn
            r["contact_lastname"] = ln
        else:
            # 1 seul mot => NOM
            r["contact_firstname"] = ""
            r["contact_lastname"] = interloc
        return

    # 2) si OCR a déjà rempli prénom/nom, on garde
    if (r.get("contact_firstname") or "").strip() or (r.get("contact_lastname") or "").strip():
        return

    # si OCR a rempli un interlocuteur mais pas contact_*
    if (r.get("interlocuteur") or "").strip():
        fn, ln = split_name_simple(r["interlocuteur"])
        if fn and ln:
            r["contact_firstname"] = fn
            r["contact_lastname"] = ln
        else:
            r["contact_firstname"] = ""
            r["contact_lastname"] = r["interlocuteur"]
        return

    # 3) fallback dirigeant
    dirg = (r.get("dirigeant") or "").strip()
    if dirg:
        fn, ln = split_name_simple(dirg)
        r["contact_firstname"] = fn
        r["contact_lastname"] = ln or dirg


def maybe_enrich_from_ocr(records: List[dict]) -> List[Tuple[str, bytes]]:
    """
    Télécharge les cartes de visite et complète :
    - interlocuteur (si vide)
    - email (si vide)
    - phone2 (mobile si vide)
    - phone (fixe si vide)
    - contact_firstname/contact_lastname via ensure_contact_fields()
    Retourne les PJ images.
    """
    attachments: List[Tuple[str, bytes]] = []

    for r in records:
        file_id = (r.get("card_photo_file_id") or "").strip()
        url = (r.get("card_photo_url") or "").strip()

        if not file_id and not url:
            continue

        img_bytes = b""
        try:
            if url:
                resp = requests.get(url, timeout=60)
                resp.raise_for_status()
                img_bytes = resp.content
            else:
                fp = telegram_get_file_path(file_id)
                if not fp:
                    continue
                img_bytes = telegram_download_file(fp)
        except Exception:
            continue

        o = ocr_extract_contact_fields(img_bytes)

        # Interlocuteur si vide
        if not (r.get("interlocuteur") or "").strip() and o.get("contact_full"):
            r["interlocuteur"] = o["contact_full"]

        # Email si vide
        if not (r.get("email") or "").strip() and o.get("email"):
            r["email"] = o["email"]

        # Phones
        if not (r.get("phone") or "").strip() and o.get("phone"):
            r["phone"] = o["phone"]
        if not (r.get("phone2") or "").strip() and o.get("mobile"):
            r["phone2"] = o["mobile"]

        # Contact prenom/nom (avec priorité interlocuteur>ocr>dirigeant)
        ensure_contact_fields(r)

        att_name = safe_filename(f"carte_{r.get('name','societe')}_{file_id or 'url'}.jpg")
        attachments.append((att_name, img_bytes))

    return attachments


# =====================================================
# EXCEL BUILDER (avec template si présent)
# =====================================================

def build_excel(records: List[dict], xlsx_path: str, template_path: Optional[str] = None) -> None:
    if template_path and os.path.exists(template_path):
        wb = load_workbook(template_path)
        ws = wb.active
        # Nettoie les lignes données (garde l'entête)
        if ws.max_row >= 2:
            ws.delete_rows(2, ws.max_row - 1)
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "PROSPECTION"
        ws.append(EXCEL_COLUMNS)

    for r in records:
        ensure_contact_fields(r)

        row = [
            r.get("name", "") or "",
            r.get("address", "") or "",
            r.get("postal_code", "") or "",
            r.get("city", "") or "",
            r.get("phone", "") or "",
            r.get("phone2", "") or "",
            r.get("email", "") or "",
            r.get("siret", "") or "",
            normalize_naf(r.get("naf", "")),
            r.get("website", "") or "",
            r.get("contact_civility", "") or "",
            r.get("contact_firstname", "") or "",
            r.get("contact_lastname", "") or "",
            r.get("dirigeant", "") or "",
            r.get("resume", "") or "",
            r.get("commande", "") or "",
            r.get("infos_commerciales", "") or "",
            r.get("card_photo_file_id", "") or "",
            r.get("agency", "") or "",
            r.get("initials", "") or "",
        ]
        ws.append(row)

    # Largeur colonnes (si template, on laisse souvent les widths existants)
    if not template_path:
        for col in range(1, len(EXCEL_COLUMNS) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 20

    wb.save(xlsx_path)


# =====================================================
# BREVO SEND
# =====================================================

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


# =====================================================
# STATS ADMIN (CLOSES)
# =====================================================

def build_admin_stats_from_closes(closes: List[dict]) -> str:
    stats: Dict[str, Dict[str, Dict[str, int]]] = {}
    total = {"prospects": 0, "clients": 0, "commandes": 0}

    for c in closes:
        ag = (c.get("agency") or "").strip().upper()
        ini = (c.get("initials") or "").strip().upper()
        if not ag or not ini:
            continue

        p = int(c.get("visits_prospects") or 0)
        cl = int(c.get("visits_clients") or 0)
        co = int(c.get("commandes") or 0)

        stats.setdefault(ag, {})
        stats[ag].setdefault(ini, {"prospects": 0, "clients": 0, "commandes": 0})

        stats[ag][ini]["prospects"] += p
        stats[ag][ini]["clients"] += cl
        stats[ag][ini]["commandes"] += co

        total["prospects"] += p
        total["clients"] += cl
        total["commandes"] += co

    html = "<h2>📊 RÉCAP GLOBAL</h2>"
    for ag in sorted(stats.keys()):
        html += f"<h3>🏢 {ag}</h3>"
        html += """
        <table border="1" cellpadding="6" cellspacing="0">
        <tr><th>Collaborateur</th><th>Prospects</th><th>Clients</th><th>Commandes</th></tr>
        """
        for ini in sorted(stats[ag].keys()):
            d = stats[ag][ini]
            html += f"<tr><td>{ini}</td><td>{d['prospects']}</td><td>{d['clients']}</td><td>{d['commandes']}</td></tr>"
        html += "</table><br>"

    html += f"""
    <h3>🔢 TOTAL GLOBAL</h3>
    <ul>
      <li>Prospects : {total['prospects']}</li>
      <li>Clients : {total['clients']}</li>
      <li>Commandes : {total['commandes']}</li>
    </ul>
    """
    return html


# =====================================================
# late_update logic
# =====================================================

def has_new_since_1745(prospects: List[dict], closes: List[dict]) -> bool:
    threshold = dtime(17, 45)
    for r in prospects:
        dt = iso_to_paris(r.get("created_at", ""))
        if dt and dt.time() > threshold:
            return True
    for c in closes:
        dt = iso_to_paris(c.get("closed_at", ""))
        if dt and dt.time() > threshold:
            return True
    return False


# =====================================================
# Filtering helpers
# =====================================================

def filter_records(records: List[dict], date: str, agency: Optional[str], initials: Optional[str]) -> List[dict]:
    out: List[dict] = []
    for r in records:
        if (r.get("date") or "").strip() != date:
            continue
        if agency and (r.get("agency") or "").strip().upper() != agency.upper():
            continue
        if initials and (r.get("initials") or "").strip().upper() != initials.upper():
            continue
        out.append(r)
    return out


# =====================================================
# MAIN
# =====================================================

def main():
    send_mode = (opt_env("SEND_MODE") or "individual").strip()
    routing = load_routing()

    run_date = opt_env("RUN_DATE")
    if not run_date:
        run_date = datetime.now(PARIS).strftime("%Y-%m-%d")

    agency = opt_env("AGENCY").strip().upper()
    initials = opt_env("INITIALS").strip().upper()

    prospects_all = fetch_dump_jsonl(run_date, "prospects")
    closes_all = fetch_dump_jsonl(run_date, "closes")

    template_path = "Table.2.xlsx" if os.path.exists("Table.2.xlsx") else None

    # -------------------------------------------------
    # INDIVIDUAL: mail immédiat collaborateur
    # -------------------------------------------------
    if send_mode == "individual":
        if not agency or not initials:
            raise RuntimeError("SEND_MODE=individual requires AGENCY and INITIALS")

        prospects = filter_records(prospects_all, run_date, agency, initials)

        # OCR cartes: complète champs + PJ cartes
        card_attachments = maybe_enrich_from_ocr(prospects)

        xlsx_name = safe_filename(f"prospection_{agency}_{initials}_{run_date}.xlsx")
        build_excel(prospects, xlsx_name, template_path=template_path)

        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        attachments = [(xlsx_name, excel_bytes)] + card_attachments

        to_list = routing.get(agency, {}).get(initials, [])
        if not to_list:
            raise RuntimeError(f"No recipient mapping for {agency}/{initials}")

        subject = f"[PROSPECTION] {agency}/{initials} - {run_date}"
        html = f"""
        <p>Bonjour,</p>
        <p>Voici l’export de prospection <b>{agency}/{initials}</b> du <b>{run_date}</b>.</p>
        <p>Pièces jointes : Excel + cartes de visite.</p>
        """
        send_mail_brevo(to_list, subject, html, attachments)
        print(f"OK individual {agency}/{initials} records={len(prospects)} cards={len(card_attachments)}")
        return

    # -------------------------------------------------
    # AGENCY_MANAGER: 1 mail par agence au responsable
    # -------------------------------------------------
    if send_mode == "agency_manager":
        for ag in ["GR", "VR", "GRS", "SLS"]:
            prospects = filter_records(prospects_all, run_date, ag, None)

            xlsx_name = safe_filename(f"prospection_{ag}_CONSOLIDE_{run_date}.xlsx")
            build_excel(prospects, xlsx_name, template_path=template_path)

            with open(xlsx_name, "rb") as f:
                excel_bytes = f.read()

            to_list = routing.get(ag, {}).get("_manager", [])
            if not to_list:
                print(f"Skip agency {ag}: no _manager")
                continue

            subject = f"[PROSPECTION] Récap agence {ag} - {run_date}"
            html = f"""
            <p>Bonjour,</p>
            <p>Voici le récapitulatif consolidé de l’agence <b>{ag}</b> pour le <b>{run_date}</b>.</p>
            <p>Pièce jointe : Excel consolidé.</p>
            """
            send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
            print(f"OK agency_manager {ag} records={len(prospects)}")
        return

    # -------------------------------------------------
    # ADMIN: mail global + stats closes
    # -------------------------------------------------
    if send_mode == "admin":
        prospects = filter_records(prospects_all, run_date, None, None)
        closes = filter_records(closes_all, run_date, None, None)

        xlsx_name = safe_filename(f"prospection_GLOBAL_{run_date}.xlsx")
        build_excel(prospects, xlsx_name, template_path=template_path)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        stats_html = build_admin_stats_from_closes(closes)

        to_list = routing.get("_admin", [])
        if not to_list:
            raise RuntimeError("No _admin routing configured")

        subject = f"[PROSPECTION] Récap GLOBAL - {run_date}"
        html = f"""
        <p>Bonsoir,</p>
        {stats_html}
        <p>Excel global en pièce jointe.</p>
        """
        send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
        print(f"OK admin records={len(prospects)} closes={len(closes)}")
        return

    # -------------------------------------------------
    # LATE_UPDATE: 23:59 uniquement si nouveautés après 17:45
    # -------------------------------------------------
    if send_mode == "late_update":
        prospects = filter_records(prospects_all, run_date, None, None)
        closes = filter_records(closes_all, run_date, None, None)

        if not has_new_since_1745(prospects, closes):
            print("No new records since 17:45 -> no mail sent")
            return

        xlsx_name = safe_filename(f"prospection_GLOBAL_{run_date}_LATE.xlsx")
        build_excel(prospects, xlsx_name, template_path=template_path)
        with open(xlsx_name, "rb") as f:
            excel_bytes = f.read()

        stats_html = build_admin_stats_from_closes(closes)
        to_list = routing.get("_admin", [])
        if not to_list:
            raise RuntimeError("No _admin routing configured")

        subject = f"[PROSPECTION] MAJ 23:59 - {run_date}"
        html = f"""
        <p>Bonsoir,</p>
        <p><b>Mise à jour tardive</b> (nouvelles saisies après 17:45).</p>
        {stats_html}
        <p>Excel en pièce jointe.</p>
        """
        send_mail_brevo(to_list, subject, html, [(xlsx_name, excel_bytes)])
        print("OK late_update sent")
        return

    raise RuntimeError(f"Invalid SEND_MODE: {send_mode}")


if __name__ == "__main__":
    main()
