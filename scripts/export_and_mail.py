import os
import re
import json
import time
import io
from datetime import datetime
from typing import Dict, List, Tuple, Optional

import requests
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from PIL import Image
import pytesseract


# =========================
# CONFIG
# =========================

# Mapping destinataires (TEST comme tu l'as donné)
# Clé = (AGENCY, INITIALS) => email destinataire immédiat
RECIPIENTS_SESSION = {
    ("GR", "JL"): "sam9999@free.fr",
    ("GR", "CZ"): "slengue@gmail.com",
    ("VR", "JB"): "sam9999@free.fr",
    ("VR", "LB"): "slengue@gmail.com",
    ("GRS", "PV"): "sam9999@free.fr",
    ("GRS", "ST"): "slengue@gmail.com",
    ("SLS", "AC"): "sam9999@free.fr",
}

# Récap agence fin de journée => responsable
RECIPIENTS_DAILY_AGENCY = {
    "GR": "sam9999@free.fr",     # JL
    "VR": "sam9999@free.fr",     # JB
    "GRS": "sam9999@free.fr",    # PV
    "SLS": "sam9999@free.fr",    # AC
}

# Ton récap global (17:47)
RECIPIENTS_DAILY_GLOBAL = ["samuel.lengue@ras-interim.fr"]


# Colonnes Excel attendues (ordre)
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
]

# OCR — on cherche surtout nom + email + tel
EMAIL_RE = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:(?:\+33|0)\s*[1-9](?:[\s.\-]*\d{2}){4})")


def must_env(name: str) -> str:
    v = os.getenv(name, "").strip()
    if not v:
        raise RuntimeError(f"Missing required env var: {name}")
    return v


def normalize_naf(naf: str) -> str:
    # "23.12Z" -> "2312Z"
    if not naf:
        return ""
    naf = str(naf).strip().upper()
    naf = naf.replace(".", "").replace(" ", "")
    return naf


def join_address(address: str, postal: str, city: str) -> str:
    parts = [str(x).strip() for x in [address, postal, city] if str(x).strip()]
    return " ".join(parts).strip()


def safe_filename(s: str) -> str:
    s = re.sub(r"[^\w\-\. ]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:120] if len(s) > 120 else s


def build_dump_url(base: str, date: str, kind: str) -> str:
    base = base.strip()
    if base.endswith("/"):
        base = base[:-1]
    if not base.startswith("http://") and not base.startswith("https://"):
        raise RuntimeError("WORKER_BASE_URL must start with https:// or http://")
    return f"{base}/dump?date={date}&kind={kind}"


def fetch_dump_jsonl(date: str, kind: str) -> List[dict]:
    base = must_env("WORKER_BASE_URL")
    token = must_env("EXPORT_TOKEN")

    url = build_dump_url(base, date, kind)
    headers = {"X-Export-Token": token}

    r = requests.get(url, headers=headers, timeout=60)
    r.raise_for_status()

    lines = [ln.strip() for ln in r.text.splitlines() if ln.strip()]
    out = []
    for ln in lines:
        try:
            out.append(json.loads(ln))
        except Exception:
            # si jamais une ligne est déjà JSON stringifié bizarrement
            continue
    return out


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


def ocr_extract_contact_fields(image_bytes: bytes) -> Dict[str, str]:
    """
    Retourne: {contact_full, email, mobile, phone}
    - on privilégie mobile si trouvé 06/07
    - on tente de récupérer un "nom" en détectant une ligne nom/prénom probable (heuristique)
    """
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    except Exception:
        return {"contact_full": "", "email": "", "mobile": "", "phone": ""}

    # OCR
    text = pytesseract.image_to_string(img, lang="fra+eng")
    text = re.sub(r"[ \t]+", " ", text)
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    # Email
    email = ""
    m = EMAIL_RE.search(text)
    if m:
        email = m.group(0).strip()

    # Phones
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
        # autre fixe si pas 04
        phone = next((x for x in normalized if x.startswith(("01", "02", "03", "05", "09", "08"))), "")

    # Heuristique contact: chercher une ligne "Prénom NOM" (ex Céline MEKH...),
    # souvent au-dessus du poste.
    contact_full = ""
    # 1) lignes courtes candidates
    candidates = []
    for ln in lines[:20]:
        if len(ln) < 4 or len(ln) > 60:
            continue
        # éviter "VEOLIA", "SAS", etc.
        upper = ln.upper()
        if any(k in upper for k in ["SAS", "SARL", "EURL", "FRANCE", "GROUPE", "VEOLIA", "WWW", "HTTP", "@"]):
            continue
        # au moins 2 mots
        if len(ln.split()) < 2:
            continue
        # contient des lettres
        if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", ln):
            continue
        candidates.append(ln)

    # 2) préférer la ligne où il y a un mot en MAJ (souvent NOM)
    for ln in candidates:
        if re.search(r"\b[A-ZÀ-ÖØ-Ý]{3,}\b", ln):
            contact_full = ln
            break
    if not contact_full and candidates:
        contact_full = candidates[0]

    return {"contact_full": contact_full, "email": email, "mobile": mobile, "phone": phone}


def split_contact(contact_full: str) -> Tuple[str, str, str]:
    """
    Retourne (civilité, prenom, nom)
    """
    if not contact_full:
        return ("", "", "")
    s = contact_full.strip()
    civ = ""
    # petites civilités possibles
    s = re.sub(r"^(M\.?|MME|Mme|Monsieur|Madame)\s+", lambda m: (m.group(0).strip() + " "), s, flags=re.IGNORECASE)
    # on extrait civ si présent
    m = re.match(r"^(M\.?|MME|Mme|Monsieur|Madame)\s+(.*)$", s, flags=re.IGNORECASE)
    if m:
        civ = m.group(1).replace("Monsieur", "M.").replace("Madame", "Mme").strip()
        s = m.group(2).strip()

    parts = s.split()
    if len(parts) == 1:
        return (civ, "", parts[0])
    prenom = parts[0]
    nom = " ".join(parts[1:])
    return (civ, prenom, nom)


def filter_records(records: List[dict], date: str, agency: str, initials: Optional[str], mode: str) -> List[dict]:
    """
    mode:
      - session: filtre date + agency + initials
      - agency_daily: filtre date + agency (tous collaborateurs)
      - global_daily: filtre date (toutes agences)
    """
    out = []
    for r in records:
        if str(r.get("date", "")).strip() != date:
            continue
        if mode in ("session", "agency_daily") and str(r.get("agency", "")).strip().upper() != agency.upper():
            continue
        if mode == "session":
            if str(r.get("initials", "")).strip().upper() != (initials or "").upper():
                continue
        out.append(r)
    return out


def build_excel(records: List[dict], xlsx_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "PROSPECTION"

    ws.append(EXCEL_COLUMNS)

    for r in records:
        naf = normalize_naf(r.get("naf", ""))
        row = [
            r.get("name", "") or "",
            join_address(r.get("address", ""), r.get("postal_code", ""), r.get("city", "")),
            r.get("postal_code", "") or "",
            r.get("city", "") or "",
            r.get("phone", "") or "",
            r.get("phone2", "") or "",  # tu utilises phone2 comme "téléphone 2"
            r.get("email", "") or "",
            r.get("siret", "") or "",
            naf,
            r.get("website", "") or "",
            r.get("contact_civility", "") or "",
            r.get("contact_firstname", "") or "",
            r.get("contact_lastname", "") or "",
            r.get("dirigeant", "") or "",
            r.get("resume", "") or "",
            r.get("commande", "") or "",
            r.get("infos_commerciales", "") or "",
            r.get("card_photo_file_id", "") or "",
        ]
        ws.append(row)

    # autosize simple
    for col in range(1, len(EXCEL_COLUMNS) + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = 20

    wb.save(xlsx_path)


def send_mail_brevo(to_list: List[str], subject: str, html: str, attachments: List[Tuple[str, bytes]]) -> None:
    api_key = must_env("BREVO_API_KEY")
    mail_from = must_env("MAIL_FROM")

    payload = {
        "sender": {"email": mail_from, "name": "Prospection Bot"},
        "to": [{"email": x} for x in to_list],
        "subject": subject,
        "htmlContent": html,
        "attachment": [
            {"name": name, "content": (content_bytes).decode("latin1")}
            for (name, content_bytes) in []
        ],
    }

    # Brevo attend base64 pour attachment.content
    import base64
    payload["attachment"] = [{"name": name, "content": base64.b64encode(b).decode("utf-8")} for name, b in attachments]

    r = requests.post(
        "https://api.brevo.com/v3/smtp/email",
        headers={"api-key": api_key, "content-type": "application/json"},
        json=payload,
        timeout=60,
    )
    r.raise_for_status()


def main():
    """
    Paramètres via env (venant du workflow_dispatch):
      MODE = session | agency_daily | global_daily
      RUN_DATE = YYYY-MM-DD
      AGENCY = GR/VR/GRS/SLS (si mode session/agency_daily)
      INITIALS = JL/CZ/... (si mode session)
      VISITS_CLIENTS / VISITS_PROSPECTS / COMMANDES (optionnel: stats)
    """
    mode = os.getenv("MODE", "session").strip()
    run_date = os.getenv("RUN_DATE", "").strip()
    if not run_date:
        # fallback = date du jour (UTC -> pas parfait, mais ton worker stocke en Europe/Paris; normalement RUN_DATE est fourni)
        run_date = datetime.utcnow().strftime("%Y-%m-%d")

    agency = os.getenv("AGENCY", "").strip().upper()
    initials = os.getenv("INITIALS", "").strip().upper()

    visits_clients = os.getenv("VISITS_CLIENTS", "").strip()
    visits_prospects = os.getenv("VISITS_PROSPECTS", "").strip()
    commandes = os.getenv("COMMANDES", "").strip()

    # 1) récup prospects du jour
    prospects_all = fetch_dump_jsonl(run_date, "prospects")

    # 2) filtre selon mode
    if mode == "session":
        if not agency or not initials:
            raise RuntimeError("MODE=session requires AGENCY and INITIALS")
        prospects = filter_records(prospects_all, run_date, agency, initials, mode)
    elif mode == "agency_daily":
        if not agency:
            raise RuntimeError("MODE=agency_daily requires AGENCY")
        prospects = filter_records(prospects_all, run_date, agency, None, mode)
    elif mode == "global_daily":
        prospects = filter_records(prospects_all, run_date, "", None, mode)
    else:
        raise RuntimeError("Invalid MODE. Use session|agency_daily|global_daily")

    # 3) OCR sur cartes (uniquement MODE=session : tu veux PJ + enrichir contact)
    attachments: List[Tuple[str, bytes]] = []
    if mode == "session":
        for r in prospects:
            file_id = (r.get("card_photo_file_id") or "").strip()
            if not file_id:
                continue

            # download
            fp = telegram_get_file_path(file_id)
            if not fp:
                continue
            img_bytes = telegram_download_file(fp)

            # OCR -> enrichit contact + email/phone si manquant
            o = ocr_extract_contact_fields(img_bytes)

            # contact
            if not r.get("interlocuteur") and o.get("contact_full"):
                r["interlocuteur"] = o["contact_full"]

            civ, prenom, nom = split_contact(o.get("contact_full", ""))
            if not r.get("contact_civility") and civ:
                r["contact_civility"] = civ
            if not r.get("contact_firstname") and prenom:
                r["contact_firstname"] = prenom
            if not r.get("contact_lastname") and nom:
                r["contact_lastname"] = nom

            # email
            if not r.get("email") and o.get("email"):
                r["email"] = o["email"]

            # phones
            if not r.get("phone") and o.get("phone"):
                r["phone"] = o["phone"]
            if not r.get("phone2") and o.get("mobile"):
                r["phone2"] = o["mobile"]

            # PJ
            attachments.append((safe_filename(f"carte_{r.get('name','societe')}_{file_id}.jpg"), img_bytes))

    # 4) Excel
    if mode == "session":
        xlsx_name = safe_filename(f"prospection_{agency}_{initials}_{run_date}.xlsx")
    elif mode == "agency_daily":
        xlsx_name = safe_filename(f"prospection_{agency}_CONSOLIDE_{run_date}.xlsx")
    else:
        xlsx_name = safe_filename(f"prospection_GLOBAL_{run_date}.xlsx")

    xlsx_path = os.path.join(os.getcwd(), xlsx_name)
    build_excel(prospects, xlsx_path)

    # Excel en PJ
    with open(xlsx_path, "rb") as f:
        excel_bytes = f.read()
    attachments.insert(0, (xlsx_name, excel_bytes))

    # 5) destinataires + mail
    if mode == "session":
        to_email = RECIPIENTS_SESSION.get((agency, initials))
        if not to_email:
            raise RuntimeError(f"No recipient mapped for ({agency},{initials}) in RECIPIENTS_SESSION")
        to_list = [to_email]
        subject = f"[PROSPECTION] {agency}/{initials} - {run_date}"
        stats_html = ""
        if visits_clients or visits_prospects or commandes:
            stats_html = f"""
            <p><b>Clôture session</b></p>
            <ul>
              <li>Visites clients: {visits_clients or "-"}</li>
              <li>Visites prospects: {visits_prospects or "-"}</li>
              <li>Commandes: {commandes or "-"}</li>
            </ul>
            """
        html = f"""
        <p>Bonjour,</p>
        <p>Voici l’export de prospection <b>{agency}/{initials}</b> du <b>{run_date}</b>.</p>
        {stats_html}
        <p>Pièces jointes : Excel + cartes de visite.</p>
        """
        send_mail_brevo(to_list, subject, html, attachments)

    elif mode == "agency_daily":
        to_email = RECIPIENTS_DAILY_AGENCY.get(agency)
        if not to_email:
            raise RuntimeError(f"No agency daily recipient mapped for {agency}")
        to_list = [to_email]
        subject = f"[PROSPECTION] Récap agence {agency} - {run_date}"
        html = f"""
        <p>Bonjour,</p>
        <p>Voici le récapitulatif consolidé de l’agence <b>{agency}</b> pour le <b>{run_date}</b>.</p>
        <p>Pièce jointe : Excel consolidé.</p>
        """
        send_mail_brevo(to_list, subject, html, attachments)

    else:  # global_daily
        to_list = RECIPIENTS_DAILY_GLOBAL
        subject = f"[PROSPECTION] Récap GLOBAL - {run_date}"
        html = f"""
        <p>Bonsoir,</p>
        <p>Voici le récapitulatif global pour le <b>{run_date}</b> (toutes agences / collaborateurs).</p>
        <p>Pièce jointe : Excel global.</p>
        """
        send_mail_brevo(to_list, subject, html, attachments)

    print(f"OK: mode={mode} date={run_date} records={len(prospects)} attachments={len(attachments)}")


if __name__ == "__main__":
    main()