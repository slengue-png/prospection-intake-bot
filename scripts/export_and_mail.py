"""
export_and_mail.py — Script principal export prospection v8.4
=============================================================
Modes :
  individual     : export commercial + manager + admin
  agency_manager : recap journalier par agence → managers
  admin          : recap toutes agences → SL

Pièces jointes :
  - Prospection_{date}_{ag}_{ini}.xlsx  (consultation)
  - Prospection_{date}_{ag}_{ini}.xls   (import CRM)
  - Photos_{date}_{ag}_{ini}.zip        (cartes + façades)
"""

import os, sys, json, datetime, base64, io, zipfile, re
from typing import Any, Dict, List, Optional, Tuple

import requests

from excel_export import (
    export_commercial, export_agency_manager, export_admin,
    build_email_attachments, get_xlsx_bytes,
)

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────

def _env(k, d=""): return os.environ.get(k, d).strip()

WORKER_BASE_URL    = _env("WORKER_BASE_URL").rstrip("/")
EXPORT_TOKEN       = _env("EXPORT_TOKEN")
TELEGRAM_TOKEN     = _env("TELEGRAM_TOKEN")
BREVO_API_KEY      = _env("BREVO_API_KEY")
BREVO_SENDER_EMAIL = _env("BREVO_SENDER_EMAIL")
BREVO_SENDER_NAME  = _env("BREVO_SENDER_NAME", "Prospection Bot")
MAIL_ROUTING_JSON  = _env("MAIL_ROUTING_JSON")
SEND_MODE          = _env("SEND_MODE", "individual").lower()
RUN_DATE           = _env("RUN_DATE") or datetime.date.today().strftime("%Y-%m-%d")
AGENCY             = _env("AGENCY").upper()
INITIALS           = _env("INITIALS").upper()

# OUT_DIR robuste
OUT_DIR = _env("OUT_DIR", "/tmp/prospection_out")
if os.path.exists(OUT_DIR) and not os.path.isdir(OUT_DIR):
    OUT_DIR = "/tmp/prospection_out"
os.makedirs(OUT_DIR, exist_ok=True)

# Routing avec fallback hardcodé
DEFAULT_ROUTING = {
    "agencies": {
        "GR":  {"manager": {"initials":"JL","email":"jennifer.laurens@ras-interim.fr"},
                "commercial": {"initials":"CZ","email":"celine.zunarelli@ras-interim.fr"}},
        "VR":  {"manager": {"initials":"JB","email":"jelena.carrasso@ras-interim.fr"},
                "commercial": {"initials":"LB","email":"laura.berthet@ras-interim.fr"}},
        "GRS": {"manager": {"initials":"PV","email":"pauline.vieira@ras-interim.fr"},
                "commercial": {"initials":"ST","email":"severine.thevenin@ras-interim.fr"}},
        "SLS": {"manager": {"initials":"AC","email":"aurelie.curt@ras-interim.fr"},
                "commercial": {"initials":"AC","email":"aurelie.curt@ras-interim.fr"}},
    },
    "users": {
        "JL":"jennifer.laurens@ras-interim.fr","CZ":"celine.zunarelli@ras-interim.fr",
        "JB":"jelena.carrasso@ras-interim.fr","LB":"laura.berthet@ras-interim.fr",
        "PV":"pauline.vieira@ras-interim.fr","ST":"severine.thevenin@ras-interim.fr",
        "AC":"aurelie.curt@ras-interim.fr","SL":"samuel.lengue@ras-interim.fr",
    },
    "admin": {"initials":"SL","email":"samuel.lengue@ras-interim.fr"},
}

try:
    ROUTING = json.loads(MAIL_ROUTING_JSON) if MAIL_ROUTING_JSON else {}
    if not ROUTING.get("agencies"): raise ValueError
except Exception:
    ROUTING = DEFAULT_ROUTING
    print("  ⚠️  MAIL_ROUTING_JSON invalide — routing hardcodé")

ADMIN_EMAIL = ROUTING.get("admin", {}).get("email", "samuel.lengue@ras-interim.fr")

# Cache images Telegram
_TG_CACHE: Dict[str, bytes] = {}


# ─────────────────────────────────────────────
# TELEGRAM
# ─────────────────────────────────────────────

def tg_download(file_id: str) -> bytes:
    """Télécharge une image Telegram via file_id (fraîche, pas d'expiration)."""
    if not file_id or not TELEGRAM_TOKEN:
        return b""
    if file_id in _TG_CACHE:
        return _TG_CACHE[file_id]
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile",
            json={"file_id": file_id}, timeout=20,
        )
        fp = r.json().get("result", {}).get("file_path", "")
        if not fp:
            return b""
        img = requests.get(
            f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{fp}",
            timeout=60,
        ).content
        _TG_CACHE[file_id] = img
        return img
    except Exception as e:
        print(f"  ⚠️  tg_download({file_id[:20]}...): {e}")
        return b""


# ─────────────────────────────────────────────
# FETCH PROSPECTS
# ─────────────────────────────────────────────

def fetch_prospects(date="", agency="", initials="") -> List[Dict]:
    params = {"date": date} if date else {}
    try:
        r = requests.get(
            f"{WORKER_BASE_URL}/dump",
            params=params,
            headers={"X-Export-Token": EXPORT_TOKEN},
            timeout=60,
        )
        r.raise_for_status()
        out = []
        for line in r.text.strip().splitlines():
            line = line.strip()
            if not line: continue
            try:
                obj = json.loads(line)
                if isinstance(obj, dict): out.append(obj)
            except: pass
        print(f"  → {len(out)} prospect(s) bruts")
        return out
    except Exception as e:
        print(f"  ✗ fetch: {e}")
        return []


def normalize_prospect(p: Dict) -> Dict:
    d = p.get("draft", p)
    return {
        "name":               d.get("name", ""),
        "nom_commercial":     d.get("nom_commercial", ""),
        "address":            d.get("address", ""),
        "postal_code":        d.get("postal_code", ""),
        "city":               d.get("city", ""),
        "phone":              d.get("phone", ""),
        "phone2":             d.get("phone2", ""),
        "email":              d.get("email", ""),
        "email_card":         d.get("email_card", ""),
        "titre":              d.get("titre", ""),
        "siret":              d.get("siret", ""),
        "naf":                d.get("naf", ""),
        "activity_summary":   d.get("secteur_activite") or d.get("activity_summary") or "",
        "effectif":           d.get("effectif", ""),
        "date_creation":      d.get("date_creation", ""),
        "capital":            d.get("capital", ""),
        "website":            d.get("website", ""),
        "gps_lat":            d.get("gps_lat", ""),
        "gps_lon":            d.get("gps_lon", ""),
        "contact_civility":   d.get("contact_civility", ""),
        "interlocuteur":      d.get("interlocuteur", ""),
        "dirigeant":          d.get("dirigeant", ""),
        "resume":             d.get("resume", ""),
        "commande":           d.get("commande", ""),
        "qualification":      d.get("qualification", ""),
        # Photos — file_id pour le ZIP, url = placeholder
        "card_photo_file_id":   d.get("card_photo_file_id", ""),
        "card_photo_url":       "voir_zip" if d.get("card_photo_file_id") else "",
        "facade_photo_file_id": d.get("facade_photo_file_id", ""),
        "facade_photo_url":     "voir_zip" if d.get("facade_photo_file_id") else "",
        "agency":    p.get("agency", d.get("agency", "")),
        "initials":  p.get("initials", d.get("initials", "")),
        "_confidence": int(p.get("_confidence", d.get("_confidence", 0)) or 0),
    }


def filter_rows(prospects_raw, agency="", initials="") -> List[Dict]:
    rows = []
    seen_visit_ids = set()
    seen_sirets = set()

    for p in prospects_raw:
        ag  = (p.get("agency") or p.get("draft", {}).get("agency", "")).upper()
        ini = (p.get("initials") or p.get("draft", {}).get("initials", "")).upper()
        if agency  and ag  != agency.upper():  continue
        if initials and ini != initials.upper(): continue

        # Dédupliquer par visit_id
        vid = p.get("visit_id") or p.get("draft", {}).get("visit_id", "")
        if vid and vid in seen_visit_ids:
            print(f"  ⏭ doublon visit_id: {vid[:30]}")
            continue
        if vid: seen_visit_ids.add(vid)

        # Dédupliquer par SIRET si pas de visit_id
        row = normalize_prospect(p)
        siret = row.get("siret", "").strip()
        name  = row.get("name", "").strip()
        dedup_key = f"{siret}|{name}" if siret else ""
        if dedup_key and dedup_key in seen_sirets:
            print(f"  ⏭ doublon SIRET: {siret} — {name}")
            continue
        if dedup_key: seen_sirets.add(dedup_key)

        rows.append(row)
    return rows


# ─────────────────────────────────────────────
# ZIP PHOTOS
# ─────────────────────────────────────────────

def _safe_name(name: str) -> str:
    """Nom de fichier sûr depuis le nom de société."""
    s = re.sub(r"[^\w\s-]", "", (name or "INCONNU").upper())
    s = re.sub(r"\s+", "_", s.strip())
    return s[:40] or "INCONNU"


def build_photos_zip(rows: List[Dict], date: str, agency: str, initials: str) -> Optional[Tuple[str, bytes]]:
    """Construit un ZIP avec toutes les cartes de visite et photos de façade."""
    buf = io.BytesIO()
    count = 0
    seen_fids = set()  # Éviter les doublons par file_id

    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as z:
        index_lines = ["société;fichier;type"]

        for row in rows:
            safe = _safe_name(row.get("name", ""))

            # Carte de visite
            card_fid = (row.get("card_photo_file_id") or "").strip()
            if card_fid and card_fid not in seen_fids:
                seen_fids.add(card_fid)
                img = tg_download(card_fid)
                if img:
                    fname = f"cartes/{safe}_carte.jpg"
                    z.writestr(fname, img)
                    index_lines.append(f"{row.get('name','')}; {fname}; carte")
                    count += 1

            # Photo façade
            facade_fid = (row.get("facade_photo_file_id") or "").strip()
            if facade_fid and facade_fid not in seen_fids:
                seen_fids.add(facade_fid)
                img = tg_download(facade_fid)
                if img:
                    fname = f"facades/{safe}_facade.jpg"
                    z.writestr(fname, img)
                    index_lines.append(f"{row.get('name','')}; {fname}; facade")
                    count += 1

        if count == 0:
            return None

        z.writestr("INDEX.csv", "\n".join(index_lines).encode("utf-8"))

    zip_name = f"Photos_{date}_{agency}_{initials}.zip"
    print(f"  📷 ZIP photos: {count} image(s)")
    return zip_name, buf.getvalue()


# ─────────────────────────────────────────────
# EMAIL BREVO
# ─────────────────────────────────────────────

def send_brevo(to_email: str, to_name: str, subject: str, html: str,
               attachments: List = None) -> bool:
    if not BREVO_API_KEY or not to_email:
        print(f"  ✗ Email vide ou Brevo non configuré")
        return False
    payload = {
        "sender":      {"email": BREVO_SENDER_EMAIL, "name": BREVO_SENDER_NAME},
        "to":          [{"email": to_email, "name": to_name}],
        "subject":     subject,
        "htmlContent": html,
    }
    if attachments:
        payload["attachment"] = [
            {"name": name, "content": base64.b64encode(content).decode()}
            for name, content in attachments
        ]
    try:
        r = requests.post(
            "https://api.brevo.com/v3/smtp/email",
            headers={"api-key": BREVO_API_KEY, "Content-Type": "application/json"},
            json=payload, timeout=60,
        )
        if r.status_code in (200, 201, 202):
            print(f"  ✓ → {to_email}")
            return True
        print(f"  ✗ Brevo {r.status_code}: {r.text[:200]}")
        return False
    except Exception as e:
        print(f"  ✗ Brevo: {e}")
        return False


def html_body(count: int, title: str, subtitle: str) -> str:
    return f"""<html><body style="font-family:Arial,sans-serif;color:#333;max-width:600px">
    <h2 style="color:#2C3E50;border-bottom:2px solid #2C3E50;padding-bottom:8px">{title}</h2>
    <p>{subtitle}</p>
    <p><b>{count} fiche(s)</b> dans les fichiers joints.</p>
    <p style="margin-top:20px">Fichiers joints :<br>
    &nbsp;• <b>.xlsx</b> — consultation avec mise en page<br>
    &nbsp;• <b>.xls</b> — import CRM<br>
    &nbsp;• <b>.zip</b> — photos cartes de visite et façades (si disponibles)</p>
    <hr style="border:none;border-top:1px solid #eee;margin-top:30px">
    <p style="color:#aaa;font-size:11px">Généré par Prospection Bot v8.4 — {datetime.date.today()}</p>
    </body></html>"""


def make_attachments(path_xlsx: str, path_xls: str,
                     zip_data: Optional[Tuple[str, bytes]] = None,
                     xls_only: bool = False, xlsx_only: bool = False) -> List:
    att = []
    if not xls_only and os.path.exists(path_xlsx):
        att.append((os.path.basename(path_xlsx), get_xlsx_bytes(path_xlsx)))
    if not xlsx_only and os.path.exists(path_xls):
        att.append((os.path.basename(path_xls), get_xlsx_bytes(path_xls)))
    if zip_data:
        att.append(zip_data)
    return att


# ─────────────────────────────────────────────
# MODE INDIVIDUAL
# ─────────────────────────────────────────────

def run_individual():
    print(f"\n=== INDIVIDUAL — {AGENCY}/{INITIALS} — {RUN_DATE} ===")
    if not AGENCY or not INITIALS:
        print("✗ AGENCY et INITIALS requis"); sys.exit(1)

    prospects_raw = fetch_prospects(date=RUN_DATE)
    rows = filter_rows(prospects_raw, AGENCY, INITIALS)
    print(f"  → {len(rows)} fiche(s) {AGENCY}/{INITIALS}")
    if not rows: return

    path_xlsx, path_xls = export_commercial(RUN_DATE, AGENCY, INITIALS, rows, OUT_DIR)
    zip_data = build_photos_zip(rows, RUN_DATE, AGENCY, INITIALS)

    ag_cfg     = ROUTING.get("agencies", {}).get(AGENCY, {})
    commercial = ag_cfg.get("commercial", {})
    manager    = ag_cfg.get("manager", {})
    subject    = f"Prospection {RUN_DATE} — {AGENCY}/{INITIALS} ({len(rows)} fiche(s))"
    html       = html_body(len(rows), f"Prospection {RUN_DATE} — {AGENCY}/{INITIALS}",
                           f"Export de prospection du <b>{RUN_DATE}</b> pour <b>{INITIALS}</b> ({AGENCY}).")

    sent = set()

    # Commercial → xlsx + xls + zip
    if commercial.get("email"):
        att = make_attachments(path_xlsx, path_xls, zip_data)
        send_brevo(commercial["email"], f"Commercial {INITIALS}", subject, html, att)
        sent.add(commercial["email"])

    # Manager → xlsx + xls + zip
    if manager.get("email") and manager["email"] not in sent:
        att = make_attachments(path_xlsx, path_xls, zip_data)
        send_brevo(manager["email"], f"Manager {AGENCY}", subject, html, att)
        sent.add(manager["email"])

    # Admin SL → xlsx + xls + zip (toujours)
    if ADMIN_EMAIL and ADMIN_EMAIL not in sent:
        att = make_attachments(path_xlsx, path_xls, zip_data)
        send_brevo(ADMIN_EMAIL, "Admin SL", subject, html, att)
        sent.add(ADMIN_EMAIL)

    print(f"✅ Individual terminé — {len(sent)} destinataire(s)")


# ─────────────────────────────────────────────
# MODE AGENCY_MANAGER
# ─────────────────────────────────────────────

def run_agency_manager():
    print(f"\n=== AGENCY_MANAGER — {RUN_DATE} ===")
    prospects_raw = fetch_prospects(date=RUN_DATE)

    for agency in ["GR", "VR", "GRS", "SLS"]:
        rows = filter_rows(prospects_raw, agency=agency)
        print(f"\n  Agence {agency}: {len(rows)} fiche(s)")
        if not rows: continue

        path_xlsx, path_xls = export_agency_manager(RUN_DATE, agency, rows, OUT_DIR)
        zip_data = build_photos_zip(rows, RUN_DATE, agency, "AGENCE")

        manager = ROUTING.get("agencies", {}).get(agency, {}).get("manager", {})
        subject = f"Récap {RUN_DATE} — Agence {agency} ({len(rows)} fiche(s))"
        html    = html_body(len(rows), f"Récap Agence {agency} — {RUN_DATE}",
                            f"Récapitulatif journalier de l'agence <b>{agency}</b>.")

        # Manager → xls uniquement + zip
        if manager.get("email"):
            att = make_attachments(path_xlsx, path_xls, zip_data, xls_only=True)
            send_brevo(manager["email"], f"Manager {agency}", subject, html, att)

    print(f"\n✅ Agency_manager terminé")


# ─────────────────────────────────────────────
# MODE ADMIN
# ─────────────────────────────────────────────

def run_admin():
    print(f"\n=== ADMIN — {RUN_DATE} ===")
    prospects_raw = fetch_prospects(date=RUN_DATE)
    rows = [normalize_prospect(p) for p in prospects_raw]
    print(f"  → {len(rows)} fiche(s) toutes agences")
    if not rows: return

    path_xlsx, path_xls = export_admin(RUN_DATE, rows, OUT_DIR)
    zip_data = build_photos_zip(rows, RUN_DATE, "ADMIN", "ALL")

    subject = f"Récap admin {RUN_DATE} — Toutes agences ({len(rows)} fiche(s))"
    html    = html_body(len(rows), f"Récap Admin — {RUN_DATE}",
                        "Récapitulatif complet toutes agences.")

    # Admin SL → xlsx uniquement + zip
    if ADMIN_EMAIL:
        att = make_attachments(path_xlsx, path_xls, zip_data, xlsx_only=True)
        send_brevo(ADMIN_EMAIL, "Admin SL", subject, html, att)

    print(f"✅ Admin terminé")


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

if __name__ == "__main__":
    print(f"Prospection Export v8.4 — mode={SEND_MODE} date={RUN_DATE} agency={AGENCY} initials={INITIALS}")
    print(f"  WORKER_BASE_URL: {'OK' if WORKER_BASE_URL else 'MANQUANT'}")
    print(f"  EXPORT_TOKEN: {'OK' if EXPORT_TOKEN else 'MANQUANT'}")
    print(f"  BREVO_API_KEY: {'OK' if BREVO_API_KEY else 'MANQUANT'}")
    print(f"  MAIL_ROUTING: agencies={list(ROUTING.get('agencies',{}).keys())} admin={ADMIN_EMAIL}")
    if not WORKER_BASE_URL: print("✗ WORKER_BASE_URL manquant"); sys.exit(1)
    if not EXPORT_TOKEN:    print("✗ EXPORT_TOKEN manquant");    sys.exit(1)

    if   SEND_MODE == "individual":     run_individual()
    elif SEND_MODE == "agency_manager": run_agency_manager()
    elif SEND_MODE == "admin":          run_admin()
    else: print(f"✗ SEND_MODE inconnu: {SEND_MODE}"); sys.exit(1)
