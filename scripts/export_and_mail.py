"""
excel_export.py — Module export Excel prospection
==================================================
Produit deux fichiers par destinataire :
  1. Prospection_{date}_{destinataire}.xlsx  — consultation (avec miniatures)
  2. Prospection_{date}_{destinataire}.xls   — import CRM (sans images, données brutes)

Colonnes CRM dans l'ordre exact :
  NOM | RUE | CODE POSTAL | VILLE | Téléphone | Téléphone (Portable) |
  Mail générique | SIRET | NAF | ACTIVITE | SITE WEB | PREFIXE |
  INTERLOCUTEUR | DIRIGEANT | RESUME ENTRETIEN | COMMANDE |
  CARTE DE VISITE | AGENCE | INITIALS

Couleurs agence (colonne AGENCE) :
  GR  → Orange  #FF8C00
  VR  → Vert    #228B22
  GRS → Jaune   #FFD700
  SLS → Bleu    #1E90FF

Score de confiance (fond de ligne) :
  ≥ 60 → blanc  (fiche complète)
  30-59 → jaune très pâle  #FFFDE7
  < 30  → rouge très pâle  #FFF0F0
"""

import io
import os
import re
import datetime as zoneinfo
from typing import Any, Dict, List, Optional, Tuple

import requests
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────

CRM_HEADERS = [
    "NOM",
    "RUE",
    "CODE POSTAL",
    "VILLE",
    "Téléphone",
    "Téléphone (Portable)",
    "Mail générique",
    "SIRET",
    "NAF",
    "ACTIVITE",
    "EFFECTIF",
    "DATE CREATION",
    "CAPITAL",
    "SITE WEB",
    "PREFIXE",
    "INTERLOCUTEUR",
    "DIRIGEANT",
    "RESUME ENTRETIEN",
    "COMMANDE",
    "QUALIFICATION",
    "CARTE DE VISITE",
    "AGENCE",
    "INITIALS",
]

# Mapping interne → colonne CRM
FIELD_MAP = {
    "NOM":                   "name",
    "RUE":                   "address",
    "CODE POSTAL":           "postal_code",
    "VILLE":                 "city",
    "Téléphone":             "phone",
    "Téléphone (Portable)":  "phone2",
    "Mail générique":        "email",
    "SIRET":                 "siret",
    "NAF":                   "naf",
    "ACTIVITE":              "activity_summary",
    "EFFECTIF":              "effectif",
    "DATE CREATION":         "date_creation",
    "CAPITAL":               "capital",
    "SITE WEB":              "website",
    "PREFIXE":               "contact_civility",
    "INTERLOCUTEUR":         "interlocuteur",
    "DIRIGEANT":             "dirigeant",
    "RESUME ENTRETIEN":      "resume",
    "COMMANDE":              "commande",
    "QUALIFICATION":         "qualification",
    "CARTE DE VISITE":       "card_photo_url",
    "AGENCE":                "agency",
    "INITIALS":              "initials",
}

# Couleurs agence (fond cellule AGENCE)
AGENCY_COLORS = {
    "GR":  "FF8C00",   # Orange
    "VR":  "228B22",   # Vert
    "GRS": "FFD700",   # Jaune
    "SLS": "1E90FF",   # Bleu
}

AGENCY_TEXT_COLORS = {
    "GR":  "FFFFFF",
    "VR":  "FFFFFF",
    "GRS": "000000",
    "SLS": "FFFFFF",
}

# Fond de ligne selon score de confiance
SCORE_FILLS = {
    "high":   "FFFFFF",   # ≥ 60 : blanc
    "medium": "FFFDE7",   # 30-59 : jaune très pâle
    "low":    "FFF0F0",   # < 30  : rouge très pâle
}

# Hauteur de ligne avec miniature (px → pts Excel ≈ px * 0.75)
ROW_HEIGHT_WITH_IMG = 60
IMG_MAX_W = 90
IMG_MAX_H = 55

# Largeurs colonnes (caractères)
COL_WIDTHS = {
    "NOM":                  28,
    "RUE":                  30,
    "CODE POSTAL":           9,
    "VILLE":                18,
    "Téléphone":            14,
    "Téléphone (Portable)": 14,
    "Mail générique":       30,
    "SIRET":                16,
    "NAF":                   7,
    "ACTIVITE":             28,
    "EFFECTIF":             14,
    "DATE CREATION":        14,
    "CAPITAL":              12,
    "SITE WEB":             28,
    "PREFIXE":               8,
    "INTERLOCUTEUR":        22,
    "DIRIGEANT":            22,
    "RESUME ENTRETIEN":     40,
    "COMMANDE":             20,
    "QUALIFICATION":        12,
    "CARTE DE VISITE":      14,
    "AGENCE":                8,
    "INITIALS":              8,
}

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def paris_today() -> str:
    try:
        import zoneinfo as zi
        tz = zi.ZoneInfo("Europe/Paris")
        return datetime.datetime.now(tz).strftime("%Y-%m-%d")
    except Exception:
        return datetime.date.today().strftime("%Y-%m-%d")


def score_level(score: int) -> str:
    if score >= 60:
        return "high"
    if score >= 30:
        return "medium"
    return "low"


def thin_border() -> Border:
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def header_fill() -> PatternFill:
    return PatternFill("solid", start_color="2C3E50", end_color="2C3E50")


def header_font() -> Font:
    return Font(name="Arial", bold=True, color="FFFFFF", size=10)


def cell_font(bold: bool = False) -> Font:
    return Font(name="Arial", size=9, bold=bold)


def centered() -> Alignment:
    return Alignment(horizontal="center", vertical="center", wrap_text=False)


def left_aligned() -> Alignment:
    return Alignment(horizontal="left", vertical="center", wrap_text=False)


def wrap_aligned() -> Alignment:
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


def row_fill(score: int) -> PatternFill:
    color = SCORE_FILLS[score_level(score)]
    return PatternFill("solid", start_color=color, end_color=color)


def agency_fill(agency: str) -> PatternFill:
    color = AGENCY_COLORS.get(agency.upper(), "EEEEEE")
    return PatternFill("solid", start_color=color, end_color=color)


def agency_font(agency: str) -> Font:
    color = AGENCY_TEXT_COLORS.get(agency.upper(), "000000")
    return Font(name="Arial", bold=True, size=9, color=color)


# ─────────────────────────────────────────────
# TÉLÉCHARGEMENT + REDIMENSIONNEMENT IMAGE
# ─────────────────────────────────────────────

_img_cache: Dict[str, bytes] = {}


def fetch_image_bytes(url: str) -> bytes:
    """Télécharge une image depuis une URL Telegram ou HTTP."""
    if not url:
        return b""
    if url in _img_cache:
        return _img_cache[url]
    try:
        headers = {"User-Agent": "ProspectionBot/8.2"}
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code == 200:
            _img_cache[url] = r.content
            return r.content
    except Exception:
        pass
    return b""


def fetch_image_from_file_id(file_id: str) -> bytes:
    """Télécharge une image depuis un file_id Telegram."""
    if not file_id or not TELEGRAM_TOKEN:
        return b""
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile",
            json={"file_id": file_id},
            timeout=15,
        )
        j = r.json()
        if not j.get("ok"):
            return b""
        fp = j["result"]["file_path"]
        return fetch_image_bytes(
            f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{fp}"
        )
    except Exception:
        return b""


def resize_image_bytes(img_bytes: bytes, max_w: int, max_h: int) -> Optional[bytes]:
    """Redimensionne une image en conservant le ratio, retourne des bytes PNG."""
    if not img_bytes:
        return None
    try:
        img = PILImage.open(io.BytesIO(img_bytes)).convert("RGB")
        img.thumbnail((max_w, max_h), PILImage.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
    except Exception:
        return None


def build_xl_image(img_bytes: bytes) -> Optional[XLImage]:
    """Crée un objet openpyxl Image depuis des bytes."""
    resized = resize_image_bytes(img_bytes, IMG_MAX_W, IMG_MAX_H)
    if not resized:
        return None
    try:
        xl_img = XLImage(io.BytesIO(resized))
        xl_img.width = IMG_MAX_W
        xl_img.height = IMG_MAX_H
        return xl_img
    except Exception:
        return None


def get_card_image(row: Dict[str, Any]) -> bytes:
    """
    Récupère l'image de carte de visite d'une fiche.
    Priorité : card_photo_url > card_photo_file_id > facade_photo_url > facade_photo_file_id
    """
    for url_key in ["card_photo_url", "facade_photo_url"]:
        url = (row.get(url_key) or "").strip()
        if url:
            img = fetch_image_bytes(url)
            if img:
                return img

    for fid_key in ["card_photo_file_id", "facade_photo_file_id"]:
        fid = (row.get(fid_key) or "").strip()
        if fid:
            img = fetch_image_from_file_id(fid)
            if img:
                return img

    return b""


# ─────────────────────────────────────────────
# CONSTRUCTION D'UN ONGLET
# ─────────────────────────────────────────────

def build_sheet(
    ws,
    rows: List[Dict[str, Any]],
    with_images: bool = True,
    sheet_title: str = "DATA",
):
    """
    Remplit un onglet openpyxl avec les données de prospection.
    with_images=True  → version consultation (xlsx)
    with_images=False → version import CRM (xls)
    """
    ws.title = sheet_title

    # ── En-têtes ──────────────────────────────
    for col_idx, header in enumerate(CRM_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font()
        cell.fill = header_fill()
        cell.alignment = centered()
        cell.border = thin_border()

    # Figer la première ligne + première colonne
    ws.freeze_panes = "B2"

    # Filtres automatiques
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CRM_HEADERS))}1"

    # ── Largeurs colonnes ──────────────────────
    for col_idx, header in enumerate(CRM_HEADERS, start=1):
        width = COL_WIDTHS.get(header, 15)
        # Colonne image plus large si on affiche les photos
        if header == "CARTE DE VISITE" and with_images:
            width = max(width, int(IMG_MAX_W / 6) + 2)
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Index colonnes utiles
    col_agence_idx = CRM_HEADERS.index("AGENCE") + 1
    col_card_idx   = CRM_HEADERS.index("CARTE DE VISITE") + 1
    col_qualif_idx = CRM_HEADERS.index("QUALIFICATION") + 1 if "QUALIFICATION" in CRM_HEADERS else None

    # ── Données ───────────────────────────────
    for row_idx, row in enumerate(rows, start=2):
        score  = int(row.get("_confidence", 0) or row.get("_score", 0))
        agency = (row.get("agency") or "").upper().strip()
        r_fill = row_fill(score)

        # Hauteur de ligne
        if with_images:
            ws.row_dimensions[row_idx].height = ROW_HEIGHT_WITH_IMG

        for col_idx, header in enumerate(CRM_HEADERS, start=1):
            field = FIELD_MAP.get(header, "")
            value = str(row.get(field) or "").strip()

            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border()

            # Colonne CARTE DE VISITE : image ou lien
            if header == "CARTE DE VISITE":
                if with_images:
                    # On insère l'image plus bas (après la boucle colonnes)
                    cell.value = ""
                    cell.fill = r_fill
                else:
                    # Version CRM : URL texte
                    cell.value = value
                    cell.font = cell_font()
                    cell.alignment = left_aligned()
                    cell.fill = r_fill
                continue

            # Colonne QUALIFICATION : couleur selon valeur
            if header == "QUALIFICATION":
                qualif = value.lower().strip()
                if qualif == "chaud":
                    q_fill = PatternFill("solid", fgColor="FFD0D0")  # rouge pâle
                    q_val  = "🔥 Chaud"
                elif qualif == "tiede":
                    q_fill = PatternFill("solid", fgColor="FFF3CD")  # jaune pâle
                    q_val  = "📅 Tiède"
                elif qualif == "froid":
                    q_fill = PatternFill("solid", fgColor="D0E8FF")  # bleu pâle
                    q_val  = "❄️ Froid"
                else:
                    q_fill = r_fill
                    q_val  = value
                cell.value = q_val
                cell.font = cell_font(bold=qualif in ("chaud", "froid"))
                cell.fill = q_fill
                cell.alignment = centered()
                continue

            # Colonne AGENCE : couleur dédiée
            if header == "AGENCE":
                cell.value = agency
                cell.font = agency_font(agency)
                cell.fill = agency_fill(agency)
                cell.alignment = centered()
                continue

            # Colonnes texte long : wrap
            if header in ("RESUME ENTRETIEN", "COMMANDE"):
                cell.value = value
                cell.font = cell_font()
                cell.alignment = wrap_aligned()
                cell.fill = r_fill
                continue

            # Toutes les autres colonnes
            cell.value = value
            cell.font = cell_font()
            cell.alignment = left_aligned()
            cell.fill = r_fill

        # ── Image carte de visite ──────────────
        if with_images:
            img_bytes = get_card_image(row)
            if img_bytes:
                xl_img = build_xl_image(img_bytes)
                if xl_img:
                    cell_addr = f"{get_column_letter(col_card_idx)}{row_idx}"
                    ws.add_image(xl_img, cell_addr)
            else:
                # Pas d'image : afficher l'URL en texte
                url = str(row.get("card_photo_url") or "").strip()
                ws.cell(row=row_idx, column=col_card_idx).value = url or "—"
                ws.cell(row=row_idx, column=col_card_idx).font = cell_font()
                ws.cell(row=row_idx, column=col_card_idx).alignment = left_aligned()

    # ── Ligne de total (nombre de fiches) ─────
    total_row = len(rows) + 2
    ws.cell(row=total_row, column=1, value=f"Total : {len(rows)} fiche(s)")
    ws.cell(row=total_row, column=1).font = Font(name="Arial", bold=True, size=9, italic=True)


# ─────────────────────────────────────────────
# CONSTRUCTION WORKBOOK
# ─────────────────────────────────────────────

def build_workbook_commercial(
    rows: List[Dict[str, Any]],
    with_images: bool,
) -> Workbook:
    """1 onglet DATA pour le commercial."""
    wb = Workbook()
    ws = wb.active
    build_sheet(ws, rows, with_images=with_images, sheet_title="DATA")
    return wb


def build_workbook_agency(
    rows_by_initials: Dict[str, List[Dict[str, Any]]],
    agency: str,
    with_images: bool,
) -> Workbook:
    """
    1 onglet unique avec toutes les fiches de l'agence.
    Trié par initials puis score décroissant.
    """
    wb = Workbook()
    ws = wb.active

    all_rows = []
    for initials in sorted(rows_by_initials.keys()):
        all_rows.extend(rows_by_initials[initials])

    # Tri : score décroissant
    all_rows.sort(key=lambda r: int(r.get("_confidence", 0) or 0), reverse=True)

    build_sheet(ws, all_rows, with_images=with_images, sheet_title=agency)
    return wb


def build_workbook_admin(
    rows_by_agency: Dict[str, List[Dict[str, Any]]],
    with_images: bool,
) -> Workbook:
    """
    1 onglet par agence pour le manager général.
    Ordre des onglets : GR, VR, GRS, SLS.
    """
    wb = Workbook()
    first = True

    for agency in ["GR", "VR", "GRS", "SLS"]:
        rows = rows_by_agency.get(agency, [])
        if not rows:
            continue

        rows_sorted = sorted(
            rows,
            key=lambda r: int(r.get("_confidence", 0) or 0),
            reverse=True,
        )

        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet()

        build_sheet(ws, rows_sorted, with_images=with_images, sheet_title=agency)

    return wb


# ─────────────────────────────────────────────
# SAUVEGARDE DOUBLE FORMAT
# ─────────────────────────────────────────────

def save_pair(
    wb_consultation: Workbook,
    wb_crm: Workbook,
    out_dir: str,
    base_name: str,
) -> Tuple[str, str]:
    """
    Sauvegarde deux fichiers :
      - base_name.xlsx  (consultation avec images)
      - base_name.xls   (import CRM sans images, format xlsx renommé)

    Retourne (path_xlsx, path_xls).

    Note sur le .xls : les CRM modernes acceptent le format Office Open XML
    même avec l'extension .xls. Si le CRM est strict (Excel 97-2003 binaire),
    il faudra passer par LibreOffice en CLI : `libreoffice --headless
    --convert-to xls base_name.xlsx` — la fonction le fait automatiquement
    si LibreOffice est disponible.
    """
    os.makedirs(out_dir, exist_ok=True)

    path_xlsx = os.path.join(out_dir, f"{base_name}.xlsx")
    path_xls  = os.path.join(out_dir, f"{base_name}.xls")

    wb_consultation.save(path_xlsx)

    # Tenter une conversion LibreOffice vers vrai .xls
    crm_xlsx_tmp = os.path.join(out_dir, f"{base_name}_CRM_tmp.xlsx")
    wb_crm.save(crm_xlsx_tmp)

    converted = _try_libreoffice_convert(crm_xlsx_tmp, out_dir, base_name)

    if not converted:
        # Fallback : renommer le xlsx en .xls (compatible avec la plupart des CRM)
        import shutil
        shutil.copy(crm_xlsx_tmp, path_xls)

    try:
        os.remove(crm_xlsx_tmp)
    except Exception:
        pass

    return path_xlsx, path_xls


def _try_libreoffice_convert(xlsx_path: str, out_dir: str, base_name: str) -> bool:
    """
    Tente de convertir un .xlsx en vrai binaire .xls via LibreOffice headless.
    Retourne True si la conversion a réussi.
    """
    import subprocess
    import shutil

    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo:
        return False

    try:
        result = subprocess.run(
            [lo, "--headless", "--convert-to", "xls", "--outdir", out_dir, xlsx_path],
            capture_output=True,
            timeout=60,
        )
        if result.returncode == 0:
            # LibreOffice nomme le fichier {stem}.xls dans out_dir
            stem = os.path.splitext(os.path.basename(xlsx_path))[0]
            lo_output = os.path.join(out_dir, f"{stem}.xls")
            target    = os.path.join(out_dir, f"{base_name}.xls")
            if os.path.exists(lo_output) and lo_output != target:
                os.rename(lo_output, target)
            return os.path.exists(target)
    except Exception:
        pass

    return False


# ─────────────────────────────────────────────
# FONCTIONS PUBLIQUES D'EXPORT
# ─────────────────────────────────────────────

def export_commercial(
    date: str,
    agency: str,
    initials: str,
    rows: List[Dict[str, Any]],
    out_dir: str = "out",
) -> Tuple[str, str]:
    """
    Génère les deux fichiers pour un commercial.
    Retourne (path_xlsx, path_xls).
    """
    base = f"Prospection_{date}_{agency}_{initials}"
    wb_cons = build_workbook_commercial(rows, with_images=True)
    wb_crm  = build_workbook_commercial(rows, with_images=False)
    return save_pair(wb_cons, wb_crm, out_dir, base)


def export_agency_manager(
    date: str,
    agency: str,
    rows: List[Dict[str, Any]],
    out_dir: str = "out",
) -> Tuple[str, str]:
    """
    Génère les deux fichiers pour le responsable d'agence.
    rows = toutes les fiches de l'agence (tous commerciaux).
    Retourne (path_xlsx, path_xls).
    """
    # Grouper par initials pour le tri
    by_initials: Dict[str, List] = {}
    for r in rows:
        ini = (r.get("initials") or "").upper()
        by_initials.setdefault(ini, []).append(r)

    base = f"Prospection_{date}_{agency}_MANAGER"
    wb_cons = build_workbook_agency(by_initials, agency, with_images=True)
    wb_crm  = build_workbook_agency(by_initials, agency, with_images=False)
    return save_pair(wb_cons, wb_crm, out_dir, base)


def export_admin(
    date: str,
    rows_all: List[Dict[str, Any]],
    out_dir: str = "out",
) -> Tuple[str, str]:
    """
    Génère les deux fichiers pour le manager général.
    1 onglet par agence.
    Retourne (path_xlsx, path_xls).
    """
    by_agency: Dict[str, List] = {}
    for r in rows_all:
        ag = (r.get("agency") or "").upper()
        by_agency.setdefault(ag, []).append(r)

    base = f"Prospection_{date}_ADMIN"
    wb_cons = build_workbook_admin(by_agency, with_images=True)
    wb_crm  = build_workbook_admin(by_agency, with_images=False)
    return save_pair(wb_cons, wb_crm, out_dir, base)


# ─────────────────────────────────────────────
# INTÉGRATION DANS export_and_mail.py
# ─────────────────────────────────────────────

def get_xlsx_bytes(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def build_email_attachments(
    path_xlsx: str,
    path_xls: str,
) -> List[Tuple[str, bytes]]:
    """
    Retourne la liste des pièces jointes pour Brevo :
    [(nom_fichier, bytes), ...]
    """
    attachments = []
    if os.path.exists(path_xlsx):
        attachments.append((os.path.basename(path_xlsx), get_xlsx_bytes(path_xlsx)))
    if os.path.exists(path_xls):
        attachments.append((os.path.basename(path_xls), get_xlsx_bytes(path_xls)))
    return attachments


# ─────────────────────────────────────────────
# TEST LOCAL
# ─────────────────────────────────────────────

def _make_test_rows() -> List[Dict[str, Any]]:
    """Données de test sans vraie API."""
    return [
        {
            "name": "MEANDRE CREATION",
            "address": "63 RUE DU MOIRON",
            "postal_code": "38420",
            "city": "LE VERSOUD",
            "phone": "0476523471",
            "phone2": "",
            "email": "prenom.nom@votremail.co",
            "siret": "45235898900021",
            "naf": "4332A",
            "activity_summary": "Menuiserie / Agencement",
            "website": "https://www.meandre-creation.com",
            "contact_civility": "M.",
            "interlocuteur": "MARC PAUL MAIRE",
            "dirigeant": "MARC PAUL MAIRE",
            "resume": "Fjddd",
            "commande": "1 menuisier",
            "card_photo_url": "",
            "card_photo_file_id": "",
            "facade_photo_url": "",
            "agency": "GR",
            "initials": "CZ",
            "_confidence": 72,
        },
        {
            "name": "CIMATEL",
            "address": "1 RUE YOURI GAGARINE",
            "postal_code": "38420",
            "city": "LE VERSOUD",
            "phone": "0476756781",
            "phone2": "",
            "email": "Durand@laposte.fr",
            "siret": "81165556200023",
            "naf": "3250A",
            "activity_summary": "Fabrication matériel médical",
            "website": "https://cimatel.fr/",
            "contact_civility": "",
            "interlocuteur": "MARC MOKARRAMI",
            "dirigeant": "M durand",
            "resume": "Bon bah j'ai vu l'entreprise",
            "commande": "1clitunier",
            "card_photo_url": "",
            "card_photo_file_id": "",
            "facade_photo_url": "",
            "agency": "VR",
            "initials": "LB",
            "_confidence": 45,
        },
        {
            "name": "DUPONT MENUISERIE",
            "address": "12 AV DES ALPES",
            "postal_code": "38000",
            "city": "GRENOBLE",
            "phone": "0476001122",
            "phone2": "0612345678",
            "email": "",
            "siret": "",
            "naf": "4332B",
            "activity_summary": "Menuiserie bois",
            "website": "",
            "contact_civility": "M.",
            "interlocuteur": "JEAN DUPONT",
            "dirigeant": "JEAN DUPONT",
            "resume": "Interesse par interim",
            "commande": "",
            "card_photo_url": "",
            "card_photo_file_id": "",
            "facade_photo_url": "",
            "agency": "GRS",
            "initials": "ST",
            "_confidence": 22,
        },
        {
            "name": "TRANSPORT SLS",
            "address": "5 ZI DES CHARTRONS",
            "postal_code": "69000",
            "city": "LYON",
            "phone": "0478901234",
            "phone2": "",
            "email": "contact@transport-sls.fr",
            "siret": "77123456700012",
            "naf": "4941A",
            "activity_summary": "Transport routier de fret",
            "website": "https://transport-sls.fr",
            "contact_civility": "Mme",
            "interlocuteur": "SOPHIE MARTIN",
            "dirigeant": "PIERRE MARTIN",
            "resume": "RDV semaine prochaine",
            "commande": "2 chauffeurs",
            "card_photo_url": "",
            "card_photo_file_id": "",
            "facade_photo_url": "",
            "agency": "SLS",
            "initials": "AC",
            "_confidence": 65,
        },
    ]


if __name__ == "__main__":
    import datetime
    date = datetime.date.today().strftime("%Y-%m-%d")
    rows = _make_test_rows()
    out  = "out_test"

    print("── Export commercial (CZ / GR) ──")
    gr_rows = [r for r in rows if r["agency"] == "GR"]
    p1, p2 = export_commercial(date, "GR", "CZ", gr_rows, out)
    print(f"  xlsx : {p1}")
    print(f"  xls  : {p2}")

    print("── Export manager agence GR ──")
    p3, p4 = export_agency_manager(date, "GR", gr_rows, out)
    print(f"  xlsx : {p3}")
    print(f"  xls  : {p4}")

    print("── Export admin (toutes agences) ──")
    p5, p6 = export_admin(date, rows, out)
    print(f"  xlsx : {p5}")
    print(f"  xls  : {p6}")

    print("\n✅ Tous les fichiers générés dans", out)
