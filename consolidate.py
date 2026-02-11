import os, re, json, subprocess
import pandas as pd
import yaml
from pathlib import Path

OUT_DIR = Path("out")
OUT_DIR.mkdir(exist_ok=True)

def gh_api(path):
    token = os.environ.get("GH_TOKEN") or os.environ.get("GITHUB_TOKEN")
    if not token:
        raise RuntimeError("Missing GH_TOKEN")
    repo = os.environ.get("REPO")
    url = f"https://api.github.com/repos/{repo}/{path.lstrip('/')}"
    cmd = f'curl -sS -H "Authorization: Bearer {token}" -H "Accept: application/vnd.github+json" "{url}"'
    res = subprocess.check_output(["bash", "-lc", cmd])
    return json.loads(res)

def extract_yaml(issue_body: str):
    m = re.search(r"```yaml\s*(.*?)\s*```", issue_body, re.S)
    if not m:
        return None
    raw = m.group(1)
    return yaml.safe_load(raw)

def main():
    issues = []
    page = 1
    while True:
        batch = gh_api(f"issues?state=all&labels=intake&per_page=100&page={page}")
        if not batch:
            break
        issues.extend(batch)
        page += 1

    rows = []
    for it in issues:
        y = extract_yaml(it.get("body", "") or "")
        if not isinstance(y, dict):
            continue
        enriched = y.get("enriched") or {}
        rows.append({
            "NOM": y.get("name", "") or "",
            "ADRESSE": enriched.get("address", "") or "",
            "CODE POSTAL": enriched.get("postalCode", "") or "",
            "VILLE": enriched.get("city", "") or (y.get("city", "") or ""),
            "TELEPHONE": enriched.get("phone1", "") or (y.get("phone", "") or ""),
            "TELEPHONE 2": enriched.get("phone2", "") or "",
            "MAIL": enriched.get("email", "") or "",
            "SIRET": enriched.get("siret", "") or "",
            "NAF": enriched.get("naf", "") or "",
            "SITE WEB": enriched.get("website", "") or "",
            "Contact: civilité": "",
            "Contact : prénom": "",
            "Contact : nom": "",
            "RESUME ENTRETIEN": (y.get("resume_entretien") or "").strip(),
            "COMMANDE": (y.get("commande") or "").strip(),
        })

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(by=["NOM", "VILLE"], kind="stable")

    csv_path = OUT_DIR / "prospection_reception.csv"
    xlsx_path = OUT_DIR / "prospection_reception.xlsx"
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="RECEPTION")
        ws = writer.sheets["RECEPTION"]
        widths = {"A": 40, "B": 45, "C": 12, "D": 22, "E": 18, "F": 18, "G": 30,
                  "H": 18, "I": 10, "J": 30, "N": 60, "O": 40}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

    print(f"Wrote {csv_path} and {xlsx_path} with {len(df)} rows")

if __name__ == "__main__":
    main()
