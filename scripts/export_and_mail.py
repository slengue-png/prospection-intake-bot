def tg_api_get_file(token: str, file_id: str) -> Optional[str]:
    url = f"https://api.telegram.org/bot{token}/getFile"
    try:
        r = requests.get(url, params={"file_id": file_id}, timeout=20)
        # si Telegram renvoie 400/401/403, on veut voir le body
        if r.status_code != 200:
            print("[TG][getFile][HTTP]", r.status_code, r.text[:200])
            return None
        j = r.json()
        if not j.get("ok"):
            print("[TG][getFile][NOTOK]", str(j)[:200])
            return None
        return j["result"].get("file_path")
    except Exception as e:
        print("[TG][getFile][EXC]", str(e)[:160])
        return None

def tg_download_file(token: str, file_path: str) -> Optional[bytes]:
    url = f"https://api.telegram.org/file/bot{token}/{file_path}"
    try:
        r = requests.get(url, timeout=30)
        if r.status_code != 200:
            print("[TG][download][HTTP]", r.status_code, r.text[:200])
            return None
        return r.content
    except Exception as e:
        print("[TG][download][EXC]", str(e)[:160])
        return None