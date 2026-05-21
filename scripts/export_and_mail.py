DIS DONC ON GALERE : Run python scripts/export_and_mail.py
  python scripts/export_and_mail.py
shell: /usr/bin/bash -e ***0***
env:
  pythonLocation: /opt/hostedtoolcache/Python/3.11.15/x64
  PKG_CONFIG_PATH: /opt/hostedtoolcache/Python/3.11.15/x64/lib/pkgconfig
  Python_ROOT_DIR: /opt/hostedtoolcache/Python/3.11.15/x64
  Python2_ROOT_DIR: /opt/hostedtoolcache/Python/3.11.15/x64
  Python3_ROOT_DIR: /opt/hostedtoolcache/Python/3.11.15/x64
  LD_LIBRARY_PATH: /opt/hostedtoolcache/Python/3.11.15/x64/lib
  WORKER_BASE_URL: ***
  EXPORT_TOKEN: ***
  TELEGRAM_TOKEN: ***
  BREVO_API_KEY: ***
  BREVO_SENDER_EMAIL: ***
  BREVO_SENDER_NAME: ***
  MAIL_ROUTING_JSON:
***
  GEMINI_API_KEY: ***
  GOOGLE_PLACES_API_KEY: ***
  SEND_MODE: individual
  RUN_DATE: 2026-05-21
  AGENCY: GRS
  INITIALS: SL
  File "/home/runner/work/prospection-intake-bot/prospection-intake-bot/scripts/export_and_mail.py", line 272
    if (p.get("agency") or "").upper() == agency.upper()]        path_xlsx, path_xls = export_agency_manager(RUN_DATE, agency, rows, OUT_DIR)
                                                                 ^^^^^^^^^
SyntaxError: invalid syntax
Error: Process completed with exit code 1.
