# -*- coding: utf-8 -*-
"""Code.gs から Apps Script プロジェクトのインポート用JSONを作る。"""
import json
from pathlib import Path

d = Path(__file__).parent
code = (d / "Code.gs").read_text(encoding="utf-8")
manifest = {"timeZone": "Asia/Tokyo", "exceptionLogging": "STACKDRIVER",
            "runtimeVersion": "V8",
            "webapp": {"executeAs": "USER_DEPLOYING", "access": "ANYONE_ANONYMOUS"}}
payload = {"files": [
    {"name": "appsscript", "type": "JSON", "source": json.dumps(manifest)},
    {"name": "Code", "type": "SERVER_JS", "source": code},
]}
out = d / "import_payload.json"
out.write_text(json.dumps(payload, ensure_ascii=True), encoding="utf-8")
print(f"OK {out} ({out.stat().st_size} bytes)")
