# -*- coding: utf-8 -*-
"""Code.gs を Monaco の setValue 呼び出し1行に変換して出力する。"""
import json
from pathlib import Path

d = Path(__file__).parent
code = (d / "Code.gs").read_text(encoding="utf-8")
js = "monaco.editor.getModels()[0].setValue(" + json.dumps(code, ensure_ascii=True) + "); 'set:' + monaco.editor.getModels()[0].getValue().length"
(d / "setvalue.js").write_text(js, encoding="utf-8")
print(f"OK setvalue.js ({len(js)} chars)")
