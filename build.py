"""Arma asientosPid.html (un solo archivo) a partir de src/."""
from pathlib import Path

ROOT = Path(__file__).parent
SRC = ROOT / "src"

html = (SRC / "template.html").read_text(encoding="utf-8")
html = html.replace("/*STYLES*/", (SRC / "styles.css").read_text(encoding="utf-8"))
html = html.replace("/*LOGIC*/", (SRC / "logic.js").read_text(encoding="utf-8"))
html = html.replace("/*APP*/", (SRC / "app.js").read_text(encoding="utf-8"))
(ROOT / "asientosPid.html").write_text(html, encoding="utf-8")
print("asientosPid.html generado")
