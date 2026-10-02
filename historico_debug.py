from pathlib import Path
from scraper import create_session, BASE_URL, LIST_PARAMS
s = create_session(); s.get(BASE_URL, timeout=30)
out = Path("historico/debug"); out.mkdir(parents=True, exist_ok=True)
(out / "base.html").write_text(s.get(BASE_URL, timeout=30).text)
for y in ["2024", "2025", "2026"]:
    r = s.post(BASE_URL, params=LIST_PARAMS, data={"folder": y}, timeout=60)
    (out / f"list_{y}.html").write_text(r.text)
print("ok")
