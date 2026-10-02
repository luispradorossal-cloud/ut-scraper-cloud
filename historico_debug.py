from pathlib import Path
from scraper import create_session, BASE_URL, DOWNLOAD_PARAMS_BASE, PORTLET_ID
s = create_session(); s.get(BASE_URL, timeout=30)
out = Path("historico/debug"); out.mkdir(parents=True, exist_ok=True)
tests = [("Prog_Diaria_Inicial_Aislado_200626.xlsx","2026"),("Prog_Diaria200626.xlsx","2026"),
         ("Prog_Diaria_Inicial_Aislado_010126.xlsx","2026"),("Prog_Diaria010126.xlsx","2026"),
         ("Prog_Diaria_Inicial_Aislado_150925.xlsx","2025"),("Prog_Diaria150925.xlsx","2025"),
         ("Prog_Diaria_Inicial_Aislado_150924.xlsx","2024"),("Prog_Diaria150924.xlsx","2024")]
lines=[]
for f,y in tests:
    p = DOWNLOAD_PARAMS_BASE.copy(); p["p_p_resource_id"]=f; p[f"_{PORTLET_ID}_folder"]=y
    r = s.get(BASE_URL+"?"+"&".join(f"{k}={v}" for k,v in p.items()), timeout=60)
    lines.append(f"{f},{y},{r.status_code},{r.headers.get('Content-Type','')},{len(r.content)},{r.content[:2]!r}")
(out/"probe.csv").write_text("\n".join(lines))
print("\n".join(lines))
