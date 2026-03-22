"""
Descarga los 3 archivos de una fecha especifica y los envia en UN solo correo.
- UT Aislado (Prog_Diaria_Inicial_Aislado_DDMMYY.xlsx)
- UT Prog Diaria (Prog_DiariaDDMMYY.xlsx)
- EOR Predespacho (PUB004-PRE-YYYYMMDD-OSO002.zip)

Uso:
    TARGET_DATE=2026-03-22 python descargar_fecha.py
"""

import base64
import email.encoders
import email.mime.base
import email.mime.multipart
import email.mime.text
import os
import smtplib
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# Zona horaria El Salvador (UTC-6)
TZ_SV = timezone(timedelta(hours=-6))


def now_sv():
    """Retorna la hora actual en zona horaria de El Salvador."""
    return datetime.now(TZ_SV)


# --- Config UT ----------------------------------------------------------------
UT_BASE_URL = "https://www.ut.com.sv/programacion-diaria1"
UT_PORTLET = "ProgramacionDiaria_WAR_PredespachoPublico_INSTANCE_Sw0UJdEgNCl7"

UT_LIST_PARAMS = {
    "p_p_id": UT_PORTLET,
    "p_p_lifecycle": "1",
    "p_p_state": "normal",
    "p_p_mode": "view",
    f"_{UT_PORTLET}_javax.portlet.action": "detalleArchivos",
}

UT_DL_PARAMS = {
    "p_p_id": UT_PORTLET,
    "p_p_lifecycle": "2",
    "p_p_state": "normal",
    "p_p_mode": "view",
    "p_p_cacheability": "cacheLevelPage",
    f"_{UT_PORTLET}_myaction": "detalles",
}

# --- Config EOR ---------------------------------------------------------------
EOR_BASE_URL = "https://www.enteoperador.org/"
EOR_PAGE_URL = (
    "https://www.enteoperador.org/mer/gestion-comercial/"
    "informes-publicos-de-procesos-comerciales/"
    "informe-de-procesos-comerciales-el-salvador/predespacho-sv/"
)
EOR_DL_PARAMS = {
    "red_fm_connect": "true",
    "front": "user",
    "fid": "117",
    "defaults": "0",
    "access_all": "0",
    "cmd": "file",
}


def create_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0.0.0",
        "Accept": "*/*",
    })
    return s


# --- UT Downloads -------------------------------------------------------------

def ut_download(session, filename, year, output_dir):
    filepath = output_dir / filename
    params = UT_DL_PARAMS.copy()
    params["p_p_resource_id"] = filename
    params[f"_{UT_PORTLET}_folder"] = year
    query = "&".join(f"{k}={v}" for k, v in params.items())
    url = f"{UT_BASE_URL}?{query}"

    try:
        session.headers["Referer"] = UT_BASE_URL
        r = session.get(url, timeout=60)
        r.raise_for_status()
        if "text/html" in r.headers.get("Content-Type", ""):
            print(f"  [ERROR] {filename}: respuesta HTML")
            return None
        with open(filepath, "wb") as f:
            f.write(r.content)
        print(f"  [OK] {filename} ({filepath.stat().st_size / 1024:.0f} KB)")
        return filepath
    except Exception as e:
        print(f"  [ERROR] {filename}: {e}")
        return None


def ut_check_exists(session, filename, year):
    try:
        session.headers["Referer"] = UT_BASE_URL
        session.get(UT_BASE_URL, timeout=15)
    except Exception:
        pass
    try:
        r = session.post(UT_BASE_URL, params=UT_LIST_PARAMS, data={"folder": year}, timeout=15)
        r.raise_for_status()
        soup = BeautifulSoup(r.text, "html.parser")
        for link in soup.find_all("a"):
            if link.get_text(strip=True) == filename:
                return True
    except Exception:
        pass
    return False


# --- EOR Download -------------------------------------------------------------

def eor_download(session, filename, output_dir):
    filepath = output_dir / filename
    file_hash = "l1_" + base64.b64encode(filename.encode()).decode()
    params = EOR_DL_PARAMS.copy()
    params["target"] = file_hash
    params["_t"] = str(int(time.time()))

    try:
        session.headers["Referer"] = EOR_PAGE_URL
        session.get(EOR_PAGE_URL, timeout=15)
    except Exception:
        pass

    try:
        r = session.get(EOR_BASE_URL, params=params, timeout=120)
        r.raise_for_status()
        if "text/html" in r.headers.get("Content-Type", ""):
            print(f"  [ERROR] {filename}: respuesta HTML")
            return None
        with open(filepath, "wb") as f:
            f.write(r.content)
        size = filepath.stat().st_size
        if size < 10240:
            print(f"  [ERROR] {filename}: muy pequeno ({size} bytes)")
            filepath.unlink()
            return None
        print(f"  [OK] {filename} ({size / 1024:.0f} KB)")
        return filepath
    except Exception as e:
        print(f"  [ERROR] {filename}: {e}")
        return None


# --- Email --------------------------------------------------------------------

def send_combined_email(files, fecha_str):
    gmail_user = os.environ["GMAIL_USER"]
    gmail_pass = os.environ["GMAIL_PASS"]
    dest_email = os.environ["DEST_EMAIL"]

    ahora = now_sv()

    # Convertir fecha para el asunto
    dt = datetime.strptime(fecha_str, "%Y-%m-%d")
    fecha_display = dt.strftime("%d/%m/%Y")

    msg = email.mime.multipart.MIMEMultipart()
    msg["From"] = gmail_user
    msg["To"] = dest_email
    msg["Subject"] = f"Predespacho {fecha_display} - UT & EOR"

    lista = []
    for f in files:
        lista.append(f"  - {f.name} ({f.stat().st_size / 1024:.0f} KB)")
    archivos_txt = "\n".join(lista)

    body = (
        f"Descarga por Fecha - {fecha_display}\n"
        f"{'=' * 50}\n\n"
        f"Archivos adjuntos ({len(files)}):\n"
        f"{archivos_txt}\n\n"
        f"Hora de descarga: {ahora.strftime('%H:%M')} (hora SV)\n\n"
        f"---\n"
        f"Enviado automaticamente\n"
    )
    msg.attach(email.mime.text.MIMEText(body, "plain", "utf-8"))

    for filepath in files:
        with open(filepath, "rb") as f:
            att = email.mime.base.MIMEBase("application", "octet-stream")
            att.set_payload(f.read())
        email.encoders.encode_base64(att)
        att.add_header("Content-Disposition", f"attachment; filename={filepath.name}")
        msg.attach(att)

    print(f"\n  [CORREO] Enviando {len(files)} archivos a {dest_email}...", end=" ")
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(gmail_user, gmail_pass)
        server.send_message(msg)
    print("OK!")


# --- Main ---------------------------------------------------------------------

def main():
    target_date = os.environ.get("TARGET_DATE")
    if not target_date:
        print("ERROR: Falta variable TARGET_DATE (YYYY-MM-DD)")
        sys.exit(1)

    dt = datetime.strptime(target_date, "%Y-%m-%d")
    ddmmyy = dt.strftime("%d%m%y")
    yyyymmdd = dt.strftime("%Y%m%d")
    year = str(dt.year)

    aislado_name = f"Prog_Diaria_Inicial_Aislado_{ddmmyy}.xlsx"
    prog_name = f"Prog_Diaria{ddmmyy}.xlsx"
    eor_name = f"PUB004-PRE-{yyyymmdd}-OSO002.zip"

    ahora = now_sv()
    print("=" * 60)
    print("  DESCARGA POR FECHA - TODOS LOS ARCHIVOS")
    print("=" * 60)
    print(f"  Fecha:     {target_date}")
    print(f"  Aislado:   {aislado_name}")
    print(f"  ProgDia:   {prog_name}")
    print(f"  EOR:       {eor_name}")
    print(f"  Inicio:    {ahora.strftime('%H:%M:%S')} (hora SV)")
    print("=" * 60)

    session = create_session()
    output_dir = Path("downloads")
    output_dir.mkdir(exist_ok=True)

    downloaded = []

    # 1. UT Aislado
    print(f"\n  [1/3] UT Aislado...")
    result = ut_download(session, aislado_name, year, output_dir)
    if result:
        downloaded.append(result)

    # 2. UT Prog Diaria
    print(f"\n  [2/3] UT Prog Diaria...")
    result = ut_download(session, prog_name, year, output_dir)
    if result:
        downloaded.append(result)

    # 3. EOR Predespacho
    print(f"\n  [3/3] EOR Predespacho...")
    result = eor_download(session, eor_name, output_dir)
    if result:
        downloaded.append(result)

    # Resumen
    print(f"\n{'=' * 60}")
    print(f"  RESULTADO: {len(downloaded)}/3 archivos descargados")
    print(f"{'=' * 60}")

    if downloaded:
        send_combined_email(downloaded, target_date)
    else:
        print("  [CORREO] No hay archivos para enviar")
        sys.exit(1)

    if len(downloaded) < 3:
        print(f"\n  AVISO: {3 - len(downloaded)} archivo(s) no encontrado(s)")


if __name__ == "__main__":
    main()
