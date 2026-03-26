"""
UT El Salvador - Descargador de Programacion Diaria (GitHub Actions)
Busca el archivo del dia siguiente, reintenta cada 5 min por 30 min,
descarga y envia por correo.

Variables de entorno requeridas:
    GMAIL_USER      - Correo Gmail remitente
    GMAIL_PASS      - Contrasena de aplicacion Gmail
    DEST_EMAIL      - Correo destinatario
    TIPO            - "aislado" o "prog"
"""

import email.encoders
import email.mime.base
import email.mime.multipart
import email.mime.text
import os
import re
import smtplib
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# --- Config -------------------------------------------------------------------
BASE_URL = "https://www.ut.com.sv/programacion-diaria1"
PORTLET_ID = "ProgramacionDiaria_WAR_PredespachoPublico_INSTANCE_Sw0UJdEgNCl7"

LIST_PARAMS = {
    "p_p_id": PORTLET_ID,
    "p_p_lifecycle": "1",
    "p_p_state": "normal",
    "p_p_mode": "view",
    f"_{PORTLET_ID}_javax.portlet.action": "detalleArchivos",
}

DOWNLOAD_PARAMS_BASE = {
    "p_p_id": PORTLET_ID,
    "p_p_lifecycle": "2",
    "p_p_state": "normal",
    "p_p_mode": "view",
    "p_p_cacheability": "cacheLevelPage",
    f"_{PORTLET_ID}_myaction": "detalles",
}

RETRY_INTERVAL = 300  # 5 minutos
MAX_RETRIES = 6       # 30 minutos total

# Zona horaria El Salvador (UTC-6)
TZ_SV = timezone(timedelta(hours=-6))


def now_sv():
    """Retorna la hora actual en zona horaria de El Salvador."""
    return datetime.now(TZ_SV)


def extract_date_from_filename(filename):
    """Extrae la fecha DDMMYY del nombre del archivo y retorna DD/MM/YYYY.
    Ej: Prog_Diaria_Inicial_Aislado_260326.xlsx -> 26/03/2026
    Ej: Prog_Diaria260326.xlsx -> 26/03/2026"""
    match = re.search(r'(\d{6})\.xlsx$', filename)
    if match:
        ddmmyy = match.group(1)
        try:
            dt = datetime.strptime(ddmmyy, "%d%m%y")
            return dt.strftime("%d/%m/%Y")
        except ValueError:
            pass
    return None


# --- Email --------------------------------------------------------------------

def send_email(filepath, tipo, target_date_str=None):
    gmail_user = os.environ["GMAIL_USER"]
    gmail_pass = os.environ["GMAIL_PASS"]
    dest_email = os.environ["DEST_EMAIL"]

    ahora = now_sv()
    tipo_nombre = "Prog_Diaria_Inicial_Aislado" if tipo == "aislado" else "Prog_Diaria"

    # Prioridad: fecha del nombre del archivo > target_date_str > fecha actual
    fecha_archivo = extract_date_from_filename(filepath.name)
    if not fecha_archivo and target_date_str:
        try:
            dt = datetime.strptime(target_date_str, "%Y-%m-%d")
            fecha_archivo = dt.strftime("%d/%m/%Y")
        except ValueError:
            fecha_archivo = None
    if not fecha_archivo:
        fecha_archivo = ahora.strftime("%d/%m/%Y")

    print(f"  [CORREO] Fecha extraida del archivo: {fecha_archivo} (de: {filepath.name})")

    msg = email.mime.multipart.MIMEMultipart()
    msg["From"] = gmail_user
    msg["To"] = dest_email
    msg["Subject"] = f"UT - {tipo_nombre} - {fecha_archivo}"

    body = (
        f"Programacion Diaria - UT El Salvador\n"
        f"{'=' * 45}\n\n"
        f"Tipo: {tipo_nombre}\n"
        f"Fecha del reporte: {fecha_archivo}\n"
        f"Hora de descarga: {ahora.strftime('%H:%M')} (hora SV)\n"
        f"Archivo: {filepath.name} ({filepath.stat().st_size / 1024:.0f} KB)\n\n"
        f"---\n"
        f"Enviado automaticamente por GitHub Actions\n"
    )
    msg.attach(email.mime.text.MIMEText(body, "plain", "utf-8"))

    with open(filepath, "rb") as f:
        attachment = email.mime.base.MIMEBase("application", "octet-stream")
        attachment.set_payload(f.read())
    email.encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", f"attachment; filename={filepath.name}")
    msg.attach(attachment)

    print(f"  [CORREO] Enviando a {dest_email}...", end=" ")
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(gmail_user, gmail_pass)
        server.send_message(msg)
    print("OK!")


# --- Scraping -----------------------------------------------------------------

def create_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "es-ES,es;q=0.9,en;q=0.8",
        "Referer": BASE_URL,
    })
    return session


def list_files(session, year):
    try:
        session.get(BASE_URL, timeout=30)
    except requests.RequestException:
        pass

    try:
        response = session.post(
            BASE_URL, params=LIST_PARAMS, data={"folder": year}, timeout=30
        )
        response.raise_for_status()
    except requests.RequestException as e:
        print(f"  [ERROR] {e}")
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    files = []

    for link in soup.find_all("a"):
        text = link.get_text(strip=True)
        if text.endswith(".xlsx") and link.get("href"):
            files.append(text)

    return files


def download_file(session, filename, year, output_dir):
    filepath = output_dir / filename
    params = DOWNLOAD_PARAMS_BASE.copy()
    params["p_p_resource_id"] = filename
    params[f"_{PORTLET_ID}_folder"] = year
    query = "&".join(f"{k}={v}" for k, v in params.items())
    url = f"{BASE_URL}?{query}"

    try:
        response = session.get(url, timeout=60)
        response.raise_for_status()

        if "text/html" in response.headers.get("Content-Type", ""):
            print(f"  [ERROR] {filename}: respuesta HTML")
            return None

        with open(filepath, "wb") as f:
            f.write(response.content)

        size_kb = filepath.stat().st_size / 1024
        print(f"  [OK] {filename} ({size_kb:.0f} KB)")
        return filepath

    except requests.RequestException as e:
        print(f"  [ERROR] {filename}: {e}")
        return None


def get_target_filename(tipo, target_date=None):
    """Genera nombre de archivo y ano para la fecha objetivo.
    Si TARGET_DATE env var esta definida (YYYY-MM-DD), usa esa fecha.
    Si no, usa manana (hora de El Salvador)."""
    if target_date:
        dt = datetime.strptime(target_date, "%Y-%m-%d")
    else:
        dt = now_sv() + timedelta(days=1)
    ddmmyy = dt.strftime("%d%m%y")
    year = str(dt.year)
    if tipo == "aislado":
        return f"Prog_Diaria_Inicial_Aislado_{ddmmyy}.xlsx", year, dt.strftime("%Y-%m-%d")
    else:
        return f"Prog_Diaria{ddmmyy}.xlsx", year, dt.strftime("%Y-%m-%d")


# --- Main ---------------------------------------------------------------------

def main():
    tipo = os.environ.get("TIPO", "aislado")
    target_date = os.environ.get("TARGET_DATE")
    expected_file, year, date_str = get_target_filename(tipo, target_date)

    max_retries = 1 if target_date else MAX_RETRIES
    retry_interval = 30 if target_date else RETRY_INTERVAL

    ahora = now_sv()
    print("=" * 60)
    print("  UT EL SALVADOR - GITHUB ACTIONS")
    print("=" * 60)
    print(f"  Tipo:             {'Aislado' if tipo == 'aislado' else 'Prog_Diaria'}")
    print(f"  Archivo esperado: {expected_file}")
    print(f"  Fecha objetivo:   {date_str}")
    print(f"  Carpeta ano:      {year}")
    print(f"  Reintentos:       cada {retry_interval // 60} min, max {max_retries}")
    print(f"  Inicio:           {ahora.strftime('%Y-%m-%d %H:%M:%S')} (hora SV)")
    print("=" * 60)

    session = create_session()
    output_dir = Path("downloads")
    output_dir.mkdir(exist_ok=True)

    for attempt in range(1, max_retries + 1):
        print(f"\n  --- Intento {attempt}/{max_retries} - {now_sv().strftime('%H:%M:%S')} (hora SV) ---")

        files = list_files(session, year)
        print(f"  [BUSCAR] {len(files)} archivos en carpeta {year}")

        if expected_file in files:
            print(f"\n  [ENCONTRADO] {expected_file}")

            result = download_file(session, expected_file, year, output_dir)

            if result:
                send_email(result, tipo, date_str)
                print(f"\n  COMPLETADO - {now_sv().strftime('%H:%M:%S')} (hora SV)")
                return

            print(f"  [ERROR] Descarga fallida")

        else:
            recent = files[:3] if files else ["(vacio)"]
            print(f"  [NO ENCONTRADO] {expected_file}")
            print(f"  Mas recientes: {', '.join(recent)}")

        if attempt < max_retries:
            next_t = (now_sv() + timedelta(seconds=retry_interval)).strftime('%H:%M:%S')
            print(f"  [ESPERAR] Proximo intento a las {next_t} (hora SV)...")
            time.sleep(retry_interval)

    print(f"\n  TIMEOUT - No se encontro {expected_file} despues de {max_retries} intentos")
    sys.exit(1)


if __name__ == "__main__":
    main()
