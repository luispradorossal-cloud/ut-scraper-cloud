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
import smtplib
import sys
import time
from datetime import datetime, timedelta
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


# --- Email --------------------------------------------------------------------

def send_email(filepath: Path, tipo: str):
    gmail_user = os.environ["GMAIL_USER"]
    gmail_pass = os.environ["GMAIL_PASS"]
    dest_email = os.environ["DEST_EMAIL"]

    now = datetime.now()
    tipo_nombre = "Prog_Diaria_Inicial_Aislado" if tipo == "aislado" else "Prog_Diaria"

    msg = email.mime.multipart.MIMEMultipart()
    msg["From"] = gmail_user
    msg["To"] = dest_email
    msg["Subject"] = f"UT - {tipo_nombre} - {now.strftime('%d/%m/%Y')}"

    body = (
        f"Programacion Diaria - UT El Salvador\n"
        f"{'=' * 45}\n\n"
        f"Tipo: {tipo_nombre}\n"
        f"Fecha: {now.strftime('%d/%m/%Y')}\n"
        f"Hora de descarga: {now.strftime('%H:%M')}\n"
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

def create_session() -> requests.Session:
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


def list_files(session: requests.Session, year: str) -> list:
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


def get_tomorrow_filename(tipo: str) -> str:
    tomorrow = datetime.now() + timedelta(days=1)
    ddmmyy = tomorrow.strftime("%d%m%y")
    if tipo == "aislado":
        return f"Prog_Diaria_Inicial_Aislado_{ddmmyy}.xlsx"
    else:
        return f"Prog_Diaria{ddmmyy}.xlsx"


# --- Main ---------------------------------------------------------------------

def main():
    tipo = os.environ.get("TIPO", "aislado")
    expected_file = get_tomorrow_filename(tipo)
    tomorrow = datetime.now() + timedelta(days=1)
    year = str(tomorrow.year)

    print("=" * 60)
    print("  UT EL SALVADOR - GITHUB ACTIONS")
    print("=" * 60)
    print(f"  Tipo:             {'Aislado' if tipo == 'aislado' else 'Prog_Diaria'}")
    print(f"  Archivo esperado: {expected_file}")
    print(f"  Carpeta ano:      {year}")
    print(f"  Reintentos:       cada {RETRY_INTERVAL // 60} min, max {MAX_RETRIES}")
    print(f"  Inicio:           {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 60)

    session = create_session()
    output_dir = Path("downloads")
    output_dir.mkdir(exist_ok=True)

    for attempt in range(1, MAX_RETRIES + 1):
        print(f"\n  --- Intento {attempt}/{MAX_RETRIES} - {datetime.now().strftime('%H:%M:%S')} ---")

        files = list_files(session, year)
        print(f"  [BUSCAR] {len(files)} archivos en carpeta {year}")

        if expected_file in files:
            print(f"\n  [ENCONTRADO] {expected_file}")

            result = download_file(session, expected_file, year, output_dir)

            if result:
                send_email(result, tipo)
                print(f"\n  COMPLETADO - {datetime.now().strftime('%H:%M:%S')}")
                return

            print(f"  [ERROR] Descarga fallida")

        else:
            recent = files[:3] if files else ["(vacio)"]
            print(f"  [NO ENCONTRADO] {expected_file}")
            print(f"  Mas recientes: {', '.join(recent)}")

        if attempt < MAX_RETRIES:
            next_t = (datetime.now() + timedelta(seconds=RETRY_INTERVAL)).strftime('%H:%M:%S')
            print(f"  [ESPERAR] Proximo intento a las {next_t}...")
            time.sleep(RETRY_INTERVAL)

    print(f"\n  TIMEOUT - No se encontro {expected_file} despues de {MAX_RETRIES} intentos")
    sys.exit(1)


if __name__ == "__main__":
    main()
