"""
EOR (Ente Operador Regional) - Descargador de Predespacho SV
Descarga archivos ZIP de predespacho desde www.enteoperador.org

Patron de archivos: PUB004-PRE-YYYYMMDD-OSO002.zip

Uso:
    python eor_scraper.py --list-only
    python eor_scraper.py --date 2026-03-22
    python eor_scraper.py --watch --email
"""

import argparse
import base64
import email.encoders
import email.mime.base
import email.mime.multipart
import email.mime.text
import io
import json
import os
import smtplib
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

# Forzar UTF-8 en stdout para Windows
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

try:
    import requests
    from bs4 import BeautifulSoup
except ImportError:
    print("Instalando dependencias necesarias...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests", "beautifulsoup4"])
    import requests
    from bs4 import BeautifulSoup


# --- Config -------------------------------------------------------------------
BASE_URL = "https://www.enteoperador.org"
DOWNLOAD_URL = BASE_URL + "/"
DOWNLOAD_PARAMS = {
    "red_fm_connect": "true",
    "front": "user",
    "fid": "117",
    "defaults": "0",
    "access_all": "0",
    "cmd": "file",
}

PAGE_URL = (
    "https://www.enteoperador.org/mer/gestion-comercial/"
    "informes-publicos-de-procesos-comerciales/"
    "informe-de-procesos-comerciales-el-salvador/predespacho-sv/"
)

LIST_PARAMS = {
    "red_fm_connect": "true",
    "front": "user",
    "fid": "117",
    "defaults": "0",
    "access_all": "0",
    "cmd": "open",
    "target": "l1_Lw",
}

SCRIPT_DIR = Path(__file__).parent.resolve()
CONFIG_FILE = SCRIPT_DIR / "eor_config.json"

FILE_PREFIX = "PUB004-PRE-"
FILE_SUFFIX = "-OSO002.zip"

# Zona horaria El Salvador (UTC-6)
TZ_SV = timezone(timedelta(hours=-6))


def now_sv():
    """Retorna la hora actual en zona horaria de El Salvador."""
    return datetime.now(TZ_SV)


# --- Utilidades ---------------------------------------------------------------

def filename_to_hash(filename: str) -> str:
    encoded = base64.b64encode(filename.encode()).decode()
    return f"l1_{encoded}"


def build_expected_filename(date_str: str) -> str:
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{FILE_PREFIX}{dt.strftime('%Y%m%d')}{FILE_SUFFIX}"


def get_today_date() -> str:
    return now_sv().strftime("%Y-%m-%d")


def get_tomorrow_date() -> str:
    return (now_sv() + timedelta(days=1)).strftime("%Y-%m-%d")


# --- Correo -------------------------------------------------------------------

def load_config() -> dict:
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_config(config: dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


def setup_email():
    print("=" * 60)
    print("  CONFIGURACION DE CORREO (Gmail) - EOR Scraper")
    print("=" * 60)
    print()
    print("  Para Gmail necesitas una 'Contrasena de aplicacion'.")
    print("  https://myaccount.google.com/apppasswords")
    print()

    config = load_config()
    gmail_user = input("  Tu correo Gmail: ").strip()
    gmail_pass = input("  Contrasena de aplicacion: ").strip()
    dest_email = input("  Correo destinatario: ").strip()

    config["email"] = {
        "gmail_user": gmail_user,
        "gmail_pass": gmail_pass,
        "dest_email": dest_email,
    }
    save_config(config)

    print(f"\n  Guardado en: {CONFIG_FILE}")
    print("  Probando conexion...", end=" ")
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as s:
            s.starttls()
            s.login(gmail_user, gmail_pass)
        print("OK!")
    except Exception as e:
        print(f"ERROR: {e}")


def send_email(filepath: Path, target_date_str: str = None) -> bool:
    if os.environ.get("GMAIL_USER"):
        ec = {
            "gmail_user": os.environ["GMAIL_USER"],
            "gmail_pass": os.environ["GMAIL_PASS"],
            "dest_email": os.environ["DEST_EMAIL"],
        }
    else:
        config = load_config()
        if "email" not in config:
            shared_config = SCRIPT_DIR / "ut_config.json"
            if shared_config.exists():
                with open(shared_config, "r", encoding="utf-8") as f:
                    config = json.load(f)
            if "email" not in config:
                print("  [CORREO] No configurado.")
                return False
        ec = config["email"]

    ahora = now_sv()

    if target_date_str:
        dt = datetime.strptime(target_date_str, "%Y-%m-%d")
        fecha_archivo = dt.strftime("%d/%m/%Y")
    else:
        fecha_archivo = ahora.strftime("%d/%m/%Y")

    msg = email.mime.multipart.MIMEMultipart()
    msg["From"] = ec["gmail_user"]
    msg["To"] = ec["dest_email"]
    msg["Subject"] = f"EOR Predespacho SV - {fecha_archivo}"

    body = (
        f"Predespacho El Salvador - Ente Operador Regional\n"
        f"{'=' * 50}\n\n"
        f"Fecha: {fecha_archivo}\n"
        f"Archivo: {filepath.name}\n"
        f"Tamano: {filepath.stat().st_size / 1024:.0f} KB\n"
        f"Hora de descarga: {ahora.strftime('%H:%M')} (hora SV)\n\n"
        f"---\n"
        f"Enviado automaticamente por eor_scraper.py\n"
    )
    msg.attach(email.mime.text.MIMEText(body, "plain", "utf-8"))

    with open(filepath, "rb") as f:
        att = email.mime.base.MIMEBase("application", "octet-stream")
        att.set_payload(f.read())
    email.encoders.encode_base64(att)
    att.add_header("Content-Disposition", f"attachment; filename={filepath.name}")
    msg.attach(att)

    try:
        print(f"  [CORREO] Enviando a {ec['dest_email']}...", end=" ")
        with smtplib.SMTP("smtp.gmail.com", 587) as s:
            s.starttls()
            s.login(ec["gmail_user"], ec["gmail_pass"])
            s.send_message(msg)
        print("OK!")
        return True
    except Exception as e:
        print(f"ERROR: {e}")
        return False


# --- Scraping -----------------------------------------------------------------

def create_session() -> requests.Session:
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Referer": PAGE_URL,
    })
    return session


def list_files(session: requests.Session, limit: int = 30) -> list:
    print("  [BUSCAR] Consultando archivos en EOR...")

    try:
        session.get(PAGE_URL, timeout=30)
    except requests.RequestException:
        pass

    try:
        response = session.get(DOWNLOAD_URL, params=LIST_PARAMS, timeout=30)
        response.raise_for_status()
        data = response.json()
    except (requests.RequestException, json.JSONDecodeError) as e:
        print(f"  [ERROR] {e}")
        return []

    files = []
    file_list = data.get("files", [])

    for f in file_list:
        name = f.get("name", "")
        if name.startswith(FILE_PREFIX) and name.endswith(".zip"):
            files.append({
                "filename": name,
                "hash": f.get("hash", ""),
                "size": f.get("size", 0),
                "date": f.get("ts", 0),
            })

    files.sort(key=lambda x: x["filename"], reverse=True)

    if limit:
        files = files[:limit]

    print(f"  [OK] Encontrados: {len(files)} archivo(s)")
    return files


def download_file(session: requests.Session, filename: str, output_dir: Path, overwrite: bool = False):
    filepath = output_dir / filename

    if filepath.exists() and not overwrite:
        print(f"  [SKIP] {filename} (ya existe)")
        return filepath

    file_hash = filename_to_hash(filename)
    params = DOWNLOAD_PARAMS.copy()
    params["target"] = file_hash
    params["_t"] = str(int(time.time()))

    try:
        response = session.get(DOWNLOAD_URL, params=params, timeout=120, stream=True)
        response.raise_for_status()

        content_type = response.headers.get("Content-Type", "")
        if "text/html" in content_type:
            print(f"  [ERROR] {filename}: respuesta HTML, no ZIP")
            return None

        with open(filepath, "wb") as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        size_kb = filepath.stat().st_size / 1024
        if size_kb < 10:
            print(f"  [ERROR] {filename}: archivo muy pequeno ({size_kb:.0f} KB), posible error")
            filepath.unlink()
            return None

        print(f"  [OK] {filename} ({size_kb:.0f} KB)")
        return filepath

    except requests.RequestException as e:
        print(f"  [ERROR] {filename}: {e}")
        if filepath.exists():
            filepath.unlink()
        return None


def watch_and_download(target_date: str, send_mail: bool, retry_interval: int = 300, max_retries: int = 6):
    expected_file = build_expected_filename(target_date)

    print("=" * 60)
    print("  EOR - MODO WATCH (Predespacho SV)")
    print("=" * 60)
    print(f"  Archivo esperado: {expected_file}")
    print(f"  Fecha objetivo:   {target_date}")
    print(f"  Reintento cada:   {retry_interval // 60} minutos")
    print(f"  Intentos maximos: {max_retries}")
    print(f"  Tiempo maximo:    {(retry_interval * max_retries) // 60} minutos")
    print(f"  Enviar correo:    {'Si' if send_mail else 'No'}")
    print(f"  Inicio:           {now_sv().strftime('%Y-%m-%d %H:%M:%S')} (hora SV)")
    print("=" * 60)

    session = create_session()
    output_dir = SCRIPT_DIR / "eor_predespacho"
    output_dir.mkdir(parents=True, exist_ok=True)

    for attempt in range(1, max_retries + 1):
        print(f"\n  --- Intento {attempt}/{max_retries} - {now_sv().strftime('%H:%M:%S')} (hora SV) ---")

        files = list_files(session, limit=50)
        filenames = [f["filename"] for f in files]

        if expected_file in filenames:
            print(f"\n  [ENCONTRADO] {expected_file}")

            result = download_file(session, expected_file, output_dir, overwrite=True)

            if result:
                if send_mail:
                    print()
                    send_email(result, target_date)

                print(f"\n{'=' * 60}")
                print(f"  COMPLETADO - {now_sv().strftime('%H:%M:%S')} (hora SV)")
                print(f"{'=' * 60}")
                return True
        else:
            recent = filenames[:3] if filenames else ["(vacio)"]
            print(f"  [NO ENCONTRADO] {expected_file}")
            print(f"  Mas recientes: {', '.join(recent)}")

        if attempt < max_retries:
            next_t = (now_sv() + timedelta(seconds=retry_interval)).strftime('%H:%M:%S')
            print(f"  [ESPERAR] Proximo intento a las {next_t} (hora SV)...")
            time.sleep(retry_interval)

    print(f"\n{'=' * 60}")
    print(f"  TIMEOUT - No se encontro {expected_file} despues de {max_retries} intentos")
    print(f"{'=' * 60}")
    return False


# --- Main ---------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Descarga archivos de Predespacho SV desde EOR",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument("--date", default=None, help="Fecha del archivo a descargar (YYYY-MM-DD)")
    parser.add_argument("--list-only", action="store_true", help="Solo listar archivos")
    parser.add_argument("-o", "--output", default=None, help="Carpeta de salida")
    parser.add_argument("--overwrite", action="store_true", help="Sobrescribir existentes")
    parser.add_argument("--email", action="store_true", help="Enviar por correo")
    parser.add_argument("--setup-email", action="store_true", help="Configurar correo")
    parser.add_argument("--watch", action="store_true",
                        help="Modo watch: busca archivo de hoy, reintenta cada 5 min por 30 min")
    parser.add_argument("--watch-date", default=None,
                        help="Fecha especifica para modo watch (YYYY-MM-DD, default: hoy)")
    parser.add_argument("--retry-interval", type=int, default=300, help="Segundos entre reintentos (default: 300)")
    parser.add_argument("--max-retries", type=int, default=6, help="Max reintentos (default: 6)")
    parser.add_argument("--limit", type=int, default=30, help="Archivos a listar (default: 30)")

    args = parser.parse_args()

    if args.setup_email:
        setup_email()
        return

    if args.watch:
        target = args.watch_date or os.environ.get("TARGET_DATE") or get_tomorrow_date()
        watch_and_download(
            target_date=target,
            send_mail=args.email,
            retry_interval=args.retry_interval,
            max_retries=args.max_retries,
        )
        return

    # Modo normal
    print("=" * 60)
    print("  EOR - PREDESPACHO EL SALVADOR")
    print("=" * 60)
    print(f"  Fecha:   {now_sv().strftime('%Y-%m-%d %H:%M:%S')} (hora SV)")
    print("=" * 60)

    session = create_session()
    files = list_files(session, limit=args.limit)

    if not files:
        print("  [WARN] No se encontraron archivos")
        return

    print(f"\n  {'_' * 56}")
    print(f"  {len(files)} archivo(s) mas recientes:")
    print(f"  {'_' * 56}")

    for i, f in enumerate(files, 1):
        size = int(f["size"]) if f["size"] else 0
        size_str = f"{size / 1024:.0f} KB" if size < 1048576 else f"{size / 1048576:.1f} MB"
        print(f"  {i:3d}. {f['filename']:<45s} {size_str:>10s}")

    if args.list_only:
        return

    output_dir = Path(args.output) if args.output else SCRIPT_DIR / "eor_predespacho"
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.date:
        target = build_expected_filename(args.date)
        print(f"\n  [DESCARGAR] {target}")
        result = download_file(session, target, output_dir, args.overwrite)
        if result and args.email:
            send_email(result, args.date)
    else:
        latest = files[0]
        print(f"\n  [DESCARGAR] {latest['filename']} (mas reciente)")
        result = download_file(session, latest["filename"], output_dir, args.overwrite)
        if result and args.email:
            try:
                date_part = latest["filename"].replace(FILE_PREFIX, "").replace(FILE_SUFFIX, "")
                file_date = datetime.strptime(date_part, "%Y%m%d").strftime("%Y-%m-%d")
            except ValueError:
                file_date = None
            send_email(result, file_date)

    print(f"\n{'=' * 60}")


if __name__ == "__main__":
    main()
