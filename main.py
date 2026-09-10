# main.py

import logging
import sys
from datetime import datetime, date
from pathlib import Path

import yaml
from colorama import init, Fore, Style
from openpyxl import load_workbook

from jira_client import JiraClient
from worklog_analyzer import WorklogAnalyzer
from report_generator import ReportGenerator
from email_notifier import EmailNotifier


init(autoreset=True)

# ── Logging Ayarları ──
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.FileHandler("jira_worklog.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)


def load_config(path: str = "config.yaml") -> dict:
    config_path = Path(path)
    if not config_path.exists():
        print(Fore.RED + f"Konfigürasyon dosyası bulunamadı: {path}")
        print(Fore.YELLOW + "Örnek config.yaml oluşturuluyor...")
        create_sample_config(path)
        sys.exit(1)

    with open(config_path, "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)

    # ── YENİ: Excel'den kullanıcı yükleme ──
    if "users_file" in config:
        config["users"] = load_users_from_excel(config["users_file"])
    elif "users" not in config or not config["users"]:
        print(Fore.RED + "Config'de 'users' veya 'users_file' tanımlanmalı!")
        sys.exit(1)

    validate_config(config)
    return config


def validate_config(config: dict):
    """Konfigürasyonu doğrular."""
    required_keys = ["jira", "users", "date_range", "issue_filters"]
    for key in required_keys:
        if key not in config:
            raise ValueError(f"Config'de '{key}' anahtarı eksik!")

    # ── GÜNCELLEME: users artık Excel'den de gelebilir ──
    if not config.get("users"):
        raise ValueError(
            "Çalışan listesi boş! "
            "Config'de 'users_file' ile Excel yolu belirtin."
        )

    dr = config["date_range"]
    start = datetime.strptime(str(dr["start_date"]), "%d.%m.%Y").date()
    end = datetime.strptime(str(dr["end_date"]), "%d.%m.%Y").date()
    if start > end:
        raise ValueError("start_date, end_date'den büyük olamaz!")

    logger.info("Konfigürasyon doğrulandı ✓")


def create_sample_config(path: str):
    """Örnek config dosyası oluşturur."""
    sample = """
# JIRA Worklog Checker - Konfigürasyon
jira:
  server: "https://your-domain.atlassian.net"
  auth:
    email: "your-email@company.com"
    api_token: "YOUR_API_TOKEN"

users:
  - username: "user1"
    display_name: "User One"

date_range:
  start_date: "2025-01-01"
  end_date: "2025-01-31"

issue_filters:
  issue_types: ["Sub-task"]
  statuses: ["In Progress"]
  projects: ["MYPROJECT"]
  require_due_date: true
  require_assignee: true

worklog_rules:
  min_daily_hours: 4.0
  max_daily_hours: 10.0
  check_weekends: false

report:
  output_format: "both"
  excel_filename: "worklog_report.xlsx"
"""
    with open(path, "w", encoding="utf-8") as f:
        f.write(sample.strip())
    print(Fore.GREEN + f"Örnek config oluşturuldu: {path}")


def run_interactive():
    """
    Config dosyası olmadan interaktif çalıştırma.
    Hızlı test için kullanılabilir.
    """
    print(Fore.CYAN + "\n  🔧 İnteraktif Mod\n")

    server = input("  Jira Server URL: ").strip()
    email = input("  Email: ").strip()
    token = input("  API Token: ").strip()

    user_input = input(
        "  Çalışanlar (virgülle ayır): "
    ).strip()
    users = [
        {"username": u.strip(), "display_name": u.strip()}
        for u in user_input.split(",")
    ]

    start = input("  Başlangıç tarihi (YYYY-MM-DD): ").strip()
    end = input("  Bitiş tarihi (YYYY-MM-DD): ").strip()

    projects_input = input(
        "  Projeler (virgülle ayır, boş=tümü): "
    ).strip()
    projects = (
        [p.strip() for p in projects_input.split(",")]
        if projects_input else []
    )

    config = {
        "jira": {
            "server": server,
            "auth": {"email": email, "api_token": token},
        },
        "users": users,
        "date_range": {"start_date": start, "end_date": end},
        "issue_filters": {
            "issue_types": ["Sub-task"],
            "statuses": ["In Progress"],
            "projects": projects,
            "require_due_date": True,
            "require_assignee": True,
        },
        "worklog_rules": {
            "min_daily_hours": 4.0,
            "max_daily_hours": 10.0,
            "check_weekends": False,
        },
        "report": {
            "output_format": "console",
            "excel_filename": "worklog_report.xlsx",
        },
    }

    return config


def load_users_from_excel(filepath: str) -> list[dict]:
    """Excel dosyasından kullanıcı listesini okur."""
    path = Path(filepath)
    if not path.exists():
        print(Fore.RED + f"Kullanıcı dosyası bulunamadı: {filepath}")
        sys.exit(1)

    wb = load_workbook(path, read_only=True)
    ws = wb.active
    users = []

    headers = [
        str(cell.value).strip().lower() if cell.value else ""
        for cell in next(ws.iter_rows(min_row=1, max_row=1))
    ]

    username_col = None
    display_col = None
    email_col = None

    for idx, h in enumerate(headers):
        if h in ("username", "kullanıcı", "kullanıcı adı", "user"):
            username_col = idx
        elif h in ("displayname", "display name", "display_name", "görünen ad", "ad", "name"):
            display_col = idx
        elif h in ("email", "e-posta", "eposta", "mail"):
            email_col = idx

    if username_col is None:
        print(Fore.RED + "Excel'de 'username' sütunu bulunamadı!")
        print(Fore.YELLOW + f"  Bulunan sütunlar: {headers}")
        sys.exit(1)

    for row in ws.iter_rows(min_row=2, values_only=True):
        uname = row[username_col]
        if not uname or str(uname).strip() == "":
            continue

        uname = str(uname).strip()
        dname = (
            str(row[display_col]).strip()
            if display_col is not None and row[display_col]
            else uname
        )
        # ── YENİ: email ──
        email = (
            str(row[email_col]).strip()
            if email_col is not None and row[email_col]
            else None
        )

        users.append({
            "username": uname,
            "display_name": dname,
            "email": email,  # ← YENİ
        })

    wb.close()

    if not users:
        print(Fore.RED + f"Excel dosyasında kullanıcı bulunamadı: {filepath}")
        sys.exit(1)

    logger.info(f"{len(users)} kullanıcı Excel'den yüklendi: {filepath}")
    return users


def main():
    """Ana çalıştırma fonksiyonu."""
    print(Fore.CYAN + Style.BRIGHT + """
    ╔══════════════════════════════════════════════╗
    ║        JIRA WORKLOG CHECKER v1.0             ║
    ║    Worklog kontrol ve raporlama aracı        ║
    ╚══════════════════════════════════════════════╝
    """)

    # Komut satırı argümanları
    config_path = "config.yaml"
    interactive = False

    for arg in sys.argv[1:]:
        if arg == "--interactive" or arg == "-i":
            interactive = True
        elif arg.startswith("--config="):
            config_path = arg.split("=", 1)[1]
        elif arg == "--help" or arg == "-h":
            print("""
  Kullanım:
    python main.py                    config.yaml ile çalıştır
    python main.py --config=my.yaml   özel config dosyası
    python main.py --interactive      interaktif mod
    python main.py --help             bu yardım mesajı
            """)
            sys.exit(0)

    # Config yükle
    if interactive:
        config = run_interactive()
    else:
        config = load_config(config_path)

    # Jira bağlantısı
    try:
        client = JiraClient(
            server=config["jira"]["server"],
            username=config["jira"]["auth"]["username"],
            password=config["jira"]["auth"]["password"],
        )
    except Exception as e:
        print(Fore.RED + f"\n    Jira bağlantı hatası: {e}")
        sys.exit(1)

    # Analiz
    print(Fore.YELLOW + "\n    Analiz başlatılıyor...\n")

    analyzer = WorklogAnalyzer(client, config)
    reports = analyzer.analyze()

    # Rapor
    reporter = ReportGenerator(config)
    reporter.generate(reports)

    # ── YENİ: Mail bildirimi ──
    if config.get("email", {}).get("enabled", False):
        print(Fore.YELLOW + "\n  📧 Mail bildirimleri kontrol ediliyor...\n")
        notifier = EmailNotifier(config)
        notifier.process(reports)
    else:
        logger.info("E-posta bildirimleri devre dışı.")

    # Sonuç
    total_warnings = sum(len(r.warnings) for r in reports)
    if total_warnings > 0:
        print(
            Fore.YELLOW +
            f"\n    Toplam {total_warnings} uyarı bulundu."
        )
    else:
        print(Fore.GREEN + "\n    Tüm kontroller başarılı!")


if __name__ == "__main__":
    main()