# email_notifier.py

import logging
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Optional

from models import UserReport

logger = logging.getLogger(__name__)


class EmailNotifier:
    """Düşük worklog puanı olan kullanıcılara hatırlatma maili gönderir."""

    def __init__(self, config: dict):
        self.config = config
        self.email_config = config.get("email", {})
        self.enabled = self.email_config.get("enabled", False)
        self.threshold = self.email_config.get("threshold", 70)
        self.smtp_config = self.email_config.get("smtp", {})
        self.from_address = self.email_config.get("from_address", "")
        self.from_name = self.email_config.get("from_name", "Jira Worklog Sistem")

        # Kullanıcı email eşleme tablosu
        self.user_emails: dict[str, str] = {}
        for u in config.get("users", []):
            if u.get("email"):
                self.user_emails[u["username"]] = u["email"]
                # display_name ile de eşle
                if u.get("display_name"):
                    self.user_emails[u["display_name"]] = u["email"]

    def process(self, reports: list[UserReport]):
        """Raporları kontrol eder, gerekirse mail gönderir."""
        if not self.enabled:
            logger.info("E-posta bildirimleri devre dışı.")
            return

        low_performers = [
            r for r in reports if r.success_percentage < self.threshold
        ]

        if not low_performers:
            logger.info(
                f"Tüm kullanıcılar %{self.threshold} üzerinde. "
                f"Mail gönderilmeyecek."
            )
            return

        logger.info(
            f"{len(low_performers)} kullanıcı eşik altında, "
            f"mail gönderiliyor..."
        )

        # Bireysel hatırlatma mailleri
        sent_count = 0
        failed_count = 0
        for report in low_performers:
            email = self._find_email(report)
            if not email:
                logger.warning(
                    f"  ⚠️ {report.display_name}: email adresi bulunamadı, "
                    f"atlanıyor."
                )
                failed_count += 1
                continue

            success = self._send_user_reminder(report, email)
            if success:
                sent_count += 1
            else:
                failed_count += 1

        # Yönetici özet maili
        if self.email_config.get("send_manager_summary", False):
            manager_emails = self.email_config.get("manager_emails", [])
            if manager_emails:
                self._send_manager_summary(reports, low_performers, manager_emails)

        logger.info(
            f"Mail işlemi tamamlandı: "
            f"{sent_count} gönderildi, {failed_count} başarısız"
        )

    def _find_email(self, report: UserReport) -> Optional[str]:
        """Kullanıcının email adresini bulur."""
        return (
            self.user_emails.get(report.username)
            or self.user_emails.get(report.display_name)
        )

    def _send_email(self, to: str, subject: str, html_body: str) -> bool:
        """SMTP üzerinden mail gönderir."""
        try:
            msg = MIMEMultipart("alternative")
            msg["From"] = f"{self.from_name} <{self.from_address}>"
            msg["To"] = to
            msg["Subject"] = subject
            msg.attach(MIMEText(html_body, "html", "utf-8"))

            server_addr = self.smtp_config.get("server", "")
            port = self.smtp_config.get("port", 587)
            use_tls = self.smtp_config.get("use_tls", True)

            with smtplib.SMTP(server_addr, port, timeout=30) as server:
                if use_tls:
                    server.starttls()
                username = self.smtp_config.get("username", "")
                password = self.smtp_config.get("password", "")
                if username and password:
                    server.login(username, password)
                server.sendmail(self.from_address, to, msg.as_string())

            logger.info(f"  ✅ Mail gönderildi: {to}")
            return True

        except smtplib.SMTPAuthenticationError:
            logger.error(f"  ❌ SMTP kimlik doğrulama hatası: {to}")
            return False
        except smtplib.SMTPRecipientsRefused:
            logger.error(f"  ❌ Geçersiz alıcı adresi: {to}")
            return False
        except smtplib.SMTPException as e:
            logger.error(f"  ❌ SMTP hatası ({to}): {e}")
            return False
        except Exception as e:
            logger.error(f"  ❌ Mail gönderme hatası ({to}): {e}")
            return False

    # ═══════════════════════════════════════════
    #  KULLANICI HATIRLATMA MAİLİ
    # ═══════════════════════════════════════════
    def _send_user_reminder(self, report: UserReport, to_email: str) -> bool:
        """Kullanıcıya kişisel hatırlatma maili gönderir."""
        subject_tpl = self.email_config.get(
            "subject_template",
            "⚠️ Worklog Hatırlatma — {name} (%{percentage})"
        )
        subject = subject_tpl.format(
            name=report.display_name,
            percentage=report.success_percentage,
        )

        html = self._build_user_html(report)
        return self._send_email(to_email, subject, html)

    def _build_user_html(self, report: UserReport) -> str:
        """Kullanıcı hatırlatma mailinin HTML içeriğini oluşturur."""
        pct = report.success_percentage
        dr = self.config["date_range"]

        # Durum rengi
        if pct >= 70:
            status_color = "#BF8F00"
            status_bg = "#FFF8E1"
            status_text = "ORTA"
        else:
            status_color = "#C00000"
            status_bg = "#FFF0F0"
            status_text = "KÖTÜ"

        # Worklog girilmemiş günler
        missing_days_html = ""
        missing_days = []
        DAY_TR = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
        for d in sorted(report.daily_summaries.keys()):
            ds = report.daily_summaries[d]
            if d.weekday() < 5 and ds.total_seconds == 0:
                missing_days.append(
                    f"{d.strftime('%d.%m.%Y')} ({DAY_TR[d.weekday()]})"
                )

        if missing_days:
            rows = "".join(
                f'<tr><td style="padding:6px 12px;border-bottom:1px solid #eee;">'
                f'❌ {day}</td></tr>'
                for day in missing_days
            )
            missing_days_html = f"""
            <div style="margin-top:20px;">
                <h3 style="color:#C00000;margin-bottom:8px;">
                    📅 Worklog Girilmemiş Günler ({len(missing_days)})
                </h3>
                <table style="border-collapse:collapse;width:100%;
                              background:#fff;border-radius:6px;
                              border:1px solid #eee;">
                    {rows}
                </table>
            </div>
            """

        # Worklog girilmemiş issue'lar
        missing_issues_html = ""
        no_wl_issues = [
            i for i in report.matching_issues if not i.get("has_worklog")
        ]
        if no_wl_issues:
            rows = "".join(
                f'<tr>'
                f'<td style="padding:6px 12px;border-bottom:1px solid #eee;">'
                f'<strong>{i["key"]}</strong></td>'
                f'<td style="padding:6px 12px;border-bottom:1px solid #eee;">'
                f'{i["summary"][:50]}</td>'
                f'<td style="padding:6px 12px;border-bottom:1px solid #eee;">'
                f'{i["status"]}</td>'
                f'</tr>'
                for i in no_wl_issues
            )
            missing_issues_html = f"""
            <div style="margin-top:20px;">
                <h3 style="color:#C00000;margin-bottom:8px;">
                    📋 Worklog Girilmemiş Issue'lar ({len(no_wl_issues)})
                </h3>
                <table style="border-collapse:collapse;width:100%;
                              background:#fff;border-radius:6px;
                              border:1px solid #eee;">
                    <tr style="background:#f5f5f5;">
                        <th style="padding:8px 12px;text-align:left;">Key</th>
                        <th style="padding:8px 12px;text-align:left;">Özet</th>
                        <th style="padding:8px 12px;text-align:left;">Statü</th>
                    </tr>
                    {rows}
                </table>
            </div>
            """

        from datetime import datetime as _dt
        _start = _dt.strptime(str(dr['start_date']), "%Y-%m-%d")
        _end = _dt.strptime(str(dr['end_date']), "%Y-%m-%d")

        return f"""
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;
             background:#f4f6f8;">

  <!-- Ana Konteyner -->
  <div style="max-width:650px;margin:20px auto;background:#ffffff;
              border-radius:12px;overflow:hidden;
              box-shadow:0 2px 12px rgba(0,0,0,0.1);">

    <!-- Üst Banner -->
    <div style="background:linear-gradient(135deg,#2F5496,#4472C4);
                padding:24px 30px;text-align:center;">
      <h1 style="color:#fff;margin:0;font-size:22px;">
        ⚠️ Worklog Hatırlatması
      </h1>
      <p style="color:#B4C6E7;margin:8px 0 0;font-size:14px;">
        {_start.strftime('%d.%m.%Y')} — {_end.strftime('%d.%m.%Y')}
      </p>
    </div>

    <!-- İçerik -->
    <div style="padding:24px 30px;">

      <!-- Selamlama -->
      <p style="font-size:15px;color:#333;">
        Merhaba <strong>{report.display_name}</strong>,
      </p>
      <p style="font-size:14px;color:#555;line-height:1.6;">
        Belirtilen tarih aralığındaki worklog kayıtlarınız incelendi.
        Başarı oranınız beklenen seviyenin altında kalmıştır.
        Lütfen eksik worklog girişlerinizi tamamlayınız.
      </p>

      <!-- Skor Kartı -->
      <div style="background:{status_bg};border-left:4px solid {status_color};
                  border-radius:8px;padding:20px;margin:20px 0;
                  text-align:center;">
        <div style="font-size:42px;font-weight:bold;color:{status_color};">
          %{pct}
        </div>
        <div style="font-size:14px;color:{status_color};
                    margin-top:4px;font-weight:600;">
          Durum: {status_text}
        </div>
        <div style="margin-top:12px;font-size:13px;color:#666;">
          Worklog girilen gün: <strong>{report.total_days_worked}</strong> /
          Toplam iş günü: <strong>{report.total_working_days}</strong>
          &nbsp;|&nbsp;
          Toplam çalışma: <strong>{report.total_hours} saat</strong>
        </div>
      </div>

      <!-- İlerleme Çubuğu -->
      <div style="background:#eee;border-radius:10px;height:20px;
                  overflow:hidden;margin:16px 0;">
        <div style="background:{status_color};height:100%;
                    width:{min(pct, 100)}%;border-radius:10px;
                    transition:width 0.3s;">
        </div>
      </div>
      <div style="display:flex;justify-content:space-between;
                  font-size:11px;color:#999;">
        <span>0%</span>
        <span style="color:#BF8F00;">70% (Orta)</span>
        <span style="color:#548235;">90% (İyi)</span>
        <span>100%</span>
      </div>

      {missing_days_html}

      {missing_issues_html}

      <!-- Aksiyon Butonu -->
      <div style="text-align:center;margin:28px 0 12px;">
        <a href="{self.config['jira']['server']}"
           style="display:inline-block;background:#2F5496;color:#fff;
                  padding:12px 32px;border-radius:6px;text-decoration:none;
                  font-weight:600;font-size:14px;">
          🔗 Jira'ya Git
        </a>
      </div>

    </div>

    <!-- Alt Bilgi -->
    <div style="background:#f8f9fa;padding:16px 30px;
                border-top:1px solid #eee;text-align:center;">
      <p style="font-size:11px;color:#999;margin:0;">
        Bu mail otomatik olarak Jira Worklog Kontrol Sistemi
        tarafından gönderilmiştir.
        <br>
        Gönderim zamanı: {datetime.now().strftime('%d.%m.%Y %H:%M')}
      </p>
    </div>

  </div>

</body>
</html>
"""

    # ═══════════════════════════════════════════
    #  YÖNETİCİ ÖZET MAİLİ
    # ═══════════════════════════════════════════
    def _send_manager_summary(
        self,
        all_reports: list[UserReport],
        low_reports: list[UserReport],
        manager_emails: list[str],
    ):
        """Yöneticilere özet mail gönderir."""
        subject = self.email_config.get(
            "manager_subject",
            "📊 Worklog Haftalık Özet Raporu"
        )
        html = self._build_manager_html(all_reports, low_reports)

        for email in manager_emails:
            self._send_email(email, subject, html)

    def _build_manager_html(
        self,
        all_reports: list[UserReport],
        low_reports: list[UserReport],
    ) -> str:
        """Yönetici özet mailinin HTML içeriğini oluşturur."""
        dr = self.config["date_range"]
        from datetime import datetime as _dt
        _start = _dt.strptime(str(dr['start_date']), "%Y-%m-%d")
        _end = _dt.strptime(str(dr['end_date']), "%Y-%m-%d")

        total_users = len(all_reports)
        good = len([r for r in all_reports if r.success_percentage >= 90])
        medium = len([
            r for r in all_reports
            if 70 <= r.success_percentage < 90
        ])
        bad = len([r for r in all_reports if r.success_percentage < 70])

        # Tüm kullanıcı satırları
        all_rows = ""
        for r in sorted(all_reports, key=lambda x: x.success_percentage):
            pct = r.success_percentage
            if pct >= 90:
                row_color = "#f0fff0"
                badge_color = "#548235"
                badge_text = "İYİ"
            elif pct >= 70:
                row_color = "#fffef0"
                badge_color = "#BF8F00"
                badge_text = "ORTA"
            else:
                row_color = "#fff0f0"
                badge_color = "#C00000"
                badge_text = "KÖTÜ"

            all_rows += f"""
            <tr style="background:{row_color};">
                <td style="padding:10px 12px;border-bottom:1px solid #eee;">
                    {r.display_name}
                </td>
                <td style="padding:10px 12px;border-bottom:1px solid #eee;
                           text-align:center;">
                    <strong style="color:{badge_color};font-size:16px;">
                        %{pct}
                    </strong>
                </td>
                <td style="padding:10px 12px;border-bottom:1px solid #eee;
                           text-align:center;">
                    <span style="background:{badge_color};color:#fff;
                                padding:3px 10px;border-radius:12px;
                                font-size:11px;font-weight:600;">
                        {badge_text}
                    </span>
                </td>
                <td style="padding:10px 12px;border-bottom:1px solid #eee;
                           text-align:center;">
                    {r.total_days_worked}/{r.total_working_days}
                </td>
                <td style="padding:10px 12px;border-bottom:1px solid #eee;
                           text-align:center;">
                    {r.total_hours}s
                </td>
                <td style="padding:10px 12px;border-bottom:1px solid #eee;
                           text-align:center;">
                    {len(r.matching_issues)}
                </td>
            </tr>
            """

        return f"""
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;
             background:#f4f6f8;">

  <div style="max-width:750px;margin:20px auto;background:#ffffff;
              border-radius:12px;overflow:hidden;
              box-shadow:0 2px 12px rgba(0,0,0,0.1);">

    <!-- Banner -->
    <div style="background:linear-gradient(135deg,#1a365d,#2F5496);
                padding:24px 30px;text-align:center;">
      <h1 style="color:#fff;margin:0;font-size:22px;">
        📊 Worklog Özet Raporu
      </h1>
      <p style="color:#B4C6E7;margin:8px 0 0;font-size:14px;">
        {_start.strftime('%d.%m.%Y')} — {_end.strftime('%d.%m.%Y')}
      </p>
    </div>

    <div style="padding:24px 30px;">

      <!-- Özet Kartları -->
      <div style="display:flex;gap:12px;margin-bottom:24px;">
        <div style="flex:1;background:#f0fff0;border-radius:8px;
                    padding:16px;text-align:center;
                    border:1px solid #c6efce;">
          <div style="font-size:28px;font-weight:bold;color:#548235;">
            {good}
          </div>
          <div style="font-size:12px;color:#548235;margin-top:4px;">
            🟢 İYİ (≥90%)
          </div>
        </div>
        <div style="flex:1;background:#fffef0;border-radius:8px;
                    padding:16px;text-align:center;
                    border:1px solid #ffeb9c;">
          <div style="font-size:28px;font-weight:bold;color:#BF8F00;">
            {medium}
          </div>
          <div style="font-size:12px;color:#BF8F00;margin-top:4px;">
            🟡 ORTA (70-89%)
          </div>
        </div>
        <div style="flex:1;background:#fff0f0;border-radius:8px;
                    padding:16px;text-align:center;
                    border:1px solid #ffc7ce;">
          <div style="font-size:28px;font-weight:bold;color:#C00000;">
            {bad}
          </div>
          <div style="font-size:12px;color:#C00000;margin-top:4px;">
            🔴 KÖTÜ (&lt;70%)
          </div>
        </div>
      </div>

      <!-- Tüm Kullanıcılar Tablosu -->
      <h3 style="color:#2F5496;margin-bottom:8px;">
        👥 Tüm Kullanıcılar ({total_users})
      </h3>
      <table style="border-collapse:collapse;width:100%;
                    border-radius:8px;overflow:hidden;
                    border:1px solid #ddd;">
        <tr style="background:#2F5496;">
          <th style="padding:10px 12px;color:#fff;text-align:left;">
            Kullanıcı</th>
          <th style="padding:10px 12px;color:#fff;text-align:center;">
            Başarı</th>
          <th style="padding:10px 12px;color:#fff;text-align:center;">
            Durum</th>
          <th style="padding:10px 12px;color:#fff;text-align:center;">
            Gün</th>
          <th style="padding:10px 12px;color:#fff;text-align:center;">
            Saat</th>
          <th style="padding:10px 12px;color:#fff;text-align:center;">
            Issue</th>
        </tr>
        {all_rows}
      </table>

      <!-- Eşik altı uyarı -->
      {"" if not low_reports else f'''
      <div style="margin-top:20px;background:#fff0f0;
                  border-left:4px solid #C00000;
                  border-radius:8px;padding:16px;">
        <strong style="color:#C00000;">
          ⚠️ {len(low_reports)} kullanıcıya hatırlatma maili gönderildi
        </strong>
        <p style="font-size:13px;color:#666;margin:8px 0 0;">
          {", ".join(r.display_name for r in low_reports)}
        </p>
      </div>
      '''}

    </div>

    <!-- Footer -->
    <div style="background:#f8f9fa;padding:16px 30px;
                border-top:1px solid #eee;text-align:center;">
      <p style="font-size:11px;color:#999;margin:0;">
        Jira Worklog Kontrol Sistemi — Otomatik Rapor
        <br>
        {datetime.now().strftime('%d.%m.%Y %H:%M')}
      </p>
    </div>

  </div>

</body>
</html>
"""