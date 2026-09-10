# email_notifier.py

import logging
from datetime import datetime
from typing import Optional

from models import UserReport

logger = logging.getLogger(__name__)


class OutlookClient:
    """pywin32 COM ile yerel Outlook uygulaması üzerinden mail gönderir."""

    def __init__(self, from_address: str):
        self.from_address = from_address
        self._outlook = None

    def _get_outlook(self):
        """Outlook COM nesnesini lazy başlatır."""
        if not self._outlook:
            try:
                import win32com.client
                self._outlook = win32com.client.Dispatch("Outlook.Application")
                logger.info("Outlook bağlantısı kuruldu")
            except Exception as e:
                logger.error(f"Outlook başlatılamadı: {e}")
                raise RuntimeError(
                    "Outlook uygulaması bulunamadı. "
                    "pywin32 yüklü ve Outlook açık olduğundan emin olun."
                ) from e
        return self._outlook

    def send_mail(
        self,
        to_addresses: list[str],
        subject: str,
        html_body: str,
        cc_addresses: Optional[list[str]] = None,
        importance: int = 1,
    ) -> bool:
        """
        Outlook COM üzerinden mail gönderir.

        Args:
            to_addresses: Alıcı adresleri
            subject: Konu
            html_body: HTML içerik
            cc_addresses: CC adresleri
            importance: 0=Low, 1=Normal, 2=High
        """
        try:
            outlook = self._get_outlook()
            mail = outlook.CreateItem(0)  # 0 = olMailItem

            mail.To = "; ".join(to_addresses)
            mail.Subject = subject
            mail.HTMLBody = html_body
            mail.Importance = importance

            if cc_addresses:
                mail.CC = "; ".join(cc_addresses)

            # Farklı hesaptan göndermek istenirse
            if self.from_address:
                try:
                    for account in outlook.Session.Accounts:
                        if account.SmtpAddress.lower() == self.from_address.lower():
                            mail._oleobj_.Invoke(
                                *(64209, 0, 8, 0, account)
                            )
                            break
                except Exception:
                    # Hesap bulunamazsa varsayılan hesapla gönderir
                    pass

            mail.Send()
            logger.info(f"  ✅ Mail gönderildi: {to_addresses}")
            return True

        except Exception as e:
            logger.error(f"  ❌ Mail gönderilemedi ({to_addresses}): {e}")
            return False


class EmailNotifier:
    """Düşük worklog puanı olan kullanıcılara hatırlatma maili gönderir."""

    def __init__(self, config: dict):
        self.config = config
        self.email_config = config.get("email", {})
        self.enabled = self.email_config.get("enabled", False)
        self.threshold = self.email_config.get("threshold", 70)
        self.from_address = self.email_config.get("from_address", "")

        self.client = OutlookClient(from_address=self.from_address)

        # Kullanıcı email eşleme
        self.user_emails: dict[str, str] = {}
        for u in config.get("users", []):
            if u.get("email"):
                self.user_emails[u["username"]] = u["email"]
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

        sent_count = 0
        failed_count = 0
        for report in low_performers:
            email = self._find_email(report)
            if not email:
                logger.warning(
                    f"  ⚠️ {report.display_name}: email bulunamadı, atlanıyor."
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
                self._send_manager_summary(
                    reports, low_performers, manager_emails
                )

        logger.info(
            f"Mail işlemi tamamlandı: "
            f"{sent_count} gönderildi, {failed_count} başarısız"
        )

    def _find_email(self, report: UserReport) -> Optional[str]:
        return (
            self.user_emails.get(report.username)
            or self.user_emails.get(report.display_name)
        )

    def _send_user_reminder(self, report: UserReport, to_email: str) -> bool:
        subject_tpl = self.email_config.get(
            "subject_template",
            "⚠️ Worklog Hatırlatma — {name} (%{percentage})"
        )
        subject = subject_tpl.format(
            name=report.display_name,
            percentage=report.success_percentage,
        )
        html = self._build_user_html(report)

        return self.client.send_mail(
            to_addresses=[to_email],
            subject=subject,
            html_body=html,
            importance=2,  # High
        )

    def _send_manager_summary(
        self,
        all_reports: list[UserReport],
        low_reports: list[UserReport],
        manager_emails: list[str],
    ):
        subject = self.email_config.get(
            "manager_subject",
            "📊 Worklog Haftalık Özet Raporu"
        )
        html = self._build_manager_html(all_reports, low_reports)

        self.client.send_mail(
            to_addresses=manager_emails,
            subject=subject,
            html_body=html,
        )

    # ═══════════════════════════════════════════
    #  HTML ŞABLONLARI
    # ═══════════════════════════════════════════
    def _build_user_html(self, report: UserReport) -> str:
        pct = report.success_percentage
        dr = self.config["date_range"]

        if pct >= 90:
            status_color = "#548235"
            status_bg = "#C6EFCE"
            status_text = "Tam Uyumlu Giriş"
            status_message = "Belirtilen tarih aralığındaki worklog kayıtlarınız incelendi. Girişleriniz beklenen seviyeyi tam olarak karşılamakta ve ilgili dönem için gerekli kayıtlar eksiksiz şekilde sağlanmış görünmektedir."
        elif pct >= 80:
            status_color = "#BF8F00"
            status_bg = "#FFEB9C"
            status_text = "Kabul Edilebilir Giriş"
            status_message = "Belirtilen tarih aralığındaki worklog kayıtlarınız incelendi. Girişleriniz genel olarak beklenen seviyeye yakın olmakla birlikte bazı günlerde eksiklikler bulunmaktadır. Worklog kayıtlarınızı daha düzenli ve eksiksiz şekilde giriniz."
        elif pct >= 70:
            status_color = "#E36C09"
            status_bg = "#FBD5B5"
            status_text = "Artırılması Gereken Giriş"
            status_message = "Belirtilen tarih aralığındaki worklog kayıtlarınız incelendi. Girişleriniz beklenen seviyenin altında kalmakla birlikte tamamen yetersiz değildir. Worklog giriş sıklığınızı arttırınız ve eksik günlere ait kayıtları tamamlayınız."
        else:
            status_color = "#C00000"
            status_bg = "#FFC7CE"
            status_text = "Yetersiz Giriş"
            status_message = "Belirtilen tarih aralığındaki worklog kayıtlarınız incelendi. Başarı oranınız beklenen seviyenin altında kalmıştır. Lütfen eksik worklog girişlerinizi tamamlayınız. İzinli günlerinizi dikkate almayınız."

        missing_days_html = ""
        missing_days = []
        DAY_TR = ["Pazartesi", "Salı", "Çarşamba", "Perşembe",
                  "Cuma", "Cumartesi", "Pazar"]
        for d in sorted(report.daily_summaries.keys()):
            ds = report.daily_summaries[d]
            if d.weekday() < 5 and ds.total_seconds == 0:
                missing_days.append(
                    f"{d.strftime('%d.%m.%Y')} ({DAY_TR[d.weekday()]})"
                )

        if missing_days:
            rows = "".join(
                f'<tr><td style="padding:6px 12px;'
                f'border-bottom:1px solid #eee;">'
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

        missing_issues_html = ""
        no_wl_issues = [
            i for i in report.matching_issues if not i.get("has_worklog")
        ]
        if report.matching_issues:
            rows = "".join(
                f'<tr>'
                f'<td style="padding:6px 12px;border-bottom:1px solid #eee;">'
                f'<strong>{i["key"]}</strong></td>'
                f'<td style="padding:6px 12px;border-bottom:1px solid #eee;">'
                f'{i["summary"][:50]}</td>'
                f'<td style="padding:6px 12px;border-bottom:1px solid #eee;">'
                f'{i["due_date"]}</td>'
                f'</tr>'
                for i in report.matching_issues
            )
            missing_issues_html = f"""
            <div style="margin-top:20px;">
                <h3 style="color:#5B9BD5;margin-bottom:8px;">
                    📋 Worklog Girişi İçin Uygun Issue'lar ({len(report.matching_issues)})
                </h3>
                <table style="border-collapse:collapse;width:100%;
                              background:#fff;border-radius:6px;
                              border:1px solid #eee;">
                    <tr style="background:#f5f5f5;">
                        <th style="padding:8px 12px;text-align:left;">Key</th>
                        <th style="padding:8px 12px;text-align:left;">Başlık</th>
                        <th style="padding:8px 12px;text-align:left;">Bitiş Tarihi</th>
                    </tr>
                    {rows}
                </table>
            </div>
            """

        from datetime import datetime as _dt
        _start = _dt.strptime(str(dr['start_date']), "%d.%m.%Y")
        _end = _dt.strptime(str(dr['end_date']), "%d.%m.%Y")

        return f"""
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;font-family:'Segoe UI',Arial,sans-serif;
             background:#ffffff;">
  <div style="max-width:650px;margin:20px auto;background:#ffffff;
              border-radius:12px;overflow:hidden;
              box-shadow:0 2px 12px rgba(0,0,0,0.1);">
    <div style="background:linear-gradient(135deg,#eee,#fff);
                padding:24px 30px;text-align:center;">
      <h1 style="color:#222;margin:0;font-size:22px;">
        ⚠️ Worklog Hatırlatması ⚠️
      </h1>
      <p style="color:#333;margin:8px 0 0;font-size:14px;">
        {_start.strftime('%d.%m.%Y')} — {_end.strftime('%d.%m.%Y')}</p>
    </div>
    <div style="padding:24px 30px;">
      <p style="font-size:15px;color:#333;">
        Merhaba <strong>{report.display_name}</strong>,</p>
      <p style="font-size:14px;color:#555;line-height:1.6;">
        {status_message}  
      </p>
      <div style="background:{status_bg};border-left:4px solid {status_color};
                  border-radius:8px;padding:20px;margin:20px 0;
                  text-align:center;">
        <div style="font-size:42px;font-weight:bold;color:{status_color};">
          %{pct}</div>
        <div style="font-size:14px;color:{status_color};
                    padding-top:4px;font-weight:600;">
          Durum: {status_text}</div>
        <div style="padding-top:12px;font-size:13px;color:#666;">
          Worklog girilen gün: <strong>{report.total_days_worked}</strong> /
          Toplam iş günü: <strong>{report.total_working_days}</strong>
          &nbsp;|&nbsp;
          Toplam çalışma: <strong>{report.total_hours} saat</strong></div>
      </div>

      <!--[if mso]>
      <table cellpadding="0" cellspacing="0" width="100%"
             style="margin:8px 0;">
        <tr>
          <td width="{min(pct, 100)}%"
              style="background:{status_color};height:20px;
                     font-size:1px;line-height:1px;">
            &nbsp;
          </td>
          <td width="{100 - min(pct, 100)}%"
              style="background:#E0E0E0;height:20px;
                     font-size:1px;line-height:1px;">
            &nbsp;
          </td>
        </tr>
      </table>
      <![endif]-->
      <!--[if !mso]><!-->
      <table cellpadding="0" cellspacing="0" width="100%"
             style="margin:8px 0;border-collapse:separate;">
        <tr>
          <td style="background:#E0E0E0;border-radius:10px;
                     padding:0;height:20px;">
            <table cellpadding="0" cellspacing="0"
                   width="{min(pct, 100)}%"
                   style="border-collapse:separate;">
              <tr>
                <td style="background:{status_color};
                           border-radius:10px;height:20px;
                           font-size:12px;color:#fff;
                           text-align:center;font-weight:bold;
                           font-family:'Segoe UI',Arial,sans-serif;">
                  {f'%{pct}' if pct >= 15 else '&nbsp;'}
                </td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      <!--<![endif]-->

      <table style="width:100%;font-size:11px;color:#999;margin-top:2px;"
             cellpadding="0" cellspacing="0">
        <tr>
          <td style="text-align:center;width:30%; color:#C00000">&lt;70% (Yetersiz Giriş)</td>
          <td style="text-align:center;width:20%;color:#E36C09;">70%-80% (Artırılması Gereken Giriş)</td>
          <td style="text-align:center;width:20%;color:#BF8F00;">80%-90% (Kabul Edilebilir Giriş)</td>
          <td style="text-align:center;width:30%;color:#548235">≥90 (Tam Uyumlu Giriş)</td>
        </tr>
      </table>
      {missing_days_html}
      {missing_issues_html}
      
      <div style="text-align:center;">
        <table cellpadding="0" cellspacing="16">
          <tr>
            <td align="center"
                style="background:#2F5496;border-radius:6px;
                       mso-padding-alt:10px 18px;">
              <a href="{self.config['jira']['server']}"
                 target="_blank"
                 style="display:inline-block; width: 100%; height: 100%;
                        color:#ffffff;text-decoration:none;
                        font-family:'Segoe UI',Arial,sans-serif;
                        font-size:14px;font-weight:600;
                        mso-line-height-rule:exactly;
                        line-height:20px;">
                &#128279; Jira'ya Git
              </a>
            </td>
          </tr>
        </table>

      </div>
    </div>
    <div style="background:#f8f9fa;padding:16px 30px;
                border-top:1px solid #eee;text-align:center;">
      <p style="font-size:11px;color:#999;margin:0;">
        Bu mail otomatik olarak Jira Worklog Kontrol Sistemi
        tarafından gönderilmiştir.<br>
      </p>
    </div>
  </div>
</body>
</html>
"""

    def _build_manager_html(
        self,
        all_reports: list[UserReport],
        low_reports: list[UserReport],
    ) -> str:
        dr = self.config["date_range"]
        from datetime import datetime as _dt
        _start = _dt.strptime(str(dr['start_date']), "%d.%m.%Y")
        _end = _dt.strptime(str(dr['end_date']), "%d.%m.%Y")

        total_users = len(all_reports)
        good = len([r for r in all_reports if r.success_percentage >= 90])
        medium = len([
            r for r in all_reports if 70 <= r.success_percentage < 90
        ])
        bad = len([r for r in all_reports if r.success_percentage < 70])

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
                <td style="padding:10px 12px;
                    border-bottom:1px solid #eee;">
                    {r.display_name}</td>
                <td style="padding:10px 12px;
                    border-bottom:1px solid #eee;text-align:center;">
                    <strong style="color:{badge_color};font-size:16px;">
                        %{pct}</strong></td>
                <td style="padding:10px 12px;
                    border-bottom:1px solid #eee;text-align:center;">
                    <span style="background:{badge_color};color:#fff;
                                padding:3px 10px;border-radius:12px;
                                font-size:11px;font-weight:600;">
                        {badge_text}</span></td>
                <td style="padding:10px 12px;
                    border-bottom:1px solid #eee;text-align:center;">
                    {r.total_days_worked}/{r.total_working_days}</td>
                <td style="padding:10px 12px;
                    border-bottom:1px solid #eee;text-align:center;">
                    {r.total_hours}s</td>
                <td style="padding:10px 12px;
                    border-bottom:1px solid #eee;text-align:center;">
                    {len(r.matching_issues)}</td>
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
    <div style="background:linear-gradient(135deg,#1a365d,#2F5496);
                padding:24px 30px;text-align:center;">
      <h1 style="color:#fff;margin:0;font-size:22px;">
        📊 Worklog Özet Raporu</h1>
      <p style="color:#B4C6E7;margin:8px 0 0;font-size:14px;">
        {_start.strftime('%d.%m.%Y')} — {_end.strftime('%d.%m.%Y')}</p>
    </div>
    <div style="padding:24px 30px;">
      <table style="width:100%;border-spacing:12px;">
        <tr>
          <td style="background:#f0fff0;border-radius:8px;
                     padding:16px;text-align:center;
                     border:1px solid #c6efce;">
            <div style="font-size:28px;font-weight:bold;
                        color:#548235;">{good}</div>
            <div style="font-size:12px;color:#548235;
                        margin-top:4px;">🟢 İYİ (≥90%)</div>
          </td>
          <td style="background:#fffef0;border-radius:8px;
                     padding:16px;text-align:center;
                     border:1px solid #ffeb9c;">
            <div style="font-size:28px;font-weight:bold;
                        color:#BF8F00;">{medium}</div>
            <div style="font-size:12px;color:#BF8F00;
                        margin-top:4px;">🟡 ORTA (70-89%)</div>
          </td>
          <td style="background:#fff0f0;border-radius:8px;
                     padding:16px;text-align:center;
                     border:1px solid #ffc7ce;">
            <div style="font-size:28px;font-weight:bold;
                        color:#C00000;">{bad}</div>
            <div style="font-size:12px;color:#C00000;
                        margin-top:4px;">🔴 KÖTÜ (&lt;70%)</div>
          </td>
        </tr>
      </table>
      <h3 style="color:#2F5496;margin:20px 0 8px;">
        👥 Tüm Kullanıcılar ({total_users})</h3>
      <table style="border-collapse:collapse;width:100%;
                    border-radius:8px;overflow:hidden;
                    border:1px solid #ddd;">
        <tr style="background:#2F5496;">
          <th style="padding:10px 12px;color:#fff;
                     text-align:left;">Kullanıcı</th>
          <th style="padding:10px 12px;color:#fff;
                     text-align:center;">Başarı</th>
          <th style="padding:10px 12px;color:#fff;
                     text-align:center;">Durum</th>
          <th style="padding:10px 12px;color:#fff;
                     text-align:center;">Gün</th>
          <th style="padding:10px 12px;color:#fff;
                     text-align:center;">Saat</th>
          <th style="padding:10px 12px;color:#fff;
                     text-align:center;">Issue</th>
        </tr>
        {all_rows}
      </table>
      {"" if not low_reports else f'''
      <div style="margin-top:20px;background:#fff0f0;
                  border-left:4px solid #C00000;
                  border-radius:8px;padding:16px;">
        <strong style="color:#C00000;">
          ⚠️ {len(low_reports)} kullanıcıya hatırlatma maili gönderildi
        </strong>
        <p style="font-size:13px;color:#666;margin:8px 0 0;">
          {", ".join(r.display_name for r in low_reports)}</p>
      </div>
      '''}
    </div>
    <div style="background:#f8f9fa;padding:16px 30px;
                border-top:1px solid #eee;text-align:center;">
      <p style="font-size:11px;color:#999;margin:0;">
        Jira Worklog Kontrol Sistemi — Otomatik Rapor<br>
        {datetime.now().strftime('%d.%m.%Y %H:%M')}</p>
    </div>
  </div>
</body>
</html>
"""