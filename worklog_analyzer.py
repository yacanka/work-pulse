# worklog_analyzer.py

import logging
from datetime import date, timedelta
from collections import defaultdict

from models import (
    WorklogEntry, DailySummary, UserReport, IssueInfo
)
from jira_client import JiraClient

logger = logging.getLogger(__name__)


class WorklogAnalyzer:
    """Worklog verilerini analiz eder ve raporlar oluşturur."""

    def __init__(self, client: JiraClient, config: dict):
        self.client = client
        self.config = config
        self.filters = config["issue_filters"]
        self.rules = config["worklog_rules"]
        self.date_start = self._parse_date(config["date_range"]["start_date"])
        self.date_end = self._parse_date(config["date_range"]["end_date"])

    @staticmethod
    def _parse_date(d) -> date:
        if isinstance(d, date):
            return d
        from datetime import datetime
        return datetime.strptime(str(d), "%d.%m.%Y").date()

    def _get_all_dates(self) -> list[date]:
        """Tarih aralığındaki tüm günleri döndürür."""
        dates = []
        current = self.date_start
        while current <= self.date_end:
            dates.append(current)
            current += timedelta(days=1)
        return dates

    def _get_working_dates(self) -> list[date]:
        """Tarih aralığındaki iş günlerini döndürür."""
        dates = self._get_all_dates()
        if not self.rules.get("check_weekends", False):
            dates = [d for d in dates if d.weekday() < 5]
        return dates

    def _resolve_usernames(self) -> dict[str, str]:
        """
        Config'deki kullanıcıları çözümler.
        {username_or_id: display_name} döndürür.
        """
        users = {}
        for u in self.config["users"]:
            uname = u["username"]
            dname = u.get("display_name", uname)

            # Cloud ortamında accountId'ye çevirmeyi dene
            """ account_id = self.client.get_user_account_id(uname)
            if account_id:
                users[account_id] = dname
                logger.info(f"Kullanıcı çözümlendi: {uname} -> {account_id}")
            else:
                users[uname] = dname
                logger.info(f"Kullanıcı doğrudan kullanılıyor: {uname}") """
            users[uname] = dname
            logger.info(f"Okunan çalışan adı: {uname}")

        return users

    def analyze(self) -> list[UserReport]:
        """Ana analiz fonksiyonu. Tüm kullanıcılar için rapor üretir."""
        users = self._resolve_usernames()
        user_ids = list(users.keys())
        working_dates = self._get_working_dates()
        all_dates = self._get_all_dates()

        reports: list[UserReport] = []

        # ── 1) JQL ile eşleşen issue'ları bul (tarih koşulu YOK) ──
        jql = self.client.build_jql(
            users=user_ids,
            projects=self.filters.get("projects", []),
            issue_types=self.filters.get("issue_types", []),
            statuses=self.filters.get("statuses", []),
            require_due_date=self.filters.get("require_due_date", False),
            require_assignee=self.filters.get("require_assignee", True),
            end_date=self.date_end.strftime("%Y-%m-%d")
            # start_date, end_date YOK artık
        )

        matching_issues = self.client.search_issues(jql)
        logger.info(
            f"Kriterlere uyan toplam {len(matching_issues)} issue bulundu"
        )

        # Issue'ları kullanıcıya göre grupla
        user_issues: dict[str, list[IssueInfo]] = defaultdict(list)
        for issue in matching_issues:
            if issue.assignee:
                user_issues[issue.assignee].append(issue)

        # ── 2) Her kullanıcı için analiz ──
        for user_id, display_name in users.items():
            logger.info(f"\n{'='*50}")
            logger.info(f"Çalışan analiz ediliyor: {display_name}")

            report = UserReport(
                username=user_id,
                display_name=display_name,
            )

            issues = user_issues.get(user_id, [])

            # ── Uyarı: Kriterlere uyan issue yok ──
            if not issues:
                report.warnings.append(
                    f"⚠️  {display_name} için kriterlere uyan "
                    f"issue bulunamadı!"
                )
                report.warnings.append(
                    f"   Kriterler: tür={self.filters.get('issue_types')}, "
                    f"statü={self.filters.get('statuses')}, "
                    f"due_date={'zorunlu' if self.filters.get('require_due_date') else 'opsiyonel'}"
                )

                diag = self.client.find_user_issues_without_criteria(
                    username=user_id,
                    projects=self.filters.get("projects", []),
                    issue_types=self.filters.get("issue_types", []),
                    statuses=self.filters.get("statuses", []),
                    require_due_date=self.filters.get(
                        "require_due_date", False
                    ),
                )
                if diag["wrong_type"]:
                    report.warnings.append("   ❌ Yanlış türdeki issue'lar:")
                    report.warnings.extend(diag["wrong_type"][:5])
                if diag["wrong_status"]:
                    report.warnings.append("   ❌ Yanlış statüdeki issue'lar:")
                    report.warnings.extend(diag["wrong_status"][:5])
                if diag["no_due_date"]:
                    report.warnings.append("   ❌ Due date'i olmayan issue'lar:")
                    report.warnings.extend(diag["no_due_date"][:5])

                reports.append(report)
                continue

            # ── 3) Worklog'ları topla ve issue'ları sınıflandır ──
            issues_with_worklog = []     # tarih aralığında worklog'u OLAN
            issues_without_worklog = []  # tarih aralığında worklog'u OLMAYAN
            all_worklogs: list[WorklogEntry] = []

            for iss in issues:
                wls = self.client.get_issue_worklogs(
                    issue_key=iss.key,
                    start_date=self.date_start,
                    end_date=self.date_end,
                    target_users=[user_id],
                )
                if wls:
                    issues_with_worklog.append(iss)
                    all_worklogs.extend(wls)
                    logger.info(f"  ✅ {iss.key}: {len(wls)} worklog bulundu")
                else:
                    issues_without_worklog.append(iss)
                    logger.info(f"  ❌ {iss.key}: worklog YOK (tarih aralığında)")

            for iss in issues_with_worklog:
                report.matching_issues.append({
                    "key": iss.key,
                    "summary": iss.summary,
                    "type": iss.issue_type,
                    "status": iss.status,
                    "due_date": iss.due_date.strftime("%d.%m.%Y") if iss.due_date else "-",
                    "parent": iss.parent_key or "-",
                    "has_worklog": True,
                    # ── YENİ: Workload bilgileri ──
                    "original_estimate": iss.original_estimate_seconds,
                    "remaining_estimate": iss.remaining_estimate_seconds,
                    "time_spent": iss.time_spent_seconds,
                    "workload_progress": iss.workload_progress,
                })

            for iss in issues_without_worklog:
                report.matching_issues.append({
                    "key": iss.key,
                    "summary": iss.summary,
                    "type": iss.issue_type,
                    "status": iss.status,
                    "due_date": iss.due_date.strftime("%d.%m.%Y") if iss.due_date else "-",
                    "parent": iss.parent_key or "-",
                    "has_worklog": False,
                    "original_estimate": iss.original_estimate_seconds,
                    "remaining_estimate": iss.remaining_estimate_seconds,
                    "time_spent": iss.time_spent_seconds,
                    "workload_progress": iss.workload_progress,
                })

            report.total_original_estimate = sum(
                iss.original_estimate_seconds for iss in issues
            )
            report.total_remaining_estimate = sum(
                iss.remaining_estimate_seconds for iss in issues
            )
            report.total_logged_on_issues = sum(
                iss.time_spent_seconds for iss in issues
            )

            # ── Worklog'u olmayan issue'lar için uyarı ──
            if issues_without_worklog:
                report.warnings.append(
                    f"⚠️  {len(issues_without_worklog)} issue'da "
                    f"({self.date_start} - {self.date_end}) "
                    f"aralığında worklog yok:"
                )
                for iss in issues_without_worklog:
                    report.warnings.append(
                        f"     ❌ {iss.key} - {iss.summary}"
                    )

            # ── 4) Gün bazında grupla ──
            for d in all_dates:
                summary = DailySummary(
                    user=user_id,
                    display_name=display_name,
                    date=d,
                )
                report.daily_summaries[d] = summary

            for wl in all_worklogs:
                if wl.date in report.daily_summaries:
                    ds = report.daily_summaries[wl.date]
                    ds.worklogs.append(wl)
                    ds.total_seconds += wl.time_spent_seconds

            # ── 5) Kural kontrolleri ──
            min_h = self.rules.get("min_daily_hours", 0)
            max_h = self.rules.get("max_daily_hours", 24)
            check_weekends = self.rules.get("check_weekends", False)

            for d in working_dates:
                ds = report.daily_summaries.get(d)
                if not ds:
                    continue

                if ds.total_seconds == 0:
                    report.warnings.append(
                        f"⚠️  {d.strftime('%d.%m.%Y')} "
                        f"({self._day_name(d)}): "
                        f"Hiç worklog girilmemiş!"
                    )
                elif ds.total_hours < min_h:
                    report.warnings.append(
                        f"⚠️  {d.strftime('%d.%m.%Y')} "
                        f"({self._day_name(d)}): "
                        f"{ds.total_display} çalışılmış "
                        f"(minimum {min_h}s)"
                    )
                elif ds.total_hours > max_h:
                    report.warnings.append(
                        f"⚠️  {d.strftime('%d.%m.%Y')} "
                        f"({self._day_name(d)}): "
                        f"{ds.total_display} çalışılmış "
                        f"(maksimum {max_h}s)"
                    )

            if not check_weekends:
                for d in all_dates:
                    ds = report.daily_summaries.get(d)
                    if ds and ds.is_weekend and ds.total_seconds > 0:
                        report.warnings.append(
                            f"ℹ️  {d.strftime('%d.%m.%Y')} "
                            f"({self._day_name(d)}): "
                            f"Hafta sonu {ds.total_display} çalışılmış"
                        )

            reports.append(report)

        return reports

    @staticmethod
    def _day_name(d: date) -> str:
        days = [
            "Pazartesi", "Salı", "Çarşamba",
            "Perşembe", "Cuma", "Cumartesi", "Pazar"
        ]
        return days[d.weekday()]