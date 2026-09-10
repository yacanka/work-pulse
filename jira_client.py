# jira_client.py

import logging
from datetime import date, datetime, timedelta
from typing import Optional

from jira import JIRA
from jira.exceptions import JIRAError

from models import WorklogEntry, IssueInfo

logger = logging.getLogger(__name__)


class JiraClient:
    """Jira API ile iletişim kuran istemci sınıfı."""

    def __init__(self, server: str, username: str, password: str):
        self.server = server
        try:
            self.jira = JIRA(
                server=server,
                basic_auth=(username, password),
                options={"verify": "JIRA_Chain.crt"}
            )
            logger.info(f"Jira bağlantısı başarılı: {server}")
        except JIRAError as e:
            logger.error(f"Jira bağlantı hatası: {e}")
            raise

    def build_jql(
        self,
        users: list[str],
        projects: list[str],
        issue_types: list[str],
        statuses: list[str],
        require_due_date: bool,
        require_assignee: bool,
        end_date: str,
        # start_date ve end_date parametreleri KALDIRILDI
    ) -> str:
        """Filtreleme kriterlerine göre JQL sorgusu oluşturur (worklog tarihi HARİÇ)."""
        conditions = []

        if projects:
            proj_str = ", ".join(f'"{p}"' for p in projects)
            conditions.append(f"project IN ({proj_str})")

        if issue_types:
            types_str = ", ".join(f'"{t}"' for t in issue_types)
            conditions.append(f"issuetype IN ({types_str})")

        if statuses:
            status_str = ", ".join(f'"{s}"' for s in statuses)
            conditions.append(f"status IN ({status_str})")

        if users:
            users_str = ", ".join(f'"{u}"' for u in users)
            conditions.append(f"assignee IN ({users_str})")

        if require_due_date:
            conditions.append("duedate IS NOT EMPTY")
            conditions.append(f'duedate >= "{end_date}"')

        # ❌ KALDIRILDI: worklogDate filtreleri artık yok
        # conditions.append(f'worklogDate >= "{start_date}"')
        # conditions.append(f'worklogDate <= "{end_date}"')

        jql = " AND ".join(conditions)
        jql += " ORDER BY assignee ASC, key ASC"

        logger.info(f"Oluşturulan JQL: {jql}")
        return jql


    def search_issues(self, jql: str) -> list[IssueInfo]:
        """JQL sorgusuna göre issue'ları arar."""
        issues = []
        start_at = 0
        max_results = 100

        while True:
            try:
                results = self.jira.search_issues(
                    jql,
                    startAt=start_at,
                    maxResults=max_results,
                    # ── GÜNCELLEME: timetracking eklendi ──
                    fields="summary,issuetype,status,assignee,"
                           "duedate,parent,worklog,timetracking",
                )
            except JIRAError as e:
                logger.error(f"Issue arama hatası: {e}")
                raise

            for issue in results:
                f = issue.fields

                assignee = None
                if f.assignee:
                    assignee = getattr(
                        f.assignee, "accountId",
                        getattr(f.assignee, "name", str(f.assignee))
                    )

                due = None
                if f.duedate:
                    due = datetime.strptime(f.duedate, "%Y-%m-%d").date()

                parent_key = None
                parent_summary = None
                if hasattr(f, "parent") and f.parent:
                    parent_key = f.parent.key
                    parent_summary = f.parent.fields.summary

                # ── YENİ: Workload alanları ──
                tt = getattr(f, "timetracking", None)
                original_est = 0
                remaining_est = 0
                time_spent = 0
                if tt:
                    original_est = getattr(
                        tt, "originalEstimateSeconds", 0
                    ) or 0
                    remaining_est = getattr(
                        tt, "remainingEstimateSeconds", 0
                    ) or 0
                    time_spent = getattr(
                        tt, "timeSpentSeconds", 0
                    ) or 0

                issues.append(IssueInfo(
                    key=issue.key,
                    summary=f.summary,
                    issue_type=str(f.issuetype),
                    status=str(f.status),
                    assignee=assignee,
                    due_date=due,
                    parent_key=parent_key,
                    parent_summary=parent_summary,
                    original_estimate_seconds=original_est,
                    remaining_estimate_seconds=remaining_est,
                    time_spent_seconds=time_spent,
                ))

            if start_at + max_results >= results.total:
                break
            start_at += max_results

        logger.info(f"Toplam {len(issues)} issue bulundu")
        return issues

    def get_issue_worklogs(
        self,
        issue_key: str,
        start_date: date,
        end_date: date,
        target_users: list[str],
    ) -> list[WorklogEntry]:
        """
        Bir issue'nun worklog'larını getirir.
        Tarih aralığı ve kullanıcı filtrelemesi yapar.
        """
        entries = []

        try:
            issue = self.jira.issue(issue_key)
            worklogs = self.jira.worklogs(issue_key)
        except JIRAError as e:
            logger.error(f"Worklog getirme hatası ({issue_key}): {e}")
            return entries

        for wl in worklogs:
            # Tarih parse
            started = datetime.strptime(
                wl.started[:10], "%Y-%m-%d"
            ).date()

            # Tarih filtresi
            if started < start_date or started > end_date:
                continue

            # Yazar bilgisi
            author_id = getattr(
                wl.author, "accountId",
                getattr(wl.author, "name", str(wl.author))
            )
            author_display = getattr(
                wl.author, "displayName", author_id
            )

            # Kullanıcı filtresi
            if target_users and author_id not in target_users:
                # displayName ile de kontrol et
                if author_display not in target_users:
                    continue

            comment = getattr(wl, "comment", "") or ""

            entries.append(WorklogEntry(
                issue_key=issue_key,
                issue_summary=issue.fields.summary,
                issue_type=str(issue.fields.issuetype),
                issue_status=str(issue.fields.status),
                author=author_id,
                author_display_name=author_display,
                date=started,
                time_spent_seconds=wl.timeSpentSeconds,
                comment=comment,
                worklog_id=wl.id,
            ))

        return entries

    def find_user_issues_without_criteria(
        self,
        username: str,
        projects: list[str],
        issue_types: list[str],
        statuses: list[str],
        require_due_date: bool,
    ) -> dict:
        """
        Kullanıcının kriterlere UYMAYAN issue'larını bulur.
        Uyarı mesajları oluşturmak için kullanılır.
        """
        result = {
            "wrong_type": [],
            "wrong_status": [],
            "no_due_date": [],
        }

        # Kullanıcının tüm issue'larını getir
        base_jql = f'assignee = "{username}"'
        if projects:
            proj_str = ", ".join(f'"{p}"' for p in projects)
            base_jql += f" AND project IN ({proj_str})"

        try:
            issues = self.jira.search_issues(
                base_jql,
                maxResults=200,
                fields="summary,issuetype,status,duedate",
            )
        except JIRAError:
            return result

        for issue in issues:
            f = issue.fields
            itype = str(f.issuetype)
            status = str(f.status)

            if issue_types and itype not in issue_types:
                result["wrong_type"].append(
                    f"  {issue.key}: tür='{itype}' "
                    f"(beklenen: {issue_types})"
                )

            if statuses and status not in statuses:
                result["wrong_status"].append(
                    f"  {issue.key}: statü='{status}' "
                    f"(beklenen: {statuses})"
                )

            if require_due_date and not f.duedate:
                result["no_due_date"].append(
                    f"  {issue.key}: due date boş"
                )

        return result

    def get_user_account_id(self, query: str) -> Optional[str]:
        """Kullanıcı adından account ID bulur (Cloud için)."""
        try:
            users = self.jira.search_users(query=query, maxResults=1)
            if users:
                return users[0].accountId
        except Exception:
            pass
        return None