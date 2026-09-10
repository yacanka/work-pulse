# models.py

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Optional


@dataclass
class WorklogEntry:
    """Tek bir worklog kaydını temsil eder."""
    issue_key: str
    issue_summary: str
    issue_type: str
    issue_status: str
    author: str
    author_display_name: str
    date: date
    time_spent_seconds: int
    comment: str = ""
    worklog_id: str = ""

    @property
    def hours(self) -> float:
        return round(self.time_spent_seconds / 3600, 2)

    @property
    def time_spent_display(self) -> str:
        h = self.time_spent_seconds // 3600
        m = (self.time_spent_seconds % 3600) // 60
        if h and m:
            return f"{h}s {m}dk"
        elif h:
            return f"{h}s"
        return f"{m}dk"


@dataclass
class DailySummary:
    """Bir kullanıcının belirli bir gündeki worklog özeti."""
    user: str
    display_name: str
    date: date
    total_seconds: int = 0
    worklogs: list[WorklogEntry] = field(default_factory=list)

    @property
    def total_hours(self) -> float:
        return round(self.total_seconds / 3600, 2)

    @property
    def total_display(self) -> str:
        h = self.total_seconds // 3600
        m = (self.total_seconds % 3600) // 60
        return f"{h}s {m}dk"

    @property
    def is_weekend(self) -> bool:
        return self.date.weekday() >= 5


@dataclass
class IssueInfo:
    """Issue detay bilgisi."""
    key: str
    summary: str
    issue_type: str
    status: str
    assignee: Optional[str]
    due_date: Optional[date]
    parent_key: Optional[str] = None
    parent_summary: Optional[str] = None
    # ── YENİ: Workload alanları ──
    original_estimate_seconds: int = 0
    remaining_estimate_seconds: int = 0
    time_spent_seconds: int = 0

    @property
    def original_estimate_hours(self) -> float:
        return round(self.original_estimate_seconds / 3600, 2)

    @property
    def remaining_estimate_hours(self) -> float:
        return round(self.remaining_estimate_seconds / 3600, 2)

    @property
    def time_spent_hours(self) -> float:
        return round(self.time_spent_seconds / 3600, 2)

    @property
    def workload_progress(self) -> float:
        """Tamamlanma yüzdesi (time spent / original estimate)."""
        if self.original_estimate_seconds == 0:
            return 0.0
        return round(
            (self.time_spent_seconds / self.original_estimate_seconds) * 100, 1
        )

    @staticmethod
    def format_duration(seconds: int) -> str:
        if seconds <= 0:
            return "-"
        h = seconds // 3600
        m = (seconds % 3600) // 60
        if h and m:
            return f"{h}s {m}dk"
        elif h:
            return f"{h}s"
        return f"{m}dk"


@dataclass
class UserReport:
    """Bir kullanıcının tüm tarih aralığı için raporu."""
    username: str
    display_name: str
    daily_summaries: dict[date, DailySummary] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)
    matching_issues: list[dict] = field(default_factory=list)
    # ── YENİ: Workload toplam bilgileri ──
    total_original_estimate: int = 0
    total_remaining_estimate: int = 0
    total_logged_on_issues: int = 0

    @property
    def total_hours(self) -> float:
        return round(
            sum(ds.total_hours for ds in self.daily_summaries.values()), 2
        )

    @property
    def total_days_worked(self) -> int:
        return sum(
            1 for ds in self.daily_summaries.values() if ds.total_seconds > 0
        )

    # ── YENİ ──
    @property
    def total_working_days(self) -> int:
        """Hafta içi gün sayısı."""
        return sum(
            1 for ds in self.daily_summaries.values()
            if ds.date.weekday() < 5
        )

    @property
    def success_percentage(self) -> float:
        """Worklog başarı yüzdesi: worklog girilen iş günü / toplam iş günü * 100."""
        working_days = self.total_working_days
        if working_days == 0:
            return 0.0
        days_with_worklog = sum(
            1 for ds in self.daily_summaries.values()
            if ds.total_seconds > 0 and ds.date.weekday() < 5
        )
        return round((days_with_worklog / working_days) * 100, 1)

    @property
    def success_status(self) -> str:
        pct = self.success_percentage
        if pct >= 90:
            return "Tam Uyumlu Giriş"
        elif pct >= 80:
            return "Kabul Edilebilir Giriş"
        elif pct >= 70:
            return "Artırılması Gereken Giriş"
        else:
            return "Yetersiz Giriş"

    @property
    def success_color(self) -> str:
        """Colorama renk kodu."""
        pct = self.success_percentage
        if pct >= 90:
            return "GREEN"
        elif pct >= 80:
            return "YELLOW"
        elif pct >= 70:
            return "ORANGE"
        return "RED"