"""② 비작업일수 산정 — 가이드라인 제2장 2.(3)

    비작업일수 = A + B - C,  C = round(A × B ÷ 달력일수)   (소수점 첫째자리에서 반올림)
    A : 기상조건으로 인한 월별 비작업일수 (부록3 관측지점 통계 × 공정 기상프로필 적용률)
    B : 법정공휴일수 (부록1)
    월별 비작업일수 < 주40시간 근무제 비작업일수(월 8일) 이면 8일 적용
"""
import calendar
from dataclasses import dataclass, field

from . import refdata as R


def round_half_up(x, nd=0):
    m = 10 ** nd
    return int(x * m + 0.5 + 1e-9) / m if x >= 0 else -round_half_up(-x, nd)


@dataclass
class MonthNonwork:
    year: int
    month: int
    days: int
    A: float                 # 기상 비작업일수
    A_detail: dict           # 조건번호 → 일수(적용률 반영)
    B: int                   # 법정공휴일
    C: int                   # 중복일수
    raw: float               # A+B-C
    applied: float           # 주40시간 하한·반올림·동절기 중지 반영
    winter_stop: bool = False

    @property
    def workable(self):
        return max(self.days - self.applied, 0.0)

    @property
    def rate(self):
        return self.workable / self.days


@dataclass
class NonworkSetup:
    station: str
    profiles: dict                     # 프로필명 → {조건번호: 적용률}
    dust: list                         # 미세먼지 월별 일수(12)
    floor_days: float = 8.0            # 주 40시간 근무제 하한
    round_int: bool = True             # 월별 비작업일수 정수 반올림
    winter_stop: bool = False          # 동절기 공사중지 반영 여부
    winter_months: tuple = (12, 1, 2)
    winter_profiles: tuple = None      # None=전체 프로필 중지
    no_holiday_profiles: set = field(default_factory=lambda: set(R.NO_HOLIDAY_PROFILES))
    winter_exempt: set = field(default_factory=set)   # 동절기 중지 미적용 프로필(시운전 등)
    custom: list = field(default_factory=lambda: [0.0] * 12)   # 사용자정의 기상조건(15) 월별 일수
    custom_label: str = "사용자정의"

    def station_data(self):
        st = R.WEATHER["stations"].get(self.station)
        if st is None:
            raise ValueError(f"기상관측지점 '{self.station}'이(가) 부록3 DB에 없습니다. 사용가능: {', '.join(sorted(R.WEATHER['stations']))}")
        return st

    def cond_value(self, cond, month):
        if cond == 14:
            return float(self.dust[month - 1] or 0)
        if cond == 15:
            return float(self.custom[month - 1] or 0)
        return self.station_data()["cond"][str(cond)][month - 1]

    def month(self, profile, year, month):
        days = calendar.monthrange(year, month)[1]
        weights = self.profiles.get(profile, {})
        detail = {c: round(self.cond_value(c, month) * w, 2) for c, w in weights.items() if w}
        A = round(sum(detail.values()), 1)
        if profile in self.no_holiday_profiles or profile == R.PREP_PROFILE:
            mn = MonthNonwork(year, month, days, 0.0, {}, 0, 0, 0.0, 0.0)
            if self.winter_stop and month in self.winter_months and profile not in self.winter_exempt:
                mn.applied, mn.winter_stop = float(days), True
            return mn
        B = R.holidays(year, month)
        C = int(round_half_up(A * B / days))
        raw = round(A + B - C, 1)
        applied = max(raw, self.floor_days)
        if self.round_int:
            applied = float(round_half_up(applied))
        stop = False
        if self.winter_stop and month in self.winter_months and profile not in self.winter_exempt:
            applied, stop = float(days), True
        return MonthNonwork(year, month, days, A, detail, B, C, raw, applied, stop)

    def annual_climate(self, profile):
        """프로필별 월평균 기상 비작업일수(A) 12개월 (표 2-5 형식)."""
        out = {}
        for c, w in self.profiles[profile].items():
            if w:
                out[c] = [round(self.cond_value(c, m) * w, 1) for m in range(1, 13)]
        return out
