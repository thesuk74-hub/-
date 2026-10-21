"""가동률 달력 — 공정별 기상프로필에 따른 일(日)별 작업가능 비율.

가이드라인 제2장 3.(3) '공종별 공사기간 = 작업일수 + 비작업일수'
(개시일부터 월별 작업가능일수를 차감해 가며 종료일 산정, 마지막 달은 비례 배분)을
일 단위 가동률(작업가능일수 ÷ 달력일수)의 누적으로 동일하게 구현한다.
"""
import datetime as dt
from bisect import bisect_left, bisect_right


class CapacityCalendar:
    def __init__(self, setup, start_date, horizon_days=365 * 15):
        self.setup = setup
        self.start = start_date
        self.horizon = horizon_days
        self.dates = [start_date + dt.timedelta(days=i) for i in range(horizon_days)]
        self._cum = {}
        self._month_cache = {}

    def month_info(self, profile, y, m):
        key = (profile, y, m)
        if key not in self._month_cache:
            self._month_cache[key] = self.setup.month(profile, y, m)
        return self._month_cache[key]

    def cum(self, profile):
        """cum[i] = 0일차~(i-1)일차 작업가능일 누계."""
        if profile not in self._cum:
            c = [0.0]
            for d in self.dates:
                c.append(c[-1] + self.month_info(profile, d.year, d.month).rate)
            self._cum[profile] = c
        return self._cum[profile]

    def cap(self, profile, day):
        d = self.dates[day]
        return self.month_info(profile, d.year, d.month).rate

    def finish(self, profile, start, work):
        """start일차에 착수하여 작업일수 work를 소진하는 종료일차(포함)."""
        if work <= 0:
            return start - 1
        c = self.cum(profile)
        target = c[start] + work - 1e-9
        e = bisect_left(c, target, lo=start + 1)
        if e >= len(c):
            raise OverflowError("공사기간이 계산범위(15년)를 초과합니다. 입력값을 확인하십시오.")
        return e - 1

    def start_for_finish(self, profile, finish, work):
        """finish일차(포함)에 종료되도록 하는 가장 늦은 착수일차."""
        if work <= 0:
            return finish + 1
        c = self.cum(profile)
        target = c[finish + 1] - work + 1e-9
        s = bisect_right(c, target) - 1   # 가장 큰 s : c[s] <= target
        return max(min(s, finish), 0)

    def day_to_date(self, day):
        return self.start + dt.timedelta(days=int(day))

    def date_to_day(self, date):
        return (date - self.start).days
