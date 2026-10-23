"""③·④ 작업일수 산정 — 가이드라인 제2장 2.(4), 부록4

    단위작업 작업일수 = ROUNDUP( 작업수량 ÷ (1일 작업량 × 보정계수 × 작업조수) )   (1일 미만은 1일)
    공정 작업일수     = Σ 단위작업 작업일수(중복가능 제외) + 간접작업일수 + 보정일수 + 직접입력일수

1일 작업량은 ① 수동입력(품셈·실적) ② DB 지정(번호/공종명) ③ 공종명·규격·단위 자동매칭 순으로 결정한다.
보정계수 = 현장여건 × 공사규모 × 지질조건(지반 관련 공종) × 기상기후(옥외 공정) × 개별계수
"""
import math
import re
from dataclasses import dataclass, field

from . import refdata as R

DB = R.PRODUCTIVITY["items"]
DB_BY_NO = {it["no"]: it for it in DB}

UNIT_ALIAS = {"m3": "㎥", "m³": "㎥", "공㎥": "㎥", "m2": "㎡", "m²": "㎡", "공㎡": "㎡", "톤": "ton", "t": "ton",
              "ea": "개소", "EA": "개소", "개": "개소", "개소": "개소", "주": "주", "M": "m", "본": "본",
              "km": "km", "식": "식", "경간": "경간", "span": "span", "Lot": "Lot", "회": "회"}


def norm_unit(u):
    u = (u or "").strip().split("/")[0].strip()
    return UNIT_ALIAS.get(u, UNIT_ALIAS.get(u.lower(), u))


# 자동매칭 규칙 (공종명+규격 정규식 → DB 공종명). 위에서부터 우선 적용
RULES = [
    (r"말뚝박기용\s*천공|파일\s*천공|H-?PILE\s*천공", "가설흙막이(말뚝박기용천공)"),
    (r"강널말뚝.*(항타|근입|설치)|SHEET\s*PILE", "가설흙막이(강널말뚝)"),
    (r"(터파기|굴착).*(연암|경암|보통암|발파암)", "터파기(발파암)"),
    (r"(터파기|굴착).*(풍화암|리핑)", "터파기(리핑암)"),
    (r"(터파기|굴착).*(토사|보통토|사질)", "터파기(토사)"),
    (r"(흙깎기|흙깍기).*(토사)", "흙깍기(토사)"),
    (r"(흙깎기|흙깍기).*(리핑|풍화)", "흙깍기(리핑암)"),
    (r"되메우기", "되메우기"),
    (r"기초지정|모래지정", "기초지정"),
    (r"표토\s*제거", "표토제거"),
    (r"벌개\s*제근", "벌개제근"),
    (r"콘크리트.*타설.*무근|무근.*타설", "콘크리트타설(무근,f1-Ⅰ)"),
    (r"콘크리트\s*타설|타설.*철근|레미콘\s*타설", "콘크리트타설(철근,f1-Ⅱ)"),
    (r"유로폼", "유로폼 설치(보통)"),
    (r"합판\s*거푸집", "합판거푸집 설치(보통)"),
    (r"시스템\s*동바리", "시스템동바리"),
    (r"강관\s*동바리", "강관동바리(토목)"),
    (r"시스템\s*비계", "시스템비계"),
    (r"강관.*비계", "강관비계"),
    (r"철근.*가공.*조립|철근.*조립", "철근조립(토목)"),
    (r"철근.*가공", "철근가공"),
    (r"포장.*절단|ASP.*절단", "포장절단"),
    (r"(아스콘|아스팔트).*깨기|포장\s*깨기", "기존시설물 깨기"),
    (r"(아스팔트|아스콘).*(표층|중간층)", "아스팔트 중간/표층"),
    (r"(아스팔트|아스콘).*기층", "아스팔트 기층(5~7㎝)"),
    (r"보조기층", "아스팔트 보조기층"),
    (r"(택|텍)\s*코팅", "택코팅(본선)"),
    (r"프라임\s*코팅", "프라임코팅(본선)"),
    (r"(흄관|VR관|V\.R)", "배수관(흄관/VR)"),
    (r"(PVC|PE|이중벽|고강성|하수관|오수관|우수관|폴리에틸렌|파형).*(부설|설치|접합)", "배수관(흄관/VR)"),
    (r"맨홀\s*뚜껑", "맨홀뚜껑 설치"),
    (r"맨홀.*(설치|부설)|(설치).*맨홀", "조립식 PC맨홀 설치"),
    (r"(오수|우수|빗물)\s*받이", "PE 우수받이 설치"),
    (r"벽돌\s*쌓기", "벽돌쌓기"),
    (r"블록.*쌓기", "블록보강쌓기"),
    (r"모르타르\s*바름|미장", "시멘트모르타르 바름"),
    (r"도막\s*방수", "도막방수"),
    (r"시트\s*방수", "시트방수"),
    (r"액체\s*방수", "시멘트액체방수"),
    (r"철골.*세우기", "철골 현장세우기"),
    (r"데크\s*플레이트", "데크플레이트 설치"),
    (r"각형\s*덕트", "각형덕트(0.6mm)"),
    (r"H-?BEAM.*철거|H-?PILE.*(뽑기|인발)", "H-Beam 철거"),
    (r"PHC|기성\s*말뚝", "기성말뚝기초(PHC)"),
    (r"차선\s*도색", "차선도색"),
    (r"가설\s*울타리", "가설울타리"),
    (r"가드레일", "가드레일"),
]
NAME_INDEX = {}
for _it in DB:
    NAME_INDEX.setdefault(_it["name"], _it)


def _bigrams(s):
    s = re.sub(r"[\s()\[\],·./\-_~]", "", s or "")
    return {s[i:i + 2] for i in range(len(s) - 1)} or {s}


def dice(a, b):
    A, B = _bigrams(a), _bigrams(b)
    return 2 * len(A & B) / (len(A) + len(B)) if A and B else 0.0


def find_db(hint):
    """DB 지정값(번호 또는 공종명)으로 항목 조회."""
    if hint is None or str(hint).strip() == "":
        return None
    h = str(hint).strip()
    if re.fullmatch(r"\d+(\.0)?", h):
        return DB_BY_NO.get(int(float(h)))
    if h in NAME_INDEX:
        return NAME_INDEX[h]
    best = max(DB, key=lambda it: dice(h, it["name"]))
    return best if dice(h, best["name"]) >= 0.6 else None


def auto_match(name, spec, unit, field_name=""):
    """자동매칭: (DB항목, 근거문자열, 신뢰도) 또는 (None, 사유, 0)."""
    text = f"{name} {spec or ''}"
    u = norm_unit(unit)
    for pat, dbname in RULES:
        if re.search(pat, text, re.I):
            it = NAME_INDEX.get(dbname)
            if dbname == "철근조립(토목)" and field_name == "건축":
                it = NAME_INDEX["철근조립(건축Ⅱ)"]
            if it and norm_unit(it["unit"]) == u:
                return it, f"규칙매칭('{pat}')", 0.9
    cands = [it for it in DB if norm_unit(it["unit"]) == u]
    if not cands:
        return None, f"단위 '{unit}'와 일치하는 DB 항목 없음", 0.0
    scored = sorted(((dice(text, it["name"] + it["condition"]), it) for it in cands), key=lambda x: -x[0])
    s, it = scored[0]
    if s >= 0.45:
        return it, f"유사도매칭({s:.2f})", round(s, 2)
    return None, f"매칭불가(최고유사도 {s:.2f}: {it['name']})", 0.0


@dataclass
class Factors:
    site: str = "보통"
    scale: str = "보통"
    geology: str = "양호(토사)"
    climate: str = "일반"
    tables: dict = field(default_factory=lambda: {k: dict(v) for k, v in R.ADJUST_FACTORS.items()})

    def value(self, item, profile):
        t = self.tables
        f_site = t["현장여건"].get(self.site, 1.0)
        f_scale = t["공사규모"].get(self.scale, 1.0)
        grp = f"{item.group} {item.db['group'] if item.db else ''} {item.name}"
        f_geo = t["지질조건"].get(self.geology, 1.0) if any(g in grp for g in R.GEO_GROUPS) else 1.0
        f_cli = t["기상기후"].get(self.climate, 1.0) if profile in ("옥외공사", "구조물공사") else 1.0
        return f_site, f_scale, f_geo, f_cli


@dataclass
class Item:
    no: int
    act: str
    name: str
    spec: str
    unit: str
    qty: float
    db_hint: str = None
    manual_rate: float = None
    crews: float = 1
    max_crews: float = None
    item_factor: float = 1.0
    indirect: float = 0.0
    correction: float = 0.0
    parallel: bool = False          # 중복가능(타 작업과 동시진행 → 공정 작업일수 미합산)
    step: float = 1                 # 작업일수 산정단위(1일 또는 0.1일)
    note: str = ""
    group: str = ""
    # 계산 결과
    db: dict = None
    source: str = ""
    match_note: str = ""
    confidence: float = 0.0
    base_rate: float = None
    factors: tuple = (1.0, 1.0, 1.0, 1.0)
    factor: float = 1.0
    rate: float = None
    days: int = 0
    formula: str = ""
    warn: str = ""

    def resolve(self, field_name, profile, fac: Factors):
        self.warn = ""
        if self.manual_rate:
            self.base_rate, self.source, self.match_note, self.confidence = float(self.manual_rate), "품셈·실적(수동입력)", "수동입력", 1.0
            if self.db_hint:
                self.db = find_db(self.db_hint)
        else:
            it = find_db(self.db_hint)
            if it:
                self.db, self.match_note, self.confidence = it, "DB지정", 1.0
            elif self.db_hint:
                self.warn = f"DB지정 '{self.db_hint}' 항목을 찾을 수 없음 → 자동매칭 시도"
            if not self.db:
                self.db, self.match_note, self.confidence = auto_match(self.name, self.spec, self.unit, field_name)
            if self.db:
                self.base_rate, self.source = float(self.db["rate"]), self.db["source"]
                if norm_unit(self.db["unit"]) != norm_unit(self.unit):
                    self.warn += f" 단위불일치(수량 {self.unit} / DB {self.db['unit']})"
        self.factors = fac.value(self, profile)
        self.factor = round(self.factors[0] * self.factors[1] * self.factors[2] * self.factors[3] * (self.item_factor or 1.0), 4)
        if not self.base_rate:
            self.rate, self.days = None, 0
            if not self.parallel:
                self.warn = (self.warn + " 1일작업량 미결정 — 수동1일작업량 또는 DB지정 필요").strip()
            else:
                self.match_note = "중복가능(작업일수 미합산)"
            self.formula = "-"
            return
        self.rate = self.base_rate * self.factor
        self.recalc()

    def recalc(self):
        crews = float(self.crews or 1)
        crews = int(crews) if crews == int(crews) else round(crews, 2)   # 구간분할 배수 등 소수 작업조 허용
        raw = self.qty / (self.rate * crews) if self.qty else 0
        st = self.step or 1
        if raw > 0:
            d = math.ceil(raw / st - 1e-9) * st
            self.days = int(round(d)) if st >= 1 else round(d, 1)
        else:
            self.days = 0
        unit = "" if st >= 1 else f",{st:g}일단위"
        self.formula = f"ROUNDUP({self.qty:g}÷({self.base_rate:g}×{self.factor:g}×{crews}){unit})={self.days:g}"

    @property
    def total(self):
        return 0 if self.parallel else self.days + (self.indirect or 0) + (self.correction or 0)


def activity_workdays(act, items):
    """공정(Activity) 작업일수 = Σ 단위작업(중복가능 제외) + 직접입력."""
    s = sum(it.total for it in items)
    return s + (act.direct_days or 0)
