"""공동구매 매출 실적 리포트 생성 엔진.

원본 예약 엑셀(개별 예약 raw 데이터)을 읽어 요약/분석용 데이터프레임들을 계산한다.
Streamlit UI(app.py)에서 이 모듈의 함수들을 불러써서 화면 표시 / XLSX / HTML을 만든다.
"""
import re
import io
import datetime
import pandas as pd
import numpy as np

REQUIRED_COLUMNS = [
    "숙소", "예약 상태", "예약번호", "예약일", "체크인", "체크아웃",
    "박수", "객실수", "객실명", "요금제명(코드)", "통화타입",
    "판매가", "입금가", "판매채널", "예약경로", "취소일자", "취소 수수료",
]

CONFIRMED_STATUSES = {"확정", "완료"}
CANCELLED_STATUSES = {"취소"}

WEEKDAY_KO = ["월", "화", "수", "목", "금", "토", "일"]

# 이모지 및 기타 심볼 제거용 (유니코드 emoji 블록 대부분 커버)
_EMOJI_PATTERN = re.compile(
    "["
    "\U0001F300-\U0001FAFF"
    "\U00002600-\U000027BF"
    "\U0001F1E6-\U0001F1FF"
    "\U00002190-\U000021FF"
    "\U0000FE0F"
    "\U0001F3FB-\U0001F3FF"  # skin tone modifiers
    "]+",
    flags=re.UNICODE,
)


def auto_clean_name(raw: str) -> str:
    """이모지 / 'PKG' 표기 / '+' 구분자를 정리해 사람이 읽기 좋은 이름으로 변환.

    완벽한 정규화는 불가능하므로(딜마다 표기 방식이 달라) 이 결과는 '기본 추정값'이고,
    UI의 이름 매핑 표에서 사용자가 최종적으로 고칠 수 있게 한다.
    """
    if raw is None:
        return ""
    s = str(raw)
    s = _EMOJI_PATTERN.sub("", s)
    s = re.sub(r"\bPKG\b", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^\s*마켓딜\s*", "", s)  # 딜 페이지 마킹 접두사 제거
    s = s.replace("+", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def build_name_mapping(values) -> pd.DataFrame:
    """고유값 목록으로부터 원본→정리이름 매핑 기본 테이블 생성 (수정 가능한 데이터에디터용)."""
    uniques = sorted({str(v) for v in values if v is not None and str(v).strip() != ""})
    return pd.DataFrame(
        {"원본": uniques, "정리된 이름": [auto_clean_name(u) for u in uniques]}
    )


def load_reservations(file) -> pd.DataFrame:
    """예약 엑셀 파일(업로드된 파일 객체 또는 경로)을 읽어 DataFrame으로 반환."""
    df = pd.read_excel(file, sheet_name=0, engine="openpyxl")
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {', '.join(missing)}")

    df["예약일"] = pd.to_datetime(df["예약일"], errors="coerce")
    df["체크인"] = pd.to_datetime(df["체크인"], errors="coerce")
    df["체크아웃"] = pd.to_datetime(df["체크아웃"], errors="coerce")
    for col in ["박수", "객실수", "판매가", "입금가", "취소 수수료"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["예약경로"] = df["예약경로"].fillna("").astype(str)
    df["판매채널"] = df["판매채널"].fillna("").astype(str)
    df["예약 상태"] = df["예약 상태"].fillna("").astype(str)
    df["객실명"] = df["객실명"].fillna("").astype(str)
    df["요금제명(코드)"] = df["요금제명(코드)"].fillna("").astype(str)
    return df


def filter_by_channel(df: pd.DataFrame, keyword: str, column: str = "예약경로") -> pd.DataFrame:
    """지정 컬럼 값에 keyword(대소문자 무시)가 포함된 행만 남긴다. keyword가 비어있으면 전체 반환."""
    if not keyword:
        return df.copy()
    mask = df[column].str.lower().str.contains(keyword.lower(), na=False)
    return df[mask].copy()


def apply_name_cleaning(df: pd.DataFrame, room_map: pd.DataFrame, package_map: pd.DataFrame) -> pd.DataFrame:
    """사용자가 수정한 이름 매핑 표를 적용해 '객실명_정리', '패키지명_정리' 컬럼 추가."""
    df = df.copy()
    room_dict = dict(zip(room_map["원본"], room_map["정리된 이름"]))
    pkg_dict = dict(zip(package_map["원본"], package_map["정리된 이름"]))
    df["객실명_정리"] = df["객실명"].map(lambda v: room_dict.get(str(v), auto_clean_name(v)))
    df["패키지명_정리"] = df["요금제명(코드)"].map(lambda v: pkg_dict.get(str(v), auto_clean_name(v)))
    df["is_confirmed"] = df["예약 상태"].isin(CONFIRMED_STATUSES)
    df["is_cancelled"] = df["예약 상태"].isin(CANCELLED_STATUSES)
    return df


def fmt_won(n) -> str:
    try:
        return f"{int(round(n)):,}원"
    except (ValueError, TypeError):
        return "0원"


def fmt_int(n) -> str:
    try:
        return f"{int(round(n)):,}"
    except (ValueError, TypeError):
        return "0"


def fmt_pct(n) -> str:
    try:
        return f"{n * 100:.1f}%"
    except (ValueError, TypeError):
        return "0.0%"


class ReportData:
    """리포트 생성에 필요한 모든 집계 결과를 담는 컨테이너."""

    def __init__(self, df: pd.DataFrame, meta: dict):
        self.df = df
        self.meta = meta  # 인플루언서명, 판매기간, 호텔명, 투숙기간
        self._compute()

    def _compute(self):
        df = self.df
        conf = df[df["is_confirmed"]]
        canc = df[df["is_cancelled"]]

        self.total_count = len(df)
        self.total_nights = int(df["박수"].sum())
        self.total_revenue = float(df["판매가"].sum())

        self.confirmed_count = len(conf)
        self.confirmed_nights = int(conf["박수"].sum())
        self.confirmed_revenue = float(conf["판매가"].sum())

        self.cancelled_count = len(canc)
        self.cancelled_nights = int(canc["박수"].sum())
        self.cancelled_revenue = float(canc["판매가"].sum())

        self.cancel_rate = (
            self.cancelled_count / self.total_count if self.total_count else 0.0
        )

        # ---- 전체 예약 현황 (숙소별로 전체/확정/취소 상태를 행으로 구분) ----
        def _status_rows(label, g):
            gc = g[g["is_confirmed"]]
            gx = g[g["is_cancelled"]]
            return [
                {"숙소": label, "상태": "전체", "건수": len(g), "박수": int(g["박수"].sum()), "매출액": float(g["판매가"].sum())},
                {"숙소": label, "상태": "확정", "건수": len(gc), "박수": int(gc["박수"].sum()), "매출액": float(gc["판매가"].sum())},
                {"숙소": label, "상태": "취소", "건수": len(gx), "박수": int(gx["박수"].sum()), "매출액": float(gx["판매가"].sum())},
            ]

        status_rows = []
        for hotel, g in df.groupby("숙소"):
            status_rows += _status_rows(hotel, g)
        status_rows += _status_rows("전체", df)
        self.overall_status = pd.DataFrame(status_rows)

        # ---- 패키지별 분석 (확정 기준) ----
        pkg_rows = []
        for pkg, g in conf.groupby("패키지명_정리"):
            cnt, nights, rev = len(g), int(g["박수"].sum()), float(g["판매가"].sum())
            cancel_cnt = int((canc["패키지명_정리"] == pkg).sum())
            denom = cnt + cancel_cnt
            pkg_rows.append({
                "패키지": pkg, "건수": cnt, "박수": nights, "매출액": rev,
                "ADR_박": rev / nights if nights else 0,
                "ADR_건": rev / cnt if cnt else 0,
                "취소건": cancel_cnt,
                "취소율": cancel_cnt / denom if denom else 0,
                "매출비중": rev / self.confirmed_revenue if self.confirmed_revenue else 0,
            })
        pkg_df = pd.DataFrame(pkg_rows).sort_values("매출액", ascending=False).reset_index(drop=True)
        total_row = {
            "패키지": "합계", "건수": self.confirmed_count, "박수": self.confirmed_nights,
            "매출액": self.confirmed_revenue,
            "ADR_박": self.confirmed_revenue / self.confirmed_nights if self.confirmed_nights else 0,
            "ADR_건": self.confirmed_revenue / self.confirmed_count if self.confirmed_count else 0,
            "취소건": self.cancelled_count,
            "취소율": self.cancel_rate,
            "매출비중": 1.0,
        }
        self.package_analysis = pd.concat([pkg_df, pd.DataFrame([total_row])], ignore_index=True)

        # ---- 체크인 월별 현황 ----
        df2 = df.copy()
        df2["월"] = df2["체크인"].dt.to_period("M")
        months = sorted(df2["월"].dropna().unique())
        month_data = {"전체": {}, "확정": {}, "취소": {}}
        for m in months:
            gm = df2[df2["월"] == m]
            gmc = gm[gm["is_confirmed"]]
            gmx = gm[gm["is_cancelled"]]
            month_data["전체"][m] = (len(gm), int(gm["박수"].sum()), float(gm["판매가"].sum()))
            month_data["확정"][m] = (len(gmc), int(gmc["박수"].sum()), float(gmc["판매가"].sum()))
            month_data["취소"][m] = (len(gmx), int(gmx["박수"].sum()), float(gmx["판매가"].sum()))
        self.months = months
        self.monthly_checkin = month_data

        # ---- 일별 판매 현황 (예약일 기준, 전체 상태 포함) ----
        df3 = df.copy()
        df3["판매일"] = df3["예약일"].dt.date
        daily = (
            df3.groupby("판매일")
            .agg(건수=("예약번호", "count"), 박수=("박수", "sum"), 매출액=("판매가", "sum"))
            .reset_index()
            .sort_values("판매일")
        )
        daily["누적매출액"] = daily["매출액"].cumsum()
        daily["요일"] = daily["판매일"].map(lambda d: WEEKDAY_KO[d.weekday()])
        self.daily_sales = daily

        # ---- 체크인 요일별 현황 (확정 기준) ----
        conf2 = conf.copy()
        conf2["요일"] = conf2["체크인"].dt.weekday.map(lambda i: WEEKDAY_KO[i])
        wd_rows = []
        for wd in WEEKDAY_KO:
            g = conf2[conf2["요일"] == wd]
            cnt, nights, rev = len(g), int(g["박수"].sum()), float(g["판매가"].sum())
            wd_rows.append({
                "요일": wd, "건수": cnt, "박수": nights, "매출액": rev,
                "ADR": rev / nights if nights else 0,
                "비중": rev / self.confirmed_revenue if self.confirmed_revenue else 0,
            })
        self.weekday_checkin = pd.DataFrame(wd_rows)

        # ---- 시설별(객실별) 분석 ----
        room_rows = []
        for room, g in df.groupby("객실명_정리"):
            gc = g[g["is_confirmed"]]
            gx = g[g["is_cancelled"]]
            room_rows.append({
                "객실명": room,
                "전체_건수": len(g), "전체_박수": int(g["박수"].sum()), "전체_매출": float(g["판매가"].sum()),
                "확정_건수": len(gc), "확정_박수": int(gc["박수"].sum()), "확정_매출": float(gc["판매가"].sum()),
                "취소_건수": len(gx), "취소_박수": int(gx["박수"].sum()), "취소_매출": float(gx["판매가"].sum()),
            })
        room_df = pd.DataFrame(room_rows).sort_values("전체_매출", ascending=False).reset_index(drop=True)
        room_total = {
            "객실명": "전체 소계",
            "전체_건수": self.total_count, "전체_박수": self.total_nights, "전체_매출": self.total_revenue,
            "확정_건수": self.confirmed_count, "확정_박수": self.confirmed_nights, "확정_매출": self.confirmed_revenue,
            "취소_건수": self.cancelled_count, "취소_박수": self.cancelled_nights, "취소_매출": self.cancelled_revenue,
        }
        self.room_analysis = pd.concat([room_df, pd.DataFrame([room_total])], ignore_index=True)

        # ---- 패키지×객실 매트릭스 (확정 기준) ----
        rooms_sorted = room_df["객실명"].tolist()
        packages_sorted = pkg_df["패키지"].tolist()
        matrix = {}
        for room in rooms_sorted:
            row = {}
            for pkg in packages_sorted:
                g = conf[(conf["객실명_정리"] == room) & (conf["패키지명_정리"] == pkg)]
                row[pkg] = (len(g), int(g["박수"].sum()), float(g["판매가"].sum()))
            gr = conf[conf["객실명_정리"] == room]
            row["소계"] = (len(gr), int(gr["박수"].sum()), float(gr["판매가"].sum()))
            matrix[room] = row
        total_row_m = {}
        for pkg in packages_sorted:
            gp = conf[conf["패키지명_정리"] == pkg]
            total_row_m[pkg] = (len(gp), int(gp["박수"].sum()), float(gp["판매가"].sum()))
        total_row_m["소계"] = (self.confirmed_count, self.confirmed_nights, self.confirmed_revenue)
        matrix["합계"] = total_row_m
        self.matrix_rooms = rooms_sorted + ["합계"]
        self.matrix_packages = packages_sorted
        self.matrix = matrix

        # ---- ADR 분석 ----
        self.overall_adr = self.confirmed_revenue / self.confirmed_nights if self.confirmed_nights else 0

        adr_pkg_rows = []
        for _, r in pkg_df.iterrows():
            adr_pkg_rows.append(r.to_dict())
        self.adr_package = pd.DataFrame(adr_pkg_rows)

        adr_room_rows = []
        for room, g in conf.groupby("객실명_정리"):
            cnt, nights, rev = len(g), int(g["박수"].sum()), float(g["판매가"].sum())
            adr_room_rows.append({
                "객실명": room, "건수": cnt, "박수": nights, "확정매출": rev,
                "ADR": rev / nights if nights else 0,
                "건당평균단가": rev / cnt if cnt else 0,
                "비중": rev / self.confirmed_revenue if self.confirmed_revenue else 0,
            })
        self.adr_room = (
            pd.DataFrame(adr_room_rows).sort_values("ADR", ascending=False).reset_index(drop=True)
        )

        # ---- 체크인 일자별 상세 (확정 기준) ----
        conf3 = conf.copy()
        conf3["체크인일"] = conf3["체크인"].dt.date
        detail_rows = []
        for day, g in conf3.groupby("체크인일"):
            wd = WEEKDAY_KO[day.weekday()]
            sub = (
                g.groupby(["객실명_정리", "패키지명_정리"])
                .agg(건수=("예약번호", "count"), 박수=("박수", "sum"), 매출액=("판매가", "sum"))
                .reset_index()
                .sort_values(["객실명_정리", "패키지명_정리"])
            )
            for i, sr in enumerate(sub.itertuples(index=False)):
                detail_rows.append({
                    "체크인": day.strftime("%m-%d") if i == 0 else None,
                    "요일": wd if i == 0 else None,
                    "객실명": sr[0],
                    "패키지": sr[1],
                    "건수": sr[2], "박수": sr[3], "매출액": sr[4],
                })
        self.checkin_detail = pd.DataFrame(detail_rows)
