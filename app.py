import datetime
import streamlit as st
import pandas as pd

from report_engine import (
    load_reservations, filter_by_channel, build_name_mapping, apply_name_cleaning,
    ReportData, fmt_int, fmt_won, fmt_pct,
)
from exporters import build_workbook, build_html, build_email_body, render_overall_status_table
import charts

st.set_page_config(page_title="공동구매 매출 실적 리포트", page_icon="📊", layout="wide")
st.title("📊 공동구매 매출 실적 리포트 생성기")

if "report" not in st.session_state:
    st.session_state.report = None

uploaded = st.file_uploader("예약 엑셀 파일 (.xlsx)", type=["xlsx"])

if uploaded is None:
    st.info("예약 엑셀 파일을 업로드하면 시작할 수 있어요.")
    st.stop()

try:
    raw_df = load_reservations(uploaded)
except ValueError as e:
    st.error(str(e))
    st.stop()

st.success(f"✅ {len(raw_df)}건 로드 완료")

# ---------------- 채널 필터 ----------------
st.subheader("판매채널 필터")
c1, c2 = st.columns([1, 2])
with c1:
    use_filter = st.checkbox("예약경로 필터 사용", value=True)
with c2:
    keyword = st.text_input("포함 키워드 (예약경로 컬럼 기준, 대소문자 무시)", value="tceleb", disabled=not use_filter)

df = filter_by_channel(raw_df, keyword if use_filter else "")
st.caption(f"필터 적용 후 {len(df)}건 (전체 {len(raw_df)}건 중)")

# ---------------- 이름 정리 매핑 (수정 가능) ----------------
with st.expander("🏷️ 객실명 · 패키지명 정리 (자동 추정 — 직접 수정 가능)", expanded=False):
    st.caption("이모지/PKG 표기 등을 자동으로 정리한 값입니다. '정리된 이름' 칸을 더블클릭해 직접 고칠 수 있어요.")
    room_map_default = build_name_mapping(df["객실명"])
    pkg_map_default = build_name_mapping(df["요금제명(코드)"])

    st.markdown("**객실명**")
    room_map = st.data_editor(
        room_map_default, key="room_map", num_rows="fixed", use_container_width=True,
        disabled=["원본"],
    )
    st.markdown("**패키지명**")
    pkg_map = st.data_editor(
        pkg_map_default, key="pkg_map", num_rows="fixed", use_container_width=True,
        disabled=["원본"],
    )

df = apply_name_cleaning(df, room_map, pkg_map)

# ---------------- 리포트 정보 ----------------
st.subheader("리포트 정보")

default_hotel = df["숙소"].mode().iat[0] if len(df) else ""
sale_dates = df["예약일"].dropna()
stay_dates = df["체크인"].dropna()
default_sale_range = (
    (sale_dates.min().date(), sale_dates.max().date()) if len(sale_dates) else (datetime.date.today(), datetime.date.today())
)
default_stay_range = (
    (stay_dates.min().date(), stay_dates.max().date()) if len(stay_dates) else (datetime.date.today(), datetime.date.today())
)

col1, col2 = st.columns(2)
with col1:
    influencer = st.text_input("인플루언서명", value="")
    sale_range = st.date_input("판매 기간", value=default_sale_range)
with col2:
    hotel = st.text_input("호텔·리조트명", value=default_hotel)
    stay_range = st.date_input("투숙 기간", value=default_stay_range)

generate = st.button("🚀 리포트 생성", type="primary", use_container_width=True)

if generate:
    if len(df) == 0:
        st.error("필터 적용 후 남은 예약 건이 없습니다. 필터 키워드를 확인해주세요.")
        st.stop()

    def fmt_range(r):
        if isinstance(r, (tuple, list)) and len(r) == 2:
            return f"{r[0].strftime('%Y.%m.%d')} ~ {r[1].strftime('%Y.%m.%d')}"
        return r.strftime('%Y.%m.%d')

    def fmt_range_kr(r):
        if isinstance(r, (tuple, list)) and len(r) == 2:
            return f"{r[0].year}년 {r[0].month:02d}월 {r[0].day:02d}일부터 {r[1].year}년 {r[1].month:02d}월 {r[1].day:02d}일까지"
        return f"{r.year}년 {r.month:02d}월 {r.day:02d}일"

    meta = {
        "influencer": influencer or "(인플루언서명 미입력)",
        "hotel": hotel or "(호텔명 미입력)",
        "sale_period": fmt_range(sale_range),
        "stay_period": fmt_range(stay_range),
        "sale_period_kr": fmt_range_kr(sale_range),
    }
    rd = ReportData(df, meta)
    st.session_state.report = {
        "rd": rd, "meta": meta,
        "xlsx": build_workbook(rd, meta),
        "html": build_html(rd, meta),
        "email": build_email_body(rd, meta),
    }

report = st.session_state.report
if report:
    st.success("✅ 생성 완료!")
    rd = report["rd"]
    meta = report["meta"]

    pkg_no_total = rd.package_analysis[rd.package_analysis["패키지"] != "합계"]
    room_no_total = rd.room_analysis[rd.room_analysis["객실명"] != "전체 소계"]

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("확정 건수(확정+완료)", fmt_int(rd.confirmed_count))
    k2.metric("확정 박수", fmt_int(rd.confirmed_nights))
    k3.metric("확정 매출액", fmt_won(rd.confirmed_revenue))
    k4.metric("취소 건수", fmt_int(rd.cancelled_count))
    k5.metric("취소율", fmt_pct(rd.cancel_rate))

    dcol1, dcol2 = st.columns(2)
    with dcol1:
        st.download_button(
            "📄 HTML 다운로드", data=report["html"].encode("utf-8"),
            file_name=f"report_{meta['influencer']}_X_{meta['hotel']}.html",
            mime="text/html", use_container_width=True,
        )
    with dcol2:
        st.download_button(
            "📊 XLSX 다운로드", data=report["xlsx"],
            file_name=f"report_{meta['influencer']}_X_{meta['hotel']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.subheader("📈 시각화 대시보드")

    st.markdown("**일별 판매 추이** (판매기간 기준, 전체 상태 포함)")
    st.plotly_chart(charts.daily_sales_chart(rd.daily_sales), use_container_width=True)

    vcol1, vcol2 = st.columns(2)
    with vcol1:
        st.markdown("**패키지별 매출** (확정 기준)")
        st.plotly_chart(
            charts.horizontal_bar_chart(pkg_no_total["패키지"].tolist(), pkg_no_total["매출액"].tolist()),
            use_container_width=True,
        )
    with vcol2:
        st.markdown("**객실별 매출** (전체 기준)")
        st.plotly_chart(
            charts.horizontal_bar_chart(room_no_total["객실명"].tolist(), room_no_total["전체_매출"].tolist()),
            use_container_width=True,
        )

    vcol3, vcol4 = st.columns(2)
    with vcol3:
        st.markdown("**체크인 요일별 매출** (확정 기준)")
        st.plotly_chart(charts.weekday_chart(rd.weekday_checkin), use_container_width=True)
    with vcol4:
        st.markdown("**패키지별 취소율**")
        st.plotly_chart(
            charts.horizontal_bar_chart(
                pkg_no_total["패키지"].tolist(), [v * 100 for v in pkg_no_total["취소율"].tolist()],
                color=charts.ORANGE, value_suffix="%",
            ),
            use_container_width=True,
        )

    st.markdown("**체크인 월별 매출** (확정 vs 취소)")
    st.plotly_chart(charts.monthly_stacked_chart(rd.months, rd.monthly_checkin), use_container_width=True)

    with st.expander("▶ 전체 예약 현황", expanded=True):
        st.markdown(render_overall_status_table(rd.overall_status), unsafe_allow_html=True)
    with st.expander("▶ 패키지별 분석 (확정 기준)"):
        st.dataframe(rd.package_analysis, use_container_width=True, hide_index=True)
    with st.expander("▶ 시설별 객실 분석"):
        st.dataframe(rd.room_analysis, use_container_width=True, hide_index=True)
    with st.expander("▶ 일별 판매 현황"):
        st.dataframe(rd.daily_sales, use_container_width=True, hide_index=True)
    with st.expander("▶ 체크인 요일별 현황 (확정 기준)"):
        st.dataframe(rd.weekday_checkin, use_container_width=True, hide_index=True)
    with st.expander("▶ 객실별 ADR (확정 기준)"):
        st.dataframe(rd.adr_room, use_container_width=True, hide_index=True)
    with st.expander("▶ 체크인 일자별 상세 (확정 기준)"):
        st.dataframe(rd.checkin_detail, use_container_width=True, hide_index=True)

    st.subheader("📧 메일 본문")
    email_text = st.text_area("자유롭게 수정할 수 있어요", value=report["email"], height=420)
    st.download_button(
        "메일 본문 텍스트 다운로드", data=email_text.encode("utf-8"),
        file_name="mail_body.txt", mime="text/plain",
    )
