"""리포트 데이터를 시각화하기 위한 Plotly 차트 빌더.

색상/마크 규칙은 dataviz 스킬 가이드(단일 계열=시퀀셜 1색, 2계열 비교=카테고리컬 고정 순서,
막대 두께 제한, 둥근 데이터 끝, 옅은 그리드라인, 값 라벨은 절제해서)를 따른다.
"""
import plotly.graph_objects as go

BLUE = "#2a78d6"    # 시퀀셜 기본 / 확정
RED = "#e34948"     # 발산 대응쌍 / 취소
ORANGE = "#eb6834"  # 취소율 등 위험 신호용 보조 색

TEXT_PRIMARY = "#0b0b0b"
TEXT_SECONDARY = "#52514e"
TEXT_MUTED = "#898781"
GRIDLINE = "#e1e0d9"
BASELINE = "#c3c2b7"

FONT = dict(family="system-ui, -apple-system, 'Segoe UI', sans-serif", color=TEXT_PRIMARY)

BAR_WIDTH = 0.55  # 트랙 대비 ~24px 두께감을 주는 상대폭


def _base_layout(height=360, showlegend=False):
    return dict(
        height=height,
        margin=dict(l=8, r=24, t=8, b=8),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=FONT,
        showlegend=showlegend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0,
                     font=dict(color=TEXT_SECONDARY, size=12)),
        hoverlabel=dict(bgcolor="white", font=dict(color=TEXT_PRIMARY, size=12)),
    )


def _axis(title=None, tickformat=None, categoryorder=None, categoryarray=None):
    ax = dict(
        title=dict(text=title, font=dict(color=TEXT_MUTED, size=12)) if title else None,
        gridcolor=GRIDLINE, zerolinecolor=BASELINE, linecolor=BASELINE,
        tickfont=dict(color=TEXT_MUTED, size=12),
    )
    if tickformat:
        ax["tickformat"] = tickformat
    if categoryorder:
        ax["categoryorder"] = categoryorder
    if categoryarray:
        ax["categoryarray"] = categoryarray
    return ax


def daily_sales_chart(daily_df):
    x = [f"{d.strftime('%m-%d')}({wd})" for d, wd in zip(daily_df["판매일"], daily_df["요일"])]
    fig = go.Figure(go.Bar(
        x=x, y=daily_df["매출액"], marker_color=BLUE, width=BAR_WIDTH,
        marker_line_width=0,
        text=[f"{v:,.0f}" for v in daily_df["매출액"]], textposition="outside",
        textfont=dict(color=TEXT_SECONDARY, size=11),
        hovertemplate="%{x}<br>매출액 %{y:,.0f}원<extra></extra>",
    ))
    fig.update_traces(marker=dict(cornerradius=4))
    fig.update_layout(**_base_layout())
    fig.update_xaxes(**_axis())
    fig.update_yaxes(**_axis(tickformat=","))
    return fig


def horizontal_bar_chart(labels, values, color=BLUE, value_suffix="원", height=None):
    order = sorted(range(len(labels)), key=lambda i: values[i])
    labels = [labels[i] for i in order]
    values = [values[i] for i in order]
    height = height or max(280, 40 * len(labels))
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h", marker_color=color, width=BAR_WIDTH,
        marker_line_width=0,
        text=[f"{v:,.0f}{value_suffix}" for v in values], textposition="outside",
        textfont=dict(color=TEXT_SECONDARY, size=11),
        hovertemplate="%{y}<br>%{x:,.0f}" + value_suffix + "<extra></extra>",
    ))
    fig.update_traces(marker=dict(cornerradius=4))
    fig.update_layout(**_base_layout(height=height))
    fig.update_xaxes(**_axis(tickformat=","))
    fig.update_yaxes(**_axis())
    return fig


def weekday_chart(weekday_df):
    order = ["월", "화", "수", "목", "금", "토", "일"]
    fig = go.Figure(go.Bar(
        x=weekday_df["요일"], y=weekday_df["매출액"], marker_color=BLUE, width=BAR_WIDTH,
        marker_line_width=0,
        text=[f"{v:,.0f}" for v in weekday_df["매출액"]], textposition="outside",
        textfont=dict(color=TEXT_SECONDARY, size=11),
        hovertemplate="%{x}요일<br>매출액 %{y:,.0f}원<extra></extra>",
    ))
    fig.update_traces(marker=dict(cornerradius=4))
    fig.update_layout(**_base_layout())
    fig.update_xaxes(**_axis(categoryorder="array", categoryarray=order))
    fig.update_yaxes(**_axis(tickformat=","))
    return fig


def monthly_stacked_chart(months, monthly_checkin):
    x = [f"{m.year}.{m.month:02d}" for m in months]
    confirmed = [monthly_checkin["확정"][m][2] for m in months]
    cancelled = [monthly_checkin["취소"][m][2] for m in months]
    fig = go.Figure()
    fig.add_bar(x=x, y=confirmed, name="확정", marker_color=BLUE, width=BAR_WIDTH,
                marker_line=dict(color="white", width=2),
                hovertemplate="%{x}<br>확정 매출 %{y:,.0f}원<extra></extra>")
    fig.add_bar(x=x, y=cancelled, name="취소", marker_color=RED, width=BAR_WIDTH,
                marker_line=dict(color="white", width=2),
                hovertemplate="%{x}<br>취소 매출 %{y:,.0f}원<extra></extra>")
    fig.update_layout(barmode="stack", **_base_layout(showlegend=True))
    fig.update_xaxes(**_axis())
    fig.update_yaxes(**_axis(tickformat=","))
    return fig
