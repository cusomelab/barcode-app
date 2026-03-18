"""쿠팡 일일 재고/판매량 대시보드

구글 시트 "판매현황" 탭에서 판매된 상품만 표시
+ 브랜드별(마켓피아/올하이) 전체 재고 현황
"""
import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="재고/판매량 추적", page_icon="📦", layout="wide")

SHEET_ID = st.secrets.get("GOOGLE_SHEET_ID", "")
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_gc():
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"구글 인증 실패: {e}")
        st.info("Streamlit 앱 Settings > Secrets에 서비스 계정 정보를 설정하세요.")
        st.stop()


@st.cache_data(ttl=300)
def load_sheet(name: str) -> pd.DataFrame:
    try:
        gc = get_gc()
        ws = gc.open_by_key(SHEET_ID).worksheet(name)
        return pd.DataFrame(ws.get_all_records())
    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=300)
def get_ws_names() -> list:
    try:
        return [ws.title for ws in get_gc().open_by_key(SHEET_ID).worksheets()]
    except Exception:
        return []


def parse_stock(val) -> int:
    s = str(val)
    if "품절" in s:
        return 0
    nums = re.findall(r"\d+", s)
    return int(nums[0]) if nums else 0


# ─── 메인 ──────────────────────────────────────────

def main():
    if not SHEET_ID:
        st.error("GOOGLE_SHEET_ID가 설정되지 않았습니다.")
        st.stop()

    st.title("📦 쿠팡 재고/판매량 대시보드")

    tab_sold, tab_brand = st.tabs(["🔥 판매 현황", "📋 브랜드별 재고"])

    # ─── 탭1: 판매현황 (마켓피아+올하이 합산) ───
    with tab_sold:
        show_sales_dashboard()

    # ─── 탭2: 브랜드별 전체 재고 ───
    with tab_brand:
        show_brand_dashboard()


def show_sales_dashboard():
    """판매현황 시트 — 판매된 것만"""
    df = load_sheet("판매현황")

    if df.empty:
        st.warning("'판매현황' 시트에 데이터가 없습니다. daily_run.py 실행 후 확인하세요.")
        return

    # 숫자 변환
    for col in ["판매자가격", "판매량", "별점", "후기수"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    st.subheader("🔥 오늘 판매된 상품")

    # 요약 카드
    col1, col2, col3, col4 = st.columns(4)
    total_items = len(df)
    total_sales = int(df["판매량"].sum()) if "판매량" in df.columns else 0
    brands = df["브랜드"].value_counts() if "브랜드" in df.columns else pd.Series()

    with col1:
        st.metric("판매 상품 수", f"{total_items}건")
    with col2:
        st.metric("총 판매량", f"{total_sales}개")
    with col3:
        mk = int(brands.get("마켓피아", 0))
        st.metric("마켓피아", f"{mk}건")
    with col4:
        oh = int(brands.get("올하이", 0))
        st.metric("올하이", f"{oh}건")

    # 수집일시
    if "수집일시" in df.columns and len(df) > 0:
        st.caption(f"수집: {df['수집일시'].iloc[0]}")

    # 필터
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        brand_filter = st.multiselect("브랜드 필터", ["마켓피아", "올하이"], default=["마켓피아", "올하이"])
    with col_f2:
        search = st.text_input("상품명 검색", "")

    filtered = df.copy()
    if brand_filter and "브랜드" in filtered.columns:
        filtered = filtered[filtered["브랜드"].isin(brand_filter)]
    if search and "상품명" in filtered.columns:
        filtered = filtered[filtered["상품명"].str.contains(search, case=False, na=False)]

    # 테이블
    st.dataframe(filtered, use_container_width=True, height=400)

    # 판매량 TOP 10 차트
    if len(filtered) > 0 and "상품명" in filtered.columns:
        st.subheader("📊 판매량 TOP 10")
        top10 = filtered.nlargest(10, "판매량").copy()
        top10["_name"] = top10["상품명"].str[:30] + ".."
        fig = px.bar(top10, x="판매량", y="_name", orientation="h",
                     color="브랜드" if "브랜드" in top10.columns else None,
                     color_discrete_map={"마켓피아": "#FF6B35", "올하이": "#4A90D9"},
                     labels={"_name": "", "판매량": "판매량(개)"})
        fig.update_layout(yaxis=dict(autorange="reversed"), showlegend=True, height=400)
        st.plotly_chart(fig, use_container_width=True)


def show_brand_dashboard():
    """브랜드별 전체 재고 현황"""
    ws_names = get_ws_names()
    brands = [n for n in ws_names if n in ["마켓피아", "올하이"]]

    if not brands:
        st.warning("마켓피아/올하이 시트가 없습니다.")
        return

    selected = st.selectbox("브랜드 선택", brands)
    df = load_sheet(selected)

    if df.empty:
        st.warning(f"'{selected}' 시트에 데이터가 없습니다.")
        return

    # 재고 파싱
    if "재고량" in df.columns:
        df["_재고수"] = df["재고량"].apply(parse_stock)
    else:
        df["_재고수"] = 0

    total = len(df)
    soldout = len(df[df["_재고수"] == 0])
    in_stock = total - soldout

    # 요약
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("전체", f"{total:,}건")
    with col2:
        st.metric("재고 있음", f"{in_stock:,}건")
    with col3:
        st.metric("품절", f"{soldout:,}건")

    # 차트
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        bins = [0, 1, 5, 10, 50, 100, float("inf")]
        labels = ["품절", "1~4", "5~9", "10~49", "50~99", "100+"]
        df["_구간"] = pd.cut(df["_재고수"], bins=bins, labels=labels, right=False)
        dist = df["_구간"].value_counts().reindex(labels).fillna(0)
        fig = px.bar(x=dist.index, y=dist.values, labels={"x": "재고", "y": "상품수"},
                     color_discrete_sequence=["#FF6B35"])
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with col_c2:
        fig2 = go.Figure(data=[go.Pie(
            labels=["재고있음", "품절"], values=[in_stock, soldout],
            hole=0.4, marker_colors=["#2ecc71", "#e74c3c"])])
        st.plotly_chart(fig2, use_container_width=True)

    # 상품 목록
    search = st.text_input("🔍 검색", "", key="brand_search")
    display = df.copy()
    if search:
        name_col = next((c for c in ["상품명"] if c in display.columns), None)
        if name_col:
            display = display[display[name_col].str.contains(search, case=False, na=False)]

    cols = [c for c in display.columns if not c.startswith("_")]
    st.dataframe(display[cols], use_container_width=True, height=400)


if __name__ == "__main__":
    main()
