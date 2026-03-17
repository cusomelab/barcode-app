"""쿠팡 일일 재고/판매량 대시보드

구글 시트에서 데이터를 읽어 시각화
- 마켓피아 / 올하이 브랜드별 상품 현황
- 재고 분포, 품절 현황, 판매량 랭킹
"""
import re
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="재고/판매량 추적", page_icon="📦", layout="wide")

# ─── 구글 시트 설정 ─────────────────────────────────
SHEET_ID = st.secrets.get("GOOGLE_SHEET_ID", "")
KEYWORDS = ["마켓피아", "올하이"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_gspread_client():
    """Streamlit secrets에서 서비스 계정 인증"""
    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"구글 인증 실패: {e}")
        st.info("`.streamlit/secrets.toml`에 서비스 계정 정보를 설정하세요.")
        st.stop()


@st.cache_data(ttl=300)
def load_sheet(worksheet_name: str) -> pd.DataFrame:
    """구글 시트 워크시트 로드"""
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet(worksheet_name)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except gspread.exceptions.WorksheetNotFound:
        return pd.DataFrame()
    except Exception as e:
        st.warning(f"'{worksheet_name}' 로드 실패: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=300)
def get_worksheet_names() -> list:
    """시트의 모든 워크시트명 가져오기"""
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(SHEET_ID)
        return [ws.title for ws in sh.worksheets()]
    except Exception:
        return []


def parse_stock(val) -> int:
    """'재고량 : 5개' 또는 '품절' → 숫자"""
    if pd.isna(val):
        return 0
    s = str(val)
    if "품절" in s:
        return 0
    nums = re.findall(r"\d+", s)
    return int(nums[0]) if nums else 0


# ─── 메인 ──────────────────────────────────────────

def main():
    st.title("📦 쿠팡 재고/판매량 대시보드")

    if not SHEET_ID:
        st.error("GOOGLE_SHEET_ID가 설정되지 않았습니다.")
        st.stop()

    # 사이드바
    st.sidebar.header("설정")

    # 워크시트 목록
    ws_names = get_worksheet_names()
    if not ws_names:
        st.warning("시트를 불러올 수 없습니다.")
        st.stop()

    # 브랜드 선택
    keyword = st.sidebar.selectbox("브랜드", KEYWORDS)

    # 해당 브랜드의 워크시트 찾기
    brand_sheets = [n for n in ws_names if keyword in n]
    if not brand_sheets:
        st.warning(f"'{keyword}' 데이터가 시트에 없습니다.")
        st.stop()

    selected_sheet = st.sidebar.selectbox("시트 선택", brand_sheets)

    # 데이터 로드
    df = load_sheet(selected_sheet)
    if df.empty:
        st.warning(f"'{selected_sheet}' 시트에 데이터가 없습니다.")
        st.stop()

    st.sidebar.success(f"{len(df)}건 로드됨")

    # ─── 재고 수치 파싱 ───
    if "재고량" in df.columns:
        df["_재고수"] = df["재고량"].apply(parse_stock)
    elif "오늘재고" in df.columns:
        df["_재고수"] = pd.to_numeric(df["오늘재고"], errors="coerce").fillna(0).astype(int)
    else:
        df["_재고수"] = 0

    # ─── 요약 카드 ───
    total = len(df)
    soldout = len(df[df["_재고수"] == 0])
    in_stock = total - soldout

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("전체 상품", f"{total:,}건")
    with col2:
        st.metric("재고 있음", f"{in_stock:,}건", delta=f"{in_stock/max(total,1)*100:.0f}%")
    with col3:
        st.metric("품절", f"{soldout:,}건", delta=f"-{soldout/max(total,1)*100:.0f}%")
    with col4:
        if "판매량" in df.columns:
            total_sales = pd.to_numeric(df["판매량"], errors="coerce").sum()
            st.metric("일일 판매량", f"{int(total_sales):,}개")
        else:
            st.metric("데이터 기준", selected_sheet)

    # ─── 탭 ───
    tab1, tab2, tab3, tab4 = st.tabs(["📋 상품 목록", "📈 판매 TOP", "🔴 품절", "📊 차트"])

    with tab1:
        show_list(df)
    with tab2:
        show_sales_top(df)
    with tab3:
        show_soldout(df)
    with tab4:
        show_charts(df)


def show_list(df: pd.DataFrame):
    """상품 목록"""
    search = st.text_input("🔍 상품명 검색", "")
    name_col = next((c for c in ["상품명", "itemnm"] if c in df.columns), None)

    filtered = df.copy()
    if search and name_col:
        filtered = filtered[filtered[name_col].str.contains(search, case=False, na=False)]

    display_cols = [c for c in filtered.columns if not c.startswith("_")]
    st.dataframe(filtered[display_cols], use_container_width=True, height=500)
    st.caption(f"{len(filtered)}건")


def show_sales_top(df: pd.DataFrame):
    """판매량 TOP"""
    if "판매량" not in df.columns:
        st.info("판매량 데이터가 없습니다. daily_run.py 실행 후 확인하세요.")
        return

    df["_판매"] = pd.to_numeric(df["판매량"], errors="coerce").fillna(0)
    top = df.nlargest(50, "_판매")

    name_col = next((c for c in ["상품명", "itemnm"] if c in top.columns), None)
    cols = [c for c in [name_col, "vendorItemId", "판매자가격", "전일재고", "오늘재고", "판매량", "품절여부", "아이템위너"] if c and c in top.columns]

    st.dataframe(top[cols], use_container_width=True, height=500)

    # 차트
    if name_col and len(top) > 0:
        top10 = top.head(10).copy()
        top10["_name"] = top10[name_col].str[:25] + ".."
        fig = px.bar(top10, x="_판매", y="_name", orientation="h",
                     color="_판매", color_continuous_scale="Oranges",
                     labels={"_판매": "판매량", "_name": ""})
        fig.update_layout(yaxis=dict(autorange="reversed"), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)


def show_soldout(df: pd.DataFrame):
    """품절 현황"""
    soldout = df[df["_재고수"] == 0]
    if soldout.empty:
        st.success("품절 상품이 없습니다!")
        return

    st.warning(f"품절: {len(soldout)}건 / 전체: {len(df)}건")
    display_cols = [c for c in soldout.columns if not c.startswith("_")]
    st.dataframe(soldout[display_cols], use_container_width=True, height=500)


def show_charts(df: pd.DataFrame):
    """차트"""
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**재고 분포**")
        bins = [0, 1, 5, 10, 50, 100, float("inf")]
        labels = ["품절", "1~4개", "5~9개", "10~49개", "50~99개", "100+"]
        df["_구간"] = pd.cut(df["_재고수"], bins=bins, labels=labels, right=False)
        dist = df["_구간"].value_counts().reindex(labels).fillna(0)
        fig = px.bar(x=dist.index, y=dist.values,
                     labels={"x": "재고 구간", "y": "상품 수"},
                     color=dist.index,
                     color_discrete_sequence=px.colors.sequential.RdBu_r)
        fig.update_layout(showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.markdown("**품절 비율**")
        soldout = len(df[df["_재고수"] == 0])
        in_stock = len(df) - soldout
        fig2 = go.Figure(data=[go.Pie(
            labels=["재고있음", "품절"], values=[in_stock, soldout],
            hole=0.4, marker_colors=["#2ecc71", "#e74c3c"],
        )])
        st.plotly_chart(fig2, use_container_width=True)


if __name__ == "__main__":
    main()
