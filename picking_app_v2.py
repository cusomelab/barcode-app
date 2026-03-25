"""
==========================================================
쿠팡 로켓배송 피킹 검증 시스템 v2
==========================================================
핵심 기능:
  1. 구글 시트 실시간 연동 (gspread + 서비스 계정)
  2. 송장번호 입력 → 해당 쉽먼트 피킹 리스트 로드
  3. 바코드 스캔 시 3중 처리:
     - 출고지시서: 쉽먼트별 피킹 수량 차감 + 오피킹 경고
     - 배대지 입고: 회차 기호(●/★/■) 매칭하여 재고 차감
     - 박스 단위: 박스별 수량 추적 → 부족분 감지
  4. 피킹 로그 자동 기록

구글 시트 설정:
  1. 서비스 계정 JSON 키 파일을 이 앱과 같은 폴더에 둠
  2. 구글 시트에 서비스 계정 이메일을 편집자로 공유
  3. .env 파일 또는 아래 CONFIG에서 시트 ID 설정

사용법:
  pip install streamlit pandas gspread google-auth
  streamlit run picking_app_v2.py
==========================================================
"""

import streamlit as st
import pandas as pd
import re
import time
from datetime import datetime
from pathlib import Path

# ============================================================
# ❶ 구글 시트 연동 설정
# ============================================================
# ── 여기만 수정하면 됩니다 ──
CONFIG = {
    # 서비스 계정 JSON 키 파일 경로
    "SERVICE_ACCOUNT_FILE": "service_account.json",

    # 구글 스프레드시트 ID (URL에서 /d/ 뒤의 문자열)
    # 예: https://docs.google.com/spreadsheets/d/여기가_시트ID/edit
    "SPREADSHEET_ID": "여기에_스프레드시트_ID_입력",

    # 시트(탭) 이름
    "SHEET_출고지시서": "출고지시서",          # 출고지시서 시트 탭 이름
    "SHEET_배대지입고": "배대지입고리스트",     # 배대지 입고 리스트 시트 탭 이름
    "SHEET_피킹로그": "피킹로그",              # 피킹 로그 기록용 시트 (자동 생성)
}

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(
    page_title="피킹 검증 시스템 v2",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# 커스텀 CSS
# ============================================================
st.markdown("""
<style>
    /* 스캔 결과 피드백 */
    .scan-ok {
        background: #d4edda; border-left: 6px solid #28a745;
        padding: 1.2rem 1.5rem; border-radius: 8px; margin: 0.5rem 0; color: #155724;
    }
    .scan-error {
        background: #f8d7da; border-left: 6px solid #dc3545;
        padding: 1.2rem 1.5rem; border-radius: 8px; margin: 0.5rem 0; color: #721c24;
        animation: shake 0.5s ease-in-out;
    }
    .scan-warning {
        background: #fff3cd; border-left: 6px solid #ffc107;
        padding: 1.2rem 1.5rem; border-radius: 8px; margin: 0.5rem 0; color: #856404;
    }
    .scan-complete {
        background: #cce5ff; border-left: 6px solid #007bff;
        padding: 1.2rem 1.5rem; border-radius: 8px; margin: 0.5rem 0; color: #004085;
    }
    .scan-shortage {
        background: #e2e3f1; border-left: 6px solid #6c63ff;
        padding: 1.2rem 1.5rem; border-radius: 8px; margin: 0.5rem 0; color: #383467;
    }
    @keyframes shake {
        0%, 100% { transform: translateX(0); }
        20% { transform: translateX(-10px); }
        40% { transform: translateX(10px); }
        60% { transform: translateX(-6px); }
        80% { transform: translateX(6px); }
    }
    /* 스캔 입력창 크게 */
    div[data-testid="stTextInput"] input {
        font-size: 1.3rem !important;
        padding: 0.8rem !important;
        font-weight: 600 !important;
    }
    /* 송장 입력 다이얼로그 강조 */
    .shipment-input {
        background: #f0f2f6;
        padding: 2rem;
        border-radius: 12px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================
# ❷ 구글 시트 연결 함수들
# ============================================================
@st.cache_resource(ttl=600)
def get_gsheet_client():
    """
    구글 시트 클라이언트를 초기화합니다.
    서비스 계정 JSON이 없으면 None 반환 (CSV 모드로 폴백).
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        # 서비스 계정 인증
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_file(
            CONFIG["SERVICE_ACCOUNT_FILE"], scopes=scopes
        )
        client = gspread.authorize(creds)
        return client
    except FileNotFoundError:
        return None
    except Exception as e:
        st.warning(f"구글 시트 연결 실패: {e}")
        return None


def load_sheet_as_df(client, sheet_name):
    """
    구글 시트의 특정 탭을 DataFrame으로 로드합니다.

    Parameters:
        client: gspread 클라이언트
        sheet_name: 시트(탭) 이름

    Returns:
        pd.DataFrame 또는 None
    """
    try:
        spreadsheet = client.open_by_key(CONFIG["SPREADSHEET_ID"])
        worksheet = spreadsheet.worksheet(sheet_name)

        # 전체 데이터를 가져와서 DataFrame으로 변환
        data = worksheet.get_all_records()
        if not data:
            return pd.DataFrame()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"시트 '{sheet_name}' 로드 실패: {e}")
        return None


def update_sheet_cell(client, sheet_name, row_idx, col_idx, value):
    """
    구글 시트의 특정 셀을 업데이트합니다.
    row_idx, col_idx는 1-based (구글 시트 기준).
    """
    try:
        spreadsheet = client.open_by_key(CONFIG["SPREADSHEET_ID"])
        worksheet = spreadsheet.worksheet(sheet_name)
        worksheet.update_cell(row_idx, col_idx, value)
        return True
    except Exception as e:
        st.error(f"시트 업데이트 실패: {e}")
        return False


def append_log_to_sheet(client, log_entry):
    """
    피킹 로그를 구글 시트에 한 행 추가합니다.
    시트가 없으면 자동 생성합니다.
    """
    try:
        spreadsheet = client.open_by_key(CONFIG["SPREADSHEET_ID"])

        # 피킹로그 시트가 없으면 생성
        try:
            ws = spreadsheet.worksheet(CONFIG["SHEET_피킹로그"])
        except Exception:
            ws = spreadsheet.add_worksheet(
                title=CONFIG["SHEET_피킹로그"], rows=1000, cols=10
            )
            # 헤더 추가
            ws.append_row([
                "시간", "송장번호", "바코드", "상품명",
                "결과", "스캔수량", "필요수량", "회차기호", "박스번호"
            ])

        ws.append_row(log_entry)
        return True
    except Exception as e:
        # 로그 기록 실패는 경고만 (피킹 작업 중단하지 않음)
        return False


# ============================================================
# ❸ 박스번호 파싱 — 회차 기호 + 박스번호 + 수량 분리
# ============================================================
def parse_box_info(box_str):
    """
    출고지시서의 박스번호 문자열을 파싱합니다.

    입력 예시와 결과:
        "●6(3)"        → 기호=●, 박스=6, 수량=3, 상태=피킹가능
        "★17(11)"      → 기호=★, 박스=17, 수량=11, 상태=피킹가능
        "★1"           → 기호=★, 박스=1, 수량=None, 상태=피킹가능 (배대지 시트)
        "국내재고"       → 기호=국내재고, 박스=None, 수량=None, 상태=피킹가능
        "국내재고(2)"    → 기호=국내재고, 박스=None, 수량=2, 상태=피킹가능
        "부족(-1)"      → 기호=부족, 박스=None, 수량=-1, 상태=부족
        "국내부족(-5)"   → 기호=국내부족, 박스=None, 수량=-5, 상태=부족

    Returns:
        dict: {"기호", "박스", "수량", "상태"}
              상태: "피킹가능" | "부족" | "알수없음"
    """
    if pd.isna(box_str) or str(box_str).strip() == "":
        return {"기호": None, "박스": None, "수량": None, "상태": "알수없음"}

    box_str = str(box_str).strip()

    # ── 패턴 1: "부족(-1)" 또는 "국내부족(-5)" — 부족분 (피킹 불가) ──
    match = re.match(r"((?:국내)?부족)\((-?\d+)\)", box_str)
    if match:
        return {
            "기호": match.group(1),
            "박스": None,
            "수량": int(match.group(2)),  # 음수값 유지
            "상태": "부족",
        }

    # ── 패턴 2: "●6(3)" — 특수기호 + 박스번호 + (수량) ──
    match = re.match(r"([●★■▲◆◇○□△▼♦♠♣♥☆※·]+)(\d+)\((\d+)\)", box_str)
    if match:
        return {
            "기호": match.group(1),
            "박스": match.group(2),
            "수량": int(match.group(3)),
            "상태": "피킹가능",
        }

    # ── 패턴 3: "★1" — 특수기호 + 박스번호 (배대지 시트 형태) ──
    match = re.match(r"([●★■▲◆◇○□△▼♦♠♣♥☆※·]+)(\d+)", box_str)
    if match:
        return {
            "기호": match.group(1),
            "박스": match.group(2),
            "수량": None,
            "상태": "피킹가능",
        }

    # ── 패턴 4: "국내재고" 또는 "국내재고(2)" — 국내 보유분 ──
    match = re.match(r"(국내재고)\((\d+)\)", box_str)
    if match:
        return {
            "기호": match.group(1),
            "박스": None,
            "수량": int(match.group(2)),
            "상태": "피킹가능",
        }
    if box_str == "국내재고":
        return {"기호": "국내재고", "박스": None, "수량": None, "상태": "피킹가능"}

    # ── 매칭 실패 ──
    return {"기호": box_str, "박스": None, "수량": None, "상태": "알수없음"}


# ============================================================
# ❹ 세션 상태 초기화
# ============================================================
def init_session():
    """앱 전역에서 사용하는 세션 변수를 초기화합니다."""
    defaults = {
        "df_출고": None,           # 출고지시서 전체 DataFrame
        "df_배대지": None,          # 배대지 입고 전체 DataFrame
        "selected_shipment": None,  # 현재 선택된 쉽먼트 운송장번호
        "picking_state": {},        # {바코드: {필요수량, 스캔수량, 상품명, 회차기호, 박스번호, ...}}
        "inventory_state": {},      # 배대지 재고: {(기호, 바코드): 잔여수량}
        "scan_log": [],             # 스캔 이력
        "last_scan_result": None,   # 마지막 스캔 피드백
        "scan_counter": 0,          # 입력 필드 리셋용 카운터
        "completed_shipments": set(),
        "shortage_items": [],           # 부족분 목록
        "data_loaded": False,       # 데이터 로드 완료 여부
        "gsheet_client": None,      # 구글 시트 클라이언트
        "use_gsheet": False,        # 구글 시트 모드 여부
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val


init_session()


# ============================================================
# ❺ 데이터 로드 (구글 시트 or CSV 자동 감지)
# ============================================================
def load_all_data():
    """
    출고지시서 + 배대지 입고 데이터를 로드합니다.
    구글 시트 연결이 가능하면 시트에서, 아니면 CSV 업로드에서 가져옵니다.
    """
    client = get_gsheet_client()

    if client:
        st.session_state.gsheet_client = client
        st.session_state.use_gsheet = True

        # 출고지시서 로드
        df_출고 = load_sheet_as_df(client, CONFIG["SHEET_출고지시서"])
        if df_출고 is not None and not df_출고.empty:
            st.session_state.df_출고 = clean_출고지시서(df_출고)

        # 배대지 입고 로드
        df_배대지 = load_sheet_as_df(client, CONFIG["SHEET_배대지입고"])
        if df_배대지 is not None and not df_배대지.empty:
            st.session_state.df_배대지 = clean_배대지(df_배대지)

        if st.session_state.df_출고 is not None:
            st.session_state.data_loaded = True
            return True

    return False


def clean_출고지시서(df):
    """출고지시서 DataFrame을 정제합니다."""
    df = df.copy()

    # 필수 컬럼 확인
    required = ["바코드", "상품명", "수량", "쉽먼트운송장번호"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"출고지시서 필수 컬럼 누락: {missing}")
        return None

    df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    df["쉽먼트운송장번호"] = df["쉽먼트운송장번호"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["바코드"] = df["바코드"].astype(str).str.strip()

    # 박스번호 파싱 → 회차 기호 + 부족 상태 추출
    if "박스번호" in df.columns:
        parsed = df["박스번호"].apply(parse_box_info)
        df["회차기호"] = parsed.apply(lambda x: x["기호"])
        df["박스넘버"] = parsed.apply(lambda x: x["박스"])
        df["박스내수량"] = parsed.apply(lambda x: x["수량"])
        df["피킹상태"] = parsed.apply(lambda x: x["상태"])  # 피킹가능/부족/알수없음

    return df


def clean_배대지(df):
    """배대지 입고 리스트 DataFrame을 정제합니다."""
    df = df.copy()

    # 바코드 컬럼 확인
    if "바코드" not in df.columns:
        st.error("배대지 시트에 '바코드' 컬럼이 없습니다")
        return None

    df["바코드"] = df["바코드"].astype(str).str.strip()

    # 수량 컬럼 정제
    if "수량" in df.columns:
        df["수량"] = pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)

    # 배대지주문수량도 정제
    if "배대지주문수량" in df.columns:
        df["배대지주문수량"] = pd.to_numeric(df["배대지주문수량"], errors="coerce").fillna(0).astype(int)

    # 박스번호에서 회차 기호 추출
    if "박스번호" in df.columns:
        parsed = df["박스번호"].apply(parse_box_info)
        df["회차기호"] = parsed.apply(lambda x: x["기호"])
        df["박스넘버"] = parsed.apply(lambda x: x["박스"])

    return df


def load_csv_fallback():
    """CSV 파일 업로드로 데이터를 로드합니다 (구글 시트 연결 실패 시)."""
    st.sidebar.markdown("---")
    st.sidebar.subheader("📂 CSV 파일 업로드")

    uploaded_출고 = st.sidebar.file_uploader(
        "출고지시서 CSV", type=["csv"], key="csv_출고"
    )
    uploaded_배대지 = st.sidebar.file_uploader(
        "배대지 입고 CSV (선택)", type=["csv"], key="csv_배대지"
    )

    if uploaded_출고:
        df = pd.read_csv(uploaded_출고, encoding="utf-8-sig")
        st.session_state.df_출고 = clean_출고지시서(df)
        if st.session_state.df_출고 is not None:
            st.session_state.data_loaded = True

    if uploaded_배대지:
        df = pd.read_csv(uploaded_배대지, encoding="utf-8-sig")
        st.session_state.df_배대지 = clean_배대지(df)


# ============================================================
# ❻ 배대지 재고 초기화
# ============================================================
def init_inventory():
    """
    배대지 입고 데이터로부터 회차 기호별 재고를 초기화합니다.

    inventory_state 구조:
        {(회차기호, 바코드): 잔여수량}

    예: ("★", "R248320860001"): 2
        ("●", "R245781620003"): 5
    """
    df = st.session_state.df_배대지
    if df is None or df.empty:
        return

    inventory = {}
    for _, row in df.iterrows():
        barcode = row["바코드"]
        symbol = row.get("회차기호", "기타")
        qty = row.get("수량", 0)

        key = (symbol, barcode)
        if key in inventory:
            inventory[key] += qty
        else:
            inventory[key] = qty

    st.session_state.inventory_state = inventory


# ============================================================
# ❼ 쉽먼트 선택 → 피킹 상태 초기화
# ============================================================
def init_picking(shipment_id):
    """
    선택한 쉽먼트의 피킹 상태를 초기화합니다.

    picking_state 구조:
        {바코드: {
            상품명, 필요수량, 스캔수량,
            회차기호, 박스번호, 박스내수량,
            배대지잔여,  ← 해당 회차의 배대지 재고
        }}
    """
    df = st.session_state.df_출고
    shipment_df = df[df["쉽먼트운송장번호"] == shipment_id]

    if shipment_df.empty:
        st.error(f"쉽먼트 {shipment_id}를 찾을 수 없습니다")
        return

    picking = {}
    shortage_items = []  # 부족분 따로 추적

    for _, row in shipment_df.iterrows():
        bc = row["바코드"]
        symbol = row.get("회차기호", "")
        qty = row["수량"]
        pick_status = row.get("피킹상태", "피킹가능")

        # ── 부족 항목: 피킹 대상에서 제외하고 별도 표시 ──
        if pick_status == "부족":
            shortage_items.append({
                "바코드": bc,
                "상품명": row["상품명"],
                "부족수량": abs(row.get("박스내수량", 0) or 0),
                "박스번호": row.get("박스번호", ""),
            })
            continue

        if bc in picking:
            # 같은 바코드가 여러 행일 수 있음 (다른 박스에서)
            picking[bc]["필요수량"] += qty
        else:
            # 배대지 재고 확인
            inv_key = (symbol, bc)
            inv_qty = st.session_state.inventory_state.get(inv_key, None)

            picking[bc] = {
                "상품명": row["상품명"],
                "필요수량": qty,
                "스캔수량": 0,
                "회차기호": symbol if symbol else "N/A",
                "박스번호": row.get("박스번호", ""),
                "박스넘버": row.get("박스넘버", ""),
                "박스내수량": row.get("박스내수량", None),
                "배대지잔여": inv_qty,  # None이면 배대지 데이터 없음
                "SKU_ID": row.get("SKU ID", ""),
                "물류센터": row.get("물류센터(FC)", ""),
            }

    st.session_state.picking_state = picking
    st.session_state.shortage_items = shortage_items  # 부족분 저장
    st.session_state.selected_shipment = shipment_id
    st.session_state.scan_log = []
    st.session_state.last_scan_result = None
    st.session_state.scan_counter = 0


# ============================================================
# ❽ 바코드 스캔 처리 — 3중 검증 핵심 로직
# ============================================================
def process_scan(barcode):
    """
    스캔된 바코드를 3중 검증합니다.

    검증 1: 출고지시서 — 이 쉽먼트에 있는 바코드인가?
    검증 2: 수량 확인 — 필요 수량을 초과하지 않았는가?
    검증 3: 배대지 재고 — 해당 회차에 실제 재고가 있는가?

    스캔 성공 시 배대지 재고도 동시 차감합니다.
    """
    barcode = barcode.strip()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state = st.session_state.picking_state
    inventory = st.session_state.inventory_state

    if barcode not in state:
        # ── 오피킹: 이 쉽먼트에 없는 바코드 ──
        hint = _find_barcode_hint(barcode)
        result = {
            "status": "error",
            "message": "🚨 오피킹! 이 쉽먼트에 없는 바코드",
            "detail": f"{barcode}{hint}",
            "barcode": barcode, "상품명": "", "시간": now,
        }
        _log_scan(result, now)
        return result

    item = state[barcode]
    item["스캔수량"] += 1

    # ── 수량 초과 체크 ──
    if item["스캔수량"] > item["필요수량"]:
        result = {
            "status": "over",
            "message": f"⚠️ 수량 초과! {item['상품명'][:35]}",
            "detail": f"필요 {item['필요수량']}개인데 {item['스캔수량']}번째 스캔",
            "barcode": barcode, "상품명": item["상품명"], "시간": now,
        }
        _log_scan(result, now)
        return result

    # ── 배대지 재고 차감 ──
    symbol = item["회차기호"]
    inv_key = (symbol, barcode)
    shortage_warning = ""

    if inv_key in inventory:
        if inventory[inv_key] > 0:
            inventory[inv_key] -= 1
            item["배대지잔여"] = inventory[inv_key]
        else:
            # 배대지 재고 소진 — 실제로 물건이 안 온 것
            shortage_warning = f" | ⚠ {symbol}회차 배대지 재고 소진!"
            item["배대지잔여"] = 0

    # ── 정상 스캔 ──
    remaining = item["필요수량"] - item["스캔수량"]
    result = {
        "status": "ok" if not shortage_warning else "shortage",
        "message": f"✅ {item['상품명'][:35]}",
        "detail": f"스캔 {item['스캔수량']}/{item['필요수량']} (남은: {remaining}){shortage_warning}",
        "barcode": barcode, "상품명": item["상품명"], "시간": now,
    }
    _log_scan(result, now)
    return result


def _find_barcode_hint(barcode):
    """오피킹 시 이 바코드가 다른 쉽먼트에 있는지 찾아줍니다."""
    df = st.session_state.df_출고
    if df is None:
        return ""
    match = df[df["바코드"] == barcode]
    if not match.empty:
        name = match["상품명"].iloc[0][:25]
        others = match["쉽먼트운송장번호"].unique()[:3]
        return f" → [{name}] 다른 쉽먼트에 있음: {', '.join(s[-6:] for s in others)}"
    return " → 출고지시서에 없는 바코드"


def _log_scan(result, now):
    """스캔 결과를 로그에 기록하고 구글 시트에도 저장합니다."""
    st.session_state.scan_log.append(result)
    st.session_state.last_scan_result = result
    st.session_state.scan_counter += 1

    # 구글 시트 로그 기록 (비동기적으로 — 실패해도 피킹 중단 안 함)
    if st.session_state.use_gsheet and st.session_state.gsheet_client:
        item = st.session_state.picking_state.get(result["barcode"], {})
        log_row = [
            now,
            st.session_state.selected_shipment or "",
            result["barcode"],
            result.get("상품명", "")[:40],
            result["status"],
            item.get("스캔수량", ""),
            item.get("필요수량", ""),
            item.get("회차기호", ""),
            item.get("박스번호", ""),
        ]
        append_log_to_sheet(st.session_state.gsheet_client, log_row)


# ============================================================
# ❾ 진행률 계산
# ============================================================
def get_progress():
    """피킹 진행 상황을 계산합니다."""
    state = st.session_state.picking_state
    if not state:
        return {"total": 0, "scanned": 0, "skus": 0, "done_skus": 0,
                "pct": 0.0, "is_complete": False, "over": 0, "shortage": 0}

    total = sum(v["필요수량"] for v in state.values())
    scanned = sum(min(v["스캔수량"], v["필요수량"]) for v in state.values())
    over = sum(max(0, v["스캔수량"] - v["필요수량"]) for v in state.values())
    skus = len(state)
    done_skus = sum(1 for v in state.values() if v["스캔수량"] >= v["필요수량"])

    # 배대지 재고 부족 건수
    shortage = sum(
        1 for v in state.values()
        if v.get("배대지잔여") is not None and v["배대지잔여"] == 0 and v["스캔수량"] < v["필요수량"]
    )

    pct = scanned / total if total > 0 else 0

    return {
        "total": total, "scanned": scanned, "skus": skus,
        "done_skus": done_skus, "pct": pct,
        "is_complete": scanned >= total, "over": over, "shortage": shortage,
    }


# ============================================================
# ❿ 메인 UI
# ============================================================
def main():
    st.markdown("## 📦 쿠팡 피킹 검증 시스템")
    st.caption("바코드 스캔 → 출고지시서 검증 + 배대지 재고 동시 차감")

    # ── 사이드바: 설정 & 데이터 연결 ──
    with st.sidebar:
        st.header("⚙️ 설정")

        # 구글 시트 연결 시도
        mode = st.radio(
            "데이터 소스",
            ["📊 구글 시트 (실시간)", "📂 CSV 파일 업로드"],
            index=0,
        )

        if mode == "📊 구글 시트 (실시간)":
            if st.button("🔄 구글 시트 연결", use_container_width=True):
                with st.spinner("구글 시트 연결 중..."):
                    success = load_all_data()
                    if success:
                        init_inventory()
                        st.success("✅ 구글 시트 연결 완료!")
                        st.rerun()
                    else:
                        st.error("연결 실패 — CSV 모드를 사용하세요")
        else:
            load_csv_fallback()
            if st.session_state.data_loaded and st.session_state.df_배대지 is not None:
                init_inventory()

        # 데이터 상태 표시
        if st.session_state.df_출고 is not None:
            n_rows = len(st.session_state.df_출고)
            n_ship = st.session_state.df_출고["쉽먼트운송장번호"].nunique()
            st.success(f"출고지시서: {n_rows}행 / {n_ship}개 쉽먼트")

        if st.session_state.df_배대지 is not None:
            n_inv = len(st.session_state.df_배대지)
            st.success(f"배대지 입고: {n_inv}행 로드됨")
        else:
            st.info("배대지 시트 미연결 (피킹 검증만 가능)")

        # 로그 다운로드
        if st.session_state.scan_log:
            st.sidebar.divider()
            log_df = pd.DataFrame(st.session_state.scan_log)
            st.sidebar.download_button(
                "📥 스캔 로그 CSV",
                data=log_df.to_csv(index=False, encoding="utf-8-sig"),
                file_name=f"picking_log_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True,
            )

    # ── 메인: 데이터 없으면 안내 ──
    if not st.session_state.data_loaded:
        st.info("👈 사이드바에서 데이터를 연결하세요 (구글 시트 또는 CSV)")
        _show_setup_guide()
        return

    # ── 송장번호 입력 (쉽먼트 선택) ──
    if not st.session_state.selected_shipment:
        _show_shipment_selector()
        return

    # ── 피킹 진행 화면 ──
    _show_picking_screen()


# ============================================================
# UI 서브 화면들
# ============================================================
def _show_setup_guide():
    """초기 설정 가이드를 표시합니다."""
    st.markdown("""
    **초기 설정 방법 (구글 시트 모드):**
    1. `service_account.json` 파일을 앱 폴더에 배치
    2. 구글 스프레드시트에 서비스 계정 이메일을 **편집자**로 공유
    3. `picking_app_v2.py` 상단의 `CONFIG`에서 스프레드시트 ID와 시트 탭 이름 설정
    4. 사이드바에서 '구글 시트 연결' 클릭

    **또는 CSV 모드:**
    1. 사이드바에서 'CSV 파일 업로드' 선택
    2. 출고지시서 CSV 업로드 (필수)
    3. 배대지 입고 CSV 업로드 (선택 — 재고 연동 시 필요)
    """)


def _show_shipment_selector():
    """송장번호 입력 팝업을 표시합니다."""
    st.markdown('<div class="shipment-input">', unsafe_allow_html=True)
    st.markdown("### 📋 쉽먼트 선택")
    st.markdown("송장번호를 입력하거나 목록에서 선택하세요")

    col1, col2 = st.columns([2, 1])

    with col1:
        # 직접 입력
        input_shipment = st.text_input(
            "송장번호 직접 입력",
            placeholder="운송장번호 입력 후 Enter",
            key="shipment_input",
        )

    with col2:
        # 드롭다운 선택
        df = st.session_state.df_출고

        # 물류센터 필터
        centers = ["전체"] + sorted(df["물류센터(FC)"].unique().tolist()) if "물류센터(FC)" in df.columns else ["전체"]
        center = st.selectbox("물류센터", centers, key="center_filter")

        if center != "전체":
            filtered = df[df["물류센터(FC)"] == center]
        else:
            filtered = df

    # 쉽먼트 목록 테이블
    summary = filtered.groupby("쉽먼트운송장번호").agg(
        SKU수=("바코드", "nunique"),
        총수량=("수량", "sum"),
        센터=("물류센터(FC)", "first") if "물류센터(FC)" in filtered.columns else ("수량", "count"),
        회차=("회차기호", "first") if "회차기호" in filtered.columns else ("수량", "count"),
    ).reset_index().sort_values("총수량", ascending=False)

    # 선택 가능한 테이블
    selected_shipment = st.selectbox(
        "또는 목록에서 선택",
        options=summary["쉽먼트운송장번호"].tolist(),
        format_func=lambda x: (
            f"{'✅ ' if x in st.session_state.completed_shipments else ''}"
            f"{x[-6:]} | "
            f"{summary[summary['쉽먼트운송장번호']==x]['센터'].values[0] if '센터' in summary.columns else ''} | "
            f"{summary[summary['쉽먼트운송장번호']==x]['SKU수'].values[0]}종 "
            f"{summary[summary['쉽먼트운송장번호']==x]['총수량'].values[0]}개"
        ),
        key="shipment_select",
    )

    # 시작 버튼
    target = input_shipment.strip() if input_shipment else selected_shipment

    if st.button("🚀 피킹 시작", type="primary", use_container_width=True):
        if target:
            # 입력된 송장번호가 유효한지 확인
            valid_ids = df["쉽먼트운송장번호"].unique()
            if target in valid_ids:
                init_picking(target)
                st.rerun()
            else:
                # 부분 매칭 시도 (뒤 6자리)
                matches = [s for s in valid_ids if s.endswith(target)]
                if len(matches) == 1:
                    init_picking(matches[0])
                    st.rerun()
                elif len(matches) > 1:
                    st.warning(f"'{target}'에 매칭되는 쉽먼트가 {len(matches)}개입니다. 전체 번호를 입력해주세요.")
                else:
                    st.error(f"'{target}'에 해당하는 쉽먼트를 찾을 수 없습니다.")

    st.markdown('</div>', unsafe_allow_html=True)


def _show_picking_screen():
    """피킹 진행 화면을 표시합니다."""
    progress = get_progress()
    shipment_id = st.session_state.selected_shipment

    # ── 헤더: 쉽먼트 정보 + 다른 쉽먼트로 전환 버튼 ──
    hcol1, hcol2 = st.columns([4, 1])
    with hcol1:
        item0 = list(st.session_state.picking_state.values())[0] if st.session_state.picking_state else {}
        center = item0.get("물류센터", "")
        symbol = item0.get("회차기호", "")
        st.markdown(f"**쉽먼트:** `{shipment_id}` | **센터:** {center} | **회차:** {symbol}")
    with hcol2:
        if st.button("🔄 다른 쉽먼트", use_container_width=True):
            st.session_state.selected_shipment = None
            st.rerun()

    # ── 진행률 대시보드 ──
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("스캔", f"{progress['scanned']}/{progress['total']}")
    c2.metric("SKU 완료", f"{progress['done_skus']}/{progress['skus']}")
    c3.metric("진행률", f"{progress['pct']:.0%}")
    c4.metric("초과 스캔", f"{progress['over']}건",
              delta=f"+{progress['over']}" if progress['over'] > 0 else None,
              delta_color="inverse")
    c5.metric("재고 부족", f"{progress['shortage']}건",
              delta=f"{progress['shortage']}" if progress['shortage'] > 0 else None,
              delta_color="inverse")

    st.progress(progress["pct"])

    # ── 피킹 완료 ──
    if progress["is_complete"]:
        st.markdown(
            '<div class="scan-complete">'
            f'<strong style="font-size:1.3rem;">🎉 피킹 완료!</strong><br>'
            f'쉽먼트 {shipment_id[-6:]} — {progress["total"]}개 전부 검증 완료'
            '</div>', unsafe_allow_html=True
        )
        st.session_state.completed_shipments.add(shipment_id)

    # ── 바코드 스캔 입력 ──
    st.markdown("---")
    scan_key = f"scan_{st.session_state.scan_counter}"
    scanned = st.text_input(
        "🔫 바코드 스캔 (스캐너 또는 직접 입력)",
        key=scan_key,
        placeholder="스캐너 대기 중... 바코드를 스캔하세요",
    )

    if scanned:
        process_scan(scanned)
        st.rerun()

    # ── 마지막 스캔 결과 피드백 ──
    r = st.session_state.last_scan_result
    if r:
        css_class = {
            "ok": "scan-ok", "over": "scan-warning",
            "error": "scan-error", "shortage": "scan-shortage",
        }.get(r["status"], "scan-ok")
        st.markdown(
            f'<div class="{css_class}">'
            f'<strong style="font-size:1.1rem;">{r["message"]}</strong><br>'
            f'{r["detail"]}'
            f'</div>', unsafe_allow_html=True
        )

    # ── 피킹 현황 테이블 ──
    st.markdown("---")
    st.subheader("📋 피킹 현황")

    rows = []
    for bc, info in st.session_state.picking_state.items():
        s, n = info["스캔수량"], info["필요수량"]
        if s > n:
            status = f"⚠️ 초과 ({s}/{n})"
        elif s >= n:
            status = "✅ 완료"
        elif s > 0:
            status = f"🔄 {s}/{n}"
        else:
            status = "⬜ 대기"

        # 배대지 재고 표시
        inv = info.get("배대지잔여")
        inv_display = f"{inv}" if inv is not None else "-"

        rows.append({
            "상태": status,
            "바코드": bc,
            "상품명": info["상품명"][:35] + ("..." if len(info["상품명"]) > 35 else ""),
            "필요": n,
            "스캔": s,
            "남은": max(0, n - s),
            "회차": info.get("회차기호", ""),
            "박스": info.get("박스번호", ""),
            "배대지재고": inv_display,
        })

    # 정렬: 진행중 → 대기 → 완료 → 초과
    order = {"🔄": 0, "⬜": 1, "✅": 2, "⚠️": 3}
    rows.sort(key=lambda x: order.get(x["상태"][0], 9))

    st.dataframe(
        pd.DataFrame(rows),
        use_container_width=True,
        hide_index=True,
        height=min(500, len(rows) * 38 + 40),
    )

    # ── 부족분 표시 (피킹 불가 항목) ──
    shortage = st.session_state.get("shortage_items", [])
    if shortage:
        with st.expander(f"⛔ 부족분 — 피킹 불가 ({len(shortage)}건)", expanded=False):
            st.caption("출고지시서에 '부족'으로 표시된 항목입니다. 중국에서 미입고되었거나 국내 재고가 부족합니다.")
            short_df = pd.DataFrame(shortage)
            st.dataframe(short_df, use_container_width=True, hide_index=True)

    # ── 스캔 로그 ──
    if st.session_state.scan_log:
        with st.expander(f"📜 스캔 로그 ({len(st.session_state.scan_log)}건)"):
            log_display = []
            for entry in reversed(st.session_state.scan_log[-50:]):
                icon = {"ok": "✅", "over": "⚠️", "error": "🚨", "shortage": "📦"}.get(entry["status"], "?")
                log_display.append({
                    "시간": entry["시간"],
                    "결과": icon,
                    "바코드": entry["barcode"],
                    "내용": entry["message"],
                })
            st.dataframe(pd.DataFrame(log_display), use_container_width=True, hide_index=True)


# ============================================================
# 앱 실행
# ============================================================
if __name__ == "__main__":
    main()
