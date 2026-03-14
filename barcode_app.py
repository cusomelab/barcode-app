# ╔══════════════════════════════════════════════════════╗
# ║         [쿠썸] 바코드 라벨 생성기 - Streamlit         ║
# ╚══════════════════════════════════════════════════════╝
import os, io, urllib.request, csv, zipfile
from datetime import datetime, timedelta
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable, PageBreak)
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from pypdf import PdfWriter, PdfReader

# ── 폰트 준비 ──────────────────────────────────────────
FONT_PATH = 'NanumGothicBold.ttf'
FONT_REG_PATH = 'NanumGothic.ttf'
if not os.path.exists(FONT_PATH):
    urllib.request.urlretrieve(
        'https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Bold.ttf',
        FONT_PATH
    )
if not os.path.exists(FONT_REG_PATH):
    urllib.request.urlretrieve(
        'https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Regular.ttf',
        FONT_REG_PATH
    )

# reportlab 한국어 폰트 등록
try:
    pdfmetrics.registerFont(TTFont('NanumBold', FONT_PATH))
    pdfmetrics.registerFont(TTFont('NanumReg', FONT_REG_PATH))
except Exception:
    pass

# ══════════════════════════════════════════════════════
# ── 공통 헬퍼 (바코드 라벨용) ──────────────────────────
# ══════════════════════════════════════════════════════
def wrap_text(text, font, max_w, draw):
    lines, cur = [], []
    for word in text.split(' '):
        test = ' '.join(cur + [word])
        bbox = draw.textbbox((0,0), test, font=font)
        if bbox[2]-bbox[0] <= max_w: cur.append(word)
        else:
            if cur: lines.append(' '.join(cur)); cur=[word]
            else:
                tmp=''
                for ch in word:
                    tb=draw.textbbox((0,0),tmp+ch,font=font)
                    if tb[2]-tb[0]<=max_w: tmp+=ch
                    else:
                        if tmp: lines.append(tmp)
                        tmp=ch
                if tmp: lines.append(tmp)
                cur=[]
    if cur: lines.append(' '.join(cur))
    return lines

def get_barcode_img(barcode_number, write_text=False):
    bc = barcode.get('code128', barcode_number, writer=ImageWriter())
    bc.writer.set_options({'module_height':80,'module_width':1.6,
                           'quiet_zone':6,'write_text':write_text})
    raw = bc.render()
    g=raw.convert('L'); w_r,h_r=g.size; p=g.load()
    l =next(x for x in range(w_r)          if any(p[x,yy]<255 for yy in range(h_r)))
    rr=next(x for x in range(w_r-1,-1,-1)   if any(p[x,yy]<255 for yy in range(h_r)))+1
    t =next(yy for yy in range(h_r)         if any(p[xx,yy]<255 for xx in range(w_r)))
    if write_text:
        bar_end=t
        for yy in range(t,h_r):
            if sum(1 for xx in range(l,rr) if p[xx,yy]<100)/(rr-l)>=0.25: bar_end=yy
        return raw.crop((l,t,rr,bar_end+2))
    else:
        b=next(yy for yy in range(h_r-1,-1,-1) if any(p[xx,yy]<255 for xx in range(w_r)))+1
        return raw.crop((l,t,rr,b))

# ── 소형 라벨 생성 ─────────────────────────────────────
def create_small(product_name, barcode_number, material, fixed_origin, fixed_age):
    CANVAS_W, CANVAS_H = 650, 450; PAD = 30
    img=Image.new('RGB',(CANVAS_W,CANVAS_H),'white'); draw=ImageDraw.Draw(img)
    font_big=ImageFont.truetype(FONT_PATH,26)
    font_mid=ImageFont.truetype(FONT_PATH,20)

    y=PAD
    for line in wrap_text(product_name,font_big,CANVAS_W-PAD*2,draw)[:2]:
        bb=draw.textbbox((0,0),line,font=font_big)
        draw.text(((CANVAS_W-(bb[2]-bb[0]))//2,y),line,font=font_big,fill='black')
        y+=bb[3]-bb[1]+6

    bc_img=get_barcode_img(barcode_number,write_text=False)
    c1b=draw.textbbox((0,0),fixed_origin,font=font_mid)
    c2b=draw.textbbox((0,0),fixed_age,font=font_mid)
    mat_h=font_mid.size+8 if material else 0
    fixed_h=(c1b[3]-c1b[1])+(c2b[3]-c2b[1])+mat_h+18
    bc_y=y+10; cue_y=CANVAS_H-fixed_h-PAD
    BAR_W=CANVAS_W-PAD*2; BAR_H=cue_y-bc_y-6
    if BAR_H<60: BAR_H=60
    img.paste(bc_img.resize((BAR_W,BAR_H),Image.LANCZOS),(PAD,bc_y))
    cur_y=bc_y+BAR_H+8

    if material:
        mt=f'재질 : {material}'; mb=draw.textbbox((0,0),mt,font=font_mid)
        draw.text(((CANVAS_W-(mb[2]-mb[0]))//2,cur_y),mt,font=font_mid,fill='black')
        cur_y+=mb[3]-mb[1]+6
    for txt in (fixed_origin,fixed_age):
        tb=draw.textbbox((0,0),txt,font=font_mid)
        draw.text(((CANVAS_W-(tb[2]-tb[0]))//2,cur_y),txt,font=font_mid,fill='black')
        cur_y+=tb[3]-tb[1]+5
    return img

# ── 대형 라벨 생성 ─────────────────────────────────────
def fit_font(text, max_w, draw, max_size=42, min_size=8):
    for size in range(max_size,min_size-1,-1):
        f=ImageFont.truetype(FONT_PATH,size)
        bb=draw.textbbox((0,0),text,font=f)
        if bb[2]-bb[0]<=max_w: return f
    return ImageFont.truetype(FONT_PATH,min_size)

def create_large(product_name, barcode_number, material, fix_list):
    CANVAS_W,CANVAS_H=450,640; PAD=22
    img=Image.new('RGB',(CANVAS_W,CANVAS_H),'white'); draw=ImageDraw.Draw(img)
    fn=ImageFont.truetype(FONT_PATH,22)
    fm=ImageFont.truetype(FONT_PATH,20)
    ff=ImageFont.truetype(FONT_PATH,16)

    y=PAD
    for line in wrap_text(product_name,fn,CANVAS_W-PAD*2,draw)[:3]:
        bb=draw.textbbox((0,0),line,font=fn)
        draw.text(((CANVAS_W-(bb[2]-bb[0]))//2,y),line,font=fn,fill='black')
        y+=bb[3]-bb[1]+4
    y+=8

    bc_img=get_barcode_img(barcode_number,write_text=True)
    fix_h=0
    for txt in fix_list:
        for ln in wrap_text(txt,ff,CANVAS_W-PAD*2,draw):
            bb=draw.textbbox((0,0),ln,font=ff); fix_h+=bb[3]-bb[1]+3
        fix_h+=6

    BAR_W=CANVAS_W-PAD*2
    BAR_H=CANVAS_H-y-50-30-14-fix_h-PAD-16
    if BAR_H<80: BAR_H=80
    img.paste(bc_img.resize((BAR_W,BAR_H),Image.LANCZOS),(PAD,y)); y+=BAR_H+6

    font_bc=fit_font(barcode_number,BAR_W,draw)
    nb=draw.textbbox((0,0),barcode_number,font=font_bc)
    draw.text(((CANVAS_W-(nb[2]-nb[0]))//2,y),barcode_number,font=font_bc,fill='black')
    y+=nb[3]-nb[1]+8

    if material:
        mt=f'재질 : {material}'; mb=draw.textbbox((0,0),mt,font=fm)
        draw.text(((CANVAS_W-(mb[2]-mb[0]))//2,y),mt,font=fm,fill='black')
        y+=mb[3]-mb[1]+8

    draw.line([(PAD,y),(CANVAS_W-PAD,y)],fill=(160,160,160),width=1); y+=10
    for txt in fix_list:
        for ln in wrap_text(txt,ff,CANVAS_W-PAD*2,draw):
            bb=draw.textbbox((0,0),ln,font=ff)
            draw.text((PAD,y),ln,font=ff,fill='black'); y+=bb[3]-bb[1]+3
        y+=6

    draw.rectangle([1,1,CANVAS_W-2,CANVAS_H-2],outline=(180,180,180),width=1)
    return img.rotate(90,expand=True)

# ── 엑셀 처리 공통 ─────────────────────────────────────
def process_excel(uploaded_file, mode, settings):
    wb=load_workbook(uploaded_file); ws=wb.active
    col_insert=settings['col_insert']
    ws.column_dimensions[get_column_letter(col_insert)].width=settings['col_width']

    temp_dir='_temp'; os.makedirs(temp_dir,exist_ok=True)
    ok=0; errors=[]
    progress=st.progress(0)
    status=st.empty()

    total=sum(1 for r in range(settings['start_row'],ws.max_row+1)
              if ws.cell(r,settings['col_barcode']).value)

    for r in range(settings['start_row'],ws.max_row+1):
        bv=ws.cell(r,settings['col_barcode']).value
        if not bv: continue
        bv=str(bv).strip()
        nm=str(ws.cell(r,settings['col_name']).value or '').strip()
        mt=str(ws.cell(r,settings['col_material']).value or '').strip()
        img_path=f'{temp_dir}/label_{r}.png'

        try:
            if mode=='소형':
                img=create_small(nm,bv,mt,settings['origin'],settings['age'])
            else:
                img=create_large(nm,bv,mt,settings['fix_list'])
            img.save(img_path)
        except Exception as e:
            errors.append(f'{r}행 실패: {e}'); continue

        xl=XLImage(img_path)
        xl.width=settings['insert_w']; xl.height=settings['insert_h']
        ws.add_image(xl,f'{get_column_letter(col_insert)}{r}')
        ws.row_dimensions[r].height=settings['row_height']
        ok+=1
        progress.progress(ok/max(total,1))
        status.text(f'✅ {r}행 처리 중... ({ok}/{total})')

    output=io.BytesIO(); wb.save(output); output.seek(0)
    progress.progress(1.0); status.text(f'🎉 완료! {ok}개 생성')
    return output, ok, errors


# ══════════════════════════════════════════════════════
# ── 출고 작업 지시서 PDF 생성 함수들 ──────────────────
# ══════════════════════════════════════════════════════

def parse_date(date_str):
    """날짜 문자열을 파싱해서 datetime 반환. 실패하면 None."""
    if not date_str or not date_str.strip():
        return None
    s = date_str.strip()
    nums = [n for n in __import__('re').findall(r'\d+', s)]
    try:
        if len(nums) == 1 and len(nums[0]) == 8:
            ymd = nums[0]
            return datetime(int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8]))
        elif len(nums) >= 3:
            y, m, d = int(nums[0]), int(nums[1]), int(nums[2])
            if y < 100: y += 2000
            return datetime(y, m, d)
    except Exception:
        pass
    return None


def calc_deadline(date_str):
    """입고예정일 + 20일 → 입고마감일 문자열 반환"""
    dt = parse_date(date_str)
    if dt is None:
        return '날짜 없음'
    deadline = dt + timedelta(days=20)
    return deadline.strftime('%Y-%m-%d')


def parse_csv_to_items(file_bytes):
    """CSV 바이트 → WorkOrderItem 리스트 반환"""
    text = file_bytes.decode('utf-8-sig', errors='replace')
    reader = csv.reader(text.splitlines())
    rows = list(reader)
    if not rows:
        return []

    # 헤더 행 건너뛰기 (첫 행)
    data_rows = rows[1:]
    items = []
    for row in data_rows:
        if len(row) < 11 or not (row[1] if len(row) > 1 else '').strip():
            continue
        def safe(idx, default=''):
            return row[idx].strip() if idx < len(row) else default
        try:
            qty = int(safe(7, '0') or '0')
        except ValueError:
            qty = 0
        items.append({
            'logisticsCenter': safe(1),
            'expectedDate':    safe(3),
            'productBarcode':  safe(5),
            'productName':     safe(6),
            'quantity':        qty,
            'shipmentNumber':  safe(8),
            'orderDate':       safe(9),
            'boxNumber':       safe(10),
            'location':        safe(12),
        })
    return items


def group_items(items, grouping_keys):
    """그룹화 기준에 따라 dict로 묶음"""
    grouped = {}
    for item in items:
        key_parts = [item.get(k, '') for k in grouping_keys if item.get(k)]
        key = '_'.join(key_parts) if key_parts else '(미분류)'
        grouped.setdefault(key, []).append(item)

    # 각 그룹 내부 정렬: 박스번호 → 상품명
    import locale
    for key in grouped:
        grouped[key].sort(key=lambda x: (
            [int(c) if c.isdigit() else c.lower()
             for c in __import__('re').split(r'(\d+)', x.get('boxNumber',''))],
            x.get('productName','')
        ))
    return grouped


def create_work_order_pdf(group_key, items):
    """reportlab으로 출고 작업 지시서 PDF 생성 → BytesIO 반환"""
    buf = io.BytesIO()
    PAGE_W, PAGE_H = A4
    MARGIN = 18 * mm

    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=MARGIN
    )

    # 스타일 정의
    s_title   = ParagraphStyle('title',   fontName='NanumBold', fontSize=18, leading=22, textColor=colors.HexColor('#111111'))
    s_sub     = ParagraphStyle('sub',     fontName='NanumBold', fontSize=9,  leading=12, textColor=colors.HexColor('#888888'), spaceAfter=2)
    s_card_lbl= ParagraphStyle('cardlbl', fontName='NanumBold', fontSize=8,  leading=10, textColor=colors.HexColor('#888888'))
    s_card_val= ParagraphStyle('cardval', fontName='NanumBold', fontSize=12, leading=15, textColor=colors.HexColor('#111111'))
    s_card_big= ParagraphStyle('cardbig', fontName='NanumBold', fontSize=22, leading=26, textColor=colors.HexColor('#1a56db'))
    s_th      = ParagraphStyle('th',      fontName='NanumBold', fontSize=9,  leading=11, textColor=colors.white)
    s_td      = ParagraphStyle('td',      fontName='NanumReg',  fontSize=9,  leading=12, textColor=colors.HexColor('#111111'))
    s_td_bold = ParagraphStyle('tdbold',  fontName='NanumBold', fontSize=10, leading=12, textColor=colors.HexColor('#1a56db'))
    s_mono    = ParagraphStyle('mono',    fontName='NanumReg',  fontSize=9,  leading=12, textColor=colors.HexColor('#111111'))
    s_footer  = ParagraphStyle('footer',  fontName='NanumReg',  fontSize=8,  leading=10, textColor=colors.HexColor('#888888'))

    total_qty   = sum(i['quantity'] for i in items)
    first       = items[0]
    deadline    = calc_deadline(first.get('expectedDate',''))
    created_at  = datetime.now().strftime('%Y-%m-%d %H:%M')
    usable_w    = PAGE_W - MARGIN * 2

    story = []

    # ── 헤더 ──────────────────────────────────────────
    story.append(Paragraph('출고 작업 지시서', s_title))
    story.append(Paragraph(group_key, s_sub))
    story.append(Spacer(1, 1*mm))
    story.append(HRFlowable(width='100%', thickness=2, color=colors.HexColor('#1a56db')))
    story.append(Spacer(1, 4*mm))

    # ── 정보 카드 (3열 테이블) ─────────────────────────
    col_w = usable_w / 3

    def info_card(label, value, big=False):
        lbl = Paragraph(label, s_card_lbl)
        val = Paragraph(str(value), s_card_big if big else s_card_val)
        return [lbl, val]

    card_data = [[
        info_card('물류센터',  first.get('logisticsCenter','-')),
        info_card('입고예정일', first.get('expectedDate','-')),
        info_card('총 수량',   total_qty, big=True),
    ],[
        info_card('송장번호',  first.get('shipmentNumber','-')),
        info_card('입고마감일', deadline),
        info_card('품목 수',   f'{len(items)}개'),
    ]]

    def make_card_cell(label_val_list):
        """[lbl_para, val_para] → 테이블 셀용 nested table"""
        t = Table([[label_val_list[0]], [label_val_list[1]]], colWidths=[col_w - 6*mm])
        t.setStyle(TableStyle([
            ('LEFTPADDING',  (0,0),(-1,-1), 0),
            ('RIGHTPADDING', (0,0),(-1,-1), 0),
            ('TOPPADDING',   (0,0),(-1,-1), 1),
            ('BOTTOMPADDING',(0,0),(-1,-1), 1),
        ]))
        return t

    card_table_data = []
    for row in card_data:
        card_table_data.append([make_card_cell(cell) for cell in row])

    card_bg = [colors.HexColor('#f8fafc'), colors.HexColor('#eff6ff')]
    card_tbl = Table(card_table_data, colWidths=[col_w]*3)
    card_style = [
        ('BACKGROUND', (0,0),(2,0), card_bg[0]),
        ('BACKGROUND', (0,1),(2,1), card_bg[1]),
        ('BOX',        (0,0),(2,1), 1, colors.HexColor('#e2e8f0')),
        ('INNERGRID',  (0,0),(2,1), 0.5, colors.HexColor('#e2e8f0')),
        ('LEFTPADDING',  (0,0),(-1,-1), 4*mm),
        ('RIGHTPADDING', (0,0),(-1,-1), 2*mm),
        ('TOPPADDING',   (0,0),(-1,-1), 3*mm),
        ('BOTTOMPADDING',(0,0),(-1,-1), 3*mm),
        ('VALIGN',     (0,0),(-1,-1), 'MIDDLE'),
        ('ROUNDEDCORNERS', [4]),
    ]
    card_tbl.setStyle(TableStyle(card_style))
    story.append(card_tbl)
    story.append(Spacer(1, 5*mm))

    # ── 상품 테이블 ────────────────────────────────────
    # 컬럼 폭: 바코드 27%, 상품명 35%, 수량 8%, 위치 14%, 박스 16%
    cw = [usable_w*p for p in [0.25, 0.35, 0.08, 0.16, 0.16]]

    header_row = [
        Paragraph('바코드', s_th),
        Paragraph('상품명', s_th),
        Paragraph('수량', s_th),
        Paragraph('위치', s_th),
        Paragraph('박스', s_th),
    ]
    table_data = [header_row]

    for i, item in enumerate(items):
        row = [
            Paragraph(item.get('productBarcode',''), s_mono),
            Paragraph(item.get('productName',''),    s_td),
            Paragraph(str(item.get('quantity',0)),   s_td_bold),
            Paragraph(item.get('location',''),       s_td),
            Paragraph(item.get('boxNumber',''),      s_td),
        ]
        table_data.append(row)

    # 합계 행
    table_data.append([
        Paragraph('합  계', ParagraphStyle('sum', fontName='NanumBold', fontSize=9, textColor=colors.HexColor('#111111'))),
        Paragraph('', s_td),
        Paragraph(str(total_qty), ParagraphStyle('sumqty', fontName='NanumBold', fontSize=12, textColor=colors.HexColor('#1a56db'))),
        Paragraph('', s_td),
        Paragraph('', s_td),
    ])

    tbl = Table(table_data, colWidths=cw, repeatRows=1)

    row_colors = []
    for i in range(1, len(table_data)-1):
        bg = colors.white if i % 2 == 1 else colors.HexColor('#f8fafc')
        row_colors.append(('BACKGROUND', (0,i),(4,i), bg))

    tbl_style = [
        # 헤더
        ('BACKGROUND', (0,0),(4,0), colors.HexColor('#1e293b')),
        ('TEXTCOLOR',  (0,0),(4,0), colors.white),
        ('ALIGN',      (0,0),(4,0), 'CENTER'),
        # 합계 행
        ('BACKGROUND', (0,-1),(4,-1), colors.HexColor('#eff6ff')),
        ('LINEABOVE',  (0,-1),(4,-1), 1, colors.HexColor('#93c5fd')),
        # 전체
        ('FONTSIZE',   (0,0),(-1,-1), 9),
        ('TOPPADDING', (0,0),(-1,-1), 4),
        ('BOTTOMPADDING',(0,0),(-1,-1), 4),
        ('LEFTPADDING',(0,0),(-1,-1), 3*mm),
        ('RIGHTPADDING',(0,0),(-1,-1), 2*mm),
        ('VALIGN',     (0,0),(-1,-1), 'MIDDLE'),
        ('ALIGN',      (2,1),(2,-1), 'CENTER'),  # 수량 가운데
        ('GRID',       (0,0),(-1,-1), 0.4, colors.HexColor('#e2e8f0')),
        ('LINEBELOW',  (0,0),(4,0), 1, colors.HexColor('#1a56db')),
    ] + row_colors

    tbl.setStyle(TableStyle(tbl_style))
    story.append(tbl)
    story.append(Spacer(1, 5*mm))

    # ── 푸터 ──────────────────────────────────────────
    story.append(HRFlowable(width='100%', thickness=0.5, color=colors.HexColor('#e2e8f0')))
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(
        f'※ 바코드와 수량을 작업 전 반드시 대조해 주세요. (자동 생성 문서) · 생성일시: {created_at}',
        s_footer
    ))

    doc.build(story)
    buf.seek(0)
    return buf


def merge_pdfs(pdf_buffers):
    """여러 PDF BytesIO를 하나로 병합 → BytesIO 반환"""
    writer = PdfWriter()
    for buf in pdf_buffers:
        reader = PdfReader(buf)
        for page in reader.pages:
            writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out


# ══════════════════════════════════════════════════════
# Streamlit UI
# ══════════════════════════════════════════════════════
st.set_page_config(page_title='로켓배송 출고 생성기', page_icon='🚀', layout='centered')
st.title('🚀 로켓배송 출고 생성기')
st.caption('바코드 라벨 생성 · 출고 작업 지시서 PDF 변환')

tab1, tab2, tab3 = st.tabs(['📦 소형 라벨', '📋 대형 라벨 (90도 회전)', '📄 출고 작업 지시서 PDF'])

# ── 소형 탭 ────────────────────────────────────────────
with tab1:
    st.subheader('📋 열 번호 설정')
    st.caption('A=1, B=2, C=3, D=4 ... L=12, M=13')
    c1,c2 = st.columns(2)
    with c1:
        s_col_name     = st.number_input('상품명 열',     min_value=1, max_value=50, value=4,  key='s_name')
        s_col_barcode  = st.number_input('바코드 열',     min_value=1, max_value=50, value=12, key='s_barcode')
        s_col_material = st.number_input('재질 열',       min_value=1, max_value=50, value=13, key='s_material')
    with c2:
        s_col_insert   = st.number_input('이미지 삽입 열', min_value=1, max_value=50, value=13, key='s_insert')
        s_start_row    = st.number_input('시작 행',       min_value=1, max_value=10,  value=2,  key='s_startrow')

    st.divider()
    st.subheader('✍️ 고정 문구')
    s_origin = st.text_input('제조국',   value='제조국 Made in China',             key='s_origin')
    s_age    = st.text_input('사용연령', value='본 제품은 14세 이상 사용가능합니다', key='s_age')

    st.divider()
    s_file = st.file_uploader('📂 엑셀 파일 업로드', type=['xlsx'], key='s_file')

    if st.button('🚀 소형 라벨 생성 시작', type='primary', key='s_btn'):
        if not s_file:
            st.warning('⚠️ 엑셀 파일을 먼저 업로드해주세요!')
        else:
            with st.spinner('처리 중...'):
                settings={
                    'col_name':s_col_name,'col_barcode':s_col_barcode,
                    'col_material':s_col_material,'col_insert':s_col_insert,
                    'start_row':s_start_row,'origin':s_origin,'age':s_age,
                    'insert_w':244,'insert_h':157,'row_height':120,'col_width':38
                }
                output, ok, errors = process_excel(s_file, '소형', settings)
            if errors:
                for e in errors: st.error(e)
            fname = s_file.name.replace('.xlsx','_완성.xlsx')
            st.download_button('⬇️ 완성 파일 다운로드', output, file_name=fname,
                             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ── 대형 탭 ────────────────────────────────────────────
with tab2:
    st.subheader('📋 열 번호 설정')
    st.caption('A=1, B=2, C=3, D=4 ... L=12, M=13')
    c1,c2 = st.columns(2)
    with c1:
        l_col_name     = st.number_input('상품명 열',     min_value=1, max_value=50, value=4,  key='l_name')
        l_col_barcode  = st.number_input('바코드 열',     min_value=1, max_value=50, value=12, key='l_barcode')
        l_col_material = st.number_input('재질 열',       min_value=1, max_value=50, value=13, key='l_material')
    with c2:
        l_col_insert   = st.number_input('이미지 삽입 열', min_value=1, max_value=50, value=13, key='l_insert')
        l_start_row    = st.number_input('시작 행',       min_value=1, max_value=10,  value=2,  key='l_startrow')

    st.divider()
    st.subheader('✍️ 고정 문구')
    l_caution = st.text_area('취급주의', value='취급 상 주의 사항 : 화기에 주의 하세요. 의류의 경우 단독 세탁 권장, 표백제 사용 금지, 다림질 주의, 착용시 손톱이나 날카로운 곳에 긁히지 않도록 주의', height=80, key='l_caution')
    l_addr    = st.text_input('주소/전화', value='표시자 주소 및 전화번호 : (주) 폰이지 서울시 영등포구 영등포로 109, 749호 0507-1311-1108', key='l_addr')
    l_origin  = st.text_input('제조국',   value='제조국 : Made in China',  key='l_origin')
    l_age     = st.text_input('사용연령', value='사용연령 : 만14세이상',    key='l_age')

    st.divider()
    l_file = st.file_uploader('📂 엑셀 파일 업로드', type=['xlsx'], key='l_file')

    if st.button('🚀 대형 라벨 생성 시작', type='primary', key='l_btn'):
        if not l_file:
            st.warning('⚠️ 엑셀 파일을 먼저 업로드해주세요!')
        else:
            with st.spinner('처리 중...'):
                settings={
                    'col_name':l_col_name,'col_barcode':l_col_barcode,
                    'col_material':l_col_material,'col_insert':l_col_insert,
                    'start_row':l_start_row,
                    'fix_list':[l_caution,l_addr,l_origin,l_age],
                    'insert_w':298,'insert_h':208,'row_height':160,'col_width':58
                }
                output, ok, errors = process_excel(l_file, '대형', settings)
            if errors:
                for e in errors: st.error(e)
            fname = l_file.name.replace('.xlsx','_완성.xlsx')
            st.download_button('⬇️ 완성 파일 다운로드', output, file_name=fname,
                             mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# ── 출고 작업 지시서 탭 ───────────────────────────────
with tab3:
    st.subheader('📄 출고 작업 지시서 PDF 생성')
    st.caption('CSV 파일을 업로드하면 그룹별 지시서를 PDF로 생성하고 하나로 병합합니다')

    # CSV 컬럼 안내
    with st.expander('📌 CSV 컬럼 형식 안내', expanded=False):
        st.markdown("""
| 열 번호 | 내용 |
|--------|------|
| B (2번째) | 물류센터 |
| D (4번째) | 입고예정일 |
| F (6번째) | 바코드 |
| G (7번째) | 상품명 |
| H (8번째) | 수량 |
| I (9번째) | 송장번호 |
| J (10번째) | 발주일 |
| K (11번째) | 박스번호 |
| M (13번째) | 위치 |
        """)
        st.info('첫 번째 행은 헤더로 자동 스킵됩니다.')

    st.divider()

    # 그룹화 기준
    st.subheader('🗂️ 그룹화 기준 선택')
    grouping_options = {
        '물류센터': 'logisticsCenter',
        '송장번호': 'shipmentNumber',
        '박스번호': 'boxNumber',
    }
    selected_labels = st.multiselect(
        '그룹화 기준 (복수 선택 가능)',
        options=list(grouping_options.keys()),
        default=['물류센터', '송장번호'],
        key='p_grouping'
    )
    selected_keys = [grouping_options[lbl] for lbl in selected_labels]

    st.divider()

    # 다운로드 옵션
    st.subheader('⬇️ 출력 옵션')
    download_mode = st.radio(
        '다운로드 형태',
        ['📄 전체 병합 PDF (1개 파일)', '🗜️ 그룹별 ZIP (개별 PDF)'],
        key='p_dl_mode'
    )

    st.divider()

    p_file = st.file_uploader('📂 CSV 파일 업로드', type=['csv'], key='p_file')

    if p_file:
        # 파일 분석 미리보기
        raw = p_file.read()
        items = parse_csv_to_items(raw)
        p_file.seek(0)

        if not items:
            st.error('⚠️ CSV에서 데이터를 읽지 못했습니다. 컬럼 형식을 확인해주세요.')
        else:
            if selected_keys:
                grouped = group_items(items, selected_keys)

                # 그룹 미리보기 테이블
                st.markdown(f'**총 {len(grouped)}개 그룹** · {len(items)}개 품목')
                preview_rows = []
                for gk, gitems in grouped.items():
                    preview_rows.append({
                        '그룹 키': gk,
                        '품목 수': len(gitems),
                        '총 수량': sum(i['quantity'] for i in gitems),
                        '입고예정일': gitems[0].get('expectedDate','-'),
                        '입고마감일': calc_deadline(gitems[0].get('expectedDate','')),
                    })
                st.dataframe(preview_rows, use_container_width=True, hide_index=True)

                st.divider()

                if st.button('🚀 PDF 생성 시작', type='primary', key='p_btn'):
                    group_keys = list(grouped.keys())
                    total_g = len(group_keys)
                    progress = st.progress(0)
                    status   = st.empty()

                    pdf_buffers = {}
                    errors = []

                    for i, gk in enumerate(group_keys):
                        try:
                            status.text(f'📝 생성 중: {gk}  ({i+1}/{total_g})')
                            pdf_buf = create_work_order_pdf(gk, grouped[gk])
                            pdf_buffers[gk] = pdf_buf
                        except Exception as e:
                            errors.append(f'[{gk}] 실패: {e}')
                        progress.progress((i+1) / total_g)

                    if errors:
                        for e in errors:
                            st.error(e)

                    if pdf_buffers:
                        status.text(f'✅ {len(pdf_buffers)}개 PDF 생성 완료!')
                        progress.progress(1.0)
                        today = datetime.now().strftime('%Y%m%d_%H%M')

                        if '병합' in download_mode:
                            # 전체 병합 PDF
                            merged = merge_pdfs(list(pdf_buffers.values()))
                            st.download_button(
                                label=f'⬇️ 전체 병합 PDF 다운로드 ({len(pdf_buffers)}페이지)',
                                data=merged,
                                file_name=f'출고지시서_전체_{today}.pdf',
                                mime='application/pdf',
                                key='p_dl_merged'
                            )
                        else:
                            # 개별 ZIP
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                                for gk, pdf_b in pdf_buffers.items():
                                    safe_name = gk.replace('/', '_').replace('\\', '_')[:50]
                                    zf.writestr(f'{safe_name}.pdf', pdf_b.read())
                            zip_buf.seek(0)
                            st.download_button(
                                label=f'⬇️ ZIP 다운로드 ({len(pdf_buffers)}개 PDF)',
                                data=zip_buf,
                                file_name=f'출고지시서_{today}.zip',
                                mime='application/zip',
                                key='p_dl_zip'
                            )
            else:
                st.warning('⚠️ 그룹화 기준을 하나 이상 선택해주세요.')
