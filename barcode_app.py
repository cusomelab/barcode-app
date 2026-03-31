# ╔══════════════════════════════════════════════════════╗
# ║         [쿠썸] 바코드 라벨 생성기 - Streamlit         ║
# ╚══════════════════════════════════════════════════════╝
import os, io, re, urllib.request, csv, zipfile
from collections import OrderedDict
from datetime import datetime, timedelta
import streamlit as st
import streamlit.components.v1 as components_v1
import time
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
import pdfplumber
import pypdfium2 as pdfium

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
    fn=ImageFont.truetype(FONT_PATH,33)
    fm=ImageFont.truetype(FONT_PATH,20)
    ff=ImageFont.truetype(FONT_PATH,16)

    y=PAD
    for line in wrap_text(product_name,fn,CANVAS_W-PAD*2,draw)[:3]:
        bb=draw.textbbox((0,0),line,font=fn)
        draw.text(((CANVAS_W-(bb[2]-bb[0]))//2,y),line,font=fn,fill='black')
        y+=bb[3]-bb[1]+4
    y+=8

    bc_img=get_barcode_img(barcode_number,write_text=False)
    fix_h=0
    for txt in fix_list:
        for ln in wrap_text(txt,ff,CANVAS_W-PAD*2,draw):
            bb=draw.textbbox((0,0),ln,font=ff); fix_h+=bb[3]-bb[1]+3
        fix_h+=6

    BAR_W=CANVAS_W-PAD*2
    BAR_H=CANVAS_H-y-30-14-fix_h-PAD
    if BAR_H<60: BAR_H=60
    img.paste(bc_img.resize((BAR_W,BAR_H),Image.LANCZOS),(PAD,y)); y+=BAR_H+6

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


def create_work_order_pdf(group_key, items, shipment_id=None, box_number=None):
    """reportlab으로 출고 작업 지시서 PDF 생성 → BytesIO 반환
    shipment_id/box_number가 있으면 상단 우측에 표시"""
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
    s_shipment= ParagraphStyle('shipment',fontName='NanumBold', fontSize=14, leading=18, textColor=colors.HexColor('#1a56db'), alignment=2)

    total_qty   = sum(i['quantity'] for i in items)
    first       = items[0]
    deadline    = calc_deadline(first.get('expectedDate',''))
    created_at  = datetime.now().strftime('%Y-%m-%d %H:%M')
    usable_w    = PAGE_W - MARGIN * 2

    story = []

    # ── 헤더: 좌측 타이틀 + 우측 쉽먼트/박스 ──────────
    if shipment_id and box_number:
        shipment_text = f'쉽먼트 {shipment_id} | 박스 {box_number}'
    elif shipment_id:
        shipment_text = f'쉽먼트 {shipment_id}'
    else:
        shipment_text = ''

    if shipment_text:
        header_data = [[
            Paragraph('출고 작업 지시서', s_title),
            Paragraph(shipment_text, s_shipment),
        ]]
        header_tbl = Table(header_data, colWidths=[usable_w * 0.5, usable_w * 0.5])
        header_tbl.setStyle(TableStyle([
            ('VALIGN', (0, 0), (-1, -1), 'BOTTOM'),
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ]))
        story.append(header_tbl)
    else:
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
st.set_page_config(page_title='로켓배송 운영 관리', page_icon='🚀', layout='centered')
st.markdown("""<style>
    .block-container { max-width: 58rem !important; }
    /* 피킹 스캔 결과 피드백 */
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
    .shipment-input {
        background: #f0f2f6; padding: 2rem; border-radius: 12px; text-align: center;
    }
</style>""", unsafe_allow_html=True)
st.title('🚀 로켓배송 운영 관리')
st.caption('엑셀 파일을 업로드하면 바코드 이미지를 자동으로 삽입합니다')

# ══════════════════════════════════════════════════════
# 피킹 검증 시스템 — 헬퍼 함수 & 설정
# ══════════════════════════════════════════════════════
PICKING_CONFIG = {
    "SERVICE_ACCOUNT_FILE": "service_account.json",
}

def _extract_sheet_id(url_or_id):
    """구글 시트 URL 또는 ID에서 스프레드시트 ID만 추출"""
    m = re.search(r'/d/([a-zA-Z0-9_-]+)', url_or_id)
    if m:
        return m.group(1)
    return url_or_id.strip()

@st.cache_resource(ttl=600)
def get_gsheet_client():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        # 1순위: Streamlit Secrets (Cloud 배포용)
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                dict(st.secrets["gcp_service_account"]), scopes=scopes
            )
        # 2순위: 로컬 JSON 파일
        else:
            creds = Credentials.from_service_account_file(
                PICKING_CONFIG["SERVICE_ACCOUNT_FILE"], scopes=scopes
            )
        return gspread.authorize(creds)
    except FileNotFoundError:
        return None
    except Exception as e:
        st.warning(f"구글 시트 연결 실패: {e}")
        return None

def pick_load_sheet_as_df(client, sheet_url, tab_name):
    try:
        import pandas as _pd
        sheet_id = _extract_sheet_id(sheet_url)
        spreadsheet = client.open_by_key(sheet_id)
        worksheet = spreadsheet.worksheet(tab_name)
        # 중복 헤더 대응: get_all_values로 읽고 직접 DataFrame 생성
        all_values = worksheet.get_all_values()
        if len(all_values) < 2:
            return _pd.DataFrame()
        headers = all_values[0]
        # 중복 컬럼명 처리: 같은 이름이면 _2, _3 붙임
        seen = {}
        unique_headers = []
        for h in headers:
            h = str(h).strip()
            if h == '':
                h = f'_unnamed_{len(unique_headers)}'
            if h in seen:
                seen[h] += 1
                unique_headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 1
                unique_headers.append(h)
        df = _pd.DataFrame(all_values[1:], columns=unique_headers)
        return df
    except Exception as e:
        st.error(f"시트 '{tab_name}' 로드 실패: {e}")
        return None

def pick_append_log(client, sheet_url, log_entry):
    try:
        sheet_id = _extract_sheet_id(sheet_url)
        spreadsheet = client.open_by_key(sheet_id)
        try:
            ws = spreadsheet.worksheet("피킹로그")
        except Exception:
            ws = spreadsheet.add_worksheet(title="피킹로그", rows=1000, cols=10)
            ws.append_row(["시간","송장번호","바코드","상품명","결과","스캔수량","필요수량","회차기호","박스번호"])
        ws.append_row(log_entry)
        return True
    except Exception:
        return False

def pick_update_sheet_inventory(client, sheet_url, tab_name, barcode, decrement=1):
    """배대지 시트의 실제 수량을 차감"""
    try:
        sheet_id = _extract_sheet_id(sheet_url)
        spreadsheet = client.open_by_key(sheet_id)
        ws = spreadsheet.worksheet(tab_name)
        all_values = ws.get_all_values()
        if len(all_values) < 2:
            return False
        headers = all_values[0]
        bc_col = None
        qty_col = None
        for i, h in enumerate(headers):
            h_strip = str(h).strip()
            if h_strip == "바코드":
                bc_col = i
            if h_strip == "수량":
                qty_col = i
        if bc_col is None or qty_col is None:
            return False
        for row_idx in range(1, len(all_values)):
            if str(all_values[row_idx][bc_col]).strip() == barcode:
                current = all_values[row_idx][qty_col]
                try:
                    current_val = int(float(current))
                except (ValueError, TypeError):
                    current_val = 0
                new_val = max(0, current_val - decrement)
                ws.update_cell(row_idx + 1, qty_col + 1, new_val)
                return True
        return False
    except Exception:
        return False

def pick_parse_box(box_str):
    import pandas as _pd
    if _pd.isna(box_str) or str(box_str).strip() == "":
        return {"기호": None, "박스": None, "수량": None, "상태": "알수없음"}
    box_str = str(box_str).strip()
    match = re.match(r"((?:국내)?부족)\((-?\d+)\)", box_str)
    if match:
        return {"기호": match.group(1), "박스": None, "수량": int(match.group(2)), "상태": "부족"}
    match = re.match(r"([●★■▲◆◇○□△▼♦♠♣♥☆※·]+)(\d+)\((\d+)\)", box_str)
    if match:
        return {"기호": match.group(1), "박스": match.group(2), "수량": int(match.group(3)), "상태": "피킹가능"}
    match = re.match(r"([●★■▲◆◇○□△▼♦♠♣♥☆※·]+)(\d+)", box_str)
    if match:
        return {"기호": match.group(1), "박스": match.group(2), "수량": None, "상태": "피킹가능"}
    match = re.match(r"(국내재고)\((\d+)\)", box_str)
    if match:
        return {"기호": match.group(1), "박스": None, "수량": int(match.group(2)), "상태": "피킹가능"}
    if box_str == "국내재고":
        return {"기호": "국내재고", "박스": None, "수량": None, "상태": "피킹가능"}
    return {"기호": box_str, "박스": None, "수량": None, "상태": "알수없음"}

def pick_clean_출고(df):
    import pandas as _pd
    df = df.copy()
    required = ["바코드", "상품명", "수량", "쉽먼트운송장번호"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(f"출고지시서 필수 컬럼 누락: {missing}")
        return None
    df["수량"] = _pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    df["쉽먼트운송장번호"] = df["쉽먼트운송장번호"].astype(str).str.replace(r"\.0$", "", regex=True)
    df["바코드"] = df["바코드"].astype(str).str.strip()
    if "박스번호" in df.columns:
        parsed = df["박스번호"].apply(pick_parse_box)
        df["회차기호"] = parsed.apply(lambda x: x["기호"])
        df["박스넘버"] = parsed.apply(lambda x: x["박스"])
        df["박스내수량"] = parsed.apply(lambda x: x["수량"])
        df["피킹상태"] = parsed.apply(lambda x: x["상태"])
    return df

def pick_clean_배대지(df):
    import pandas as _pd
    df = df.copy()
    if "바코드" not in df.columns:
        st.error("배대지 시트에 '바코드' 컬럼이 없습니다")
        return None
    df["바코드"] = df["바코드"].astype(str).str.strip()
    if "수량" in df.columns:
        df["수량"] = _pd.to_numeric(df["수량"], errors="coerce").fillna(0).astype(int)
    if "배대지주문수량" in df.columns:
        df["배대지주문수량"] = _pd.to_numeric(df["배대지주문수량"], errors="coerce").fillna(0).astype(int)
    if "박스번호" in df.columns:
        parsed = df["박스번호"].apply(pick_parse_box)
        df["회차기호"] = parsed.apply(lambda x: x["기호"])
        df["박스넘버"] = parsed.apply(lambda x: x["박스"])
    return df

def pick_init_session():
    defaults = {
        "pick_df_출고": None, "pick_df_배대지": None,
        "pick_selected_shipment": None, "pick_picking_state": {},
        "pick_inventory_state": {}, "pick_scan_log": [],
        "pick_last_scan_result": None, "pick_scan_counter": 0,
        "pick_completed_shipments": set(), "pick_shortage_items": [],
        "pick_data_loaded": False, "pick_gsheet_client": None,
        "pick_use_gsheet": False,
        "pick_sheet_url_출고": "", "pick_sheet_tab_출고": "",
        "pick_sheet_url_배대지": "", "pick_sheet_tab_배대지": "",
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

pick_init_session()

def pick_load_all_data(url_출고, tab_출고, url_배대지="", tab_배대지=""):
    import pandas as _pd
    client = get_gsheet_client()
    if client:
        st.session_state.pick_gsheet_client = client
        st.session_state.pick_use_gsheet = True
        # 출고지시서(쉽먼트시트) 로드
        if url_출고 and tab_출고:
            df_출고 = pick_load_sheet_as_df(client, url_출고, tab_출고)
            if df_출고 is not None and not df_출고.empty:
                st.session_state.pick_df_출고 = pick_clean_출고(df_출고)
                st.session_state.pick_sheet_url_출고 = url_출고
                st.session_state.pick_sheet_tab_출고 = tab_출고
        # 배대지 입고 로드
        if url_배대지 and tab_배대지:
            df_배대지 = pick_load_sheet_as_df(client, url_배대지, tab_배대지)
            if df_배대지 is not None and not df_배대지.empty:
                st.session_state.pick_df_배대지 = pick_clean_배대지(df_배대지)
                st.session_state.pick_sheet_url_배대지 = url_배대지
                st.session_state.pick_sheet_tab_배대지 = tab_배대지
        if st.session_state.pick_df_출고 is not None:
            st.session_state.pick_data_loaded = True
            return True
    return False

def pick_init_inventory():
    df = st.session_state.pick_df_배대지
    if df is None or df.empty:
        return
    inventory = {}
    for _, row in df.iterrows():
        barcode = row["바코드"]
        symbol = row.get("회차기호", "기타")
        qty = row.get("수량", 0)
        key = (symbol, barcode)
        inventory[key] = inventory.get(key, 0) + qty
    st.session_state.pick_inventory_state = inventory

def pick_init_picking(shipment_id):
    df = st.session_state.pick_df_출고
    shipment_df = df[df["쉽먼트운송장번호"] == shipment_id]
    if shipment_df.empty:
        st.error(f"쉽먼트 {shipment_id}를 찾을 수 없습니다")
        return
    picking = {}
    shortage_items = []
    for _, row in shipment_df.iterrows():
        bc = row["바코드"]
        symbol = row.get("회차기호", "")
        qty = row["수량"]
        pick_status = row.get("피킹상태", "피킹가능")
        if pick_status == "부족":
            shortage_items.append({
                "바코드": bc, "상품명": row["상품명"],
                "부족수량": abs(row.get("박스내수량", 0) or 0),
                "박스번호": row.get("박스번호", ""),
            })
            continue
        if bc in picking:
            picking[bc]["필요수량"] += qty
        else:
            inv_key = (symbol, bc)
            inv_qty = st.session_state.pick_inventory_state.get(inv_key, None)
            picking[bc] = {
                "상품명": row["상품명"], "필요수량": qty, "스캔수량": 0,
                "회차기호": symbol if symbol else "N/A",
                "박스번호": row.get("박스번호", ""), "박스넘버": row.get("박스넘버", ""),
                "박스내수량": row.get("박스내수량", None), "배대지잔여": inv_qty,
                "SKU_ID": row.get("SKU ID", ""), "물류센터": row.get("물류센터(FC)", ""),
            }
    st.session_state.pick_picking_state = picking
    st.session_state.pick_shortage_items = shortage_items
    st.session_state.pick_selected_shipment = shipment_id
    st.session_state.pick_scan_log = []
    st.session_state.pick_last_scan_result = None
    st.session_state.pick_scan_counter = 0

def pick_process_scan(barcode):
    barcode = barcode.strip()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    state = st.session_state.pick_picking_state
    inventory = st.session_state.pick_inventory_state

    if barcode not in state:
        df = st.session_state.pick_df_출고
        hint = ""
        if df is not None:
            match = df[df["바코드"] == barcode]
            if not match.empty:
                name = match["상품명"].iloc[0][:25]
                others = match["쉽먼트운송장번호"].unique()[:3]
                hint = f" → [{name}] 다른 쉽먼트에 있음: {', '.join(s[-6:] for s in others)}"
            else:
                hint = " → 출고지시서에 없는 바코드"
        result = {"status": "error", "message": "🚨 오피킹! 이 쉽먼트에 없는 바코드",
                  "detail": f"{barcode}{hint}", "barcode": barcode, "상품명": "", "시간": now}
        st.session_state.pick_scan_log.append(result)
        st.session_state.pick_last_scan_result = result
        st.session_state.pick_scan_counter += 1
        return result

    item = state[barcode]
    item["스캔수량"] += 1

    if item["스캔수량"] > item["필요수량"]:
        result = {"status": "over", "message": f"⚠️ 수량 초과! {item['상품명'][:35]}",
                  "detail": f"필요 {item['필요수량']}개인데 {item['스캔수량']}번째 스캔",
                  "barcode": barcode, "상품명": item["상품명"], "시간": now}
        st.session_state.pick_scan_log.append(result)
        st.session_state.pick_last_scan_result = result
        st.session_state.pick_scan_counter += 1
        return result

    symbol = item["회차기호"]
    inv_key = (symbol, barcode)
    shortage_warning = ""
    if inv_key in inventory:
        if inventory[inv_key] > 0:
            inventory[inv_key] -= 1
            item["배대지잔여"] = inventory[inv_key]
        else:
            shortage_warning = f" | ⚠ {symbol}회차 배대지 재고 소진!"
            item["배대지잔여"] = 0

    remaining = item["필요수량"] - item["스캔수량"]
    result = {
        "status": "ok" if not shortage_warning else "shortage",
        "message": f"✅ {item['상품명'][:35]}",
        "detail": f"스캔 {item['스캔수량']}/{item['필요수량']} (남은: {remaining}){shortage_warning}",
        "barcode": barcode, "상품명": item["상품명"], "시간": now,
    }
    st.session_state.pick_scan_log.append(result)
    st.session_state.pick_last_scan_result = result
    st.session_state.pick_scan_counter += 1

    if st.session_state.pick_use_gsheet and st.session_state.pick_gsheet_client:
        log_row = [now, st.session_state.pick_selected_shipment or "", barcode,
                   item["상품명"][:40], result["status"], item["스캔수량"],
                   item["필요수량"], item.get("회차기호",""), item.get("박스번호","")]
        log_url = st.session_state.pick_sheet_url_출고
        pick_append_log(st.session_state.pick_gsheet_client, log_url, log_row)
        # 배대지 시트 실제 수량 차감
        if result["status"] in ("ok", "shortage") and st.session_state.pick_sheet_url_배대지 and st.session_state.pick_sheet_tab_배대지:
            pick_update_sheet_inventory(
                st.session_state.pick_gsheet_client,
                st.session_state.pick_sheet_url_배대지,
                st.session_state.pick_sheet_tab_배대지,
                barcode, decrement=1
            )
    return result

def pick_get_progress():
    state = st.session_state.pick_picking_state
    if not state:
        return {"total":0,"scanned":0,"skus":0,"done_skus":0,"pct":0.0,"is_complete":False,"over":0,"shortage":0}
    total = sum(v["필요수량"] for v in state.values())
    scanned = sum(min(v["스캔수량"], v["필요수량"]) for v in state.values())
    over = sum(max(0, v["스캔수량"] - v["필요수량"]) for v in state.values())
    skus = len(state)
    done_skus = sum(1 for v in state.values() if v["스캔수량"] >= v["필요수량"])
    shortage = sum(1 for v in state.values()
                   if v.get("배대지잔여") is not None and v["배대지잔여"] == 0 and v["스캔수량"] < v["필요수량"])
    pct = scanned / total if total > 0 else 0
    return {"total":total,"scanned":scanned,"skus":skus,"done_skus":done_skus,
            "pct":pct,"is_complete":scanned>=total,"over":over,"shortage":shortage}

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(['📦 소형 라벨', '📋 대형 라벨 (90도 회전)', '📄 출고 작업 지시서 PDF', '📎 PDF 병합', '📝 발주중단 공문', '🚛 쉽먼트 통합', '🔄 쉽먼트 재출력', '📦 피킹 검증'])

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

# ══════════════════════════════════════════════════════
# 탭4: PDF 병합
# ══════════════════════════════════════════════════════
with tab4:
    st.header('📎 PDF 병합')
    st.caption('여러 PDF 파일을 하나로 합칩니다. 드래그로 순서 조정 가능')

    uploaded_pdfs = st.file_uploader(
        'PDF 파일 업로드 (여러 개 선택 가능)',
        type=['pdf'],
        accept_multiple_files=True,
        key='merge_pdfs'
    )

    if uploaded_pdfs:
        st.markdown(f'**{len(uploaded_pdfs)}개 파일 선택됨**')
        
        # 파일 순서 표시
        for i, f in enumerate(uploaded_pdfs):
            st.text(f'  {i+1}. {f.name}')

        if st.button('🔗 PDF 병합 시작', type='primary', key='merge_btn'):
            try:
                writer = PdfWriter()
                for pdf_file in uploaded_pdfs:
                    reader = PdfReader(pdf_file)
                    for page in reader.pages:
                        writer.add_page(page)
                
                out_buf = io.BytesIO()
                writer.write(out_buf)
                out_buf.seek(0)
                
                today = datetime.now().strftime('%Y%m%d_%H%M')
                st.success(f'✅ {len(uploaded_pdfs)}개 PDF 병합 완료!')
                st.download_button(
                    label='⬇️ 병합된 PDF 다운로드',
                    data=out_buf,
                    file_name=f'병합_{today}.pdf',
                    mime='application/pdf',
                    key='merge_dl'
                )
            except Exception as e:
                st.error(f'❌ 오류: {e}')

# ══════════════════════════════════════════════════════
# 탭5: PDF → 이미지
# ══════════════════════════════════════════════════════
# PDF→이미지 기능 제거됨
    st.caption('PDF 각 페이지를 PNG 이미지로 변환합니다')

    pdf_to_img_file = st.file_uploader(
        'PDF 파일 업로드',
        type=['pdf'],
        key='pdf_to_img'
    )

    col1, col2 = st.columns(2)
    with col1:
        img_format = st.selectbox('이미지 형식', ['PNG', 'JPEG'], key='img_fmt')
    with col2:
        dpi = st.selectbox('해상도 (DPI)', [72, 150, 200, 300], index=2, key='img_dpi')

    if pdf_to_img_file:
        if st.button('🖼️ 변환 시작', type='primary', key='pdf_img_btn'):
            try:
                import fitz  # PyMuPDF
                
                pdf_bytes = pdf_to_img_file.read()
                doc = fitz.open(stream=pdf_bytes, filetype='pdf')
                total_pages = len(doc)
                
                st.info(f'총 {total_pages}페이지 변환 중...')
                progress = st.progress(0)
                
                img_buffers = []
                for i, page in enumerate(doc):
                    mat = fitz.Matrix(dpi/72, dpi/72)
                    pix = page.get_pixmap(matrix=mat)
                    img_buf = io.BytesIO(pix.tobytes(output=img_format.lower()))
                    img_buffers.append((f'page_{i+1:03d}.{img_format.lower()}', img_buf.getvalue()))
                    progress.progress((i+1)/total_pages)
                
                doc.close()
                
                if len(img_buffers) == 1:
                    st.success('✅ 변환 완료!')
                    st.download_button(
                        label=f'⬇️ 이미지 다운로드',
                        data=img_buffers[0][1],
                        file_name=img_buffers[0][0],
                        mime=f'image/{img_format.lower()}',
                        key='pdf_img_dl'
                    )
                else:
                    # ZIP으로 묶기
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for fname, data in img_buffers:
                            zf.writestr(fname, data)
                    zip_buf.seek(0)
                    today = datetime.now().strftime('%Y%m%d_%H%M')
                    st.success(f'✅ {total_pages}페이지 변환 완료!')
                    st.download_button(
                        label=f'⬇️ 전체 이미지 ZIP 다운로드 ({total_pages}장)',
                        data=zip_buf,
                        file_name=f'pdf_images_{today}.zip',
                        mime='application/zip',
                        key='pdf_img_zip_dl'
                    )
                    
            except ImportError:
                st.error('❌ PyMuPDF 라이브러리가 필요합니다. requirements.txt에 pymupdf를 추가하세요.')
            except Exception as e:
                st.error(f'❌ 오류: {e}')

# ══════════════════════════════════════════════════════
# 탭6: 이미지 → PDF
# ══════════════════════════════════════════════════════
# 이미지→PDF 기능 제거됨
    st.caption('여러 이미지를 하나의 PDF로 변환합니다')

    img_files = st.file_uploader(
        '이미지 파일 업로드 (JPG, PNG, WEBP)',
        type=['jpg', 'jpeg', 'png', 'webp'],
        accept_multiple_files=True,
        key='img_to_pdf'
    )

    col1, col2 = st.columns(2)
    with col1:
        page_size = st.selectbox('페이지 크기', ['A4', 'A3', '원본 비율 유지'], key='page_size')
    with col2:
        margin_mm = st.slider('여백 (mm)', 0, 30, 10, key='img_margin')

    if img_files:
        st.markdown(f'**{len(img_files)}개 이미지 선택됨**')
        
        # 미리보기
        cols = st.columns(min(len(img_files), 4))
        for i, img_f in enumerate(img_files[:4]):
            with cols[i]:
                st.image(img_f, caption=img_f.name, use_container_width=True)
        if len(img_files) > 4:
            st.caption(f'... 외 {len(img_files)-4}개')

        if st.button('📄 PDF 변환 시작', type='primary', key='img_pdf_btn'):
            try:
                from reportlab.lib.pagesizes import A4, A3
                from reportlab.platypus import SimpleDocTemplate, Image as RLImage
                from reportlab.lib.units import mm
                
                out_buf = io.BytesIO()
                
                if page_size == 'A4':
                    ps = A4
                elif page_size == 'A3':
                    ps = A3
                else:
                    ps = None
                
                margin = margin_mm * mm
                
                if ps:
                    doc = SimpleDocTemplate(
                        out_buf,
                        pagesize=ps,
                        leftMargin=margin, rightMargin=margin,
                        topMargin=margin, bottomMargin=margin
                    )
                    w = ps[0] - 2*margin
                    h = ps[1] - 2*margin
                    
                    story = []
                    for i, img_f in enumerate(img_files):
                        img_buf = io.BytesIO(img_f.read())
                        pil_img = Image.open(img_buf)
                        iw, ih = pil_img.size
                        ratio = min(w/iw, h/ih)
                        rw, rh = iw*ratio, ih*ratio
                        img_buf.seek(0)
                        rl_img = RLImage(img_buf, width=rw, height=rh)
                        story.append(rl_img)
                        if i < len(img_files)-1:
                            from reportlab.platypus import PageBreak
                            story.append(PageBreak())
                    
                    doc.build(story)
                else:
                    # 원본 비율 유지 - 첫 이미지 크기로 페이지 설정
                    writer_pdf = PdfWriter()
                    for img_f in img_files:
                        img_buf = io.BytesIO(img_f.read())
                        pil_img = Image.open(img_buf).convert('RGB')
                        iw, ih = pil_img.size
                        single_buf = io.BytesIO()
                        
                        tmp_doc = SimpleDocTemplate(
                            single_buf,
                            pagesize=(iw, ih),
                            leftMargin=margin, rightMargin=margin,
                            topMargin=margin, bottomMargin=margin
                        )
                        pw = iw - 2*margin
                        ph = ih - 2*margin
                        img_buf.seek(0)
                        story = [RLImage(img_buf, width=pw, height=ph)]
                        tmp_doc.build(story)
                        single_buf.seek(0)
                        r = PdfReader(single_buf)
                        for page in r.pages:
                            writer_pdf.add_page(page)
                    
                    writer_pdf.write(out_buf)
                
                out_buf.seek(0)
                today = datetime.now().strftime('%Y%m%d_%H%M')
                st.success(f'✅ {len(img_files)}개 이미지 → PDF 변환 완료!')
                st.download_button(
                    label='⬇️ PDF 다운로드',
                    data=out_buf,
                    file_name=f'images_to_pdf_{today}.pdf',
                    mime='application/pdf',
                    key='img_pdf_dl'
                )
            except Exception as e:
                st.error(f'❌ 오류: {e}')

# ══════════════════════════════════════════════════════
# 탭7: 로켓배송 발주 중단 공문 작성
# ══════════════════════════════════════════════════════
with tab5:
    st.header('📝 로켓배송 발주 중단 공문')
    st.caption('쿠팡 로켓배송 상품 영구적 발주 중단 요청 공문을 자동으로 생성합니다')

    BANNED_KEYWORDS = ['공급가 협의','발주량 협의','가격 인상','단가','시즌 종료','일시적','잠정적']
    REQUIRED_KEYWORDS = ['영구적 생산 중단','영구적 취급 중단','영구적 생산중단','영구적 취급중단']

    col_left, col_right = st.columns([2, 3])

    with col_left:
        st.subheader('📋 공문 정보 입력')

        # 기본 정보
        with st.expander('🏢 업체 정보', expanded=True):
            company_name = st.text_input('업체명 *', placeholder='예: (주)마켓피아', key='gm_company')
            representative = st.text_input('대표이사 성함 *', placeholder='예: 홍길동', key='gm_rep')
            manager_name = st.text_input('담당자명', placeholder='예: 김담당', key='gm_mgr')
            manager_contact = st.text_input('담당자 연락처', placeholder='예: 010-1234-5678', key='gm_contact')

        with st.expander('📄 문서 정보', expanded=True):
            doc_number = st.text_input(
                '문서번호',
                value=f'제 {datetime.now().year}-001호',
                key='gm_docnum'
            )
            doc_date = st.date_input('문서 날짜', value=datetime.now(), key='gm_date')

        with st.expander('✍️ 발주 중단 사유', expanded=True):
            reason_type = st.radio(
                '사유 유형',
                ['생산 중단', '취급 중단', '직접 입력'],
                horizontal=True,
                key='gm_reason_type'
            )

            if reason_type == '생산 중단':
                default_reason = '당사 제조사의 영구적 생산 중단으로 인하여 해당 상품의 지속적인 공급이 불가능하게 되었습니다.'
            elif reason_type == '취급 중단':
                default_reason = '당사의 영구적 취급 중단 결정으로 인하여 해당 상품의 지속적인 공급이 불가능하게 되었습니다.'
            else:
                default_reason = ''

            reason_detail = st.text_area(
                '사유 상세 내용 *',
                value=default_reason,
                height=120,
                placeholder='반드시 "영구적 생산 중단" 또는 "영구적 취급 중단" 문구가 포함되어야 합니다.',
                key='gm_reason'
            )

            # 사유 유효성 검사
            has_banned = any(w in reason_detail for w in BANNED_KEYWORDS)
            has_required = any(w in reason_detail for w in REQUIRED_KEYWORDS)

            if reason_detail:
                if has_banned:
                    banned_found = [w for w in BANNED_KEYWORDS if w in reason_detail]
                    st.error(f'⛔ 금지 키워드 포함: {", ".join(banned_found)}')
                elif not has_required:
                    st.warning('⚠️ "영구적 생산 중단" 또는 "영구적 취급 중단" 문구가 필요합니다')
                else:
                    st.success('✅ 사유 검증 통과')

        with st.expander('🖊️ 직인 이미지 (선택)', expanded=False):
            stamp_file = st.file_uploader('직인 이미지 업로드 (PNG 권장)', type=['png','jpg','jpeg'], key='gm_stamp')
            if stamp_file:
                st.image(stamp_file, width=100)
                stamp_size = st.slider('직인 크기', 50, 200, 80, key='gm_stamp_size')
                stamp_x = st.slider('직인 좌우 위치 (%)', 0, 100, 58, key='gm_stamp_x')
                stamp_y = st.slider('직인 상하 위치 (%)', 0, 100, 50, key='gm_stamp_y')
            else:
                stamp_size = 80
                stamp_x = 58
                stamp_y = 50

        st.subheader('📊 SKU 목록')
        sku_input_type = st.radio('입력 방식', ['엑셀 파일 업로드', '직접 입력'], horizontal=True, key='gm_sku_type')

        sku_list = []

        if sku_input_type == '엑셀 파일 업로드':
            st.caption('A열: SKU ID, B열: 상품명 (1행은 헤더)')
            sku_excel = st.file_uploader('엑셀 파일 업로드', type=['xlsx','xls'], key='gm_excel')
            if sku_excel:
                try:
                    from openpyxl import load_workbook
                    wb = load_workbook(sku_excel)
                    ws = wb.active
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        if row[0] and row[1]:
                            sku_list.append({'id': str(row[0]).strip(), 'name': str(row[1]).strip()})
                    if sku_list:
                        st.success(f'✅ {len(sku_list)}개 SKU 로드됨')
                        st.dataframe(sku_list[:5], hide_index=True)
                        if len(sku_list) > 5:
                            st.caption(f'... 외 {len(sku_list)-5}개')
                    else:
                        st.error('SKU 데이터를 읽지 못했습니다')
                except Exception as e:
                    st.error(f'파일 읽기 오류: {e}')
        else:
            st.caption('SKU ID와 상품명을 입력하세요')
            num_skus = st.number_input('SKU 수량', min_value=1, max_value=50, value=3, key='gm_num_sku')
            for i in range(int(num_skus)):
                c1, c2 = st.columns([1, 2])
                with c1:
                    sid = st.text_input(f'SKU ID {i+1}', key=f'gm_sid_{i}')
                with c2:
                    sname = st.text_input(f'상품명 {i+1}', key=f'gm_sname_{i}')
                if sid and sname:
                    sku_list.append({'id': sid, 'name': sname})

    with col_right:
        st.subheader('👁️ 미리보기')

        # 공문 미리보기 HTML 생성
        def make_letter_html(company, rep, mgr, contact, docnum, date, reason, skus, stamp_data=None, stamp_sz=80, stamp_x=58, stamp_y=50):
            date_str = date.strftime('%Y년 %m월 %d일') if hasattr(date, 'strftime') else str(date)
            sku_rows = ''
            for sku in skus:
                sku_rows += f"""
                <tr>
                    <td style="border:1px solid black;padding:6px 8px;text-align:center">{sku['id']}</td>
                    <td style="border:1px solid black;padding:6px 8px">{sku['name']}</td>
                    <td style="border:1px solid black;padding:6px 8px;font-size:7.5pt">{reason}</td>
                </tr>"""
            if not sku_rows:
                sku_rows = '<tr><td colspan="3" style="border:1px solid black;padding:20px;text-align:center;color:#999">엑셀 파일을 업로드해 주세요</td></tr>'

            stamp_html = ''
            if stamp_data:
                import base64
                b64 = base64.b64encode(stamp_data).decode()
                stamp_html = f'<img src="data:image/png;base64,{b64}" style="position:absolute;left:{stamp_x}%;top:{stamp_y}%;transform:translate(-50%,-50%);width:{stamp_sz}px;height:{stamp_sz}px;object-fit:contain;mix-blend-mode:multiply;pointer-events:none"/>'

            return f"""
            <div style="background:white;padding:15px;width:100%;box-sizing:border-box;font-family:serif;font-size:9pt;line-height:1.6;color:black">
                <table style="width:100%;border-collapse:collapse;margin-bottom:8px">
                    <tr><td style="width:110px;font-weight:bold;padding:3px 0">문서번호 :</td><td style="padding:3px 0">{docnum}</td></tr>
                    <tr><td style="font-weight:bold;padding:3px 0">수&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;신 :</td><td style="font-weight:bold;padding:3px 0">쿠팡 주식회사 귀중</td></tr>
                    <tr><td style="font-weight:bold;padding:3px 0">발&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;신 :</td><td style="padding:3px 0"><b>{company or '[업체명]'}</b>&nbsp;/&nbsp;{mgr or '[담당자명]'}&nbsp;{contact or '[연락처]'}</td></tr>
                    <tr><td style="font-weight:bold;padding:3px 0">제&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;목 :</td><td style="font-weight:bold;text-decoration:underline;padding:3px 0">로켓배송 상품 영구적 발주 중단 요청의 건</td></tr>
                </table>
                <hr style="border:1.5px solid black;margin:20px 0"/>
                <p>1. 귀사의 무궁한 발전을 기원합니다.</p>
                <p>2. 당사는 아래와 같은 불가피한 사유(영구적 생산 및 취급 중단)로 인해 해당 상품들의 공급을 지속할 수 없게 되었습니다. 이에 따라 로켓배송 서비스의 안정적인 운영을 위해 해당 제품들의 발주 중단을 정중히 요청드리는 바입니다.</p>
                <div style="text-align:center;font-weight:bold;font-size:13pt;margin:20px 0;letter-spacing:4px">- 아 래 -</div>
                <table style="width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:30px">
                    <thead>
                        <tr style="background:#f3f4f6;text-align:center;font-weight:bold">
                            <th style="border:1px solid black;padding:8px;width:20%">SKU ID</th>
                            <th style="border:1px solid black;padding:8px;width:45%">SKU 명칭</th>
                            <th style="border:1px solid black;padding:8px;width:35%">발주 중단 사유</th>
                        </tr>
                    </thead>
                    <tbody>{sku_rows}</tbody>
                </table>
                <div style="position:relative;text-align:center;margin-top:40px;padding-top:20px;border-top:1px solid #e5e7eb">
                    <div style="font-size:13pt;font-weight:bold;margin-bottom:30px;letter-spacing:2px">{date_str}</div>
                    <h1 style="font-size:20pt;font-weight:bold;letter-spacing:6px;margin-bottom:30px">{company or '[업체명]'}</h1>
                    <div style="display:inline-flex;align-items:center;gap:10px">
                        <span style="font-size:14pt;font-weight:bold;letter-spacing:4px">대표이사&nbsp;&nbsp;{rep or '[성함]'}</span>
                        <span style="font-size:13pt;font-weight:bold">(인)</span>
                    </div>
                    {stamp_html}
                </div>
            </div>"""

        stamp_data = stamp_file.read() if stamp_file else None
        if stamp_data:
            stamp_file.seek(0)

        html = make_letter_html(
            company_name, representative, manager_name, manager_contact,
            doc_number, doc_date, reason_detail, sku_list,
            stamp_data, stamp_size, stamp_x, stamp_y
        )
        components_v1.html(html, height=1000, scrolling=True)

        st.divider()

        # PDF 생성 & 다운로드
        is_valid = (
            company_name and representative and sku_list
            and not has_banned and has_required
        ) if reason_detail else False

        if not is_valid:
            if not company_name:
                st.warning('업체명을 입력해주세요')
            elif not representative:
                st.warning('대표이사 성함을 입력해주세요')
            elif not sku_list:
                st.warning('SKU를 1개 이상 입력해주세요')
            elif not reason_detail:
                st.warning('발주 중단 사유를 입력해주세요')

        if st.button('📄 공문 PDF 생성', type='primary', key='gm_pdf_btn', disabled=not is_valid):
            try:
                from reportlab.lib.pagesizes import A4
                from reportlab.lib.units import mm
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
                from reportlab.lib.styles import ParagraphStyle
                from reportlab.lib import colors

                buf = io.BytesIO()
                doc = SimpleDocTemplate(
                    buf, pagesize=A4,
                    leftMargin=25*mm, rightMargin=25*mm,
                    topMargin=20*mm, bottomMargin=20*mm
                )

                styles_normal = ParagraphStyle('normal', fontName='NanumReg', fontSize=10, leading=16)
                styles_bold   = ParagraphStyle('bold',   fontName='NanumBold', fontSize=10, leading=16)
                styles_title  = ParagraphStyle('title',  fontName='NanumBold', fontSize=14, leading=20, alignment=1)
                styles_center = ParagraphStyle('center', fontName='NanumBold', fontSize=10, leading=16, alignment=1)
                styles_big    = ParagraphStyle('big',    fontName='NanumBold', fontSize=18, leading=26, alignment=1, spaceAfter=10)
                styles_small  = ParagraphStyle('small',  fontName='NanumReg', fontSize=8, leading=11)

                date_str = doc_date.strftime('%Y년 %m월 %d일')
                story = []

                # 헤더 테이블
                header_data = [
                    [Paragraph('문서번호 :', styles_bold), Paragraph(doc_number, styles_normal)],
                    [Paragraph('수      신 :', styles_bold), Paragraph('쿠팡 주식회사 귀중', styles_bold)],
                    [Paragraph('발      신 :', styles_bold), Paragraph(f'{company_name}  /  {manager_name}  {manager_contact}', styles_normal)],
                    [Paragraph('제      목 :', styles_bold), Paragraph('<u>로켓배송 상품 영구적 발주 중단 요청의 건</u>', styles_bold)],
                ]
                header_table = Table(header_data, colWidths=[35*mm, None])
                header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'TOP'), ('BOTTOMPADDING', (0,0), (-1,-1), 4)]))
                story.append(header_table)
                story.append(HRFlowable(width='100%', thickness=1.5, color=colors.black, spaceAfter=12))

                story.append(Paragraph('1. 귀사의 무궁한 발전을 기원합니다.', styles_normal))
                story.append(Spacer(1, 8))
                story.append(Paragraph('2. 당사는 아래와 같은 불가피한 사유(영구적 생산 및 취급 중단)로 인해 해당 상품들의 공급을 지속할 수 없게 되었습니다. 이에 따라 로켓배송 서비스의 안정적인 운영을 위해 해당 제품들의 발주 중단을 정중히 요청드리는 바입니다.', styles_normal))
                story.append(Spacer(1, 16))
                story.append(Paragraph('- 아 래 -', styles_center))
                story.append(Spacer(1, 12))

                # SKU 테이블
                sku_table_data = [
                    [Paragraph('SKU ID', styles_bold), Paragraph('SKU 명칭', styles_bold), Paragraph('발주 중단 사유', styles_bold)]
                ]
                for sku in sku_list:
                    sku_table_data.append([
                        Paragraph(sku['id'], styles_normal),
                        Paragraph(sku['name'], styles_normal),
                        Paragraph(reason_detail, styles_small)
                    ])

                sku_table = Table(sku_table_data, colWidths=[35*mm, 75*mm, 50*mm])
                sku_table.setStyle(TableStyle([
                    ('GRID', (0,0), (-1,-1), 1, colors.black),
                    ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                    ('ALIGN', (0,0), (0,-1), 'CENTER'),
                    ('ALIGN', (2,0), (2,-1), 'CENTER'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('TOPPADDING', (0,0), (-1,-1), 6),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                ]))
                story.append(sku_table)
                story.append(Spacer(1, 30))

                # 서명
                story.append(Paragraph(date_str, styles_center))
                story.append(Spacer(1, 20))
                story.append(Paragraph(company_name, styles_big))
                story.append(Spacer(1, 10))

                # 직인 + 대표이사 서명
                if stamp_data:
                    from reportlab.platypus import Image as RLImage, Flowable
                    sz = stamp_size * 0.8

                    class StampWithText(Flowable):
                        def __init__(self, stamp_bytes, stamp_sz, rep_name):
                            Flowable.__init__(self)
                            self.stamp_bytes = stamp_bytes
                            self.stamp_sz = stamp_sz
                            self.rep_name = rep_name
                            self.width = 160*mm
                            self.height = stamp_sz + 5

                        def draw(self):
                            canvas = self.canv
                            # 텍스트를 중앙에 배치
                            text = f'대표이사   {self.rep_name}   (인)'
                            canvas.setFont('NanumBold', 10)
                            text_w = canvas.stringWidth(text, 'NanumBold', 10)
                            text_x = self.width / 2 - text_w / 2
                            text_y = 5
                            canvas.drawString(text_x, text_y, text)
                            # 도장을 이름 왼쪽에 겹치도록 배치
                            from reportlab.lib.utils import ImageReader
                            stamp_reader = ImageReader(io.BytesIO(self.stamp_bytes))
                            stamp_x = text_x - self.stamp_sz * 0.15
                            stamp_y = text_y - self.stamp_sz * 0.35
                            canvas.drawImage(stamp_reader, stamp_x, stamp_y,
                                           width=self.stamp_sz, height=self.stamp_sz,
                                           preserveAspectRatio=True, mask='auto')

                    stamp_flow = StampWithText(stamp_data, sz, representative)
                    stamp_flow.hAlign = 'CENTER'
                    story.append(stamp_flow)
                else:
                    story.append(Paragraph(f'대표이사&nbsp;&nbsp;&nbsp;{representative}&nbsp;&nbsp;&nbsp;(인)', styles_center))

                doc.build(story)
                buf.seek(0)

                today = datetime.now().strftime('%Y%m%d')
                st.success('✅ 공문 PDF 생성 완료!')
                st.download_button(
                    label='⬇️ 공문 PDF 다운로드',
                    data=buf,
                    file_name=f'발주중단공문_{company_name}_{today}.pdf',
                    mime='application/pdf',
                    key='gm_dl'
                )
            except Exception as e:
                st.error(f'❌ PDF 생성 오류: {e}')
                import traceback
                st.code(traceback.format_exc())

# ══════════════════════════════════════════════════════
# 탭6: 쉽먼트 통합 관리
# ══════════════════════════════════════════════════════

# ── 쉽먼트 분석 헬퍼 함수들 ────────────────────────────

def _extract_manifest_info(pdf_bytes):
    """매니페스트 PDF 바이트에서 박스/송장 정보 추출"""
    pdf = pdfplumber.open(io.BytesIO(pdf_bytes))
    pages_info = []
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ''
        box_match = re.search(r'박스\s*(\d+-\d+)', text)
        invoice_match = re.search(r'송장번호\s*\n?\s*(\d{12,})', text)
        if not invoice_match:
            invoice_match = re.search(r'(4\d{11})', text)
        pages_info.append({
            'page_idx': i,
            'box_number': box_match.group(1) if box_match else None,
            'invoice_number': invoice_match.group(1) if invoice_match else None,
            'is_main_page': box_match is not None
        })
    pdf.close()
    return pages_info


def _extract_label_info(pdf_bytes):
    """라벨 PDF 바이트에서 박스/송장 정보 추출"""
    pdf = pdfplumber.open(io.BytesIO(pdf_bytes))
    pages_info = []
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ''
        box_match = re.search(r'박스\s*(\d+-\d+)', text)
        invoice_match = re.search(r'(4\d{11})', text)
        pages_info.append({
            'page_idx': i,
            'box_number': box_match.group(1) if box_match else None,
            'invoice_number': invoice_match.group(1) if invoice_match else None,
        })
    pdf.close()
    return pages_info


def _group_manifest_pages(pages_info):
    """매니페스트 박스별 그룹핑"""
    groups, current = [], None
    for info in pages_info:
        if info['is_main_page']:
            if current:
                groups.append(current)
            current = {'box_number': info['box_number'], 'invoice_number': info['invoice_number'], 'page_indices': [info['page_idx']]}
        else:
            if current:
                current['page_indices'].append(info['page_idx'])
    if current:
        groups.append(current)
    return groups


def _group_label_pages(pages_info):
    """라벨 박스별 그룹핑"""
    box_groups = OrderedDict()
    for info in pages_info:
        box = info['box_number']
        if box not in box_groups:
            box_groups[box] = {'box_number': box, 'invoice_number': info['invoice_number'], 'page_indices': []}
        box_groups[box]['page_indices'].append(info['page_idx'])
        if not box_groups[box]['invoice_number'] and info['invoice_number']:
            box_groups[box]['invoice_number'] = info['invoice_number']
    return list(box_groups.values())


def _render_labels_4up(pdf_bytes, sorted_groups, dpi=300):
    """라벨 PDF → 이미지 렌더링 → 4분할 A4"""
    pdf_doc = pdfium.PdfDocument(pdf_bytes)
    all_indices = []
    for g in sorted_groups:
        all_indices.extend(g['page_indices'])

    images = []
    for idx in all_indices:
        page = pdf_doc[idx]
        bitmap = page.render(scale=dpi / 72)
        images.append(bitmap.to_pil())
    pdf_doc.close()

    # 4-up
    a4_w, a4_h = int(8.27 * dpi), int(11.69 * dpi)
    slot_w, slot_h = a4_w // 2, a4_h // 2
    result = []

    for start in range(0, len(images), 4):
        chunk = images[start:start + 4]
        canvas = Image.new('RGB', (a4_w, a4_h), 'white')
        positions = [(0, 0), (slot_w, 0), (0, slot_h), (slot_w, slot_h)]
        for i, img in enumerate(chunk):
            ratio = img.width / img.height
            s_ratio = slot_w / slot_h
            if ratio > s_ratio:
                nw, nh = slot_w, int(slot_w / ratio)
            else:
                nh, nw = slot_h, int(slot_h * ratio)
            resized = img.resize((nw, nh), Image.LANCZOS)
            x = positions[i][0] + (slot_w - nw) // 2
            y = positions[i][1] + (slot_h - nh) // 2
            canvas.paste(resized, (x, y))
        result.append(canvas)
    return result


def _parse_csv_bytes(csv_bytes):
    """CSV 바이트 → 아이템 리스트 (pandas 기반, 구분자/인코딩 자동감지)"""
    import pandas as pd

    text = csv_bytes.decode('utf-8-sig', errors='replace')

    # 1) 구분자 자동 감지
    delimiter = ','
    try:
        dialect = csv.Sniffer().sniff(text[:4000], delimiters=',\t;')
        delimiter = dialect.delimiter
    except csv.Error:
        if text[:4000].count('\t') > text[:4000].count(','):
            delimiter = '\t'

    # 2) pandas로 읽기 (인용부호, 콤마 포함 필드 등 자동 처리)
    try:
        df = pd.read_csv(io.StringIO(text), sep=delimiter, dtype=str, keep_default_na=False)
    except Exception:
        # 파싱 실패 시 탭으로 재시도
        try:
            df = pd.read_csv(io.StringIO(text), sep='\t', dtype=str, keep_default_na=False)
        except Exception:
            return []

    if df.empty:
        return []

    # 3) 컬럼명 정규화 (공백 제거)
    df.columns = [str(c).strip() for c in df.columns]

    # 4) 컬럼 매핑 — 헤더명으로 자동 탐지
    FIELD_MAP = {
        'logisticsCenter': ['물류센터', 'FC'],
        'expectedDate':    ['입고예정일', '예정일'],
        'productBarcode':  ['바코드', '상품바코드'],
        'productName':     ['상품명', '품명', '상품이름'],
        'quantity':        ['수량'],
        'shipmentNumber':  ['쉽먼트운송장', '송장번호', '운송장번호'],
        'orderDate':       ['발주일', '주문일'],
        'boxNumber':       ['박스번호'],
        'location':        ['위치', '적재위치'],
    }

    col_map = {}  # 내부필드명 → 실제 DF 컬럼명
    for field, keywords in FIELD_MAP.items():
        for kw in keywords:
            for col in df.columns:
                if col == kw or col.startswith(kw) or kw in col:
                    col_map[field] = col
                    break
            if field in col_map:
                break

    # 핵심 컬럼 없으면 인덱스 기반 폴백
    if 'productBarcode' not in col_map or 'productName' not in col_map:
        idx_map = {
            'logisticsCenter': 1, 'expectedDate': 3,
            'productBarcode': 5, 'productName': 6, 'quantity': 7,
            'shipmentNumber': 8, 'orderDate': 9, 'boxNumber': 10,
        }
        cols = list(df.columns)
        col_map = {}
        for field, idx in idx_map.items():
            if idx < len(cols):
                col_map[field] = cols[idx]
        if 12 < len(cols):
            col_map['location'] = cols[12]

    # 5) 아이템 리스트 생성
    def safe_get(row, field):
        col = col_map.get(field)
        if col and col in row.index:
            return str(row[col]).strip()
        return ''

    items = []
    for _, row in df.iterrows():
        barcode = safe_get(row, 'productBarcode')
        shipment = safe_get(row, 'shipmentNumber')
        if not barcode and not shipment:
            continue
        try:
            qty = int(float(safe_get(row, 'quantity') or '0'))
        except (ValueError, TypeError):
            qty = 0
        items.append({
            'logisticsCenter': safe_get(row, 'logisticsCenter'),
            'expectedDate': safe_get(row, 'expectedDate'),
            'productBarcode': barcode,
            'productName': safe_get(row, 'productName'),
            'quantity': qty,
            'shipmentNumber': shipment,
            'orderDate': safe_get(row, 'orderDate'),
            'boxNumber': safe_get(row, 'boxNumber'),
            'location': safe_get(row, 'location'),
        })
    return items


with tab6:
    st.header('🚛 쉽먼트 통합 관리')
    st.caption('CSV + 매니페스트 + 라벨을 한번에 업로드하면 출고지시서 생성 + 송장번호순 정렬 + 라벨 4분할 + 전체 통합 PDF를 만듭니다')

    st.divider()

    # ── 파일 업로드 ──────────────────────────────────
    st.subheader('📂 파일 업로드')
    st.caption('CSV 1개 + 매니페스트/라벨 PDF 여러 개를 한꺼번에 드래그 & 드롭하세요')

    uploaded_files = st.file_uploader(
        '파일 선택 (CSV + PDF)',
        type=['csv', 'pdf'],
        accept_multiple_files=True,
        key='shipment_files'
    )

    if uploaded_files:
        # 파일 분류
        csv_file = None
        manifest_files = {}
        label_files = {}

        for f in uploaded_files:
            if f.name.lower().endswith('.csv'):
                csv_file = f
            elif 'manifest' in f.name.lower():
                sid = re.search(r'\((\d+)\)', f.name)
                if sid:
                    manifest_files[sid.group(1)] = f
            elif 'label' in f.name.lower():
                sid = re.search(r'\((\d+)\)', f.name)
                if sid:
                    label_files[sid.group(1)] = f

        # 짝 매칭
        pairs = []
        for sid in sorted(manifest_files.keys()):
            if sid in label_files:
                pairs.append((sid, manifest_files[sid], label_files[sid]))

        # 미리보기 표시
        st.markdown(f'**분류 결과:**')
        if csv_file:
            st.markdown(f'- CSV: `{csv_file.name}`')
        else:
            st.warning('CSV 파일이 없습니다. 출고지시서 없이 쉽먼트만 처리합니다.')

        st.markdown(f'- 쉽먼트 세트: **{len(pairs)}개**')
        for sid, m, l in pairs:
            st.markdown(f'  - `[{sid}]` {m.name} + {l.name}')

        # 매칭 안 된 파일 경고
        for sid in manifest_files:
            if sid not in label_files:
                st.warning(f'쉽먼트 {sid}: 매니페스트만 있고 라벨 없음')
        for sid in label_files:
            if sid not in manifest_files:
                st.warning(f'쉽먼트 {sid}: 라벨만 있고 매니페스트 없음')

        if not pairs:
            st.error('매칭되는 매니페스트/라벨 세트가 없습니다.')
        else:
            st.divider()

            if st.button('🚀 통합 PDF 생성 시작', type='primary', key='shipment_btn'):
                progress = st.progress(0)
                status = st.empty()
                total_steps = len(pairs) + (1 if csv_file else 0) + 1
                step = 0

                try:
                    # ===== 1. 매니페스트 분석 → 송장→(쉽먼트ID, 박스) 매핑 =====
                    status.text('📊 매니페스트 분석 중...')
                    invoice_mapping = {}
                    manifest_data = {}  # sid → (bytes, sorted_groups)

                    for sid, mf, lf in pairs:
                        m_bytes = mf.read()
                        mf.seek(0)
                        m_info = _extract_manifest_info(m_bytes)
                        m_groups = _group_manifest_pages(m_info)
                        sorted_m = sorted(m_groups, key=lambda g: g['invoice_number'] or '')
                        manifest_data[sid] = (m_bytes, sorted_m)

                        for g in sorted_m:
                            if g['invoice_number']:
                                invoice_mapping[g['invoice_number']] = (sid, g['box_number'])

                    # ===== 2. CSV → 출고지시서 PDF 생성 (송장번호별) =====
                    so_by_invoice = {}  # 송장번호 → PDF BytesIO
                    so_pdf_buf = None
                    so_pages = 0
                    if csv_file:
                        status.text('📄 출고지시서 생성 중...')
                        csv_bytes = csv_file.read()
                        items = _parse_csv_bytes(csv_bytes)

                        if items:
                            grouped = OrderedDict()
                            for item in items:
                                key = item.get('shipmentNumber', '')
                                grouped.setdefault(key, []).append(item)

                            for key in grouped:
                                grouped[key].sort(key=lambda x: (
                                    [int(c) if c.isdigit() else c.lower()
                                     for c in re.split(r'(\d+)', x.get('boxNumber', ''))],
                                    x.get('productName', '')
                                ))

                            all_so_bufs = []
                            for inv_num in sorted(grouped.keys()):
                                inv_items = grouped[inv_num]
                                ship_id, box_num = invoice_mapping.get(inv_num, (None, None))
                                center = inv_items[0].get('logisticsCenter', '')
                                gk = f"{center}_{inv_num}" if center else inv_num
                                pdf_buf = create_work_order_pdf(gk, inv_items, ship_id, box_num)
                                so_by_invoice[inv_num] = pdf_buf
                                all_so_bufs.append(pdf_buf)

                            # 출고지시서 전체 병합본
                            so_writer = PdfWriter()
                            for buf in all_so_bufs:
                                buf.seek(0)
                                reader = PdfReader(buf)
                                for page in reader.pages:
                                    so_writer.add_page(page)
                            so_pdf_buf = io.BytesIO()
                            so_writer.write(so_pdf_buf)
                            so_pdf_buf.seek(0)
                            so_pages = len(PdfReader(io.BytesIO(so_pdf_buf.getvalue())).pages)

                            st.success(f'출고지시서 {len(all_so_bufs)}건 생성 완료')

                        step += 1
                        progress.progress(step / total_steps)

                    # ===== 3. 각 세트별 처리 =====
                    label_data = {}  # sid → (l_bytes, sorted_l)
                    shipment_only_writer = PdfWriter()  # 쉽먼트만 (매니페스트+라벨)

                    for sid, mf, lf in pairs:
                        status.text(f'📦 쉽먼트 {sid} 처리 중...')

                        m_bytes, sorted_m = manifest_data[sid]
                        m_reader = PdfReader(io.BytesIO(m_bytes))

                        # 매니페스트 정렬 → 쉽먼트 전용 PDF에 추가
                        for g in sorted_m:
                            for pidx in g['page_indices']:
                                shipment_only_writer.add_page(m_reader.pages[pidx])

                        # 라벨 분석
                        l_bytes = lf.read()
                        lf.seek(0)
                        l_info = _extract_label_info(l_bytes)
                        l_groups = _group_label_pages(l_info)
                        sorted_l = sorted(l_groups, key=lambda g: g['invoice_number'] or '')
                        label_data[sid] = (l_bytes, sorted_l)

                        step += 1
                        progress.progress(step / total_steps)

                    # 라벨 4분할 → 쉽먼트 전용에 추가
                    for sid, mf, lf in pairs:
                        l_bytes, sorted_l = label_data[sid]
                        four_up = _render_labels_4up(l_bytes, sorted_l)
                        for img in four_up:
                            img_buf = io.BytesIO()
                            img.save(img_buf, format='PDF', resolution=300)
                            img_buf.seek(0)
                            lp = PdfReader(img_buf)
                            shipment_only_writer.add_page(lp.pages[0])

                    shipment_only_buf = io.BytesIO()
                    shipment_only_writer.write(shipment_only_buf)
                    shipment_only_buf.seek(0)
                    shipment_only_pages = len(PdfReader(io.BytesIO(shipment_only_buf.getvalue())).pages)

                    # ===== 4. 전체 통합 =====
                    # 순서: [출고지시서→매니페스트→라벨] 송장번호순 통합
                    status.text('📎 전체 통합 PDF 생성 중...')
                    final_writer = PdfWriter()
                    total_pages = 0
                    label_total = 0

                    # 송장번호별 라벨 그룹 매핑 (sid → {invoice → [groups]})
                    label_by_invoice = {}
                    for sid, mf, lf in pairs:
                        l_bytes, sorted_l = label_data[sid]
                        inv_map = {}
                        for g in sorted_l:
                            inv = g['invoice_number'] or ''
                            inv_map.setdefault(inv, []).append(g)
                        label_by_invoice[sid] = inv_map

                    # 송장번호순 출고지시서→매니페스트→라벨 통합 배치
                    for sid, mf, lf in pairs:
                        m_bytes, sorted_m = manifest_data[sid]
                        m_reader = PdfReader(io.BytesIO(m_bytes))
                        l_bytes, sorted_l = label_data[sid]

                        for g in sorted_m:
                            inv = g['invoice_number']
                            # 출고지시서 먼저
                            if inv and inv in so_by_invoice:
                                so_buf = so_by_invoice[inv]
                                so_buf.seek(0)
                                so_reader = PdfReader(so_buf)
                                for page in so_reader.pages:
                                    final_writer.add_page(page)
                                total_pages += len(so_reader.pages)
                            # 매니페스트
                            for pidx in g['page_indices']:
                                final_writer.add_page(m_reader.pages[pidx])
                                total_pages += 1
                            # 해당 송장의 라벨 바로 뒤에 배치
                            inv_key = inv or ''
                            if inv_key in label_by_invoice.get(sid, {}):
                                inv_label_groups = label_by_invoice[sid].pop(inv_key)
                                four_up = _render_labels_4up(l_bytes, inv_label_groups)
                                for img in four_up:
                                    img_buf = io.BytesIO()
                                    img.save(img_buf, format='PDF', resolution=300)
                                    img_buf.seek(0)
                                    lp = PdfReader(img_buf)
                                    final_writer.add_page(lp.pages[0])
                                    total_pages += 1
                                    label_total += 1

                        # 매칭 안 된 나머지 라벨 처리
                        for inv_key, groups in label_by_invoice.get(sid, {}).items():
                            four_up = _render_labels_4up(l_bytes, groups)
                            for img in four_up:
                                img_buf = io.BytesIO()
                                img.save(img_buf, format='PDF', resolution=300)
                                img_buf.seek(0)
                                lp = PdfReader(img_buf)
                                final_writer.add_page(lp.pages[0])
                                total_pages += 1
                                label_total += 1

                    final_buf = io.BytesIO()
                    final_writer.write(final_buf)
                    final_buf.seek(0)

                    step += 1
                    progress.progress(1.0)
                    status.text('✅ 완료!')

                    # ===== 결과 표시 =====
                    st.divider()
                    st.subheader('📋 처리 결과')
                    st.markdown(f"""
| 구분 | 페이지 |
|------|--------|
| 출고지시서 + 매니페스트 (교차) | {total_pages - label_total}p |
| 라벨 4분할 | {label_total}p |
| **전체 합계** | **{total_pages}p** |
""")
                    st.caption('순서: [출고지시서→매니페스트→라벨] 송장번호순 통합 배치')

                    st.divider()

                    # ── 다운로드 버튼 3개 ────────────────────
                    today = datetime.now().strftime('%Y%m%d_%H%M')

                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.download_button(
                            label=f'⬇️ 전체 통합 PDF ({total_pages}p)',
                            data=final_buf,
                            file_name=f'shipment_ALL_merged_{today}.pdf',
                            mime='application/pdf',
                            key='ship_dl_all',
                            type='primary'
                        )

                    with col_b:
                        shipment_only_buf.seek(0)
                        st.download_button(
                            label=f'⬇️ 쉽먼트만 ({shipment_only_pages}p)',
                            data=shipment_only_buf,
                            file_name=f'shipment_only_{today}.pdf',
                            mime='application/pdf',
                            key='ship_dl_shipment'
                        )

                    with col_c:
                        if so_pdf_buf:
                            so_pdf_buf.seek(0)
                            st.download_button(
                                label=f'⬇️ 출고지시서만 ({so_pages}p)',
                                data=so_pdf_buf,
                                file_name=f'출고지시서_{today}.pdf',
                                mime='application/pdf',
                                key='ship_dl_so'
                            )

                except Exception as e:
                    st.error(f'❌ 오류 발생: {e}')
                    import traceback
                    st.code(traceback.format_exc())

# ── 쉽먼트 재출력 탭 ──────────────────────────────────────
with tab7:
    st.header('🔄 쉽먼트 재출력')
    st.caption('CSV(출고지시서)의 송장번호와 매니페스트/라벨의 송장번호를 매칭하여, CSV에 있는 송장만 골라 재출력합니다')

    st.divider()

    st.subheader('📂 파일 업로드')
    st.caption('CSV 1개(필수) + 매니페스트/라벨 PDF를 업로드하세요')

    reprint_files = st.file_uploader(
        '파일 선택 (CSV + PDF)',
        type=['csv', 'pdf'],
        accept_multiple_files=True,
        key='reprint_files'
    )

    if reprint_files:
        rp_csv = None
        rp_manifests = {}
        rp_labels = {}

        for f in reprint_files:
            if f.name.lower().endswith('.csv'):
                rp_csv = f
            elif 'manifest' in f.name.lower():
                sid = re.search(r'\((\d+)\)', f.name)
                if sid:
                    rp_manifests[sid.group(1)] = f
            elif 'label' in f.name.lower():
                sid = re.search(r'\((\d+)\)', f.name)
                if sid:
                    rp_labels[sid.group(1)] = f

        rp_pairs = []
        for sid in sorted(rp_manifests.keys()):
            if sid in rp_labels:
                rp_pairs.append((sid, rp_manifests[sid], rp_labels[sid]))

        st.markdown(f'**분류 결과:**')
        if rp_csv:
            st.markdown(f'- CSV: `{rp_csv.name}`')
        else:
            st.error('⚠️ CSV 파일은 필수입니다. 송장번호 매칭에 사용됩니다.')

        st.markdown(f'- 쉽먼트 세트: **{len(rp_pairs)}개**')
        for sid, m, l in rp_pairs:
            st.markdown(f'  - `[{sid}]` {m.name} + {l.name}')

        for sid in rp_manifests:
            if sid not in rp_labels:
                st.warning(f'쉽먼트 {sid}: 매니페스트만 있고 라벨 없음')
        for sid in rp_labels:
            if sid not in rp_manifests:
                st.warning(f'쉽먼트 {sid}: 라벨만 있고 매니페스트 없음')

        if not rp_csv:
            st.stop()

        if not rp_pairs:
            st.error('매칭되는 매니페스트/라벨 세트가 없습니다.')
        else:
            st.divider()

            if st.button('🔄 쉽먼트 재출력 시작', type='primary', key='reprint_btn'):
                rp_progress = st.progress(0)
                rp_status = st.empty()

                try:
                    # ===== 1. CSV 파싱 → 송장번호 목록 추출 =====
                    rp_status.text('📄 CSV 분석 중...')
                    csv_bytes = rp_csv.read()
                    rp_items = _parse_csv_bytes(csv_bytes)

                    if not rp_items:
                        st.error('CSV에서 항목을 찾을 수 없습니다.')
                        st.stop()

                    # 파싱 결과 확인용
                    with st.expander(f'🔍 CSV 파싱 결과 확인 ({len(rp_items)}건)', expanded=False):
                        import pandas as _rpd
                        preview = [{
                            '바코드': it['productBarcode'],
                            '상품명': it['productName'],
                            '수량': it['quantity'],
                            '송장번호': it['shipmentNumber'][-6:] if it['shipmentNumber'] else '',
                            '박스': it['boxNumber'],
                        } for it in rp_items[:20]]
                        st.dataframe(_rpd.DataFrame(preview), use_container_width=True, hide_index=True)

                    # CSV의 송장번호(shipmentNumber) 목록
                    csv_invoices = set()
                    rp_grouped = OrderedDict()
                    for item in rp_items:
                        inv = item.get('shipmentNumber', '')
                        if inv:
                            csv_invoices.add(inv)
                            rp_grouped.setdefault(inv, []).append(item)

                    for key in rp_grouped:
                        rp_grouped[key].sort(key=lambda x: (
                            [int(c) if c.isdigit() else c.lower()
                             for c in re.split(r'(\d+)', x.get('boxNumber', ''))],
                            x.get('productName', '')
                        ))

                    st.info(f'CSV 송장번호: **{len(csv_invoices)}건** 감지')
                    rp_progress.progress(0.2)

                    # ===== 2. 매니페스트/라벨 분석 =====
                    rp_status.text('📊 매니페스트/라벨 분석 중...')
                    rp_manifest_data = {}
                    rp_label_data = {}
                    rp_invoice_mapping = {}

                    for sid, mf, lf in rp_pairs:
                        m_bytes = mf.read(); mf.seek(0)
                        m_info = _extract_manifest_info(m_bytes)
                        m_groups = _group_manifest_pages(m_info)
                        sorted_m = sorted(m_groups, key=lambda g: g['invoice_number'] or '')
                        rp_manifest_data[sid] = (m_bytes, sorted_m)

                        for g in sorted_m:
                            if g['invoice_number']:
                                rp_invoice_mapping[g['invoice_number']] = (sid, g['box_number'])

                        l_bytes = lf.read(); lf.seek(0)
                        l_info = _extract_label_info(l_bytes)
                        l_groups = _group_label_pages(l_info)
                        sorted_l = sorted(l_groups, key=lambda g: g['invoice_number'] or '')
                        rp_label_data[sid] = (l_bytes, sorted_l)

                    rp_progress.progress(0.4)

                    # ===== 3. 매칭 결과 확인 =====
                    all_manifest_invoices = set(rp_invoice_mapping.keys())
                    matched = csv_invoices & all_manifest_invoices
                    not_in_manifest = csv_invoices - all_manifest_invoices
                    not_in_csv = all_manifest_invoices - csv_invoices

                    if not matched:
                        st.error('CSV와 매니페스트 간 매칭되는 송장번호가 없습니다.')
                        st.stop()

                    col_m1, col_m2, col_m3 = st.columns(3)
                    with col_m1:
                        st.metric('매칭됨', f'{len(matched)}건')
                    with col_m2:
                        st.metric('CSV에만 존재', f'{len(not_in_manifest)}건')
                    with col_m3:
                        st.metric('쉽먼트에만 존재', f'{len(not_in_csv)}건')

                    if not_in_manifest:
                        with st.expander(f'CSV에만 있는 송장 ({len(not_in_manifest)}건) - 매니페스트 없음'):
                            st.code('\n'.join(sorted(not_in_manifest)))
                    if not_in_csv:
                        with st.expander(f'쉽먼트에만 있는 송장 ({len(not_in_csv)}건) - 이번에 미출력'):
                            st.code('\n'.join(sorted(not_in_csv)))

                    # ===== 4. 출고지시서 생성 (매칭된 송장만) =====
                    rp_status.text('📄 출고지시서 생성 중...')
                    rp_so_by_invoice = {}
                    all_so_bufs = []

                    for inv_num in sorted(matched):
                        inv_items = rp_grouped[inv_num]
                        ship_id, box_num = rp_invoice_mapping.get(inv_num, (None, None))
                        center = inv_items[0].get('logisticsCenter', '')
                        gk = f"{center}_{inv_num}" if center else inv_num
                        pdf_buf = create_work_order_pdf(gk, inv_items, ship_id, box_num)
                        rp_so_by_invoice[inv_num] = pdf_buf
                        all_so_bufs.append(pdf_buf)

                    rp_progress.progress(0.6)

                    # ===== 5. 매칭된 송장만 필터링하여 통합 PDF 생성 =====
                    rp_status.text('📎 통합 PDF 생성 중...')
                    rp_final_writer = PdfWriter()
                    rp_total = 0
                    rp_label_total = 0

                    # 송장별 라벨 그룹 매핑
                    rp_label_by_inv = {}
                    for sid, mf, lf in rp_pairs:
                        l_bytes, sorted_l = rp_label_data[sid]
                        inv_map = {}
                        for g in sorted_l:
                            inv = g['invoice_number'] or ''
                            if inv in matched:
                                inv_map.setdefault(inv, []).append(g)
                        rp_label_by_inv[sid] = inv_map

                    for sid, mf, lf in rp_pairs:
                        m_bytes, sorted_m = rp_manifest_data[sid]
                        m_reader = PdfReader(io.BytesIO(m_bytes))
                        l_bytes, sorted_l = rp_label_data[sid]

                        for g in sorted_m:
                            inv = g['invoice_number']
                            if inv not in matched:
                                continue

                            # 출고지시서
                            if inv in rp_so_by_invoice:
                                so_buf = rp_so_by_invoice[inv]
                                so_buf.seek(0)
                                so_reader = PdfReader(so_buf)
                                for page in so_reader.pages:
                                    rp_final_writer.add_page(page)
                                rp_total += len(so_reader.pages)

                            # 매니페스트
                            for pidx in g['page_indices']:
                                rp_final_writer.add_page(m_reader.pages[pidx])
                                rp_total += 1

                            # 라벨
                            inv_key = inv or ''
                            if inv_key in rp_label_by_inv.get(sid, {}):
                                inv_label_groups = rp_label_by_inv[sid].pop(inv_key)
                                four_up = _render_labels_4up(l_bytes, inv_label_groups)
                                for img in four_up:
                                    img_buf = io.BytesIO()
                                    img.save(img_buf, format='PDF', resolution=300)
                                    img_buf.seek(0)
                                    lp = PdfReader(img_buf)
                                    rp_final_writer.add_page(lp.pages[0])
                                    rp_total += 1
                                    rp_label_total += 1

                    rp_progress.progress(0.9)

                    rp_final_buf = io.BytesIO()
                    rp_final_writer.write(rp_final_buf)
                    rp_final_buf.seek(0)

                    rp_progress.progress(1.0)
                    rp_status.text('✅ 완료!')

                    # ===== 결과 표시 =====
                    st.divider()
                    st.subheader('📋 재출력 결과')
                    st.markdown(f"""
| 구분 | 수량 |
|------|------|
| 매칭된 송장 | {len(matched)}건 |
| 출고지시서 + 매니페스트 | {rp_total - rp_label_total}p |
| 라벨 4분할 | {rp_label_total}p |
| **전체 합계** | **{rp_total}p** |
""")
                    st.caption('순서: [출고지시서→매니페스트→라벨] 매칭 송장번호순 통합 배치')

                    st.divider()

                    today = datetime.now().strftime('%Y%m%d_%H%M')
                    st.download_button(
                        label=f'⬇️ 재출력 통합 PDF ({rp_total}p)',
                        data=rp_final_buf,
                        file_name=f'shipment_reprint_{today}.pdf',
                        mime='application/pdf',
                        key='reprint_dl',
                        type='primary'
                    )

                except Exception as e:
                    st.error(f'❌ 오류 발생: {e}')
                    import traceback
                    st.code(traceback.format_exc())

# ══════════════════════════════════════════════════════
# 탭8: 피킹 검증 시스템
# ══════════════════════════════════════════════════════
with tab8:
    import pandas as _pd

    st.header('📦 피킹 검증 시스템')
    st.caption('바코드 스캔 → 출고지시서 검증 + 배대지 재고 동시 차감')

    # ── 데이터 소스 선택 ──
    pick_mode = st.radio(
        "데이터 소스",
        ["📊 구글 시트 (실시간)", "📂 CSV 파일 업로드"],
        index=0, key="pick_mode", horizontal=True,
    )

    if pick_mode == "📊 구글 시트 (실시간)":
        st.markdown("##### 쉽먼트 시트 (출고지시서)")
        gs_col1, gs_col2 = st.columns([3, 1])
        with gs_col1:
            url_출고 = st.text_input("구글 시트 URL", placeholder="https://docs.google.com/spreadsheets/d/...", key="pick_url_출고")
        with gs_col2:
            tab_출고 = st.text_input("탭 이름", value="쉽먼트시트", key="pick_tab_출고")

        st.markdown("##### 배대지 입고 시트 (선택)")
        gs_col3, gs_col4 = st.columns([3, 1])
        with gs_col3:
            url_배대지 = st.text_input("구글 시트 URL", placeholder="비워두면 같은 시트 사용", key="pick_url_배대지")
        with gs_col4:
            tab_배대지 = st.text_input("탭 이름", value="배대지입고리스트", key="pick_tab_배대지")

        # 배대지 URL 비어있으면 출고지시서 URL 사용
        if not url_배대지.strip() and url_출고.strip():
            url_배대지 = url_출고

        if st.button("🔄 구글 시트 연결", use_container_width=True, key="pick_gsheet_btn"):
            if not url_출고.strip():
                st.error("쉽먼트 시트 URL을 입력해주세요")
            else:
                with st.spinner("구글 시트 연결 중..."):
                    success = pick_load_all_data(url_출고, tab_출고, url_배대지, tab_배대지)
                    if success:
                        pick_init_inventory()
                        st.success("✅ 구글 시트 연결 완료!")
                        st.rerun()
                    else:
                        st.error("연결 실패 — URL과 탭 이름을 확인하세요")
    else:
        pc1, pc2 = st.columns(2)
        with pc1:
            pick_csv_출고 = st.file_uploader("출고지시서 CSV", type=["csv"], key="pick_csv_출고")
        with pc2:
            pick_csv_배대지 = st.file_uploader("배대지 입고 CSV (선택)", type=["csv"], key="pick_csv_배대지")
        if pick_csv_출고:
            df = _pd.read_csv(pick_csv_출고, encoding="utf-8-sig")
            st.session_state.pick_df_출고 = pick_clean_출고(df)
            if st.session_state.pick_df_출고 is not None:
                st.session_state.pick_data_loaded = True
        if pick_csv_배대지:
            df = _pd.read_csv(pick_csv_배대지, encoding="utf-8-sig")
            st.session_state.pick_df_배대지 = pick_clean_배대지(df)
            if st.session_state.pick_data_loaded and st.session_state.pick_df_배대지 is not None:
                pick_init_inventory()

    # ── 데이터 상태 표시 ──
    if st.session_state.pick_df_출고 is not None:
        n_rows = len(st.session_state.pick_df_출고)
        n_ship = st.session_state.pick_df_출고["쉽먼트운송장번호"].nunique()
        st.success(f"출고지시서: {n_rows}행 / {n_ship}개 쉽먼트")
    if st.session_state.pick_df_배대지 is not None:
        st.success(f"배대지 입고: {len(st.session_state.pick_df_배대지)}행 로드됨")

    # ── 데이터 없으면 가이드 ──
    if not st.session_state.pick_data_loaded:
        st.info("위에서 데이터를 연결하세요 (구글 시트 또는 CSV)")
        st.markdown("""
**구글 시트 모드:**
1. 구글 시트 URL을 위 입력칸에 붙여넣기
2. 탭 이름을 정확히 입력 (예: 쉽먼트시트, 배대지입고리스트)
3. '구글 시트 연결' 클릭

**CSV 모드:**
1. 'CSV 파일 업로드' 선택
2. 출고지시서 CSV 업로드 (필수)
3. 배대지 입고 CSV 업로드 (선택)
        """)
    elif not st.session_state.pick_selected_shipment:
        # ── 송장번호 선택 ──
        st.markdown('<div class="shipment-input">', unsafe_allow_html=True)
        st.markdown("### 📋 쉽먼트 선택")

        p_col1, p_col2 = st.columns([2, 1])
        with p_col1:
            input_shipment = st.text_input("송장번호 직접 입력", placeholder="운송장번호 입력 후 Enter", key="pick_shipment_input")
        with p_col2:
            pick_df = st.session_state.pick_df_출고
            centers = ["전체"] + sorted(pick_df["물류센터(FC)"].unique().tolist()) if "물류센터(FC)" in pick_df.columns else ["전체"]
            center = st.selectbox("물류센터", centers, key="pick_center_filter")

        pick_df = st.session_state.pick_df_출고
        if center != "전체" and "물류센터(FC)" in pick_df.columns:
            filtered = pick_df[pick_df["물류센터(FC)"] == center]
        else:
            filtered = pick_df

        summary = filtered.groupby("쉽먼트운송장번호").agg(
            SKU수=("바코드", "nunique"), 총수량=("수량", "sum"),
        ).reset_index().sort_values("총수량", ascending=False)

        selected_shipment = st.selectbox(
            "또는 목록에서 선택",
            options=summary["쉽먼트운송장번호"].tolist(),
            format_func=lambda x: (
                f"{'✅ ' if x in st.session_state.pick_completed_shipments else ''}"
                f"{x[-6:]} | "
                f"{summary[summary['쉽먼트운송장번호']==x]['SKU수'].values[0]}종 "
                f"{summary[summary['쉽먼트운송장번호']==x]['총수량'].values[0]}개"
            ),
            key="pick_shipment_select",
        )

        target = input_shipment.strip() if input_shipment else selected_shipment

        if st.button("🚀 피킹 시작", type="primary", use_container_width=True, key="pick_start_btn"):
            if target:
                valid_ids = pick_df["쉽먼트운송장번호"].unique()
                if target in valid_ids:
                    pick_init_picking(target)
                    st.rerun()
                else:
                    matches = [s for s in valid_ids if s.endswith(target)]
                    if len(matches) == 1:
                        pick_init_picking(matches[0])
                        st.rerun()
                    elif len(matches) > 1:
                        st.warning(f"'{target}'에 매칭되는 쉽먼트가 {len(matches)}개입니다.")
                    else:
                        st.error(f"'{target}'에 해당하는 쉽먼트를 찾을 수 없습니다.")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        # ── 피킹 진행 화면 ──
        progress = pick_get_progress()
        shipment_id = st.session_state.pick_selected_shipment

        hcol1, hcol2 = st.columns([4, 1])
        with hcol1:
            item0 = list(st.session_state.pick_picking_state.values())[0] if st.session_state.pick_picking_state else {}
            st.markdown(f"**쉽먼트:** `{shipment_id}` | **센터:** {item0.get('물류센터','')} | **회차:** {item0.get('회차기호','')}")
        with hcol2:
            if st.button("🔄 다른 쉽먼트", use_container_width=True, key="pick_change_btn"):
                st.session_state.pick_selected_shipment = None
                st.rerun()

        pc1, pc2, pc3, pc4, pc5 = st.columns(5)
        pc1.metric("스캔", f"{progress['scanned']}/{progress['total']}")
        pc2.metric("SKU 완료", f"{progress['done_skus']}/{progress['skus']}")
        pc3.metric("진행률", f"{progress['pct']:.0%}")
        pc4.metric("초과 스캔", f"{progress['over']}건",
                   delta=f"+{progress['over']}" if progress['over'] > 0 else None, delta_color="inverse")
        pc5.metric("재고 부족", f"{progress['shortage']}건",
                   delta=f"{progress['shortage']}" if progress['shortage'] > 0 else None, delta_color="inverse")
        st.progress(progress["pct"])

        if progress["is_complete"]:
            st.markdown(
                f'<div class="scan-complete"><strong style="font-size:1.3rem;">🎉 피킹 완료!</strong><br>'
                f'쉽먼트 {shipment_id[-6:]} — {progress["total"]}개 전부 검증 완료</div>',
                unsafe_allow_html=True)
            st.session_state.pick_completed_shipments.add(shipment_id)

        st.markdown("---")
        scan_key = f"pick_scan_{st.session_state.pick_scan_counter}"
        scanned = st.text_input("🔫 바코드 스캔 (스캐너 또는 직접 입력)", key=scan_key,
                                placeholder="스캐너 대기 중... 바코드를 스캔하세요")
        if scanned:
            pick_process_scan(scanned)
            st.rerun()

        r = st.session_state.pick_last_scan_result
        if r:
            css_class = {"ok":"scan-ok","over":"scan-warning","error":"scan-error","shortage":"scan-shortage"}.get(r["status"],"scan-ok")
            st.markdown(
                f'<div class="{css_class}"><strong style="font-size:1.1rem;">{r["message"]}</strong><br>{r["detail"]}</div>',
                unsafe_allow_html=True)
            # 스캔 결과 소리
            sound_js = {
                "ok": "o.frequency.value=880;g.gain.value=0.3;o.start();setTimeout(()=>g.gain.value=0,150);setTimeout(()=>o.stop(),200);",
                "error": "o.type='square';o.frequency.value=200;g.gain.value=0.5;o.start();setTimeout(()=>{o.frequency.value=150},150);setTimeout(()=>g.gain.value=0,500);setTimeout(()=>o.stop(),600);",
                "over": "o.type='sawtooth';o.frequency.value=400;g.gain.value=0.4;o.start();setTimeout(()=>{o.frequency.value=300},100);setTimeout(()=>g.gain.value=0,300);setTimeout(()=>o.stop(),400);",
                "shortage": "o.frequency.value=600;g.gain.value=0.3;o.start();setTimeout(()=>{o.frequency.value=400},100);setTimeout(()=>g.gain.value=0,250);setTimeout(()=>o.stop(),300);",
            }
            js_code = sound_js.get(r["status"], sound_js["ok"])
            from streamlit.components.v1 import html as st_html
            st_html(f"""<script>
            try{{var a=new(window.AudioContext||window.webkitAudioContext)();var o=a.createOscillator();var g=a.createGain();o.connect(g);g.connect(a.destination);{js_code}}}catch(e){{}}
            </script>""", height=0)

        st.markdown("---")
        st.subheader("📋 피킹 현황")
        rows = []
        for bc, info in st.session_state.pick_picking_state.items():
            s, n = info["스캔수량"], info["필요수량"]
            if s > n: status_txt = f"⚠️ 초과 ({s}/{n})"
            elif s >= n: status_txt = "✅ 완료"
            elif s > 0: status_txt = f"🔄 {s}/{n}"
            else: status_txt = "⬜ 대기"
            inv = info.get("배대지잔여")
            rows.append({
                "상태": status_txt, "바코드": bc,
                "상품명": info["상품명"][:35] + ("..." if len(info["상품명"]) > 35 else ""),
                "필요": n, "스캔": s, "남은": max(0, n - s),
                "회차": info.get("회차기호",""), "박스": info.get("박스번호",""),
                "배대지재고": f"{inv}" if inv is not None else "-",
            })
        pick_order = {"🔄":0,"⬜":1,"✅":2,"⚠️":3}
        rows.sort(key=lambda x: pick_order.get(x["상태"][0], 9))
        st.dataframe(_pd.DataFrame(rows), use_container_width=True, hide_index=True,
                     height=min(500, len(rows) * 38 + 40))

        shortage = st.session_state.get("pick_shortage_items", [])
        if shortage:
            with st.expander(f"⛔ 부족분 — 피킹 불가 ({len(shortage)}건)", expanded=False):
                st.caption("출고지시서에 '부족'으로 표시된 항목입니다.")
                st.dataframe(_pd.DataFrame(shortage), use_container_width=True, hide_index=True)

        if st.session_state.pick_scan_log:
            with st.expander(f"📜 스캔 로그 ({len(st.session_state.pick_scan_log)}건)"):
                log_display = []
                for entry in reversed(st.session_state.pick_scan_log[-50:]):
                    icon = {"ok":"✅","over":"⚠️","error":"🚨","shortage":"📦"}.get(entry["status"],"?")
                    log_display.append({"시간":entry["시간"],"결과":icon,"바코드":entry["barcode"],"내용":entry["message"]})
                st.dataframe(_pd.DataFrame(log_display), use_container_width=True, hide_index=True)

            st.download_button(
                "📥 스캔 로그 CSV",
                data=_pd.DataFrame(st.session_state.pick_scan_log).to_csv(index=False, encoding="utf-8-sig"),
                file_name=f"picking_log_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv", use_container_width=True, key="pick_log_dl")
