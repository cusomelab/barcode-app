# ╔══════════════════════════════════════════════════════╗
# ║         [쿠썸] 바코드 라벨 생성기 - Streamlit         ║
# ╚══════════════════════════════════════════════════════╝
import os, io, urllib.request
import streamlit as st
from PIL import Image, ImageDraw, ImageFont
import barcode
from barcode.writer import ImageWriter
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter

# ── 폰트 준비 ──────────────────────────────────────────
FONT_PATH = 'NanumGothicBold.ttf'
if not os.path.exists(FONT_PATH):
    urllib.request.urlretrieve(
        'https://github.com/google/fonts/raw/main/ofl/nanumgothic/NanumGothic-Bold.ttf',
        FONT_PATH
    )

# ── 공통 헬퍼 ──────────────────────────────────────────
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
def create_small(product_name, barcode_number, material,
                 fixed_origin, fixed_age):
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
# Streamlit UI
# ══════════════════════════════════════════════════════
st.set_page_config(page_title='바코드 라벨 생성기', page_icon='🏷️', layout='centered')
st.title('🏷️ 바코드 라벨 생성기')
st.caption('엑셀 파일을 업로드하면 바코드 이미지를 자동으로 삽입합니다')

tab1, tab2 = st.tabs(['📦 소형 라벨', '📋 대형 라벨 (90도 회전)'])

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
