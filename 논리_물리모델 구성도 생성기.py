#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
민통선 출입관리체계 — 논리/물리 데이터 모델 구성도 생성기 v2.0
================================================================
기능:
  [논리 모델] 개념 데이터 모델 구성도 (엔티티/속성 기반)
    • 엔티티 박스 내 기본키(PK) / 외래키(FK) 속성 표시
    • 항목 표시 순서: 기본키/외래키 마커 → 속성명 → 데이터타입
    • 삼발이(Crow's Foot) 곡선 관계선
    • 교차 관계선 색상 자동 변경
    • 주제영역별 글꼴·색상·좌표 설정
    • 설정 저장 / 불러오기 (JSON)
    • 모든 탭에 수직 스크롤바 적용

  [물리 모델] ER 다이어그램 생성기 (엔티티/속성 기반)
    • FK 관계 자동 감지 + 수동 편집
    • 주제영역별 슬라이드 생성
    • 삼발이 화살표 + 직교 꺾인 선
    • 설정 저장 / 불러오기 (JSON)
    • 모든 탭에 수직 스크롤바 적용

  [UI 구조]
    • 상단 메뉴 바: [논리 모델] / [물리 모델] 전환 버튼
    • 논리 모델 전용 Notebook (탭 3개) — 스크롤 가능
    • 물리 모델 전용 Notebook (탭 1개) — 스크롤 가능
    • FK 편집 다이얼로그에도 스크롤바 적용

실행:
  python "논리_물리모델 구성도 생성기.py"

pip install python-pptx pandas openpyxl
"""
import sys, os, re, json, math, argparse
from pathlib import Path
from collections import OrderedDict, defaultdict, namedtuple
from dataclasses import dataclass, field
from datetime import datetime

try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Cm, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError as e:
    print(f"[오류] {e}\npip install python-pptx pandas openpyxl")
    sys.exit(1)

# ════════════════════════════════════════════════════════════════
# 공통 상수
# ════════════════════════════════════════════════════════════════
FONT           = "맑은 고딕"
LOGICAL_CFG    = "논리모델_설정.json"
PHYSICAL_CFG   = "물리모델_설정.json"
PHYS_FK_FILE   = "물리모델_FK관계.json"

# 논리 모델 기본 색상
CLR = dict(
    da_bg='BFBFBF', da_bd='404040',
    sa_bg='D6D6D6', sa_bd='585858',
    za_bg='EBEBEB', za_bd='787878',
    en_bg='FFFFFF', en_bd='606060',
    pk_txt='1F3864',
    fk_txt='4A4A4A',
    sep='AAAAAA',
    line='404040',
    cross='FF6600',
    title_bg='2E4057', title_txt='FFFFFF',
)

AUDIT_ATTRS = {
    '등록자ID','수정자ID','등록일시','수정일시','등록일자','수정일자',
    '처리일시','처리자ID','처리일','작성자ID','삭제여부','사용여부',
}

# 물리 모델 팔레트
PALETTE = [
    "#B0BEC5","#90A4AE","#78909C","#607D8B",
    "#BDBDBD","#9E9E9E","#BCAAA4","#A1887F",
    "#CFD8DC","#B0BEC5","#D7CCC8","#E0E0E0",
    "#546E7A","#455A64","#ECEFF1","#8D6E63",
]

# 물리 모델 기본값
DEFAULT_TBL_W    = 9.2
DEFAULT_H_GAP    = 0.6
DEFAULT_V_GAP    = 0.5
DEFAULT_RI_H_SPC = 0.15
DEFAULT_RI_V_SPC = 0.10
DEFAULT_MAX_COLS = 10
PPT_SLIDE_W = 33.87
PPT_SLIDE_H = 19.05
MARGIN_L = 0.5; MARGIN_T = 1.75; MARGIN_R = 0.4; MARGIN_B = 0.3
AVAIL_W = PPT_SLIDE_W - MARGIN_L - MARGIN_R
AVAIL_H = PPT_SLIDE_H - MARGIN_T - MARGIN_B
HDR_H = 0.85; ROW_H = 0.42

LayoutCfg = namedtuple("LayoutCfg", "tbl_w hdr_h row_h h_gap v_gap num_cols sky_y")

# ════════════════════════════════════════════════════════════════
# 공통 유틸
# ════════════════════════════════════════════════════════════════
def rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def emu(v_cm):
    return int(Cm(float(v_cm)))

_color_cache: dict = {}

def _clr(h):
    h = h.lstrip("#")
    return RGBColor(int(h[:2],16), int(h[2:4],16), int(h[4:],16))

def _subject_color(key):
    if key not in _color_cache:
        _color_cache[key] = PALETTE[len(_color_cache) % len(PALETTE)]
    return _color_cache[key]

def _set_inset(shape, l=0.12, t=0.03, r=0.08, b=0.03):
    from pptx.oxml.ns import qn
    bp = shape.text_frame._txBody.find(qn("a:bodyPr"))
    if bp is not None:
        bp.set("lIns",str(int(Cm(l)))); bp.set("tIns",str(int(Cm(t))))
        bp.set("rIns",str(int(Cm(r)))); bp.set("bIns",str(int(Cm(b))))

# ════════════════════════════════════════════════════════════════
# ① 논리 모델 — 데이터 로드
# ════════════════════════════════════════════════════════════════
def logical_load_tree(excel_path):
    """엔티티속성정의서 → OrderedDict 계층 구조"""
    df = pd.read_excel(excel_path, sheet_name='엔티티정의서', header=1)
    df = df.dropna(subset=['대주제영역','엔티티명'])
    tree = OrderedDict()
    for _, row in df.iterrows():
        da  = str(row.get('대주제영역','')).strip()
        sa  = str(row.get('상위주제영역','')).strip()
        za  = str(row.get('주제영역','')).strip()
        en  = str(row.get('엔티티명','')).strip()
        att = str(row.get('속성명','')).strip()
        pk  = str(row.get('PK여부','N')).strip()
        fk  = str(row.get('FK여부','N')).strip()
        dt  = str(row.get('데이터타입','')).strip()
        if not da or not en or att in ('','nan'):
            continue
        (tree.setdefault(da, OrderedDict())
             .setdefault(sa, OrderedDict())
             .setdefault(za, OrderedDict())
             .setdefault(en, {'pk':[], 'fk':[], 'all':[], 'dtype':{}}))
        e = tree[da][sa][za][en]
        if att not in e['all']:
            e['all'].append(att)
            e['dtype'][att] = dt
        if pk == 'Y' and att not in e['pk']:
            e['pk'].append(att)
        if fk == 'Y' and att not in e['fk']:
            e['fk'].append(att)
    return tree


def logical_extract_relations(tree, extra_rels=None):
    pk_map = defaultdict(list)
    for da,sad in tree.items():
        for sa,zad in sad.items():
            for za,end in zad.items():
                for en,att in end.items():
                    for pk in att['pk']:
                        pk_map[pk].append((da,sa,za,en))
    en_loc = {}
    for da,sad in tree.items():
        for sa,zad in sad.items():
            for za,end in zad.items():
                for en in end:
                    en_loc[en] = (da,sa,za)
    rels, seen = [], set()

    def add_rel(par_en, fk_attr, chi_en):
        key = (par_en, chi_en, fk_attr)
        if key in seen or par_en == chi_en: return
        seen.add(key)
        pda,psa,pza = en_loc.get(par_en, ('','',''))
        cda,csa,cza = en_loc.get(chi_en, ('','',''))
        rels.append(dict(
            par=par_en, par_da=pda, par_sa=psa, par_za=pza,
            chi=chi_en, chi_da=cda, chi_sa=csa, chi_za=cza,
            attr=fk_attr
        ))

    for da,sad in tree.items():
        for sa,zad in sad.items():
            for za,end in zad.items():
                for chi_en, chi_att in end.items():
                    for attr in chi_att['all']:
                        if attr in AUDIT_ATTRS: continue
                        if attr in chi_att['pk']: continue
                        for par_da,par_sa,par_za,par_en in pk_map.get(attr, []):
                            add_rel(par_en, attr, chi_en)
    if extra_rels:
        for par_en, pk_attr, chi_en, fk_attr in extra_rels:
            add_rel(par_en, fk_attr, chi_en)
    return rels

# ════════════════════════════════════════════════════════════════
# ② 논리 모델 — 레이아웃 크기 계산 (LC)
# ════════════════════════════════════════════════════════════════
class LC:
    def __init__(self, c):
        g = lambda k,d: c.get(k,d)
        self.SW       = emu(g('slide_w',    100))
        self.SH       = emu(g('slide_h',     55))
        self.MARG     = emu(g('margin',      0.8))
        self.GAP      = emu(g('gap',         0.35))
        self.EW       = emu(g('entity_w',    4.8))
        self.EH_BASE  = emu(g('entity_h',    1.3))
        self.AW       = emu(g('attr_w',  g('entity_w', 4.8)))
        self.EH_ROW   = emu(g('attr_row_h',  0.48))
        self.EH_ATTR_MIN = emu(g('attr_min_h', 0.48))
        self.PAD      = emu(g('pad',         0.55))
        self.HDR_ZA   = emu(g('hdr_za',      0.82))
        self.HDR_SA   = emu(g('hdr_sa',      0.88))
        self.HDR_DA   = emu(g('hdr_da',      0.95))
        self.TTL      = emu(g('title_h',     1.25))
        self.DA_FSZ   = float(g('da_font_sz', 13))
        self.SA_FSZ   = float(g('sa_font_sz', 11))
        self.ZA_FSZ   = float(g('za_font_sz', 10))
        self.EN_FSZ   = float(g('en_font_sz',  9))
        self.AT_FSZ   = float(g('attr_font_sz',7.5))
        self.TOFF     = emu(g('text_offset',  0))
        self.EN_ML    = emu(g('en_margin_l',  0.12))
        self.EN_MR    = emu(g('en_margin_r',  0.08))
        self.EN_MT    = emu(g('en_margin_t',  0.06))
        self.EN_MB    = emu(g('en_margin_b',  0.04))
        self.EN_GAP_H = emu(g('en_gap_h',     0.35))
        self.EN_GAP_V = emu(g('en_gap_v',     0.35))
        self.EN_PAD_L = emu(g('en_pad_l',     0.55))
        self.EN_PAD_T = emu(g('en_pad_t',     0.0))
        self.EN_Y_OFF = emu(g('en_y_offset',  0.0))
        self.PK_FSZ   = float(g('pk_font_sz',  7.5))
        self.FK_FSZ   = float(g('fk_font_sz',  7.5))
        self.PK_CLR   = g('pk_color',  '1F3864')
        self.FK_CLR   = g('fk_color',  '4A4A4A')
        self.PK_BOLD  = bool(g('pk_bold',  True))
        self.FK_BOLD  = bool(g('fk_bold',  False))
        self.SHOW_DTYPE = bool(g('show_dtype', True))


def en_name_h(L): return L.EH_BASE

def en_attr_h(attrs, L):
    n = len(attrs['pk']) + len(attrs['fk'])
    row_h = max(1, n) * L.EH_ROW
    return max(row_h, L.EH_ATTR_MIN)

def en_height(attrs, L): return en_name_h(L) + en_attr_h(attrs, L)

def za_wh(en_dict, L, cols=2):
    n = max(len(en_dict), 1)
    c = min(n, cols); r = math.ceil(n / c)
    max_eh = max(en_height(a,L) for a in en_dict.values()) if en_dict else L.EH_BASE
    col_w = max(L.EW, L.AW)
    w = c*col_w + (c-1)*L.EN_GAP_H + 2*L.EN_PAD_L
    h = L.HDR_ZA + L.EN_PAD_T + r*max_eh + max(0,r-1)*L.EN_GAP_V + 2*L.PAD
    return w, h

def sa_wh(za_dict, L):
    max_w, total_h = 0, L.HDR_SA + L.PAD
    for za,end in za_dict.items():
        zw,zh = za_wh(end, L)
        max_w = max(max_w, zw); total_h += zh + L.GAP
    total_h -= L.GAP
    return max_w + 2*L.PAD, total_h + L.PAD

def da_wh(sa_dict, L):
    total_w, max_h = L.PAD, 0
    for sa,zad in sa_dict.items():
        sw,sh = sa_wh(zad, L)
        total_w += sw + L.GAP; max_h = max(max_h, sh)
    total_w -= L.GAP
    return total_w + L.PAD, L.HDR_DA + max_h + 2*L.PAD

# ════════════════════════════════════════════════════════════════
# ③ 논리 모델 — PPTX 도형 그리기
# ════════════════════════════════════════════════════════════════
def add_rect(slide, x, y, w, h, bg, bd, text='',
             fsz=10, bold=False, fc='000000',
             bw=None, va='top', ha=PP_ALIGN.CENTER, toff=0):
    if bw is None: bw = Pt(0.75)
    s = slide.shapes.add_shape(1, int(x), int(y), int(w), int(h))
    s.fill.solid(); s.fill.fore_color.rgb = rgb(bg)
    s.line.color.rgb = rgb(bd); s.line.width = bw
    tf = s.text_frame; tf.word_wrap = True
    tf.margin_left = emu(0.1); tf.margin_right = emu(0.1)
    tf.margin_top  = emu(0.05) + int(toff); tf.margin_bottom = emu(0.05)
    tf.vertical_anchor = {'top':1,'middle':3,'bottom':4}[va]
    if text:
        p = tf.paragraphs[0]; p.alignment = ha
        r = p.add_run(); r.text = text; r.font.name = FONT
        r.font.size = Pt(fsz); r.font.bold = bold; r.font.color.rgb = rgb(fc)
    return s


def add_entity_box(slide, x, y, w, en_name, pk_list, fk_list, L, sty, toff=0, dtype_map=None):
    """
    엔티티를 2개 도형으로 분리:
      ① 상단 — 엔티티명 박스
      ② 하단 — 속성 박스 (순서: 기본키/외래키 마커, 속성명, 데이터타입)
    """
    attrs   = {'pk': pk_list, 'fk': fk_list, 'all': pk_list + fk_list}
    name_h  = en_name_h(L)
    attr_h  = en_attr_h(attrs, L)
    total_h = name_h + attr_h
    dtype_map = dtype_map or {}

    en_bg  = sty.get('en_bg',      CLR['en_bg'])
    en_bd  = sty.get('en_bd',      CLR['en_bd'])
    en_nfc = sty.get('en_name_fc', '000000')
    name_bg = sty.get('en_name_bg', 'D0D8E8')

    # ① 엔티티명 박스
    sn = slide.shapes.add_shape(1, int(x), int(y), int(w), int(name_h))
    sn.fill.solid(); sn.fill.fore_color.rgb = rgb(name_bg)
    sn.line.color.rgb = rgb(en_bd); sn.line.width = Pt(0.75)
    tfn = sn.text_frame; tfn.word_wrap = True
    tfn.margin_left = L.EN_ML; tfn.margin_right = L.EN_MR
    tfn.margin_top  = L.EN_MT + int(toff); tfn.margin_bottom = L.EN_MB
    tfn.vertical_anchor = 3
    pn = tfn.paragraphs[0]; pn.alignment = PP_ALIGN.CENTER
    rn = pn.add_run(); rn.text = en_name
    rn.font.name = FONT; rn.font.size = Pt(L.EN_FSZ)
    rn.font.bold = True; rn.font.color.rgb = rgb(en_nfc)

    # ② 속성 박스
    attr_y = y + name_h
    attr_w = L.AW
    sa2 = slide.shapes.add_shape(1, int(x), int(attr_y), int(attr_w), int(attr_h))
    sa2.fill.solid(); sa2.fill.fore_color.rgb = rgb(en_bg)
    sa2.line.color.rgb = rgb(en_bd); sa2.line.width = Pt(0.75)
    tfa = sa2.text_frame; tfa.word_wrap = True
    tfa.margin_left = L.EN_ML; tfa.margin_right = L.EN_MR
    tfa.margin_top  = L.EN_MT; tfa.margin_bottom = L.EN_MB
    tfa.vertical_anchor = 1

    # 항목 순서: 기본키(PK) 먼저, 외래키(FK) 다음
    # 각 행: [마커] 속성명  데이터타입
    all_attrs = [('pk', a) for a in pk_list] + [('fk', a) for a in fk_list]
    if all_attrs:
        first = True
        for kind, attr in all_attrs:
            if first:
                p = tfa.paragraphs[0]; first = False
            else:
                p = tfa.add_paragraph()
            p.alignment = PP_ALIGN.LEFT

            if kind == 'pk':
                marker = '▶ '
                fsz_v  = L.PK_FSZ; bold_v = L.PK_BOLD; clr_v = L.PK_CLR
            else:
                marker = '◆ '
                fsz_v  = L.FK_FSZ; bold_v = L.FK_BOLD; clr_v = L.FK_CLR

            r = p.add_run()
            dt_str = dtype_map.get(attr, '')
            if L.SHOW_DTYPE and dt_str:
                r.text = f'{marker}{attr}  {dt_str}'
            else:
                r.text = f'{marker}{attr}'
            r.font.name = FONT; r.font.size = Pt(fsz_v)
            r.font.bold = bold_v; r.font.color.rgb = rgb(clr_v)
    else:
        tfa.paragraphs[0].alignment = PP_ALIGN.LEFT

    return sn, sa2, total_h


# ── 엣지 포인트 ──
def edge_pt_entity(ex, ey, ew, total_h, name_h, tx, ty):
    cx  = ex + ew // 2
    ncy = ey + name_h // 2
    tcy = ey + total_h // 2
    dx, dy = tx - cx, ty - tcy
    hw = ew // 2
    hh_total = total_h // 2
    if abs(dx) < 1 and abs(dy) < 1:
        return cx, tcy
    if abs(dx) < 1:
        return (cx, ey) if dy < 0 else (cx, ey + total_h)
    sl = dy / dx
    if abs(sl) * hw < hh_total:
        if dx > 0:
            edge_y = int(ncy + sl * hw)
        else:
            edge_y = int(ncy - sl * hw)
        edge_y = max(ey, min(edge_y, ey + name_h))
        ex_side = ex + ew if dx > 0 else ex
        return ex_side, edge_y
    else:
        if dy > 0:
            denom = sl if abs(sl) > 1e-9 else 1e-9
            return int(cx + hh_total / denom), ey + total_h
        else:
            if abs(sl) < 0.5:
                return (ex + ew, int(ncy)) if dx > 0 else (ex, int(ncy))
            denom = sl if abs(sl) > 1e-9 else 1e-9
            return int(cx - hh_total / denom), ey


def _exit_direction(px, py, bx, by, bw, bh):
    tol = max(4, emu(0.02))
    on_left   = abs(px - bx)        < tol
    on_right  = abs(px - (bx + bw)) < tol
    on_top    = abs(py - by)        < tol
    on_bottom = abs(py - (by + bh)) < tol
    in_y = by - tol <= py <= by + bh + tol
    in_x = bx - tol <= px <= bx + bw + tol
    if (on_left or on_right) and in_y: return 'H'
    if (on_top or on_bottom) and in_x: return 'V'
    return 'unknown'


def _inside_box(px, py, bx, by, bw, bh):
    return bx < px < bx + bw and by < py < by + bh


def draw_crowfoot(slide, x1, y1, x2, y2, color='404040', box1=None, box2=None):
    lc2 = rgb(color); lw = Pt(1.25); shapes = []

    def line(ax, ay, bx2, by2, clr=None, width=None):
        try:
            c = slide.shapes.add_connector(1, int(ax), int(ay), int(bx2), int(by2))
            c.line.color.rgb = clr or lc2; c.line.width = width or lw
            shapes.append(c)
        except Exception: pass

    dx, dy = x2 - x1, y2 - y1
    length = math.sqrt(dx*dx + dy*dy)
    if length < emu(0.1): return shapes
    ux, uy = dx/length, dy/length
    px2, py2 = -uy, ux
    TL = emu(0.3); CL = emu(0.42); CS = emu(0.22); SNAP = emu(0.3)
    abs_dx, abs_dy = abs(dx), abs(dy)

    if abs_dx < SNAP or abs_dy < SNAP:
        line(x1, y1, x2, y2)
    elif box1 is not None and box2 is not None:
        bx1,by1,bw1,bh1 = box1; bx2b,by2b,bw2,bh2 = box2
        dir1 = _exit_direction(x1,y1,bx1,by1,bw1,bh1)
        dir2 = _exit_direction(x2,y2,bx2b,by2b,bw2,bh2)
        if dir1 == 'unknown': dir1 = 'H' if abs_dx >= abs_dy else 'V'
        if dir2 == 'unknown': dir2 = 'H' if abs_dx >= abs_dy else 'V'
        if dir1 == 'H' and dir2 == 'H':
            mid_x = (bx1+bw1+bx2b)//2 if x1 < x2 else (bx2b+bw2+bx1)//2
            line(x1,y1,mid_x,y1); line(mid_x,y1,mid_x,y2); line(mid_x,y2,x2,y2)
        elif dir1 == 'V' and dir2 == 'V':
            mid_y = (by1+bh1+by2b)//2 if y1 < y2 else (by2b+bh2+by1)//2
            line(x1,y1,x1,mid_y); line(x1,mid_y,x2,mid_y); line(x2,mid_y,x2,y2)
        else:
            m1x,m1y = x2,y1; m2x,m2y = x1,y2
            def corner_ok(cx,cy):
                return (not _inside_box(cx,cy,bx1,by1,bw1,bh1) and
                        not _inside_box(cx,cy,bx2b,by2b,bw2,bh2))
            prefer_h = (dir1 == 'H')
            if prefer_h:
                if corner_ok(m1x,m1y): line(x1,y1,m1x,m1y); line(m1x,m1y,x2,y2)
                else:                   line(x1,y1,m2x,m2y); line(m2x,m2y,x2,y2)
            else:
                if corner_ok(m2x,m2y): line(x1,y1,m2x,m2y); line(m2x,m2y,x2,y2)
                else:                   line(x1,y1,m1x,m1y); line(m1x,m1y,x2,y2)
    else:
        mx,my = (x2,y1) if abs_dx >= abs_dy else (x1,y2)
        line(x1,y1,mx,my); line(mx,my,x2,y2)

    line(x1+px2*TL, y1+py2*TL, x1-px2*TL, y1-py2*TL)
    bx3 = x2 - ux*CL; by3 = y2 - uy*CL
    line(bx3,by3,x2,y2); line(bx3,by3,x2-px2*CS,y2-py2*CS)
    line(bx3,by3,x2+px2*CS,y2+py2*CS)
    line(bx3+px2*TL,by3+py2*TL, bx3-px2*TL,by3-py2*TL)
    return shapes


def segs_cross(p1,p2,p3,p4):
    def c2d(a,b): return a[0]*b[1]-a[1]*b[0]
    def sub(a,b): return (a[0]-b[0],a[1]-b[1])
    r=sub(p2,p1); s=sub(p4,p3); rxs=c2d(r,s)
    if abs(rxs)<1e-9: return False
    t=c2d(sub(p3,p1),s)/rxs; u=c2d(sub(p3,p1),r)/rxs
    return 0.05<t<0.95 and 0.05<u<0.95

# ════════════════════════════════════════════════════════════════
# ④ 논리 모델 — 메인 PPTX 빌더
# ════════════════════════════════════════════════════════════════
def logical_build_pptx(tree, relations, cfg, positions=None):
    L   = LC(cfg)
    prs = Presentation()
    prs.slide_width  = Emu(L.SW); prs.slide_height = Emu(L.SH)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    en_pos = {}; line_segs = []; styles = cfg.get('styles', {})

    add_rect(slide, L.MARG, L.MARG, L.SW-2*L.MARG, L.TTL,
             CLR['title_bg'], '1A2540',
             text='민통선 출입관리체계 — 논리 데이터 모델',
             fsz=14, bold=True, fc=CLR['title_txt'], bw=Pt(2.0), va='middle')

    cur_x = L.MARG; cur_y = L.MARG + L.TTL + L.GAP

    for da, sa_dict in tree.items():
        sty = styles.get(da, {}); pos = (positions or {}).get(da, {})
        dw, dh = da_wh(sa_dict, L)
        if pos:
            cur_x = emu(pos.get('x', 1.0)); cur_y = emu(pos.get('y', 2.5))
        toff_v = emu(sty.get('text_offset', 0))

        add_rect(slide, cur_x, cur_y, dw, dh,
                 sty.get('da_bg', CLR['da_bg']), sty.get('da_bd', CLR['da_bd']),
                 text=da, fsz=float(sty.get('da_font_sz', L.DA_FSZ)),
                 bold=True, fc=sty.get('da_fc','000000'), bw=Pt(2.25), va='top', toff=toff_v)

        sa_x = cur_x + L.PAD
        for sa, za_dict in sa_dict.items():
            sw, sh = sa_wh(za_dict, L)
            sa_y   = cur_y + L.HDR_DA + L.PAD
            add_rect(slide, sa_x, sa_y, sw, sh,
                     sty.get('sa_bg', CLR['sa_bg']), sty.get('sa_bd', CLR['sa_bd']),
                     text=sa, fsz=float(sty.get('sa_font_sz', L.SA_FSZ)),
                     bold=True, fc=sty.get('sa_fc','000000'), bw=Pt(1.5), va='top', toff=toff_v)

            za_x = sa_x + L.PAD; za_y = sa_y + L.HDR_SA + L.PAD
            for za, en_dict in za_dict.items():
                zdw = sw - 2*L.PAD
                zw, zh = za_wh(en_dict, L)
                add_rect(slide, za_x, za_y, zdw, zh,
                         sty.get('za_bg', CLR['za_bg']), sty.get('za_bd', CLR['za_bd']),
                         text=za, fsz=float(sty.get('za_font_sz', L.ZA_FSZ)),
                         bold=False, fc=sty.get('za_fc','000000'), bw=Pt(1.0), va='top', toff=toff_v)

                cols = 2 if len(en_dict) > 1 else 1
                ex0  = za_x + L.EN_PAD_L
                ey0  = za_y + L.HDR_ZA + L.EN_PAD_T + L.PAD//2 + L.EN_Y_OFF

                for idx, (en, attrs) in enumerate(en_dict.items()):
                    col = idx % cols; row = idx // cols
                    ey_row = ey0
                    for r2 in range(row):
                        row_ens = [list(en_dict.values())[ri]
                                   for ri in range(r2*cols, min((r2+1)*cols, len(en_dict)))]
                        ey_row += max(en_height(a,L) for a in row_ens) + L.EN_GAP_V
                    ex_pos = ex0 + col * (L.EW + L.EN_GAP_H)
                    dtype_map = attrs.get('dtype', {})
                    _, __, actual_h = add_entity_box(
                        slide, ex_pos, ey_row, L.EW,
                        en, attrs['pk'], attrs['fk'], L, sty, toff=toff_v,
                        dtype_map=dtype_map)
                    en_pos[en] = (ex_pos, ey_row, L.EW, actual_h, en_name_h(L))

                za_y += zh + L.GAP
            sa_x += sw + L.GAP
        if not pos:
            cur_x += dw + L.GAP * 2

    line_clr  = cfg.get('line_color',  CLR['line'])
    cross_clr = cfg.get('cross_color', CLR['cross'])

    for rel in relations:
        pen, cen = rel['par'], rel['chi']
        if pen not in en_pos or cen not in en_pos: continue
        px2,py2,pw,ph,pnh = en_pos[pen]; cx2,cy2,cw,ch,cnh = en_pos[cen]
        pcx,pcy = px2+pw//2, py2+ph//2; ccx,ccy = cx2+cw//2, cy2+ch//2
        x1,y1 = edge_pt_entity(px2,py2,pw,ph,pnh,ccx,ccy)
        x2,y2 = edge_pt_entity(cx2,cy2,cw,ch,cnh,pcx,pcy)
        shps = draw_crowfoot(slide,x1,y1,x2,y2,line_clr,
                             box1=(px2,py2,pw,ph), box2=(cx2,cy2,cw,ch))
        line_segs.append(((x1,y1),(x2,y2),shps))

    n = len(line_segs); crossed = set()
    for i in range(n):
        for j in range(i+1,n):
            if segs_cross(line_segs[i][0],line_segs[i][1],
                          line_segs[j][0],line_segs[j][1]):
                crossed.add(i); crossed.add(j)
    for idx in crossed:
        _,_,shps = line_segs[idx]
        for sh in shps:
            try: sh.line.color.rgb = rgb(cross_clr)
            except Exception: pass

    leg_items = [
        ('■ 대주제영역', CLR['da_bg'], CLR['da_bd']),
        ('■ 상위주제영역',CLR['sa_bg'], CLR['sa_bd']),
        ('■ 주제영역',    CLR['za_bg'], CLR['za_bd']),
        ('■ 엔티티',      CLR['en_bg'], CLR['en_bd']),
    ]
    lx = L.MARG; ly = L.SH - L.MARG - emu(0.72)
    for lbl, bg, bd in leg_items:
        add_rect(slide, lx, ly, emu(3.3), emu(0.65), bg, bd,
                 text=lbl, fsz=8, bw=Pt(0.5), va='middle')
        lx += emu(3.5)
    return prs

# ════════════════════════════════════════════════════════════════
# ⑤ 물리 모델 — 데이터 클래스 및 파싱
# ════════════════════════════════════════════════════════════════
COL_IDX = {
    "시스템":0,"DB명":1,"스키마명":2,"주제영역":3,
    "테이블명":4,"테이블한글명":5,"순서":6,
    "컬럼명":7,"컬럼한글명":8,"PK":9,"FK":10,
    "NULL허용":11,"Datatype":12,"Default":13,
    "정의특이사항":14,"NEW컬럼명":15,
}

@dataclass
class PhysColumn:
    seq:int; name:str; name_kor:str
    is_pk:bool; is_fk:bool; nullable:bool
    datatype:str; default:str; definition:str

@dataclass
class PhysTable:
    name:str; name_kor:str; subject:str
    columns:list = field(default_factory=list)


def phys_load_excel(path):
    df = pd.read_excel(path, sheet_name=0, header=0, dtype=str).fillna("")
    return df

def phys_parse_tables(df):
    tables = {}
    def get(row, key):
        idx = COL_IDX.get(key, -1)
        if idx < 0 or idx >= len(row): return ""
        return str(row.iloc[idx]).strip()
    for _, row in df.iterrows():
        tn = get(row,"테이블명")
        if not tn or tn.lower() in ("nan","테이블명"): continue
        subj = get(row,"주제영역")
        if tn not in tables:
            tables[tn] = PhysTable(tn, get(row,"테이블한글명") or tn, subj)
        elif subj and not tables[tn].subject:
            tables[tn].subject = subj
        seq_raw = get(row,"순서")
        try: seq = int(float(seq_raw)) if seq_raw else 0
        except: seq = 0
        col = PhysColumn(
            seq=seq, name=get(row,"컬럼명"), name_kor=get(row,"컬럼한글명"),
            is_pk=get(row,"PK").upper()=="Y", is_fk=get(row,"FK").upper()=="Y",
            nullable=get(row,"NULL허용").upper()!="N",
            datatype=get(row,"Datatype"), default=get(row,"Default"),
            definition=get(row,"정의특이사항"))
        if col.name: tables[tn].columns.append(col)
    for t in tables.values(): t.columns.sort(key=lambda c: c.seq)
    return tables

REF_PAT = re.compile(r"\b(TB_\w+)\.(\w+)", re.IGNORECASE)

def phys_extract_relations(tables):
    rels = []
    for tbl in tables.values():
        for col in tbl.columns:
            if not col.is_fk: continue
            for m in REF_PAT.finditer(col.definition):
                dst = m.group(1).upper()
                if dst in tables:
                    rels.append((tbl.name, col.name, dst, m.group(2).upper()))
    return rels

def phys_inject_fk_cols(tables, all_rels):
    added = 0
    for child_tbl, fk_col, parent_tbl, pk_col in all_rels:
        if child_tbl not in tables or parent_tbl not in tables: continue
        ctbl = tables[child_tbl]; ptbl = tables[parent_tbl]
        existing = next((c for c in ctbl.columns if c.name == fk_col), None)
        if existing is not None:
            if not existing.is_fk: existing.is_fk = True
            continue
        ref_col_name = pk_col if (pk_col and pk_col != "PK") else ""
        parent_col = None
        if ref_col_name:
            parent_col = next((c for c in ptbl.columns if c.name == ref_col_name), None)
        if parent_col is None:
            parent_col = next((c for c in ptbl.columns if c.is_pk), None)
        if parent_col:
            new_col = PhysColumn(seq=0, name=fk_col, name_kor=parent_col.name_kor,
                                 is_pk=False, is_fk=True, nullable=True,
                                 datatype=parent_col.datatype, default="",
                                 definition=f"FK → {parent_tbl}.{parent_col.name}")
        else:
            new_col = PhysColumn(seq=0, name=fk_col, name_kor=fk_col,
                                 is_pk=False, is_fk=True, nullable=True,
                                 datatype="", default="", definition=f"FK → {parent_tbl}")
        insert_idx = sum(1 for c in ctbl.columns if c.is_pk)
        ctbl.columns.insert(insert_idx, new_col); added += 1
    return added

def phys_save_fk(extra_rels, path=None):
    p = Path(path or PHYS_FK_FILE)
    data = {"version":2, "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "relations":[{"child_table":ct,"fk_col":fc,"parent_table":pt,"parent_pk_col":pc}
                         for ct,fc,pt,pc in extra_rels]}
    try:
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return p
    except Exception as e:
        print(f"FK 저장 실패: {e}"); return None

def phys_load_fk(path=None):
    p = Path(path or PHYS_FK_FILE)
    if not p.exists(): return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return [(r["child_table"],r["fk_col"],r["parent_table"],r.get("parent_pk_col",""))
                for r in data.get("relations",[])]
    except: return []

def phys_detect_parent_child(relations):
    parents, children = set(), set()
    for s,_,d,__ in relations: children.add(s); parents.add(d)
    return parents, children

def phys_resolve_groups(tables):
    if sum(1 for t in tables.values() if t.subject.strip())/max(len(tables),1) >= 0.5:
        g = defaultdict(list)
        for n,t in tables.items(): g[t.subject.strip() or "기타"].append(n)
        return dict(g)
    g = defaultdict(list)
    for n in tables:
        parts = n.split("_"); g["_".join(parts[:2]) if len(parts)>=2 else n].append(n)
    return dict(g)

# ════════════════════════════════════════════════════════════════
# ⑥ 물리 모델 — 레이아웃 / 도형 / 관계선
# ════════════════════════════════════════════════════════════════
def _auto_cfg(tbl_list, tables, max_cols=DEFAULT_MAX_COLS,
              user_tbl_w=None, user_h_gap=None, user_v_gap=None):
    tbls = [t for t in tbl_list if t in tables]; n = len(tbls)
    pref_w = float(user_tbl_w) if user_tbl_w else DEFAULT_TBL_W
    h_gap  = float(user_h_gap) if user_h_gap  else DEFAULT_H_GAP
    v_gap  = float(user_v_gap) if user_v_gap  else DEFAULT_V_GAP
    if n == 0:
        return LayoutCfg(pref_w,HDR_H,ROW_H,h_gap,v_gap,
                         max(1,int((AVAIL_W+h_gap)/(pref_w+h_gap))), MARGIN_T-0.45)
    avg_cols = min(max_cols+1, sum(len(tables[t].columns) for t in tbls)/n)
    for nc in range(3,9):
        tbl_w = min(pref_w,(AVAIL_W-(nc-1)*h_gap)/nc)
        if tbl_w < 4.5: continue
        n_per = (n+nc-1)//nc
        col_h = n_per*(HDR_H+avg_cols*ROW_H+v_gap)-v_gap
        if col_h <= AVAIL_H:
            return LayoutCfg(tbl_w,HDR_H,ROW_H,h_gap,v_gap,nc,MARGIN_T-0.45)
    nc = 7; tbl_w = max(4.5,(AVAIL_W-(nc-1)*h_gap)/nc)
    n_per = (n+nc-1)//nc; avg_h = HDR_H+avg_cols*ROW_H
    scale = min(1.0,max(0.38,AVAIL_H/max(0.1,n_per*(avg_h+v_gap))))
    return LayoutCfg(min(pref_w,tbl_w),max(0.42,HDR_H*scale),max(0.22,ROW_H*scale),
                     h_gap,v_gap,nc,MARGIN_T-0.30)

def _phys_layout(tbl_list, tables, cfg, height_map=None):
    col_y = [0.0]*cfg.num_cols; positions = []
    for tn in tbl_list:
        if tn not in tables: continue
        c = col_y.index(min(col_y))
        left = MARGIN_L + c*(cfg.tbl_w+cfg.h_gap)
        top  = MARGIN_T + col_y[c]
        positions.append((tn,left,top,c))
        h = (height_map[tn] if height_map and tn in height_map
             else cfg.hdr_h + len(tables[tn].columns)*cfg.row_h)
        col_y[c] += h + cfg.v_gap
    return positions

def _tbl_display_height(tbl, cfg, max_cols):
    n = min(len(tbl.columns), max_cols)
    extra = 1 if len(tbl.columns) > max_cols else 0
    return cfg.hdr_h + (n+extra)*cfg.row_h

def _sort_by_role(tbl_list, parents, children):
    def _key(tn):
        p = tn in parents; c = tn in children
        if p and not c: return 0
        if p and c:     return 1
        if c and not p: return 2
        return 3
    return sorted(tbl_list, key=_key)

def _get_related_parents(tbl_names, all_rels, all_tables):
    on = set(tbl_names)
    return {d: all_tables[d] for s,sc,d,dc in all_rels
            if s in on and d not in on and d in all_tables}

def _get_related_children(tbl_names, all_rels, all_tables):
    on = set(tbl_names)
    return {s: all_tables[s] for s,sc,d,dc in all_rels
            if d in on and s not in on and s in all_tables}

def _phys_add_table(slide, tbl, left, top, is_parent, is_child,
                    cfg, max_cols=DEFAULT_MAX_COLS, fk_col_keys=None):
    from pptx.enum.text import PP_ALIGN as PA
    RECT = 1
    hdr_bg = "#EEEEEE"; txt_clr = _clr("#212121")
    if   is_parent and is_child: role = "  ▲▼ 상위/하위"
    elif is_parent:               role = "  ▲ 상위"
    elif is_child:                role = "  ▼ 하위"
    else:                         role = ""
    if   is_parent: bclr,bw = _clr("#424242"),Pt(2.0)
    elif is_child:  bclr,bw = _clr("#616161"),Pt(1.5)
    else:           bclr,bw = _clr("#9E9E9E"),Pt(0.75)
    pt_h = max(6,min(9,round(cfg.hdr_h*10.5)))
    pt_s = max(5,min(8,round(cfg.hdr_h*9.0)))
    pt_c = max(5,min(7,round(cfg.row_h*16.5)))
    pt_d = max(5,min(7,round(cfg.row_h*14.5)))

    hdr = slide.shapes.add_shape(RECT,Cm(left),Cm(top),Cm(cfg.tbl_w),Cm(cfg.hdr_h))
    hdr.fill.solid(); hdr.fill.fore_color.rgb = _clr(hdr_bg)
    hdr.line.color.rgb = bclr; hdr.line.width = bw
    _set_inset(hdr,0.10,0.04,0.08,0.0)
    tf = hdr.text_frame; tf.word_wrap = False
    p0 = tf.paragraphs[0]; p0.alignment = PA.CENTER
    r0 = p0.add_run(); r0.text = tbl.name
    r0.font.name=FONT; r0.font.size=Pt(pt_h); r0.font.bold=True; r0.font.color.rgb=txt_clr
    p1 = tf.add_paragraph(); p1.alignment = PA.CENTER
    r1 = p1.add_run(); r1.text = tbl.name_kor + role
    r1.font.name=FONT; r1.font.size=Pt(pt_s); r1.font.bold=bool(role)
    r1.font.color.rgb=_clr("#444444")

    col_pts = {}
    display_cols = tbl.columns[:max_cols]
    for i, col in enumerate(display_cols):
        rt = top + cfg.hdr_h + i*cfg.row_h
        force_fk = bool(fk_col_keys and (tbl.name,col.name) in fk_col_keys)
        disp_pk  = col.is_pk and not force_fk
        disp_fk  = col.is_fk or force_fk
        if   disp_pk: row_bg="#F0F0F0"
        elif disp_fk: row_bg="#F8F8F8"
        else:         row_bg="#FFFFFF"
        row = slide.shapes.add_shape(RECT,Cm(left),Cm(rt),Cm(cfg.tbl_w),Cm(cfg.row_h))
        row.fill.solid(); row.fill.fore_color.rgb=_clr(row_bg)
        row.line.color.rgb=_clr("#CCCCCC"); row.line.width=Pt(0.25)
        _set_inset(row,0.12,0.0,0.06,0.0)
        tf2=row.text_frame; tf2.word_wrap=False
        pr=tf2.paragraphs[0]; pr.alignment=PA.LEFT
        # 순서: 기본키/외래키 태그 | 속성명 | 한글명 | 데이터타입
        tag = "기본키 " if disp_pk else ("외래키 " if disp_fk else "      ")
        r_tag=pr.add_run(); r_tag.text=tag
        r_tag.font.name=FONT; r_tag.font.size=Pt(pt_d); r_tag.font.bold=True
        r_tag.font.color.rgb=(_clr("#212121") if disp_pk else
                               _clr("#555555") if disp_fk else _clr("#CCCCCC"))
        r_col=pr.add_run(); r_col.text=col.name
        r_col.font.name=FONT; r_col.font.size=Pt(pt_c)
        r_col.font.bold=disp_pk; r_col.font.color.rgb=_clr("#212121")
        r_kor=pr.add_run(); r_kor.text=f"  {col.name_kor}" if col.name_kor else ""
        r_kor.font.name=FONT; r_kor.font.size=Pt(pt_d); r_kor.font.color.rgb=_clr("#555555")
        r_dt=pr.add_run(); r_dt.text=f"  {col.datatype}" if col.datatype else ""
        r_dt.font.name=FONT; r_dt.font.size=Pt(pt_d); r_dt.font.color.rgb=_clr("#777777")
        cy = rt + cfg.row_h/2
        col_pts[col.name] = {"L":(left,cy),"R":(left+cfg.tbl_w,cy)}

    remaining = len(tbl.columns)-len(display_cols)
    if remaining > 0:
        ri = len(display_cols); rt = top+cfg.hdr_h+ri*cfg.row_h
        more=slide.shapes.add_shape(RECT,Cm(left),Cm(rt),Cm(cfg.tbl_w),Cm(cfg.row_h))
        more.fill.solid(); more.fill.fore_color.rgb=_clr("#EEEEEE")
        more.line.color.rgb=_clr("#CCCCCC"); more.line.width=Pt(0.25)
        _set_inset(more,0.14,0.0,0.06,0.0)
        tf3=more.text_frame; tf3.word_wrap=False
        pm=tf3.paragraphs[0]; pm.alignment=PA.CENTER
        rm=pm.add_run(); rm.text=f"... +{remaining}개 속성 더 있음"
        rm.font.name=FONT; rm.font.size=Pt(pt_d)
        rm.font.italic=True; rm.font.color.rgb=_clr("#888888")
    return col_pts

def _phys_add_ref_table(slide, tbl, left, top, cfg, max_cols=DEFAULT_MAX_COLS):
    from pptx.enum.text import PP_ALIGN as PA
    from pptx.oxml.ns import qn
    try: from lxml import etree
    except ImportError: etree = None
    RECT = 1
    pk_cols = [c for c in tbl.columns if c.is_pk][:max_cols]
    pt_h = max(6,min(9,round(cfg.hdr_h*10.5)))
    pt_c = max(5,round(cfg.row_h*14))
    hdr=slide.shapes.add_shape(RECT,Cm(left),Cm(top),Cm(cfg.tbl_w),Cm(cfg.hdr_h))
    hdr.fill.solid(); hdr.fill.fore_color.rgb=_clr("#E0E0E0")
    hdr.line.color.rgb=_clr("#9E9E9E"); hdr.line.width=Pt(1.0)
    if etree is not None:
        try:
            sp=hdr._element.find(qn("p:spPr"))
            if sp is not None:
                ln=sp.find(qn("a:ln"))
                if ln is not None:
                    pd2=etree.SubElement(ln,qn("a:prstDash")); pd2.set("val","dash")
        except: pass
    _set_inset(hdr,0.10,0.04,0.08,0.0)
    tf=hdr.text_frame; tf.word_wrap=False
    p0=tf.paragraphs[0]; p0.alignment=PA.CENTER
    r0=p0.add_run(); r0.text=tbl.name
    r0.font.name=FONT; r0.font.size=Pt(pt_h); r0.font.bold=True; r0.font.color.rgb=_clr("#424242")
    p1=tf.add_paragraph(); p1.alignment=PA.CENTER
    r1=p1.add_run(); r1.text=f"{tbl.name_kor}  [외부참조]"
    r1.font.name=FONT; r1.font.size=Pt(max(5,round(cfg.hdr_h*9)))
    r1.font.italic=True; r1.font.color.rgb=_clr("#757575")
    col_pts={}
    for i,col in enumerate(pk_cols):
        rt=top+cfg.hdr_h+i*cfg.row_h
        row=slide.shapes.add_shape(RECT,Cm(left),Cm(rt),Cm(cfg.tbl_w),Cm(cfg.row_h))
        row.fill.solid(); row.fill.fore_color.rgb=_clr("#F5F5F5")
        row.line.color.rgb=_clr("#CCCCCC"); row.line.width=Pt(0.25)
        _set_inset(row,0.12,0.0,0.06,0.0)
        tf2=row.text_frame; tf2.word_wrap=False
        pr=tf2.paragraphs[0]; pr.alignment=PA.LEFT
        r2=pr.add_run(); r2.text="기본키 "
        r2.font.name=FONT; r2.font.size=Pt(pt_c); r2.font.bold=True; r2.font.color.rgb=_clr("#616161")
        r3=pr.add_run(); r3.text=col.name
        r3.font.name=FONT; r3.font.size=Pt(pt_c); r3.font.color.rgb=_clr("#424242")
        cy=rt+cfg.row_h/2
        col_pts[col.name]={"L":(left,cy),"R":(left+cfg.tbl_w,cy)}
    return col_pts, cfg.hdr_h+len(pk_cols)*cfg.row_h

def _gap_right(c,cfg): return MARGIN_L+c*(cfg.tbl_w+cfg.h_gap)+cfg.tbl_w+cfg.h_gap/2
def _gap_left(c,cfg):
    return (MARGIN_L-cfg.h_gap/2) if c==0 else MARGIN_L+c*(cfg.tbl_w+cfg.h_gap)-cfg.h_gap/2

def _compute_route(sc,sy,dc,dy,cfg):
    nc=cfg.num_cols
    if sc==dc:
        rx=_gap_right(sc,cfg)
        if sc==nc-1: rx=MARGIN_L+nc*(cfg.tbl_w+cfg.h_gap)+0.25
        return "R","R",[(rx,sy),(rx,dy)]
    elif sc<dc:
        r1=_gap_right(sc,cfg)
        if dc==sc+1: return "R","L",[(r1,sy),(r1,dy)]
        r2=_gap_left(dc,cfg)
        return "R","L",[(r1,sy),(r1,cfg.sky_y),(r2,cfg.sky_y),(r2,dy)]
    else:
        r1=_gap_left(sc,cfg)
        if sc==dc+1: return "L","R",[(r1,sy),(r1,dy)]
        r2=_gap_right(dc,cfg)
        return "L","R",[(r1,sy),(r1,cfg.sky_y),(r2,cfg.sky_y),(r2,dy)]

def _draw_segment(slide,x1,y1,x2,y2,clr,w):
    if abs(x1-x2)<0.001 and abs(y1-y2)<0.001: return
    try:
        from pptx.enum.shapes import MSO_CONNECTOR_TYPE
        conn=slide.shapes.add_connector(MSO_CONNECTOR_TYPE.STRAIGHT,Cm(x1),Cm(y1),Cm(x2),Cm(y2))
    except:
        conn=slide.shapes.add_connector(1,Cm(x1),Cm(y1),Cm(x2),Cm(y2))
    conn.line.color.rgb=clr; conn.line.width=w

def _draw_bezier(slide,x0,y0,cx0,cy0,cx1,cy1,x1,y1,hex_clr,pt_w):
    try:
        from lxml import etree
    except ImportError:
        _draw_segment(slide,x0,y0,x1,y1,_clr(hex_clr),Pt(pt_w)); return
    EMU2=360000
    all_x=[x0,cx0,cx1,x1]; all_y=[y0,cy0,cy1,y1]
    bx=min(all_x); by=min(all_y)
    bw=max(all_x)-bx; bh=max(all_y)-by
    pad=0.02
    if bw<pad: bx-=pad; bw=2*pad
    if bh<pad: by-=pad; bh=2*pad
    def _e(v,base): return str(max(0,int(round((v-base)*EMU2))))
    off_x=int(bx*EMU2); off_y=int(by*EMU2)
    ext_w=int(bw*EMU2); ext_h=int(bh*EMU2)
    lw_emu=int(pt_w*12700); clr_val=hex_clr.lstrip('#')
    xml_str=(
        '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<p:nvSpPr><p:cNvPr id="1" name="RI_curve"/>'
        '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>'
        '<p:spPr>'
        f'<a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_w}" cy="{ext_h}"/></a:xfrm>'
        '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        '<a:rect l="0" t="0" r="0" b="0"/><a:pathLst>'
        f'<a:path w="{ext_w}" h="{ext_h}">'
        f'<a:moveTo><a:pt x="{_e(x0,bx)}" y="{_e(y0,by)}"/></a:moveTo>'
        '<a:cubicBezTo>'
        f'<a:pt x="{_e(cx0,bx)}" y="{_e(cy0,by)}"/>'
        f'<a:pt x="{_e(cx1,bx)}" y="{_e(cy1,by)}"/>'
        f'<a:pt x="{_e(x1,bx)}" y="{_e(y1,by)}"/>'
        '</a:cubicBezTo></a:path></a:pathLst></a:custGeom>'
        '<a:noFill/>'
        f'<a:ln w="{lw_emu}"><a:solidFill><a:srgbClr val="{clr_val}"/></a:solidFill>'
        '<a:round/></a:ln></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )
    try:
        from lxml import etree as et2
        sp=et2.fromstring(xml_str); slide.shapes._spTree.append(sp)
    except: pass

def _bezier_sample(x0,y0,cx0,cy0,cx1,cy1,x1,y1,n=20):
    pts=[]
    for i in range(n+1):
        t=0.05+0.90*i/n; u=1-t
        px=u**3*x0+3*u**2*t*cx0+3*u*t**2*cx1+t**3*x1
        py=u**3*y0+3*u**2*t*cy0+3*u*t**2*cy1+t**3*y1
        pts.append((px,py))
    return pts

def _segs_cross2(a1,a2,b1,b2):
    def _cross(o,a,b): return (a[0]-o[0])*(b[1]-o[1])-(a[1]-o[1])*(b[0]-o[0])
    d1=_cross(b1,b2,a1);d2=_cross(b1,b2,a2);d3=_cross(a1,a2,b1);d4=_cross(a1,a2,b2)
    return (((d1>0 and d2<0)or(d1<0 and d2>0)) and ((d3>0 and d4<0)or(d3<0 and d4>0)))

def _beziers_cross(bzA,bzB,n=20):
    sA=_bezier_sample(*bzA,n=n); sB=_bezier_sample(*bzB,n=n)
    for i in range(len(sA)-1):
        for j in range(len(sB)-1):
            if _segs_cross2(sA[i],sA[i+1],sB[j],sB[j+1]): return True
    return False

def _draw_crowsfoot(slide,pts,clr,cfg):
    D=max(0.20,min(0.42,cfg.row_h*0.80)); H=D*0.55
    mlw=Pt(max(0.6,min(1.2,cfg.row_h*2.2))); plw=Pt(max(0.5,min(1.0,cfg.row_h*2.0)))
    if len(pts)<2: return
    sx,sy=pts[0]; nx,ny=pts[1]
    if abs(nx-sx)>=abs(ny-sy):
        bk=-1 if nx>sx else 1; bx2=sx+bk*D
        _draw_segment(slide,sx,sy,bx2,sy-H,clr,mlw)
        _draw_segment(slide,sx,sy,bx2,sy,  clr,mlw)
        _draw_segment(slide,sx,sy,bx2,sy+H,clr,mlw)
    else:
        bk=-1 if ny>sy else 1; by2=sy+bk*D
        _draw_segment(slide,sx,sy,sx-H,by2,clr,mlw)
        _draw_segment(slide,sx,sy,sx,  by2,clr,mlw)
        _draw_segment(slide,sx,sy,sx+H,by2,clr,mlw)
    ex2,ey2=pts[-1]; px2,py2=pts[-2]
    off=D*0.42; H2=D*0.50
    if abs(ex2-px2)>=abs(ey2-py2):
        dir_x=1 if ex2>px2 else -1
        _draw_segment(slide,ex2,ey2-H2,ex2,ey2+H2,clr,plw)
        _draw_segment(slide,ex2-dir_x*off,ey2-H2,ex2-dir_x*off,ey2+H2,clr,plw)
    else:
        dir_y=1 if ey2>py2 else -1
        _draw_segment(slide,ex2-H2,ey2,ex2+H2,ey2,clr,plw)
        _draw_segment(slide,ex2-H2,ey2-dir_y*off,ex2+H2,ey2-dir_y*off,clr,plw)

def _draw_relations(slide,rels,tbl_positions,cfg,
                    ri_h_spc=DEFAULT_RI_H_SPC,ri_v_spc=DEFAULT_RI_V_SPC):
    CLR_NORM="#C62828"; CLR_CROSS="#E65100"
    LW_PT=max(0.5,min(1.2,cfg.row_h*2.1))
    TENSION=0.45; U_RADIUS=max(0.25,cfg.h_gap*0.55)
    HY_BASE=cfg.hdr_h/2; TW=cfg.tbl_w
    drawn=set(); plans=[]
    for st,sc,dt,dc in rels:
        key=(st,sc,dt,dc)
        if key in drawn or st==dt: continue
        drawn.add(key)
        if st not in tbl_positions or dt not in tbl_positions: continue
        sl2,s_top,sci=tbl_positions[st]; dl2,d_top,dci=tbl_positions[dt]
        plans.append({'st':st,'dt':dt,'sl':sl2,'s_top':s_top,'sci':sci,
                      'dl':dl2,'d_top':d_top,'dci':dci})
    n=len(plans); sy_off=[0.0]*n; dy_off=[0.0]*n
    src_grp={}; dst_grp={}
    for i,rp in enumerate(plans):
        src_grp.setdefault(rp['st'],[]).append((i,rp['dci']))
        dst_grp.setdefault(rp['dt'],[]).append((i,rp['sci']))
    for _,lst in src_grp.items():
        if len(lst)<=1: continue
        lst_s=sorted(lst,key=lambda t:t[1]); nc=len(lst_s)
        for rank,(idx,_) in enumerate(lst_s):
            sy_off[idx]=(rank-(nc-1)/2.0)*ri_v_spc
    for _,lst in dst_grp.items():
        if len(lst)<=1: continue
        lst_s=sorted(lst,key=lambda t:t[1]); nc=len(lst_s)
        for rank,(idx,_) in enumerate(lst_s):
            dy_off[idx]=(rank-(nc-1)/2.0)*ri_v_spc
    beziers=[]
    for i,rp in enumerate(plans):
        sci=rp['sci'];dci=rp['dci'];sl2=rp['sl'];dl2=rp['dl']
        s_y=rp['s_top']+HY_BASE+sy_off[i]; d_y=rp['d_top']+HY_BASE+dy_off[i]
        if sci==dci:
            x0=sl2+TW;y0=s_y;x1=dl2+TW;y1=d_y
            u_r=U_RADIUS+(sy_off[i]*0.5); cx0=x0+u_r;cy0=y0;cx1=x1+u_r;cy1=y1
        elif sci<dci:
            x0=sl2+TW;y0=s_y;x1=dl2;y1=d_y;dx2=x1-x0
            cx0=x0+dx2*TENSION;cy0=y0;cx1=x1-dx2*TENSION;cy1=y1
        else:
            x0=sl2;y0=s_y;x1=dl2+TW;y1=d_y;dx2=x0-x1
            cx0=x0-dx2*TENSION;cy0=y0;cx1=x1+dx2*TENSION;cy1=y1
        beziers.append((x0,y0,cx0,cy0,cx1,cy1,x1,y1))
    crossed=[False]*n
    for i in range(n):
        for j in range(i+1,n):
            if _beziers_cross(beziers[i],beziers[j],n=20):
                crossed[i]=True; crossed[j]=True
    for i,rp in enumerate(plans):
        x0,y0,cx0,cy0,cx1,cy1,x1,y1=beziers[i]
        clr_hex=CLR_CROSS if crossed[i] else CLR_NORM
        clr_rgb=_clr(clr_hex)
        _draw_bezier(slide,x0,y0,cx0,cy0,cx1,cy1,x1,y1,clr_hex,LW_PT)
        marker_pts=[(x0,y0),(cx0,cy0),(cx1,cy1),(x1,y1)]
        _draw_crowsfoot(slide,marker_pts,clr_rgb,cfg)

def _phys_add_header(slide,title,n_main,n_ref=0,sub_areas=None):
    from pptx.enum.text import PP_ALIGN as PA
    ref_txt=f"  +외부참조 {n_ref}개" if n_ref else ""
    tb=slide.shapes.add_textbox(Cm(0.5),Cm(0.10),Cm(PPT_SLIDE_W-10),Cm(1.50))
    tf=tb.text_frame; tf.word_wrap=False
    p0=tf.paragraphs[0]; p0.alignment=PA.LEFT
    r0=p0.add_run(); r0.text=f"{title}  ({n_main}개 엔티티{ref_txt})"
    r0.font.size=Pt(13); r0.font.bold=True; r0.font.color.rgb=_clr("#1A1A2E")
    if sub_areas:
        shown=sub_areas[:6]; more=f"  외 {len(sub_areas)-6}개" if len(sub_areas)>6 else ""
        p1=tf.add_paragraph(); p1.alignment=PA.LEFT
        r1=p1.add_run(); r1.text="포함: "+" · ".join(shown)+more
        r1.font.size=Pt(8); r1.font.color.rgb=_clr("#555555")
    leg=slide.shapes.add_textbox(Cm(PPT_SLIDE_W-10),Cm(0.10),Cm(9.5),Cm(1.50))
    lf2=leg.text_frame; lf2.word_wrap=False
    lp=lf2.paragraphs[0]; lp.alignment=PA.RIGHT
    for txt,ch in [("▲ 상위  ","#1565C0"),("▼ 하위  ","#E65100"),
                   ("→| FK관계(삼발이)  ","#C62828"),("⊡ 외부참조","#888888")]:
        r=lp.add_run(); r.text=txt; r.font.size=Pt(8); r.font.color.rgb=_clr(ch)

def phys_render_slide(prs,blank,title,main_tbls,ref_tbls,all_tbl_map,slide_rels,
                      parents,children,cfg,max_cols,sub_areas=None,
                      ri_h_spc=DEFAULT_RI_H_SPC,ri_v_spc=DEFAULT_RI_V_SPC):
    actual_main=[t for t in main_tbls if t in all_tbl_map]
    actual_ref=list(ref_tbls.keys())
    if not actual_main: return
    slide=prs.slides.add_slide(blank)
    _phys_add_header(slide,title,len(actual_main),len(actual_ref),sub_areas=sub_areas)
    h_map={}
    for tn in actual_ref:
        pk_n=sum(1 for c in ref_tbls[tn].columns if c.is_pk)
        h_map[tn]=cfg.hdr_h+min(pk_n,max_cols)*cfg.row_h
    for tn in actual_main:
        h_map[tn]=_tbl_display_height(all_tbl_map[tn],cfg,max_cols)
    sorted_main=_sort_by_role(actual_main,parents,children)
    order=sorted_main+sorted(actual_ref)
    pos_list=_phys_layout(order,all_tbl_map,cfg,height_map=h_map)
    fk_col_keys={(st,sc) for st,sc,dt,dc in slide_rels}
    all_col_pts={}
    for tn,left,top,_ in pos_list:
        if tn in ref_tbls:
            pts,_=_phys_add_ref_table(slide,ref_tbls[tn],left,top,cfg,max_cols)
            all_col_pts[tn]=pts
        else:
            all_col_pts[tn]=_phys_add_table(
                slide,all_tbl_map[tn],left,top,tn in parents,tn in children,
                cfg,max_cols,fk_col_keys=fk_col_keys)
    on_slide=set(order)
    rels=[r for r in slide_rels if r[0] in on_slide and r[2] in on_slide]
    tbl_pos={tn:(left,top,ci) for tn,left,top,ci in pos_list}
    _draw_relations(slide,rels,tbl_pos,cfg,ri_h_spc=ri_h_spc,ri_v_spc=ri_v_spc)

def phys_export_pptx(all_tables,all_rels,opts,parents,children,out_path):
    global _color_cache; _color_cache={}
    prs=Presentation()
    prs.slide_width=Cm(PPT_SLIDE_W); prs.slide_height=Cm(PPT_SLIDE_H)
    blank=prs.slide_layouts[6]
    mode=opts["mode"]; mc=opts.get("max_cols",DEFAULT_MAX_COLS)
    user_tw=opts.get("tbl_w",DEFAULT_TBL_W); user_hg=opts.get("h_gap",DEFAULT_H_GAP)
    user_vg=opts.get("v_gap",DEFAULT_V_GAP)
    ri_h=opts.get("ri_h_spc",DEFAULT_RI_H_SPC); ri_v=opts.get("ri_v_spc",DEFAULT_RI_V_SPC)
    out_path=Path(out_path)
    if mode=="all":
        tbls=sorted(all_tables.keys())
        cfg=_auto_cfg(tbls,all_tables,mc,user_tbl_w=user_tw,user_h_gap=user_hg,user_v_gap=user_vg)
        phys_render_slide(prs,blank,"전체 ER 다이어그램",tbls,{},all_tables,all_rels,
                          parents,children,cfg,mc,ri_h_spc=ri_h,ri_v_spc=ri_v)
    elif mode=="areas":
        for aslide in opts["area_slides"]:
            title=aslide["title"]; tbls=aslide["tables"]
            incl=aslide.get("incl_rel",True)
            sub_areas=aslide.get("sub_areas",[])
            ref_map=_get_related_parents(tbls,all_rels,all_tables) if incl else {}
            child_ext=_get_related_children(tbls,all_rels,all_tables) if incl else {}
            ext_main=list(tbls)+[t for t in child_ext if t not in set(tbls)]
            combined={**{t:all_tables[t] for t in ext_main if t in all_tables},**ref_map}
            cfg=_auto_cfg(sorted(combined.keys()),combined,mc,
                          user_tbl_w=user_tw,user_h_gap=user_hg,user_v_gap=user_vg)
            phys_render_slide(prs,blank,title,ext_main,ref_map,combined,all_rels,
                              parents,children,cfg,mc,sub_areas=sub_areas,
                              ri_h_spc=ri_h,ri_v_spc=ri_v)
    elif mode=="tables":
        for ts in opts["table_slides"]:
            name=ts["name"]; tbls=ts["tables"]
            work={t:all_tables[t] for t in tbls if t in all_tables}
            cfg=_auto_cfg(tbls,work,mc,user_tbl_w=user_tw,user_h_gap=user_hg,user_v_gap=user_vg)
            on=set(tbls); rels=[r for r in all_rels if r[0] in on and r[2] in on]
            phys_render_slide(prs,blank,name,tbls,{},work,rels,parents,children,cfg,mc,
                              ri_h_spc=ri_h,ri_v_spc=ri_v)
    out_path.parent.mkdir(parents=True,exist_ok=True)
    try: prs.save(str(out_path))
    except PermissionError:
        ts2=datetime.now().strftime("%Y%m%d_%H%M%S")
        p2=out_path.with_stem(out_path.stem+f"_{ts2}"); prs.save(str(p2))

# ════════════════════════════════════════════════════════════════
# ⑦ GUI — 스크롤 탭 헬퍼
# ════════════════════════════════════════════════════════════════
def _make_scrollable_tab(nb, label):
    """노트북에 탭 추가 + 수직 스크롤바 + 내부 프레임 반환"""
    import tkinter as tk
    from tkinter import ttk
    tab = ttk.Frame(nb); nb.add(tab, text=label)
    cnv = tk.Canvas(tab, bg='#F0F0F0', highlightthickness=0)
    sb  = ttk.Scrollbar(tab, orient='vertical', command=cnv.yview)
    cnv.configure(yscrollcommand=sb.set)
    sb.pack(side='right', fill='y')
    cnv.pack(side='left', fill='both', expand=True)
    inner = tk.Frame(cnv, bg='#F0F0F0')
    win   = cnv.create_window((0,0), window=inner, anchor='nw')

    def _on_inner(e): cnv.configure(scrollregion=cnv.bbox('all'))
    def _on_cnv(e):   cnv.itemconfig(win, width=e.width)
    def _on_mw(e):    cnv.yview_scroll(int(-1*(e.delta/120)),'units')

    inner.bind('<Configure>', _on_inner)
    cnv.bind('<Configure>', _on_cnv)
    cnv.bind('<MouseWheel>', _on_mw)
    cnv.bind('<Button-4>', lambda e: cnv.yview_scroll(-1,'units'))
    cnv.bind('<Button-5>', lambda e: cnv.yview_scroll( 1,'units'))
    inner.bind('<MouseWheel>', _on_mw)
    inner.bind('<Button-4>', lambda e: cnv.yview_scroll(-1,'units'))
    inner.bind('<Button-5>', lambda e: cnv.yview_scroll( 1,'units'))
    return tab, inner


def _make_scrollable_frame(parent):
    """범용 스크롤 프레임 (탭이 아닌 일반 Frame 안에 삽입)"""
    import tkinter as tk
    from tkinter import ttk
    cnv = tk.Canvas(parent, bg='#F0F0F0', highlightthickness=0)
    sb  = ttk.Scrollbar(parent, orient='vertical', command=cnv.yview)
    cnv.configure(yscrollcommand=sb.set)
    sb.pack(side='right', fill='y')
    cnv.pack(side='left', fill='both', expand=True)
    inner = tk.Frame(cnv, bg='#F0F0F0')
    win   = cnv.create_window((0,0), window=inner, anchor='nw')

    def _on_inner(e): cnv.configure(scrollregion=cnv.bbox('all'))
    def _on_cnv(e):   cnv.itemconfig(win, width=e.width)
    def _on_mw(e):    cnv.yview_scroll(int(-1*(e.delta/120)),'units')

    inner.bind('<Configure>', _on_inner)
    cnv.bind('<Configure>', _on_cnv)
    cnv.bind('<MouseWheel>', _on_mw)
    cnv.bind('<Button-4>', lambda e: cnv.yview_scroll(-1,'units'))
    cnv.bind('<Button-5>', lambda e: cnv.yview_scroll( 1,'units'))
    inner.bind('<MouseWheel>', _on_mw)
    inner.bind('<Button-4>', lambda e: cnv.yview_scroll(-1,'units'))
    inner.bind('<Button-5>', lambda e: cnv.yview_scroll( 1,'units'))
    return inner

# ════════════════════════════════════════════════════════════════
# ⑧ GUI — 논리 모델 탭들
# ════════════════════════════════════════════════════════════════
def build_logical_tabs(nb):
    """논리 모델용 탭 3개 생성 후 모든 변수/함수 딕셔너리 반환"""
    import tkinter as tk
    from tkinter import ttk, filedialog, colorchooser

    # ── 탭1: 기본 설정 ──────────────────────────────────────────
    _, inner1 = _make_scrollable_tab(nb, '  [논리] 기본 설정  ')

    def lf(parent, title, pady=(4,4)):
        f = ttk.LabelFrame(parent, text=f' {title} ', padding=8)
        f.pack(fill='x', padx=10, pady=pady); return f

    def row_kv(f, label, var, unit='', r=None, c=0):
        ri = f.grid_size()[1] if r is None else r
        ttk.Label(f, text=label).grid(row=ri, column=c, sticky='w', pady=2)
        ttk.Entry(f, textvariable=var, width=9).grid(row=ri, column=c+1, padx=6, sticky='w')
        if unit:
            ttk.Label(f, text=unit, foreground='#666').grid(row=ri, column=c+2, sticky='w')

    # 파일
    ff = lf(inner1, '파일')
    xl_v  = tk.StringVar(value='민통선_엔티티속성정의서.xlsx')
    out_v = tk.StringVar(value='민통선_논리모델.pptx')
    for ri, (lbl, var, cmd_fn) in enumerate([
        ('입력 파일 (엑셀)', xl_v,
         lambda: xl_v.set(filedialog.askopenfilename(
             filetypes=[('Excel','*.xlsx *.xls'),('모든','*.*')]) or xl_v.get())),
        ('출력 파일 (.pptx)', out_v,
         lambda: out_v.set(filedialog.asksaveasfilename(
             defaultextension='.pptx',
             filetypes=[('PowerPoint','*.pptx'),('모든','*.*')],
             initialfile=out_v.get()) or out_v.get())),
    ]):
        ttk.Label(ff, text=lbl).grid(row=ri, column=0, sticky='w', pady=2)
        ttk.Entry(ff, textvariable=var, width=44).grid(row=ri, column=1, padx=6)
        ttk.Button(ff, text='찾아보기', command=cmd_fn).grid(row=ri, column=2, padx=4)

    # 슬라이드 크기
    sf = lf(inner1, '슬라이드 크기'); sf.columnconfigure(5, weight=1)
    sw_v=tk.StringVar(value='100'); sh_v=tk.StringVar(value='55')
    mg_v=tk.StringVar(value='0.8'); gp_v=tk.StringVar(value='0.35')
    row_kv(sf,'너비',sw_v,'cm',r=0,c=0); row_kv(sf,'높이',sh_v,'cm',r=1,c=0)
    row_kv(sf,'여백',mg_v,'cm',r=0,c=4); row_kv(sf,'간격',gp_v,'cm',r=1,c=4)

    # 엔티티명 박스
    enf = lf(inner1, '엔티티명 박스 크기 (상단 — 엔티티 이름 표시)')
    enf.columnconfigure(5, weight=1)
    ew_v=tk.StringVar(value='4.8'); eh_v=tk.StringVar(value='1.3')
    row_kv(enf,'너비',ew_v,'cm',r=0,c=0); row_kv(enf,'높이 (고정)',eh_v,'cm',r=1,c=0)
    ttk.Label(enf, text='※ 엔티티명이 표시되는 상단 헤더 도형의 너비·고정 높이입니다.',
              foreground='#888').grid(row=2,column=0,columnspan=6,sticky='w',pady=(2,0))

    # 속성 박스
    ef = lf(inner1, '속성 박스 크기 (하단 — 기본키/외래키 속성명 표시)')
    ef.columnconfigure(5, weight=1)
    aw_v=tk.StringVar(value='4.8'); er_v=tk.StringVar(value='0.48'); eam_v=tk.StringVar(value='0.48')
    row_kv(ef,'너비',aw_v,'cm',r=0,c=0)
    row_kv(ef,'행 높이 (기본키/외래키 1개당)',er_v,'cm',r=1,c=0)
    row_kv(ef,'최소 높이',eam_v,'cm',r=0,c=4)
    ttk.Label(ef,
              text=('※ 항목 표시 순서: ▶기본키(PK), ◆외래키(FK) — 속성명  데이터타입\n'
                    '   행 높이: 속성 1개당 높이 / 최소 높이: 속성이 없을 때의 하한'),
              foreground='#888').grid(row=2,column=0,columnspan=6,sticky='w',pady=(2,0))

    # 데이터타입 표시
    dtf = lf(inner1, '속성 표시 옵션')
    show_dtype_v = tk.BooleanVar(value=True)
    ttk.Checkbutton(dtf, text='속성명 옆에 데이터타입 함께 표시',
                    variable=show_dtype_v).grid(row=0,column=0,sticky='w',pady=2)

    # 엔티티 박스 내부 여백
    epf = lf(inner1, '엔티티 박스 내부 여백'); epf.columnconfigure(7, weight=1)
    en_ml_v=tk.StringVar(value='0.12'); en_mr_v=tk.StringVar(value='0.08')
    en_mt_v=tk.StringVar(value='0.06'); en_mb_v=tk.StringVar(value='0.04')
    row_kv(epf,'왼쪽 여백',en_ml_v,'cm',r=0,c=0)
    row_kv(epf,'오른쪽 여백',en_mr_v,'cm',r=1,c=0)
    row_kv(epf,'위쪽 여백',en_mt_v,'cm',r=0,c=4)
    row_kv(epf,'아래쪽 여백',en_mb_v,'cm',r=1,c=4)

    # 엔티티 박스 외부 여백
    egf = lf(inner1, '엔티티 박스 외부 여백'); egf.columnconfigure(7, weight=1)
    en_gh_v=tk.StringVar(value='0.35'); en_gv_v=tk.StringVar(value='0.35')
    en_pl_v=tk.StringVar(value='0.55'); en_pt_v=tk.StringVar(value='0.0')
    row_kv(egf,'가로 간격 (열 사이)',en_gh_v,'cm',r=0,c=0)
    row_kv(egf,'세로 간격 (행 사이)',en_gv_v,'cm',r=1,c=0)
    row_kv(egf,'주제영역 좌우 패딩',en_pl_v,'cm',r=0,c=4)
    row_kv(egf,'주제영역 상단 추가 패딩',en_pt_v,'cm',r=1,c=4)

    # 엔티티 박스 세로 오프셋
    eyof = lf(inner1, '엔티티 박스 세로 위치 오프셋'); eyof.columnconfigure(5, weight=1)
    en_yoff_v = tk.StringVar(value='0.0')
    row_kv(eyof,'세로 오프셋',en_yoff_v,'cm',r=0,c=0)

    # 레이아웃 여백
    lf2 = lf(inner1, '레이아웃 여백'); lf2.columnconfigure(5, weight=1)
    pd_v=tk.StringVar(value='0.55'); za_v=tk.StringVar(value='0.82')
    sa2_v=tk.StringVar(value='0.88'); da2_v=tk.StringVar(value='0.95')
    ttl_v=tk.StringVar(value='1.25')
    row_kv(lf2,'내부 패딩',pd_v,'cm',r=0,c=0)
    row_kv(lf2,'주제영역 H',za_v,'cm',r=1,c=0)
    row_kv(lf2,'상위주제 H',sa2_v,'cm',r=0,c=4)
    row_kv(lf2,'대주제 H',da2_v,'cm',r=1,c=4)
    row_kv(lf2,'슬라이드 제목 H',ttl_v,'cm',r=2,c=0)

    # 관계선
    rf = lf(inner1, '관계선'); rf.columnconfigure(5, weight=1)
    lc_v=tk.StringVar(value='404040'); cc_v=tk.StringVar(value='FF6600')
    ttk.Label(rf,text='관계선 색상 (HEX)').grid(row=0,column=0,sticky='w',pady=2)
    ttk.Entry(rf,textvariable=lc_v,width=10).grid(row=0,column=1,padx=6)
    ttk.Label(rf,text='교차선 색상 (HEX)').grid(row=0,column=3,sticky='w',pady=2)
    ttk.Entry(rf,textvariable=cc_v,width=10).grid(row=0,column=4,padx=6)

    # ── 탭2: 주제영역별 설정 ────────────────────────────────────
    _, inner2 = _make_scrollable_tab(nb, '  [논리] 주제영역별 설정  ')
    subj_vars = {}

    def mk_clr(parent, lbl, var):
        ttk.Label(parent, text=f' {lbl}').pack(side='left')
        e = ttk.Entry(parent, textvariable=var, width=8); e.pack(side='left', padx=1)
        def pick(_v=var):
            init='#'+_v.get().lstrip('#')
            res=colorchooser.askcolor(color=init,title='색상 선택')
            if res and res[1]: _v.set(res[1].lstrip('#'))
        ttk.Button(parent, text='●', width=2, command=pick).pack(side='left')

    def populate_tab2(da_list):
        for w in inner2.winfo_children(): w.destroy()
        subj_vars.clear()
        ttk.Label(inner2,
                  text='대주제영역별로 위치·글꼴·배경색·텍스트 오프셋을 설정하세요.',
                  foreground='#555').pack(pady=(6,4),padx=10,anchor='w')
        for da in da_list:
            sv = {
                'use_pos':     tk.BooleanVar(value=False),
                'x':           tk.StringVar(value='1.0'),
                'y':           tk.StringVar(value='2.5'),
                'da_font_sz':  tk.StringVar(value='13'),
                'sa_font_sz':  tk.StringVar(value='11'),
                'za_font_sz':  tk.StringVar(value='10'),
                'en_font_sz':  tk.StringVar(value='9'),
                'da_fc':       tk.StringVar(value='000000'),
                'da_bg':       tk.StringVar(value='BFBFBF'),
                'sa_bg':       tk.StringVar(value='D6D6D6'),
                'za_bg':       tk.StringVar(value='EBEBEB'),
                'en_name_bg':  tk.StringVar(value='D0D8E8'),
                'text_offset': tk.StringVar(value='0'),
            }
            subj_vars[da] = sv
            frm = ttk.LabelFrame(inner2, text=f'  {da}  ', padding=8)
            frm.pack(fill='x', padx=10, pady=4)
            r0 = tk.Frame(frm, bg='#F0F0F0'); r0.pack(fill='x', pady=2)
            ttk.Checkbutton(r0, text='위치 직접 지정', variable=sv['use_pos']).pack(side='left')
            for lbl2,key in [('X:','x'),('Y:','y')]:
                ttk.Label(r0, text=f'  {lbl2}').pack(side='left')
                ttk.Entry(r0, textvariable=sv[key], width=6).pack(side='left', padx=2)
                ttk.Label(r0, text='cm').pack(side='left')
            r1 = tk.Frame(frm, bg='#F0F0F0'); r1.pack(fill='x', pady=2)
            for lbl2,key in [('대주제 pt:','da_font_sz'),('상위주제:','sa_font_sz'),
                              ('주제영역:','za_font_sz'),('엔티티:','en_font_sz')]:
                ttk.Label(r1, text=f' {lbl2}').pack(side='left')
                ttk.Entry(r1, textvariable=sv[key], width=5).pack(side='left', padx=1)
            r2 = tk.Frame(frm, bg='#F0F0F0'); r2.pack(fill='x', pady=2)
            mk_clr(r2,'대주제 배경:',sv['da_bg'])
            mk_clr(r2,'상위주제 배경:',sv['sa_bg'])
            mk_clr(r2,'주제영역 배경:',sv['za_bg'])
            mk_clr(r2,'글꼴색:',sv['da_fc'])
            r2b = tk.Frame(frm, bg='#F0F0F0'); r2b.pack(fill='x', pady=2)
            mk_clr(r2b,'엔티티명 박스 배경:',sv['en_name_bg'])
            r3 = tk.Frame(frm, bg='#F0F0F0'); r3.pack(fill='x', pady=2)
            ttk.Label(r3, text='텍스트 오프셋:').pack(side='left')
            ttk.Entry(r3, textvariable=sv['text_offset'], width=7).pack(side='left', padx=4)
            ttk.Label(r3, text='cm').pack(side='left')

    # ── 탭3: 전체 글꼴 설정 ──────────────────────────────────────
    _, inner3 = _make_scrollable_tab(nb, '  [논리] 전체 글꼴 설정  ')

    def mk_clr_grid(parent, lbl, var, row_i, col_start):
        ttk.Label(parent, text=lbl).grid(row=row_i,column=col_start,sticky='w',pady=2,padx=(0,2))
        ttk.Entry(parent, textvariable=var, width=8).grid(row=row_i,column=col_start+1,padx=2)
        def pick(_v=var):
            init='#'+_v.get().lstrip('#')
            res=colorchooser.askcolor(color=init,title='색상 선택')
            if res and res[1]: _v.set(res[1].lstrip('#'))
        ttk.Button(parent,text='●',width=2,command=pick).grid(row=row_i,column=col_start+2,padx=(0,8))

    gf  = lf(inner3,'기본 글꼴 크기'); gf.columnconfigure(5,weight=1)
    gda_v=tk.StringVar(value='13'); gsa_v=tk.StringVar(value='11')
    gza_v=tk.StringVar(value='10'); gen_v=tk.StringVar(value='9')
    gat_v=tk.StringVar(value='7.5')
    row_kv(gf,'대주제 글꼴',gda_v,'pt',r=0,c=0)
    row_kv(gf,'상위주제 글꼴',gsa_v,'pt',r=1,c=0)
    row_kv(gf,'주제영역 글꼴',gza_v,'pt',r=2,c=0)
    row_kv(gf,'엔티티명 글꼴',gen_v,'pt',r=0,c=4)
    row_kv(gf,'속성 글꼴',gat_v,'pt',r=1,c=4)

    af = lf(inner3,'속성명 (기본키/외래키) 글꼴 크기·색상·굵기'); af.columnconfigure(9,weight=1)
    pk_fsz_v=tk.StringVar(value='7.5'); fk_fsz_v=tk.StringVar(value='7.5')
    pk_clr_v=tk.StringVar(value='1F3864'); fk_clr_v=tk.StringVar(value='4A4A4A')
    pk_bold_v=tk.BooleanVar(value=True); fk_bold_v=tk.BooleanVar(value=False)
    row_kv(af,'기본키 글꼴 크기',pk_fsz_v,'pt',r=0,c=0)
    mk_clr_grid(af,'기본키 색상:',pk_clr_v,0,4)
    ttk.Checkbutton(af,text='굵게',variable=pk_bold_v).grid(row=0,column=7,padx=(4,0),sticky='w')
    row_kv(af,'외래키 글꼴 크기',fk_fsz_v,'pt',r=1,c=0)
    mk_clr_grid(af,'외래키 색상:',fk_clr_v,1,4)
    ttk.Checkbutton(af,text='굵게',variable=fk_bold_v).grid(row=1,column=7,padx=(4,0),sticky='w')
    ttk.Label(af,
              text='※ 기본키(▶) · 외래키(◆) 속성 행에 각각 독립 적용됩니다.',
              foreground='#888').grid(row=2,column=0,columnspan=10,sticky='w',pady=(2,0))

    tf3b = lf(inner3,'텍스트 오프셋 (전체 기본)')
    gtoff_v = tk.StringVar(value='0')
    ttk.Label(tf3b,text='전체 기본 오프셋:').grid(row=0,column=0,sticky='w')
    ttk.Entry(tf3b,textvariable=gtoff_v,width=10).grid(row=0,column=1,padx=6)
    ttk.Label(tf3b,text='cm').grid(row=0,column=2,sticky='w')

    # ── get_cfg ─────────────────────────────────────────────────
    def get_cfg():
        cfg = {
            'slide_w':float(sw_v.get()),'slide_h':float(sh_v.get()),
            'margin':float(mg_v.get()),'gap':float(gp_v.get()),
            'entity_w':float(ew_v.get()),'entity_h':float(eh_v.get()),
            'attr_w':float(aw_v.get()),'attr_row_h':float(er_v.get()),
            'attr_min_h':float(eam_v.get()),
            'show_dtype':show_dtype_v.get(),
            'pad':float(pd_v.get()),'hdr_za':float(za_v.get()),
            'hdr_sa':float(sa2_v.get()),'hdr_da':float(da2_v.get()),
            'title_h':float(ttl_v.get()),
            'da_font_sz':float(gda_v.get()),'sa_font_sz':float(gsa_v.get()),
            'za_font_sz':float(gza_v.get()),'en_font_sz':float(gen_v.get()),
            'attr_font_sz':float(gat_v.get()),'text_offset':float(gtoff_v.get()),
            'line_color':lc_v.get().strip(),'cross_color':cc_v.get().strip(),
            'en_margin_l':float(en_ml_v.get()),'en_margin_r':float(en_mr_v.get()),
            'en_margin_t':float(en_mt_v.get()),'en_margin_b':float(en_mb_v.get()),
            'en_gap_h':float(en_gh_v.get()),'en_gap_v':float(en_gv_v.get()),
            'en_pad_l':float(en_pl_v.get()),'en_pad_t':float(en_pt_v.get()),
            'en_y_offset':float(en_yoff_v.get()),
            'pk_font_sz':float(pk_fsz_v.get()),'fk_font_sz':float(fk_fsz_v.get()),
            'pk_color':pk_clr_v.get().strip(),'fk_color':fk_clr_v.get().strip(),
            'pk_bold':pk_bold_v.get(),'fk_bold':fk_bold_v.get(),
            'styles':{},
        }
        positions = {}
        for da, sv in subj_vars.items():
            cfg['styles'][da] = {
                'da_font_sz':float(sv['da_font_sz'].get()),
                'sa_font_sz':float(sv['sa_font_sz'].get()),
                'za_font_sz':float(sv['za_font_sz'].get()),
                'en_font_sz':float(sv['en_font_sz'].get()),
                'da_fc':sv['da_fc'].get().strip(),
                'da_bg':sv['da_bg'].get().strip(),
                'sa_bg':sv['sa_bg'].get().strip(),
                'za_bg':sv['za_bg'].get().strip(),
                'en_name_bg':sv['en_name_bg'].get().strip(),
                'text_offset':float(sv['text_offset'].get()),
            }
            if sv['use_pos'].get():
                positions[da]={'x':float(sv['x'].get()),'y':float(sv['y'].get())}
        return cfg, positions

    return dict(
        xl_v=xl_v, out_v=out_v,
        get_cfg=get_cfg,
        populate_tab2=populate_tab2,
        subj_vars=subj_vars,
        sw_v=sw_v,sh_v=sh_v,mg_v=mg_v,gp_v=gp_v,
        ew_v=ew_v,eh_v=eh_v,aw_v=aw_v,er_v=er_v,eam_v=eam_v,
        show_dtype_v=show_dtype_v,
        pd_v=pd_v,za_v=za_v,sa2_v=sa2_v,da2_v=da2_v,ttl_v=ttl_v,
        gda_v=gda_v,gsa_v=gsa_v,gza_v=gza_v,gen_v=gen_v,gat_v=gat_v,
        gtoff_v=gtoff_v,lc_v=lc_v,cc_v=cc_v,
        en_ml_v=en_ml_v,en_mr_v=en_mr_v,en_mt_v=en_mt_v,en_mb_v=en_mb_v,
        en_gh_v=en_gh_v,en_gv_v=en_gv_v,en_pl_v=en_pl_v,en_pt_v=en_pt_v,
        en_yoff_v=en_yoff_v,
        pk_fsz_v=pk_fsz_v,fk_fsz_v=fk_fsz_v,
        pk_clr_v=pk_clr_v,fk_clr_v=fk_clr_v,
        pk_bold_v=pk_bold_v,fk_bold_v=fk_bold_v,
    )

# ════════════════════════════════════════════════════════════════
# ⑨ GUI — 물리 모델 탭
# ════════════════════════════════════════════════════════════════
def build_physical_tab(nb):
    """물리 모델용 탭 1개 생성 후 변수/함수 딕셔너리 반환"""
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog

    _, inner = _make_scrollable_tab(nb, '  [물리] ER 다이어그램  ')

    # ── 파일 ────────────────────────────────────────────────────
    ff = ttk.LabelFrame(inner, text=' 파일 ', padding=8)
    ff.pack(fill='x', padx=10, pady=4)
    phys_xl_v  = tk.StringVar(value='물리모델_엔티티속성정의서.xlsx')
    phys_out_v = tk.StringVar(value='물리모델_ER다이어그램.pptx')
    for ri,(lbl,var,cmd_fn) in enumerate([
        ('입력 파일 (엑셀)', phys_xl_v,
         lambda: phys_xl_v.set(filedialog.askopenfilename(
             filetypes=[('Excel','*.xlsx *.xls'),('모든','*.*')]) or phys_xl_v.get())),
        ('출력 파일 (.pptx)', phys_out_v,
         lambda: phys_out_v.set(filedialog.asksaveasfilename(
             defaultextension='.pptx',
             filetypes=[('PowerPoint','*.pptx'),('모든','*.*')],
             initialfile=phys_out_v.get()) or phys_out_v.get())),
    ]):
        ttk.Label(ff,text=lbl).grid(row=ri,column=0,sticky='w',pady=2)
        ttk.Entry(ff,textvariable=var,width=44).grid(row=ri,column=1,padx=6)
        ttk.Button(ff,text='찾아보기',command=cmd_fn).grid(row=ri,column=2,padx=4)

    # ── 도형·관계선 설정 ─────────────────────────────────────────
    setf = ttk.LabelFrame(inner, text=' 도형·관계선 설정 ', padding=8)
    setf.pack(fill='x', padx=10, pady=4)

    def _lbl(p,t,w=None):
        kw={"font":(FONT,10)}
        if w: kw["width"]=w
        return ttk.Label(p,text=t,**kw)
    def _spin(p,var,fr,to,inc,fmt,w=6):
        return tk.Spinbox(p,from_=fr,to=to,increment=inc,
                          textvariable=var,width=w,format=fmt,font=(FONT,10))

    row1=tk.Frame(setf); row1.pack(fill='x',pady=(0,3))
    _lbl(row1,"엔티티당 최대 속성 표시 수:").pack(side='left')
    max_cols_var=tk.IntVar(value=DEFAULT_MAX_COLS)
    tk.Spinbox(row1,from_=1,to=50,textvariable=max_cols_var,width=5,
               font=(FONT,10)).pack(side='left',padx=(4,6))
    ttk.Label(row1,text="개  (초과 시 '+N개 더' 표시)").pack(side='left')

    row2=tk.Frame(setf); row2.pack(fill='x',pady=(0,3))
    _lbl(row2,"엔티티 도형 너비:").pack(side='left')
    tbl_w_var=tk.DoubleVar(value=DEFAULT_TBL_W)
    _spin(row2,tbl_w_var,4.0,16.0,0.5,"%.1f").pack(side='left',padx=(4,6))
    ttk.Label(row2,text="cm").pack(side='left')

    row3=tk.Frame(setf); row3.pack(fill='x',pady=(0,3))
    _lbl(row3,"엔티티 간격 —  가로:").pack(side='left')
    h_gap_var=tk.DoubleVar(value=DEFAULT_H_GAP)
    _spin(row3,h_gap_var,0.1,5.0,0.1,"%.1f").pack(side='left',padx=(4,10))
    _lbl(row3,"cm    세로:").pack(side='left')
    v_gap_var=tk.DoubleVar(value=DEFAULT_V_GAP)
    _spin(row3,v_gap_var,0.1,5.0,0.1,"%.1f").pack(side='left',padx=(4,6))
    ttk.Label(row3,text="cm").pack(side='left')

    row4=tk.Frame(setf); row4.pack(fill='x')
    _lbl(row4,"관계선(RI) 간격 —  가로:").pack(side='left')
    ri_h_var=tk.DoubleVar(value=DEFAULT_RI_H_SPC)
    _spin(row4,ri_h_var,0.02,2.0,0.05,"%.2f").pack(side='left',padx=(4,10))
    _lbl(row4,"cm    세로:").pack(side='left')
    ri_v_var=tk.DoubleVar(value=DEFAULT_RI_V_SPC)
    _spin(row4,ri_v_var,0.02,2.0,0.05,"%.2f").pack(side='left',padx=(4,6))
    ttk.Label(row4,text="cm").pack(side='left')

    # ── 주제영역 설정 (트리) ─────────────────────────────────────
    sa_lf = ttk.LabelFrame(inner, text=' 주제영역 설정 (L1 기준 슬라이드 생성) ', padding=6)
    sa_lf.pack(fill='both', expand=True, padx=10, pady=4)

    sa_data = {}; cur_area = [None]
    mode_var = tk.StringVar(value="areas")

    mlf2 = tk.Frame(sa_lf); mlf2.pack(fill='x', pady=(0,4))
    for val,lbl2 in [("all","전체 생성"),("areas","주제영역 직접 작성"),("tables","엔티티 선택")]:
        ttk.Radiobutton(mlf2,text=lbl2,variable=mode_var,value=val).pack(side='left',padx=8)

    area_body = tk.Frame(sa_lf); area_body.pack(fill='both', expand=True)
    a_left = tk.Frame(area_body); a_left.pack(side='left', fill='y', padx=(0,5))
    a_right = ttk.LabelFrame(area_body, text=' 엔티티 할당 ', padding=4)
    a_right.pack(side='left', fill='both', expand=True)

    tree_lf2 = ttk.LabelFrame(a_left, text=' 주제영역 계층 ', padding=4)
    tree_lf2.pack(fill='both', expand=True)
    tr_sb2 = tk.Scrollbar(tree_lf2, orient='vertical')
    sa_tree = ttk.Treeview(tree_lf2, yscrollcommand=tr_sb2.set,
                            selectmode='browse', height=12, show='tree headings')
    sa_tree["columns"] = ("level","count")
    sa_tree.heading("#0",text="주제영역"); sa_tree.heading("level",text="레벨")
    sa_tree.heading("count",text="엔티티")
    sa_tree.column("#0",width=180); sa_tree.column("level",width=40,anchor='center')
    sa_tree.column("count",width=50,anchor='center')
    tr_sb2.config(command=sa_tree.yview)
    tr_sb2.pack(side='right',fill='y'); sa_tree.pack(fill='both',expand=True)

    def _sa_insert(parent_iid, title, level):
        iid = sa_tree.insert(parent_iid,'end',text=f"  {title}",
                             values=(f"L{level}","0"))
        sa_data[iid]={"title":title,"level":level,"tables":set()}
        if parent_iid: sa_tree.item(parent_iid,open=True)
        return iid

    def _upd_count(iid):
        if iid in sa_data:
            n=len(sa_data[iid]["tables"]); lv=sa_data[iid]["level"]
            sa_tree.item(iid,values=(f"L{lv}",str(n)))

    def _add_root():
        t=simpledialog.askstring("주제영역 추가","주제영역 제목 입력:")
        if t and t.strip(): _sa_insert("",t.strip(),1)

    def _add_child():
        sel=sa_tree.selection()
        if not sel: return
        iid=sel[0]; lv=sa_data.get(iid,{}).get("level",1)
        if lv>=4: return
        t=simpledialog.askstring("하위 주제영역","하위 제목 입력:")
        if t and t.strip(): _sa_insert(iid,t.strip(),lv+1)

    def _rename():
        sel=sa_tree.selection()
        if not sel: return
        iid=sel[0]; old=sa_data.get(iid,{}).get("title","")
        nt=simpledialog.askstring("이름 변경","새 제목:",initialvalue=old)
        if nt and nt.strip():
            sa_data[iid]["title"]=nt.strip(); sa_tree.item(iid,text=f"  {nt.strip()}")

    def _delete():
        sel=sa_tree.selection()
        if not sel: return
        def _del_r(n):
            for c in sa_tree.get_children(n): _del_r(c)
            sa_data.pop(n,None)
        _del_r(sel[0]); sa_tree.delete(sel[0])

    tbf2 = tk.Frame(a_left); tbf2.pack(fill='x',pady=(4,2))
    for txt2,cmd2,bg2 in [("L1추가",_add_root,"#D0E8FF"),("하위추가",_add_child,"#E8F5E9"),
                           ("이름변경",_rename,"#FFF9C4"),("삭제",_delete,"#FFEBEE")]:
        tk.Button(tbf2,text=txt2,command=cmd2,font=(FONT,8),relief='flat',
                  bg=bg2,padx=6,pady=3).pack(side='left',padx=2)

    # 오른쪽: 엔티티 할당
    a_title_lbl = ttk.Label(a_right,text="← 왼쪽 주제영역을 클릭하면 엔티티 목록이 표시됩니다",
                             foreground='#888888')
    a_title_lbl.pack(anchor='w',pady=(0,3))
    a_btn_row2=tk.Frame(a_right); a_btn_row2.pack(fill='x',pady=(0,3))
    a_sv2=tk.StringVar()
    a_se2=ttk.Entry(a_btn_row2,textvariable=a_sv2,width=18)
    a_se2.pack(side='left',padx=(0,6))
    incl_rel_var=tk.BooleanVar(value=True)

    a_tbl_vars = {}

    def _aall2():
        if not cur_area[0]: return
        for t,v in a_tbl_vars.items():
            q=a_sv2.get()
            if q.upper() in t.upper() or not q: v.set(True)
        _save_area2()

    def _anone2():
        if not cur_area[0]: return
        for t,v in a_tbl_vars.items():
            q=a_sv2.get()
            if q.upper() in t.upper() or not q: v.set(False)
        _save_area2()

    tk.Button(a_btn_row2,text="전체선택",command=_aall2,font=(FONT,8),
              relief='flat',bg="#D0E8FF",padx=7,pady=2).pack(side='left',padx=(0,3))
    tk.Button(a_btn_row2,text="전체해제",command=_anone2,font=(FONT,8),
              relief='flat',bg="#FFE0CC",padx=7,pady=2).pack(side='left')
    ttk.Label(a_btn_row2,text="  관계 엔티티 자동포함:").pack(side='left',padx=(14,2))
    ttk.Checkbutton(a_btn_row2,variable=incl_rel_var).pack(side='left')

    a_cv2=tk.Canvas(a_right,bg='#F4F4F4',highlightthickness=0)
    a_sb2=ttk.Scrollbar(a_right,orient='vertical',command=a_cv2.yview)
    a_inner2=tk.Frame(a_cv2,bg='#F4F4F4')
    a_inner2.bind('<Configure>',lambda e:a_cv2.configure(scrollregion=a_cv2.bbox('all')))
    a_win2=a_cv2.create_window((0,0),window=a_inner2,anchor='nw')
    a_cv2.configure(yscrollcommand=a_sb2.set)

    def _on_a_cv2_cnf(e): a_cv2.itemconfig(a_win2, width=e.width)
    def _on_a_cv2_mw(e):  a_cv2.yview_scroll(int(-1*(e.delta/120)),'units')
    a_cv2.bind('<Configure>', _on_a_cv2_cnf)
    a_cv2.bind('<MouseWheel>', _on_a_cv2_mw)
    a_cv2.bind('<Button-4>', lambda e: a_cv2.yview_scroll(-1,'units'))
    a_cv2.bind('<Button-5>', lambda e: a_cv2.yview_scroll( 1,'units'))
    a_inner2.bind('<MouseWheel>', _on_a_cv2_mw)
    a_inner2.bind('<Button-4>', lambda e: a_cv2.yview_scroll(-1,'units'))
    a_inner2.bind('<Button-5>', lambda e: a_cv2.yview_scroll( 1,'units'))

    a_sb2.pack(side='right',fill='y'); a_cv2.pack(fill='both',expand=True)

    all_tbls_cache = [None]

    def _refresh_area_tbls2(*_):
        for w in a_inner2.winfo_children(): w.destroy()
        if not cur_area[0] or not all_tbls_cache[0]: return
        q=a_sv2.get().strip()
        parents2,children2=phys_detect_parent_child(all_tbls_cache[0].get('rels',[]))
        for t in sorted(a_tbl_vars.keys()):
            if q.upper() in t.upper() or not q:
                role=(" ▲▼" if t in parents2 and t in children2 else
                      " ▲" if t in parents2 else " ▼" if t in children2 else "")
                tk.Checkbutton(a_inner2,text=t+role,variable=a_tbl_vars[t],
                               bg='#F4F4F4',font=('Consolas',9),anchor='w',
                               command=_save_area2).pack(fill='x',padx=2,pady=1)

    a_sv2.trace_add('write',_refresh_area_tbls2)

    def _save_area2():
        if cur_area[0] and cur_area[0] in sa_data:
            sa_data[cur_area[0]]["tables"]={t for t,v in a_tbl_vars.items() if v.get()}
            _upd_count(cur_area[0])

    def _load_area2(iid):
        cur_area[0]=iid
        tbls2=sa_data.get(iid,{}).get("tables",set())
        for t,v in a_tbl_vars.items(): v.set(t in tbls2)
        title2=sa_data.get(iid,{}).get("title","")
        lv2=sa_data.get(iid,{}).get("level",1)
        a_title_lbl.config(text=f"[L{lv2}] {title2}  —  엔티티 선택")
        _refresh_area_tbls2()

    def _on_sa_tree_sel(evt):
        _save_area2()
        sel=sa_tree.selection()
        if sel: _load_area2(sel[0])

    sa_tree.bind("<<TreeviewSelect>>",_on_sa_tree_sel)

    def refresh_tbl_vars(tables_dict):
        """물리 모델 로드 후 엔티티 체크박스 갱신"""
        a_tbl_vars.clear()
        a_tbl_vars.update({t:tk.BooleanVar(value=False) for t in sorted(tables_dict.keys())})
        _refresh_area_tbls2()

    return dict(
        phys_xl_v=phys_xl_v, phys_out_v=phys_out_v,
        max_cols_var=max_cols_var, tbl_w_var=tbl_w_var,
        h_gap_var=h_gap_var, v_gap_var=v_gap_var,
        ri_h_var=ri_h_var, ri_v_var=ri_v_var,
        mode_var=mode_var, sa_data=sa_data, sa_tree=sa_tree,
        incl_rel_var=incl_rel_var, a_tbl_vars=a_tbl_vars,
        all_tbls_cache=all_tbls_cache,
        refresh_tbl_vars=refresh_tbl_vars,
        _save_area2=_save_area2,
    )

# ════════════════════════════════════════════════════════════════
# ⑩ GUI — FK 편집 다이얼로그 (물리 모델 전용)
# ════════════════════════════════════════════════════════════════
def _show_fk_dialog(parent_root, tables, existing_rels, on_done=None):
    import tkinter as tk
    from tkinter import ttk, messagebox

    auto_set  = set((r[0],r[1],r[2],r[3]) for r in existing_rels)
    saved_rels = phys_load_fk()
    manual_init = []
    for r in saved_rels:
        key=(r[0],r[1],r[2],r[3] if len(r)>3 else "")
        if key not in auto_set: manual_init.append(key)

    final_rels = [None]
    dlg = tk.Toplevel(parent_root)
    dlg.title("물리 모델 — FK 관계 현황 및 편집")
    dlg.geometry("1100x700"); dlg.resizable(True,True)
    dlg.configure(bg="#F4F4F4"); dlg.grab_set()

    # 헤더
    hdr2=tk.Frame(dlg,bg="#1A237E",pady=7); hdr2.pack(fill='x')
    tk.Label(hdr2,text="FK 관계 현황 및 편집",bg="#1A237E",fg="white",
             font=(FONT,12,"bold")).pack()
    tk.Label(hdr2,text="자동 감지(파랑) / 수동 지정(녹)을 통합 관리합니다",
             bg="#1A237E",fg="#C5CAE9",font=(FONT,8)).pack()

    # 트리뷰 (스크롤바 포함)
    tv_lf2=tk.LabelFrame(dlg,text="  전체 FK 관계 현황  ",bg="#F4F4F4",
                          font=(FONT,9,"bold"),padx=5,pady=4)
    tv_lf2.pack(fill='both',expand=True,padx=8,pady=(4,2))

    # 열 머리글: 엔티티/속성 용어 적용
    TV_COLS2=("하위 엔티티","외래키 속성","상위 엔티티","기본키 속성","구분")
    tv2=ttk.Treeview(tv_lf2,columns=TV_COLS2,show="headings",height=14,selectmode="extended")
    for cn,w in [("하위 엔티티",200),("외래키 속성",140),("상위 엔티티",200),("기본키 속성",140),("구분",65)]:
        tv2.heading(cn,text=cn); tv2.column(cn,width=w,minwidth=60,
                                              anchor="center" if cn=="구분" else "w")
    vsb=ttk.Scrollbar(tv_lf2,orient="vertical",command=tv2.yview)
    hsb=ttk.Scrollbar(tv_lf2,orient="horizontal",command=tv2.xview)
    tv2.configure(yscrollcommand=vsb.set,xscrollcommand=hsb.set)
    hsb.pack(side='bottom',fill='x'); vsb.pack(side='right',fill='y'); tv2.pack(fill='both',expand=True)
    tv2.tag_configure("auto",background="#EBF5FB",foreground="#1A237E")
    tv2.tag_configure("manual",background="#E9F7EF",foreground="#1B5E20")

    # 마우스휠 스크롤 (Windows / Linux)
    def _tv_mw(e):
        tv2.yview_scroll(int(-1*(e.delta/120)),'units')
    tv2.bind('<MouseWheel>', _tv_mw)
    tv2.bind('<Button-4>', lambda e: tv2.yview_scroll(-1,'units'))
    tv2.bind('<Button-5>', lambda e: tv2.yview_scroll( 1,'units'))

    def _tv_add(ct,fc,pt,pc,typ):
        tag="auto" if typ=="자동" else "manual"
        tv2.insert("",tk.END,values=(ct,fc,pt,pc,typ),tags=(tag,))

    for ct,fc,pt,pc in sorted(auto_set): _tv_add(ct,fc,pt,pc,"자동")
    for ct,fc,pt,pc in sorted(manual_init): _tv_add(ct,fc,pt,pc,"수동")

    all_tbl_names=sorted(tables.keys())

    def _open_edit(prefill=None,edit_iid=None):
        d=tk.Toplevel(dlg); d.title("관계 편집" if edit_iid else "관계 추가")
        d.geometry("760x260"); d.configure(bg="#F4F4F4"); d.grab_set()

        # 스크롤 가능한 내부 프레임
        d_scroll_outer = tk.Frame(d, bg="#F4F4F4"); d_scroll_outer.pack(fill='both', expand=True)
        d_inner = _make_scrollable_frame(d_scroll_outer)

        frm=tk.Frame(d_inner,bg="#F4F4F4"); frm.pack(padx=16,pady=10,fill='both')
        tk.Label(frm,text="하위 엔티티(외래키)",bg="#F4F4F4",font=(FONT,9)).grid(row=0,column=0,padx=6,pady=5,sticky="w")
        child_var=tk.StringVar(value=prefill[0] if prefill else "")
        child_cb=ttk.Combobox(frm,textvariable=child_var,values=all_tbl_names,width=30)
        child_cb.grid(row=0,column=1,padx=6,pady=5,sticky="w")
        tk.Label(frm,text="외래키 속성명",bg="#F4F4F4",font=(FONT,9)).grid(row=0,column=2,padx=6,pady=5,sticky="w")
        fk_col_var=tk.StringVar(value=prefill[1] if prefill else "")
        fk_col_cb=ttk.Combobox(frm,textvariable=fk_col_var,width=25)
        fk_col_cb.grid(row=0,column=3,padx=6,pady=5,sticky="w")
        tk.Label(frm,text="상위 엔티티(기본키)",bg="#F4F4F4",font=(FONT,9)).grid(row=1,column=0,padx=6,pady=5,sticky="w")
        parent_var=tk.StringVar(value=prefill[2] if prefill else "")
        parent_cb=ttk.Combobox(frm,textvariable=parent_var,values=all_tbl_names,width=30)
        parent_cb.grid(row=1,column=1,padx=6,pady=5,sticky="w")
        tk.Label(frm,text="상위 기본키 속성명",bg="#F4F4F4",font=(FONT,9)).grid(row=1,column=2,padx=6,pady=5,sticky="w")
        pk_col_var=tk.StringVar(value=prefill[3] if prefill else "")
        pk_col_cb=ttk.Combobox(frm,textvariable=pk_col_var,width=25)
        pk_col_cb.grid(row=1,column=3,padx=6,pady=5,sticky="w")

        def _upd_fk(e=None):
            ct2=child_var.get(); pt2=parent_var.get()
            child_cols=[c.name for c in tables[ct2].columns] if ct2 in tables else []
            pk_cols=[c.name for c in tables[pt2].columns if c.is_pk] if pt2 in tables else []
            merged=child_cols.copy()
            for p3 in pk_cols:
                if p3 not in merged: merged.append(p3)
            fk_col_cb["values"]=merged

        def _upd_pk(e=None):
            pt2=parent_var.get()
            if pt2 in tables:
                pks=[c.name for c in tables[pt2].columns if c.is_pk]
                pk_col_cb["values"]=pks
                if pks and not pk_col_var.get(): pk_col_var.set(pks[0])
            _upd_fk()

        child_cb.bind("<<ComboboxSelected>>",_upd_fk)
        parent_cb.bind("<<ComboboxSelected>>",_upd_pk)
        _upd_fk(); _upd_pk()

        def _do_apply():
            ct2=child_var.get().strip(); fc2=fk_col_var.get().strip()
            pt2=parent_var.get().strip(); pc2=pk_col_var.get().strip()
            if not(ct2 and fc2 and pt2 and pc2):
                messagebox.showwarning("입력 필요","모든 항목을 입력하세요.",parent=d); return
            if edit_iid: tv2.delete(edit_iid)
            _tv_add(ct2,fc2,pt2,pc2,"수동"); d.destroy()

        btn_f=tk.Frame(d,bg="#F4F4F4"); btn_f.pack(pady=(0,10))
        tk.Button(btn_f,text="  적용  ",command=_do_apply,bg="#1A1A2E",fg="white",
                  font=(FONT,10,"bold"),relief='flat',padx=12,pady=4).pack(side='left',padx=4)
        tk.Button(btn_f,text="  취소  ",command=d.destroy,bg="#888888",fg="white",
                  font=(FONT,10),relief='flat',padx=12,pady=4).pack(side='left')

    def _del_sel():
        sel=tv2.selection()
        if not sel: return
        if messagebox.askyesno("삭제 확인",f"{len(sel)}개 관계를 삭제하시겠습니까?",parent=dlg):
            for iid in sel: tv2.delete(iid)

    def _edit_sel():
        sel=tv2.selection()
        if len(sel)!=1: return
        vals=tv2.item(sel[0],"values")
        _open_edit(prefill=(vals[0],vals[1],vals[2],vals[3]),edit_iid=sel[0])

    tv2.bind("<Double-1>",lambda e:_edit_sel())
    tv_btn2=tk.Frame(tv_lf2,bg="#F4F4F4"); tv_btn2.pack(fill='x',pady=(4,0))
    tk.Button(tv_btn2,text="  관계 추가  ",command=lambda:_open_edit(),
              bg="#C8E6C9",fg="#1B5E20",font=(FONT,9),relief='flat',padx=10,pady=3).pack(side='left',padx=(0,4))
    tk.Button(tv_btn2,text="  선택 편집  ",command=_edit_sel,
              bg="#E3F2FD",fg="#0D47A1",font=(FONT,9),relief='flat',padx=10,pady=3).pack(side='left',padx=(0,4))
    tk.Button(tv_btn2,text="  선택 삭제  ",command=_del_sel,
              bg="#FFCDD2",fg="#B71C1C",font=(FONT,9),relief='flat',padx=10,pady=3).pack(side='left')

    def _collect2():
        result=[]
        for iid in tv2.get_children():
            v=tv2.item(iid,"values"); result.append((v[0],v[1],v[2],v[3]))
        return result

    def _ok2():
        rels2=_collect2()
        manual_save=[(ct2,fc2,pt2,pc2) for ct2,fc2,pt2,pc2 in rels2
                     if (ct2,fc2,pt2,pc2) not in auto_set]
        phys_save_fk(manual_save)
        final_rels[0]=rels2; dlg.destroy()
        if on_done: on_done(rels2)

    def _cancel2():
        final_rels[0]=list(existing_rels); dlg.destroy()

    bf3=tk.Frame(dlg,bg="#F4F4F4"); bf3.pack(fill='x',padx=10,pady=(0,8))
    tk.Button(bf3,text="  확인  ",command=_ok2,bg="#1A1A2E",fg="white",
              font=(FONT,10,"bold"),relief='flat',padx=14,pady=5).pack(side='right',padx=(5,0))
    tk.Button(bf3,text="  취소  ",command=_cancel2,bg="#888888",fg="white",
              font=(FONT,10),relief='flat',padx=14,pady=5).pack(side='right')

# ════════════════════════════════════════════════════════════════
# ⑪ GUI — 메인 실행
# ════════════════════════════════════════════════════════════════
def run_gui():
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox
    except ImportError:
        print("[오류] tkinter 없음"); return

    root = tk.Tk()
    root.title("민통선 출입관리체계 — 논리/물리 모델 구성도 생성기 v2.0")
    root.geometry("1000x820")
    root.resizable(True, True)
    root.configure(bg='#F0F0F0')

    sty = ttk.Style(); sty.theme_use('clam')
    sty.configure('TLabel',       background='#F0F0F0', font=(FONT,10))
    sty.configure('TEntry',       font=(FONT,10))
    sty.configure('TButton',      font=(FONT,10,'bold'), padding=6)
    sty.configure('TCheckbutton', background='#F0F0F0', font=(FONT,10))
    sty.configure('TNotebook.Tab',font=(FONT,10,'bold'), padding=(10,4))
    sty.configure('H.TLabel',     background='#2E4057', foreground='white',
                                  font=(FONT,13,'bold'), padding=12)
    sty.configure('Logical.TButton',  font=(FONT,10,'bold'), padding=6,
                  background='#1B4F72', foreground='white')
    sty.configure('Physical.TButton', font=(FONT,10,'bold'), padding=6,
                  background='#4A235A', foreground='white')
    sty.configure('Active.TButton',   font=(FONT,11,'bold'), padding=8,
                  background='#2E4057', foreground='white')
    sty.configure('Inactive.TButton', font=(FONT,11,'bold'), padding=8,
                  background='#AAAAAA', foreground='#555555')

    # ── 헤더 ──
    ttk.Label(root,
              text='  ▣  민통선 출입관리체계  논리/물리 모델 구성도 생성기  v2.0',
              style='H.TLabel').pack(fill='x')

    # ════════════════════════════════════════════════════════
    # 모드 전환 버튼 (논리 모델 / 물리 모델)
    # ════════════════════════════════════════════════════════
    mode_bar = tk.Frame(root, bg='#2E4057', pady=4)
    mode_bar.pack(fill='x')

    current_mode = tk.StringVar(value='logical')

    # 각 모드의 노트북 컨테이너
    nb_container = tk.Frame(root, bg='#F0F0F0')
    nb_container.pack(fill='both', expand=True, padx=10, pady=(4,2))

    logical_frame  = tk.Frame(nb_container, bg='#F0F0F0')
    physical_frame = tk.Frame(nb_container, bg='#F0F0F0')

    # 논리 모델 노트북
    logical_nb = ttk.Notebook(logical_frame)
    logical_nb.pack(fill='both', expand=True)

    # 물리 모델 노트북
    physical_nb = ttk.Notebook(physical_frame)
    physical_nb.pack(fill='both', expand=True)

    # ── 논리 모델 탭들 구축 ──
    log_vars = build_logical_tabs(logical_nb)

    # ── 물리 모델 탭 구축 ──
    phys_vars = build_physical_tab(physical_nb)

    # 첫 화면: 논리 모델
    logical_frame.pack(fill='both', expand=True)

    def _switch_mode(mode):
        current_mode.set(mode)
        if mode == 'logical':
            physical_frame.pack_forget()
            logical_frame.pack(fill='both', expand=True)
            btn_logical.config(style='Active.TButton')
            btn_physical.config(style='Inactive.TButton')
            _update_bottom_buttons('logical')
        else:
            logical_frame.pack_forget()
            physical_frame.pack(fill='both', expand=True)
            btn_logical.config(style='Inactive.TButton')
            btn_physical.config(style='Active.TButton')
            _update_bottom_buttons('physical')

    btn_logical  = ttk.Button(mode_bar, text='  📐 논리 모델  ',
                               style='Active.TButton',
                               command=lambda: _switch_mode('logical'))
    btn_physical = ttk.Button(mode_bar, text='  🗄 물리 모델  ',
                               style='Inactive.TButton',
                               command=lambda: _switch_mode('physical'))
    tk.Label(mode_bar, text='   모드 전환 →  ', bg='#2E4057',
             fg='#AAAAAA', font=(FONT,9)).pack(side='left', padx=(8,0))
    btn_logical.pack(side='left', padx=4)
    btn_physical.pack(side='left', padx=4)

    # ── 상태바 ──
    stat_v = tk.StringVar(value='● 준비')
    ttk.Label(root, textvariable=stat_v,
              foreground='#333', background='#F0F0F0').pack(pady=(2,0))

    # ── 버튼 영역 ──
    bf_outer = tk.Frame(root, bg='#F0F0F0')
    bf_outer.pack(pady=(2,8))

    bf_logical  = tk.Frame(bf_outer, bg='#F0F0F0')
    bf_physical = tk.Frame(bf_outer, bg='#F0F0F0')

    def _update_bottom_buttons(mode):
        if mode == 'logical':
            bf_physical.pack_forget()
            bf_logical.pack()
        else:
            bf_logical.pack_forget()
            bf_physical.pack()

    # ════════════════════════════════════════════════════════
    # ── 논리 모델 콜백 ──
    # ════════════════════════════════════════════════════════
    def on_logical_load():
        path = log_vars['xl_v'].get().strip()
        if not os.path.isfile(path):
            messagebox.showerror('오류', f'파일 없음:\n{path}'); return None
        stat_v.set('● [논리] 데이터 로드 중…'); root.update()
        tree = logical_load_tree(path)
        log_vars['populate_tab2'](list(tree.keys()))
        stat_v.set(f'● [논리] 로드 완료: 대주제 {len(tree)}개')
        return tree

    def on_logical_generate():
        tree = on_logical_load()
        if not tree: return
        if not log_vars['subj_vars']:
            log_vars['populate_tab2'](list(tree.keys()))
        rels = logical_extract_relations(tree)
        stat_v.set('● [논리] PPTX 생성 중…'); root.update()
        try:
            cfg, pos = log_vars['get_cfg']()
            prs = logical_build_pptx(tree, rels, cfg, pos or None)
            out = log_vars['out_v'].get().strip()
            prs.save(out)
            ne = sum(len(ed) for sad in tree.values()
                     for zad in sad.values() for ed in zad.values())
            stat_v.set(f'● [논리] 완료! 엔티티 {ne}개 / 관계 {len(rels)}개')
            messagebox.showinfo('논리 모델 생성 완료',
                f'PPTX 저장 완료!\n\n  파일   : {out}\n'
                f'  엔티티 : {ne}개\n  관계선 : {len(rels)}개')
        except Exception as e:
            import traceback
            messagebox.showerror('오류', f'{e}\n\n{traceback.format_exc()[-600:]}')
            stat_v.set('● 오류')

    def on_logical_save():
        try:
            sp = filedialog.asksaveasfilename(
                title='논리 모델 설정 파일 저장',
                defaultextension='.json', initialfile=LOGICAL_CFG,
                filetypes=[('JSON','*.json'),('모든','*.*')])
            if not sp: return
            cfg, pos = log_vars['get_cfg']()
            data = {'cfg':cfg,'positions':pos,
                    'excel':log_vars['xl_v'].get(),'output':log_vars['out_v'].get()}
            with open(sp,'w',encoding='utf-8') as f:
                json.dump(data,f,ensure_ascii=False,indent=2)
            stat_v.set(f'● [논리] 설정 저장: {sp}')
        except Exception as e:
            messagebox.showerror('오류',f'저장 실패:\n{e}')

    def on_logical_load_cfg():
        path = filedialog.askopenfilename(
            filetypes=[('JSON','*.json'),('모든','*.*')], initialfile=LOGICAL_CFG)
        if not path: return
        try:
            with open(path,'r',encoding='utf-8') as f: data = json.load(f)
            c = data.get('cfg', {})
            for var,key,dv in [
                (log_vars['sw_v'],'slide_w',100),(log_vars['sh_v'],'slide_h',55),
                (log_vars['mg_v'],'margin',0.8),(log_vars['gp_v'],'gap',0.35),
                (log_vars['ew_v'],'entity_w',4.8),(log_vars['eh_v'],'entity_h',1.3),
                (log_vars['aw_v'],'attr_w',4.8),(log_vars['er_v'],'attr_row_h',0.48),
                (log_vars['eam_v'],'attr_min_h',0.48),
                (log_vars['pd_v'],'pad',0.55),(log_vars['za_v'],'hdr_za',0.82),
                (log_vars['sa2_v'],'hdr_sa',0.88),(log_vars['da2_v'],'hdr_da',0.95),
                (log_vars['ttl_v'],'title_h',1.25),
                (log_vars['gda_v'],'da_font_sz',13),(log_vars['gsa_v'],'sa_font_sz',11),
                (log_vars['gza_v'],'za_font_sz',10),(log_vars['gen_v'],'en_font_sz',9),
                (log_vars['gat_v'],'attr_font_sz',7.5),(log_vars['gtoff_v'],'text_offset',0),
                (log_vars['lc_v'],'line_color','404040'),(log_vars['cc_v'],'cross_color','FF6600'),
                (log_vars['en_ml_v'],'en_margin_l',0.12),(log_vars['en_mr_v'],'en_margin_r',0.08),
                (log_vars['en_mt_v'],'en_margin_t',0.06),(log_vars['en_mb_v'],'en_margin_b',0.04),
                (log_vars['en_gh_v'],'en_gap_h',0.35),(log_vars['en_gv_v'],'en_gap_v',0.35),
                (log_vars['en_pl_v'],'en_pad_l',0.55),(log_vars['en_pt_v'],'en_pad_t',0.0),
                (log_vars['en_yoff_v'],'en_y_offset',0.0),
                (log_vars['pk_fsz_v'],'pk_font_sz',7.5),(log_vars['fk_fsz_v'],'fk_font_sz',7.5),
                (log_vars['pk_clr_v'],'pk_color','1F3864'),(log_vars['fk_clr_v'],'fk_color','4A4A4A'),
            ]:
                var.set(str(c.get(key,dv)))
            log_vars['pk_bold_v'].set(bool(c.get('pk_bold',True)))
            log_vars['fk_bold_v'].set(bool(c.get('fk_bold',False)))
            log_vars['show_dtype_v'].set(bool(c.get('show_dtype',True)))
            if data.get('excel'): log_vars['xl_v'].set(data['excel'])
            if data.get('output'): log_vars['out_v'].set(data['output'])
            saved_styles=c.get('styles',{})
            saved_pos=data.get('positions',{})
            da_list=list(saved_styles.keys()) or list(saved_pos.keys())
            if da_list:
                log_vars['populate_tab2'](da_list)
                for da,sv in log_vars['subj_vars'].items():
                    sty2=saved_styles.get(da,{}); pos2=saved_pos.get(da,{})
                    for key,default in [
                        ('da_font_sz','13'),('sa_font_sz','11'),('za_font_sz','10'),
                        ('en_font_sz','9'),('da_fc','000000'),('da_bg','BFBFBF'),
                        ('sa_bg','D6D6D6'),('za_bg','EBEBEB'),
                        ('en_name_bg','D0D8E8'),('text_offset','0'),
                    ]:
                        sv[key].set(str(sty2.get(key,default)))
                    if pos2:
                        sv['use_pos'].set(True)
                        sv['x'].set(str(pos2.get('x','1.0')))
                        sv['y'].set(str(pos2.get('y','2.5')))
                    else:
                        sv['use_pos'].set(False)
            stat_v.set(f'● [논리] 설정 불러옴: {path}')
        except Exception as e:
            messagebox.showerror('오류',f'불러오기 실패:\n{e}')

    # ── 논리 모델 버튼 ──
    logical_btns = [
        ('  데이터 로드  ',       on_logical_load,        '#D0E8FF', '#1B4F72'),
        ('  ▶ 생성  ',            on_logical_generate,    '#C8E6C9', '#1B5E20'),
        ('  💾 설정 저장  ',      on_logical_save,        '#FFF9C4', '#7D6608'),
        ('  📂 설정 불러오기  ', on_logical_load_cfg,    '#F3E5F5', '#4A235A'),
    ]
    for txt,cmd,bg2,fg2 in logical_btns:
        tk.Button(bf_logical,text=txt,command=cmd,
                  font=(FONT,10,'bold'),relief='flat',
                  bg=bg2,fg=fg2,padx=10,pady=6).pack(side='left',padx=4)
    tk.Button(bf_logical,text='  ✕ 종료  ',command=root.destroy,
              font=(FONT,10,'bold'),relief='flat',
              bg='#FFEBEE',fg='#B71C1C',padx=10,pady=6).pack(side='left',padx=4)

    # ════════════════════════════════════════════════════════
    # ── 물리 모델 콜백 ──
    # ════════════════════════════════════════════════════════
    _phys_tables  = [None]; _phys_rels     = [None]
    _phys_parents = [None]; _phys_children = [None]

    def on_phys_load():
        path = phys_vars['phys_xl_v'].get().strip()
        if not os.path.isfile(path):
            messagebox.showerror('오류',f'파일 없음:\n{path}'); return
        stat_v.set('● [물리] 데이터 로드 중…'); root.update()
        try:
            df = phys_load_excel(path)
            tables = phys_parse_tables(df)
            rels   = phys_extract_relations(tables)
            saved_fk = phys_load_fk()
            all_rels = list(rels)
            existing_keys = {(r[0],r[1],r[2],r[3]) for r in rels}
            for r in saved_fk:
                if (r[0],r[1],r[2],r[3]) not in existing_keys:
                    all_rels.append(r)
            phys_inject_fk_cols(tables, all_rels)
            parents, children = phys_detect_parent_child(all_rels)
            _phys_tables[0]  = tables
            _phys_rels[0]    = all_rels
            _phys_parents[0] = parents
            _phys_children[0]= children
            phys_vars['all_tbls_cache'][0] = {'tables':tables,'rels':all_rels}
            phys_vars['refresh_tbl_vars'](tables)
            stat_v.set(f'● [물리] 로드 완료: 엔티티 {len(tables)}개 / FK관계 {len(all_rels)}개')
        except Exception as e:
            import traceback
            messagebox.showerror('오류',f'{e}\n\n{traceback.format_exc()[-400:]}')
            stat_v.set('● 오류')

    def on_phys_generate():
        if _phys_tables[0] is None:
            on_phys_load()
        if _phys_tables[0] is None: return
        tables  = _phys_tables[0]
        all_rels= _phys_rels[0]
        parents = _phys_parents[0]
        children= _phys_children[0]

        pv2 = phys_vars
        mode = pv2['mode_var'].get()
        mc   = pv2['max_cols_var'].get()

        opts = {'mode':mode,'max_cols':mc,
                'tbl_w':pv2['tbl_w_var'].get(),'h_gap':pv2['h_gap_var'].get(),
                'v_gap':pv2['v_gap_var'].get(),'ri_h_spc':pv2['ri_h_var'].get(),
                'ri_v_spc':pv2['ri_v_var'].get()}

        if mode == "areas":
            pv2['_save_area2']()
            sa_data2 = pv2['sa_data']
            sa_tree2 = pv2['sa_tree']

            def _collect(iid):
                d=sa_data2.get(iid,{})
                return {"title":d.get("title",""),"level":d.get("level",1),
                        "tables":list(d.get("tables",set())),
                        "children":[_collect(c) for c in sa_tree2.get_children(iid)]}

            areas=[_collect(iid) for iid in sa_tree2.get_children()]

            def _build_l1(l1_node):
                color_map={}; all_tbls2=[]; sub_labels=[]; pidx=[0]
                def _gather(node,hex_c):
                    for t in node["tables"]:
                        if t not in color_map:
                            color_map[t]=hex_c; all_tbls2.append(t)
                    for child in node["children"]:
                        cidx=pidx[0]%len(PALETTE); pidx[0]+=1
                        sub_labels.append(f"L{child['level']} {child['title']}")
                        _gather(child,PALETTE[cidx])
                l1_hex=PALETTE[pidx[0]%len(PALETTE)]; pidx[0]+=1
                _gather(l1_node,l1_hex)
                return {"title":l1_node["title"],"tables":all_tbls2,
                        "table_colors":color_map,"sub_areas":sub_labels,
                        "incl_rel":pv2['incl_rel_var'].get()}

            slides=[_build_l1(l1) for l1 in areas]
            slides=[s for s in slides if s["tables"]]
            if not slides:
                messagebox.showwarning("주제영역 없음","L1 주제영역을 추가하고 엔티티를 할당하세요.")
                return
            opts['area_slides']=slides

        elif mode == "tables":
            sel_tbls = [t for t,v in pv2['a_tbl_vars'].items() if v.get()]
            if not sel_tbls:
                messagebox.showwarning("엔티티 없음","엔티티를 선택하세요."); return
            opts['table_slides']=[{"name":"슬라이드 1","tables":sel_tbls}]

        stat_v.set('● [물리] PPTX 생성 중…'); root.update()
        try:
            out = pv2['phys_out_v'].get().strip()
            phys_export_pptx(tables, all_rels, opts, parents, children, out)
            stat_v.set(f'● [물리] 완료! → {out}')
            messagebox.showinfo('물리 모델 생성 완료',
                f'PPTX 저장 완료!\n\n  파일: {out}\n  엔티티: {len(tables)}개\n  FK관계: {len(all_rels)}개')
        except Exception as e:
            import traceback
            messagebox.showerror('오류',f'{e}\n\n{traceback.format_exc()[-600:]}')
            stat_v.set('● 오류')

    def on_phys_fk_edit():
        if _phys_tables[0] is None:
            messagebox.showwarning("로드 필요","먼저 물리 모델 엑셀 파일을 로드하세요."); return
        tables = _phys_tables[0]; all_rels2 = _phys_rels[0]
        _show_fk_dialog(root, tables, all_rels2,
                        on_done=lambda rels: _after_fk_edit(rels,tables))

    def _after_fk_edit(new_rels, tables):
        phys_inject_fk_cols(tables, new_rels)
        parents2, children2 = phys_detect_parent_child(new_rels)
        _phys_rels[0]    = new_rels
        _phys_parents[0] = parents2
        _phys_children[0]= children2
        phys_vars['all_tbls_cache'][0]={'tables':tables,'rels':new_rels}
        stat_v.set(f'● [물리] FK 관계 업데이트: {len(new_rels)}개')

    # ── 물리 모델 설정 저장 ──
    def on_phys_save_cfg():
        try:
            sp = filedialog.asksaveasfilename(
                title='물리 모델 설정 파일 저장',
                defaultextension='.json', initialfile=PHYSICAL_CFG,
                filetypes=[('JSON','*.json'),('모든','*.*')])
            if not sp: return
            pv2 = phys_vars
            sa_tree2 = pv2['sa_tree']
            sa_data2 = pv2['sa_data']
            def _tree_dump(iid):
                d = sa_data2.get(iid, {})
                return {
                    'title':    d.get('title',''),
                    'level':    d.get('level',1),
                    'tables':   list(d.get('tables',set())),
                    'children': [_tree_dump(c) for c in sa_tree2.get_children(iid)]
                }
            data = {
                'max_cols':  pv2['max_cols_var'].get(),
                'tbl_w':     pv2['tbl_w_var'].get(),
                'h_gap':     pv2['h_gap_var'].get(),
                'v_gap':     pv2['v_gap_var'].get(),
                'ri_h_spc':  pv2['ri_h_var'].get(),
                'ri_v_spc':  pv2['ri_v_var'].get(),
                'excel':     pv2['phys_xl_v'].get(),
                'output':    pv2['phys_out_v'].get(),
                'mode':      pv2['mode_var'].get(),
                'incl_rel':  pv2['incl_rel_var'].get(),
                'area_tree': [_tree_dump(iid) for iid in sa_tree2.get_children()],
            }
            with open(sp,'w',encoding='utf-8') as f:
                json.dump(data,f,ensure_ascii=False,indent=2)
            stat_v.set(f'● [물리] 설정 저장: {sp}')
        except Exception as e:
            messagebox.showerror('오류',f'저장 실패:\n{e}')

    # ── 물리 모델 설정 불러오기 ──
    def on_phys_load_cfg():
        path = filedialog.askopenfilename(
            filetypes=[('JSON','*.json'),('모든','*.*')], initialfile=PHYSICAL_CFG)
        if not path: return
        try:
            with open(path,'r',encoding='utf-8') as f: data = json.load(f)
            pv2 = phys_vars
            if 'max_cols' in data:
                try: pv2['max_cols_var'].set(int(data['max_cols']))
                except: pass
            for key, var in [('tbl_w','tbl_w_var'),('h_gap','h_gap_var'),
                              ('v_gap','v_gap_var'),('ri_h_spc','ri_h_var'),
                              ('ri_v_spc','ri_v_var')]:
                if key in data:
                    try: pv2[var].set(float(data[key]))
                    except: pass
            if data.get('excel'):  pv2['phys_xl_v'].set(data['excel'])
            if data.get('output'): pv2['phys_out_v'].set(data['output'])
            if data.get('mode'):   pv2['mode_var'].set(data['mode'])
            if 'incl_rel' in data: pv2['incl_rel_var'].set(bool(data['incl_rel']))
            # 주제영역 트리 복원
            sa_tree2 = pv2['sa_tree']
            sa_data2 = pv2['sa_data']
            for iid in list(sa_tree2.get_children()): sa_tree2.delete(iid)
            sa_data2.clear()
            def _tree_load(parent_iid, node):
                iid = sa_tree2.insert(parent_iid, 'end',
                                      text=f"  {node.get('title','')}",
                                      values=(f"L{node.get('level',1)}","0"))
                sa_data2[iid] = {
                    'title':  node.get('title',''),
                    'level':  node.get('level',1),
                    'tables': set(node.get('tables',[])),
                }
                if parent_iid: sa_tree2.item(parent_iid, open=True)
                for child in node.get('children',[]):
                    _tree_load(iid, child)
                n_t = len(sa_data2[iid]['tables'])
                lv_t = sa_data2[iid]['level']
                sa_tree2.item(iid, values=(f"L{lv_t}", str(n_t)))
            for node in data.get('area_tree', []):
                _tree_load('', node)
            stat_v.set(f'● [물리] 설정 불러옴: {path}')
        except Exception as e:
            messagebox.showerror('오류',f'불러오기 실패:\n{e}')

    # ── 물리 모델 버튼 ──
    physical_btns = [
        ('  데이터 로드  ',   on_phys_load,       '#D0E8FF', '#1B4F72'),
        ('  FK 관계 편집  ', on_phys_fk_edit,    '#F3E5F5', '#4A235A'),
        ('  ▶ 생성  ',        on_phys_generate,   '#C8E6C9', '#1B5E20'),
        ('  설정 저장  ',     on_phys_save_cfg,   '#FFF9C4', '#7D6608'),
        ('  설정 불러오기  ', on_phys_load_cfg,   '#FCE4EC', '#880E4F'),
    ]
    for txt,cmd,bg2,fg2 in physical_btns:
        tk.Button(bf_physical,text=txt,command=cmd,
                  font=(FONT,10,'bold'),relief='flat',
                  bg=bg2,fg=fg2,padx=10,pady=6).pack(side='left',padx=4)
    tk.Button(bf_physical,text='  ✕ 종료  ',command=root.destroy,
              font=(FONT,10,'bold'),relief='flat',
              bg='#FFEBEE',fg='#B71C1C',padx=10,pady=6).pack(side='left',padx=4)

    # 첫 화면 버튼 표시
    bf_logical.pack()

    root.mainloop()


# ════════════════════════════════════════════════════════════════
# 진입점
# ════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    run_gui()
