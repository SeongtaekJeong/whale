#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
민통선 출입관리체계 — 개념 데이터 모델 구성도 생성기 v2.0
===========================================================
기능:
  • 엔티티 박스 내 PK / FK 속성 표시
  • 삼발이(Crow's Foot) 곡선 관계선
  • 교차 관계선 색상 자동 변경
  • 주제영역별 글꼴·색상·좌표·텍스트 오프셋 설정
  • 설정 저장 / 불러오기 (JSON)
  • GUI 입력 창 (tkinter)

실행:
  [GUI] python 개념모델_구성도_생성기.py
  [CLI] python 개념모델_구성도_생성기.py -i 민통선_엔티티속성정의서.xlsx -o 출력.pptx

pip install python-pptx pandas openpyxl
"""
import sys, os, json, math, argparse
from collections import OrderedDict, defaultdict

try:
    import pandas as pd
    from pptx import Presentation
    from pptx.util import Cm, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError as e:
    print(f"[오류] {e}\npip install python-pptx pandas openpyxl")
    sys.exit(1)

# ════════════════════════════════════════════════════════════════════
# 전역 상수
# ════════════════════════════════════════════════════════════════════
FONT          = "맑은 고딕"
SETTINGS_FILE = "개념모델_설정.json"

# 기본 색상 (연한 회색 계열)
CLR = dict(
    da_bg='BFBFBF', da_bd='404040',   # 대주제영역
    sa_bg='D6D6D6', sa_bd='585858',   # 상위주제영역
    za_bg='EBEBEB', za_bd='787878',   # 주제영역
    en_bg='FFFFFF', en_bd='606060',   # 엔티티
    pk_txt='1F3864',                  # PK 속성 글꼴 (남색)
    fk_txt='4A4A4A',                  # FK 속성 글꼴
    sep='AAAAAA',                     # 구분선
    line='404040',                    # 관계선
    cross='FF6600',                   # 교차 관계선 (주황)
    title_bg='2E4057', title_txt='FFFFFF',
)

# FK 자동 감지 시 제외할 공통 감사 속성명
AUDIT_ATTRS = {
    '등록자ID','수정자ID','등록일시','수정일시','등록일자','수정일자',
    '처리일시','처리자ID','처리일','작성자ID','삭제여부','사용여부',
}

def rgb(h):
    h = h.lstrip('#')
    return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

def emu(v_cm):
    return int(Cm(float(v_cm)))

# ════════════════════════════════════════════════════════════════════
# 1. 데이터 로드
# ════════════════════════════════════════════════════════════════════
def load_tree(excel_path):
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

        if not da or not en or att in ('','nan'):
            continue

        (tree.setdefault(da, OrderedDict())
             .setdefault(sa, OrderedDict())
             .setdefault(za, OrderedDict())
             .setdefault(en, {'pk':[], 'fk':[], 'all':[]}))

        e = tree[da][sa][za][en]
        if att not in e['all']:
            e['all'].append(att)
        if pk == 'Y' and att not in e['pk']:
            e['pk'].append(att)
        if fk == 'Y' and att not in e['fk']:
            e['fk'].append(att)

    return tree


def extract_relations(tree, extra_rels=None):
    """
    PK 속성명 매칭으로 엔티티 간 관계 추출
    extra_rels: [(parent_en, pk_attr, child_en, fk_attr), ...]  수동 추가 관계
    """
    # pk_attr → [(da,sa,za,en)]
    pk_map = defaultdict(list)
    for da,sad in tree.items():
        for sa,zad in sad.items():
            for za,end in zad.items():
                for en,att in end.items():
                    for pk in att['pk']:
                        pk_map[pk].append((da,sa,za,en))

    # entity_name → (da,sa,za)
    en_loc = {}
    for da,sad in tree.items():
        for sa,zad in sad.items():
            for za,end in zad.items():
                for en in end:
                    en_loc[en] = (da,sa,za)

    rels, seen = [], set()

    def add_rel(par_en, fk_attr, chi_en):
        key = (par_en, chi_en, fk_attr)
        if key in seen or par_en == chi_en:
            return
        seen.add(key)
        pda,psa,pza = en_loc.get(par_en, ('','',''))
        cda,csa,cza = en_loc.get(chi_en, ('','',''))
        rels.append(dict(
            par=par_en, par_da=pda, par_sa=psa, par_za=pza,
            chi=chi_en, chi_da=cda, chi_sa=csa, chi_za=cza,
            attr=fk_attr
        ))

    # ── 자동 감지: PK 속성명이 다른 엔티티 속성 목록에 있으면 관계로 간주 ──
    for da,sad in tree.items():
        for sa,zad in sad.items():
            for za,end in zad.items():
                for chi_en, chi_att in end.items():
                    for attr in chi_att['all']:
                        if attr in AUDIT_ATTRS:
                            continue
                        if attr in chi_att['pk']:
                            continue  # PK는 자신의 PK → 제외
                        for par_da,par_sa,par_za,par_en in pk_map.get(attr, []):
                            add_rel(par_en, attr, chi_en)

    # ── 수동 추가 관계 ─────────────────────────────────────────────
    if extra_rels:
        for par_en, pk_attr, chi_en, fk_attr in extra_rels:
            add_rel(par_en, fk_attr, chi_en)

    return rels

# ════════════════════════════════════════════════════════════════════
# 2. 레이아웃 크기 계산
# ════════════════════════════════════════════════════════════════════
class LC:
    """cm 입력 → Emu 변환"""
    def __init__(self, c):
        g = lambda k,d: c.get(k,d)
        self.SW       = emu(g('slide_w',    100))
        self.SH       = emu(g('slide_h',     55))
        self.MARG     = emu(g('margin',      0.8))
        self.GAP      = emu(g('gap',         0.35))
        self.EW       = emu(g('entity_w',    4.8))
        self.EH_BASE  = emu(g('entity_h',    1.3))
        self.EH_ROW   = emu(g('attr_row_h',  0.48))
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


def en_height(attrs, L):
    rows = 1 + len(attrs['pk']) + len(attrs['fk'])
    return max(L.EH_BASE, L.EH_BASE + max(0,(rows-1)) * L.EH_ROW)


def za_wh(en_dict, L, cols=2):
    n = max(len(en_dict), 1)
    c = min(n, cols)
    r = math.ceil(n / c)
    max_eh = max(en_height(a,L) for a in en_dict.values()) if en_dict else L.EH_BASE
    w = c*L.EW + (c-1)*L.GAP + 2*L.PAD
    h = L.HDR_ZA + r*max_eh + max(0,r-1)*L.GAP + 2*L.PAD
    return w, h


def sa_wh(za_dict, L):
    max_w, total_h = 0, L.HDR_SA + L.PAD
    for za,end in za_dict.items():
        zw,zh = za_wh(end, L)
        max_w = max(max_w, zw)
        total_h += zh + L.GAP
    total_h -= L.GAP
    return max_w + 2*L.PAD, total_h + L.PAD


def da_wh(sa_dict, L):
    total_w, max_h = L.PAD, 0
    for sa,zad in sa_dict.items():
        sw,sh = sa_wh(zad, L)
        total_w += sw + L.GAP
        max_h = max(max_h, sh)
    total_w -= L.GAP
    return total_w + L.PAD, L.HDR_DA + max_h + 2*L.PAD

# ════════════════════════════════════════════════════════════════════
# 3. PPTX 도형 그리기
# ════════════════════════════════════════════════════════════════════
def add_rect(slide, x, y, w, h, bg, bd, text='',
             fsz=10, bold=False, fc='000000',
             bw=Pt(0.75), va='top', ha=PP_ALIGN.CENTER, toff=0):
    """편집 가능한 직사각형 추가"""
    s = slide.shapes.add_shape(1, int(x), int(y), int(w), int(h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(bg)
    s.line.color.rgb = rgb(bd)
    s.line.width = bw
    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left   = emu(0.1)
    tf.margin_right  = emu(0.1)
    tf.margin_top    = emu(0.05) + int(toff)
    tf.margin_bottom = emu(0.05)
    tf.vertical_anchor = {'top':1,'middle':3,'bottom':4}[va]
    if text:
        p = tf.paragraphs[0]; p.alignment = ha
        r = p.add_run()
        r.text = text; r.font.name = FONT
        r.font.size = Pt(fsz); r.font.bold = bold
        r.font.color.rgb = rgb(fc)
    return s


def add_entity_box(slide, x, y, w, en_name, pk_list, fk_list, L,
                   sty, toff=0):
    """엔티티 박스 (이름 + PK/FK 속성)"""
    attrs   = {'pk': pk_list, 'fk': fk_list, 'all': pk_list + fk_list}
    box_h   = en_height(attrs, L)
    en_bg   = sty.get('en_bg',   CLR['en_bg'])
    en_bd   = sty.get('en_bd',   CLR['en_bd'])
    en_nfc  = sty.get('en_name_fc', '000000')

    s = slide.shapes.add_shape(1, int(x), int(y), int(w), int(box_h))
    s.fill.solid()
    s.fill.fore_color.rgb = rgb(en_bg)
    s.line.color.rgb = rgb(en_bd)
    s.line.width = Pt(0.75)

    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left   = emu(0.12)
    tf.margin_right  = emu(0.08)
    tf.margin_top    = emu(0.06) + int(toff)
    tf.margin_bottom = emu(0.04)
    tf.vertical_anchor = 1   # top

    # 엔티티명
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = en_name
    r.font.name = FONT; r.font.size = Pt(L.EN_FSZ)
    r.font.bold = True; r.font.color.rgb = rgb(en_nfc)

    # 구분선
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.LEFT
    r2 = p2.add_run(); r2.text = '─' * 14
    r2.font.name = FONT; r2.font.size = Pt(5.5)
    r2.font.color.rgb = rgb(CLR['sep'])

    # PK
    for pk in pk_list:
        p = tf.add_paragraph(); p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = f'▶ {pk}'
        r.font.name = FONT; r.font.size = Pt(L.AT_FSZ)
        r.font.bold = True; r.font.color.rgb = rgb(CLR['pk_txt'])

    # FK
    for fk in fk_list:
        p = tf.add_paragraph(); p.alignment = PP_ALIGN.LEFT
        r = p.add_run(); r.text = f'◆ {fk}'
        r.font.name = FONT; r.font.size = Pt(L.AT_FSZ)
        r.font.bold = False; r.font.color.rgb = rgb(CLR['fk_txt'])

    return s, box_h

# ── 엣지 포인트 계산 ─────────────────────────────────────────────
def edge_pt(ex, ey, ew, eh, tx, ty):
    """박스(ex,ey,ew,eh) 에서 대상(tx,ty) 방향의 엣지 점"""
    cx, cy = ex + ew//2, ey + eh//2
    dx, dy = tx - cx, ty - cy
    if abs(dx) < 1 and abs(dy) < 1:
        return cx, cy
    hw, hh = ew//2, eh//2
    if abs(dx) < 1:
        return cx, ey if dy < 0 else ey + eh
    sl = dy / dx
    if abs(sl) * hw < hh:
        if dx > 0:
            return ex+ew, int(cy + sl*hw)
        else:
            return ex,    int(cy - sl*hw)
    else:
        if dy > 0:
            return int(cx + hh/sl), ey+eh
        else:
            return int(cx - hh/sl), ey

# ── 삼발이 관계선 ──────────────────────────────────────────────
def draw_crowfoot(slide, x1, y1, x2, y2, color='404040'):
    """
    곡선 커넥터 + crow's foot 마커
    Parent (x1,y1): 단일 틱(|)
    Child  (x2,y2): 삼발이(>|)
    """
    lc = rgb(color)
    lw = Pt(1.25)
    shapes = []

    def line(ax, ay, bx, by, clr=None, width=None):
        try:
            c = slide.shapes.add_connector(1,
                int(ax), int(ay), int(bx), int(by))
            c.line.color.rgb = clr or lc
            c.line.width = width or lw
            shapes.append(c)
        except Exception:
            pass

    def curve(ax, ay, bx, by):
        try:
            c = slide.shapes.add_connector(3,
                int(ax), int(ay), int(bx), int(by))
            c.line.color.rgb = lc
            c.line.width = lw
            shapes.append(c)
            return c
        except Exception:
            line(ax, ay, bx, by)
            return None

    # ── 방향 벡터 ────────────────────────────────────────────────
    dx, dy = x2 - x1, y2 - y1
    length = math.sqrt(dx*dx + dy*dy)
    if length < emu(0.1):
        return shapes

    ux, uy = dx/length, dy/length     # 단위 방향
    px, py = -uy, ux                  # 수직(법선)

    TL = emu(0.3)    # 틱 반길이
    CL = emu(0.42)   # 삼발이 길이
    CS = emu(0.22)   # 삼발이 반폭

    # ── 메인 곡선 커넥터 ────────────────────────────────────────
    curve(x1, y1, x2, y2)

    # ── Parent 끝: 단일 수직 틱 ─────────────────────────────────
    line(x1 + px*TL, y1 + py*TL,
         x1 - px*TL, y1 - py*TL)

    # ── Child 끝: 삼발이 ────────────────────────────────────────
    # 백 포인트 (Child 에서 안쪽으로)
    bx = x2 - ux*CL
    by = y2 - uy*CL

    # 중앙선
    line(bx, by, x2, y2)
    # 왼쪽 발
    line(bx, by, x2 - px*CS, y2 - py*CS)
    # 오른쪽 발
    line(bx, by, x2 + px*CS, y2 + py*CS)
    # 틱 (필수 마크)
    line(bx + px*TL, by + py*TL,
         bx - px*TL, by - py*TL)

    return shapes

# ── 선분 교차 감지 ────────────────────────────────────────────
def segs_cross(p1, p2, p3, p4):
    """선분 p1p2 와 p3p4 의 내부 교차 여부"""
    def c2d(a, b): return a[0]*b[1] - a[1]*b[0]
    def sub(a, b): return (a[0]-b[0], a[1]-b[1])
    r = sub(p2,p1); s = sub(p4,p3)
    rxs = c2d(r,s)
    if abs(rxs) < 1e-9:
        return False
    t = c2d(sub(p3,p1), s) / rxs
    u = c2d(sub(p3,p1), r) / rxs
    return 0.05 < t < 0.95 and 0.05 < u < 0.95

# ════════════════════════════════════════════════════════════════════
# 4. 메인 PPTX 빌더
# ════════════════════════════════════════════════════════════════════
def build_pptx(tree, relations, cfg, positions=None):
    L   = LC(cfg)
    prs = Presentation()
    prs.slide_width  = Emu(L.SW)
    prs.slide_height = Emu(L.SH)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    en_pos   = {}   # en_name → (x, y, w, h)
    line_segs = []  # [(p1,p2, shapes_list)]
    styles    = cfg.get('styles', {})

    # ── 슬라이드 제목 ─────────────────────────────────────────────
    add_rect(slide, L.MARG, L.MARG,
             L.SW - 2*L.MARG, L.TTL,
             CLR['title_bg'], '1A2540',
             text='민통선 출입관리체계 — 개념 데이터 모델',
             fsz=14, bold=True, fc=CLR['title_txt'],
             bw=Pt(2.0), va='middle')

    cur_x = L.MARG
    cur_y = L.MARG + L.TTL + L.GAP

    # ── 대주제영역 순회 ───────────────────────────────────────────
    for da, sa_dict in tree.items():
        sty = styles.get(da, {})
        pos = (positions or {}).get(da, {})

        if pos:
            cur_x = emu(pos.get('x', 1.0))
            cur_y = emu(pos.get('y', 2.5))

        dw, dh = da_wh(sa_dict, L)
        toff_v = emu(sty.get('text_offset', 0))

        # 대주제영역 박스
        add_rect(slide, cur_x, cur_y, dw, dh,
                 sty.get('da_bg', CLR['da_bg']),
                 sty.get('da_bd', CLR['da_bd']),
                 text=da,
                 fsz=float(sty.get('da_font_sz', L.DA_FSZ)),
                 bold=True, fc=sty.get('da_fc','000000'),
                 bw=Pt(2.25), va='top', toff=toff_v)

        sa_x = cur_x + L.PAD

        for sa, za_dict in sa_dict.items():
            sw, sh = sa_wh(za_dict, L)
            sa_y = cur_y + L.HDR_DA + L.PAD

            # 상위주제영역 박스
            add_rect(slide, sa_x, sa_y, sw, sh,
                     sty.get('sa_bg', CLR['sa_bg']),
                     sty.get('sa_bd', CLR['sa_bd']),
                     text=sa,
                     fsz=float(sty.get('sa_font_sz', L.SA_FSZ)),
                     bold=True, fc=sty.get('sa_fc','000000'),
                     bw=Pt(1.5), va='top', toff=toff_v)

            za_x = sa_x + L.PAD
            za_y = sa_y + L.HDR_SA + L.PAD

            for za, en_dict in za_dict.items():
                zdw = sw - 2*L.PAD
                zw, zh = za_wh(en_dict, L)

                # 주제영역 박스
                add_rect(slide, za_x, za_y, zdw, zh,
                         sty.get('za_bg', CLR['za_bg']),
                         sty.get('za_bd', CLR['za_bd']),
                         text=za,
                         fsz=float(sty.get('za_font_sz', L.ZA_FSZ)),
                         bold=False, fc=sty.get('za_fc','000000'),
                         bw=Pt(1.0), va='top', toff=toff_v)

                # ── 엔티티 배치 ───────────────────────────────────
                cols = 2 if len(en_dict) > 1 else 1
                ex0  = za_x + L.PAD
                ey0  = za_y + L.HDR_ZA + L.PAD // 2

                for idx, (en, attrs) in enumerate(en_dict.items()):
                    col = idx % cols
                    row = idx // cols

                    # 행 Y 누적
                    ey_row = ey0
                    for r in range(row):
                        row_ens = [list(en_dict.values())[ri]
                                   for ri in range(r*cols, min((r+1)*cols, len(en_dict)))]
                        ey_row += max(en_height(a,L) for a in row_ens) + L.GAP

                    ex_pos = ex0 + col * (L.EW + L.GAP)

                    _, actual_h = add_entity_box(
                        slide, ex_pos, ey_row, L.EW,
                        en, attrs['pk'], attrs['fk'],
                        L, sty, toff=toff_v
                    )
                    en_pos[en] = (ex_pos, ey_row, L.EW, actual_h)

                za_y += zh + L.GAP

            sa_x += sw + L.GAP

        cur_x += dw + L.GAP * 2

    # ── 관계선 그리기 ─────────────────────────────────────────────
    line_clr  = cfg.get('line_color',  CLR['line'])
    cross_clr = cfg.get('cross_color', CLR['cross'])

    for rel in relations:
        pen, cen = rel['par'], rel['chi']
        if pen not in en_pos or cen not in en_pos:
            continue

        px2, py2, pw, ph = en_pos[pen]
        cx2, cy2, cw, ch = en_pos[cen]

        pcx, pcy = px2 + pw//2, py2 + ph//2
        ccx, ccy = cx2 + cw//2, cy2 + ch//2

        x1, y1 = edge_pt(px2, py2, pw, ph, ccx, ccy)
        x2, y2 = edge_pt(cx2, cy2, cw, ch, pcx, pcy)

        shps = draw_crowfoot(slide, x1, y1, x2, y2, line_clr)
        line_segs.append(((x1,y1),(x2,y2), shps))

    # ── 교차 감지 → 색상 변경 ────────────────────────────────────
    n = len(line_segs)
    crossed = set()
    for i in range(n):
        for j in range(i+1, n):
            (p1,p2,_) = line_segs[i]
            (p3,p4,_) = line_segs[j]
            if segs_cross(p1,p2,p3,p4):
                crossed.add(i); crossed.add(j)

    for idx in crossed:
        _,_,shps = line_segs[idx]
        for sh in shps:
            try:
                sh.line.color.rgb = rgb(cross_clr)
            except Exception:
                pass

    # ── 범례 ──────────────────────────────────────────────────────
    leg_items = [
        ('■ 대주제영역',   CLR['da_bg'], CLR['da_bd']),
        ('■ 상위주제영역', CLR['sa_bg'], CLR['sa_bd']),
        ('■ 주제영역',     CLR['za_bg'], CLR['za_bd']),
        ('■ 엔티티',       CLR['en_bg'], CLR['en_bd']),
    ]
    lx = L.MARG
    ly = L.SH - L.MARG - emu(0.72)
    for lbl, bg, bd in leg_items:
        add_rect(slide, lx, ly, emu(3.3), emu(0.65),
                 bg, bd, text=lbl, fsz=8, bw=Pt(0.5), va='middle')
        lx += emu(3.5)

    return prs

# ════════════════════════════════════════════════════════════════════
# 5. GUI
# ════════════════════════════════════════════════════════════════════
def run_gui():
    try:
        import tkinter as tk
        from tkinter import ttk, filedialog, messagebox, colorchooser
    except ImportError:
        print("[오류] tkinter 없음 — CLI 모드로 실행하세요")
        return

    root = tk.Tk()
    root.title("민통선 개념모델 구성도 생성기 v2.0")
    root.geometry("900x740")
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

    ttk.Label(root, text='  ▣  민통선 개념모델 구성도 생성기  v2.0',
              style='H.TLabel').pack(fill='x')

    nb = ttk.Notebook(root)
    nb.pack(fill='both', expand=True, padx=10, pady=(8,4))

    # ─────────────────────────────────────────────────────────────
    # 탭 1: 기본 설정
    # ─────────────────────────────────────────────────────────────
    tab1 = ttk.Frame(nb); nb.add(tab1, text='  기본 설정  ')

    def lf(parent, title, pady=(4,4)):
        f = ttk.LabelFrame(parent, text=f' {title} ', padding=8)
        f.pack(fill='x', padx=10, pady=pady)
        return f

    def row_kv(f, label, var, unit='', r=None, c=0):
        ri = f.grid_size()[1] if r is None else r
        ttk.Label(f, text=label).grid(row=ri, column=c, sticky='w', pady=2)
        ttk.Entry(f, textvariable=var, width=9).grid(row=ri, column=c+1, padx=6, sticky='w')
        if unit:
            ttk.Label(f, text=unit, foreground='#666').grid(row=ri, column=c+2, sticky='w')

    # 파일
    ff = lf(tab1, '파일')
    xl_v  = tk.StringVar(value='민통선_엔티티속성정의서.xlsx')
    out_v = tk.StringVar(value='민통선_개념모델.pptx')
    for row_i, (lbl, var, cmd_fn) in enumerate([
        ('입력 파일', xl_v,  lambda: xl_v.set(filedialog.askopenfilename(
            filetypes=[('Excel','*.xlsx *.xls'),('모든','*.*')]) or xl_v.get())),
        ('출력 파일', out_v, lambda: out_v.set(filedialog.asksaveasfilename(
            defaultextension='.pptx',
            filetypes=[('PowerPoint','*.pptx'),('모든','*.*')],
            initialfile=out_v.get()) or out_v.get())),
    ]):
        ttk.Label(ff, text=lbl).grid(row=row_i, column=0, sticky='w', pady=2)
        ttk.Entry(ff, textvariable=var, width=44).grid(row=row_i, column=1, padx=6)
        ttk.Button(ff, text='찾아보기', command=cmd_fn).grid(row=row_i, column=2, padx=4)

    # 슬라이드 크기
    sf = lf(tab1, '슬라이드 크기'); sf.columnconfigure(5, weight=1)
    sw_v=tk.StringVar(value='100');  sh_v=tk.StringVar(value='55')
    mg_v=tk.StringVar(value='0.8');  gp_v=tk.StringVar(value='0.35')
    row_kv(sf, '너비', sw_v, 'cm', r=0, c=0)
    row_kv(sf, '높이', sh_v, 'cm', r=1, c=0)
    row_kv(sf, '여백', mg_v, 'cm', r=0, c=4)
    row_kv(sf, '간격', gp_v, 'cm', r=1, c=4)

    # 엔티티 박스
    ef = lf(tab1, '엔티티 박스'); ef.columnconfigure(5, weight=1)
    ew_v=tk.StringVar(value='4.8');  eh_v=tk.StringVar(value='1.3')
    er_v=tk.StringVar(value='0.48')
    row_kv(ef, '엔티티 너비', ew_v, 'cm', r=0, c=0)
    row_kv(ef, '기본 높이',   eh_v, 'cm', r=1, c=0)
    row_kv(ef, '속성 행 높이', er_v,'cm', r=0, c=4)

    # 레이아웃 여백
    lf2 = lf(tab1, '레이아웃 여백'); lf2.columnconfigure(5, weight=1)
    pd_v=tk.StringVar(value='0.55'); za_v=tk.StringVar(value='0.82')
    sa2_v=tk.StringVar(value='0.88'); da2_v=tk.StringVar(value='0.95')
    ttl_v=tk.StringVar(value='1.25')
    row_kv(lf2, '내부 패딩',     pd_v, 'cm', r=0, c=0)
    row_kv(lf2, '주제영역 H',    za_v, 'cm', r=1, c=0)
    row_kv(lf2, '상위주제 H',   sa2_v,'cm', r=0, c=4)
    row_kv(lf2, '대주제 H',     da2_v,'cm', r=1, c=4)
    row_kv(lf2, '슬라이드 제목 H', ttl_v,'cm', r=2, c=0)

    # 관계선
    rf = lf(tab1, '관계선'); rf.columnconfigure(5, weight=1)
    lc_v=tk.StringVar(value='404040');  cc_v=tk.StringVar(value='FF6600')
    ttk.Label(rf, text='관계선 색상 (HEX)').grid(row=0, column=0, sticky='w', pady=2)
    ttk.Entry(rf, textvariable=lc_v, width=10).grid(row=0, column=1, padx=6)
    ttk.Label(rf, text='교차선 색상 (HEX)').grid(row=0, column=3, sticky='w', pady=2)
    ttk.Entry(rf, textvariable=cc_v, width=10).grid(row=0, column=4, padx=6)

    # ─────────────────────────────────────────────────────────────
    # 탭 2: 주제영역별 설정
    # ─────────────────────────────────────────────────────────────
    tab2 = ttk.Frame(nb); nb.add(tab2, text='  주제영역별 설정  ')

    cnv2 = tk.Canvas(tab2, bg='#F0F0F0', highlightthickness=0)
    sb2  = ttk.Scrollbar(tab2, orient='vertical', command=cnv2.yview)
    cnv2.configure(yscrollcommand=sb2.set)
    sb2.pack(side='right', fill='y')
    cnv2.pack(side='left', fill='both', expand=True)
    inner2 = tk.Frame(cnv2, bg='#F0F0F0')
    cnv2.create_window((0,0), window=inner2, anchor='nw')
    inner2.bind('<Configure>',
                lambda e: cnv2.configure(scrollregion=cnv2.bbox('all')))

    subj_vars = {}   # {da: {key: StringVar/BooleanVar}}

    def populate_tab2(da_list):
        for w in inner2.winfo_children():
            w.destroy()
        subj_vars.clear()
        ttk.Label(inner2, text='대주제영역별로 위치·글꼴·배경색·텍스트 오프셋을 설정하세요.',
                  foreground='#555').pack(pady=(6,4), padx=10, anchor='w')

        for da in da_list:
            sv = {
                'use_pos':    tk.BooleanVar(value=False),
                'x':          tk.StringVar(value='1.0'),
                'y':          tk.StringVar(value='2.5'),
                'da_font_sz': tk.StringVar(value='13'),
                'sa_font_sz': tk.StringVar(value='11'),
                'za_font_sz': tk.StringVar(value='10'),
                'en_font_sz': tk.StringVar(value='9'),
                'da_fc':      tk.StringVar(value='000000'),
                'da_bg':      tk.StringVar(value='BFBFBF'),
                'sa_bg':      tk.StringVar(value='D6D6D6'),
                'za_bg':      tk.StringVar(value='EBEBEB'),
                'text_offset':tk.StringVar(value='0'),
            }
            subj_vars[da] = sv

            frm = ttk.LabelFrame(inner2, text=f'  {da}  ', padding=8)
            frm.pack(fill='x', padx=10, pady=4)

            # 위치 좌표
            r0 = tk.Frame(frm, bg='#F0F0F0'); r0.pack(fill='x', pady=2)
            ttk.Checkbutton(r0, text='위치 직접 지정',
                            variable=sv['use_pos']).pack(side='left')
            for lbl, key in [('X:', 'x'), ('Y:', 'y')]:
                ttk.Label(r0, text=f'  {lbl}').pack(side='left')
                ttk.Entry(r0, textvariable=sv[key], width=6).pack(side='left', padx=2)
                ttk.Label(r0, text='cm').pack(side='left')

            # 글꼴 크기
            r1 = tk.Frame(frm, bg='#F0F0F0'); r1.pack(fill='x', pady=2)
            for lbl, key in [('대주제 pt:', 'da_font_sz'),
                             ('상위주제:', 'sa_font_sz'),
                             ('주제영역:', 'za_font_sz'),
                             ('엔티티:',  'en_font_sz')]:
                ttk.Label(r1, text=f' {lbl}').pack(side='left')
                ttk.Entry(r1, textvariable=sv[key], width=5).pack(side='left', padx=1)

            # 색상
            r2 = tk.Frame(frm, bg='#F0F0F0'); r2.pack(fill='x', pady=2)

            def mk_clr(parent, lbl, var):
                ttk.Label(parent, text=f' {lbl}').pack(side='left')
                e = ttk.Entry(parent, textvariable=var, width=8)
                e.pack(side='left', padx=1)
                def pick(_v=var):
                    init = '#' + _v.get().lstrip('#')
                    res  = colorchooser.askcolor(color=init, title='색상 선택')
                    if res and res[1]:
                        _v.set(res[1].lstrip('#'))
                ttk.Button(parent, text='●', width=2, command=pick).pack(side='left')

            mk_clr(r2, '대주제 배경:', sv['da_bg'])
            mk_clr(r2, '상위주제 배경:', sv['sa_bg'])
            mk_clr(r2, '주제영역 배경:', sv['za_bg'])
            mk_clr(r2, '글꼴색:', sv['da_fc'])

            # 텍스트 오프셋
            r3 = tk.Frame(frm, bg='#F0F0F0'); r3.pack(fill='x', pady=2)
            ttk.Label(r3, text='텍스트 오프셋 (도형 상단에서 아래로):').pack(side='left')
            ttk.Entry(r3, textvariable=sv['text_offset'], width=7).pack(side='left', padx=4)
            ttk.Label(r3, text='cm').pack(side='left')

    # ─────────────────────────────────────────────────────────────
    # 탭 3: 글꼴 전체 설정
    # ─────────────────────────────────────────────────────────────
    tab3 = ttk.Frame(nb); nb.add(tab3, text='  전체 글꼴 설정  ')

    gf = lf(tab3, '기본 글꼴 크기 (주제영역별 미설정 시 적용)')
    gf.columnconfigure(5, weight=1)
    gda_v=tk.StringVar(value='13'); gsa_v=tk.StringVar(value='11')
    gza_v=tk.StringVar(value='10'); gen_v=tk.StringVar(value='9')
    gat_v=tk.StringVar(value='7.5')
    row_kv(gf, '대주제 글꼴', gda_v,'pt', r=0, c=0)
    row_kv(gf, '상위주제 글꼴', gsa_v,'pt', r=1, c=0)
    row_kv(gf, '주제영역 글꼴', gza_v,'pt', r=2, c=0)
    row_kv(gf, '엔티티명 글꼴', gen_v,'pt', r=0, c=4)
    row_kv(gf, '속성 글꼴',     gat_v,'pt', r=1, c=4)

    tf3 = lf(tab3, '텍스트 오프셋 (전체 기본 — 도형 상단에서 아래 여백)')
    gtoff_v = tk.StringVar(value='0')
    ttk.Label(tf3, text='전체 기본 오프셋:').grid(row=0,column=0,sticky='w')
    ttk.Entry(tf3, textvariable=gtoff_v, width=10).grid(row=0,column=1,padx=6)
    ttk.Label(tf3, text='cm').grid(row=0,column=2,sticky='w')
    ttk.Label(tf3,
              text='각 도형의 텍스트가 상단에서 이 값만큼 아래에서 시작됩니다.',
              foreground='#666').grid(row=1,column=0,columnspan=4,sticky='w',pady=2)

    # ─────────────────────────────────────────────────────────────
    # 상태바 + 버튼
    # ─────────────────────────────────────────────────────────────
    stat_v = tk.StringVar(value='● 준비')
    ttk.Label(root, textvariable=stat_v,
              foreground='#333', background='#F0F0F0').pack(pady=(2,0))

    bf = tk.Frame(root, bg='#F0F0F0'); bf.pack(pady=(4,10))

    # ── 설정 수집 ─────────────────────────────────────────────────
    def get_cfg():
        cfg = {
            'slide_w':    float(sw_v.get()),
            'slide_h':    float(sh_v.get()),
            'margin':     float(mg_v.get()),
            'gap':        float(gp_v.get()),
            'entity_w':   float(ew_v.get()),
            'entity_h':   float(eh_v.get()),
            'attr_row_h': float(er_v.get()),
            'pad':        float(pd_v.get()),
            'hdr_za':     float(za_v.get()),
            'hdr_sa':     float(sa2_v.get()),
            'hdr_da':     float(da2_v.get()),
            'title_h':    float(ttl_v.get()),
            'da_font_sz': float(gda_v.get()),
            'sa_font_sz': float(gsa_v.get()),
            'za_font_sz': float(gza_v.get()),
            'en_font_sz': float(gen_v.get()),
            'attr_font_sz': float(gat_v.get()),
            'text_offset':  float(gtoff_v.get()),
            'line_color':   lc_v.get().strip(),
            'cross_color':  cc_v.get().strip(),
            'styles': {},
        }
        positions = {}
        for da, sv in subj_vars.items():
            cfg['styles'][da] = {
                'da_font_sz':  float(sv['da_font_sz'].get()),
                'sa_font_sz':  float(sv['sa_font_sz'].get()),
                'za_font_sz':  float(sv['za_font_sz'].get()),
                'en_font_sz':  float(sv['en_font_sz'].get()),
                'da_fc':       sv['da_fc'].get().strip(),
                'da_bg':       sv['da_bg'].get().strip(),
                'sa_bg':       sv['sa_bg'].get().strip(),
                'za_bg':       sv['za_bg'].get().strip(),
                'text_offset': float(sv['text_offset'].get()),
            }
            if sv['use_pos'].get():
                positions[da] = {
                    'x': float(sv['x'].get()),
                    'y': float(sv['y'].get()),
                }
        return cfg, positions

    def do_load():
        path = xl_v.get().strip()
        if not os.path.isfile(path):
            messagebox.showerror('오류', f'파일 없음:\n{path}')
            return None, None
        return load_tree(path), None

    def on_load():
        stat_v.set('● 데이터 로드 중…'); root.update()
        tree, _ = do_load()
        if tree:
            populate_tab2(list(tree.keys()))
            stat_v.set(f'● 로드 완료: 대주제 {len(tree)}개')

    def on_save():
        try:
            cfg, pos = get_cfg()
            data = {'cfg': cfg, 'positions': pos,
                    'excel': xl_v.get(), 'output': out_v.get()}
            with open(SETTINGS_FILE,'w',encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            stat_v.set(f'● 설정 저장: {SETTINGS_FILE}')
        except Exception as e:
            messagebox.showerror('오류', f'저장 실패:\n{e}')

    def on_load_cfg():
        path = filedialog.askopenfilename(
            filetypes=[('JSON','*.json'),('모든','*.*')],
            initialfile=SETTINGS_FILE)
        if not path: return
        try:
            with open(path,'r',encoding='utf-8') as f:
                data = json.load(f)
            c = data.get('cfg', {})
            for var, key, dv in [
                (sw_v,'slide_w',100),(sh_v,'slide_h',55),(mg_v,'margin',0.8),
                (gp_v,'gap',0.35),(ew_v,'entity_w',4.8),(eh_v,'entity_h',1.3),
                (er_v,'attr_row_h',0.48),(pd_v,'pad',0.55),(za_v,'hdr_za',0.82),
                (sa2_v,'hdr_sa',0.88),(da2_v,'hdr_da',0.95),(ttl_v,'title_h',1.25),
                (gda_v,'da_font_sz',13),(gsa_v,'sa_font_sz',11),(gza_v,'za_font_sz',10),
                (gen_v,'en_font_sz',9),(gat_v,'attr_font_sz',7.5),(gtoff_v,'text_offset',0),
                (lc_v,'line_color','404040'),(cc_v,'cross_color','FF6600'),
            ]:
                var.set(str(c.get(key, dv)))
            if data.get('excel'): xl_v.set(data['excel'])
            if data.get('output'): out_v.set(data['output'])
            stat_v.set(f'● 설정 불러옴: {path}')
        except Exception as e:
            messagebox.showerror('오류', f'불러오기 실패:\n{e}')

    def on_generate():
        stat_v.set('● 데이터 로드 중…'); root.update()
        tree, _ = do_load()
        if not tree: return
        populate_tab2(list(tree.keys()))
        rels = extract_relations(tree)
        stat_v.set('● PPTX 생성 중…'); root.update()
        try:
            cfg, pos = get_cfg()
            prs = build_pptx(tree, rels, cfg, pos or None)
            out = out_v.get().strip()
            prs.save(out)
            ne = sum(len(ed) for sad in tree.values()
                     for zad in sad.values() for ed in zad.values())
            stat_v.set(f'● 완료! 엔티티 {ne}개 / 관계 {len(rels)}개')
            messagebox.showinfo('생성 완료',
                f'PPTX 저장 완료!\n\n  파일   : {out}\n'
                f'  엔티티 : {ne}개\n  관계선 : {len(rels)}개\n\n'
                f'PowerPoint에서 열면 모든 도형을 편집할 수 있습니다.')
        except Exception as e:
            import traceback
            messagebox.showerror('오류', f'{e}\n\n{traceback.format_exc()[-400:]}')
            stat_v.set('● 오류')

    # 버튼 배치
    for col, (txt, cmd) in enumerate([
        ('  데이터 로드  ',   on_load),
        ('  ▶ 생성  ',        on_generate),
        ('  💾 설정 저장  ',  on_save),
        ('  📂 설정 불러오기  ', on_load_cfg),
        ('  ✕ 종료  ',        root.destroy),
    ]):
        ttk.Button(bf, text=txt, command=cmd).grid(row=0, column=col, padx=5)

    root.mainloop()

# ════════════════════════════════════════════════════════════════════
# 6. CLI
# ════════════════════════════════════════════════════════════════════
def parse_args():
    p = argparse.ArgumentParser(
        description='민통선 출입관리체계 개념모델 구성도 생성기',
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('-i','--input',  default='민통선_엔티티속성정의서.xlsx')
    p.add_argument('-o','--output', default='민통선_개념모델.pptx')
    p.add_argument('--slide-w',  type=float, default=100)
    p.add_argument('--slide-h',  type=float, default=55)
    p.add_argument('--entity-w', type=float, default=4.8)
    p.add_argument('--entity-h', type=float, default=1.3)
    p.add_argument('--gui',      action='store_true')
    return p.parse_args()

# ════════════════════════════════════════════════════════════════════
# 진입점
# ════════════════════════════════════════════════════════════════════
if __name__ == '__main__':
    args = parse_args()

    use_gui = args.gui
    if not use_gui:
        try:
            import tkinter
            if len(sys.argv) == 1:
                use_gui = True
        except ImportError:
            pass

    if use_gui:
        run_gui()
        sys.exit(0)

    # CLI
    print('='*62)
    print('  민통선 개념모델 구성도 생성기 v2.0 [CLI]')
    print('='*62)

    if not os.path.isfile(args.input):
        print(f'[오류] 파일 없음: {args.input}'); sys.exit(1)

    print(f'  입력  : {args.input}')
    print(f'  출력  : {args.output}')
    print('  ① 데이터 로드…')
    tree = load_tree(args.input)
    rels = extract_relations(tree)
    ne = sum(len(ed) for sad in tree.values()
             for zad in sad.values() for ed in zad.values())
    print(f'     → 대주제 {len(tree)}개 / 엔티티 {ne}개 / 관계 {len(rels)}개')

    cfg = dict(slide_w=args.slide_w, slide_h=args.slide_h,
               entity_w=args.entity_w, entity_h=args.entity_h)

    print('  ② PPTX 생성…')
    prs = build_pptx(tree, rels, cfg)
    prs.save(args.output)
    print(f'  ③ 저장 완료 → {args.output}')
    print('='*62)
