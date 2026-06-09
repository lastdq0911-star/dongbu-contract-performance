from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = '/home/user/dongbu-contract-performance/SEBANG_AI_공모전_신청서_작성본.docx'

# ── 헬퍼 ─────────────────────────────────────────────────────────
NAVY     = (0x0F, 0x29, 0x5C)
GRAY_BG  = 'F2F2F2'
NAVY_HEX = '0F295C'

def set_cell(cell, text, bold=False, size=None, align=None, color=None):
    cell.text = ''
    p = cell.paragraphs[0]
    if align:
        p.alignment = align
    run = p.add_run(text)
    run.bold = bold
    if size:  run.font.size = Pt(size)
    if color: run.font.color.rgb = RGBColor(*color)

def add_run(para, text, bold=False, size=None, color=None):
    run = para.add_run(text)
    run.bold = bold
    if size:  run.font.size = Pt(size)
    if color: run.font.color.rgb = RGBColor(*color)
    return run

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def table_border(table):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for side in ['top','left','bottom','right','insideH','insideV']:
        el = OxmlElement(f'w:{side}')
        el.set(qn('w:val'),   'single')
        el.set(qn('w:sz'),    '4')
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), '999999')
        tblBorders.append(el)
    tblPr.append(tblBorders)

def make_2col_table(doc, rows):
    """rows: list of (label, content_fn)"""
    tbl = doc.add_table(rows=len(rows), cols=2)
    tbl.style = 'Table Grid'
    table_border(tbl)
    tbl.columns[0].width = Cm(2.8)
    for i, (label, fn) in enumerate(rows):
        row = tbl.rows[i]
        set_cell(row.cells[0], label, bold=True, size=9,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_bg(row.cells[0], GRAY_BG)
        row.cells[1].text = ''
        fn(row.cells[1])
    return tbl

# ── 문서 생성 ─────────────────────────────────────────────────────
d = Document()
for sec in d.sections:
    sec.top_margin    = Cm(2.0)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin   = Cm(2.5)
    sec.right_margin  = Cm(2.5)

d.styles['Normal'].font.name = '맑은 고딕'
d.styles['Normal'].font.size = Pt(10)

# ══════════════════════════════════════════════
# [첨부#1] 제목
# ══════════════════════════════════════════════
p = d.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p, '[첨부#1] 신청서', bold=True, size=13, color=NAVY)

p2 = d.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(p2, 'SEBANG AI 아이디어 공모전 신청서', bold=True, size=14, color=NAVY)
d.add_paragraph()

# ── 1. 신청자 정보 ──────────────────────────────
s1 = d.add_paragraph()
add_run(s1, '1. 신청자 정보', bold=True, size=11, color=NAVY)

tbl1 = d.add_table(rows=8, cols=4)
tbl1.style = 'Table Grid'
table_border(tbl1)

# 참가구분
r = tbl1.rows[0]
set_cell(r.cells[0], '참가 구분', bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
set_cell_bg(r.cells[0], GRAY_BG)
m = r.cells[1].merge(r.cells[2]).merge(r.cells[3])
set_cell(m, '☑ 개인    □ 단체(팀)  총 인원 (         )명', size=9)

# 팀 대표/개인 — 소속/이름/이메일+직급/연락처
sub_rows = [
    (1, '팀 대표\n및 개인', '소속', None, None, None),
    (2, '',              '이름', None, None, None),
]
for ri, hd, l1, _, __, ___ in sub_rows:
    r = tbl1.rows[ri]
    set_cell(r.cells[0], hd, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_bg(r.cells[0], GRAY_BG)
    set_cell(r.cells[1], l1, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_bg(r.cells[1], GRAY_BG)
    m = r.cells[2].merge(r.cells[3])
    set_cell(m, '', size=9)

r = tbl1.rows[3]
set_cell(r.cells[0], '', size=9); set_cell_bg(r.cells[0], GRAY_BG)
set_cell(r.cells[1], '이메일', bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER); set_cell_bg(r.cells[1], GRAY_BG)
set_cell(r.cells[2], '직급',   bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER); set_cell_bg(r.cells[2], GRAY_BG)
set_cell(r.cells[3], '', size=9)

r = tbl1.rows[4]
set_cell(r.cells[0], '', size=9); set_cell_bg(r.cells[0], GRAY_BG)
set_cell(r.cells[1], '연락처', bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER); set_cell_bg(r.cells[1], GRAY_BG)
m = r.cells[2].merge(r.cells[3]); set_cell(m, '', size=9)

for idx, ri in enumerate([5,6,7]):
    r = tbl1.rows[ri]
    set_cell(r.cells[0], '팀원' if idx==0 else '', bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_bg(r.cells[0], GRAY_BG)
    set_cell(r.cells[1], '소속', bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER); set_cell_bg(r.cells[1], GRAY_BG)
    set_cell(r.cells[2], '이름', bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER); set_cell_bg(r.cells[2], GRAY_BG)
    set_cell(r.cells[3], '', size=9)

d.add_paragraph()

# ── 2. 제안 개요 ──────────────────────────────
s2 = d.add_paragraph()
add_run(s2, '2. 제안 개요', bold=True, size=11, color=NAVY)

def f_title(c):
    p = c.paragraphs[0]
    add_run(p, 'AI 기반 벌크운송 하불실적 관리 웹 서비스\n— 구간·협력사·화주별 실적 분석 및 구간 단가 자동 산출', bold=True, size=10)

def f_field(c):
    p = c.paragraphs[0]
    add_run(p, '□ 업무 혁신 : 반복되거나 수작업 비중이 높은 비효율적인 업무 자동화 (효율화)\n', size=9)
    add_run(p, '□ 의사결정 지원 : 보고 자료 등 작성 시, 신속하고 정확한 의사결정을 지원하는 분석 툴\n', size=9)
    add_run(p,
        '☑ 기타 : 고객 서비스 개선 및 신규 가치 창출 등 기타 아이디어 제안\n',
        bold=True, size=9)
    add_run(p, '  ※ 업무 혁신(수작업 자동화) + 의사결정 지원(실적 분석 툴) 복합 적용', size=8, color=(0x80,0x80,0x80))

def f_summary(c):
    p = c.paragraphs[0]
    add_run(p,
        '동부계약관리팀은 벌크운송 하불(외주비) 실적을 엑셀로 수작업 관리해왔으나, '
        '데이터 분산·오류·분석 한계로 신속한 의사결정이 어려운 상황이었다.\n'
        'Supabase 클라우드 DB와 AI(Claude)를 활용해 하불 실적 데이터를 실시간 조회·분석할 수 있는 '
        '웹 서비스를 개발하였으며, 별도 설치 없이 PC·모바일 브라우저에서 즉시 사용 가능하다.\n'
        '협력사별·고객(화주)별·구간별 실적 집계, 구간 단가(최저가·최고가·평균가) 자동 산출, '
        '기간 필터·엑셀 내보내기 등 현장 실무에 바로 활용 가능한 기능을 제공한다.\n'
        '수작업 보고서 작성 시간을 대폭 단축하고 실적 기반의 협력사 단가 협상·계약 검토를 지원한다.',
        size=9)

make_2col_table(d, [
    ('제목',    f_title),
    ('분야',    f_field),
    ('제안\n개요', f_summary),
])

d.add_paragraph()

# 유의사항
for txt in [
    '※ 유의사항',
    '1) 제안서 제출 시 ① 신청서, ② 아이디어 활용 동의서를 모두 제출해야 하며, 제출 서류는 반환하지 않습니다.',
    '2) 신청자는 공모안이 제3자의 저작권, 초상권 등을 침해하지 않도록 주의 의무를 준수해야 하며, 관련된 사항은 신청자의 책임으로 합니다.',
    '3) 타 공모전 입상·수상작 또는 표절·중복응모·도용 등 부정행위 발생 시 심사 제외, 수상 취소 및 향후 당사에서 진행하는 DT 및 AI 관련 모든 공모 응시 참여가 제한됩니다.',
    '4) 접수 현황 및 심사 결과에 따라 시상 규모 축소 또는 변경할 수 있습니다.',
    '5) 1차 심사 및 평가는 신청자 정보를 미공개로 진행할 예정입니다.',
]:
    np = d.add_paragraph()
    np.paragraph_format.left_indent = Cm(0.5 if txt.startswith('※') else 0.8)
    add_run(np, txt, bold=txt.startswith('※'), size=8, color=(0x60,0x60,0x60))

d.add_paragraph()
ag = d.add_paragraph()
ag.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(ag, '신청자 본인은 위 내용에 동의하며, < SEBANG AI 아이디어 공모전 >에 참가 신청합니다.', bold=True, size=10)
d.add_paragraph()

dp = d.add_paragraph()
dp.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(dp, '2026년         월         일', size=11)
d.add_paragraph()

st = d.add_table(rows=1, cols=3)
st.alignment = 1; table_border(st)
set_cell(st.rows[0].cells[0], '신청인(팀 대표)', bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
set_cell_bg(st.rows[0].cells[0], GRAY_BG)
set_cell(st.rows[0].cells[1], '(소속)\n\n', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
set_cell(st.rows[0].cells[2], '(이름)                  (인)', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

d.add_page_break()

# ── 3. 아이디어 제안서 ─────────────────────────
s3 = d.add_paragraph()
add_run(s3, '3. 아이디어 제안서', bold=True, size=11, color=NAVY)

def idea_title(c):
    p = c.paragraphs[0]
    add_run(p, 'AI 기반 벌크운송 하불실적 관리 웹 서비스\n— 구간·협력사·화주별 실적 분석 및 구간 단가 자동 산출', bold=True, size=10)

def idea_current(c):
    p = c.paragraphs[0]
    add_run(p, '① 업무 현황\n', bold=True, size=9, color=NAVY)
    add_run(p,
        '• 업무 목적 : 벌크운송 하불(외주운송비) 실적 집계·분석 및 협력사·화주별 단가 관리\n'
        '• 수행 주기 : 월별 데이터 취합 및 보고 (수시 조회 필요)\n'
        '• 수행 인력 : 동부계약관리팀 담당자 (소수 인원)\n'
        '• 현재 방식 : SAP·WMS 등 시스템에서 추출한 데이터를 엑셀에 수작업으로 취합·가공·분석\n\n',
        size=9)
    add_run(p, '② 문제점\n', bold=True, size=9, color=NAVY)
    add_run(p,
        '• 엑셀 수작업 집계로 인한 오류 발생 및 버전 관리 어려움\n'
        '• 협력사별·화주별·구간별 실적을 교차 분석하려면 매번 피벗 재작업 필요 (건당 30분~수 시간 소요)\n'
        '• 구간별 하불 단가(최저가·최고가·평균가)를 파악할 체계적 수단이 없어 협상·계약 시 근거 부족\n'
        '• 데이터가 개인 PC에 분산 저장되어 공유·열람 불편, 외부·모바일 접근 불가\n'
        '• 기간별 추이 분석, 이상 단가 감지 등 고도 분석이 사실상 불가능',
        size=9)

def idea_proposal(c):
    p = c.paragraphs[0]
    add_run(p, '① 서비스 개요\n', bold=True, size=9, color=NAVY)
    add_run(p,
        'Supabase 클라우드 DB에 하불 실적 데이터를 저장하고, '
        'AI(Claude)로 개발한 단일 파일 웹앱(HTML/JS)으로 실시간 조회·분석·다운로드하는 시스템.\n'
        '별도 서버 구축 없이 GitHub Pages로 배포하여 PC·스마트폰 브라우저에서 즉시 접근 가능.\n\n',
        size=9)
    add_run(p, '② 주요 기능\n', bold=True, size=9, color=NAVY)
    features = [
        ('전체 데이터 조회',    '하불 실적 전체를 테이블로 조회, 열 정렬·크기 조절·드래그 이동, 기간·협력사·화주 필터, 엑셀 다운로드'),
        ('협력사별 분석',       '협력사별 하불 총액·건수·월별 추이를 카드·차트로 즉시 확인, 기간 설정 가능'),
        ('고객(화주)별 분석',   '화주별 물동량·하불 실적 비교 분석, 기간별 증감 추이 파악'),
        ('구간 단가 분석',      '출발지·도착지·톤급 조합별 최저가·최고가·평균가 자동 산출 (중량=1톤: 톤급당 단가 / 그 외: 톤당 단가 자동 구분)'),
        ('스마트 검색',         '출발지·도착지 자동완성 검색, 키보드(↑↓ Enter) 및 터치 지원'),
        ('실시간 데이터 관리',  '관리자가 Supabase에 데이터 업로드 시 전체 사용자에게 즉시 반영'),
    ]
    for fn, fd in features:
        add_run(p, f'  • {fn} : ', bold=True, size=9)
        add_run(p, f'{fd}\n', size=9)

    add_run(p, '\n③ 장애요인 및 해결방안\n', bold=True, size=9, color=NAVY)
    issues = [
        ('기존 엑셀 데이터 이관',     'SAP·엑셀 데이터를 CSV 형태로 1회 업로드, 이후 월별 append 방식으로 누적 관리'),
        ('전산 시스템 연동 부재',      'Supabase REST API 활용으로 별도 IT 인프라 없이 즉시 운영 가능'),
        ('보안·접근 권한',             'Supabase Row Level Security + 사용자 인증으로 권한별 접근 제어'),
        ('유지보수 전담 인력 부족',    '단일 HTML 파일 구조로 수정·배포가 단순, AI(Claude) 보조로 비개발자도 운영 가능'),
    ]
    for iss, sol in issues:
        add_run(p, f'  • {iss} → ', bold=True, size=9)
        add_run(p, f'{sol}\n', size=9)

    add_run(p, '\n④ 시스템 구성도\n', bold=True, size=9, color=NAVY)
    add_run(p,
        '  [사용자 : 계약관리팀 담당자 / 관리자]\n'
        '          ↓ PC·모바일 브라우저 접속 (GitHub Pages)\n'
        '  [웹 프론트엔드 — HTML/CSS/JavaScript 단일 파일]\n'
        '    - 전체데이터 / 협력사별 / 고객별 / 구간단가 분석 화면\n'
        '    - 자동완성 검색 · 필터 · 정렬 · 엑셀 다운로드\n'
        '          ↕ REST API (Supabase)\n'
        '  [Supabase 클라우드 DB]\n'
        '    - 하불 실적 테이블 (용역구분·출발지·도착지·톤급·중량·하불금액·협력사·화주·연도·월)\n'
        '          ↑ 월별 데이터 업로드 (CSV / 직접 입력)\n'
        '  [SAP / WMS / 엑셀 원천 데이터]',
        size=9)

def idea_tools(c):
    p = c.paragraphs[0]
    tools = [
        ('활용 DATA',   'SAP, WMS, 기존 엑셀 하불 실적 데이터'),
        ('AI 활용',     'Claude AI — 전체 서비스 설계·코드 생성·기능 개선·디버깅 보조 (비개발자 주도 개발 실현)'),
        ('개발 도구',   'HTML5 / CSS3 / JavaScript (단일 파일 웹앱, 프레임워크 미사용)'),
        ('DB·API',      'Supabase (PostgreSQL 기반 클라우드 DB, REST API, 인증)'),
        ('배포',        'GitHub Pages (무료, 별도 서버 불필요, 브라우저에서 즉시 접근)'),
        ('기타',        'Chart.js (차트), SheetJS (엑셀 다운로드), 반응형 CSS (모바일 지원)'),
    ]
    for k, v in tools:
        add_run(p, f'• {k} : ', bold=True, size=9)
        add_run(p, f'{v}\n', size=9)

def idea_effect(c):
    p = c.paragraphs[0]
    add_run(p, '① 정량적 기대 효과\n', bold=True, size=9, color=NAVY)
    effects = [
        ('업무 시간 절감',  '월별 실적 보고서 작성 시간 수 시간 → 10분 이내 (약 80~90% 절감)'),
        ('오류 감소',       '수작업 집계 오류 제거, 단일 DB 기준으로 데이터 일관성 확보'),
        ('비용',            '서버 구축·유지 비용 ₩0 (GitHub Pages + Supabase 무료 티어 운영)'),
        ('접근성',          '사무실 외 장소·스마트폰에서도 실시간 조회 가능'),
    ]
    for k, v in effects:
        add_run(p, f'  • {k} : ', bold=True, size=9)
        add_run(p, f'{v}\n', size=9)
    add_run(p, '\n② 정성적 기대 효과\n', bold=True, size=9, color=NAVY)
    add_run(p,
        '  • 구간별 실적 기반 단가 데이터를 축적하여 협력사 단가 협상 시 객관적 근거 확보\n'
        '  • 화주별 물동량 추이 파악으로 계약 갱신·영업 전략 수립 지원\n'
        '  • 이상 단가(최저가 대비 과다 청구 등) 감지로 비용 절감 및 계약 리스크 관리\n'
        '  • AI 보조 개발 사례로서 비개발자 직원도 업무 맞춤형 디지털 도구를 직접 만들 수 있다는 가능성 입증\n'
        '  • 구조 확장 시 타 팀(영업·물류·재무) 하불 데이터와 통합 가능',
        size=9)

make_2col_table(d, [
    ('제목',       idea_title),
    ('현황\n진단', idea_current),
    ('제안\n내용', idea_proposal),
    ('활용\n도구', idea_tools),
    ('기대\n효과', idea_effect),
])

d.add_page_break()

# ══════════════════════════════════════════════
# [첨부#2] 아이디어 활용 동의서
# ══════════════════════════════════════════════
h2 = d.add_paragraph()
h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(h2, '[첨부#2] 아이디어 활용 동의서', bold=True, size=13, color=NAVY)
d.add_paragraph()

t2 = d.add_paragraph()
t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(t2, '아이디어 활용 동의서', bold=True, size=14, color=NAVY)
d.add_paragraph()

d.add_paragraph().add_run(
    '본인은 세방㈜에서 주관하는 <SEBANG AI 아이디어 공모전>(이하 "공모전")에 아이디어를 제출함에 있어, 아래 사항에 동의합니다.'
).font.size = Pt(10)
d.add_paragraph()

bel = d.add_paragraph()
bel.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(bel, '-     아          래     -', bold=True, size=10)
d.add_paragraph()

for txt in [
    '본인이 제출한 아이디어는 회사의 업무 개선, AI 활용 및 업무 자동화 과제 추진, 교육, 홍보 및 기타 내부 목적을 위해 활용될 수 있다.',
    '회사는 해당 아이디어를 사용, 수정, 보완, 재구성할 수 있으며, 아이디어를 기반으로 한 결과물을 내부자료(보고서, 사례집 등)로 제작 및 활용할 수 있다.',
    '회사는 제출된 아이디어를 시스템 개선, AI 기반 서비스 기획 및 개발 등 다양한 형태로 제한 없이 활용할 수 있다.',
]:
    cp = d.add_paragraph()
    cp.paragraph_format.left_indent = Cm(1)
    add_run(cp, '• ' + txt, size=10)

d.add_paragraph()
add_run(d.add_paragraph(), '본인은 상기 내용을 충분히 이해하였으며, 이에 동의합니다.', bold=True, size=10)
d.add_paragraph()

dp2 = d.add_paragraph()
dp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_run(dp2, '2026년         월         일', size=11)
d.add_paragraph()

st2 = d.add_table(rows=4, cols=3)
st2.alignment = 1; table_border(st2)
for i, role in enumerate(['대표자','팀원','팀원','팀원']):
    r = st2.rows[i]
    set_cell(r.cells[0], role, bold=True, size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell_bg(r.cells[0], GRAY_BG)
    set_cell(r.cells[1], '(소속)\n\n', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell(r.cells[2], '(이름)                  (인)', size=9, align=WD_ALIGN_PARAGRAPH.CENTER)

d.save(OUT)
print('저장 완료:', OUT)
