"""공모전 서식2 본문을 HWP 로 옮길 수 있는 .docx 로 생성한다.

한글(HWP)은 .docx 를 열 수 있으므로, 여기서 만든 파일을 한글에서 열고
'다른 이름으로 저장 → .hwp' 하면 된다. 공식 서식의 글꼴·여백·줄간격을
그대로 반영해 두었으므로 옮긴 뒤 손볼 것이 적다.

공식 작성요령(서식2 원본에서 추출)
    좌우여백 20mm · 위아래 10mm · 머리말/꼬리말 10mm
    줄간격 140% · 휴먼명조 15pt(본문) · 맑은고딕 12pt(표·캡션)
    개조식으로 상세하게, 본문 15쪽 이내
    파란색 글씨(작성요령)는 제출 시 삭제 → 여기서는 애초에 넣지 않는다

사용:  python system/make_form2_docx.py
출력:  internal/서식2_본문.docx
"""
from __future__ import annotations

import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

ROOT = Path(__file__).resolve().parent.parent
FIG = ROOT / "docs" / "figures"
SHOT = ROOT / "docs" / "screenshots" / "doc"   # 문서용 크롭본(폭:높이 1.9:1)
if not SHOT.exists():
    SHOT = ROOT / "docs" / "screenshots"
OUT = ROOT / "internal" / "서식2_본문.docx"

BODY_FONT = "휴먼명조"
TABLE_FONT = "맑은 고딕"
BODY_PT = 15
SMALL_PT = 11
INK = RGBColor(0x1F, 0x2D, 0x3D)
ACCENT = RGBColor(0x1F, 0x55, 0x82)


# --------------------------------------------------------------------------- #
# 서식 기초
# --------------------------------------------------------------------------- #
def set_run(run, *, size=BODY_PT, font=BODY_FONT, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.name = font
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    # 한글 글꼴은 eastAsia 속성에 따로 지정해야 적용된다
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font)
    rFonts.set(qn("w:ascii"), font)
    rFonts.set(qn("w:hAnsi"), font)


def para(doc, text="", *, size=BODY_PT, font=BODY_FONT, bold=False, color=None,
         indent=0.0, space_after=4, align=None, line=1.4):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = line              # 공식 요령의 줄간격 140%
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(0)
    if indent:
        pf.left_indent = Cm(indent)
    if align is not None:
        p.alignment = align
    if text:
        set_run(p.add_run(text), size=size, font=font, bold=bold, color=color)
    return p


def chapter(doc, no, title):
    """장 제목 — 서식2 의 번호 박스를 굵은 제목으로 대체한다."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.2
    set_run(p.add_run("%d  %s" % (no, title)), size=17, font=TABLE_FONT,
            bold=True, color=ACCENT)
    # 아래 실선
    pPr = p._element.get_or_add_pPr()
    bdr = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "12")
    bot.set(qn("w:color"), "1F5582")
    bdr.append(bot)
    pPr.append(bdr)


def head_box(doc, text):
    """□ 수준 소제목."""
    para(doc, "□ " + text, size=13.5, font=TABLE_FONT, bold=True,
         color=ACCENT, space_after=4, line=1.2)


def bullet(doc, text, *, sub=False):
    """◦ / - 수준 항목. 개조식 뼈대."""
    mark, ind = ("- ", 1.0) if sub else ("◦ ", 0.4)
    para(doc, mark + text, size=BODY_PT if not sub else BODY_PT - 1,
         indent=ind, space_after=3)


def body(doc, text, *, indent=0.8):
    """항목 아래 서술 문단."""
    para(doc, text, size=BODY_PT - 2, indent=indent, space_after=5, line=1.32)


def caption(doc, text):
    para(doc, text, size=SMALL_PT - 1, font=TABLE_FONT, color=INK,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=7, line=1.16)


def figure(doc, path, cap, width_cm=11.5):
    p = path if isinstance(path, Path) else Path(path)
    if not p.exists():
        para(doc, "[그림 누락: %s]" % p.name, size=SMALL_PT, color=RGBColor(0xC0, 0x39, 0x2B))
        return
    fp = doc.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(6)
    fp.paragraph_format.space_after = Pt(3)
    fp.add_run().add_picture(str(p), width=Cm(width_cm))
    caption(doc, cap)


def figure_pair(doc, left, right, cap, width_cm=6.9):
    """그림 두 장을 나란히. 15쪽 제한 안에서 더 많은 화면을 보이기 위한 배치."""
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = True
    for cell, path in zip(t.rows[0].cells, (left, right)):
        cell.text = ""
        pp = cell.paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.space_after = Pt(0)
        pth = path if isinstance(path, Path) else Path(path)
        if pth.exists():
            pp.add_run().add_picture(str(pth), width=Cm(width_cm))
        else:
            set_run(pp.add_run("[그림 누락: %s]" % pth.name), size=SMALL_PT)
        # 표 테두리 제거
        tcPr = cell._tc.get_or_add_tcPr()
        bd = OxmlElement("w:tcBorders")
        for side in ("top", "left", "bottom", "right"):
            e = OxmlElement("w:" + side)
            e.set(qn("w:val"), "nil")
            bd.append(e)
        tcPr.append(bd)
    caption(doc, cap)


def table(doc, headers, rows, *, widths=None, size=SMALL_PT):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        pp = cell.paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pp.paragraph_format.line_spacing = 1.15
        set_run(pp.add_run(str(h)), size=size, font=TABLE_FONT, bold=True)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "EAF1F7")
        cell._tc.get_or_add_tcPr().append(shd)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            pp = cells[i].paragraphs[0]
            pp.paragraph_format.line_spacing = 1.15
            if i > 0 and str(v).replace(",", "").replace(".", "").replace("%", "").isdigit():
                pp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            set_run(pp.add_run(str(v)), size=size, font=TABLE_FONT)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    _sp = doc.add_paragraph()
    _sp.paragraph_format.space_after = Pt(2)
    _sp.paragraph_format.line_spacing = 0.6
    return t


# --------------------------------------------------------------------------- #
def build():
    doc = Document()
    s = doc.sections[0]
    s.left_margin = s.right_margin = Mm(20)
    s.top_margin = s.bottom_margin = Mm(10)
    s.header_distance = s.footer_distance = Mm(10)

    st = doc.styles["Normal"]
    st.font.name = BODY_FONT
    st.font.size = Pt(BODY_PT)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), BODY_FONT)

    # ── 표지 성격의 제목 ────────────────────────────────────────────────
    para(doc, "자치법규 정책지도", size=22, font=TABLE_FONT, bold=True,
         color=ACCENT, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, line=1.1)
    para(doc, "전국 자치법규 199,858건의 그래프 지도와 근거 기반 조례 격차 진단",
         size=13, font=TABLE_FONT, color=INK,
         align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14, line=1.2)

    # =====================================================================
    chapter(doc, 1, "배경 및 개요")

    head_box(doc, "기획 목적")
    bullet(doc, "전국 지방자치단체가 조례를 만들 때 참고할 근거를 한자리에서 "
                "제공하고, 무엇이 확인된 사실이고 무엇이 추정인지를 함께 공시한다")
    body(doc,
        "담당자는 유사 사례를 검색엔진으로 찾고 자치법규정보시스템에서 조례를 하나씩 열람"
        "하며 다른 지방자치단체에 직접 전화해 확인한다. 전국 기초자치단체 227곳이 같"
        "은 조사를 각자 반복하지만 그 결과가 축적되거나 공유되는 구조는 없다. 본 아이"
        "디어는 전국의 자치법규·상위법령·예산·국회 의안·행정경계를 하나의 그래프로 통합"
        "하고 지도 위에 얹어, '우리와 여건이 비슷한 곳에는 있는데 우리에게는 없는 조"
        "례'를 근거와 함께 확인하게 한다.")

    head_box(doc, "배경 및 필요성")
    bullet(doc, "조례 입안의 정보 비대칭 — 조회는 가능하나 비교가 불가능하다")
    body(doc,
        "자치법규정보시스템은 조례 원문을 제공하지만 그 구조는 개별 문서의 '목록'이어서"
        " 지방자치단체 간 비교나 '우리에게 무엇이 없는가'라는 질의를 지원하지 않는다."
        " 국가법령정보센터도 위임 조항을 제공하지만 그 위임이 227개 기초자치단체 중 "
        "어디에서 이행되고 어디에서 이행되지 않았는지는 조회 단위에서 끊긴다.")

    bullet(doc, "선행 연구의 규모 제약 — 표본 수십 건 단위의 수작업 코딩")
    body(doc,
         "조례를 대상으로 한 국내 선행 연구는 특정 정책 주제를 정해 수십 건에서 "
         "수백 건 규모의 조례를 수작업으로 코딩하는 방식이 일반적이다. 이러한 "
         "접근은 깊이를 확보하는 대신 전국 단위의 구조를 보지 못하며, 연구마다 "
         "코딩 기준이 달라 결과를 서로 잇기 어렵다. 전국 자치법규 전량을 대상으로 "
         "관계 구조를 구축하고 그 위에서 비교 질의를 수행한 사례는 확인되지 않는다.")

    bullet(doc, "공간정보와의 결합 필요성 — 조례는 지역의 함수이기 때문이다")
    body(doc,
         "조례는 특정 지방자치단체의 관할 구역에 효력을 미치는 규범이므로 본질적으로 "
         "공간 데이터이다. 어떤 조례가 어디에 있고 어디에 없는지, 인접한 지방자치단체 "
         "사이에 유사한 조례가 함께 나타나는지는 행정경계 위에서만 답할 수 있는 "
         "질문이다. 본 아이디어가 행정경계와 공간통계를 핵심 구성요소로 삼는 이유가 "
         "여기에 있다.")

    head_box(doc, "아이디어의 우수성 및 타당성")
    bullet(doc, "전국 전량을 대상으로 한다")
    body(doc,
         "표본이 아니라 자치법규 199,858건, 상위법령 29,811건, 조문 본문 2,365,068건을 "
         "모두 수집하여 노드 250,416개와 엣지 1,125,036개의 단일 그래프로 구성하였다. "
         "전국 227개 기초자치단체 전부가 동일한 기준으로 비교된다.")

    bullet(doc, "공간 인접성과 정책 유사성을 함께 잇는 이중 엣지 구조를 갖는다")
    body(doc,
         "지방자치단체와 조례를 노드로 두고, 행정경계가 맞닿아 있는 공간 인접 관계 "
         "1,098건과 조례 사이의 정책 유사 관계 21,740건, 그리고 조례가 상위법령을 "
         "준거로 인용한 관계 421,627건을 각각 별도의 엣지로 구성하였다. 이로써 "
         "'지리적으로 가까워서 닮았는가' 와 '정책 내용이 닮았는가' 를 구분해 물을 "
         "수 있다.")

    bullet(doc, "자체 한계를 정량적으로 측정하여 화면에 공시한다")
    body(doc,
        "정확도를 주장하는 대신 어디까지 맞는지를 측정해 보여준다. 상위법령 인용 421"
        ",627건 중 인용 조문의 존재가 확인된 것은 182,784건(43.3%), 분"
        "야 자동 분류의 코더 간 일치도는 Cohen's κ 0.517, 이웃 추천 3개"
        " 모델 중 무작위 기준선을 넘는 것은 1개다. 이 수치를 '검증 공시' 화면으로"
        " 제공한다. 과장하지 않는 것이 행정 실무에서 쓰일 조건이라고 보았다.")

    figure(doc, SHOT / "01_dashboard.png",
           "그림 1. 전국 요약 화면 — 자치법규 199,858건, 상위법령 29,811건, 준거 인용 "
           "421,627건, 예산 세부사업 933,527행, 국회 의안 19,847건, 행정구역 556곳이 "
           "노드 250,416개·엣지 1,125,036개의 단일 그래프로 통합되어 있다.")

    # =====================================================================
    chapter(doc, 2, "구현 방법 및 결과")

    head_box(doc, "활용 데이터")
    bullet(doc, "전 항목이 무료 공개 API 이며 실제 호출로 확보하였다. 유료 데이터와 "
                "개인 데이터는 사용하지 않았다")
    table(doc,
          ["연번", "데이터명", "내용", "출처", "무상", "확보 실적"],
          [["1", "자치법규", "조례·규칙 메타와 조문 본문",
            "법제처 국가법령정보 Open API", "○", "199,858건 / 조문 2,365,068건"],
           ["2", "법령·행정규칙", "법률·시행령·시행규칙·행정규칙",
            "법제처 국가법령정보 Open API", "○", "29,811건 / 조문 86,745건"],
           ["3", "위임·인용 관계", "조례가 준거로 인용한 상위법령",
            "법제처(lsDelegated·lsStmd·lnkLsOrd)", "○", "421,627건"],
           ["4", "행정경계", "시도·시군구 경계 GeoJSON",
            "국토교통부 V-World 데이터 API", "○", "시군구 250개 폴리곤"],
           ["5", "법정동코드", "10자리 코드·행정구역 승계 이력",
            "행정안전부 행정표준코드(공공데이터포털)", "○", "행정구역 556곳 / 승계 17건"],
           ["6", "지방재정 세출", "세부사업별 예산현액·지출액",
            "행정안전부 지방재정365", "○", "933,527행"],
           ["7", "국회 의안", "발의법률안 메타",
            "국회사무처 열린국회정보", "○", "19,847건"],
           ["8", "발의자", "대표·공동 발의자 명단",
            "국회사무처 열린국회정보", "○", "241,546건"],
           ["9", "표결 기록", "의안별 국회의원 표결",
            "국회사무처 열린국회정보", "○", "57,178행 / 의원 320명"]],
          widths=[1.0, 2.4, 3.4, 3.6, 0.9, 4.0])
    body(doc,
         "법률·공간·재정·정치의 네 이질 도메인을 결합하였다. 모든 인증키는 이메일 "
         "기반으로 무료 발급되며, 발급 절차는 화면 단위로 문서화하여 공개 저장소에 "
         "함께 공개하였다(docs/11_API키_발급가이드.md). 데이터 확보 과정에서 각 API의 "
         "실제 응답을 검증하였고, 그 결과와 시행착오도 함께 기록하였다.", indent=0.4)

    head_box(doc, "구현 절차")
    bullet(doc, "수집에서 배포까지 6단계 파이프라인으로 구성하였다")
    table(doc,
          ["단계", "처리 내용", "산출"],
          [["① 수집", "공공 API 5종 호출, 재시도·속도제한·증분 갱신 처리", "SQLite 33개 테이블"],
           ["② 파싱", "조문을 제·항·호·목으로 분해, 위임 관계 4경로 추출, 분야 16종 분류",
            "조문 2,365,068건"],
           ["③ 그래프", "노드 8종·엣지 12종 구성, 행정구역 승계 반영", "250,416 노드 / 1,125,036 엣지"],
           ["④ 분석", "그래프 신경망, GraphRAG, 공간자기상관, 사건사분석", "지표 23종 · 모형 15개"],
           ["⑤ 배포", "정적 JSON 번들 생성 후 gzip 사전압축", "4,991파일 81.2MB"],
           ["⑥ 소비", "웹 화면, MCP 도구, AI 정책분석관", "화면 16종 · 도구 14종"]],
          widths=[1.8, 9.0, 4.5])
    body(doc,
         "코어는 파이썬 표준 라이브러리와 numpy 만으로 구현하여 별도의 딥러닝 "
         "프레임워크를 요구하지 않으며, 프런트엔드도 빌드 도구 없는 바닐라 "
         "자바스크립트로 작성하였다. 그 결과 서버 없이 정적 호스팅만으로 전국 "
         "데이터가 동작한다. 원본 405.6MB의 산출물을 gzip 으로 사전압축해 81.2MB로 "
         "줄였으며, 브라우저가 내려받는 시점에 해제한다.", indent=0.4)

    bullet(doc, "분석 기법은 전부 직접 구현하고 산식을 공개하였다")
    table(doc,
          ["기법", "사양", "적용 결과"],
          [["node2vec", "128차원, 랜덤워크 기반 임베딩", "조례 154,310건"],
           ["metapath2vec", "64차원, 이종 그래프 메타패스", "조례 154,310건"],
           ["GraphSAGE", "132차원, JK-Net 연결로 과평활 해소", "조례 154,310건"],
           ["GraphRAG", "BM25 어휘검색 + Dense 벡터검색을 RRF(k=60)로 융합",
            "조문 236만 건 색인"],
           ["공간자기상관", "전역 Moran's I + 국지 LISA, 조건부 순열 999회, BH-FDR 보정",
            "지표 23종"],
           ["사건사분석(EHA)", "이산시간 위험모형, 지자체 클러스터 로버스트 표준오차",
            "정책 3종 × 모형 5개"],
           ["유사 지방자치단체", "행정안전부 유사자치단체 기준 + 4개 방법 비교",
            "227곳 전수"]],
          widths=[2.6, 7.6, 5.1])
    body(doc,
         "법령 그래프의 식별자 체계는 유럽연합의 ELI(European Legislation Identifier), "
         "OASIS 의 Akoma Ntoso, 서지 표준 FRBR 에 대조하여 설계하였다. 조례의 "
         "'저작물(Work)'과 '판본(Expression)'을 분리해 두었기 때문에, 향후 개정 이력이 "
         "확보되면 같은 구조에서 판본 비교가 가능하다.", indent=0.4)

    head_box(doc, "산출물 ① — 조례 총량과 인구는 무관하다")
    figure_pair(doc, FIG / "F1_인구_대_조례수.png", FIG / "F2_분야별_변동계수.png",
                "그림 2·3. (좌) 인구와 현행 조례 수 — 인구는 135배 차이가 나지만 조례는 "
                "2.37배에 그치고 로그-로그 탄력성은 0.062다. (우) 분야별 조례 비중의 지자체 간 "
                "변동계수 — 행정·자치(0.097)와 안전·재난(0.100)은 의무·표준 영역, "
                "교통(0.831)과 농림·수산(0.639)은 지역 특성이 반영되는 재량 영역이다. "
                "점선은 조례 총량의 변동계수 0.181.")
    body(doc,
         "이 결과는 조례 수를 '입법 활동량' 의 지표로 사용하는 통념이 성립하지 않음을 "
         "보여준다. 기초자치단체의 조례집 규모는 지역의 수요가 아니라 지방자치법과 "
         "개별 법령이 요구하는 필수 조례 세트가 결정하는 제도적 상수에 가깝다. "
         "따라서 지방자치단체 사이의 실질적 차이는 조례의 개수가 아니라 그 구성에서 "
         "찾아야 한다.", indent=0.4)

    head_box(doc, "산출물 ② — 격차는 총량이 아니라 구성에 있다")
    table(doc,
          ["지표", "값", "의미"],
          [["서로 다른 정책키", "25,896개", "조례명을 정규화한 정책 단위"],
           ["한 곳만 보유", "17,784개 (68.7%)", "대부분의 정책은 지역 고유"],
           ["227곳 전부 보유", "1개", "주민투표 조례가 유일"],
           ["80% 이상 보유", "63개", "사실상의 전국 공통 핵심"],
           ["두 지자체 정책키 일치도", "Jaccard 0.158", "무작위 두 곳을 뽑으면 16%만 겹침"]],
          widths=[4.6, 3.6, 7.1])
    body(doc,
         "형식에서는 동형화가 일어났으나 내용에서는 일어나지 않았다. 이 값은 본 "
         "시스템이 별도로 산출한 유사 지방자치단체 정책 프로파일 코사인 유사도 "
         "중앙값 0.157 과 서로 다른 방법으로 계산했음에도 일치하여, 측정의 견고함을 "
         "교차 확인할 수 있었다.", indent=0.4)

    body(doc,
        "이 결과는 격차 분석의 설계를 바꾼다. 모든 분야를 같은 무게로 비교하면 도시 "
        "지방자치단체가 농림·수산 조례를 갖지 않은 것까지 결손으로 잡히기 때문이다. 본"
        " 시스템은 분야별 변동계수를 함께 제시해 실질적 누락과 지역 특성의 반영을 구분"
        "한다. 다만 분야 라벨은 규칙 기반 자동 분류이므로 순위는 견고하나 절대값에는 "
        "잡음이 있음을 함께 밝힌다.", indent=0.4)

    head_box(doc, "산출물 ③ — 조례는 공간적으로 뭉치는가")
    figure_pair(doc, FIG / "F3_공간자기상관_23지표.png", FIG / "F5_확산유형.png",
                "그림 4·5. (좌) 공간자기상관 23개 지표의 전역 Moran's I — 초록은 BH-FDR 보정을 "
                "통과한 국지 군집 보유 지표. 농림·수산이 0.764로 가장 강하게 뭉치고 동물·반려는 "
                "-0.014로 무작위와 구별되지 않는다. (우) 조문 구조 지문 공유율과 동형 군집의 "
                "시도 집중 배수 — 파란색은 전국 단일 원본형, 주황색은 인접 지자체 복제형이다. "
                "무작위 대조군의 지문 공유율은 0.8%다.")
    body(doc,
        "현행 조례 수의 전역 Moran's I 는 0.4333, 순열검정 999회 기준"
        " p=0.001로 유의하다. 국지 LISA 에서는 보정 전 55곳이 유의했으나 "
        "BH-FDR 보정 후 9곳만 남았다. 보정하지 않을 때 기대되는 위양성이 11."
        "2곳이므로 보정 없이 보고했다면 대부분이 위양성이었을 것이다. 본 시스템은 보정"
        " 전후를 함께 표시한다.", indent=0.4)
    body(doc,
         "확인된 국지 군집은 통념과 다르다. 조례 수가 낮은 값끼리 뭉친 저-저 군집 "
         "8곳은 농촌이 아니라 부산광역시와 대구광역시의 자치구였다. 자치구는 "
         "도시계획·상하수도 등 상당수 사무가 광역자치단체에 있어 조례 수가 구조적으로 "
         "적으며, 이는 행정 역량의 문제가 아니라 사무 배분의 결과로 읽는 것이 타당하다.",
         indent=0.4)
    figure_pair(doc, SHOT / "07_spatial.png", SHOT / "08_analytics.png",
                "그림 5·6. (좌) 공간자기상관 화면 — 지표별 전역 Moran's I 와 기댓값, "
                "순열검정 z·p, BH-FDR 통과 국지 군집, LISA 사분면 지도. "
                "(우) 확산 위험모형 — 이산시간 위험모형의 계수·표준오차·유의성과 "
                "확산 경로 세 가지의 비교표.")

    head_box(doc, "산출물 ④ — 확산의 유형을 조문 구조로 판별한다")
    bullet(doc, "자치법규 데이터에는 최초 제정일이 없다는 제약을 우회하였다")
    body(doc,
        "수집한 데이터에서 제정일로 기록된 값의 99.99%는 최초 제정일이 아니라 현행"
        " 판본의 공포일이었다. 판본 이력이 보존되지 않기 때문이며, 채택 시점을 전제로"
        " 하는 전통적 확산 분석은 근거가 약해진다. 본 연구는 이를 감추지 않고 시점이"
        " 필요 없는 대안 측정을 설계하였다. 조례의 조문 제목 시퀀스를 구조 지문으로 "
        "삼는 방식이다.")

    body(doc,
         "두 유형이 뚜렷하게 갈린다. 안전보안관과 자원봉사 조례는 조문 구조가 "
         "매우 비슷한데도 그 군집이 특정 시도에 몰려 있지 않다. 반면 맨발걷기와 "
         "공공심야약국 조례는 조문 구성이 제각각이지만 일치하는 소수의 사례는 "
         "인접한 지방자치단체에 집중되어 있다. 전자는 전국 어딘가의 단일 원본이 "
         "지역과 무관하게 퍼진 형태이고, 후자는 이웃을 보고 베낀 형태로 해석된다.",
         indent=0.4)
    bullet(doc, "실제 사례 — 안전보안관 조례")
    body(doc,
        "재난 및 안전관리 기본법 제66조의4 제3항(2019.12.3 신설)은 지방자치"
        "단체의 장이 주민 참여 제도를 마련해 시행할 수 있다고만 규정하고 제도의 명칭·"
        "요건·임기·지원 범위를 정하지 않았다. 6개월 뒤 아산시가 전국 최초로 안전보안"
        "관 운영 조례를 제정하였고 현재 186곳(81.9%)이 보유한다. 그런데 수집한"
        " 국가법령 조문 86,745건 어디에도 '안전보안관'은 없다. 제도의 이름과 조"
        "문 틀은 지방에서 만들어졌고 법적 근거는 사후에 부착되었다.", indent=0.4)

    head_box(doc, "산출물 ⑤ — 추천 모델의 성능을 스스로 측정한다")
    figure_pair(doc, FIG / "F4_신경망_모델평가.png", FIG / "F8_결손_로버스트니스.png",
                "그림 6·7. (좌) 이웃 조례 추천 3개 모델의 분야 일치율과 무작위 기준선 — 같은 후보 "
                "집합에서 이웃만 무작위로 바꿔도 30%대가 나오므로 기준선 없이 절대값만 보면 "
                "성능을 오독한다. (우) '없는 조례' 규모는 매칭 방법에 따라 지자체당 1.18건에서 "
                "8건까지 움직인다 — 제목만으로 보유 여부를 판정하는 방식의 한계를 보여준다.")
    body(doc,
         "세 모델을 모두 남겨 두고 평가 결과를 화면에 표로 공시하며, 기본 선택은 "
         "평가 1위 모델로 열리도록 하였다. 성능이 낮은 모델을 숨기지 않는 이유는 "
         "모델 사이의 불일치 자체가 사용자에게 정보이기 때문이다. 세 모델이 서로 "
         "다른 이웃을 제시한다면 그 조례는 구조적으로 유사한 사례가 뚜렷하지 않다는 "
         "뜻이며, 이는 추천을 그대로 신뢰하지 말라는 신호가 된다.", indent=0.4)
    figure_pair(doc, SHOT / "06_neural.png", SHOT / "10_trust.png",
                "그림 9·10. (좌) 신경망 유사도 — 모델별 이웃 조례와 함께 분야 일치율, "
                "무작위 기준선, 기준선 대비 배수를 표로 제시한다. (우) 검증 공시 — "
                "원문 확보율과 인용 대조 결과, 사람이 대조한 항목을 모집단과 함께 밝힌다.")

    head_box(doc, "산출물 ⑥ — 검증 상태를 데이터 구조에 넣는다")
    table(doc,
          ["검증 항목", "결과"],
          [["원문 확보", "조례 199,695 / 199,858건(99.92%)이 공식 원문에 연결"],
           ["인용 조문 대조", "421,627건 전건을 기계 대조 → 존재 확인 182,784건(43.3%), "
                          "불일치 28,309건(6.7%), 자동 확인 불가 210,534건(49.9%)"],
           ["사람이 직접 대조", "1,205건(법령 627·기관 578). 조례는 조문 수준 대조를 하지 않음"],
           ["조례↔예산 표본검증", "층화표본 584건 수작업 판정, 전체 정밀도 64.9%"],
           ["시간 무결성 자동감사", "규칙 위반 7,060건 탐지·보정"],
           ["분야 코딩 신뢰도", "Cohen's κ 0.517 (Landis & Koch 기준 moderate)"]],
          widths=[4.2, 11.1])
    body(doc,
         "본 시스템은 '전수 검증' 이라는 표현을 사용하지 않는다. 421,627건은 검사 "
         "대상 모집단이며 실제로 확인된 것은 43.3%다. 나머지 절반가량은 상위법령의 "
         "조문 텍스트를 확보하지 못했거나 인용 형식이 조문 단위로 특정되지 않아 "
         "자동 확인이 불가능했다. 이 구분을 화면에 그대로 노출하는 것이 본 "
         "시스템의 설계 원칙이다.", indent=0.4)

    head_box(doc, "산출물 ⑦ — 국회 입법과 지방 조례를 잇는다")
    bullet(doc, "국회 의안과 법령을 이름으로 연결해 수직 경로를 처음 측정하였다")
    body(doc,
        "국회 의안 19,847건에는 어느 법령을 다루었는지를 가리키는 식별자가 비어 있"
        "어 국회 데이터가 조례 그래프와 연결되지 않았다. 의안명에서 개정 접미를 제거하"
        "고 법령명과 정규화 비교한 결과 16,423건(82.7%)이 연결되었고, 고유 "
        "법령 1,252개 중 994개(79.4%)가 실제로 조례를 낳았다. 그 조례 인"
        "용은 210,938건으로 전체의 50.0%다.")
    figure(doc, FIG / "F7_국회_조례_연결.png",
           "그림 11. 국회에서 가장 많이 다뤄진 법령과 조례를 가장 많이 낳은 법령의 "
           "비교. 두 목록은 거의 겹치지 않는다. 국회법은 의안 261건이 제출되었으나 "
           "이를 인용한 조례는 1건이고, 국민기초생활 보장법은 조례 8,104건의 준거가 "
           "되었으나 의안 수 상위권에는 없다.", width_cm=15.5)
    body(doc,
         "순위 수준의 연관은 중간 정도이고(Spearman 0.491) 규모 수준의 연관은 "
         "약하며(Pearson 0.206), 국회에서 다뤄진 법령의 20.6%는 조례를 하나도 낳지 "
         "않았다. 국회가 법을 고치면 지방이 조례를 만든다는 단순한 그림은 성립하지 "
         "않으며, 두 입법 층위는 서로 다른 영역을 다루고 부분적으로만 연결된다.",
         indent=0.4)

    head_box(doc, "산출물 ⑧ — 지도·그래프·검색을 잇는 화면 16종")
    figure_pair(doc, SHOT / "02_map.png", SHOT / "05_graph.png",
                "그림 12·13. (좌) 시군구 코로플레스 — 조례 수·예산 집행률·분야 비중 등 "
                "지표를 바꿔가며 250개 행정경계 폴리곤에 표시하며, 조례 제정권이 없는 "
                "일반구·행정시는 상위 단위 값으로 채우고 점선으로 구분한다. "
                "(우) 법령 위계 그래프 — 법률에서 시행령·시행규칙·행정규칙을 거쳐 조례에 "
                "이르는 위계와 서브그래프 탐색.")
    figure_pair(doc, SHOT / "13_lifecycle.png", SHOT / "14_effectiveness.png",
                "그림 14·15. (좌) 정책 생애주기 — 제정·개정·폐지 이력과 행정구역 승계 추적. "
                "(우) 조례 실효성 — 조례에 연결된 예산 사업의 집행률. 연결은 확률적 매칭이므로 "
                "신뢰도 구간과 사람 확인 여부를 함께 표시한다.")
    figure_pair(doc, SHOT / "12_search.png", SHOT / "11_votes.png",
                "그림 16·17. (좌) 조문 전문검색 — 어휘검색과 벡터검색을 결합하고 그래프로 "
                "확장해 조문 본문에서 근거를 찾는다. (우) 국회 표결 — 의안별 정당 찬반 분해. "
                "개별 표결기록 보유 의안은 200건(1.0%)이며 이 사실을 화면에 함께 공시한다.")
    figure_pair(doc, SHOT / "04_gap.png", SHOT / "15_ai_agent.png",
                "그림 16·17. (좌) 유사·격차 분석 — 인구·재정 구조가 비슷한 유사 지방자치단체를 "
                "산출하고 그들이 보유했으나 기준 지자체에는 없는 조례를 상위법령 근거와 함께 "
                "제시한다. (우) AI 정책분석관 — 현재 화면과 정책 도구 결과를 함께 읽어 답하며, "
                "예산 연결처럼 확인된 값과 추정값이 섞인 항목은 이를 구분해 설명한다.")

    # =====================================================================
    chapter(doc, 3, "활용분야")

    head_box(doc, "지방자치단체 조례 담당자 — 입안 근거 조사")
    bullet(doc, "유사 지방자치단체 비교에서 시작해 상위법령 근거와 예산 연결까지 "
                "한 흐름으로 확인한다")
    body(doc,
        "담당자는 격차 분석에서 후보를 좁히고, 생애주기 화면에서 다른 지방자치단체의 폐"
        "지 사례를 확인한 뒤, 실효성 화면에서 예산이 실제로 붙는지 본다. 각 단계의 "
        "결과에 원문 링크가 함께 제시되므로 화면의 결론을 그대로 받아들이지 않고 근거를"
        " 직접 확인할 수 있다.")
    body(doc,
         "AI 정책분석관은 계산을 대신하지 않는다. 분석 모듈이 산출한 값을 근거로 "
         "받아 설명만 담당하므로 수치를 임의로 생성할 수 없으며, 확인된 값과 추정값을 "
         "구분해 말하도록 규칙을 두었다. 위 화면에서도 예산 연결에 대해 확인된 것과 "
         "추정에 해당하는 것을 나누어 답하고 있다.", indent=0.4)

    head_box(doc, "신설 지방자치단체 — 조례 백로그 산출")
    bullet(doc, "전남광주통합특별시 사례에서 우선순위 실행목록을 자동 산출하였다")
    body(doc,
        "2026년 7월 1일 시행된 전남광주통합특별시 설치 특별법은 408개 조문 중 "
        "65개에서 조례로 정하도록 위임한다. 시행 7주 시점에 제정이 확인된 것은 9개"
        "였고 현행 자치법규는 421건으로 전국 광역 중 가장 적었다. 본 시스템은 고아"
        " 위임 조항을 조번호 단위로 나열하고, 조례가 없으면 기관이 작동하지 않는 조직"
        "·기구 조항과 조례유보원칙상 지급이 불가능한 주민 금전급부 조항을 우선순위 상단"
        "에 배치한다.")

    head_box(doc, "자치법규 정비 — 판단이 필요 없는 지표")
    bullet(doc, "2005년 이전 구법 표현이 남아 있는 조례를 조문 단위로 특정한다")
    body(doc,
        "기초자치단체 현행 조문 1,895,738건을 전수 검색한 결과 국가법령에서는 사"
        "라진 구법 표현 '각호의 1에 해당'이 196개 지방자치단체의 조례 1,857건"
        "에 남아 있었다. 결손 판정과 달리 문자열 존재 여부이므로 허위 양성이 원리적으"
        "로 거의 없으며, 지방자치단체는 이 목록으로 일괄 개정 의안 한 건을 처리할 수"
        " 있다.")

    head_box(doc, "연구·언론·시민 및 융복합 확장")
    bullet(doc, "전국 전량 데이터와 신뢰도 지표를 함께 공개하여 재현 가능한 연구 기반을 "
                "제공하고, 같은 구조를 다른 규범 문서로 확장한다")
    body(doc,
         "데이터와 코드는 공개 저장소에 MIT 라이선스로 공개되어 있어 연구자가 동일한 "
         "분석을 재현하거나 다른 가설을 검증할 수 있고, 언론과 시민은 자기 지방자치단체가 "
         "무엇을 하지 않고 있는지를 근거와 함께 확인할 수 있다. 그래프 구조가 '규범 문서 - "
         "발행 주체 - 상위 근거 - 예산' 의 관계로 일반화되므로 조례 대신 주요업무계획이나 "
         "행정규칙을 노드로 삼아도 동일하게 동작하며, 행정경계와 결합하는 방식도 그대로 "
         "적용된다.")

    chapter(doc, 4, "기대효과")

    head_box(doc, "수혜자 범위")
    bullet(doc, "1차 수혜자는 전국 227개 기초자치단체와 17개 광역자치단체의 "
                "조례 담당 부서 및 지방의회 입법 담당자다")
    body(doc,
         "특히 조례 담당 인력이 적은 지방자치단체일수록 선례 조사에 드는 부담이 크므로, "
         "전국 전량 데이터를 균등하게 제공하는 것만으로 행정 역량의 격차가 완화되는 "
         "효과를 기대할 수 있다. 2차 수혜자는 중앙행정기관(위임 이행 현황 점검), "
         "연구자(재현 가능한 전국 데이터), 언론과 시민(정책 감시)이다.")

    head_box(doc, "정성적 기대효과")
    bullet(doc, "근거 기반 입법 관행의 확산")
    body(doc,
         "본 시스템의 모든 추천에는 상위법령 조문, 유사 지방자치단체 보유 현황, 예산 "
         "연결 여부가 근거로 함께 제시된다. 담당자가 '다른 곳도 하니까' 가 아니라 "
         "'어떤 법 조항에 근거해 몇 곳이 어떻게 만들었고 예산은 어떻게 붙었는가' 를 "
         "확인하고 판단하게 된다.")
    bullet(doc, "행정 정보 격차의 완화")
    body(doc,
         "조례 담당 인력과 예산이 적은 지방자치단체는 선례 조사에 투입할 여력이 "
         "부족하다. 전국 전량 데이터를 무료로 균등 제공하면 이 격차가 줄어들며, "
         "이는 지방자치단체 사이의 정책 품질 편차를 완화하는 방향으로 작용한다.")
    bullet(doc, "불확실성을 함께 제시하는 정보 제공 방식의 확립")
    body(doc,
         "행정 현장에서 데이터 기반 도구가 신뢰를 얻지 못하는 주된 이유는 정확도를 "
         "과장하기 때문이다. 본 시스템은 확인된 것과 추정한 것을 구분해 제시하고 "
         "자체 신뢰도를 화면에 공시하는 방식을 택하였다. 이는 공공 부문 데이터 서비스가 "
         "갖추어야 할 요건에 대한 하나의 제안이기도 하다.")

    head_box(doc, "정량적 기대효과")
    bullet(doc, "조사 단위의 축소 — 검색·비교·확인의 반복을 화면 조회로 대체한다")
    table(doc,
          ["항목", "현재", "본 시스템 적용 시"],
          [["유사 사례 조사", "검색·개별 열람·전화 확인의 반복", "격차 화면 1회 조회"],
           ["격차 후보 산출", "산출 수단 없음", "227곳 전수 자동 산출, 지자체당 한 자릿수"],
           ["정비 대상 특정", "육안 점검", "196곳 1,857건 조문 단위 즉시 목록화"],
           ["신설 지자체 백로그", "수작업 조문 대조", "의무위임 65개 우선순위 자동 산출"],
           ["감사 검색공간", "예산 479,683행 전수", "자체재원 경상 보조성 37,482행(7.8%)로 축소"]],
          widths=[3.6, 5.4, 6.3])
    body(doc,
         "본 신청서는 시간 절감의 절대 시간을 추정하지 않는다. 실사용자를 대상으로 한 "
         "측정을 거치지 않은 상태에서 분 단위 수치를 제시하면 근거를 요구받았을 때 "
         "답할 수 없기 때문이다. 대신 작업의 단위가 어떻게 바뀌는지를 위와 같이 "
         "제시하며, 시범 적용 기관이 확보되면 실제 소요 시간을 측정해 보완할 계획이다.",
         indent=0.4)
    bullet(doc, "운영 비용 — 서버 비용 없이 상시 서비스가 가능하다")
    body(doc,
         "전체 산출물을 정적 파일로 사전 생성하고 gzip 으로 압축해 81.2MB로 줄였기 "
         "때문에 별도의 애플리케이션 서버 없이 정적 호스팅만으로 전국 데이터가 "
         "동작한다. 데이터 갱신은 공개 API 를 호출하는 자동화 작업으로 처리되므로 "
         "지속적인 운영 부담이 낮다.")

    head_box(doc, "한계와 후속 과제")
    bullet(doc, "현 단계 산출물은 탐색적 데이터 기반이며 측정 도구가 아니다")
    table(doc,
          ["한계", "현재 값"],
          [["인용 조문 확인율", "43.3% (자동 확인 불가 49.9%)"],
           ["분야 코딩 신뢰도", "Cohen's κ 0.517"],
           ["이웃 추천 유효 모델", "3종 중 1종만 무작위 기준선 초과"],
           ["국회 개별 표결 커버리지", "200건 / 19,847건 (1.0%)"],
           ["결손 판정 허위율", "측정 방법에 따라 28~84%"],
           ["최초 제정일", "데이터에 존재하지 않음(현행 판본 공포일만 보유)"]],
          widths=[5.6, 9.7])
    body(doc,
        "정책 감시(policy surveillance) 분야의 요건은 독립 인간 코더 "
        "2인 이상, 조문 단위 변수, 신뢰도 보고, 불일치 해소 절차다. 현 단계에서는"
        " 이를 충족하지 않는다. 다만 규모가 국내 선행 연구와 비교할 수 없이 크고 신"
        "뢰도를 측정해 공개한다는 점에서 후속 연구의 기반이 된다. 향후 과제는 이중 코"
        "딩으로 분류 신뢰도를 높이고 상위법령 조문 확보 범위를 넓혀 인용 확인율을 개선"
        "하는 것이다.", indent=0.4)

    # =====================================================================
    chapter(doc, 5, "참고문헌 출처 등")

    head_box(doc, "데이터 출처 (전 항목 무료 공개 API)")
    para(doc, "1. 법제처, 국가법령정보 공동활용 Open API. https://open.law.go.kr    "
              "2. 국회사무처, 열린국회정보 Open API. https://open.assembly.go.kr    "
              "3. 국토교통부, 공간정보 오픈플랫폼(브이월드) 데이터 API. https://www.vworld.kr    "
              "4. 행정안전부, 행정표준코드관리시스템(공공데이터포털). https://www.data.go.kr    "
              "5. 행정안전부, 지방재정365 재정데이터개방. https://lofin365.go.kr    "
              "6. southkorea-maps, 대한민국 시군구 행정경계 GeoJSON(2018 단순화본). "
              "https://github.com/southkorea/southkorea-maps",
         size=SMALL_PT, indent=0.4, space_after=4, line=1.24)

    head_box(doc, "방법론 문헌")
    para(doc, "7. Walker, J. L. (1969). The Diffusion of Innovations among the American "
              "States. APSR, 63(3).    "
              "8. Berry, F. S., & Berry, W. D. (1990). State Lottery Adoptions as Policy "
              "Innovations. APSR, 84(2).    "
              "9. Shipan, C. R., & Volden, C. (2008). The Mechanisms of Policy Diffusion. "
              "AJPS, 52(4).    "
              "10. DiMaggio, P. J., & Powell, W. W. (1983). The Iron Cage Revisited. ASR, 48(2).    "
              "11. Anselin, L. (1995). Local Indicators of Spatial Association-LISA. "
              "Geographical Analysis, 27(2).    "
              "12. Benjamini, Y., & Hochberg, Y. (1995). Controlling the False Discovery "
              "Rate. JRSS-B, 57(1).    "
              "13. Fowler, J. H., & Jeon, S. (2008). The Authority of Supreme Court "
              "Precedent. Social Networks, 30(1).    "
              "14. Burris, S., et al. (2016). Policy Surveillance. J Health Polit Policy "
              "Law, 41(6).    "
              "15. Linder, F., et al. (2020). Text as Policy: Measuring Policy Similarity "
              "through Bill Text Reuse. PSJ, 48(2).    "
              "16. Grover, A., & Leskovec, J. (2016). node2vec. KDD.    "
              "17. Dong, Y., et al. (2017). metapath2vec. KDD.    "
              "18. Hamilton, W., et al. (2017). Inductive Representation Learning on Large "
              "Graphs. NeurIPS.    "
              "19. Cormack, G. V., et al. (2009). Reciprocal Rank Fusion. SIGIR.",
         size=SMALL_PT, indent=0.4, space_after=4, line=1.24)

    head_box(doc, "식별자·문서 표준 및 산출물")
    para(doc, "20. Publications Office of the EU. European Legislation Identifier (ELI).    "
              "21. OASIS. Akoma Ntoso Version 1.0 (LegalDocML).    "
              "22. IFLA. Functional Requirements for Bibliographic Records (FRBR).    "
              "23. 본 신청서의 모든 그림은 본 팀이 구축한 시스템의 실제 화면과 실측 데이터에서 "
              "직접 생성한 것이며 외부 이미지를 인용하지 않았다. 화면 캡처는 docs/screenshots, "
              "분석 도표는 docs/figures 에 재생성 스크립트와 함께 공개되어 있다.    "
              "24. 공개 저장소 https://github.com/zxsa0716/policymaps (MIT License)    "
              "25. 배포 사이트 https://policymaps.vercel.app",
         size=SMALL_PT, indent=0.4, space_after=4, line=1.24)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUT)
    print("  → %s  (%.0fKB)" % (OUT, OUT.stat().st_size / 1024))
    return 0


if __name__ == "__main__":
    raise SystemExit(build())
