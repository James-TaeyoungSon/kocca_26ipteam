"""
generate_and_upload_hwp_py.py
=============================
순수 Python으로 HWPX 파일 생성 → Notion DB 업로드.

- HWP COM / 한컴오피스 불필요 → Linux/GitHub Actions 실행 가능
- 의존 패키지: requests 만 필요
- generate_and_upload_xlsx.py 와 동일한 payload 인터페이스
"""

import base64
import csv
import io
import json
import os
import re
import sys
import zipfile
from datetime import date, datetime, timedelta
from typing import Any, Dict, List, Optional, Tuple

import requests

NOTION_VERSION = "2025-09-03"
DAYS_KR = ['월', '화', '수', '목', '금', '토', '일']

DAY_COLORS: List[Tuple[Tuple[int, int, int], Tuple[int, int, int]]] = [
    ((68, 114, 196),  (214, 228, 247)),   # 월 파랑
    ((237, 125, 49),  (252, 228, 214)),   # 화 주황
    ((112, 173, 71),  (226, 239, 218)),   # 수 초록
    ((158, 72, 14),   (249, 208, 190)),   # 목 갈색
    ((112, 48, 160),  (232, 213, 245)),   # 금 보라
]

# ── HWP 단위 ─────────────────────────────────────────────────────────────
# 1mm = 283.46 HWP 단위 (= 7200/25.4)
# A4 landscape="NARROWLY" 모드: physical width=59528(210mm), height=84186(297mm)
# → 실제 렌더링: 가로 84186, 세로 59528 (장축이 가로)
_PG_W   = 59528   # A4 단축
_PG_H   = 84186   # A4 장축
_MARGIN = 4251    # 15mm
_USABLE = _PG_H - 2 * _MARGIN   # 75684 HWP 단위 ≈ 267mm (landscape 가로 사용 폭)

# ── borderFill ID 약속 ────────────────────────────────────────────────────
_FID_NO_BORDER = 1   # 테두리 없음 (페이지 보더용)
_FID_CHARPR    = 2   # charPr borderFillIDRef 기본값
_FID_TABLE     = 3   # SOLID 테두리, 채움 없음 (표 기본)
_FID_DAY_D     = [4, 5, 6, 7, 8]      # 월~금 짙은 배경
_FID_DAY_L     = [9, 10, 11, 12, 13]  # 월~금 연한 배경
_FID_GRAY      = 14
_FID_YELLOW    = 15

# ── charPr ID 약속 ────────────────────────────────────────────────────────
_CP_BASE    = 0    # 기본
_CP_TITLE   = 7    # 18pt bold black
_CP_LABEL   = 8    # 11pt bold #404040
_CP_CONTENT = 9    # 10pt black
_CP_CAL_HD  = 10   # 13pt bold white
_CP_TIME    = 11   # 9pt #646464
_CP_CAT_DAY = [12, 13, 14, 15, 16]   # 월~금 카테고리색
_CP_DET_HD  = 17   # 12pt bold black
_CP_GUBUN   = 18   # 11pt bold black

# ── paraPr ID 약속 ───────────────────────────────────────────────────────
_PP_JUSTIFY = 0    # 양쪽정렬 (기본)
_PP_CENTER  = 20   # 가운데정렬 (새로 추가)


# ════════════════════════════════════════════════════════════════════════════
# XML 헬퍼
# ════════════════════════════════════════════════════════════════════════════

def _e(text: str) -> str:
    return (text.replace('&', '&amp;').replace('<', '&lt;')
                .replace('>', '&gt;').replace('"', '&quot;'))


def _rgb_hex(rgb: Tuple[int, int, int]) -> str:
    return f'#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}'


# ════════════════════════════════════════════════════════════════════════════
# header.xml 생성
# ════════════════════════════════════════════════════════════════════════════

_NS = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)

def _bf_solid(bf_id: int, color_hex: Optional[str] = None) -> str:
    fill = ''
    if color_hex:
        fill = f'<hc:fillBrush><hc:winBrush faceColor="{color_hex}" hatchColor="#000000" alpha="0"/></hc:fillBrush>'
    return (
        f'<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
        '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
        '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
        f'{fill}'
        '</hh:borderFill>'
    )

def _bf_none(bf_id: int, with_fill: bool = False) -> str:
    fill = '<hc:fillBrush><hc:winBrush faceColor="none" hatchColor="#999999" alpha="0"/></hc:fillBrush>' if with_fill else ''
    return (
        f'<hh:borderFill id="{bf_id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">'
        '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
        '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
        '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
        f'{fill}'
        '</hh:borderFill>'
    )

def _charpr(cp_id: int, height: int, color: str, bold: bool = False, font_ref: str = 'han') -> str:
    # font_ref: 'han'  → hangul=0, latin=0, hanja=1, ... (charPr 7-18)
    #           'han2' → hangul=1, latin=1, hanja=0, ... (charPr 1,2,4,5,6)
    #           'base' → hangul=2, latin=2, hanja=1, ... (charPr 0,3)
    if font_ref == 'han':
        fref = 'hangul="0" latin="0" hanja="1" japanese="1" other="1" symbol="1" user="1"'
    elif font_ref == 'han2':
        fref = 'hangul="1" latin="1" hanja="0" japanese="0" other="0" symbol="0" user="0"'
    else:  # 'base'
        fref = 'hangul="2" latin="2" hanja="1" japanese="1" other="1" symbol="1" user="1"'
    b = '<hh:bold/>' if bold else ''
    return (
        f'<hh:charPr id="{cp_id}" height="{height}" textColor="{color}" shadeColor="none" '
        f'useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">'
        f'<hh:fontRef {fref}/>'
        '<hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
        '<hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
        '<hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>'
        '<hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>'
        f'{b}'
        '<hh:underline type="NONE" shape="SOLID" color="#000000"/>'
        '<hh:strikeout shape="NONE" color="#000000"/>'
        '<hh:outline type="NONE"/>'
        '<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/>'
        '</hh:charPr>'
    )

def _parapr(pp_id: int, align: str = 'JUSTIFY', left: int = 0, right: int = 0,
            prev: int = 0, nxt: int = 0, spacing: int = 160,
            intent: int = 0, condense: int = 0) -> str:
    m = (
        f'<hh:margin>'
        f'<hc:intent value="{intent}" unit="HWPUNIT"/>'
        f'<hc:left value="{left}" unit="HWPUNIT"/>'
        f'<hc:right value="{right}" unit="HWPUNIT"/>'
        f'<hc:prev value="{prev}" unit="HWPUNIT"/>'
        f'<hc:next value="{nxt}" unit="HWPUNIT"/>'
        f'</hh:margin>'
        f'<hh:lineSpacing type="PERCENT" value="{spacing}" unit="HWPUNIT"/>'
    )
    return (
        f'<hh:paraPr id="{pp_id}" tabPrIDRef="0" condense="{condense}" fontLineHeight="0" '
        f'snapToGrid="1" suppressLineNumbers="0" checked="0">'
        f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
        '<hh:heading type="NONE" idRef="0" level="0"/>'
        '<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" '
        'widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>'
        '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
        f'<hp:switch>'
        f'<hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">{m}</hp:case>'
        f'<hp:default>{m}</hp:default>'
        f'</hp:switch>'
        '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" '
        'offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
        '</hh:paraPr>'
    )


def _make_header_xml() -> bytes:
    parts: List[str] = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>')
    parts.append(f'<hh:head {_NS} version="1.5" secCnt="1">')
    parts.append('<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>')
    parts.append('<hh:refList>')

    # ── 폰트 ──────────────────────────────────────────────────────────────
    f2 = lambda i, face, w, p: (
        f'<hh:font id="{i}" face="{face}" type="TTF" isEmbedded="0">'
        f'<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="{p}" contrast="0" '
        f'strokeVariation="1" armStyle="1" letterform="1" midline="1" xHeight="1"/>'
        f'</hh:font>'
    )
    fonts3 = (
        f2(0, '한컴바탕', 6, 0) + f2(1, '함초롬돋움', 6, 4) + f2(2, '함초롬바탕', 6, 4)
    )
    fonts2 = f2(0, '함초롬돋움', 6, 4) + f2(1, '함초롬바탕', 6, 4)

    parts.append('<hh:fontfaces itemCnt="7">')
    for lang, cnt, fdata in [
        ('HANGUL', 3, fonts3), ('LATIN', 3, fonts3),
        ('HANJA', 2, fonts2), ('JAPANESE', 2, fonts2),
        ('OTHER', 2, fonts2), ('SYMBOL', 2, fonts2), ('USER', 2, fonts2),
    ]:
        parts.append(f'<hh:fontface lang="{lang}" fontCnt="{cnt}">{fdata}</hh:fontface>')
    parts.append('</hh:fontfaces>')

    # ── borderFills (1~15) ────────────────────────────────────────────────
    bf_list = [
        _bf_none(1),          # ID 1: 테두리 없음
        _bf_none(2, True),    # ID 2: charPr 기본 (faceColor=none)
        _bf_solid(3),         # ID 3: 표 기본 (SOLID, 채움 없음)
    ]
    for i, (dark, _light) in enumerate(DAY_COLORS):
        bf_list.append(_bf_solid(4 + i, _rgb_hex(dark)))    # 4~8: 요일 짙은
    for i, (_dark, light) in enumerate(DAY_COLORS):
        bf_list.append(_bf_solid(9 + i, _rgb_hex(light)))   # 9~13: 요일 연한
    bf_list.append(_bf_solid(14, '#D9D9D9'))   # 14: 회색
    bf_list.append(_bf_solid(15, '#FFFF00'))   # 15: 노랑

    parts.append(f'<hh:borderFills itemCnt="{len(bf_list)}">')
    parts.extend(bf_list)
    parts.append('</hh:borderFills>')

    # ── charProperties (0-18) ─────────────────────────────────────────────
    parts.append('<hh:charProperties itemCnt="19">')
    # 0-6: 시스템 기본 charPr
    parts.append(_charpr(0, 1000, '#000000', font_ref='base'))
    parts.append(_charpr(1, 1000, '#000000', font_ref='han2'))
    parts.append(_charpr(2, 900,  '#000000', font_ref='han2'))
    parts.append(_charpr(3, 900,  '#000000', font_ref='base'))
    parts.append(_charpr(4, 900,  '#000000', font_ref='han2'))
    parts.append(_charpr(5, 1600, '#2E74B5', font_ref='han2'))
    parts.append(_charpr(6, 1100, '#000000', font_ref='han2'))
    # 7-18: 보고서 전용
    parts.append(_charpr(7,  1800, '#000000', bold=True))
    parts.append(_charpr(8,  1100, '#404040', bold=True))
    parts.append(_charpr(9,  1000, '#000000'))
    parts.append(_charpr(10, 1300, '#FFFFFF', bold=True))
    parts.append(_charpr(11, 900,  '#646464'))
    parts.append(_charpr(12, 1100, '#4472C4', bold=True))
    parts.append(_charpr(13, 1100, '#ED7D31', bold=True))
    parts.append(_charpr(14, 1100, '#70AD47', bold=True))
    parts.append(_charpr(15, 1100, '#9E480E', bold=True))
    parts.append(_charpr(16, 1100, '#7030A0', bold=True))
    parts.append(_charpr(17, 1200, '#000000', bold=True))
    parts.append(_charpr(18, 1100, '#000000', bold=True))
    parts.append('</hh:charProperties>')

    # ── tabProperties ─────────────────────────────────────────────────────
    parts.append(
        '<hh:tabProperties itemCnt="3">'
        '<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>'
        '<hh:tabPr id="1" autoTabLeft="1" autoTabRight="0"/>'
        '<hh:tabPr id="2" autoTabLeft="0" autoTabRight="1"/>'
        '</hh:tabProperties>'
    )

    # ── numberings ────────────────────────────────────────────────────────
    parts.append(
        '<hh:numberings itemCnt="1"><hh:numbering id="1" start="0">'
        '<hh:paraHead start="1" level="1" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" '
        'charPrIDRef="4294967295" checkable="0">^1.</hh:paraHead>'
        '<hh:paraHead start="1" level="2" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_SYLLABLE" '
        'charPrIDRef="4294967295" checkable="0">^2.</hh:paraHead>'
        '<hh:paraHead start="1" level="3" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" '
        'charPrIDRef="4294967295" checkable="0">^3)</hh:paraHead>'
        '<hh:paraHead start="1" level="4" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_SYLLABLE" '
        'charPrIDRef="4294967295" checkable="0">^4)</hh:paraHead>'
        '<hh:paraHead start="1" level="5" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="DIGIT" '
        'charPrIDRef="4294967295" checkable="0">(^5)</hh:paraHead>'
        '<hh:paraHead start="1" level="6" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_SYLLABLE" '
        'charPrIDRef="4294967295" checkable="0">(^6)</hh:paraHead>'
        '<hh:paraHead start="1" level="7" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="CIRCLED_DIGIT" '
        'charPrIDRef="4294967295" checkable="1">^7</hh:paraHead>'
        '<hh:paraHead start="1" level="8" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="CIRCLED_HANGUL_SYLLABLE" '
        'charPrIDRef="4294967295" checkable="1">^8</hh:paraHead>'
        '<hh:paraHead start="1" level="9" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="HANGUL_JAMO" '
        'charPrIDRef="4294967295" checkable="0"/>'
        '<hh:paraHead start="1" level="10" align="LEFT" useInstWidth="1" autoIndent="1" '
        'widthAdjust="0" textOffsetType="PERCENT" textOffset="50" numFormat="ROMAN_SMALL" '
        'charPrIDRef="4294967295" checkable="1"/>'
        '</hh:numbering></hh:numberings>'
    )

    # ── paraProperties (0-20: 기본 20개 + CENTER 1개) ─────────────────────
    parts.append('<hh:paraProperties itemCnt="21">')
    # id=0-8: outline 스타일 (다양한 들여쓰기)
    parts.append(_parapr(0))
    parts.append(_parapr(1, left=1500))
    parts.append(_parapr(2, left=1000, condense=20))
    parts.append(_parapr(3, left=2000, condense=20))
    parts.append(_parapr(4, left=3000, condense=20))
    parts.append(_parapr(5, left=4000, condense=20))
    parts.append(_parapr(6, left=5000, condense=20))
    parts.append(_parapr(7, left=6000, condense=20))
    parts.append(_parapr(8, left=7000, condense=20))
    # id=9: 머리말 (150% 행간)
    parts.append(_parapr(9, spacing=150))
    # id=10: 각주 (130% 행간, 들여쓰기 -1310)
    parts.append(_parapr(10, spacing=130, intent=-1310))
    # id=11-15: 기타
    parts.append(_parapr(11, align='LEFT', spacing=130))
    parts.append(_parapr(12, align='LEFT', prev=1200, nxt=300))
    parts.append(_parapr(13, align='LEFT', nxt=700))
    parts.append(_parapr(14, align='LEFT', left=1100, nxt=700))
    parts.append(_parapr(15, align='LEFT', left=2200, nxt=700))
    # id=16-19: outline 9~10 + misc
    parts.append(_parapr(16, left=9000))
    parts.append(_parapr(17, left=10000))
    parts.append(_parapr(18, left=8000))
    parts.append(_parapr(19, spacing=150, nxt=800))
    # id=20: CENTER 정렬 (제목용)
    parts.append(_parapr(20, align='CENTER'))
    parts.append('</hh:paraProperties>')

    # ── styles ────────────────────────────────────────────────────────────
    parts.append('<hh:styles itemCnt="22">')
    sty = [
        ('바탕글', 'Normal', 0, 0, 0), ('본문', 'Body', 1, 0, 1),
        ('개요 1', 'Outline 1', 2, 0, 2), ('개요 2', 'Outline 2', 3, 0, 3),
        ('개요 3', 'Outline 3', 4, 0, 4), ('개요 4', 'Outline 4', 5, 0, 5),
        ('개요 5', 'Outline 5', 6, 0, 6), ('개요 6', 'Outline 6', 7, 0, 7),
        ('개요 7', 'Outline 7', 8, 0, 8), ('개요 8', 'Outline 8', 18, 0, 9),
        ('개요 9', 'Outline 9', 16, 0, 10), ('개요 10', 'Outline 10', 17, 0, 11),
    ]
    for i, (nm, en, pp, cp, nxt) in enumerate(sty):
        parts.append(
            f'<hh:style id="{i}" type="PARA" name="{nm}" engName="{en}" '
            f'paraPrIDRef="{pp}" charPrIDRef="{cp}" nextStyleIDRef="{nxt}" '
            f'langID="1042" lockForm="0"/>'
        )
    parts.append(
        '<hh:style id="12" type="CHAR" name="쪽 번호" engName="Page Number" '
        'paraPrIDRef="0" charPrIDRef="1" nextStyleIDRef="0" langID="1042" lockForm="0"/>'
        '<hh:style id="13" type="PARA" name="머리말" engName="Header" '
        'paraPrIDRef="9" charPrIDRef="2" nextStyleIDRef="13" langID="1042" lockForm="0"/>'
        '<hh:style id="14" type="PARA" name="각주" engName="Footnote" '
        'paraPrIDRef="10" charPrIDRef="3" nextStyleIDRef="14" langID="1042" lockForm="0"/>'
        '<hh:style id="15" type="PARA" name="미주" engName="Endnote" '
        'paraPrIDRef="10" charPrIDRef="3" nextStyleIDRef="15" langID="1042" lockForm="0"/>'
        '<hh:style id="16" type="PARA" name="메모" engName="Memo" '
        'paraPrIDRef="11" charPrIDRef="4" nextStyleIDRef="16" langID="1042" lockForm="0"/>'
        '<hh:style id="17" type="PARA" name="차례 제목" engName="TOC Heading" '
        'paraPrIDRef="12" charPrIDRef="5" nextStyleIDRef="17" langID="1042" lockForm="0"/>'
        '<hh:style id="18" type="PARA" name="차례 1" engName="TOC 1" '
        'paraPrIDRef="13" charPrIDRef="6" nextStyleIDRef="18" langID="1042" lockForm="0"/>'
        '<hh:style id="19" type="PARA" name="차례 2" engName="TOC 2" '
        'paraPrIDRef="14" charPrIDRef="6" nextStyleIDRef="19" langID="1042" lockForm="0"/>'
        '<hh:style id="20" type="PARA" name="차례 3" engName="TOC 3" '
        'paraPrIDRef="15" charPrIDRef="6" nextStyleIDRef="20" langID="1042" lockForm="0"/>'
        '<hh:style id="21" type="PARA" name="캡션" engName="Caption" '
        'paraPrIDRef="19" charPrIDRef="0" nextStyleIDRef="21" langID="1042" lockForm="0"/>'
    )
    parts.append('</hh:styles>')
    parts.append('</hh:refList>')
    parts.append(
        '<hh:compatibleDocument targetProgram="HWP201X">'
        '<hh:layoutCompatibility/></hh:compatibleDocument>'
        '<hh:docOption>'
        '<hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/>'
        '</hh:docOption>'
        '<hh:trackchageConfig flags="56"/>'
    )
    parts.append('</hh:head>')
    return ''.join(parts).encode('utf-8')


# ════════════════════════════════════════════════════════════════════════════
# section0.xml 생성
# ════════════════════════════════════════════════════════════════════════════

_SEC_PR = (
    '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" '
    'tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" '
    'textVerticalWidthHead="0" masterPageCnt="0">'
    '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
    '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
    '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" '
    'border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
    '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
    # landscape="NARROWLY" + A4 물리 치수 → HWP가 가로(landscape)로 렌더링
    f'<hp:pagePr landscape="NARROWLY" width="{_PG_W}" height="{_PG_H}" gutterType="LEFT_ONLY">'
    f'<hp:margin header="0" footer="0" gutter="0" '
    f'left="{_MARGIN}" right="{_MARGIN}" top="{_MARGIN}" bottom="{_MARGIN}"/>'
    '</hp:pagePr>'
    '<hp:footNotePr>'
    '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
    '<hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
    '<hp:numbering type="CONTINUOUS" newNum="1"/>'
    '<hp:placement place="EACH_COLUMN" beneathText="0"/>'
    '</hp:footNotePr>'
    '<hp:endNotePr>'
    '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
    '<hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>'
    '<hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
    '<hp:numbering type="CONTINUOUS" newNum="1"/>'
    '<hp:placement place="END_OF_DOCUMENT" beneathText="0"/>'
    '</hp:endNotePr>'
    '<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER" '
    'headerInside="0" footerInside="0" fillArea="PAPER">'
    '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
    '</hp:pageBorderFill>'
    '<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER" '
    'headerInside="0" footerInside="0" fillArea="PAPER">'
    '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
    '</hp:pageBorderFill>'
    '<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER" '
    'headerInside="0" footerInside="0" fillArea="PAPER">'
    '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
    '</hp:pageBorderFill>'
    '</hp:secPr>'
    '<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" colCount="1" sameSz="1" sameGap="0"/></hp:ctrl>'
)

def _para(cp_id: int, text: str, pp_id: int = _PP_JUSTIFY,
          first: bool = False) -> str:
    """단순 텍스트 단락."""
    secpr = _SEC_PR if first else ''
    t = f'<hp:t>{_e(text)}</hp:t>' if text else ''
    if first:
        return (
            f'<hp:p id="1" paraPrIDRef="{pp_id}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="{_CP_BASE}">{secpr}</hp:run>'
            f'<hp:run charPrIDRef="{cp_id}">{t}</hp:run>'
            '</hp:p>'
        )
    return (
        f'<hp:p id="0" paraPrIDRef="{pp_id}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp_id}">{t}</hp:run>'
        '</hp:p>'
    )

def _cell(col: int, row: int, fid: int, lines: List[Tuple[str, int, int]],
          width: int, height: int = 500) -> str:
    """
    표 셀 XML.
    lines: [(text, charPrId, paraPrId), ...]
    """
    paras = ''.join(
        f'<hp:p id="0" paraPrIDRef="{pp}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{cp}">'
        f'{"<hp:t>" + _e(txt) + "</hp:t>" if txt else ""}'
        f'</hp:run>'
        f'</hp:p>'
        for txt, cp, pp in lines
    )
    return (
        f'<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="{fid}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" '
        f'linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'{paras}'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="{col}" rowAddr="{row}"/>'
        f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
        f'<hp:cellSz width="{width}" height="{height}"/>'
        f'<hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
        f'</hp:tc>'
    )

def _table(rows_data: List[List[Tuple]], col_widths: List[int]) -> str:
    """
    rows_data: [ row, row, ... ]
      row:      [ (fid, lines), (fid, lines), ... ]  per cell
      lines:    [(text, charPrId, paraPrId), ...]
    col_widths: HWP units per column
    """
    total_w = sum(col_widths)
    n_rows = len(rows_data)
    n_cols = len(col_widths)

    tr_parts = []
    for r_idx, row in enumerate(rows_data):
        tcs = ''
        for c_idx, (fid, lines) in enumerate(row):
            w = col_widths[c_idx] if c_idx < len(col_widths) else col_widths[-1]
            tcs += _cell(c_idx, r_idx, fid, lines, w)
        tr_parts.append(f'<hp:tr>{tcs}</hp:tr>')

    tbl_xml = (
        f'<hp:tbl id="0" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" '
        f'textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" '
        f'repeatHeader="1" rowCnt="{n_rows}" colCnt="{n_cols}" cellSpacing="0" '
        f'borderFillIDRef="{_FID_TABLE}" noAdjust="0">'
        f'<hp:sz width="{total_w}" widthRelTo="ABSOLUTE" height="1000" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
        f'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" '
        f'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="283" right="283" top="283" bottom="283"/>'
        f'<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
        + ''.join(tr_parts) +
        f'</hp:tbl>'
    )
    return (
        f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">'
        f'<hp:run charPrIDRef="{_CP_BASE}">{tbl_xml}</hp:run>'
        '</hp:p>'
    )


def _make_section_xml(
    title: str,
    day_dates: List[date],
    day_events: Dict[date, List[Tuple[str, str, str]]],
    header: List[str],
    data_rows: List[List[str]],
    gi: int, ii: int, ci: int, bi: int,
) -> bytes:
    parts: List[str] = []
    parts.append('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>')
    parts.append(f'<hs:sec {_NS}>')

    # ── 제목 ──────────────────────────────────────────────────────────────
    parts.append(_para(_CP_TITLE, title, pp_id=_PP_CENTER, first=True))

    # ── 캘린더 레이블 ─────────────────────────────────────────────────────
    parts.append(_para(_CP_LABEL, '■ 주간 일정 캘린더'))

    # ── 캘린더 표 ────────────────────────────────────────────────────────
    max_slots = max((len(v) for v in day_events.values()), default=1)
    max_slots = max(max_slots, 2)

    # 5등분 (나머지는 마지막 열에)
    base_w = _USABLE // 5
    rem    = _USABLE - base_w * 5
    cal_ws = [base_w] * 4 + [base_w + rem]

    cal_rows: List[List[Tuple]] = []
    # 헤더 행
    hdr_row = []
    for i, (d, dk) in enumerate(zip(day_dates, ['월', '화', '수', '목', '금'])):
        label = f'{dk}  {d.month}/{d.day}'
        hdr_row.append((_FID_GRAY, [(label, _CP_DET_HD, _PP_CENTER)]))
    cal_rows.append(hdr_row)

    # 이벤트 행
    for slot in range(max_slots):
        ev_row = []
        for i, d in enumerate(day_dates):
            events = day_events.get(d, [])
            lines: List[Tuple[str, int, int]] = []
            if slot < len(events):
                gubun, content, time_str = events[slot]
                if time_str:
                    lines.append((time_str, _CP_TIME, _PP_JUSTIFY))
                if gubun:
                    lines.append((f'[{gubun}]', _CP_CAT_DAY[i], _PP_JUSTIFY))
                if content:
                    lines.append((content, _CP_CONTENT, _PP_JUSTIFY))
            if not lines:
                lines = [('', _CP_CONTENT, _PP_JUSTIFY)]
            ev_row.append((_FID_DAY_L[i], lines))
        cal_rows.append(ev_row)

    parts.append(_table(cal_rows, cal_ws))

    # ── 구분 공백 ─────────────────────────────────────────────────────────
    parts.append(_para(_CP_CONTENT, ''))

    # ── 상세 레이블 ───────────────────────────────────────────────────────
    parts.append(_para(_CP_LABEL, '■ 상세 업무 내역'))

    # ── 상세 표 ───────────────────────────────────────────────────────────
    # 열 너비 (mm → HWP 단위, 합계=75684)
    mm_widths = {gi: 22, ii: 68, ci: 117, bi: 60}
    n_cols = len(header)
    mm_list = [mm_widths.get(c, 22) for c in range(n_cols)]
    total_mm = sum(mm_list)
    det_ws = [round(w / total_mm * _USABLE) for w in mm_list]
    # 반올림 오차 보정
    diff = _USABLE - sum(det_ws)
    det_ws[-1] += diff

    detail_rows: List[List[Tuple]] = []
    # 헤더 행
    hdr = []
    for c_idx, col_name in enumerate(header):
        hdr.append((_FID_GRAY, [(col_name, _CP_DET_HD, _PP_CENTER)]))
    detail_rows.append(hdr)

    # 데이터 행
    for row in data_rows:
        bigo_val = row[bi] if bi < len(row) else ''
        is_hbj = '본부장' in bigo_val
        dr = []
        for c_idx in range(n_cols):
            val = row[c_idx] if c_idx < len(row) else ''
            bg = _FID_YELLOW if (is_hbj and 1 <= c_idx <= 3) else _FID_TABLE
            if c_idx == gi:
                cp, pp = _CP_GUBUN, _PP_CENTER
            else:
                cp, pp = _CP_CONTENT, _PP_JUSTIFY
            dr.append((bg, [(val, cp, pp)]))
        detail_rows.append(dr)

    parts.append(_table(detail_rows, det_ws))

    # ── 마지막 빈 단락 ────────────────────────────────────────────────────
    parts.append(_para(_CP_CONTENT, ''))

    parts.append('</hs:sec>')
    return ''.join(parts).encode('utf-8')


# ════════════════════════════════════════════════════════════════════════════
# HWPX ZIP 조립
# ════════════════════════════════════════════════════════════════════════════

_MIMETYPE = b'application/hwp+zip'

_VERSION_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" '
    'tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1" buildNumber="0" '
    'os="1" xmlVersion="1.5" application="Hancom Office Hangul" '
    'appVersion="12, 0, 0, 0 WIN32LEWindows_10"/>'
)

_SETTINGS_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    '<ha:HWPApplicationSetting '
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
    '<ha:CaretPosition listIDRef="0" paraIDRef="1" pos="0"/>'
    '</ha:HWPApplicationSetting>'
)

_CONTAINER_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    '<ocf:container '
    'xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
    '<ocf:rootfiles>'
    '<ocf:rootfile full-path="Contents/content.hpf" '
    'media-type="application/hwpml-package+xml"/>'
    '<ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/>'
    '<ocf:rootfile full-path="META-INF/container.rdf" media-type="application/rdf+xml"/>'
    '</ocf:rootfiles>'
    '</ocf:container>'
)

_MANIFEST_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
)

_CONTAINER_RDF = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
    '<rdf:Description rdf:about="">'
    '<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" '
    'rdf:resource="Contents/header.xml"/>'
    '</rdf:Description>'
    '<rdf:Description rdf:about="Contents/header.xml">'
    '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/>'
    '</rdf:Description>'
    '<rdf:Description rdf:about="">'
    '<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#" '
    'rdf:resource="Contents/section0.xml"/>'
    '</rdf:Description>'
    '<rdf:Description rdf:about="Contents/section0.xml">'
    '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/>'
    '</rdf:Description>'
    '</rdf:RDF>'
)

def _content_hpf(title: str) -> str:
    now = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        f'<opf:package {_NS} version="" unique-identifier="" id="">'
        '<opf:metadata>'
        f'<opf:title>{_e(title)}</opf:title>'
        '<opf:language>ko</opf:language>'
        '<opf:meta name="creator" content="text">hwp_py</opf:meta>'
        '<opf:meta name="subject" content="text"/>'
        '<opf:meta name="description" content="text"/>'
        '<opf:meta name="lastsaveby" content="text">hwp_py</opf:meta>'
        f'<opf:meta name="CreatedDate" content="text">{now}</opf:meta>'
        f'<opf:meta name="ModifiedDate" content="text">{now}</opf:meta>'
        '<opf:meta name="date" content="text"/>'
        '<opf:meta name="keyword" content="text"/>'
        '</opf:metadata>'
        '<opf:manifest>'
        '<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>'
        '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
        '<opf:item id="settings" href="settings.xml" media-type="application/xml"/>'
        '</opf:manifest>'
        '<opf:spine>'
        '<opf:itemref idref="header" linear="yes"/>'
        '<opf:itemref idref="section0" linear="yes"/>'
        '</opf:spine>'
        '</opf:package>'
    )


def _pack_hwpx(header_xml: bytes, section_xml: bytes, preview_text: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        # mimetype: 압축 없이 (ZIP spec 요구)
        mi = zipfile.ZipInfo('mimetype')
        mi.compress_type = zipfile.ZIP_STORED
        zf.writestr(mi, _MIMETYPE)
        zf.writestr('version.xml',             _VERSION_XML)
        zf.writestr('Contents/header.xml',     header_xml)
        zf.writestr('Contents/section0.xml',   section_xml)
        zf.writestr('settings.xml',            _SETTINGS_XML)
        zf.writestr('META-INF/container.xml',  _CONTAINER_XML)
        zf.writestr('META-INF/manifest.xml',   _MANIFEST_XML)
        zf.writestr('META-INF/container.rdf',  _CONTAINER_RDF)
        zf.writestr('Contents/content.hpf',    _content_hpf(preview_text))
        zf.writestr('Preview/PrvText.txt',     preview_text.encode('utf-8'))
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# 날짜·CSV 파싱 (generate_and_upload_xlsx.py 와 동일 로직)
# ════════════════════════════════════════════════════════════════════════════

def unescape_csv_text(text: str) -> str:
    return (
        text.replace('\\r\\n', '\n').replace('\\n', '\n')
            .replace('\\r', '\n').replace('\\"', '"').replace('\\\\', '\\')
    )


def _parse_start_date(iljeong: str):
    m = re.match(r'(\d{4})\.(\d{2})\.(\d{2})', iljeong.strip())
    return date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None


def _parse_time_range(iljeong: str) -> str:
    m = re.match(
        r'\d{4}\.\d{2}\.\d{2}\([^)]+\)\s*(\d{2}):(\d{2})'
        r'\s*~\s*\d{4}\.\d{2}\.\d{2}\([^)]+\)\s*(\d{2}):(\d{2})',
        iljeong.strip(),
    )
    if not m:
        return ''
    sh, sm, eh, em = m.groups()
    return '종일' if (sh, sm, eh, em) == ('00', '00', '00', '00') else f'{sh}:{sm}~{eh}:{em}'


def _fix_allday(value: str) -> str:
    if not isinstance(value, str):
        return value
    m = re.match(
        r'(\d{4})\.(\d{2})\.(\d{2})\([^)]+\)\s*(\d{2}):(\d{2})'
        r'\s*~\s*(\d{4})\.(\d{2})\.(\d{2})\([^)]+\)\s*(\d{2}):(\d{2})',
        value.strip(),
    )
    if m:
        sy, smo, sd, sh, smin, ey, emo, ed, eh, emin = m.groups()
        if sh == '00' and smin == '00' and eh == '00' and emin == '00':
            start = datetime(int(sy), int(smo), int(sd))
            end = datetime(int(ey), int(emo), int(ed)) - timedelta(days=1)
            sdow = DAYS_KR[start.weekday()]
            if start.date() == end.date():
                return f'{sy}.{smo}.{sd}({sdow})'
            edow = DAYS_KR[end.weekday()]
            return (
                f'{sy}.{smo}.{sd}({sdow}) ~ '
                f'{end.year}.{end.month:02d}.{end.day:02d}({edow})'
            )
    return value


# ════════════════════════════════════════════════════════════════════════════
# 메인 빌더
# ════════════════════════════════════════════════════════════════════════════

def build_hwpx_bytes(csv_text: str, report_title: str = '') -> bytes:
    """CSV 텍스트 → HWPX 바이트 (순수 Python, COM 불필요)."""
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)
    if not rows:
        raise RuntimeError('CSV 내용이 비어 있습니다.')

    rows[0] = [
        col[2:] if len(col) > 2 and col[1] == '_' and col[0].isalpha() else col
        for col in rows[0]
    ]
    header = rows[0]
    data_rows = [list(r) for r in rows[1:]]

    gi = header.index('구분') if '구분' in header else 0
    ii = header.index('일정') if '일정' in header else 1
    ci = header.index('내용') if '내용' in header else 2
    bi = header.index('비고') if '비고' in header else len(header) - 1

    data_rows.sort(key=lambda r: (
        r[gi] if gi < len(r) else '',
        r[ii] if ii < len(r) else '',
    ))

    all_dates = [_parse_start_date(r[ii]) for r in data_rows if ii < len(r)]
    all_dates = [d for d in all_dates if d]
    if all_dates:
        mn = min(all_dates)
        week_monday = mn - timedelta(days=mn.weekday())
    else:
        today = date.today()
        week_monday = today - timedelta(days=today.weekday())

    day_dates = [week_monday + timedelta(days=i) for i in range(5)]
    day_events: Dict[date, List] = {d: [] for d in day_dates}

    for row in data_rows:
        raw = row[ii] if ii < len(row) else ''
        ev_date = _parse_start_date(raw)
        if ev_date in day_events:
            day_events[ev_date].append((
                row[gi] if gi < len(row) else '',
                row[ci] if ci < len(row) else '',
                _parse_time_range(raw),
            ))

    for row in data_rows:
        if ii < len(row):
            row[ii] = _fix_allday(row[ii])

    title = f'콘텐츠IP전략팀 {report_title}'.strip() if report_title else '콘텐츠IP전략팀'

    header_xml  = _make_header_xml()
    section_xml = _make_section_xml(
        title, day_dates, day_events, header, data_rows, gi, ii, ci, bi,
    )
    return _pack_hwpx(header_xml, section_xml, title)


# ════════════════════════════════════════════════════════════════════════════
# Notion API
# ════════════════════════════════════════════════════════════════════════════

def _notion_headers(token: str, json_ct: bool = False) -> Dict[str, str]:
    h = {'Authorization': f'Bearer {token}', 'Notion-Version': NOTION_VERSION}
    if json_ct:
        h['Content-Type'] = 'application/json'
    return h


def upload_file_to_notion(token: str, file_bytes: bytes, filename: str) -> str:
    # Notion은 .hwpx/.hwp 확장자를 지원하지 않음 → .zip으로 업로드
    upload_name = filename
    if upload_name.lower().endswith('.hwpx'):
        upload_name = upload_name[:-5] + '.zip'
    elif upload_name.lower().endswith('.hwp'):
        upload_name = upload_name[:-4] + '.zip'

    resp = requests.post(
        'https://api.notion.com/v1/file_uploads',
        headers=_notion_headers(token, json_ct=True),
        json={'content_type': 'application/zip', 'filename': upload_name},
        timeout=30,
    )
    resp.raise_for_status()
    obj = resp.json()

    send = requests.post(
        obj['upload_url'],
        headers=_notion_headers(token),
        files={'file': (upload_name, file_bytes, 'application/zip')},
        timeout=60,
    )
    send.raise_for_status()
    return obj['id'], upload_name


def _resolve_prop_name(token: str, page_id: str, wanted: str) -> str:
    """페이지 속성 중 wanted 이름이 없으면 'HWP'를 포함한 속성, 없으면 '파일과 미디어' 반환."""
    r = requests.get(
        f'https://api.notion.com/v1/pages/{page_id}',
        headers=_notion_headers(token),
        timeout=15,
    )
    if r.status_code != 200:
        return wanted
    props = r.json().get('properties', {})
    if wanted in props:
        return wanted
    # 'HWP'가 포함된 속성 탐색 (깨진 이름 포함)
    for k in props:
        if 'HWP' in k or 'hwp' in k.lower():
            print(f'[INFO] 속성명 자동 선택: {repr(k)} (wanted={repr(wanted)})')
            return k
    # 최종 fallback
    fallback = '파일과 미디어'
    print(f'[INFO] HWP 속성 없음 → {fallback} 사용')
    return fallback


def attach_file_to_page(token: str, page_id: str, file_upload_id: str,
                        filename: str, file_property_name: str) -> None:
    resolved = _resolve_prop_name(token, page_id, file_property_name)
    body = {
        'properties': {
            resolved: {
                'files': [{
                    'type': 'file_upload',
                    'file_upload': {'id': file_upload_id},
                    'name': filename,
                }]
            }
        }
    }
    r = requests.patch(
        f'https://api.notion.com/v1/pages/{page_id}',
        headers=_notion_headers(token, json_ct=True),
        json=body,
        timeout=30,
    )
    r.raise_for_status()


# ════════════════════════════════════════════════════════════════════════════
# 진입점
# ════════════════════════════════════════════════════════════════════════════

def _load_payload(event_path: str) -> Dict[str, Any]:
    with open(event_path, 'r', encoding='utf-8') as f:
        event = json.load(f)
    payload = event.get('client_payload') or event.get('inputs') or {}
    if not payload:
        raise RuntimeError('payload 없음 (client_payload or inputs)')
    return payload


def main() -> int:
    notion_token = os.getenv('NOTION_TOKEN', '').strip()
    if not notion_token:
        raise RuntimeError('NOTION_TOKEN 환경변수 필요')

    event_path = os.getenv('GITHUB_EVENT_PATH', '')
    if not event_path:
        raise RuntimeError('GITHUB_EVENT_PATH 환경변수 없음')

    payload = _load_payload(event_path)

    csv_text   = payload.get('csv_text', '')
    csv_b64    = payload.get('csv_b64', '')
    page_id    = payload.get('notion_page_id', '')
    filename   = payload.get('report_name', 'weekly_report.hwpx')
    prop_name  = payload.get('file_property_name', 'HWP 파일')
    delimiter  = payload.get('delimiter', ',')
    rpt_title  = payload.get('report_title', '')

    if not (csv_text or csv_b64):
        raise RuntimeError('payload에 csv_text 또는 csv_b64 필요')
    if not page_id:
        raise RuntimeError('payload에 notion_page_id 필요')

    if csv_b64:
        csv_text = base64.b64decode(csv_b64).decode('utf-8', errors='replace')
    else:
        csv_text = unescape_csv_text(csv_text)

    if delimiter == 'tab':
        delimiter = '\t'

    if filename.lower().endswith('.xlsx'):
        filename = filename[:-5] + '.hwpx'
    elif not filename.lower().endswith('.hwpx'):
        filename += '.hwpx'

    # "04월 20일 ~ 04월 26일 주간업무 보고" → "2026년 17주차" 형태로 변환
    if rpt_title and '주간업무 보고' in rpt_title:
        m = re.search(r'(\d{1,2})월\s*(\d{1,2})일', rpt_title)
        if m:
            from datetime import date as _date
            month, day = int(m.group(1)), int(m.group(2))
            year = _date.today().year
            try:
                d = _date(year, month, day)
                week_num = d.isocalendar()[1]
                rpt_title = f'{year}년 {week_num}주차'
            except Exception:
                pass

    hwpx_bytes               = build_hwpx_bytes(csv_text, report_title=rpt_title)
    file_upload_id, uploaded = upload_file_to_notion(notion_token, hwpx_bytes, filename)
    attach_file_to_page(notion_token, page_id, file_upload_id, uploaded, prop_name)

    print('OK')
    print(json.dumps(
        {'page_id': page_id, 'filename': uploaded, 'file_upload_id': file_upload_id},
        ensure_ascii=False,
    ))
    return 0


if __name__ == '__main__':
    try:
        raise SystemExit(main())
    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr)
        raise
