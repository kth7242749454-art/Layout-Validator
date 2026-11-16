# -*- coding: utf-8 -*-
"""
lee.py
=====================================================================
INDEX(Excel)와 'SECONDARY CABLE ROUTING LAYOUT' 도면(PDF)의
INSTRUMENT LIST(TAG NUMBER / J/B NO.)를 자동 비교/검증하는 프로그램

산출물
  1) TAG 확인 결과.xlsx
       - sheet "Old_Instrument_Name"   : 삭제가 필요한(구) TAG 목록
       - sheet "검토가 필요한 태그"      : INDEX에 존재하지 않는 TAG 목록
  2) JB NO. 검토 필요.xlsx
       - LAYOUT TAG와 INDEX상 J/B NO.가 불일치하는 항목
  3) 마킹된_PDF 폴더
       - 삭제 필요 TAG   -> 빨간색 하이라이트
       - 검토 필요 TAG   -> 파란색 하이라이트

-----------------------------------------------------------------
PyInstaller 빌드 예시 (참고용)
    pyinstaller --onefile --windowed lee.py ^
        --collect-all customtkinter ^
        --collect-all fitz ^
        --collect-all tkinterdnd2 ^
        --hidden-import PIL._tkinter_finder
-----------------------------------------------------------------
필요 패키지: customtkinter, pillow, pymupdf(fitz), openpyxl, tkinterdnd2(드래그앤드롭, 선택)
    pip install customtkinter pillow pymupdf openpyxl tkinterdnd2
=====================================================================
"""

import os
import re
import sys
import shutil
import zipfile
import tempfile
import traceback
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog, messagebox

# customtkinter가 Windows 배율(125%/150% 등)을 자동 감지해서 geometry()에
# 곱해버리면, 우리가 직접 px 단위로 계산한 캔버스/카드 크기와 실제 창 크기가
# 어긋나 오른쪽/아래에 빈 공간이 남는다. 배율을 1:1로 고정해서 이 문제를 없앤다.
ctk.deactivate_automatic_dpi_awareness()
ctk.set_widget_scaling(1.0)
ctk.set_window_scaling(1.0)

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

from PIL import Image, ImageDraw, ImageFilter

import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import ImageFont


# =====================================================================
#  색상 팔레트 (card.jpg/icon.jpg 참고 - 파스텔 연보라 플랫 배경 + 흰 카드,
#  화려한 그라디언트/블러 없이 얇은 보더 위주의 미니멀한 대시보드 톤)
# =====================================================================
PALETTES = {
    "light": {
        "BG": "#E9DDFE",
        "CARD": "#FFFFFF", "CARD_ALT": "#FFFFFF", "CARD_ALT_HOVER": "#F1ECFE",
        "ACCENT_CHIP": "#F1ECFE",
        "BORDER": "#EFEAFB",
        "ACCENT": "#8A7CF6", "ACCENT_DARK": "#5B4FD6", "ACCENT_PINK": "#6F62E0",
        "TEXT": "#2A2358", "MUTED": "#948BB8",
        "SUCCESS_BG": "#E8F6EC", "SUCCESS_TEXT": "#379157",
        "WARN_BG": "#FFF6E0", "WARN_TEXT": "#B08000",
        "DANGER_BG": "#FDEAF0", "DANGER_TEXT": "#D45C82",
        "INFO_BG": "#EAF0FD", "INFO_TEXT": "#4C6BC2",
        "LOG_BG": "#F8F5FD",
        "TOGGLE_ICON": "🌙", "TOGGLE_LABEL": "다크모드",
    },
    "dark": {
        "BG": "#221D3D",
        "CARD": "#322C55", "CARD_ALT": "#322C55", "CARD_ALT_HOVER": "#3B3462",
        "ACCENT_CHIP": "#3E3568",
        "BORDER": "#3E3766",
        "ACCENT": "#9C8FF8", "ACCENT_DARK": "#7B6DE8", "ACCENT_PINK": "#8879EC",
        "TEXT": "#EEEAFB", "MUTED": "#B4ABD6",
        "SUCCESS_BG": "#25402E", "SUCCESS_TEXT": "#8FDBA5",
        "WARN_BG": "#453A22", "WARN_TEXT": "#E8C878",
        "DANGER_BG": "#452530", "DANGER_TEXT": "#F09CB0",
        "INFO_BG": "#25304A", "INFO_TEXT": "#9CB6F0",
        "LOG_BG": "#282243",
        "TOGGLE_ICON": "☀", "TOGGLE_LABEL": "라이트모드",
    },
}

THEME_MODE = "light"


def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def load_font(size, bold=False):
    candidates = (
        ["malgunbd.ttf", "Malgun Gothic Bold.ttf"] if bold else ["malgun.ttf", "Malgun Gothic.ttf"]
    ) + ["NanumGothicBold.ttf" if bold else "NanumGothic.ttf", "arial.ttf"]
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except Exception:
            continue
    return ImageFont.load_default()


def draw_centered_text(draw, box, text, font, color):
    x0, y0, x1, y1 = box
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    cx = x0 + ((x1 - x0) - tw) / 2 - bbox[0]
    cy = y0 + ((y1 - y0) - th) / 2 - bbox[1]
    draw.text((cx, cy), text, font=font, fill=color)


def make_button_image(width, height, c1_hex, c2_hex, text, radius=14, text_size=16, muted=False):
    """단색 배경 + 텍스트를 슈퍼샘플링으로 안티에일리어싱해서 하나의 이미지로 구운다.
    (customtkinter 위젯 중첩 시 발생하는 반투명 미지원 아티팩트 방지 + 각지고 깨져 보이는 것 방지)"""
    ss = 4
    bw, bh = width * ss, height * ss
    color = hex_to_rgb(c1_hex) + (255,)

    img = Image.new("RGBA", (bw, bh), color)
    mask = Image.new("L", (bw, bh), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, bw - 1, bh - 1], radius=radius * ss, fill=255)
    rounded = Image.new("RGBA", (bw, bh), (0, 0, 0, 0))
    rounded.paste(img, (0, 0), mask)

    draw = ImageDraw.Draw(rounded)
    font = load_font(text_size * ss, bold=True)
    text_color = (255, 255, 255, 160) if muted else (255, 255, 255, 255)
    draw_centered_text(draw, (0, 0, bw, bh), text, font, text_color)

    return rounded.resize((width, height), Image.LANCZOS)


def _icon_excel(draw, cx, cy, r, c, filled=True):
    s = r * 0.52
    if filled:
        draw.rounded_rectangle([cx - s, cy - s, cx + s, cy + s], radius=r * 0.16, fill=c)
        bg = (0, 0, 0, 0)
        bw = s * 0.16
        draw.rectangle([cx - s + bw, cy - s / 3, cx + s - bw, cy - s / 3 + bw], fill=bg)
        draw.rectangle([cx - s + bw, cy + s / 3 - bw, cx + s - bw, cy + s / 3], fill=bg)
        draw.rectangle([cx - bw / 2, cy - s + bw, cx + bw / 2, cy - s / 3], fill=bg)
        draw.rectangle([cx - bw / 2, cy - s / 3 + bw, cx + bw / 2, cy + s / 3 - bw], fill=bg)
        draw.rectangle([cx - bw / 2, cy + s / 3, cx + bw / 2, cy + s - bw], fill=bg)
    else:
        # icon.jpg 스타일: 채우지 않고 얇은 선(스트로크)만으로 그리는 격자/표 아이콘
        sw = max(int(r * 0.18), 3)
        draw.rounded_rectangle([cx - s, cy - s, cx + s, cy + s], radius=r * 0.16, outline=c, width=sw)
        draw.line([(cx - s + sw, cy - s / 3), (cx + s - sw, cy - s / 3)], fill=c, width=sw)
        draw.line([(cx - s + sw, cy + s / 3), (cx + s - sw, cy + s / 3)], fill=c, width=sw)
        draw.line([(cx, cy - s + sw), (cx, cy + s - sw)], fill=c, width=sw)


def _icon_zip(draw, cx, cy, r, c, filled=True):
    s = r * 0.5
    if filled:
        draw.rounded_rectangle([cx - s, cy - s * 0.55, cx + s, cy + s], radius=r * 0.14, fill=c)
        draw.rounded_rectangle([cx - s * 0.55, cy - s * 0.85, cx + s * 0.55, cy - s * 0.45],
                                radius=r * 0.08, fill=c)
        bg = (0, 0, 0, 0)
        tw = s * 0.16
        for i in range(3):
            yy = cy - s * 0.28 + i * s * 0.38
            draw.rectangle([cx - tw / 2, yy, cx + tw / 2, yy + s * 0.18], fill=bg)
    else:
        sw = max(int(r * 0.18), 3)
        draw.rounded_rectangle([cx - s, cy - s * 0.55, cx + s, cy + s], radius=r * 0.14, outline=c, width=sw)
        draw.rounded_rectangle([cx - s * 0.55, cy - s * 0.85, cx + s * 0.55, cy - s * 0.45],
                                radius=r * 0.08, outline=c, width=sw)
        tw = s * 0.18
        for i in range(3):
            yy = cy - s * 0.22 + i * s * 0.3
            draw.line([(cx - tw / 2, yy), (cx + tw / 2, yy)], fill=c, width=max(int(sw * 0.8), 2))


def _icon_pdf(draw, cx, cy, r, c, filled=True):
    s = r * 0.56
    fold = s * 0.42
    x0, y0, x1, y1 = cx - s * 0.68, cy - s, cx + s * 0.68, cy + s
    if filled:
        draw.polygon([(x0, y0), (x1 - fold, y0), (x1, y0 + fold), (x1, y1), (x0, y1)], fill=c)
        draw.polygon([(x1 - fold, y0), (x1, y0 + fold), (x1 - fold, y0 + fold)],
                      fill=tuple(int(ch * 0.82) for ch in c[:3]) + (c[3] if len(c) > 3 else 255,))
        bg = (0, 0, 0, 0)
        bw = s * 0.14
        for i in range(3):
            yy = cy - s * 0.05 + i * s * 0.34
            draw.rectangle([x0 + s * 0.22, yy, x1 - s * 0.22, yy + bw], fill=bg)
    else:
        sw = max(int(r * 0.18), 3)
        draw.line([(x0, y0), (x1 - fold, y0)], fill=c, width=sw)
        draw.line([(x1 - fold, y0), (x1, y0 + fold)], fill=c, width=sw)
        draw.line([(x1, y0 + fold), (x1, y1)], fill=c, width=sw)
        draw.line([(x1, y1), (x0, y1)], fill=c, width=sw)
        draw.line([(x0, y1), (x0, y0)], fill=c, width=sw)
        draw.line([(x1 - fold, y0), (x1 - fold, y0 + fold)], fill=c, width=max(int(sw * 0.7), 2))
        draw.line([(x1 - fold, y0 + fold), (x1, y0 + fold)], fill=c, width=max(int(sw * 0.7), 2))
        for i in range(3):
            yy = cy - s * 0.02 + i * s * 0.32
            draw.line([(x0 + s * 0.22, yy), (x1 - s * 0.22, yy)], fill=c, width=max(int(sw * 0.7), 2))


def _icon_folder(draw, cx, cy, r, c, filled=True):
    s = r * 0.56
    if filled:
        draw.rounded_rectangle(
            [cx - s, cy - s * 0.15, cx + s, cy + s * 0.62], radius=r * 0.14, fill=c)
        draw.polygon([
            (cx - s, cy - s * 0.15), (cx - s, cy - s * 0.45), (cx - s * 0.35, cy - s * 0.45),
            (cx - s * 0.12, cy - s * 0.15),
        ], fill=c)
    else:
        sw = max(int(r * 0.18), 3)
        draw.rounded_rectangle(
            [cx - s, cy - s * 0.15, cx + s, cy + s * 0.62], radius=r * 0.14, outline=c, width=sw)
        draw.line([(cx - s, cy - s * 0.15), (cx - s, cy - s * 0.45)], fill=c, width=sw)
        draw.line([(cx - s, cy - s * 0.45), (cx - s * 0.35, cy - s * 0.45)], fill=c, width=sw)
        draw.line([(cx - s * 0.35, cy - s * 0.45), (cx - s * 0.12, cy - s * 0.15)], fill=c, width=sw)


def _icon_tag(draw, cx, cy, r, c, filled=True):
    s = r * 0.55
    pts = [(cx - s, cy), (cx - s * 0.12, cy - s * 0.82), (cx + s, cy - s * 0.82),
           (cx + s, cy + s * 0.82), (cx - s * 0.12, cy + s * 0.82)]
    if filled:
        draw.polygon(pts, fill=c)
        hr = s * 0.17
        draw.ellipse([cx - s * 0.62 - hr, cy - hr, cx - s * 0.62 + hr, cy + hr], fill=(0, 0, 0, 0))
    else:
        sw = max(int(r * 0.18), 3)
        draw.polygon(pts, outline=c, width=sw)
        hr = s * 0.14
        draw.ellipse([cx - s * 0.62 - hr, cy - hr, cx - s * 0.62 + hr, cy + hr], outline=c, width=max(sw - 1, 2))


def _icon_check(draw, cx, cy, r, c, filled=True):
    s = r * 0.5
    draw.line([(cx - s, cy), (cx - s * 0.15, cy + s * 0.72)], fill=c, width=int(r * 0.22), joint="curve")
    draw.line([(cx - s * 0.15, cy + s * 0.72), (cx + s, cy - s * 0.6)], fill=c, width=int(r * 0.22), joint="curve")


_ICON_FUNCS = {
    "excel": _icon_excel, "zip": _icon_zip, "pdf": _icon_pdf,
    "folder": _icon_folder, "tag": _icon_tag, "check": _icon_check,
}

_ICON_CACHE = {}


def make_badge_image(size, bg_hex, icon_type, icon_color_hex=None):
    """스퀴클(둥근 사각) 배지 + 슈퍼샘플링으로 안티에일리어싱한 채워진 실루엣 아이콘.
    icon_color_hex를 안 주면 bg_hex를 진하게 칠하고 흰색 아이콘을 올리는 '브랜드용'
    배지(로고 등), icon_color_hex를 주면 bg_hex를 옅은 칩 배경으로 쓰고 그 색으로
    아이콘을 그리는 '보조용' 배지(배경.png의 옅은 칩 + 컬러 아이콘 스타일)가 된다."""
    cache_key = (size, bg_hex, icon_type, icon_color_hex)
    if cache_key in _ICON_CACHE:
        return _ICON_CACHE[cache_key]

    ss = 4  # 슈퍼샘플 배율
    big = size * ss
    c1 = hex_to_rgb(bg_hex)
    icon_color = hex_to_rgb(icon_color_hex) + (255,) if icon_color_hex else (255, 255, 255, 255)

    radius = big * 0.38
    img = Image.new("RGBA", (big, big), c1 + (255,))
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).rounded_rectangle([0, 0, big - 1, big - 1], radius=radius, fill=255)
    rounded = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    rounded.paste(img, (0, 0), mask)

    draw = ImageDraw.Draw(rounded)
    fn = _ICON_FUNCS.get(icon_type, _icon_check)
    fn(draw, big / 2, big / 2, big * 0.36, icon_color, filled=(icon_color_hex is None))

    out = rounded.resize((size, size), Image.LANCZOS)
    _ICON_CACHE[cache_key] = out
    return out


def make_number_badge(size, bg_hex, bg_hex2, number_text):
    """숫자가 텍스트 오버레이 없이 이미지 안에 직접 구워진, 부드러운 단색 원형 배지.
    (겹친 위젯 안티에일리어싱 문제로 깨져 보이던 숫자 배지를 완전히 대체. 그라디언트 없음)"""
    cache_key = ("num", size, bg_hex, number_text)
    if cache_key in _ICON_CACHE:
        return _ICON_CACHE[cache_key]

    ss = 4
    big = size * ss
    c1 = hex_to_rgb(bg_hex)

    img = Image.new("RGBA", (big, big), c1 + (255,))
    mask = Image.new("L", (big, big), 0)
    ImageDraw.Draw(mask).ellipse([0, 0, big - 1, big - 1], fill=255)
    rounded = Image.new("RGBA", (big, big), (0, 0, 0, 0))
    rounded.paste(img, (0, 0), mask)

    draw = ImageDraw.Draw(rounded)
    font = load_font(int(big * 0.46), bold=True)
    draw_centered_text(draw, (0, 0, big, big), number_text, font, (255, 255, 255, 255))

    out = rounded.resize((size, size), Image.LANCZOS)
    _ICON_CACHE[cache_key] = out
    return out


def to_ctk_image(pil_img, size=None):
    size = size or pil_img.size
    return ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=size)


# =====================================================================
#  PDF INSTRUMENT LIST 추출
# =====================================================================
CYRILLIC_RE = re.compile(r'[\u0400-\u04FF]')
TAG_FORMAT_RE = re.compile(r'^\d{3,4}(-[A-Za-z0-9]+){1,4}$')


def _norm_header(s):
    return re.sub(r'[\s./]', '', s).upper()


JB_HEADER_NAMES = ("J/BNO", "JBNO", "J/BNO.", "JBNO.", "REMARK")


def _is_jb_header(s):
    """'J/B NO.' 뿐 아니라 'REMARK' 헤더도 J/B NO.와 같은 의미로 취급한다
    (도면에 따라 J/B NO. 열 제목이 REMARK로 되어 있는 경우가 있음).
    단, APC NO. 표의 REMARK 열은 진짜 비고란이라 별도 경로(_extract_apc_tag_only)에서
    처리하므로 여기서 제외해야 한다 - 호출부에서 APC NO. 헤더 존재 여부로 따로 걸러낸다."""
    return _norm_header(s) in JB_HEADER_NAMES


def _looks_like_tag(text):
    """실제 계기 태그 형식(예: 3021-XV-850LM)인지 검증. 조합 참조문자열,
    치수/서포트 표기 등 잡음을 걸러낸다."""
    return bool(TAG_FORMAT_RE.match(text))


def _page_spans(page):
    """페이지의 모든 텍스트 스팬을 (가로/세로 방향 포함) 수집한다."""
    d = page.get_text("dict")
    spans = []
    for block in d.get("blocks", []):
        if "lines" not in block:
            continue
        for line in block["lines"]:
            ldir = line.get("dir", (1, 0))
            is_vertical = abs(ldir[0]) < 0.5 and abs(ldir[1]) > 0.5
            for span in line["spans"]:
                txt = span["text"].strip()
                if not txt or CYRILLIC_RE.search(txt):
                    continue
                x0, y0, x1, y1 = span["bbox"]
                spans.append({"x0": x0, "y0": y0, "x1": x1, "y1": y1,
                              "text": txt, "vert": is_vertical})
    return spans


def _dynamic_left_bound(header_x, header_y, other_headers, max_window, y_tol=60, margin=8):
    """고정된 거리(예: 1400pt)를 억지로 정하는 대신, 실제로 가장 가까운
    '다른 표의 헤더'가 있으면 거기서 멈추고, 없으면 max_window까지 넓게 본다.
    이러면 표가 넓을 땐 알아서 더 멀리 보고, 표끼리 다닥다닥 붙어있을 땐
    옆 표를 침범하지 않는다."""
    x_lo = header_x - max_window
    neighbors = [h["x0"] for h in other_headers
                 if abs(h["y0"] - header_y) < y_tol and h["x0"] < header_x - margin]
    if neighbors:
        nearest = max(neighbors)
        x_lo = max(x_lo, nearest + margin)
    return x_lo


def _extract_horizontal_jb_tables(spans):
    """가로 방향 INSTRUMENT LIST (TAG NUMBER + ELEVATION + J/B NO. 또는 REMARK) 표 추출.

    핵심 원칙:
      1) 같은 y에 있다는 이유만으로 멀리 떨어진 TAG NUMBER와 J/B NO.를 묶지 않는다.
         반드시 '현재 TAG NUMBER ~ 같은 행의 다음 TAG NUMBER' 사이에 있는 J/B NO.만
         같은 표의 헤더로 인정한다.
      2) 실제 J/B 비교 표는 NO < TAG NUMBER < ELEVATION < J/B NO. 구조여야 한다.
         TAG와 J/B 사이에 DWG. NO. / APC NO.가 끼어 있으면 범례/APC 표이므로 제외한다.
      3) 표의 세로 끝(y_end)은 '같은 컬럼의 다음 TAG NUMBER 헤더'로만 잡는다.
         예전처럼 x 차이 350pt까지 허용하면 옆 표의 헤더가 경계로 잡혀 긴 표가
         첫 몇 행에서 잘리는 문제가 발생한다.
    """
    results = []

    tag_headers = [
        s for s in spans
        if not s["vert"] and _norm_header(s["text"]) == "TAGNUMBER"
    ]
    apc_headers_all = [
        s for s in spans
        if not s["vert"] and _norm_header(s["text"]) in ("APCNO", "APCNO.")
    ]
    jb_headers = [
        s for s in spans
        if not s["vert"] and _is_jb_header(s["text"])
        # APC 표의 REMARK를 J/B NO.로 오인하지 않도록 제외
        and not any(abs(a["y0"] - s["y0"]) < 3 for a in apc_headers_all)
    ]
    no_headers = [
        s for s in spans
        if not s["vert"] and _norm_header(s["text"]) == "NO"
    ]
    elev_headers = [
        s for s in spans
        if not s["vert"] and _norm_header(s["text"]) == "ELEVATION"
    ]
    dwg_headers = [
        s for s in spans
        if not s["vert"] and _norm_header(s["text"]) in ("DWGNO", "DWGNO.")
    ]

    header_groups = []

    for th in tag_headers:
        # 같은 행에 나란히 배치된 '다음 표'의 TAG NUMBER가 있으면 그 x를 현재 표의
        # 오른쪽 경계로 사용한다. 이 경계를 넘어 있는 J/B NO.는 절대 현재 표와 묶지 않는다.
        same_row_next_tags = [
            t for t in tag_headers
            if t is not th
            and abs(t["y0"] - th["y0"]) < 3
            and t["x0"] > th["x0"]
        ]
        next_tag_x = (
            min(t["x0"] for t in same_row_next_tags)
            if same_row_next_tags else float("inf")
        )

        jb_right = [
            j for j in jb_headers
            if abs(j["y0"] - th["y0"]) < 3
            and th["x0"] < j["x0"] < next_tag_x
        ]
        if not jb_right:
            continue
        jbh = min(jb_right, key=lambda j: j["x0"] - th["x0"])

        # TAG NUMBER + DWG. NO. 범례표 또는 TAG NUMBER + APC NO. 표를
        # J/B 비교 표로 잘못 분류하지 않는다.
        if any(
            abs(d["y0"] - th["y0"]) < 3
            and th["x0"] < d["x0"] < jbh["x0"]
            for d in dwg_headers
        ):
            continue
        if any(
            abs(a["y0"] - th["y0"]) < 3
            and th["x0"] < a["x0"] < jbh["x0"]
            for a in apc_headers_all
        ):
            continue

        # 이 프로그램이 J/B 비교 대상으로 삼는 실제 INSTRUMENT LIST는
        # TAG NUMBER -> ELEVATION -> J/B NO. 구조다. ELEVATION이 같은 로컬 표
        # 안에 없으면 안전하게 J/B 비교에서 제외한다.
        elev_right = [
            e for e in elev_headers
            if abs(e["y0"] - th["y0"]) < 3
            and th["x0"] < e["x0"] < jbh["x0"]
        ]
        if not elev_right:
            continue
        eh = min(elev_right, key=lambda e: e["x0"] - th["x0"])

        # NO. 헤더도 현재 로컬 표 범위 안에서만 선택한다.
        same_row_prev_tags = [
            t["x0"] for t in tag_headers
            if abs(t["y0"] - th["y0"]) < 3 and t["x0"] < th["x0"]
        ]
        left_boundary = max(same_row_prev_tags) if same_row_prev_tags else -float("inf")
        match_no = [
            n for n in no_headers
            if abs(n["y0"] - th["y0"]) < 3
            and left_boundary < n["x0"] < th["x0"]
        ]
        no_x0 = (
            max(match_no, key=lambda n: n["x0"])["x0"]
            if match_no else th["x0"] - 50
        )

        header_groups.append({
            "y": th["y0"],
            "no_x0": no_x0,
            "tag_x0": th["x0"],
            "elev_x0": eh["x0"],
            "jb_x0": jbh["x0"],
            "right_bound": next_tag_x,
        })

    header_groups.sort(key=lambda h: h["y"])

    for hg in header_groups:
        y_start = hg["y"] + 5

        no_tag_gap = hg["tag_x0"] - hg["no_x0"]
        tag_elev_gap = hg["elev_x0"] - hg["tag_x0"]
        elev_jb_gap = hg["jb_x0"] - hg["elev_x0"]

        # '같은 컬럼의 다음 표'만 y_end 후보로 인정한다.
        # 로컬 컬럼 폭을 기준으로 tolerance를 정해 옆 표(수백 pt 떨어짐)의 헤더가
        # 긴 INSTRUMENT LIST를 중간에서 잘라버리지 않게 한다.
        same_col_tol = max(20, min(100, abs(tag_elev_gap) * 0.75))
        nearby_ys = sorted(set(
            round(th["y0"]) for th in tag_headers
            if abs(th["x0"] - hg["tag_x0"]) < same_col_tol
        ))
        later_ys = [y for y in nearby_ys if y > hg["y"] + 10]
        y_end = min(later_ys) if later_ys else 10 ** 7

        tag_x_lo = hg["no_x0"] + no_tag_gap * 0.3
        tag_x_hi = hg["elev_x0"] - tag_elev_gap * 0.3
        jb_x_lo = hg["elev_x0"] + elev_jb_gap * 0.55
        jb_x_hi = hg["jb_x0"] + tag_elev_gap

        # 같은 행에 다음 표가 있으면 그 표의 TAG NUMBER보다 오른쪽까지 J/B 영역이
        # 뻗지 않도록 절대 경계를 한 번 더 건다.
        if hg["right_bound"] != float("inf"):
            jb_x_hi = min(jb_x_hi, hg["right_bound"] - 5)

        # 같은 y의 다음 J/B 그룹 NO. 열도 침범하지 않게 한다.
        same_row_next_x = [
            g["no_x0"] for g in header_groups
            if abs(g["y"] - hg["y"]) < 10 and g["no_x0"] > hg["jb_x0"]
        ]
        if same_row_next_x:
            jb_x_hi = min(jb_x_hi, min(same_row_next_x) - 5)

        tag_spans = [
            s for s in spans
            if not s["vert"]
            and y_start <= s["y0"] < y_end
            and tag_x_lo <= s["x0"] < tag_x_hi
        ]
        jb_spans = [
            s for s in spans
            if not s["vert"]
            and y_start <= s["y0"] < y_end
            and jb_x_lo <= s["x0"] < jb_x_hi
        ]

        for ts in tag_spans:
            tt = ts["text"].strip()
            if not _looks_like_tag(tt) or JB_PATTERN_RE.match(tt):
                continue

            jb_match = [
                j for j in jb_spans
                if abs(j["y0"] - ts["y0"]) < 2.5
                and _looks_like_jb_value(j["text"])
            ]
            jb_text = jb_match[0]["text"].strip() if jb_match else ""

            results.append({
                "tag": tt,
                "jb": jb_text,
                "page": None,
                "bbox": (ts["x0"], ts["y0"], ts["x1"], ts["y1"]),
                "needs_jb_check": True,
            })

    return results


def _extract_rotated_jb_tables(spans):
    """세로(90도 회전) 방향 INSTRUMENT LIST (TAG NUMBER + ELEVATION + J/B NO.) 표 추출.
    도면이 회전된 상태로 배치된 경우 텍스트도 함께 회전되어 있는데, 이 경우
    가로 방향 로직으로는 단 하나도 인식하지 못하므로 별도 처리가 필요하다."""
    results = []
    tag_headers = [s for s in spans if s["vert"] and _norm_header(s["text"]) == "TAGNUMBER"]
    apc_headers_all = [s for s in spans if s["vert"] and _norm_header(s["text"]) in ("APCNO", "APCNO.")]
    jb_headers = [s for s in spans if s["vert"] and _is_jb_header(s["text"])
                  and not any(abs(a["y0"] - s["y0"]) < 200 for a in apc_headers_all)]

    for th in tag_headers:
        match_jb = [j for j in jb_headers if 0 < (j["y0"] - th["y0"]) < 200]
        if not match_jb:
            continue
        jbh = min(match_jb, key=lambda j: j["y0"] - th["y0"])
        header_x = th["x0"]
        other_headers = [h for h in tag_headers if h is not th]
        x_lo = _dynamic_left_bound(header_x, th["y0"], other_headers, max_window=3000)
        x_hi = header_x - 5

        tag_spans = [s for s in spans if s["vert"] and x_lo <= s["x0"] < x_hi
                     and abs(s["y0"] - th["y0"]) < 40 and s["text"] != "TAG NUMBER"]
        jb_spans = [s for s in spans if s["vert"] and x_lo <= s["x0"] < x_hi
                    and abs(s["y0"] - jbh["y0"]) < 40 and s["text"] != "J/B NO."]

        for ts in tag_spans:
            tt = ts["text"].strip()
            if not _looks_like_tag(tt) or JB_PATTERN_RE.match(tt):
                continue
            jb_match = [j for j in jb_spans if abs(j["x0"] - ts["x0"]) < 2.5
                        and _looks_like_jb_value(j["text"])]
            jb_text = jb_match[0]["text"].strip() if jb_match else ""
            results.append({
                "tag": ts["text"].strip(), "jb": jb_text, "page": None,
                "bbox": (ts["x0"], ts["y0"], ts["x1"], ts["y1"]), "needs_jb_check": True,
            })
    return results


def _extract_apc_tag_only(spans, vert):
    """TAG NUMBER + APC NO. (J/B NO. 없음) 형식의 표에서 TAG NUMBER만 추출.
    APC NO.는 INDEX와 비교하지 않고, OLD/현재 태그 분류만 수행한다.
    (가로/세로 방향 도면 둘 다 지원. 가로 방향은 표가 여러 행에 걸쳐 아래로
    쭉 나열되므로 헤더~다음 헤더 사이 전체를 범위로 잡는다.)"""
    results = []
    tag_headers = [s for s in spans if s["vert"] == vert and _norm_header(s["text"]) == "TAGNUMBER"]
    apc_headers = [s for s in spans if s["vert"] == vert and _norm_header(s["text"]) in ("APCNO", "APCNO.")]
    jb_headers = [s for s in spans if s["vert"] == vert and _norm_header(s["text"]) in
                  ("J/BNO", "JBNO", "J/BNO.", "JBNO.")]
    own_tol = 3 if not vert else 200
    for th in tag_headers:
        has_apc = any(abs(a["y0"] - th["y0"]) < own_tol for a in apc_headers)
        has_jb = any(abs(j["y0"] - th["y0"]) < own_tol for j in jb_headers)
        if not has_apc or has_jb:
            continue  # J/B NO 표는 이미 다른 함수에서 처리됨
        header_x = th["x0"]
        other_headers = [h for h in tag_headers if h is not th]
        x_lo = _dynamic_left_bound(header_x, th["y0"], other_headers, max_window=400)
        x_hi = header_x - 5
        if vert:
            tag_spans = [s for s in spans if s["vert"] and x_lo <= s["x0"] < x_hi
                         and abs(s["y0"] - th["y0"]) < 40 and s["text"] != "TAG NUMBER"]
        else:
            nearby_ys = sorted(set(round(t["y0"]) for t in tag_headers
                                    if abs(t["x0"] - th["x0"]) < 350))
            later_ys = [y for y in nearby_ys if y > th["y0"] + 10]
            y_end = min(later_ys) if later_ys else th["y0"] + 3000
            tag_spans = [s for s in spans if not s["vert"] and x_lo <= s["x0"] < x_hi
                         and th["y0"] + 5 <= s["y0"] < y_end and s["text"] != "TAG NUMBER"]
        for ts in tag_spans:
            tt = ts["text"].strip()
            if not _looks_like_tag(tt) or JB_PATTERN_RE.match(tt):
                continue
            results.append({
                "tag": ts["text"].strip(), "jb": None, "page": None,
                "bbox": (ts["x0"], ts["y0"], ts["x1"], ts["y1"]), "needs_jb_check": False,
            })
    return results


def _extract_legend_tags_only(spans, vert):
    """XXXX-DJB-XXX(NN POLE) 같은 범례 표(TAG NUMBER + DWG NO.)에서 태그명만 추출한다.
    표 제목에서 J/B NO.를 역산하는 기능은 검증 결과 신뢰도가 너무 낮아(다른 페이지의
    동일 표 대비 90% 안팎 불일치) 제외했다. 대신 '이 태그가 도면에 존재한다'는 사실만
    안전하게 뽑아서, 메인 INSTRUMENT LIST에 없는 태그도 OLD/현재 분류에서 빠지지
    않도록 한다 (JB 비교는 하지 않음). (가로/세로 방향 도면 둘 다 지원. 가로 방향은
    표가 여러 행에 걸쳐 아래로 쭉 나열되므로 헤더~다음 헤더 사이 전체를 범위로 잡는다.)"""
    results = []
    tag_headers = [s for s in spans if s["vert"] == vert and _norm_header(s["text"]) == "TAGNUMBER"]
    dwg_headers = [s for s in spans if s["vert"] == vert and _norm_header(s["text"]) in ("DWGNO", "DWGNO.")]
    jb_headers = [s for s in spans if s["vert"] == vert and _norm_header(s["text"]) in
                  ("J/BNO", "JBNO", "J/BNO.", "JBNO.")]
    own_tol = 3 if not vert else 200
    for th in tag_headers:
        has_dwg = any(abs(dh["y0"] - th["y0"]) < own_tol for dh in dwg_headers)
        has_jb = any(abs(j["y0"] - th["y0"]) < own_tol for j in jb_headers)
        if not has_dwg or has_jb:
            continue  # J/B NO 표는 이미 다른 함수에서 처리됨
        header_x = th["x0"]
        other_headers = [h for h in tag_headers if h is not th]
        x_lo = _dynamic_left_bound(header_x, th["y0"], other_headers, max_window=400)
        x_hi = header_x - 5
        if vert:
            tag_spans = [s for s in spans if s["vert"] and x_lo <= s["x0"] < x_hi
                         and abs(s["y0"] - th["y0"]) < 40 and s["text"] != "TAG NUMBER"]
        else:
            nearby_ys = sorted(set(round(t["y0"]) for t in tag_headers
                                    if abs(t["x0"] - th["x0"]) < 350))
            later_ys = [y for y in nearby_ys if y > th["y0"] + 10]
            y_end = min(later_ys) if later_ys else th["y0"] + 3000
            tag_spans = [s for s in spans if not s["vert"] and x_lo <= s["x0"] < x_hi
                         and th["y0"] + 5 <= s["y0"] < y_end and s["text"] != "TAG NUMBER"]
        for ts in tag_spans:
            tt = ts["text"].strip()
            if not _looks_like_tag(tt) or JB_PATTERN_RE.match(tt):
                continue
            results.append({
                "tag": ts["text"].strip(), "jb": None, "page": None,
                "bbox": (ts["x0"], ts["y0"], ts["x1"], ts["y1"]), "needs_jb_check": False,
            })
    return results


JB_CODE_ALT = "SSP|DJB|GJB|SJB|MJB|XLP|DLP|PJB|PLP|FLP|SLP|MSC"
JB_PATTERN_RE = re.compile(r'^\d{3,4}-(?:' + JB_CODE_ALT + r')-\d{2,4}', re.IGNORECASE)


def _looks_like_jb_value(text):
    """J/B NO. 칸에 들어갈 수 있는 값인지 확인한다: 빈칸/대시, 'Direct Run'(직결),
    또는 태그 형식(XXXX-코드-XXX)의 값 — DJB/SJB/PJB 같은 정식 JB 코드뿐 아니라
    다른 태그 번호를 J/B NO.로 그대로 쓰는 경우(예: 5110-WB-901)도 있으므로
    12개 코드로 제한하지 않고 태그 형식이면 폭넓게 인정한다.
    (근처 다른 표의 행 번호처럼 하이픈 없는 숫자만 있는 값은 태그 형식이 아니므로
    여전히 걸러진다.)"""
    t = text.strip()
    if t in ("", "-", "N/A", "NA"):
        return True
    if t.upper().replace(" ", "") == "DIRECTRUN":
        return True
    return bool(JB_PATTERN_RE.match(t)) or bool(TAG_FORMAT_RE.match(t))


def _extract_by_jb_pattern(spans):
    """헤더("TAG NUMBER"/"J/B NO.")가 아예 없는 표를 위한 보조 추출 경로.
    도면에 따라 INSTRUMENT LIST 헤더가 표 맨 앞에만 한 번 있고 이후 페이지/구간에는
    반복되지 않는 경우가 있는데, 이때는 헤더 기준 탐색이 통째로 실패한다.
    대신 'XXXX-DJB-XXX'/'XXXX-SJB-XXX'/'XXXX-PJB-XXX' 형식의 J/B NO. 값을 직접 찾고,
    같은 x좌표(같은 세로줄) 바로 위쪽에 있는 태그와 짝지어서 추출한다."""
    results = []
    jb_candidates = [s for s in spans if s["vert"] and JB_PATTERN_RE.match(s["text"].strip())]
    tag_candidates = [s for s in spans if s["vert"] and _looks_like_tag(s["text"].strip())
                       and not JB_PATTERN_RE.match(s["text"].strip())]

    for jb in jb_candidates:
        cands = [t for t in tag_candidates if abs(t["x0"] - jb["x0"]) < 2.5
                 and 0 < (jb["y0"] - t["y0"]) < 300]
        if not cands:
            continue
        tag_span = min(cands, key=lambda t: jb["y0"] - t["y0"])
        results.append({
            "tag": tag_span["text"].strip(), "jb": jb["text"].strip(), "page": None,
            "bbox": (tag_span["x0"], tag_span["y0"], tag_span["x1"], tag_span["y1"]),
            "needs_jb_check": True,
        })
    return results


def extract_instrument_list(pdf_path):
    """
    PDF 내 INSTRUMENT LIST 표에서 TAG NUMBER / J-B NO. 를 추출한다.
    가로 방향, 세로(90도 회전) 방향 표를 모두 지원하며, J/B NO. 대신
    APC NO.를 쓰는 표와 DJB/GJB 범례 표(태그 존재 여부만 검사, JB 비교 없음),
    그리고 헤더가 아예 없는 표(JB 패턴 직접 매칭)도 지원한다.
    반환: [{"tag":..., "jb":..., "page":..., "bbox":(x0,y0,x1,y1), "needs_jb_check": bool}, ...]
    """
    results = []
    doc = fitz.open(pdf_path)
    try:
        for pno in range(len(doc)):
            page = doc[pno]
            spans = _page_spans(page)

            page_results = []
            page_results += _extract_horizontal_jb_tables(spans)
            page_results += _extract_rotated_jb_tables(spans)
            page_results += _extract_apc_tag_only(spans, vert=True)
            page_results += _extract_apc_tag_only(spans, vert=False)
            page_results += _extract_legend_tags_only(spans, vert=True)
            page_results += _extract_legend_tags_only(spans, vert=False)
            page_results += _extract_by_jb_pattern(spans)

            # 여러 경로에서 같은 태그+JB 조합이 중복 추출될 수 있으므로 페이지 내에서 dedupe
            seen = set()
            deduped = []
            for r in page_results:
                key = (r["tag"], r["jb"], round(r["bbox"][0]), round(r["bbox"][1]))
                if key in seen:
                    continue
                seen.add(key)
                deduped.append(r)
            page_results = deduped

            for r in page_results:
                r["page"] = pno
            results += page_results
    finally:
        doc.close()
    return results


def diagnose_extraction_coverage(pdf_path, entries):
    """추출이 비정상적으로 적게 됐을 가능성이 있는지 자체 점검한다.
    (TAG NUMBER 헤더/INSTRUMENT LIST 표 개수 대비 실제 추출된 태그 수가
    너무 적으면, 이 도면의 표 배치를 아직 못 알아보고 있다는 신호일 수 있다.)
    반환: 경고 메시지 리스트 (문제 없으면 빈 리스트)"""
    warnings = []
    doc = fitz.open(pdf_path)
    try:
        header_count = 0
        section_count = 0
        for pno in range(len(doc)):
            page = doc[pno]
            text = page.get_text()
            header_count += text.upper().count("TAG NUMBER")
            section_count += text.upper().count("INSTRUMENT LIST")
    finally:
        doc.close()

    extracted = len(entries)
    if header_count == 0 and section_count > 0:
        warnings.append(
            f"'INSTRUMENT LIST' 표 제목은 {section_count}번 보이는데 'TAG NUMBER' 헤더를 "
            f"하나도 못 찾았어요 — 이 도면은 아직 지원하지 않는 표 형식일 수 있어요."
        )
    elif header_count > 0 and extracted < header_count * 2:
        warnings.append(
            f"'TAG NUMBER' 헤더는 {header_count}개 보이는데 태그는 {extracted}개밖에 "
            f"못 뽑았어요 — 일부 표를 놓쳤을 가능성이 있어요."
        )
    return warnings


# =====================================================================
#  PDF 하이라이트 마킹
# =====================================================================
def mark_pdf(src_path, dst_path, mark_list):
    """
    mark_list: [{"page":int, "bbox":(x0,y0,x1,y1), "color":"red"/"blue"}, ...]

    PDF 'Highlight' 주석은 일부 뷰어에서 끝이 둥글게(타원처럼) 렌더링되어
    옆 태그를 침범하는 것처럼 보이는 문제가 있어, 페이지에 직접 사각형을
    그려 넣는 방식(draw_rect)으로 대체했다. 어떤 뷰어에서도 동일하게
    각진 사각형으로 보이고, 패딩을 딱 맞게 통제할 수 있다.
    """
    doc = fitz.open(src_path)
    try:
        colors = {"red": (1, 0, 0), "blue": (0.10, 0.35, 0.95)}
        for m in mark_list:
            page = doc[m["page"]]
            x0, y0, x1, y1 = m["bbox"]
            pad = 0.15
            rect = fitz.Rect(x0 - pad, y0 - pad, x1 + pad, y1 + pad)
            page.draw_rect(rect, color=None, fill=colors.get(m["color"], (1, 1, 0)),
                            fill_opacity=0.45, overlay=True)
        doc.save(dst_path, garbage=3, deflate=True)
    finally:
        doc.close()


# =====================================================================
#  INDEX 엑셀 로드
# =====================================================================
def _norm_cell(v):
    if v is None:
        return ""
    return str(v).strip()


JB_CORE_RE = re.compile(r'(\d{3,4}-(?:' + JB_CODE_ALT + r')-\d{2,4})', re.IGNORECASE)


def normalize_jb_value(v):
    """(B.V), (SKID J/B) 같은 부가 설명은 무시한다. 'XXXX-DJB-XXX' 같은 정식
    JB 코드는 그 핵심 패턴만 뽑고, 그 외 값(다른 태그 번호를 J/B NO.로 쓰는 경우,
    'Direct Run' 등)은 뒤에 붙은 괄호 설명만 떼고 값 그대로 비교 기준으로 삼는다."""
    v = _norm_cell(v)
    if v == "" or v.upper() in ("-", "N/A", "NA", "NONE"):
        return ""
    m = JB_CORE_RE.search(v)
    if m:
        return m.group(1).upper()
    v = re.sub(r'\s*\([^)]*\)\s*$', '', v).strip()
    return v.upper()


def find_header_col(ws, header_names, fallback_letter, max_scan_rows=5):
    """헤더 텍스트로 컬럼을 찾고, 못 찾으면 고정 열(fallback_letter)을 사용."""
    targets = {re.sub(r'[\s_]', '', h).upper() for h in header_names}
    for r in range(1, max_scan_rows + 1):
        for c in range(1, ws.max_column + 1):
            val = ws.cell(row=r, column=c).value
            if val is None:
                continue
            norm = re.sub(r'[\s_]', '', str(val)).upper()
            if norm in targets:
                return c, r
    from openpyxl.utils import column_index_from_string
    return column_index_from_string(fallback_letter), 1


def load_index(index_path, log=lambda s: None):
    wb = openpyxl.load_workbook(index_path, data_only=True)
    ws = wb.active

    old_col, old_hdr_row = find_header_col(ws, ["Old_Instrument_Name"], "L")
    cur_col, cur_hdr_row = find_header_col(ws, ["Instrument_Name"], "M")
    jb_col, jb_hdr_row = find_header_col(ws, ["JUNCTION_BOX"], "BU")

    header_row = max(old_hdr_row, cur_hdr_row, jb_hdr_row)

    log(f"  - Old_Instrument_Name 열: {get_column_letter(old_col)}열")
    log(f"  - Instrument_Name 열: {get_column_letter(cur_col)}열")
    log(f"  - JUNCTION_BOX 열: {get_column_letter(jb_col)}열")

    old_set = set()
    current_set = set()
    jb_map_norm = {}   # tag -> normalized jb ("" 이면 비어있음)
    jb_map_raw = {}     # tag -> 원본 표기 (엑셀에 표시된 그대로)

    for r in range(header_row + 1, ws.max_row + 1):
        old_v = _norm_cell(ws.cell(row=r, column=old_col).value)
        cur_v = _norm_cell(ws.cell(row=r, column=cur_col).value)
        jb_v = ws.cell(row=r, column=jb_col).value

        if old_v:
            old_set.add(old_v)
        if cur_v:
            current_set.add(cur_v)
            jb_map_norm[cur_v] = normalize_jb_value(jb_v)
            jb_map_raw[cur_v] = _norm_cell(jb_v) if _norm_cell(jb_v) else "-"

    return old_set, current_set, jb_map_norm, jb_map_raw


# =====================================================================
#  엑셀 리포트 생성
# =====================================================================
HEADER_FILL = PatternFill("solid", fgColor="D1C4E9")
HEADER_FONT = Font(bold=True, color="4A3F5C", size=11)
THIN = Side(style="thin", color="B39DDB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _style_sheet(ws, headers, rows):
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    for row in rows:
        ws.append(row)
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, len(headers) + 1):
        max_len = max([len(str(headers[c - 1]))] + [len(str(r[c - 1])) for r in rows] + [10])
        ws.column_dimensions[get_column_letter(c)].width = max_len + 4
    ws.freeze_panes = "A2"


def write_tag_report(out_path, old_rows, review_rows):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Old_Instrument_Name"
    _style_sheet(ws1, ["TAG NUMBER", "Document No."], old_rows)

    ws2 = wb.create_sheet("검토가 필요한 태그")
    _style_sheet(ws2, ["TAG NUMBER", "Document No."], review_rows)

    wb.save(out_path)


def write_jb_report(out_path, jb_rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "JB NO. 검토 필요"
    _style_sheet(
        ws,
        ["LAYOUT에서의 TAG NUMBER", "LAYOUT에서의 J/B NO.", "INDEX에서의 J/B NO.", "비고", "Document No."],
        jb_rows,
    )
    wb.save(out_path)


MANUAL_TEXT = """\
설명서 - INDEX · LAYOUT 비교/검증 결과 안내
=====================================================================

이 프로그램은 INDEX 엑셀과 SECONDARY CABLE ROUTING LAYOUT 도면(PDF)의
TAG NUMBER / J-B NO. 를 자동으로 대조해서 아래 3가지 결과물을 만듭니다.


1) 마킹된_PDF 폴더 (도면 원본에 색으로 표시된 사본)
---------------------------------------------------------------------
  - 빨간색 음영
      : 그 태그가 INDEX의 "Old_Instrument_Name"(구 태그 목록)에 존재함
      → 이미 이름이 바뀌었거나 폐기된 태그이므로, 도면에서 삭제가 필요합니다.

  - 파란색 음영
      : 그 태그가 INDEX의 "Old_Instrument_Name"과 "Instrument_Name"
        (구 태그 / 현재 태그) 어느 목록에도 존재하지 않음
      → INDEX에 아예 등록되어 있지 않은 태그이므로, 오탈자인지 신규 태그
        누락인지 사람이 직접 확인(검토)이 필요합니다.

  - 음영이 없는 태그
      : INDEX의 "Instrument_Name"(현재 유효 태그)에 정상적으로 존재
      → 문제 없이 올바르게 반영된 태그입니다.


2) TAG 확인 결과_YYMMDD.xlsx
---------------------------------------------------------------------
  - Sheet "Old_Instrument_Name"
      도면(PDF)에서 발견된 태그 중, INDEX 상 이미 old(구) 태그로 분류된
      것들의 목록입니다. TAG NUMBER / 어느 도면(Document No.)에 있는지가
      정리되어 있으며, 해당 도면에서 삭제 처리가 필요합니다.

  - Sheet "검토가 필요한 태그"
      도면(PDF)에서 발견됐지만 INDEX 어디에서도 찾을 수 없는 태그
      목록입니다. TAG NUMBER / Document No. 가 정리되어 있으며,
      오탈자·미등록 여부를 사람이 직접 확인해야 합니다.


3) JB NO. 검토 필요_YYMMDD.xlsx
---------------------------------------------------------------------
  도면(LAYOUT)의 INSTRUMENT LIST에 적힌 J/B NO.와, INDEX의
  "JUNCTION_BOX" 값이 서로 다른 태그만 모아놓은 목록입니다.
  (INDEX 값이 빈칸/"-"/"N/A"인 경우는 도면에도 J/B NO.가 비어 있으면
  정상으로 간주하며, 이 표에는 나타나지 않습니다.)

  - LAYOUT에서의 TAG NUMBER : 도면에 표기된 태그 번호
  - LAYOUT에서의 J/B NO. : 도면 INSTRUMENT LIST에 적힌 J/B 번호
  - INDEX에서의 J/B NO. : INDEX의 JUNCTION_BOX 열에 적힌 값
  - 비고 : 어느 쪽에 값이 없는지 / 값이 서로 다른지 표시
  - Document No. : 해당 태그가 있는 도면 파일명


※ APC NO.를 쓰는 INSTRUMENT LIST(J/B NO. 없음)는 TAG NUMBER의 OLD/현재
  여부만 확인하며, JB NO. 검토 필요 목록에는 포함되지 않습니다.


=====================================================================
문의/오류 제보가 있으면 이 프로그램을 만든 담당자에게 알려주세요.
"""


def write_manual(output_dir):
    path = os.path.join(output_dir, "설명서.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(MANUAL_TEXT)
    return path


# =====================================================================
#  메인 처리 로직
# =====================================================================
def gather_pdf_paths(drawing_paths, temp_dir, log):
    """drawing_paths: PDF/ZIP 파일 경로가 섞여 있는 리스트 (드래그앤드롭/다중 선택으로
    한 번에 여러 개가 들어올 수 있음). ZIP은 각각 압축을 풀어 그 안의 PDF를 모두 모으고,
    PDF는 그대로 추가한다."""
    paths = []
    zip_idx = 0
    for p in drawing_paths:
        if not p or not os.path.isfile(p):
            continue
        ext = os.path.splitext(p)[1].lower()
        if ext == ".zip":
            zip_idx += 1
            extract_dir = os.path.join(temp_dir, f"zip_extracted_{zip_idx}")
            os.makedirs(extract_dir, exist_ok=True)
            try:
                with zipfile.ZipFile(p, "r") as zf:
                    zf.extractall(extract_dir)
                found = []
                for root, _, files in os.walk(extract_dir):
                    for f in files:
                        if f.lower().endswith(".pdf"):
                            found.append(os.path.join(root, f))
                paths += found
                log(f"[ZIP] {os.path.basename(p)} - {len(found)}개 PDF 압축 해제 완료")
            except Exception as e:
                log(f"[오류] ZIP 압축 해제 실패 ({os.path.basename(p)}): {e}")
        elif ext == ".pdf":
            paths.append(p)

    return paths


def _versioned_output_paths(output_dir, ts):
    """같은 출력 폴더에 이미 같은 날짜의 결과가 있으면 겹쳐쓰지 않고
    Rev.1, Rev.2, Rev.3 ... 식으로 번호를 올려가며 새 파일/폴더 경로를 만든다."""
    rev = 0
    while True:
        suffix = "" if rev == 0 else f"_Rev.{rev}"
        tag_path = os.path.join(output_dir, f"TAG 확인 결과_{ts}{suffix}.xlsx")
        jb_path = os.path.join(output_dir, f"JB NO. 검토 필요_{ts}{suffix}.xlsx")
        marked_path = os.path.join(output_dir, f"마킹된_PDF{suffix}")
        if not os.path.exists(tag_path) and not os.path.exists(jb_path) and not os.path.exists(marked_path):
            return tag_path, jb_path, marked_path
        rev += 1


def run_pipeline(index_path, drawing_paths, output_dir, log, progress_cb=None):
    with tempfile.TemporaryDirectory() as temp_dir:
        log("=" * 50)
        log("INDEX 파일 로드 중...")
        old_set, current_set, jb_map_norm, jb_map_raw = load_index(index_path, log)
        log(f"  - Old_Instrument_Name: {len(old_set)}개")
        log(f"  - Instrument_Name: {len(current_set)}개")

        pdf_paths = gather_pdf_paths(drawing_paths, temp_dir, log)
        if not pdf_paths:
            log("[오류] 검토할 PDF 파일이 없습니다.")
            return None

        log(f"총 {len(pdf_paths)}개 도면 검토 시작")
        log("-" * 50)

        old_rows = []
        review_rows = []
        jb_rows = []

        ts = datetime.now().strftime("%y%m%d")
        tag_report_path, jb_report_path, marked_dir = _versioned_output_paths(output_dir, ts)
        os.makedirs(marked_dir, exist_ok=True)

        total = len(pdf_paths)
        coverage_warning_files = []
        failed_files = []       # [(fname, error), ...] - 열기/추출 자체가 실패해서 검토를 못한 도면
        no_list_files = []      # [fname, ...] - 열리긴 했지만 INSTRUMENT LIST 표를 못 찾은 도면
        for idx, pdf_path in enumerate(pdf_paths, 1):
            fname = os.path.basename(pdf_path)
            log(f"[{idx}/{total}] {fname} 분석 중...")
            try:
                entries = extract_instrument_list(pdf_path)
            except Exception as e:
                log(f"  [오류] TAG 추출 실패: {e}")
                failed_files.append((fname, str(e)))
                if progress_cb:
                    progress_cb(idx, total)
                continue

            if not entries:
                log("  - INSTRUMENT LIST 표를 찾지 못했습니다. (도면 형식 확인 필요)")
                no_list_files.append(fname)
            else:
                cov_warnings = diagnose_extraction_coverage(pdf_path, entries)
                for w in cov_warnings:
                    log(f"  ⚠ {w}")
                if cov_warnings:
                    coverage_warning_files.append(fname)

            marks = []
            seen_old = set()
            seen_review = set()
            seen_jb_mismatch = set()

            # 같은 태그가 여러 표에 중복으로 걸려서(정상값 1번 + 빈값 1번 식으로)
            # 오탐이 나지 않도록, 태그별로 이 파일에서 발견된 모든 J/B NO. 값을
            # 먼저 모아둔다. 그 중 하나라도 INDEX와 일치하면 정상으로 처리한다.
            tag_jb_values = {}
            for entry in entries:
                if entry.get("needs_jb_check", True) and entry["jb"]:
                    tag_jb_values.setdefault(entry["tag"], set()).add(
                        normalize_jb_value(entry["jb"]))

            for entry in entries:
                tag = entry["tag"]
                needs_jb = entry.get("needs_jb_check", True)
                pdf_jb_norm = normalize_jb_value(entry["jb"]) if needs_jb else ""

                if tag in current_set:
                    # M열(Instrument_Name)이 1순위: L열에도 같은 태그가 있어도
                    # M열에 존재하면 OLD로 표시하지 않고 정상 태그로 취급한다.
                    if needs_jb and tag not in seen_jb_mismatch:
                        expected_jb = jb_map_norm.get(tag, "")
                        found_values = tag_jb_values.get(tag, set())
                        # 이 태그의 모든 발견값 중 하나라도 INDEX와 일치하면 정상.
                        if expected_jb in found_values or (not expected_jb and not found_values):
                            pass
                        else:
                            seen_jb_mismatch.add(tag)
                            layout_jb_disp = entry["jb"] if entry["jb"] else (
                                next(iter(found_values), "-") if found_values else "-")
                            index_jb_disp = jb_map_raw.get(tag, "-")
                            if not found_values and expected_jb:
                                reason = "LAYOUT엔 없음 (INDEX엔 있음)"
                            elif found_values and not expected_jb:
                                reason = "INDEX엔 없음 (LAYOUT엔 있음)"
                            else:
                                reason = "값이 서로 다름"
                            jb_rows.append([tag, layout_jb_disp, index_jb_disp, reason, fname])

                elif tag in old_set:
                    if tag not in seen_old:
                        old_rows.append([tag, fname])
                        seen_old.add(tag)
                    marks.append({"page": entry["page"], "bbox": entry["bbox"], "color": "red"})

                else:
                    if tag not in seen_review:
                        review_rows.append([tag, fname])
                        seen_review.add(tag)
                    marks.append({"page": entry["page"], "bbox": entry["bbox"], "color": "blue"})

            if marks:
                dst = os.path.join(marked_dir, fname)
                try:
                    mark_pdf(pdf_path, dst, marks)
                    log(f"  - 마킹 완료 ({len(marks)}건) -> 마킹된_PDF/{fname}")
                except Exception as e:
                    log(f"  [오류] PDF 마킹 실패: {e}")
            else:
                log("  - 마킹할 항목 없음")

            if progress_cb:
                progress_cb(idx, total)

        log("-" * 50)
        log("엑셀 리포트 생성 중...")

        write_tag_report(tag_report_path, old_rows, review_rows)
        write_jb_report(jb_report_path, jb_rows)

        log(f"완료! - 삭제필요 TAG {len(old_rows)}건 / 검토필요 TAG {len(review_rows)}건 / "
            f"JB NO. 불일치 {len(jb_rows)}건")
        if coverage_warning_files:
            log(f"⚠ 주의: {len(coverage_warning_files)}개 파일에서 추출 커버리지 경고 발생 "
                f"({', '.join(coverage_warning_files)}) - 위 로그의 ⚠ 항목을 확인해 주세요.")

        # 60개처럼 대량으로 돌릴 때 사람이 로그를 한 줄씩 다 안 보고도 바로
        # "이 도면들은 아예 검토가 안 됐다"를 알 수 있도록 마지막에 한 번 더 모아서 보여준다.
        if failed_files or no_list_files:
            log("-" * 50)
            log(f"⚠ 검토하지 못한 도면 총 {len(failed_files) + len(no_list_files)}개")
            if failed_files:
                log(f"  [열기/추출 자체가 실패 - {len(failed_files)}개] (PDF 손상 가능성)")
                for fname, err in failed_files:
                    log(f"    - {fname}  : {err}")
            if no_list_files:
                log(f"  [INSTRUMENT LIST 표를 못 찾음 - {len(no_list_files)}개] (도면 형식 확인 필요)")
                for fname in no_list_files:
                    log(f"    - {fname}")

        log(f"결과 저장 위치: {output_dir}")
        log("=" * 50)

        return {
            "old_count": len(old_rows),
            "review_count": len(review_rows),
            "jb_count": len(jb_rows),
            "coverage_warning_files": coverage_warning_files,
            "failed_files": failed_files,
            "no_list_files": no_list_files,
            "tag_report_path": tag_report_path,
            "jb_report_path": jb_report_path,
            "marked_dir": marked_dir,
        }


# =====================================================================
#  GUI
# =====================================================================
APP_W = 620
RUN_BTN_W, RUN_BTN_H = 76, 25
BADGE_SIZE = 30


def truncate_name(name, limit=34):
    if len(name) <= limit:
        return name
    keep = limit - 3
    head = int(keep * 0.6)
    tail = keep - head
    return name[:head] + "..." + name[-tail:]


class ShadowCard(ctk.CTkFrame):
    """card.jpg 스타일의 흰 배경 + 얇은 보더 카드 (그림자 없이 플랫하게)."""

    def __init__(self, master, corner_radius=12, fg_color=None, **kwargs):
        pal = PALETTES[THEME_MODE]
        super().__init__(master, fg_color=fg_color or pal["CARD"], corner_radius=corner_radius,
                          border_width=2, border_color=pal["BORDER"], **kwargs)
        self._explicit_fg = fg_color

    def restyle(self, fg_color=None):
        pal = PALETTES[THEME_MODE]
        self.configure(fg_color=fg_color or self._explicit_fg or pal["CARD"],
                        border_color=pal["BORDER"])


class FileCard(ctk.CTkFrame):
    """흰 배경 카드 + 원형 아이콘 배지 + 파일명 표시 + 알약모양 버튼.
    파일 하나만 고르면 되는 입력칸(INDEX 엑셀, 출력 폴더)에 사용한다."""

    RADIUS = 16

    def __init__(self, master, app, title, subtitle, icon_type, filetypes, is_dir=False, **kwargs):
        pal = PALETTES[THEME_MODE]
        super().__init__(master, fg_color=pal["CARD"], corner_radius=self.RADIUS,
                          border_width=2, border_color=pal["BORDER"], **kwargs)

        self.app = app
        self.title_text = title
        self.filetypes = filetypes
        self.is_dir = is_dir
        self.icon_type = icon_type
        self.path = ""

        self.grid_columnconfigure(1, weight=1)

        self.badge_img = to_ctk_image(make_badge_image(BADGE_SIZE, pal["ACCENT_CHIP"], icon_type, pal["ACCENT"]),
                                       size=(BADGE_SIZE, BADGE_SIZE))
        self.badge = ctk.CTkLabel(self, image=self.badge_img, text="")
        self.badge.grid(row=0, column=0, rowspan=2, padx=(8, 6), pady=6, sticky="w")

        self.title_lbl = ctk.CTkLabel(self, text=title, anchor="sw", height=16,
                                       font=("Malgun Gothic", 11, "bold"), text_color=pal["TEXT"])
        self.title_lbl.grid(row=0, column=1, sticky="ew", pady=(6, 0))

        self.status_lbl = ctk.CTkLabel(self, text=subtitle, anchor="nw", height=14,
                                        font=("Malgun Gothic", 10), text_color=pal["MUTED"])
        self.status_lbl.grid(row=1, column=1, sticky="ew", pady=(0, 6))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.grid(row=0, column=2, rowspan=2, padx=(4, 8), pady=4)

        self.browse_btn = ctk.CTkButton(btn_frame, text="찾아보기", width=68, height=25,
                                         corner_radius=12, font=("Malgun Gothic", 10),
                                         fg_color=pal["ACCENT"], hover_color=pal["ACCENT_DARK"],
                                         text_color="white", command=self.browse)
        self.browse_btn.pack(side="left", padx=(0, 4))

        self.clear_btn = ctk.CTkButton(btn_frame, text="✕", width=25, height=25,
                                        corner_radius=12, font=("Malgun Gothic", 10, "bold"),
                                        fg_color=pal["CARD_ALT"], hover_color=pal["DANGER_BG"],
                                        text_color=pal["MUTED"], command=self.clear)
        self.clear_btn.pack(side="left")

    def browse(self):
        if self.is_dir:
            p = filedialog.askdirectory(title=f"{self.title_text} 선택")
        else:
            p = filedialog.askopenfilename(title=f"{self.title_text} 선택", filetypes=self.filetypes)
        self.app.focus_force()
        self.app.update_idletasks()
        if p:
            self.set(p)

    def clear(self):
        self.set("")

    def set(self, path):
        self.path = path
        pal = PALETTES[THEME_MODE]
        if path:
            name = os.path.basename(path) if not self.is_dir else path
            self.status_lbl.configure(text=truncate_name(name, 40), text_color=pal["ACCENT_DARK"])
        else:
            self.status_lbl.configure(text="선택된 파일 없음", text_color=pal["MUTED"])

    def get(self):
        return self.path

    def restyle(self):
        pal = PALETTES[THEME_MODE]
        self.configure(fg_color=pal["CARD"], border_color=pal["BORDER"])
        self.title_lbl.configure(text_color=pal["TEXT"])
        self.status_lbl.configure(
            text_color=pal["ACCENT_DARK"] if self.path else pal["MUTED"])
        self.browse_btn.configure(fg_color=pal["ACCENT"], hover_color=pal["ACCENT_DARK"])
        self.clear_btn.configure(fg_color=pal["CARD_ALT"], hover_color=pal["DANGER_BG"],
                                  text_color=pal["MUTED"])
        self.badge_img = to_ctk_image(
            make_badge_image(BADGE_SIZE, pal["ACCENT_CHIP"], self.icon_type, pal["ACCENT"]),
            size=(BADGE_SIZE, BADGE_SIZE))
        self.badge.configure(image=self.badge_img)


class MultiFileDropZone(ctk.CTkFrame):
    """PDF/ZIP 파일을 몇 개든 한꺼번에 드래그하거나, 클릭해서 다중 선택으로
    추가하는 입력 존. 추가된 파일은 팝업을 열지 않아도 카드 안 목록에 바로
    보이고, 60개 도면을 여러 ZIP/PDF로 나눠 준비해도 개수 제한 없이 담을 수 있다."""

    RADIUS = 14
    LIST_H = 46

    def __init__(self, master, app, **kwargs):
        pal = PALETTES[THEME_MODE]
        super().__init__(master, fg_color=pal["CARD_ALT"], corner_radius=self.RADIUS,
                          border_width=2, border_color=pal["BORDER"], **kwargs)
        self.app = app
        self.paths = []
        self._row_widgets = []

        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x", padx=10, pady=(8, 4))

        self.badge_img = to_ctk_image(make_badge_image(BADGE_SIZE, pal["ACCENT_CHIP"], "pdf", pal["ACCENT"]),
                                       size=(BADGE_SIZE, BADGE_SIZE))
        self.badge = ctk.CTkLabel(top, image=self.badge_img, text="")
        self.badge.pack(side="left", padx=(0, 8))

        text_box = ctk.CTkFrame(top, fg_color="transparent")
        text_box.pack(side="left", fill="x", expand=True)
        self.title_lbl = ctk.CTkLabel(text_box, text="PDF · ZIP 파일을 클릭해서 추가하세요", anchor="sw", height=16,
                                       font=("Malgun Gothic", 11, "bold"), text_color=pal["TEXT"])
        self.title_lbl.pack(anchor="w", pady=(6, 0))
        self.summary_lbl = ctk.CTkLabel(text_box, text="선택된 파일 없음", anchor="nw", height=14,
                                         font=("Malgun Gothic", 10), text_color=pal["MUTED"])
        self.summary_lbl.pack(anchor="w", pady=(0, 2))

        btn_col = ctk.CTkFrame(top, fg_color="transparent")
        btn_col.pack(side="right")
        self.browse_btn = ctk.CTkButton(btn_col, text="파일 추가", width=76, height=26, corner_radius=13,
                                         font=("Malgun Gothic", 10), fg_color=pal["ACCENT"],
                                         hover_color=pal["ACCENT_DARK"], text_color="white",
                                         command=self.browse)
        self.browse_btn.pack(side="left", padx=(0, 4))
        self.clear_btn = ctk.CTkButton(btn_col, text="지우기", width=60, height=26, corner_radius=13,
                                        font=("Malgun Gothic", 10), fg_color=pal["CARD"],
                                        hover_color=pal["DANGER_BG"], text_color=pal["MUTED"],
                                        border_width=2, border_color=pal["BORDER"], command=self.clear_all)
        self.clear_btn.pack(side="left")

        # 파일 목록: 별도 팝업 없이 드롭 즉시 카드 안에서 바로 보인다.
        self.list_area = ctk.CTkScrollableFrame(self, fg_color=pal["LOG_BG"], corner_radius=10,
                                                 height=self.LIST_H)
        # CTkScrollableFrame 내부 스크롤바가 기본 200px 높이를 그대로 요구해서
        # height=self.LIST_H를 줘도 실제로는 훨씬 크게 렌더링되는 문제가 있다.
        # 스크롤바 자체 높이도 맞춰줘야 카드가 실제로 작아진다.
        self.list_area._scrollbar.configure(height=self.LIST_H)
        self.list_area.pack(fill="x", padx=10, pady=(0, 8))
        self.empty_lbl = ctk.CTkLabel(self.list_area, text="'파일 추가' 버튼으로 PDF/ZIP을 넣어주세요",
                                       font=("Malgun Gothic", 10), text_color=pal["MUTED"])
        self.empty_lbl.pack(pady=6)

        for w in (top, text_box, self.badge, self.title_lbl, self.summary_lbl):
            w.bind("<Button-1>", lambda e: self.browse())
            w.configure(cursor="hand2")

    def browse(self):
        picked = filedialog.askopenfilenames(
            title="PDF / ZIP 파일 선택 (여러 개 선택 가능)",
            filetypes=[("PDF / ZIP files", "*.pdf *.zip"), ("PDF files", "*.pdf"), ("ZIP files", "*.zip")])
        self.app.focus_force()
        self.app.update_idletasks()
        if picked:
            self.add_paths(picked)

    def add_paths(self, paths):
        added = 0
        for p in paths:
            p = os.path.normpath(p)
            ext = os.path.splitext(p)[1].lower()
            if ext not in (".pdf", ".zip") or not os.path.isfile(p):
                continue
            if p not in self.paths:
                self.paths.append(p)
                added += 1
        if added:
            self._rebuild_list()

    def remove(self, path):
        if path in self.paths:
            self.paths.remove(path)
            self._rebuild_list()

    def clear_all(self):
        self.paths = []
        self._rebuild_list()

    def get(self):
        return list(self.paths)

    def _rebuild_list(self):
        pal = PALETTES[THEME_MODE]
        for w in self._row_widgets:
            w.destroy()
        self._row_widgets = []

        if not self.paths:
            self.summary_lbl.configure(text="선택된 파일 없음", text_color=pal["MUTED"])
            self.empty_lbl.pack(pady=14)
            return

        self.empty_lbl.pack_forget()
        n_zip = sum(1 for p in self.paths if p.lower().endswith(".zip"))
        n_pdf = sum(1 for p in self.paths if p.lower().endswith(".pdf"))
        parts = []
        if n_zip:
            parts.append(f"ZIP {n_zip}개")
        if n_pdf:
            parts.append(f"PDF {n_pdf}개")
        self.summary_lbl.configure(text=" + ".join(parts) + " 선택됨", text_color=pal["ACCENT_DARK"])

        for p in self.paths:
            row = ctk.CTkFrame(self.list_area, fg_color="transparent")
            row.pack(fill="x", pady=1)
            icon_type = "zip" if p.lower().endswith(".zip") else "pdf"
            img = to_ctk_image(make_badge_image(18, pal["ACCENT_CHIP"], icon_type, pal["ACCENT"]), size=(18, 18))
            icon_lbl = ctk.CTkLabel(row, image=img, text="")
            icon_lbl._img_ref = img
            icon_lbl.pack(side="left", padx=(4, 6))
            ctk.CTkLabel(row, text=truncate_name(os.path.basename(p), 52), anchor="w",
                         font=("Malgun Gothic", 10), text_color=pal["TEXT"]).pack(
                side="left", fill="x", expand=True)

            def _remove(path=p):
                self.remove(path)

            ctk.CTkButton(row, text="✕", width=18, height=18, corner_radius=9,
                          font=("Malgun Gothic", 9, "bold"), fg_color="transparent",
                          hover_color=pal["DANGER_BG"], text_color=pal["MUTED"],
                          command=_remove).pack(side="right", padx=(0, 4))
            self._row_widgets.append(row)

    def restyle(self):
        pal = PALETTES[THEME_MODE]
        self.configure(fg_color=pal["CARD_ALT"], border_color=pal["BORDER"])
        self.title_lbl.configure(text_color=pal["TEXT"])
        self.list_area.configure(fg_color=pal["LOG_BG"])
        self.empty_lbl.configure(text_color=pal["MUTED"])
        self.browse_btn.configure(fg_color=pal["ACCENT"], hover_color=pal["ACCENT_DARK"])
        self.clear_btn.configure(fg_color=pal["CARD"], hover_color=pal["DANGER_BG"], border_color=pal["BORDER"])
        self.badge_img = to_ctk_image(make_badge_image(BADGE_SIZE, pal["ACCENT_CHIP"], "pdf", pal["ACCENT"]),
                                       size=(BADGE_SIZE, BADGE_SIZE))
        self.badge.configure(image=self.badge_img)
        self._rebuild_list()


class SectionLabel(ctk.CTkFrame):
    def __init__(self, master, number, text, **kwargs):
        pal = PALETTES[THEME_MODE]
        super().__init__(master, fg_color="transparent", **kwargs)
        self.badge_img = to_ctk_image(
            make_number_badge(22, pal["ACCENT"], pal["ACCENT_PINK"], number), size=(22, 22))
        self.badge_lbl = ctk.CTkLabel(self, image=self.badge_img, text="")
        self.badge_lbl.pack(side="left", padx=(0, 6))
        self.text_lbl = ctk.CTkLabel(self, text=text, font=("Malgun Gothic", 13, "bold"),
                                      text_color=pal["TEXT"])
        self.text_lbl.pack(side="left")
        self._number = number

    def restyle(self):
        pal = PALETTES[THEME_MODE]
        self.badge_img = to_ctk_image(
            make_number_badge(22, pal["ACCENT"], pal["ACCENT_PINK"], self._number), size=(22, 22))
        self.badge_lbl.configure(image=self.badge_img)
        self.text_lbl.configure(text_color=pal["TEXT"])


_AppBase = (ctk.CTk, TkinterDnD.DnDWrapper) if HAS_DND else (ctk.CTk,)


class App(*_AppBase):
    def __init__(self):
        super().__init__()
        if HAS_DND:
            self.TkdndVersion = TkinterDnD._require(self)
        self.title("INDEX · LAYOUT 비교/검증")
        self.geometry(f"{APP_W}x680")
        self.minsize(600, 560)
        ctk.set_appearance_mode("Light")

        pal = PALETTES[THEME_MODE]
        self.configure(fg_color=pal["BG"])

        self._result_visible = False
        self._build_widgets()
        self._fit_height()
        # CTkScrollableFrame 등 일부 위젯은 첫 배치 직후 한 박자 늦게 크기가
        # 재계산되는 경우가 있어서, 한 번 더 재측정해 창 아래에 빈 여백이
        # 남지 않게 한다.
        self.after(60, self._fit_height)

    def _fit_height(self):
        """창 높이를 내용물에 딱 맞게 다시 잡는다. geometry()로 한 번 고정 크기를
        준 뒤에는 tkinter가 내용이 늘어나도(예: 검증 완료 후 결과 표시줄 등장)
        창을 자동으로 키워주지 않으므로, 내용이 바뀔 때마다 다시 호출해야 한다."""
        self.update_idletasks()
        self.geometry(f"{APP_W}x{self.winfo_reqheight()}")

    # ---------------- 위젯 ----------------
    def _build_widgets(self):
        pal = PALETTES[THEME_MODE]
        self.outer = ctk.CTkFrame(self, fg_color="transparent")
        self.outer.pack(fill="both", expand=True, padx=12, pady=10)

        # ---- 헤더 (카드 아님 - card.jpg처럼 제목만 플랫하게) ----
        self.header = ctk.CTkFrame(self.outer, fg_color="transparent")
        self.header.pack(fill="x", pady=(0, 8))

        logo_img = to_ctk_image(make_badge_image(28, pal["ACCENT"], "tag"), size=(28, 28))
        self.logo_lbl = ctk.CTkLabel(self.header, image=logo_img, text="")
        self.logo_lbl._img_ref = logo_img
        self.logo_lbl.pack(side="left", padx=(0, 8))

        title_box = ctk.CTkFrame(self.header, fg_color="transparent")
        title_box.pack(side="left")
        self.title_lbl = ctk.CTkLabel(title_box, text="INDEX · LAYOUT 비교/검증", anchor="sw", height=20,
                                       font=("Malgun Gothic", 15, "bold"), text_color=pal["TEXT"])
        self.title_lbl.pack(anchor="w")
        self.subtitle_lbl = ctk.CTkLabel(title_box, text="INDEX ↔ SECONDARY CABLE ROUTING LAYOUT 자동 대조",
                                          anchor="nw", height=14,
                                          font=("Malgun Gothic", 10), text_color=pal["MUTED"])
        self.subtitle_lbl.pack(anchor="w")

        self.theme_btn = ctk.CTkButton(self.header, text=f"{pal['TOGGLE_ICON']} {pal['TOGGLE_LABEL']}",
                                        width=90, height=26, corner_radius=13,
                                        fg_color=pal["ACCENT_PINK"], hover_color=pal["ACCENT_DARK"],
                                        text_color="white", font=("Malgun Gothic", 10),
                                        command=self.toggle_theme)
        self.theme_btn.pack(side="right")

        self.manual_btn = ctk.CTkButton(self.header, text="설명서",
                                         width=68, height=26, corner_radius=13,
                                         fg_color=pal["CARD_ALT"], hover_color=pal["CARD_ALT_HOVER"],
                                         text_color=pal["ACCENT_DARK"], font=("Malgun Gothic", 11, "bold"),
                                         border_width=2, border_color=pal["BORDER"],
                                         command=self.show_manual)
        self.manual_btn.pack(side="right", padx=(0, 6))

        # ---- 입력 카드 ----
        self.card1 = ShadowCard(self.outer, corner_radius=12, fg_color=pal["CARD"])
        self.card1.pack(fill="x", pady=(0, 8))
        inner1 = ctk.CTkFrame(self.card1, fg_color="transparent")
        inner1.pack(fill="x", padx=12, pady=8)

        self.sec1 = SectionLabel(inner1, "1", "입력 파일")
        self.sec1.pack(anchor="w", pady=(0, 4))

        self.hint_lbl = ctk.CTkLabel(
            inner1, text="※ PDF/ZIP을 클릭해서 추가하고, INDEX 엑셀만 고르면 바로 실행돼요",
            font=("Malgun Gothic", 10), text_color=pal["ACCENT_DARK"])
        self.hint_lbl.pack(anchor="w", padx=2, pady=(0, 4))

        self.card_index = FileCard(inner1, self, "INDEX 엑셀", "선택된 파일 없음",
                                    "excel", [("Excel files", "*.xlsx *.xls")])
        self.card_index.pack(fill="x", pady=(0, 4))
        self._enable_dnd(self.card_index, (".xlsx", ".xls"))

        self.drop_zone = MultiFileDropZone(inner1, self)
        self.drop_zone.pack(fill="x")
        self._enable_multi_dnd(self.drop_zone)

        # ---- 출력 카드 ----
        self.card2 = ShadowCard(self.outer, corner_radius=12, fg_color=pal["CARD"])
        self.card2.pack(fill="x", pady=(0, 8))
        inner2 = ctk.CTkFrame(self.card2, fg_color="transparent")
        inner2.pack(fill="x", padx=12, pady=8)
        self.sec2 = SectionLabel(inner2, "2", "결과 저장 위치")
        self.sec2.pack(anchor="w", pady=(0, 4))

        out_row = ctk.CTkFrame(inner2, fg_color="transparent")
        out_row.pack(fill="x")

        self.card_output = FileCard(out_row, self, "출력 폴더", "선택된 폴더 없음",
                                     "folder", None, is_dir=True)
        self.card_output.pack(fill="x")
        self._enable_dnd(self.card_output, None, is_dir=True)

        # ---- 검증 실행 버튼 (독립된 줄 + 테두리 카드로 분리해서
        #      출력 폴더 경로가 길어져도 버튼이 밀려나서 안 보이는 일이 없게 함) ----
        self.run_card = ShadowCard(inner2, corner_radius=10, fg_color=pal["CARD_ALT"])
        self.run_card.pack(fill="x", pady=(6, 0))
        run_row = ctk.CTkFrame(self.run_card, fg_color="transparent")
        run_row.pack(pady=6)

        self.run_active_img = to_ctk_image(
            make_button_image(RUN_BTN_W, RUN_BTN_H, pal["ACCENT"], pal["ACCENT_PINK"],
                               "▶ 검증 실행", radius=RUN_BTN_H // 2, text_size=10))
        self.run_muted_img = to_ctk_image(
            make_button_image(RUN_BTN_W, RUN_BTN_H, pal["CARD_ALT"], pal["CARD_ALT"],
                               "▶ 검증 실행", radius=RUN_BTN_H // 2, text_size=10, muted=True))

        self.run_btn = ctk.CTkLabel(run_row, image=self.run_active_img, text="", cursor="hand2")
        self.run_btn.pack(side="left")
        self.run_btn.bind("<Button-1>", lambda e: self.on_run())
        self._run_enabled = True

        self.progress_lbl = ctk.CTkLabel(run_row, text="", font=("Malgun Gothic", 11, "bold"),
                                          text_color=pal["ACCENT_DARK"], width=38, anchor="w")
        self.progress_lbl.pack(side="left", padx=(6, 0))

        # ---- 결과 표시줄 (검증 완료 후에만 표시) ----
        self.result_bar = ShadowCard(self.outer, corner_radius=12, fg_color=pal["CARD"])
        inner_r = ctk.CTkFrame(self.result_bar, fg_color="transparent")
        inner_r.pack(fill="x", padx=12, pady=8)

        # 삭제필요/검토필요/JB불일치를 대시보드 스탯 타일 3개로 상시 노출
        stats_row = ctk.CTkFrame(inner_r, fg_color="transparent")
        stats_row.pack(fill="x", pady=(0, 6))
        stats_row.grid_columnconfigure((0, 1, 2), weight=1)

        self._stat_defs = [
            ("old_count", "삭제 필요 TAG", "DANGER_BG", "DANGER_TEXT"),
            ("review_count", "검토 필요 TAG", "INFO_BG", "INFO_TEXT"),
            ("jb_count", "JB NO. 불일치", "WARN_BG", "WARN_TEXT"),
        ]
        self.stat_tiles = {}
        for i, (key, label, bg_key, fg_key) in enumerate(self._stat_defs):
            tile = ctk.CTkFrame(stats_row, fg_color=pal[bg_key], corner_radius=9)
            tile.grid(row=0, column=i, sticky="ew", padx=(0 if i == 0 else 5, 0))
            num_lbl = ctk.CTkLabel(tile, text="0", font=("Malgun Gothic", 14, "bold"), text_color=pal[fg_key])
            num_lbl.pack(anchor="w", padx=8, pady=(4, 0))
            name_lbl = ctk.CTkLabel(tile, text=label, font=("Malgun Gothic", 9), text_color=pal[fg_key])
            name_lbl.pack(anchor="w", padx=8, pady=(0, 4))
            self.stat_tiles[key] = {"tile": tile, "num": num_lbl, "name": name_lbl,
                                     "bg_key": bg_key, "fg_key": fg_key}

        path_row = ctk.CTkFrame(inner_r, fg_color="transparent")
        path_row.pack(fill="x")
        self.result_icon = ctk.CTkLabel(path_row, text="✓", width=14,
                                         font=("Malgun Gothic", 11, "bold"), text_color=pal["SUCCESS_TEXT"])
        self.result_icon.pack(side="left", padx=(0, 4))
        self.result_lbl = ctk.CTkLabel(path_row, text="", anchor="w",
                                        font=("Malgun Gothic", 10), text_color=pal["MUTED"])
        self.result_lbl.pack(side="left", fill="x", expand=True)
        self.result_open_btn = ctk.CTkButton(path_row, text="폴더 열기", width=68, height=24,
                                              corner_radius=12, font=("Malgun Gothic", 10),
                                              fg_color=pal["ACCENT"], hover_color=pal["ACCENT_DARK"],
                                              text_color="white", command=self.open_result_folder)
        self.result_open_btn.pack(side="right")
        self._result_path = ""
        # 처음엔 숨김 (검증 완료 시 pack)

        # ---- 로그 카드 ----
        self.card3 = ShadowCard(self.outer, corner_radius=12, fg_color=pal["CARD"])
        self.card3.pack(fill="x")
        inner3 = ctk.CTkFrame(self.card3, fg_color="transparent")
        inner3.pack(fill="x", padx=12, pady=8)
        self.sec3 = SectionLabel(inner3, "3", "진행 로그")
        self.sec3.pack(anchor="w", pady=(0, 4))

        # 로그가 쌓이기 전에는 안내문을 보여주고, 검증이 시작되면(첫 로그 줄)
        # 실제 로그 텍스트박스로 바꿔치기한다.
        self.log_empty_frame = ctk.CTkFrame(inner3, fg_color=pal["LOG_BG"], corner_radius=10)
        self.log_empty_frame.pack(fill="x")
        self.log_empty_lbl = ctk.CTkLabel(
            self.log_empty_frame, text="검증을 실행하면 로그가 여기에 표시됩니다.",
            font=("Malgun Gothic", 10), text_color=pal["MUTED"])
        self.log_empty_lbl.pack(pady=(0, 10))

        self.log_box = ctk.CTkTextbox(inner3, corner_radius=10, fg_color=pal["LOG_BG"],
                                       text_color=pal["TEXT"], font=("Consolas", 10), height=80)
        self.log_box.configure(state="disabled")
        self._log_started = False

    # ---------------- 테마 토글 ----------------
    def toggle_theme(self):
        global THEME_MODE
        THEME_MODE = "dark" if THEME_MODE == "light" else "light"
        pal = PALETTES[THEME_MODE]
        ctk.set_appearance_mode("Dark" if THEME_MODE == "dark" else "Light")

        self.configure(fg_color=pal["BG"])
        self.title_lbl.configure(text_color=pal["TEXT"])
        self.subtitle_lbl.configure(text_color=pal["MUTED"])
        self.theme_btn.configure(text=f"{pal['TOGGLE_ICON']} {pal['TOGGLE_LABEL']}",
                                  fg_color=pal["ACCENT_PINK"], hover_color=pal["ACCENT_DARK"])
        self.manual_btn.configure(fg_color=pal["CARD_ALT"], hover_color=pal["CARD_ALT_HOVER"],
                                   text_color=pal["ACCENT_DARK"], border_color=pal["BORDER"])

        logo_img = to_ctk_image(make_badge_image(28, pal["ACCENT"], "tag"), size=(28, 28))
        self.logo_lbl.configure(image=logo_img)
        self.logo_lbl._img_ref = logo_img

        for card in (self.card1, self.card2, self.card3, self.result_bar):
            card.restyle(fg_color=pal["CARD"])

        self.sec1.restyle()
        self.sec2.restyle()
        self.sec3.restyle()
        self.hint_lbl.configure(text_color=pal["ACCENT_DARK"])

        self.card_index.restyle()
        self.card_output.restyle()
        self.drop_zone.restyle()
        self.run_card.restyle(fg_color=pal["CARD_ALT"])

        self.run_active_img = to_ctk_image(
            make_button_image(RUN_BTN_W, RUN_BTN_H, pal["ACCENT"], pal["ACCENT_PINK"],
                               "▶ 검증 실행", radius=RUN_BTN_H // 2, text_size=10))
        self.run_muted_img = to_ctk_image(
            make_button_image(RUN_BTN_W, RUN_BTN_H, pal["CARD_ALT"], pal["CARD_ALT"],
                               "▶ 검증 실행", radius=RUN_BTN_H // 2, text_size=10, muted=True))
        self.run_btn.configure(image=self.run_active_img if self._run_enabled else self.run_muted_img)
        self.progress_lbl.configure(text_color=pal["ACCENT_DARK"])

        self.log_box.configure(fg_color=pal["LOG_BG"], text_color=pal["TEXT"])
        self.log_empty_frame.configure(fg_color=pal["LOG_BG"])
        self.log_empty_lbl.configure(text_color=pal["MUTED"])

        self.result_icon.configure(text_color=pal["SUCCESS_TEXT"])
        self.result_lbl.configure(text_color=pal["MUTED"])
        self.result_open_btn.configure(fg_color=pal["ACCENT"], hover_color=pal["ACCENT_DARK"])
        for tile in self.stat_tiles.values():
            tile["tile"].configure(fg_color=pal[tile["bg_key"]])
            tile["num"].configure(text_color=pal[tile["fg_key"]])
            tile["name"].configure(text_color=pal[tile["fg_key"]])

    # ---------------- 로그 / 진행률 ----------------
    def log(self, msg):
        def _do():
            if not self._log_started:
                self._log_started = True
                self.log_empty_frame.pack_forget()
                self.log_box.pack(fill="x")
                self._fit_height()
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _do)

    def set_progress(self, cur, total):
        pct = int(cur / total * 100) if total else 0

        def _do():
            self.progress_lbl.configure(text=f"{pct}%")
        self.after(0, _do)

    def set_run_enabled(self, enabled):
        self._run_enabled = enabled
        pal = PALETTES[THEME_MODE]
        self.run_btn.configure(image=self.run_active_img if enabled else self.run_muted_img)
        self.run_btn.configure(cursor="hand2" if enabled else "arrow")

    def show_result(self, path, result):
        self._result_path = path
        self.result_lbl.configure(text=f"결과 저장 위치:  {path}")
        for key, _label, _bg_key, _fg_key in self._stat_defs:
            self.stat_tiles[key]["num"].configure(text=str(result.get(key, 0)))
        if not self._result_visible:
            self.result_bar.pack(fill="x", pady=(0, 8), before=self.card3)
            self._result_visible = True
            self._fit_height()

    def open_result_folder(self):
        if self._result_path:
            try:
                os.startfile(self._result_path)
            except Exception:
                pass

    def _enable_dnd(self, card, valid_exts, is_dir=False):
        """카드 위에 파일/폴더를 드래그해서 놓으면 자동으로 선택되게 한다.
        (tkinterdnd2가 설치되어 있지 않으면 조용히 건너뛴다 - 찾아보기 버튼은
        그대로 동작하므로 필수 기능은 아니다.)"""
        if not HAS_DND:
            return

        def on_drop(event):
            try:
                paths = self.tk.splitlist(event.data)
            except Exception:
                paths = [event.data]
            for p in paths:
                p = p.strip("{}")
                if is_dir:
                    if os.path.isdir(p):
                        card.set(p)
                        break
                else:
                    ext = os.path.splitext(p)[1].lower()
                    if os.path.isfile(p) and (not valid_exts or ext in valid_exts):
                        card.set(p)
                        break

        try:
            card.drop_target_register(DND_FILES)
            card.dnd_bind("<<Drop>>", on_drop)
        except Exception:
            pass

    def _enable_multi_dnd(self, zone):
        """드롭 존 위에 PDF/ZIP 파일을 한 번에 여러 개 드래그해서 놓으면
        전부 목록에 추가한다 (한 개씩만 받던 기존 카드 방식과 달리 개수 제한 없음)."""
        if not HAS_DND:
            return

        def on_drop(event):
            try:
                paths = self.tk.splitlist(event.data)
            except Exception:
                paths = [event.data]
            cleaned = [p.strip("{}") for p in paths]
            zone.add_paths(cleaned)

        try:
            zone.drop_target_register(DND_FILES)
            zone.dnd_bind("<<Drop>>", on_drop)
        except Exception:
            pass

    def show_manual(self):
        pal = PALETTES[THEME_MODE]
        win = ctk.CTkToplevel(self)
        win.title("설명서")
        win.geometry("560x600")
        win.configure(fg_color=pal["BG"])
        win.transient(self)
        win.grab_set()

        header = ctk.CTkLabel(win, text="📖 설명서 - 결과 파일 보는 법",
                               font=("Malgun Gothic", 16, "bold"), text_color=pal["TEXT"])
        header.pack(anchor="w", padx=18, pady=(18, 8))

        box = ctk.CTkTextbox(win, corner_radius=12, fg_color=pal["CARD"],
                              text_color=pal["TEXT"], font=("Malgun Gothic", 12), wrap="word")
        box.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        box.insert("1.0", MANUAL_TEXT)
        box.configure(state="disabled")

        close_btn = ctk.CTkButton(win, text="닫기", width=100, height=32, corner_radius=16,
                                   fg_color=pal["ACCENT"], hover_color=pal["ACCENT_DARK"],
                                   text_color="white", command=win.destroy)
        close_btn.pack(pady=(0, 16))

    # ---------------- 실행 ----------------
    def on_run(self):
        if not self._run_enabled:
            return

        index_path = self.card_index.get()
        drawing_paths = self.drop_zone.get()
        output_dir = self.card_output.get()

        if not index_path:
            messagebox.showwarning("확인", "INDEX 엑셀 파일을 선택해 주세요.")
            return
        if not drawing_paths:
            messagebox.showwarning("확인", "PDF 또는 ZIP 파일을 최소 1개 이상 추가해 주세요.")
            return
        if not output_dir:
            messagebox.showwarning("확인", "결과를 저장할 출력 폴더를 선택해 주세요.")
            return

        self.set_run_enabled(False)
        self.progress_lbl.configure(text="0%")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        def worker():
            try:
                result = run_pipeline(index_path, drawing_paths, output_dir,
                                       log=self.log, progress_cb=self.set_progress)
                if result:
                    self.after(0, lambda: self.show_result(output_dir, result))
                    warn_files = result.get("coverage_warning_files") or []
                    warn_line = (
                        f"\n\n⚠ 주의: {len(warn_files)}개 파일에서 추출이 예상보다 적을 수 있어요.\n"
                        f"로그의 ⚠ 항목을 확인해 주세요."
                    ) if warn_files else ""

                    failed_files = result.get("failed_files") or []
                    no_list_files = result.get("no_list_files") or []
                    skipped_total = len(failed_files) + len(no_list_files)
                    if skipped_total:
                        names = [f for f, _ in failed_files] + list(no_list_files)
                        shown = ", ".join(names[:5])
                        more = f" 외 {len(names) - 5}개" if len(names) > 5 else ""
                        skipped_line = (
                            f"\n\n❗ 검토하지 못한 도면 {skipped_total}개: {shown}{more}\n"
                            f"로그에서 자세한 사유를 확인해 주세요."
                        )
                    else:
                        skipped_line = ""

                    self.after(0, lambda: messagebox.showinfo(
                        "완료",
                        f"검증이 완료되었습니다.\n\n"
                        f"- 삭제 필요 TAG: {result['old_count']}건\n"
                        f"- 검토 필요 TAG: {result['review_count']}건\n"
                        f"- JB NO. 불일치: {result['jb_count']}건\n\n"
                        f"결과 파일 위치:\n{output_dir}\n\n"
                        f"(파일/시트 설명은 상단 '설명서' 버튼에서 볼 수 있어요)"
                        f"{warn_line}"
                        f"{skipped_line}"
                    ))
                    try:
                        os.startfile(output_dir)
                    except Exception:
                        pass
            except Exception:
                err = traceback.format_exc()
                self.log(f"[치명적 오류]\n{err}")
                self.after(0, lambda: messagebox.showerror("오류", "처리 중 오류가 발생했습니다.\n로그를 확인해 주세요."))
            finally:
                self.after(0, lambda: self.set_run_enabled(True))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = App()
    app.mainloop()
