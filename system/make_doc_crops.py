#!/usr/bin/env python
"""화면 스크린샷 → 문서 삽입용 크롭본.

서식2 본문은 폭이 170mm 뿐이라 세로로 긴 원본 스크린샷을 그대로 넣으면
한 장이 반 쪽을 먹는다. 화면 상단(제목·핵심 지표·주요 시각화)만 남기고
가로:세로 = 1.9:1 로 잘라 둔다.

사용:  python system/make_doc_crops.py
입력:  docs/screenshots/*.png
출력:  docs/screenshots/doc/*.png   (1600 × 842)
"""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "docs" / "screenshots"
DST = SRC / "doc"
RATIO = 1.9          # 폭:높이
WIDTH = 1600


def main() -> int:
    if not SRC.exists():
        print(f"원본이 없다: {SRC}")
        return 1
    DST.mkdir(parents=True, exist_ok=True)
    n = 0
    for p in sorted(SRC.glob("*.png")):
        im = Image.open(p)
        w, h = im.size
        target_h = int(round(w / RATIO))
        if target_h > h:                      # 원본이 이미 납작하면 그대로
            out = im.copy()
        else:
            out = im.crop((0, 0, w, target_h))
        if out.size[0] != WIDTH:              # 폭 통일
            out = out.resize((WIDTH, int(round(out.size[1] * WIDTH / out.size[0]))),
                             Image.LANCZOS)
        out.save(DST / p.name, optimize=True)
        n += 1
        print(f"  ok  {p.name}  {w}×{h} → {out.size[0]}×{out.size[1]}")
    print(f"\n  크롭 {n}장 → {DST}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
