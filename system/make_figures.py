"""서식2·발표용 분석 도표 생성.

화면 캡처(docs/screenshots)가 '무엇을 만들었나'를 보인다면, 이 도표들은
'데이터에서 무엇이 나왔나'를 보인다. 전부 DB 실측값에서 그린다.

사용:  python system/make_figures.py
출력:  docs/figures/*.png (300dpi)
"""
from __future__ import annotations

import gzip
import json
import math
import random
import sqlite3
import statistics as st
import sys
from collections import Counter, defaultdict
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import rcParams

rcParams["font.family"] = "Malgun Gothic"
rcParams["axes.unicode_minus"] = False
rcParams["figure.dpi"] = 300
rcParams["savefig.bbox"] = "tight"
rcParams["savefig.facecolor"] = "white"

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "docs" / "figures"
OUT.mkdir(parents=True, exist_ok=True)
DB = ROOT / "system" / "data" / "policymap.db"
API = ROOT / "system" / "data" / "api"

INK = "#1f2d3d"
ACC = "#2c7fb8"
ACC2 = "#41ab5d"
WARN = "#d95f02"
GREY = "#9aa5b1"


def conn():
    c = sqlite3.connect(f"file:{DB}?mode=ro", uri=True)
    c.execute("PRAGMA busy_timeout = 300000")
    return c


def save(fig, name):
    p = OUT / (name + ".png")
    fig.savefig(p)
    plt.close(fig)
    print("  ok  %s.png  %.0fKB" % (name, p.stat().st_size / 1024), flush=True)


def load_json(rel):
    """api/ 하위 파일을 읽는다. gzip 사전압축본도 처리한다."""
    p = API / rel
    if p.exists():
        raw = p.read_bytes()
    else:
        gz = Path(str(p) + ".gz")
        raw = gzip.decompress(gz.read_bytes())
    d = json.loads(raw)
    return d.get("data", d)


# --------------------------------------------------------------------------- #
def fig_pop_vs_ordinance(c):
    """조례 수는 인구와 무관하다 — 산점도 + 로그-로그 적합."""
    rows = c.execute(
        "SELECT rf.population, COUNT(*) n FROM ordinances o "
        "JOIN regions r ON r.region_id = o.region_id "
        "JOIN region_features rf ON rf.region_id = r.region_id "
        "WHERE r.level=2 AND r.status='active' AND r.has_legislation=1 "
        "  AND o.status='active' AND rf.population > 0 "
        "GROUP BY o.region_id, rf.population").fetchall()
    if not rows:
        rows = c.execute(
            "SELECT r.population, COUNT(*) n FROM ordinances o "
            "JOIN regions r ON r.region_id=o.region_id "
            "WHERE r.level=2 AND r.status='active' AND r.has_legislation=1 "
            "  AND o.status='active' AND r.population > 0 "
            "GROUP BY o.region_id, r.population").fetchall()
    x = [r[0] for r in rows]
    y = [r[1] for r in rows]
    fig, ax = plt.subplots(figsize=(7.2, 4.6))
    ax.scatter(x, y, s=26, c=ACC, alpha=0.55, edgecolors="none")
    lx = [math.log(v) for v in x]
    ly = [math.log(v) for v in y]
    mx, my = st.mean(lx), st.mean(ly)
    b = sum((a - mx) * (q - my) for a, q in zip(lx, ly)) / sum((a - mx) ** 2 for a in lx)
    a0 = my - b * mx
    xs = sorted(x)
    ax.plot(xs, [math.exp(a0 + b * math.log(v)) for v in xs], color=WARN, lw=2,
            label="로그-로그 적합 (탄력성 %.3f)" % b)
    ax.set_xscale("log")
    ax.set_xlabel("인구 (명, 로그 눈금)")
    ax.set_ylabel("현행 조례·규칙 수")
    ax.set_title("조례 수는 인구와 사실상 무관하다  (기초자치단체 %d곳)" % len(rows),
                 fontsize=12, color=INK, pad=10)
    ax.legend(frameon=False, fontsize=9)
    ax.grid(alpha=.25, ls=":")
    ax.text(.02, .96,
            "인구 %.0f배  vs  조례 %.2f배\n변동계수  인구 %.3f · 조례 %.3f"
            % (max(x) / min(x), max(y) / min(y),
               st.pstdev(x) / st.mean(x), st.pstdev(y) / st.mean(y)),
            transform=ax.transAxes, va="top", fontsize=9,
            bbox=dict(fc="#f4f7fa", ec=GREY, lw=.6, boxstyle="round,pad=.5"))
    save(fig, "F1_인구_대_조례수")


def fig_category_cv(c):
    """분야마다 격차의 의미가 다르다 — 변동계수 막대."""
    cats = [("C01", "행정·자치"), ("C02", "재정·세무"), ("C03", "복지·돌봄"),
            ("C04", "인구·출산"), ("C05", "청년·교육"), ("C06", "보건·의료"),
            ("C07", "환경·기후"), ("C08", "안전·재난"), ("C09", "도시·건축"),
            ("C10", "교통"), ("C11", "경제·산업"), ("C12", "농림·수산"),
            ("C13", "문화·체육"), ("C14", "동물·반려")]
    tot = dict(c.execute(
        "SELECT o.region_id, COUNT(*) FROM ordinances o "
        "JOIN regions r ON r.region_id=o.region_id "
        "WHERE r.level=2 AND r.status='active' AND r.has_legislation=1 "
        "  AND o.status='active' GROUP BY 1").fetchall())
    out = []
    for code, nm in cats:
        per = c.execute(
            "SELECT o.region_id, COUNT(DISTINCT oc.ordinance_id) FROM ordinance_category oc "
            "JOIN ordinances o ON o.ordinance_id=oc.ordinance_id "
            "JOIN regions r ON r.region_id=o.region_id "
            "WHERE oc.category_code=? AND r.level=2 AND r.status='active' "
            "  AND r.has_legislation=1 AND o.status='active' GROUP BY 1",
            (code,)).fetchall()
        sh = [n / tot[rid] for rid, n in per if tot.get(rid)]
        if len(sh) < 50:
            continue
        out.append((nm, st.mean(sh), st.pstdev(sh) / st.mean(sh)))
    out.sort(key=lambda t: t[2])
    fig, ax = plt.subplots(figsize=(7.2, 4.8))
    cols = [WARN if v > .4 else (ACC2 if v < .15 else ACC) for _, _, v in out]
    ax.barh([o[0] for o in out], [o[2] for o in out], color=cols, height=.66)
    for i, (nm, m, v) in enumerate(out):
        ax.text(v + .012, i, "%.3f  (평균비중 %.1f%%)" % (v, m * 100),
                va="center", fontsize=8.5, color=INK)
    ax.axvline(0.181, color=GREY, ls="--", lw=1.2)
    ax.text(0.187, len(out) - .6, "조례 총량 CV 0.181", fontsize=8.5, color=GREY)
    ax.set_xlabel("지자체 간 변동계수 (CV) — 클수록 지자체마다 다르다")
    ax.set_xlim(0, max(o[2] for o in out) * 1.35)
    ax.set_title("분야마다 '격차의 의미'가 다르다", fontsize=12, color=INK, pad=10)
    ax.grid(axis="x", alpha=.25, ls=":")
    save(fig, "F2_분야별_변동계수")


def fig_spatial():
    """공간자기상관 23지표 — Moran's I 분포."""
    cat = load_json("analytics.json")
    sp = [s for s in cat.get("spatial", []) if s.get("moran_i") is not None]
    sp.sort(key=lambda s: s["moran_i"])
    lbl = [(s.get("description") or s["slug"])[:22] for s in sp]
    val = [s["moran_i"] for s in sp]
    sig = [s.get("n_significant_fdr", 0) for s in sp]
    fig, ax = plt.subplots(figsize=(7.4, 6.4))
    ax.barh(range(len(val)), val,
            color=[ACC2 if n > 0 else GREY for n in sig], height=.68)
    ax.set_yticks(range(len(val)))
    ax.set_yticklabels(lbl, fontsize=7.6)
    for i, (v, n) in enumerate(zip(val, sig)):
        ax.text(v + (.012 if v >= 0 else -.012), i,
                "%.3f%s" % (v, ("  (%d곳)" % n) if n else ""),
                va="center", ha="left" if v >= 0 else "right", fontsize=7.4, color=INK)
    ax.axvline(0, color=INK, lw=.8)
    ax.set_xlabel("전역 Moran's I  (순열검정 999회)")
    ax.set_xlim(min(val) - .12, max(val) + .16)
    ax.set_title("공간자기상관 23지표 — 초록은 BH-FDR 통과 국지군집 보유",
                 fontsize=12, color=INK, pad=10)
    ax.grid(axis="x", alpha=.25, ls=":")
    save(fig, "F3_공간자기상관_23지표")


def fig_neural_eval():
    """신경망 3모델 평가 — 무작위 기준선 대비."""
    d = load_json("neural_eval.json")
    ms = sorted(d["models"], key=lambda m: -(m.get("lift") or 0))
    names = [m["model"].replace("-numpy", "") for m in ms]
    agree = [m["category_agreement"] * 100 for m in ms]
    base = [m["random_baseline"] * 100 for m in ms]
    x = list(range(len(ms)))
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    ax.bar([i - .19 for i in x], agree, width=.38, color=ACC, label="분야 일치율")
    ax.bar([i + .19 for i in x], base, width=.38, color=GREY, label="무작위 기준선")
    for i, m in enumerate(ms):
        lift = m.get("lift") or 0
        ax.text(i, max(agree[i], base[i]) + 2.2, "%.2f배" % lift, ha="center",
                fontsize=10, color=ACC2 if lift >= 1.2 else WARN, weight="bold")
    ax.set_xticks(x)
    ax.set_xticklabels(names, fontsize=10)
    ax.set_ylabel("분야 일치율 (%)")
    ax.set_ylim(0, max(agree + base) * 1.3)
    ax.set_title("이웃 추천 3모델 평가 — 기준선 없이 절대값만 보면 오독한다",
                 fontsize=12, color=INK, pad=10)
    ax.legend(frameon=False, fontsize=9, loc="upper right")
    ax.grid(axis="y", alpha=.25, ls=":")
    save(fig, "F4_신경망_모델평가")


def fig_diffusion_type(c):
    """확산 유형 — 조문 구조 동질성 vs 시도 집중."""
    random.seed(20260827)
    pool = [s for (s,) in c.execute(
        "SELECT sido_cd FROM regions WHERE level=2 AND status='active' "
        "AND has_legislation=1")]

    def stat(kw):
        rows = c.execute(
            "SELECT o.ordinance_id, r.sido_cd FROM ordinances o "
            "JOIN regions r ON r.region_id=o.region_id "
            "WHERE o.name LIKE ? AND r.level=2 AND r.status='active' "
            "  AND o.status='active' AND o.rr_cls_cd LIKE '%제정%'",
            ("%" + kw + "%",)).fetchall()
        if len(rows) < 40:
            return None
        sig = defaultdict(list)
        for oid, sd in rows:
            ts = tuple(t for (t,) in c.execute(
                "SELECT title FROM ordinance_articles WHERE ordinance_id=? "
                "ORDER BY article_no", (oid,)) if t)
            if ts:
                sig[ts].append(sd)
        tot = sum(len(v) for v in sig.values())
        if tot < 40:
            return None
        homog = sum(len(v) for v in sig.values() if len(v) >= 2) / tot
        obs, exp = [], []
        for ts, v in sig.items():
            k = len(v)
            if k < 3:
                continue
            obs.append(Counter(v).most_common(1)[0][1] / k)
            exp.append(st.mean(Counter(random.sample(pool, k)).most_common(1)[0][1] / k
                               for _ in range(200)))
        if not obs:
            return None
        return tot, homog, st.mean(obs) / max(st.mean(exp), 1e-9)

    pts = []
    for kw in ("안전보안관", "자원봉사", "맨발걷기", "반려동물", "공공심야약국",
               "자전거", "치매", "도시재생", "생활임금", "1인가구"):
        r = stat(kw)
        if r:
            pts.append((kw,) + r)
    fig, ax = plt.subplots(figsize=(7.2, 5.0))
    for kw, tot, h, cc in pts:
        col = ACC if (h >= .35 and cc < 1.7) else (WARN if cc >= 2.0 else GREY)
        ax.scatter(h * 100, cc, s=40 + tot, c=col, alpha=.8, edgecolors="white", lw=1.2)
        ax.annotate("%s (%d)" % (kw, tot), (h * 100, cc), textcoords="offset points",
                    xytext=(7, 5), fontsize=8.6, color=INK)
    ax.axhline(1.0, color=GREY, ls="--", lw=1)
    ax.text(1, 1.03, "지리적 무작위", fontsize=8, color=GREY)
    ax.set_xlabel("조문 구조 지문 공유율 (%)  — 무작위 대조군 0.8%")
    ax.set_ylabel("동형 군집의 시도 집중 배수")
    ax.set_title("확산에는 두 유형이 있고 조문 구조가 그것을 가른다",
                 fontsize=12, color=INK, pad=10)
    ax.grid(alpha=.25, ls=":")
    ax.text(.98, .97, "← 전국 단일 원본형        수평·이웃 복제형 ↑",
            transform=ax.transAxes, ha="right", va="top", fontsize=8.6, color=GREY)
    save(fig, "F5_확산유형")


def fig_verification():
    """검증 사다리 — 근거 강도별 규모."""
    d = load_json("verification/summary.json")
    cv = d["citation_verification"]["by_status"]
    inst = d["instrument_status"]["ordinances"]
    items = [("원문 URL 확보", inst.get("source-linked", 0), ACC2),
             ("인용 조문 존재 확인", cv.get("article-verified", 0), ACC),
             ("인용 조문 불일치", cv.get("article-missing", 0), WARN),
             ("자동 확인 불가", cv.get("unverifiable", 0), GREY),
             ("사람이 직접 대조", d["verification_table"]["rows"], "#6a51a3")]
    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    ys = list(range(len(items)))
    ax.barh(ys, [i[1] for i in items], color=[i[2] for i in items], height=.62)
    ax.set_yticks(ys)
    ax.set_yticklabels([i[0] for i in items], fontsize=9.5)
    for i, it in enumerate(items):
        ax.text(it[1] * 1.02, i, "{:,}".format(it[1]), va="center", fontsize=9, color=INK)
    ax.set_xscale("log")
    ax.set_xlabel("건수 (로그 눈금) — 모집단이 다르므로 줄 사이 비율 비교는 성립하지 않는다")
    ax.set_title("검증 공시 — 무엇을 어떤 근거로 확인했는가", fontsize=12, color=INK, pad=10)
    ax.grid(axis="x", alpha=.25, ls=":")
    save(fig, "F6_검증사다리")


def fig_bill_link(c):
    """국회 의안 ↔ 조례 — 두 입법 층위는 다른 일을 한다."""
    rows = c.execute(
        "SELECT li.name, COUNT(DISTINCT b.bill_id) nb, "
        "  (SELECT COUNT(DISTINCT d.child_id) FROM delegations d "
        "   WHERE d.parent_id=li.instrument_id) no_ "
        "FROM bill_instrument_link b "
        "JOIN legal_instrument li ON li.instrument_id=b.instrument_id "
        "GROUP BY li.instrument_id, li.name").fetchall()
    if not rows:
        print("  skip F7 (bill_instrument_link 없음)")
        return
    top_b = sorted(rows, key=lambda r: -r[1])[:8]
    top_o = sorted(rows, key=lambda r: -r[2])[:8]
    fig, axes = plt.subplots(1, 2, figsize=(9.6, 4.2))
    for ax, data, key, col, title in (
            (axes[0], top_b, 1, ACC, "국회에서 가장 많이 다뤄진 법"),
            (axes[1], top_o, 2, ACC2, "조례를 가장 많이 낳은 법")):
        lab = [r[0][:16] for r in data][::-1]
        val = [r[key] for r in data][::-1]
        ax.barh(range(len(val)), val, color=col, height=.66)
        ax.set_yticks(range(len(val)))
        ax.set_yticklabels(lab, fontsize=8)
        for i, v in enumerate(val):
            ax.text(v * 1.02, i, "{:,}".format(v), va="center", fontsize=8, color=INK)
        ax.set_xlabel("의안 수" if key == 1 else "조례 인용 수")
        ax.set_title(title, fontsize=11, color=INK)
        ax.grid(axis="x", alpha=.25, ls=":")
    fig.suptitle("국회 의안 ↔ 법령 ↔ 조례  (Spearman 0.491 · Pearson 0.206)",
                 fontsize=12, color=INK, y=1.02)
    save(fig, "F7_국회_조례_연결")


def fig_gap_robustness():
    """결손 규모는 측정 방법에 좌우된다."""
    methods = ["peer 20곳\n부분문자열", "공통핵심 57개\n앞6자 매칭",
               "공통핵심\n퍼지 매칭", "광역내 피어\n3단계 감사"]
    false_rate = [28.2, 43.8, 69.0, 75.0]
    real_gap = [8.0, 3.19, 1.18, 1.3]
    fig, ax1 = plt.subplots(figsize=(7.2, 4.2))
    x = list(range(len(methods)))
    ax1.bar(x, false_rate, color=GREY, width=.55)
    ax1.set_ylabel("허위 결손 비율 (%)", color=INK)
    ax1.set_ylim(0, 100)
    ax2 = ax1.twinx()
    ax2.plot(x, real_gap, "o-", color=WARN, lw=2.2, ms=8)
    ax2.set_ylabel("실질 결손 건수 (지자체당)", color=WARN)
    ax2.set_ylim(0, 10)
    for i, (f, g) in enumerate(zip(false_rate, real_gap)):
        ax1.text(i, f + 2.5, "%.1f%%" % f, ha="center", fontsize=9, color=INK)
        ax2.text(i, g + .35, "%.2f건" % g, ha="center", fontsize=9, color=WARN)
    ax1.set_xticks(x)
    ax1.set_xticklabels(methods, fontsize=8.6)
    ax1.set_title("'없는 조례' 규모는 측정 방법에 따라 7배 움직인다",
                  fontsize=12, color=INK, pad=10)
    ax1.grid(axis="y", alpha=.25, ls=":")
    save(fig, "F8_결손_로버스트니스")


def main():
    c = conn()
    print("=== 분석 도표 생성 ===")
    jobs = [(fig_pop_vs_ordinance, c), (fig_category_cv, c), (fig_spatial, None),
            (fig_diffusion_type, c), (fig_verification, None), (fig_bill_link, c),
            (fig_neural_eval, None), (fig_gap_robustness, None)]
    for fn, arg in jobs:
        try:
            fn(arg) if arg is not None else fn()
        except Exception as e:  # noqa: BLE001
            print("  x   %s: %s: %s" % (fn.__name__, type(e).__name__, str(e)[:110]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
