# -*- coding: utf-8 -*-
"""캔바 업로드용 pptx 사본을 만든다. 원본은 건드리지 않는다.

캔바는 올릴 때 임베드 글꼴을 읽지 않고 애니메이션을 버린다. 그래서 둘을 뺀다.

굵기 문제도 여기서 잡는다. 캔바의 글꼴은 웨이트가 이름 단위로 갈린다
(네모고딕은 가벼운·보통·중간·가장 굵은 넷). pptx 에 "네모고딕 Heavy" 라고
적혀 있어도 `b="true"` 가 함께 있으면 캔바가 이름의 웨이트를 무시하고 한 단계
아래인 "중간" 으로 잡는다. b 를 빼면 이름대로 "가장 굵은" 을 쓴다.

파워포인트는 반대다. b 를 빼면 굵게 보이지 않는다. 그래서 녹화용 원본은
그대로 두고 캔바용 사본에서만 뺀다.

    python src/for_canva.py "대상폴더" ["대상폴더2" ...]
    python src/for_canva.py "파일.pptx"

산출물은 `대상폴더/캔바업로드용/*_캔바용.pptx`.
"""
import os
import re
import sys
import zipfile

TIMING = re.compile(rb"<p:timing>.*?</p:timing>", re.S)
EMBED = re.compile(rb"<p:embeddedFontLst>.*?</p:embeddedFontLst>", re.S)
RPR = re.compile(rb"<a:rPr([^>]*)>(.*?)</a:rPr>", re.S)
BOLD = re.compile(rb'\s*b="(?:true|1)"')
TYPEFACE = re.compile(rb'typeface="([^"]*)"')
SLIDE = re.compile(r"ppt/slides/slide\d+\.xml$")

# 이름에 이 웨이트가 들어 있으면 b 를 뺀다. 중간 이상만 대상으로 한다.
HEAVY = ("medium", "semibold", "semi-bold", "demibold", "demi-bold",
         "bold", "extrabold", "ultrabold", "black", "heavy")
HEAVY_KO = ("중간", "세미볼드", "볼드", "헤비", "굵은", "두꺼운")

# 이름에 이 웨이트가 있으면 b 는 일부러 넣은 가짜 볼드로 보고 그대로 둔다.
LIGHT = ("thin", "extralight", "ultralight", "light")
LIGHT_KO = ("가는", "얇은", "라이트")

_HEAVY_RE = re.compile(
    "|".join(r"(?<![a-z])" + re.escape(w) + r"(?![a-z])" for w in HEAVY),
    re.I)
_LIGHT_RE = re.compile(
    "|".join(r"(?<![a-z])" + re.escape(w) + r"(?![a-z])" for w in LIGHT),
    re.I)


def weight_in_name(name):
    """글꼴 이름이 웨이트를 선언하고 있으면 'heavy' 또는 'light' 를 준다."""
    if any(k in name for k in HEAVY_KO) or _HEAVY_RE.search(name):
        return "heavy"
    if any(k in name for k in LIGHT_KO) or _LIGHT_RE.search(name):
        return "light"
    return None


def unbold(data, seen=None):
    """이름에 중간 이상 웨이트가 든 run 에서만 b 속성을 뺀다."""
    hits = [0]

    def rep(m):
        attrs, inner = m.group(1), m.group(2)
        names = [n.decode("utf-8", "replace") for n in TYPEFACE.findall(inner)]
        kinds = {weight_in_name(n) for n in names}
        if seen is not None:
            for n in names:
                k = weight_in_name(n)
                if k:
                    seen.setdefault(n, k)
        if "heavy" not in kinds:
            return m.group(0)
        a2 = BOLD.sub(b"", attrs)
        if a2 != attrs:
            hits[0] += 1
        return b"<a:rPr" + a2 + b">" + inner + b"</a:rPr>"

    return RPR.sub(rep, data), hits[0]


def slim(src, dst, seen=None):
    """임베드 글꼴·애니메이션·겹친 볼드를 뺀 사본을 dst 에 쓴다."""
    z = zipfile.ZipFile(src)
    n = 0
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as w:
        for i in z.infolist():
            if i.filename.endswith(".fntdata"):
                continue                                    # 임베드 글꼴 제거
            data = z.read(i.filename)
            if SLIDE.match(i.filename):
                data = TIMING.sub(b"", data)                # 애니메이션 제거
                data, k = unbold(data, seen)                # 겹친 볼드 제거
                n += k
            elif i.filename == "ppt/presentation.xml":
                data = EMBED.sub(b"", data)
            elif i.filename == "ppt/_rels/presentation.xml.rels":
                data = re.sub(rb"<Relationship[^>]*fonts/[^>]*/>", b"", data)
            elif i.filename == "[Content_Types].xml":
                data = re.sub(rb'<Default Extension="fntdata"[^>]*/>', b"", data)
            zi = zipfile.ZipInfo(i.filename, date_time=i.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            w.writestr(zi, data)
    z.close()
    return n


def targets(path):
    if os.path.isfile(path):
        return os.path.dirname(path) or ".", [os.path.basename(path)]
    names = [f for f in sorted(os.listdir(path))
             if f.endswith(".pptx") and not f.startswith("~$")
             and "피드백" not in f and "자동 저장" not in f]
    return path, names


def run(path, seen):
    folder, names = targets(path)
    if not names:
        print("  pptx 가 없다")
        return
    out = os.path.join(folder, "캔바업로드용")
    os.makedirs(out, exist_ok=True)
    for f in names:
        src = os.path.join(folder, f)
        dst = os.path.join(out, f[:-5] + "_캔바용.pptx")
        n = slim(src, dst, seen)
        bad = zipfile.ZipFile(dst).testzip()
        print("  %-44s %5.1fMB → %4.1fMB  볼드 정리 %3d곳  %s" % (
            f[:44], os.path.getsize(src) / 1e6, os.path.getsize(dst) / 1e6, n,
            "정상" if bad is None else "★ 손상 " + bad))
    print("  → %s" % out)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    seen = {}
    for t in sys.argv[1:]:
        print("\n### %s" % t)
        run(t, seen)
    if seen:
        print("\n이름에 웨이트가 든 글꼴로 잡은 것 (눈으로 확인한다)")
        for name, kind in sorted(seen.items()):
            print("  %-34s %s" % (name, "b 제거함" if kind == "heavy" else "그대로 둠"))
