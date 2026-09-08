# -*- coding: utf-8 -*-
"""캔바에 올리기 전에 pptx 를 진단한다. 파일은 읽기만 한다.

    python src/check_pptx.py "파일.pptx" ["파일2.pptx" ...]
    python src/check_pptx.py "폴더"

찍어 주는 것
  - zip CRC 상태 (캔바는 CRC 가 깨진 파일을 거부한다. 파워포인트는 그냥 연다)
  - 임베드 글꼴 (PC 에 설치 안 돼 있어도 파워포인트는 이걸로 그린다)
  - 쓰인 글꼴 이름과 run 수
  - 이름에 웨이트가 있으면서 b 가 겹친 run 수 (캔바에서 한 단계 가늘어지는 곳)
  - 애니메이션이 있는 슬라이드 수
"""
import os
import re
import sys
import zipfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from for_canva import RPR, BOLD, TYPEFACE, SLIDE, weight_in_name  # noqa: E402

EMBED_NAME = re.compile(rb'<p:embeddedFont>\s*<p:font typeface="([^"]*)"')


def check(path):
    print("\n### %s" % os.path.basename(path))
    print("  크기 %.1fMB" % (os.path.getsize(path) / 1e6))

    z = zipfile.ZipFile(path)
    names = z.namelist()

    bad = z.testzip()
    if bad is None:
        print("  zip 검사  정상")
    else:
        info = z.getinfo(bad)
        print("  zip 검사  ★ 손상 %s  (%d바이트)" % (bad, info.file_size))
        print("            캔바가 거부한다. fix_crc.py 로 고친 뒤 올린다")

    slides = [n for n in names if SLIDE.match(n)]
    notes = [n for n in names if re.match(r"ppt/notesSlides/notesSlide\d+\.xml$", n)]
    print("  슬라이드 %d장, 발표자 노트 %d장" % (len(slides), len(notes)))

    fnt = [n for n in names if n.endswith(".fntdata")]
    pres = z.read("ppt/presentation.xml") if "ppt/presentation.xml" in names else b""
    embedded = [m.decode("utf-8", "replace") for m in EMBED_NAME.findall(pres)]
    if fnt:
        print("  임베드 글꼴 %d개 (%.1fMB)" % (
            len(fnt), sum(z.getinfo(n).file_size for n in fnt) / 1e6))
        for e in embedded:
            print("      %s" % e)
        print("    캔바는 이걸 읽지 않는다. 이름으로 자기 글꼴을 다시 맞춘다")
    else:
        print("  임베드 글꼴 없음")

    used = Counter()
    clash = Counter()
    anim = 0
    for s in sorted(slides):
        data = z.read(s)
        if b"<p:timing>" in data:
            anim += 1
        for m in RPR.finditer(data):
            attrs, inner = m.group(1), m.group(2)
            has_b = bool(BOLD.search(attrs))
            for raw in set(TYPEFACE.findall(inner)):
                name = raw.decode("utf-8", "replace")
                used[name] += 1
                if has_b and weight_in_name(name) == "heavy":
                    clash[name] += 1
    z.close()

    print("  애니메이션이 있는 슬라이드 %d장" % anim)
    print("  쓰인 글꼴")
    for name, c in used.most_common():
        w = weight_in_name(name)
        tag = {"heavy": "[이름에 굵은 웨이트]", "light": "[이름에 가는 웨이트]"}.get(w, "")
        print("      %-34s %4d곳 %s" % (name, c, tag))

    total = sum(clash.values())
    if total:
        print("  ★ 이름 웨이트와 b 가 겹친 곳 %d곳" % total)
        for name, c in clash.most_common():
            print("      %-34s %4d곳" % (name, c))
        print("    캔바에서 한 단계 가늘게 나온다. for_canva.py 로 사본을 만든다")
    else:
        print("  이름 웨이트와 b 가 겹친 곳 없음")


def main(args):
    files = []
    for a in args:
        if os.path.isdir(a):
            files += [os.path.join(a, f) for f in sorted(os.listdir(a))
                      if f.endswith(".pptx") and not f.startswith("~$")]
        else:
            files.append(a)
    if not files:
        print("pptx 가 없다")
        return 1
    for f in files:
        check(f)
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    raise SystemExit(main(sys.argv[1:]))
