# -*- coding: utf-8 -*-
"""CRC 가 깨진 pptx 를 고친다. 깨진 멤버만 성한 파일에서 가져와 갈아 끼운다.

파워포인트는 zip 멤버의 CRC 를 검사하지 않는다. 캔바는 검사하고 거부한다.
그래서 파워포인트에서 멀쩡히 열리는 파일이 캔바 업로드만 실패하는 일이 생긴다.

    python src/fix_crc.py "깨진.pptx"                    # 어디가 깨졌는지만 본다
    python src/fix_crc.py "깨진.pptx" "성한.pptx"         # 고쳐서 _복구.pptx 로 쓴다
    python src/fix_crc.py "깨진.pptx" "성한.pptx" "결과.pptx"

성한 파일은 자동 저장본이나 백업본이면 된다. 그 멤버(대개 그림 하나)만
같으면 슬라이드 수와 발표자 노트는 깨진 파일 쪽이 그대로 보존된다.

깨진 멤버를 뺀 나머지는 압축된 바이트를 그대로 옮긴다. 다시 압축하지 않는다.
"""
import os
import struct
import sys
import zipfile

LOC = 0x04034B50
CEN = 0x02014B50
EOCD = 0x06054B50


def broken_members(path):
    """CRC 가 맞지 않는 멤버 이름을 모두 찾는다."""
    bad = []
    with zipfile.ZipFile(path) as z:
        for n in z.namelist():
            if n.endswith("/"):
                continue
            try:
                z.read(n)
            except Exception:
                bad.append(n)
    return bad


def _raw(buf, zi):
    """로컬 헤더를 뜯어 압축된 바이트와 extra 필드를 그대로 준다."""
    off = zi.header_offset
    sig, _v, _f, _m, t, d, _c, _cs, _us, nl, el = struct.unpack_from(
        "<IHHHHHIIIHH", buf, off)
    if sig != LOC:
        raise ValueError("로컬 헤더가 아니다: %s" % zi.filename)
    extra = buf[off + 30 + nl: off + 30 + nl + el]
    start = off + 30 + nl + el
    return buf[start:start + zi.compress_size], extra, t, d


def repair(src, donor, out):
    bad = broken_members(src)
    if not bad:
        print("  깨진 멤버가 없다. 고칠 것이 없다")
        return 0

    print("  깨진 멤버 %d개" % len(bad))
    with zipfile.ZipFile(src) as z:
        for n in bad:
            i = z.getinfo(n)
            print("      %s  %d바이트" % (n, i.file_size))

    if donor is None:
        print("  성한 파일을 두 번째 인자로 주면 고친다")
        return 1

    with zipfile.ZipFile(donor) as dz:
        missing = [n for n in bad if n not in dz.namelist()]
        if missing:
            print("  ★ 성한 파일에 없는 멤버: %s" % ", ".join(missing))
            return 1
        for n in bad:
            dz.read(n)                       # 성한 쪽도 깨졌으면 여기서 터진다

    sbuf = open(src, "rb").read()
    dbuf = open(donor, "rb").read()
    sz = zipfile.ZipFile(src)
    dz = zipfile.ZipFile(donor)

    entries = []
    with open(out, "wb") as w:
        for zi in sz.infolist():
            take_from, buf = (dz, dbuf) if zi.filename in bad else (sz, sbuf)
            src_zi = take_from.getinfo(zi.filename)
            data, extra, t, d = _raw(buf, src_zi)
            if len(data) != src_zi.compress_size:
                raise ValueError("압축 바이트 길이가 안 맞는다: %s" % zi.filename)
            name = zi.filename.encode("utf-8") if (zi.flag_bits & 0x800) \
                else zi.orig_filename.encode("cp437", "replace")
            flag = src_zi.flag_bits & ~0x08          # 데이터 디스크립터는 안 쓴다
            offset = w.tell()
            if offset >= 0xFFFFFFFF:
                raise ValueError("4GB 를 넘는다. 이 스크립트로는 못 고친다")
            w.write(struct.pack("<IHHHHHIIIHH", LOC, src_zi.extract_version,
                                flag, src_zi.compress_type, t, d, src_zi.CRC,
                                src_zi.compress_size, src_zi.file_size,
                                len(name), len(extra)))
            w.write(name)
            w.write(extra)
            w.write(data)
            entries.append((src_zi, name, extra, flag, t, d, offset))

        cd = w.tell()
        for zi, name, extra, flag, t, d, offset in entries:
            comment = zi.comment or b""
            w.write(struct.pack("<IHHHHHHIIIHHHHHII", CEN,
                                (zi.create_system << 8) | zi.create_version,
                                zi.extract_version, flag, zi.compress_type,
                                t, d, zi.CRC, zi.compress_size, zi.file_size,
                                len(name), len(extra), len(comment),
                                0, zi.internal_attr, zi.external_attr, offset))
            w.write(name)
            w.write(extra)
            w.write(comment)
        size = w.tell() - cd
        w.write(struct.pack("<IHHHHIIH", EOCD, 0, 0, len(entries), len(entries),
                            size, cd, 0))

    sz.close()
    dz.close()

    left = broken_members(out)
    print("  결과 %s  %.1fMB  %s" % (
        os.path.basename(out), os.path.getsize(out) / 1e6,
        "정상" if not left else "★ 아직 깨진 멤버 " + ", ".join(left)))
    with zipfile.ZipFile(out) as z:
        n = sum(1 for x in z.namelist()
                if x.startswith("ppt/slides/slide") and x.endswith(".xml"))
        print("  슬라이드 %d장 보존" % n)
    return 0 if not left else 1


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(1)
    src = sys.argv[1]
    donor = sys.argv[2] if len(sys.argv) > 2 else None
    out = sys.argv[3] if len(sys.argv) > 3 else src[:-5] + "_복구.pptx"
    print("### %s" % os.path.basename(src))
    raise SystemExit(repair(src, donor, out))
