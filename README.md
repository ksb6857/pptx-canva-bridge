# pptx-canva-bridge

파워포인트로 만든 pptx 를 캔바에 올리면 같은 글꼴인데 굵기가 한 단계 가늘게 나온다.
이 문제를 잡는 Claude Code 스킬이다. 업로드 자체가 실패하는 CRC 손상도 함께 고친다.

원격연수 교안 5편(2권 2편, 3권 3편, 모두 네모고딕)을 캔바에 올리면서 원인을 잡고
만들었다. 실제로 쓴 수치와 확인 방법이 그대로 들어 있다.

```
녹화용 원본 pptx ─┬─▶ 그대로 파워포인트에서 녹화
                  │
                  └─▶ for_canva.py ─▶ 캔바용 사본 ─▶ 캔바 업로드
                       b 제거
                       임베드 글꼴 제거
                       애니메이션 제거
```

## 왜 두 벌인가

pptx 의 글자 조각(run)은 글꼴 이름과 굵기 플래그를 따로 갖는다.

```xml
<a:rPr sz="3600" b="true"><a:latin typeface="네모고딕 Heavy"/></a:rPr>
```

이름에 Heavy 가 있는데 `b="true"` 가 겹쳐 있다. 이 조합에서

- 캔바는 이름의 웨이트를 버리고 한 단계 아래("중간")로 잡는다
- 파워포인트는 `b` 를 빼면 굵게 그리지 않는다

한 파일로 둘 다 만족시킬 수 없다. 그래서 원본은 두고 캔바용 사본을 따로 만든다.

## 쓰는 법

파이썬 3.8 이상이면 된다. 외부 패키지가 필요 없다. 표준 라이브러리만 쓴다.

```bash
python src/check_pptx.py "교안.pptx"              # 진단
python src/for_canva.py  "교안이_있는_폴더"        # 캔바용 사본 만들기
python src/fix_crc.py    "깨진.pptx" "성한.pptx"   # 업로드가 거부될 때
```

`for_canva.py` 는 원본을 건드리지 않고 `대상폴더/캔바업로드용/` 에 사본을 만든다.

### 진단 결과 예시

```
### (2팀_11차시) ... - 10월.pptx
  크기 16.7MB
  zip 검사  정상
  슬라이드 37장, 발표자 노트 37장
  임베드 글꼴 7개 (10.0MB)
  ★ 이름 웨이트와 b 가 겹친 곳 198곳
      네모고딕 Medium                         158곳
      네모고딕 Heavy                           14곳
```

## 들어 있는 것

| 파일 | 하는 일 |
|---|---|
| `SKILL.md` | 판단 기준과 작업 순서. Claude 가 읽는다 |
| `src/check_pptx.py` | 읽기 전용 진단. CRC, 임베드 글꼴, 겹친 볼드 수 |
| `src/for_canva.py` | 캔바용 사본 생성 |
| `src/fix_crc.py` | 깨진 zip 멤버만 성한 파일에서 갈아 끼우기 |
| `docs/실측기록.md` | 어떻게 확인했는지, 숫자와 날짜 |

## 설치

Claude Code 스킬 폴더에 그대로 두면 된다.

```bash
git clone https://github.com/ksb6857/pptx-canva-bridge.git ~/.claude/skills/pptx-canva-bridge
```

## 라이선스

MIT
