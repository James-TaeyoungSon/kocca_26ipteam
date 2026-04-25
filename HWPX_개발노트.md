# HWPX 순수 Python 생성 — 개발 노트

> 한컴오피스 없이 GitHub Actions(Ubuntu) 환경에서 `.hwpx` 파일을 생성한 경험을 정리한 문서.  
> 앞으로 같은 실수를 반복하지 않도록 오류 원인과 해결책을 모두 기록.

---

## 1. HWPX 파일 구조

HWPX는 **ZIP 아카이브**다. 확장자를 `.zip`으로 바꾸면 내부를 볼 수 있다.

```
파일명.hwpx (= ZIP)
├── mimetype                    ← "application/hwp+zip" (압축 없이 저장)
├── META-INF/
│   ├── container.xml
│   └── container.rdf
├── Contents/
│   ├── content.hpf             ← 패키지 목차 (manifest)
│   ├── header.xml              ← 폰트·스타일·borderFill·charPr·paraPr 정의
│   └── section0.xml            ← 실제 문서 내용 (표·단락·텍스트)
├── Preview/
│   └── PrvText.txt
└── settings.xml
```

---

## 2. HWP 단위 계산

```
1mm = 283.46 HWP 단위
A4 세로 (portrait): width=59528, height=84186
A4 가로 (landscape, NARROWLY 모드): width=59528, height=84186 (물리 값은 동일)
  → landscape 렌더링 시 장축(84186≈297mm)이 시각적 가로가 됨
  → 사용 가능 가로 폭 = 84186 - (좌여백 + 우여백)

여백 환산:
  10mm = 2835 HWP
  15mm = 4251 HWP
  20mm = 5669 HWP
```

**핵심**: `sum(모든 열 너비) == 사용 가능 폭(_USABLE)` 이어야 한다.  
맞지 않으면 마지막 열이 넘치거나 빈 공간이 생긴다.

---

## 3. header.xml — 주요 ID 체계

### 3-1. borderFill

```xml
<hh:borderFills itemCnt="N">
  <hh:borderFill id="1" ...>...</hh:borderFill>
  ...
</hh:borderFills>
```

- `id`는 **1부터** 시작 (0 없음)
- `itemCnt`는 실제 `<hh:borderFill>` 개수와 **정확히 일치**해야 함
- 셀 XML의 `borderFillIDRef="N"` 이 여기 id를 참조

### 3-2. charPr (문자 속성)

```xml
<hh:charProperties itemCnt="N">
  <hh:charPr id="0" height="1000" textColor="#000000" ...>
    <hh:fontRef hangul="0" latin="0" .../>
    ...
  </hh:charPr>
  ...
</hh:charProperties>
```

- `id`는 **0부터** 시작
- `itemCnt`와 실제 개수 불일치 → 한글이 열 때 오류 또는 스타일 깨짐
- `fontRef hangul="N"` → HANGUL fontface의 N번 폰트 사용
- `height` 단위: 1pt = 100 (예: 10pt = 1000, 11pt = 1100)

### 3-3. paraPr (단락 속성)

- `id="0"`이 기본 정렬 (JUSTIFY)
- 가운데 정렬용 paraPr을 별도 id로 추가해야 함
- 역시 `itemCnt` 정확히 맞출 것

### 3-4. 폰트 등록

```python
# HANGUL, LATIN fontface에 폰트 추가 후 id로 참조
# 폰트 이름은 정확히 (공백 포함): '맑은 고딕', '함초롬돋움' 등
# GitHub Actions(Ubuntu)에는 폰트 미설치 → XML에 이름만 기록,
# 실제 렌더링은 사용자 PC에서 한글이 열 때 이루어짐
```

---

## 4. section0.xml — 표(Table) 구조

### 4-1. 기본 골격

```xml
<hp:tbl rowCnt="R" colCnt="C" ...>
  <hp:sz width="총너비" .../>
  <hp:tr>
    <hp:tc borderFillIDRef="3">
      <hp:subList>
        <hp:p><hp:run charPrIDRef="9"><hp:t>텍스트</hp:t></hp:run></hp:p>
      </hp:subList>
      <hp:cellAddr colAddr="0" rowAddr="0"/>
      <hp:cellSpan colSpan="1" rowSpan="1"/>
      <hp:cellSz width="W" height="H"/>
      <hp:cellMargin left="510" right="510" top="141" bottom="141"/>
    </hp:tc>
  </hp:tr>
</hp:tbl>
```

### 4-2. ❌ rowspan 높이 버그 (매우 중요)

**잘못된 코드:**
```python
# rowspan=3인 셀에 height * rowspan 적용 → 절대 금지
f'<hp:cellSz width="{w}" height="{500 * 3}"/>'  # ❌
```

**올바른 코드:**
```python
# rowspan과 무관하게 단일 행 높이만 지정
f'<hp:cellSz width="{w}" height="{500}"/>'       # ✅
```

**원인**: HWP는 `cellSz height`를 해당 셀이 스패닝하는 **첫 번째 행의 최소 높이**로 해석한다.  
`height * rowspan`을 주면 첫 행이 그 높이를 전부 차지해버려서 이후 행(스패닝 대상)이 0높이로 렌더링되어 **사라진다**.

**rowspan 병합 시 covered 셀 처리:**
```python
# 병합된 상위 셀이 차지하는 칸은 <hp:tr>에서 아예 생략
# (rowspan=3이면 다음 2개 행의 해당 열 <hp:tc>를 쓰지 않음)
```

### 4-3. 한 단락 안에 다른 스타일 텍스트 (인라인 run)

```xml
<!-- 단락 하나에 run 여러 개 → 서로 다른 색/볼드 인라인 -->
<hp:p paraPrIDRef="0">
  <hp:run charPrIDRef="9"><hp:t>일반 내용 텍스트</hp:t></hp:run>
  <hp:run charPrIDRef="19"><hp:t> (마감일)</hp:t></hp:run>
</hp:p>
```

별도 줄이 아닌 **같은 줄 끝에** 다른 스타일 텍스트를 붙이려면  
같은 `<hp:p>` 안에 `<hp:run>`을 이어서 추가하면 된다.

---

## 5. Make.com → GitHub Actions 데이터 흐름

```
Notion 페이지 기간 설정
  → Make.com 트리거
  → Google Calendar API 조회 (이벤트 수집)
  → CSV 문자열 생성
  → GitHub repository_dispatch (client_payload에 csv_text 또는 csv_b64 포함)
  → GitHub Actions 실행 (generate_and_upload_hwp_py.py)
  → HWPX 생성 → Notion 페이지에 파일 업로드
```

**Python에서 payload 읽기:**
```python
event_path = os.getenv('GITHUB_EVENT_PATH')
with open(event_path, 'r', encoding='utf-8') as f:
    event = json.load(f)
payload = event['client_payload']
csv_text = payload.get('csv_text', '')
csv_b64  = payload.get('csv_b64', '')
```

---

## 6. ❌ Google Calendar timeMax 배타적 범위 버그

**증상**: 주간 보고 기간 마지막 날(예: 금요일 4/24)의 일정이 보고서에 누락됨

**원인**: Google Calendar API `timeMax` 파라미터는 **exclusive** (미만)이다.  
Make.com이 Notion 기간 End(`2026-04-24`)를 그대로 timeMax로 넘기면  
4/24 당일에 시작하는 이벤트는 `timeMax` 이후로 판단되어 반환되지 않는다.

**Make.com 수정 (Google Calendar 모듈 End Date 필드):**
```
변경 전: 기간: End
변경 후: {{addDays(기간.End; 1)}}
```

`timeMax = 4/25` 로 설정하면 4/24 이벤트가 정상 포함된다.

> **주의**: Python 코드에서는 이 문제를 해결할 수 없다.  
> 데이터가 CSV payload에 포함되지 않으면 Python이 만들어낼 방법이 없다.

---

## 7. ❌ Google Calendar 종일 일정 날짜 +1일 표시 버그

**증상**: 3/25 종일 일정이 3/26으로 표시됨

**원인**: Google Calendar는 종일 일정의 종료일을 **실제 종료일 +1일 00:00:00**으로 저장  
(exclusive end date 관행). Make.com이 이 값을 그대로 CSV에 담아 전송.

**Python 수정**: 시작/종료 시간이 모두 `00:00`이면 종일 일정으로 판단, 종료일에서 -1일
```python
if sh == '00' and sm == '00' and eh == '00' and em == '00':
    end = end - timedelta(days=1)   # Google Calendar exclusive end 보정
```

이 수정은 Make.com을 건드리지 않고 **Python에서만** 적용 가능한 유형이다  
(데이터는 왔지만 날짜 표기가 틀린 경우).

---

## 8. 열 너비 계산 팁

```python
# 한글 문자 너비 추정 (HWP 단위)
# 10pt 기준: 한글 ≈ 1050, 영문 ≈ 580
# 11pt bold: 한글 ≈ 1200 (캐릭터 같이 3자짜리가 잘리지 않으려면)
def _tw(text, cjk_w=1050, ascii_w=580):
    return sum(cjk_w if ord(c) > 127 else ascii_w for c in text)

# 열 너비 = 헤더+데이터 최대 텍스트 너비 + 좌우 셀 여백(1020)
def col_width(col_idx, header_name, data_rows, cjk_w=1050):
    all_lines = [header_name]
    for row in data_rows:
        all_lines.extend(row[col_idx].split('\n'))
    return max(_tw(l, cjk_w) for l in all_lines) + 1020

# 마지막 열이 남은 공간을 흡수 → sum == _USABLE 보장
w_last = _USABLE - sum(other_widths)
```

---

## 9. 진단용 로그 추가 (GitHub Actions)

CSV 데이터 문제가 의심될 때 `main()` 함수에 임시 로그 추가:

```python
lines = csv_text.splitlines()
print(f'[CSV] 수신 행 수: {len(lines)}행')
for i, ln in enumerate(lines[:20]):
    print(f'[CSV] {i:02d}: {ln}')
```

Actions 탭 → 해당 실행 → "Run HWPX generator" 스텝에서 확인.

---

## 10. 로컬 개발 환경 설정

```bash
# 의존성 (Python 순수, COM/한컴 불필요)
pip install requests openpyxl

# 로컬 테스트 스크립트
python test_hwpx_local.py       # xlsx → hwpx 변환 후 저장
python debug_hwpx.py            # _parse_iljeong 결과·캘린더 이벤트 상세 출력
```

**로컬에서 작업 후 GitHub Actions에 반영하는 순서:**
1. `test_hwpx_local.py` 로 HWPX 생성 → 한글에서 열어 시각 확인
2. `debug_hwpx.py` 로 날짜 파싱·이벤트 배치 수치 확인
3. `git commit && git push` → Actions 자동 반영 (워크플로우 파일 수정 불필요)

---

## 11. 체크리스트 (신규 작업 시)

- [ ] `itemCnt` 값이 실제 요소 개수와 일치하는가?
- [ ] `sum(col_widths) == _USABLE` 인가?
- [ ] rowspan 셀의 `cellSz height`에 rowspan을 곱하지 않았는가?
- [ ] 새 charPr ID 추가 시 `itemCnt` 도 올렸는가?
- [ ] Make.com Google Calendar End Date가 `addDays(기간.End, 1)` 인가?
- [ ] 종일 일정 `00:00~00:00` 패턴 보정이 적용되어 있는가?
- [ ] 폰트 이름 오탈자 없는가? ('맑은 고딕' — 공백 포함)
