---
name: web-ui-design
description: 이 저장소(및 maechorom31-pixel의 교사용 웹 도구)의 UI 디자인 표준. 새 웹 도구·페이지를 만들거나 기존 인터페이스를 수정·개선할 때, 사용자가 "디자인", "인터페이스", "UI", "예쁘게", "시인성", "보기 좋게"를 언급하면 명시 요청이 없어도 먼저 읽는다. 색상·타이포그래피·다크 모드·카테고리 색 구분 규칙을 정의한다.
---

# 웹 도구 UI 디자인 표준

전남 지역 고등학교 교사(국어)와 학생·학부모가 쓰는 웹 도구 모음의 디자인 표준.
사용 환경: 학교 PC(구형 브라우저 가능성), 교사·학생의 휴대폰. 최우선 가치는 **시인성(읽기 쉬움)**, 그 다음이 심미성.

## 핵심 원칙 (우선순위 순)

1. **고대비**: 밝은 중립 배경 + 흰 카드 + 진한 잉크색 글자. 저채도 배경에 갈색·회색 글자처럼 명도 차가 작은 조합 금지.
2. **포인트 색은 하나**: 행동 유도(주 버튼, 링크, 포커스, 검색 강조)는 파랑 하나로 절제. 색이 많은 곳 = 정보 구분, 파랑 = 누르는 곳.
3. **카테고리 색 구분**: 목록·그룹이 있으면 그룹마다 고유 색(아이콘 타일 배경 + 섹션 제목 점)을 순환 배정해 훑어볼 때 색으로 인지되게 한다.
4. **현재 상태는 확실하게**: 선택된 탭·모드는 진한 잉크색 알약(solid pill)로 채워서 표시. 옅은 배경색 변화만으로 상태를 표시하지 않는다.
5. **다크 모드는 별도 조정**: 단순 반전 금지. 어두운 배경용 색을 따로 정의한다. 테마는 자동/라이트/다크 3단 순환 토글, localStorage에 저장.
6. **접근성 기본기**: `:focus-visible` 아웃라인, `prefers-reduced-motion` 존중, aria-label(아이콘 버튼), 모바일 1열 레이아웃, 터치 타깃 최소 38px.

## 디자인 토큰

```css
:root {
  --bg: #F4F5F7;          /* 밝은 중립 회색 배경 */
  --card: #FFFFFF;
  --ink: #171A20;          /* 본문 글자 */
  --ink-soft: #5A616C;     /* 보조 글자 */
  --line: #E3E5EA;         /* 테두리 */
  --accent: #3355D8;       /* 유일한 포인트 파랑 */
  --accent-strong: #2543B8;
  --accent-soft: #EBF0FF;  /* 포인트 옅은 배경 (호버·강조) */
  --star: #F59F00;         /* 즐겨찾기 금색 */
  --danger: #D6336C;
  --shadow: 0 1px 2px rgba(23,26,32,.05), 0 4px 14px rgba(23,26,32,.06);
  --shadow-lift: 0 2px 6px rgba(23,26,32,.07), 0 12px 28px rgba(23,26,32,.11);
}
/* 다크 (자동 감지 + 수동 토글 모두 지원해야 함) */
/* @media (prefers-color-scheme: dark) { :root:not([data-theme="light"]) {...} }
   :root[data-theme="dark"] {...} 두 곳에 같은 값을 정의 */
--bg: #14161B;  --card: #1D2027;  --ink: #EAECF1;  --ink-soft: #9BA2AE;
--line: #2D313B;  --accent: #8298FF;  --accent-strong: #A5B4FF;
--accent-soft: #242C4E;  --star: #FFC94D;  --danger: #FF8CB0;
```

### 카테고리 색 (6색 순환, 배경/전경 쌍)

| | 라이트 bg/fg | 다크 bg/fg |
|---|---|---|
| 0 파랑 | `#E8EEFF` / `#3B5BDB` | `#26305A` / `#91A7FF` |
| 1 보라 | `#F3EBFF` / `#7048E8` | `#2F2455` / `#B197FC` |
| 2 청록 | `#DFF4F0` / `#099268` | `#1B3A33` / `#63E6BE` |
| 3 주황 | `#FFF0DC` / `#E8590C` | `#46301A` / `#FFA94D` |
| 4 분홍 | `#FFE9F0` / `#D6336C` | `#47222F` / `#FAA2C1` |
| 5 회색 | `#EDEFF3` / `#4B5563` | `#2A2F3A` / `#ADB5BD` |

`.tint-N` 클래스가 `--tb`(배경)·`--tf`(전경) 변수를 설정하고, 컴포넌트는
`background: var(--tb, var(--accent-soft))`처럼 폴백과 함께 소비한다.
카테고리 인덱스 `% 6`으로 배정하면 사용자가 카테고리를 추가해도 자동으로 색이 붙는다.

## 타이포그래피

- 글꼴: `"Pretendard", "Noto Sans KR", "Apple SD Gothic Neo", "Malgun Gothic", sans-serif` — 웹폰트 로딩 금지(학교망 고려), 시스템 스택만.
- 페이지 제목 1.5rem/800, 카드 제목 1rem/700(자간 -0.015em), 본문·설명 0.84rem(행간 1.55), 보조 0.78rem.
- 숫자가 정렬되는 곳(개수, URL, 점수)은 `font-variant-numeric: tabular-nums`.
- 한 줄에 글자가 너무 길어지지 않게 콘텐츠 최대 폭 960~1020px.

## 컴포넌트 규칙

- **카드**: 흰 배경, 1px `--line` 테두리, radius 14px, 은은한 그림자. 호버 시 `translateY(-2px)` + 테두리 `--accent` + `--shadow-lift`.
- **버튼**: 기본은 카드색+테두리, 주 행동만 `--accent` 채움. radius 10px.
- **탭/필터**: 알약(radius 999px). 활성 = `--ink` 채움 + `--bg` 글자.
- **입력창**: 포커스 시 `border-color: --accent` + `box-shadow: 0 0 0 3px var(--accent-soft)` 링.
- **모달**: 배경 `rgba(15,18,28,.55)`, radius 16px, Esc·바깥 클릭으로 닫기.
- 브라우저 기본 UI 중복 제거: `input[type=search]::-webkit-search-cancel-button` 숨김 등.

## 작업 방식

- 저장소는 GitHub Pages로 배포되는 순수 정적 HTML(단일 파일, 빌드 없음). 프레임워크·CDN 의존성을 추가하지 않는다.
- localStorage 키는 기존 것을 유지해 사용자 데이터를 보존한다 (예: 메인 페이지 `linkhub-data-v3`, 테마 `linkhub-theme`).
- 수정 후에는 Playwright(Chromium)로 라이트/다크/모바일 화면을 실제 렌더링해 확인하고, JS 오류가 없는지 검사한 뒤 커밋한다.
- 기존 참조 구현: 루트 `index.html`(이 표준의 원본), `image/index.html`(웜 스톤+앰버 — 추후 이 표준으로 통일 가능).
