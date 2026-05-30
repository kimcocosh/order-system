# 옥수주조 주문 관리 시스템

업체용 주문 폼 + 관리자 대시보드 + Google Sheets 연동 시스템입니다.

---

## 📁 파일 구조

```
order-system/
  order.html   ← 업체용 주문 폼 (GitHub Pages 배포)
  admin.html   ← 관리자 대시보드 (GitHub Pages 배포)
  code.gs      ← Google Apps Script (구글 시트에 붙여넣기)
  README.md    ← 이 파일
```

---

## 🗂️ Step 1 — 구글 시트에 시트 4개 만들기

구글 시트 URL: `https://docs.google.com/spreadsheets/d/1TFW2-hbXTNR_AljzO68__eXlxiwqPdoJ6QEhjYSuzXo`

### 방법 A: 자동 생성 (권장)

Apps Script 배포 후 `initSheets` 함수를 한 번 실행하면 4개 시트가 자동으로 생성됩니다. (Step 2 이후 진행)

### 방법 B: 수동 생성

시트 하단 `+` 버튼으로 아래 4개를 순서대로 추가하세요.

| 시트명 | 컬럼 |
|--------|------|
| 주문접수 | 타임스탬프 / 업체명 / 전화번호 / 주소 / 내품명 / 요청사항 / 처리상태 |
| 재고 | 품목명 / 현재재고 / 마지막업데이트 |
| 알림수신자 | 이메일 |
| 상품목록 | 품목명 / 사용여부 |

**상품목록 초기 데이터** (사용여부 전부 `Y`):
```
오리지널(달콤), 새콤, 옥수수, 고구마, 바나나, 망고,
패션후르츠, 블러드오렌지, 말차, 토마토, 블루베리,
라이트&드라이, 옥수수7도, 생쌀9도
```

---

## ⚙️ Step 2 — Apps Script 배포

1. 구글 시트 상단 메뉴 **확장 프로그램 → Apps Script** 클릭
2. 기존 코드 전체 삭제 후 `code.gs` 내용 붙여넣기
3. 상단 **💾 저장** (Ctrl+S)
4. **함수 선택 드롭다운**에서 `initSheets` 선택 → ▶️ **실행**
   - Google 계정 권한 허용 팝업이 뜨면 승인
   - 실행 완료 후 구글 시트에 시트 4개가 생성됩니다
5. 다시 **배포 → 새 배포** 클릭
6. 유형: **웹 앱** 선택
7. 설정:
   - 설명: `옥수주조 주문 시스템`
   - 다음 사용자로 실행: **나**
   - 액세스 권한: **모든 사용자**
8. **배포** 클릭 → 웹 앱 URL 복사 (아래에서 사용)

> ⚠️ 코드 수정 후에는 반드시 **배포 → 기존 배포 관리 → 편집(연필 아이콘) → 버전: 새 버전** 으로 재배포해야 반영됩니다.

---

## 🔗 Step 3 — HTML 파일에 웹앱 URL 입력

`order.html`과 `admin.html` 각각 열어서 아래 부분을 찾아 URL 교체:

```javascript
const WEBAPP_URL = 'YOUR_WEBAPP_URL_HERE'; // ← 여기에 붙여넣기
```

예시:
```javascript
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
```

두 파일 모두 동일한 URL을 입력합니다.

---

## 🌐 Step 4 — GitHub Pages 배포

### 저장소 생성 및 업로드

1. [github.com](https://github.com) 로그인 → **New repository**
2. Repository name: `order-system`
3. Public 선택 → **Create repository**
4. `order.html`, `admin.html`, `README.md` 세 파일 업로드
   - (code.gs는 구글 시트용이므로 GitHub에 올릴 필요 없음)
5. **Settings → Pages → Source: main branch / root** 선택 → **Save**
6. 수 분 후 아래 URL로 접속 가능:

```
업체용:   https://[계정명].github.io/order-system/order.html?key=고유코드
관리자:   https://[계정명].github.io/order-system/admin.html
```

---

## 🔑 Step 5 — 업체별 고유 링크 생성

고유 key는 **업체명 기반 자동 생성**입니다. 별도 설정 없이 아래 방법으로 확인합니다.

### key 확인 방법

Apps Script 에디터에서 아래 함수를 임시 실행:

```javascript
function printKeys() {
  const ss = SpreadsheetApp.openById('1TFW2-hbXTNR_AljzO68__eXlxiwqPdoJ6QEhjYSuzXo');
  const data = ss.getSheets()[0].getDataRange().getValues();
  // 거래처 목록 시트에서 업체명 추출 후 key 생성
  // 실행 로그에서 확인
  const clients = getClients(ss); // JSON 결과에서 key 확인
  Logger.log(JSON.stringify(clients));
}
```

또는 브라우저에서 직접 확인:
```
https://script.google.com/macros/s/[배포ID]/exec?action=clients
```
→ JSON 결과의 `key` 값을 복사해서 링크 생성

### 업체 링크 형식
```
https://[계정명].github.io/order-system/order.html?key=[key값]
```

---

## 📧 Step 6 — 알림 수신자 관리

구글 시트 **알림수신자** 탭에서 직접 관리합니다.

| 이메일 |
|--------|
| manager@example.com |
| owner@example.com |

- **추가**: 새 행에 이메일 입력
- **삭제**: 해당 행 삭제
- 즉시 반영 (재배포 불필요)

---

## 🏷️ Step 7 — 상품 추가/비활성화

### 관리자 대시보드에서 (권장)
1. `admin.html` 접속 → **상품 관리** 탭
2. 기존 품목 `Y/N` 토글로 표시 여부 변경
3. 하단 입력란에서 새 품목 추가
4. **저장** 클릭

### 구글 시트에서 직접
**상품목록** 시트에서 사용여부 컬럼 값을 `Y` / `N` 으로 변경

---

## 🔒 관리자 비밀번호 변경

`admin.html` 내 아래 줄 수정:
```javascript
const PASSWORD = 'admin1234'; // ← 원하는 비밀번호로 변경
```

---

## 📱 모바일 홈 화면 추가 (업체용)

### iPhone (Safari)
1. 업체 링크 접속
2. 하단 **공유** 아이콘 → **홈 화면에 추가**

### Android (Chrome)
1. 업체 링크 접속
2. 우측 상단 메뉴 → **앱 설치** 또는 **홈 화면에 추가**

---

## ❓ 자주 묻는 질문

**Q. 주문 후 이메일이 오지 않아요**
- Apps Script에서 GmailApp 권한을 허용했는지 확인
- 알림수신자 시트에 이메일이 올바르게 입력되어 있는지 확인
- 스팸함 확인

**Q. 재고가 자동 차감이 안 돼요**
- 재고 시트의 품목명이 상품목록 시트와 정확히 일치하는지 확인 (공백, 특수문자 주의)

**Q. 업체가 링크 접속 시 "유효하지 않은 링크" 오류**
- key 값이 올바른지 확인
- 거래처 DB 시트의 업체명과 링크의 key가 매칭되는지 `/exec?action=clients` 로 확인

**Q. 코드 수정 후 반영이 안 돼요**
- Apps Script 재배포 필요: 배포 → 기존 배포 관리 → 버전: 새 버전 선택
