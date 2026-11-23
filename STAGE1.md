# 올리브영 1단계 자동 실행 (화·금 15:10, 예약 실행용)

대표님이 정기 작업 지시문에서 이 절차의 **저장·입고예약·PDF 생성을 미리 승인**했다. 아래 순서와 조건 그대로만 한다. 조건에서 벗어나면 그 자리에서 멈추고 알림을 보낸다. 모든 보고는 한글로 쉽게.

금지: 비밀번호 입력, 예외등록(G코드), 유통기한 입력, 2단계(실재고 반영) 작업, 이미 숫자가 있는 납품예정수량 변경, 예약 취소.

## 0. 납품일
화요일 실행 → 이번 주 목요일 / 금요일 실행 → 다음 주 화요일 (한국시간). 파일명용 요일: 화/목.

## 1. 올리브원 (Claude 앱 내장 브라우저, mcp__remote-devices__Claude_Browser__*)
1. https://oypartner.cj.net/CJOYP/nexacro/main.fo 로 이동, 6초 대기, 스크린샷.
   - 로그인 화면이면 **입력하지 말고** PushNotification: "올리브원 로그인이 풀려 있어요. 회사 PC Claude 앱 브라우저에서 로그인해 주세요 (화·금 15:10 자동 1단계 대기 중)" → 종료.
2. 메뉴 검색칸(스크린샷에서 왼쪽 위 검색 입력칸) 클릭 → `납품확인서` 입력 → Down → 하이라이트가 `[80120]납품확인서(R1)`인지 확인 → Enter.
3. JS로 폼 찾기 (`window.application` 부터 탐색, `ds_supVeridcLst` 가진 폼을 `window.__oyForm`에 저장). `div_inqCond.cal_stkinStrtYmd`, `cal_stkinEndYmd` 를 납품일 YYYYMMDD로 set_value. 화면의 [조회] 버튼 클릭(스크린샷으로 위치 확인), 3초 대기.
4. 행 0개 → "아직 발주 없음" 알림 후 종료. 모든 행 stkinYmd == 납품일 인지 확인.
5. `supTobeQty`가 비었거나(undefined/null/'') 0인 행만 `ds.setColumn(r,'supTobeQty', poQty)`. 숫자 있는 행은 유지. 화면에 경고창(Alert)이 뜨면 멈추고 알림.
6. [저장] 클릭 → "저장하시겠습니까?" [확인] → "저장되었습니다." [확인]. 다른 메시지면 멈추고 알림.
7. 다시 조회 → 행 수·합계 확인 (모든 행 supTobeQty 값 있음).

## 2. 입고예약 (양지센터 LA02만, 양지On은 안 함)
1. [입고예약등록] 클릭 → 새 탭 "입고예약등록(R1)". JS로 `ds_stkinResvLst` 가진 폼을 찾아 `fn_search()` 호출, 3초 대기.
2. 이미 `stkinResvYn=='R'`(예약완료) 행이 있으면 예약하지 않는다 (재예약 금지). 그 시간과 SKU/수량을 보고에 적는다.
3. 없으면 `14:30~15:00` 행이 `Y`(가능)일 때 그 행의 입고예약 칸(가로 스크롤 오른쪽 끝, "가능")을 클릭 → "입고 예약을 하시겠습니까?" [확인] → "입고 예약이 완료되었습니다." [확인]. 불가면 바로 위 시간대(14:00~14:30) 한 번만 시도. 둘 다 불가면 멈추고 알림.
4. 재조회해서 '예약완료' 확인.

## 3. 실재고파악용 PDF
1. 납품확인서 탭으로 돌아가 [조회], 헤더 체크박스로 전체 체크.
2. [인쇄] 클릭 → "납품예정수량을 입력하지 않은 데이터는 인쇄 할 수 없습니다" 창은 매번 뜨는 확인창이라 [확인]. 리포트 뷰어가 뜨면, 뷰어 iframe(`nexacro_rdviewer.html`)의 XMLHttpRequest.send 를 가로채는 대신 아래 방식으로 PDF를 만든다:
   - 뷰어의 저장(디스크) ▼ → PDF 를 누르기 **전에** iframe `XMLHttpRequest.prototype.send`를 감싸 `opcode=500` 요청 body를 `window.__rdBody`에 저장하고, 요청은 막는다(실제 전송 안 함 → 내장 브라우저 저장 창 방지).
3. Claude in Chrome(mcp__claude-in-chrome__*)에서 새 탭으로 `https://oypartnerrd.cj.net/ReportingServer/html5/js/jquery-1.11.0.min.js` 열고, JS로 `/ReportingServer/service`에 그 body POST → 응답 `1|경로` → `<a href="/ReportingServer/download?filename=경로&delete=true&attatchment=true">` 클릭. 크롬은 저장 창 없이 다운로드 폴더에 저장한다. 탭 닫기.
4. `C:\Users\gkint\Downloads`에서 방금 생긴 `30-*.pdf` 또는 `40-*.pdf`(최신 mtime) 를 device_stage_files → `/mnt/user-data/outputs/`로 복사 → pdftotext로 상품코드 수가 행 수와 같은지 확인 → device_commit_files 로
   `C:\Users\gkint\정산raw\올리브영납품입고예약자동화\납품확인서\YYYYMMDD(요일) 납품확인서_실재고파악용.pdf` 저장. 같은 이름이 이미 있으면 덮어쓰지 말고 `_2` 붙이기.
   - 크롬 연결이 안 되면 PDF만 건너뛰고 알림에 "PDF는 직접 저장 필요"라고 적는다.

## 4. 알림 (항상 1번, 짧게)
"11.24(화) 1단계 완료: 30품목 1,754개 저장 · 양지 14:30 예약(28품목 1,694개) · PDF 폴더 저장 — 3PL에 보내주세요"
문제가 있으면 첫 문장에 문제를 쓴다.
