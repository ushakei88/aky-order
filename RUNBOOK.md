# 올리브영 발주 확인 정기 작업 (예약 실행용 지시문)

너는 아케이(협력사코드 C20481)의 올리브영 발주를 확인해서 대시보드를 갱신한다. 올리브원에서는 **조회만** 한다. 저장, 입고예약, 예외등록, 인쇄는 절대 하지 않는다. 비밀번호는 절대 입력하지 않는다.

## 1. 납품일 정하기 (한국시간 기준)
- 화요일 실행 → 이번 주 목요일 납품
- 금요일 실행 → 다음 주 화요일 납품
- 월요일 실행 → 다음 날(화) 납품 / 수요일 실행 → 다음 날(목) 납품

## 2. 올리브원에서 발주 가져오기 (회사 PC의 Claude 앱 내장 브라우저, mcp__remote-devices__Claude_Browser__* 도구)
1. tabs_context로 열린 탭을 보고, https://oypartner.cj.net/CJOYP/nexacro/main.fo 로 이동한다.
2. 로그인 화면이면 **입력하지 말고 멈춘다.** PushNotification(또는 SendUserMessage)으로 "올리브원 로그인이 풀려 있어요. 회사 PC Claude 앱 브라우저에서 로그인해 주세요"라고 알리고 종료한다.
3. (좌표는 스크린샷 800x450 기준) 왼쪽 메뉴 검색칸(약 42,84)에 `납품확인서` 입력 → 아래 화살표 → Enter 로 "[80120]납품확인서(R1)"을 연다. 왼쪽 메뉴는 접지 않아도 된다(납품일자 칸이 이미 보임).
4. 시작(약 225,127)·종료(약 310,127) 납품일자 둘 다 납품일(YYYYMMDD)로 바꾼다(칸 클릭 → ctrl+a → 입력).
5. 오른쪽 위 [조회](약 770,104)를 누른다. 화면이 다르면 스크린샷으로 위치를 다시 찾는다. 3초 기다린다.
6. javascript_tool로 데이터셋을 뽑는다:
```js
(() => {
  const app = window.application || (nexacro.getApplication && nexacro.getApplication()); const seen = new Set(); let f = null;
  function w(o,d){ if(!o||d>9||seen.has(o)||f) return; seen.add(o);
    try{ if(o.ds_supVeridcLst&&o.ds_supVeridcLst.getRowCount){ f=o.ds_supVeridcLst; return; } }catch(e){}
    const k=[]; try{ if(o.all) for(let i=0;i<o.all.length;i++) k.push(o.all[i]); }catch(e){}
    try{ if(o.frames) for(let i=0;i<o.frames.length;i++) k.push(o.frames[i]); }catch(e){}
    for(const n of ['mainframe','form','frame','div_work','div_main','childframe']){ try{ if(o[n]) k.push(o[n]); }catch(e){} }
    k.forEach(x=>w(x,d+1)); }
  w(app,0); if(!f) return 'NOTFOUND';
  const cols=['stkinYmd','poYmd','stkinCntrCd','stkinCntrNm','gdsCd','gdsNm','inptUnitQty','inptUnitPerQty','poQty','supQty','mfgYmd','dstrTlmtYmd','rmkCont','oymaster','premBrndYn','cntrPoSlipNo','cntrPoSlipSeq','mdEmpNm','mdEmailAddr','costUprc','supCostAmt','supCostVatAmt','vatSpCd','vatRt','cnfmYn','stkinStatCd','stkinResvQty','boxPerQty'];
  const rows=[]; for(let r=0;r<f.getRowCount();r++) rows.push(cols.map(c=>f.getColumn(r,c)));
  return JSON.stringify({cols,rows});
})()
```
7. 결과 JSON 문자열을 작업 폴더에 `orders.json`으로 저장한다. 행이 0개면 "아직 발주 없음"으로 보고 종료한다.

## 3. 대시보드 갱신
1. add_repo(ushakei88/aky-order, push)로 저장소를 붙이고 clone 한다.
2. 셸에서 (DASH_PW는 예약 작업 지시문에 적힌 대시보드 비밀번호):
```bash
cd aky-order
DASH_PW="$DASH_PW" python3 vault.py enc ../orders.json inputs/orders.json.enc
DASH_PW="$DASH_PW" python3 vault.py dec inputs/stock.xlsx.enc /tmp/stock.xlsx
DASH_PW="$DASH_PW" python3 -W ignore build.py --orders ../orders.json --stock /tmp/stock.xlsx --stock-label "$(cat inputs/stock.label)" --out .
git add -A && git commit -m "발주 확인 $(TZ=Asia/Seoul date '+%m-%d %H:%M')" && git push
```
3. push가 거절되면(다른 자동 작업이 먼저 올림) `git fetch && git reset --hard origin/main` 후 위 명령을 처음부터 다시 실행한다.
4. build.py가 출력한 "변동: 신규 · 수량변경 · 삭제" 줄을 본다.

## 4. 알림 (모든 보고는 한글로 쉽게)
- 처음 확인이거나 변동(신규/수량변경/삭제)이 1건 이상이면 PushNotification으로 짧게 알린다: "11.24(화) 납품분: 부족 ○품목 ○카톤, 변동 ○건 — https://ushakei88.github.io/aky-order/"
- 변동이 없으면 알리지 않는다.
- GNC 브랜드는 전부 운영중지(이관) — build.py가 자동으로 운영중지로 표시하고 카톤 계산에서 뺀다.
