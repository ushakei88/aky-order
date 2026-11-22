# 올영 손익 수집 (예약 실행용 지시문)

올리브원 납품확인서(R1)를 **조회만** 해서 납품일별 데이터를 저장한다. 저장·입고예약·예외등록·인쇄는 절대 하지 않는다. 비밀번호는 절대 입력하지 않는다.

## 1. 조회할 납품일
한국시간 오늘 기준 **3일 전 ~ 내일** 사이의 화요일·목요일 납품일 전부 (예: 목요일 실행 → 월요일 지난 화·이번 목, 수요일 실행 → 지난 화·내일 목).

## 2. 납품일마다
1. 내장 브라우저에서 https://oypartner.cj.net/CJOYP/nexacro/main.fo 를 연다. 로그인 화면이면 멈추고 사용자에게 로그인해 달라고 알린다.
2. 왼쪽 메뉴 검색칸(약 80,152)에 `납품확인서` → 아래 화살표 → Enter 로 납품확인서(R1) 화면을 연다.
3. javascript_tool로 아래를 실행한다 (DATES에 조회할 납품일 YYYYMMDD 목록). 화면 버튼 대신 화면의 조회 함수를 그대로 부른다(조회만 함).
```js
(async () => {
  const DATES = ['20261117','20261119'];
  const app = nexacro.getApplication ? nexacro.getApplication() : window.application; const seen = new Set(); let f = null;
  function w(o,d){ if(!o||d>9||seen.has(o)||f) return; seen.add(o);
    try{ if(o.ds_supVeridcLst&&o.ds_supVeridcLst.getRowCount){ f=o; return; } }catch(e){}
    const k=[]; try{ if(o.all) for(let i=0;i<o.all.length;i++) k.push(o.all[i]); }catch(e){}
    try{ if(o.frames) for(let i=0;i<o.frames.length;i++) k.push(o.frames[i]); }catch(e){}
    for(const n of ['mainframe','form','frame','div_work','div_main','childframe']){ try{ if(o[n]) k.push(o[n]); }catch(e){} }
    k.forEach(x=>w(x,d+1)); }
  w(app,0); if(!f) return 'NOTFOUND';
  const cols=['stkinYmd','poYmd','stkinCntrCd','gdsCd','gdsNm','poQty','supQty','supTobeQty','costUprc','vatSpCd','cnfmYn','stkinStatCd','cntrPoSlipNo','cntrPoSlipSeq','rmkCont'];
  const out={};
  for (const d of DATES) {
    f.div_inqCond.cal_stkinStrtYmd.set_value(d); f.div_inqCond.cal_stkinEndYmd.set_value(d);
    f.ds_supVeridcLst.clearData(); f.fn_search_callback('',true);
    for(let t=0;t<30;t++){ await new Promise(r=>setTimeout(r,500)); if(f.ds_supVeridcLst.getRowCount()>0) break; }
    await new Promise(r=>setTimeout(r,800));
    const ds=f.ds_supVeridcLst, rows=[]; for(let r=0;r<ds.getRowCount();r++) rows.push(cols.map(c=>ds.getColumn(r,c)));
    out[d]={cols,rows};
  }
  window.__r1out=out;
  return Object.fromEntries(Object.entries(out).map(([k,v])=>[k,v.rows.length]));
})()
```
4. 납품일마다 `JSON.stringify(window.__r1out['YYYYMMDD'])` 결과를 `r1_YYYY-MM-DD.json` 으로 저장한다(행 0개면 건너뜀).

## 3. 저장
```bash
cd aky-order
for f in ../r1_*.json; do d=$(basename $f .json | sed 's/r1_//'); DASH_PW="$DASH_PW" python3 vault.py enc $f inputs/r1/$d.json.enc; done
git add inputs/r1 && git commit -m "납품확인서 수집 $(TZ=Asia/Seoul date '+%m-%d %H:%M')" && git push
```
푸시하면 깃허브 "올영 손익 갱신" 작업이 시트(자동_품목·자동_월집계)와 손익 대시보드를 자동으로 갱신한다.

## 4. 알림
`납품수량`(검수 확정)이 처음으로 들어온 납품일이 있으면 "11.24(화) 납품분 실매출로 확정됐어요: 매출 ○○, 마진 ○○" 처럼 짧게 알린다.
