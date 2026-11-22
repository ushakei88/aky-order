import openpyxl,datetime as dt,json,re,sys
from collections import defaultdict
def load(path):
    wb=openpyxl.load_workbook(path,data_only=True,read_only=True)
    out=[]
    for ws in wb.worksheets:
        if not re.fullmatch(r'\d\d\.\d\d',ws.title): continue
        yy,mm=map(int,ws.title.split('.'))
        if (yy,mm)<(25,7): continue
        rows=list(ws.iter_rows(max_row=600,max_col=21,values_only=True))
        hr=next((i for i,r in enumerate(rows) if r[0]=='구분'),None)
        if hr is None: continue
        kind=po=ib=None
        for r in rows[hr+3:]:
            if isinstance(r[0],str) and (r[0].startswith('■') or r[0]=='구분'): break
            if r[3] in (None,''): continue
            if r[0]: kind=r[0]
            if r[1]: po=r[1]
            if r[2]: ib=r[2]
            f=lambda i: float(r[i] or 0) if isinstance(r[i],(int,float)) else 0.0
            out.append({'month':f'20{yy:02d}-{mm:02d}','kind':kind,'po':po.date().isoformat() if isinstance(po,dt.datetime) else None,
              'inb':ib.date().isoformat() if isinstance(ib,dt.datetime) else None,'brand':str(r[3]).strip(),'center':str(r[4] or '').strip(),
              'cost':f(5),'costT':f(7),'sales':f(8),'salesT':f(10),'edi':f(13),'logi':f(16),'margin':f(17),'post':f(19),**({'status':'가','auto':True} if str(r[20] or '').startswith('가매출') else ({'status':'실','auto':True} if str(r[20] or '').startswith('자동') else {}))})
    return out
if __name__=='__main__':
    d=load(sys.argv[1]); json.dump(d,open('rows.json','w'),ensure_ascii=False)
    m=defaultdict(lambda:defaultdict(float))
    for r in d: m[r['month']][r['brand']]+=r['salesT']
    for k in sorted(m): 
        t=sum(m[k].values()); top=sorted(m[k].items(),key=lambda x:-x[1])[:3]
        print(k,round(t/1e6),'M',[(b,round(v/1e6)) for b,v in top])
    print(sorted({r['kind'] for r in d},key=str)); print(sorted({r['center'] for r in d}))
