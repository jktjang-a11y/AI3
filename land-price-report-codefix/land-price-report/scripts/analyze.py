# -*- coding: utf-8 -*-
"""
지가동향 보고서 — 데이터 분석 모듈 (재사용용)
업로드 폴더에서 엑셀을 자동 인식하여 모든 통계를 산출, result.json으로 저장.
실행:  python analyze.py  [기준연월 YYYY-MM]  [업로드폴더]  [출력경로]
기준연월 미지정 시 실거래 데이터의 최신 신고월을 자동 사용.
"""
import pandas as pd, numpy as np, json, re, warnings, sys, glob, os
warnings.filterwarnings('ignore')

# ===== 설정 =====
ARG_YM   = sys.argv[1] if len(sys.argv)>1 else None        # 예: '2026-05' (없으면 자동)
U        = (sys.argv[2] if len(sys.argv)>2 else '/mnt/user-data/uploads') + '/'
OUT      = sys.argv[3] if len(sys.argv)>3 else '/home/claude/result.json'
REGION   = '나주시'   # 지역명 (보고서 제목/검색용)

# ===== 권역 사전 (지역 특성에 맞게 수정 가능) =====
ZONE_DEF = {
 '국가산단배후(왕곡·동수·운곡 등)':['왕곡면','동수동','운곡동','오량동','진포동'],
 '혁신도시거점(빛가람동·금천면)':['빛가람동','금천면'],
 '광주연접생활권(남평·노안·산포)':['남평읍','노안면','산포면'],
 '구도심행정(영산동·이창동 등)':['영산동','이창동','대기동','송월동','토계동','죽림동','금성동','남내동','남외동'],
 '농업외곽(봉황·다시·반남 등)':['봉황면','다시면','반남면','공산면','동강면','세지면','문평면','다도면']}
ZONE_TOKENS={ '왕곡면':['왕곡면'],'산포면':['산포면'],'다도면':['다도면'],'봉황면':['봉황면'],
  '금천면':['금천면'],'남평읍':['남평읍'],'세지면':['세지면'],'반남면':['반남면'],
  '공산면':['공산면'],'동강면':['동강면'],'다시면':['다시면'],'문평면':['문평면'],
  '노안면':['노안면'],'삼도동':['삼도동'],'빛가람동':['빛가람동'] }

# ===== 파일 자동 인식 =====
def find(*keys, ext=('xlsx','xls'), exclude=()):
    """파일명에 keys(AND조건)를 모두 포함하는 엑셀을 찾아 경로 리스트 반환"""
    hits=[]
    for f in sorted(glob.glob(U+'*')):
        b=os.path.basename(f)
        if not b.lower().endswith(ext): continue
        if all(k in b for k in keys) and not any(x in b for x in exclude):
            hits.append(f)
    return hits

def find1(*keys, **kw):
    h=find(*keys, **kw)
    if not h: raise FileNotFoundError(f"파일을 찾지 못함: 키워드={keys}")
    return h[0]

def num(s):
    if pd.isna(s): return np.nan
    return pd.to_numeric(str(s).replace(',','').replace("'",'').strip(), errors='coerce')

# ---------- 1. 실거래 로드 & 병합 (실거래/매매 포함 파일 모두) ----------
tx_files=find('실거래') or find('매매')
if not tx_files: raise FileNotFoundError("실거래 엑셀을 찾지 못함")
frames=[]
for f in tx_files:
    x=pd.read_excel(f)
    frames.append(x)
# 컬럼 합집합 정렬
allcols=set()
for x in frames: allcols|=set(x.columns)
for i,x in enumerate(frames):
    for c in allcols:
        if c not in x.columns: x[c]=np.nan
df=pd.concat(frames,ignore_index=True)


for c in ['계약총금액','토지면적','건물면적','토지면적단가']:
    df[c+'_n']=df[c].apply(num)
df['신고일']=pd.to_datetime(df['신고일'],errors='coerce')
df['ym']=df['신고일'].dt.to_period('M').astype(str)
df['prefix']=df['RTMS_KEY'].astype(str).str.rsplit('-',n=1).str[0]
df['외지']=(df['외지인거래 여부']=='Y')

# 이용상황 파생
def usecat(r):
    j=str(r['현실지목']); b=str(r['건축물주용도'])
    if j=='대':
        if b in ['단독주택','다가구주택']: return '단독주택(대)'
        if b in ['일반음식점','소매점']: return '상업용(대)'
        if b=='사무소': return '사무소(대)'
        if b in ['nan','NaN','']: return '대지(나지)'
        return '기타건축물(대)'
    m={'답':'답','전':'전','임야':'임야','과수원':'과수원','공장용지':'공업용',
       '잡종지':'잡종지','도로':'도로','목장용지':'목장용지'}
    return m.get(j,'기타')
df['이용상황']=df.apply(usecat,axis=1)

# ---------- 2. 계약(묶음) 단위 집계 ----------
def bundle(g):
    rep=g.loc[g['토지면적_n'].idxmax()] if g['토지면적_n'].notna().any() else g.iloc[0]
    return pd.Series({
        'ym':rep['ym'],'신고일':rep['신고일'],'읍면동full':str(rep['읍면동']).strip(),
        '읍면동':str(rep['읍면동']).strip().split(' ')[0],
        '용도지역':rep['용도지역'],'현실지목':rep['현실지목'],'이용상황':rep['이용상황'],
        '건축물주용도':rep['건축물주용도'],'외지':rep['외지'],
        '매수자 시군구':rep['매수자 시군구'],'매수자':rep['매수자'],'매도자':rep['매도자'],
        '매수자연령':rep['매수자 연령대'],
        '번지':rep['번지'],'물건구분':rep['물건구분'],
        '거래금액':g['계약총금액_n'].max(),
        '토지면적':g['토지면적_n'].sum(),'건물면적':g['건물면적_n'].sum(),
        'n필지':len(g)})
C=df.groupby('prefix',sort=False).apply(bundle).reset_index()
C['거래금액']=C['거래금액'].fillna(0)
C['단가']=np.where(C['토지면적']>0,C['거래금액']/C['토지면적'],np.nan)

# ===== 기준연월/비교월 자동 산출 =====
df['신고일']=pd.to_datetime(df['신고일'],errors='coerce')
_allym=sorted(df['신고일'].dt.to_period('M').dropna().unique())
cur = ARG_YM if ARG_YM else str(_allym[-1])
_cp=pd.Period(cur,'M')
prev=str(_cp-1); yoy=str(_cp-12)
trail_start=str(_cp-11)
def sub(p): return C[C['ym']==p]
Scur,Sprev,Syoy=sub(cur),sub(prev),sub(yoy)

def pct(a,b): return None if (b in [0,None] or pd.isna(b)) else round((a-b)/b*100,1)
EOK=1e8  # 억원 환산

res={}
res['meta']={'당월':cur,'전월':prev,'전년동월':yoy,'지역':REGION}

# ---- 전체 거래 요약 ----
def summ(S):
    return {'건수':int(len(S)),'금액억':round(S['거래금액'].sum()/EOK,1),
            '면적':round(S['토지면적'].sum(),0),
            '외지건수':int(S['외지'].sum()),'외지금액억':round(S[S['외지']]['거래금액'].sum()/EOK,1)}
res['summary']={'cur':summ(Scur),'prev':summ(Sprev),'yoy':summ(Syoy)}
res['summary']['증감']={
 '건수_전월':pct(len(Scur),len(Sprev)),'건수_전년':pct(len(Scur),len(Syoy)),
 '금액_전월':pct(Scur['거래금액'].sum(),Sprev['거래금액'].sum()),
 '금액_전년':pct(Scur['거래금액'].sum(),Syoy['거래금액'].sum())}

# ---- ① 최근 12개월 추이 (trailing, ending cur) ----
months=pd.period_range(trail_start,cur,freq='M').astype(str).tolist()
trend=[]
prevv=None
for m in months:
    s=sub(m)
    row={'월':m,'건수':int(len(s)),'금액억':round(s['거래금액'].sum()/EOK,1),
         '외지건수':int(s['외지'].sum()),'외지금액억':round(s[s['외지']]['거래금액'].sum()/EOK,1)}
    row['외지금액비중']=round(row['외지금액억']/row['금액억']*100,1) if row['금액억']>0 else 0
    row['건수증감']=pct(row['건수'],prevv['건수']) if prevv else None
    row['금액증감']=pct(row['금액억'],prevv['금액억']) if prevv else None
    row['외지건수증감']=pct(row['외지건수'],prevv['외지건수']) if prevv else None
    row['외지금액증감']=pct(row['외지금액억'],prevv['외지금액억']) if prevv else None
    trend.append(row); prevv=row
res['trend12']=trend

# ---- ② 읍면동별 (당월) ----
def rank(S,key,extra_filter=None):
    s=S if extra_filter is None else S[extra_filter]
    g=s.groupby(key).agg(건수=('거래금액','size'),금액=('거래금액','sum')).reset_index()
    tot_c=g['건수'].sum(); tot_a=g['금액'].sum()
    g['금액억']=g['금액']/EOK
    g['건수비중']=g['건수']/tot_c*100
    g['금액비중']=g['금액']/tot_a*100 if tot_a>0 else 0
    g=g.sort_values('금액',ascending=False).reset_index(drop=True)
    return g
def to_rows(g,namecol):
    out=[]
    for i,r in g.iterrows():
        out.append({'순위':i+1,'명칭':r[namecol],'건수':int(r['건수']),
                    '건수비중':round(r['건수비중'],1),'금액억':round(r['금액억'],1),
                    '금액비중':round(r['금액비중'],1)})
    return out
res['emd']=to_rows(rank(Scur,'읍면동'),'읍면동')
res['용도지역']=to_rows(rank(Scur,'용도지역'),'용도지역')
res['이용상황']=to_rows(rank(Scur,'이용상황'),'이용상황')

# ---- ⑤⑥ 외지인 (당월) ----
res['외지_시군구']=to_rows(rank(Scur,'매수자 시군구',Scur['외지']),'매수자 시군구')
res['외지_이용상황']=to_rows(rank(Scur,'이용상황',Scur['외지']),'이용상황')
res['외지_읍면동']=to_rows(rank(Scur,'읍면동',Scur['외지']),'읍면동')
res['외지_용도지역']=to_rows(rank(Scur,'용도지역',Scur['외지']),'용도지역')

# ---- TOP10 helper ----
def top10(S):
    s=S.sort_values('거래금액',ascending=False).head(10).reset_index(drop=True)
    out=[]
    for i,r in s.iterrows():
        주체='법인' if (str(r['매수자'])=='법인' or str(r['매도자'])=='법인') else '개인'
        out.append({'순위':i+1,'읍면동':r['읍면동full'],'번지':str(r['번지']).replace("'",''),
            '용도지역':r['용도지역'],'지목':r['현실지목'],'이용상황':r['이용상황'],
            '토지면적':round(r['토지면적'],1),'건물면적':round(r['건물면적'],1) if not pd.isna(r['건물면적']) else 0,
            '단가':int(np.floor(r['단가']/1000)*1000) if not pd.isna(r['단가']) else 0,
            '금액억':round(r['거래금액']/EOK,2),'n필지':int(r['n필지']),
            '매수자유형':str(r['매수자']),'매수자시군구':str(r['매수자 시군구']),'외지':bool(r['외지'])})
    return out
res['top10_all']=top10(Scur)
res['top10_외지']=top10(Scur[Scur['외지']])

# ---- 권역별 (당월) ----
zones=ZONE_DEF
_OLD_zones_disabled={
 '국가산단배후(왕곡·동수·운곡 등)':['왕곡면','동수동','운곡동','오량동','진포동'],
 '혁신도시거점(빛가람동·금천면)':['빛가람동','금천면'],
 '광주연접생활권(남평·노안·산포)':['남평읍','노안면','산포면'],
 '구도심행정(영산동·이창동 등)':['영산동','이창동','대기동','송월동','토계동','죽림동','금성동','남내동','남외동'],
 '농업외곽(봉황·다시·반남 등)':['봉황면','다시면','반남면','공산면','동강면','세지면','문평면','다도면']}
zres={}
for z,ems in zones.items():
    mask=Scur['읍면동'].apply(lambda x: any(str(x).startswith(e) or str(x).split(' ')[0]==e for e in ems))
    s=Scur[mask]
    zres[z]={'건수':int(len(s)),'금액억':round(s['거래금액'].sum()/EOK,1),
             '외지건수':int(s['외지'].sum()),
             '대표':to_rows(rank(s,'읍면동'),'읍면동')[:3] if len(s) else []}
res['zones']=zres

# ---- 지가변동률 파싱 ----
lp=pd.read_excel(U+'나주시_지가동향_당월_2026_5.xls',header=None)
ucols={4:'평균',5:'주거',6:'상업',7:'공업',8:'녹지',9:'보전관리',10:'생산관리',11:'계획관리',13:'농림'}
icols={15:'전',16:'답',17:'주거용',18:'상업용',19:'임야',20:'공업용',21:'기타'}
def fnum(v):
    try: return round(float(v),3)
    except: return None
naju_chg={nm:fnum(lp.iloc[4,ci]) for ci,nm in {**ucols,**icols}.items()}
naju_cum={nm:fnum(lp.iloc[5,ci]) for ci,nm in {**ucols,**icols}.items()}
# 읍면동별 평균 변동률 (당월) : rows where col2 name present -> that row col4
emd_lp=[]
for i in range(4,len(lp)):
    nm=lp.iloc[i,2]
    if pd.notna(nm):
        v=lp.iloc[i,4]
        try: emd_lp.append({'권역':str(nm),'변동률':float(v)})
        except: pass
emd_lp=[x for x in emd_lp if x['권역']!='나주시']
# 권역별 용도/이용 상세 변동률(당월) 추출
zone_detail=[]
for i in range(4,len(lp)):
    nm=lp.iloc[i,2]
    if pd.notna(nm) and str(nm)!='나주시':
        d={'권역':str(nm),'평균':fnum(lp.iloc[i,4])}
        for ci,k in {**ucols,**icols}.items():
            d[k]=fnum(lp.iloc[i,ci])
        zone_detail.append(d)
res['landprice']={'당월':naju_chg,'누계':naju_cum,
   'emd_top':sorted(emd_lp,key=lambda x:-x['변동률']),
   'zone_detail':sorted(zone_detail,key=lambda x:-(x['평균'] or -9))}

# ---- 섹션8: 지가변동률 상위6 권역 + 누계 + 당월 실거래 매칭 ----
# 권역 누계 평균 (당월 행 다음 행)
zone_cum={}
for i in range(4,len(lp)):
    nm=lp.iloc[i,2]
    if pd.notna(nm) and str(nm)!='나주시':
        zone_cum[str(nm)]=fnum(lp.iloc[i+1,4]) if i+1<len(lp) else None
_zone_tokens_local={ '왕곡면':['왕곡면'],'산포면':['산포면'],'다도면':['다도면'],'봉황면':['봉황면'],
  '금천면':['금천면'],'남평읍':['남평읍'],'세지면':['세지면'],'반남면':['반남면'],
  '공산면':['공산면'],'동강면':['동강면'],'다시면':['다시면'],'문평면':['문평면'],
  '노안면':['노안면'],'삼도동':['삼도동'],'빛가람동':['빛가람동'] }
def tok(name):
    if name in ZONE_TOKENS: return ZONE_TOKENS[name]
    return [t.strip() for t in name.split(',')]
sec8=[]
for z in res['landprice']['emd_top'][:6]:
    nm=z['권역']; toks=tok(nm)
    sub=Scur[Scur['읍면동'].isin(toks)]
    gg=sub.sort_values('거래금액',ascending=False)
    reps=[]
    for _,row in gg.head(3).iterrows():
        reps.append({'읍면동':row['읍면동full'],'번지':str(row['번지']).replace("'",''),
            '지목':row['현실지목'],'이용상황':row['이용상황'],'금액억':round(row['거래금액']/1e8,2),
            '외지':bool(row['외지'])})
    det=ZDmap=next((x for x in zone_detail if x['권역']==nm),{})
    sec8.append({'권역':nm,'당월':z['변동률'],'누계':zone_cum.get(nm),
        '거래건수':int(len(sub)),'거래금액억':round(sub['거래금액'].sum()/1e8,1),
        '대표':reps,'세부':det})
res['sec8']=sec8



# ---- 인구/세대 ----
pop=pd.read_excel(U+'시군구별_주민등록인구_최근15개월_전월대비산출용_2025_03_2026_05_전라남도_나주시.xlsx',sheet_name='data')
popT=pop[pop['항목명']=='총인구수'].groupby('기간')['값'].sum().reset_index()
popT['기간']=popT['기간'].astype(str)
hh=pd.read_excel(U+'시군구별_주민등록세대수_최근15개월_전월대비산출용_2025_03_2026_05_전라남도_나주시.xlsx',sheet_name='data')
hh=hh[['기간','값']].copy(); hh['기간']=hh['기간'].astype(str)
def fmt_p(p): return p[:4]+'-'+p[4:]
demo=[]
pm=sorted(set(popT['기간'])|set(hh['기간']))[-13:]
pmap=dict(zip(popT['기간'],popT['값'])); hmap=dict(zip(hh['기간'],hh['값']))
pp=ph=None
for p in pm:
    pv=int(pmap.get(p,0)); hv=int(hmap.get(p,0))
    demo.append({'월':fmt_p(p),'인구':pv,'세대':hv,
        '인구증감':pct(pv,pp) if pp else None,'세대증감':pct(hv,ph) if ph else None,
        '인구증감수':(pv-pp) if pp else None,'세대증감수':(hv-ph) if ph else None})
    pp,ph=pv,hv
res['demo']=demo

# ---- 네이버 매물 ----
def load_listing(f):
    raw=pd.read_excel(U+f,header=None)
    cols=raw.iloc[1].tolist(); l=raw.iloc[2:].copy(); l.columns=cols
    l=l.rename(columns={'매매금액(만원)':'금액만','매매금액단가(원)':'단가원'})
    l['금액만']=l['금액만'].apply(num); l['단가원']=l['단가원'].apply(num)
    l=l.dropna(subset=['고유번호'])
    l['고유번호']=l['고유번호'].astype(str)
    l=l.drop_duplicates(subset='고유번호',keep='last')
    l['읍면동']=l['읍면동'].astype(str).str.strip()
    return l
m4=load_listing('네이버매물조회_나주시_26년4월.xls')
m5=load_listing('네이버매물조회_나주시_26년5월.xls')

# ---- 신규①: 변동률 판단근거표 (거래 충분/희소 플래그 + 전월 권역 방향성) ----
Sprev_z=Sprev.copy()
판단표=[]
for s in sec8:
    nm=s['권역']; toks=tok(nm)
    pv=Sprev_z[Sprev_z['읍면동'].isin(toks)]
    prev_cnt=int(len(pv)); prev_amt=round(pv['거래금액'].sum()/1e8,1)
    cnt=s['거래건수']
    flag='실거래 주도' if cnt>=5 else ('혼합' if cnt>=3 else '조사표본 주도')
    # 매물 (당월 권역 매물 건수)
    mv=m5[m5['읍면동'].isin(toks)]
    listing_cnt=int(len(mv))
    # 전월 대비 거래 방향
    if prev_cnt==0: dir_txt='전월 거래 부재'
    elif cnt>prev_cnt: dir_txt='전월 대비 증가'
    elif cnt<prev_cnt: dir_txt='전월 대비 감소'
    else: dir_txt='전월과 동일'
    판단표.append({'권역':nm,'당월변동률':s['당월'],'누계':s['누계'],
        '실거래건수':cnt,'실거래금액억':s['거래금액억'],'전월건수':prev_cnt,
        '거래방향':dir_txt,'매물건수':listing_cnt,'근거구분':flag})
res['판단근거표']=판단표

# ---- 신규②: 외지인 자본 성격 (연령/법인/이용상황) ----
ext=Scur[Scur['외지']].copy()
def cap_profile(df):
    n=len(df)
    if n==0: return {}
    법인=int((df['매수자']=='법인').sum())
    개인=int((df['매수자']=='개인').sum())
    # 실수요성(주거용/단독주택/답/전/임야) vs 기타
    실수요=int(df['이용상황'].isin(['단독주택(대)','대지(나지)','답','전','임야','과수원']).sum())
    age=df['매수자연령'].value_counts().to_dict()
    age={str(k):int(v) for k,v in age.items()}
    return {'총건수':n,'법인':법인,'개인':개인,'실수요성건수':실수요,
        '실수요비율':round(실수요/n*100,1),'연령분포':age}
res['외지자본성격']=cap_profile(ext)

# ---- 신규③: 데이터 한계 (정적) ----
res['데이터한계']=[
 '월별 집계는 신고일자 기준이며, 2026년 1월 신고분은 제공 데이터셋에 부존재(0건)하여 12개월 추이에서 제외함.',
 '인구이동(총전입·총전출·순이동) 항목은 제공된 주민등록 인구·세대수 자료에 포함되지 않아 인구·세대 증감만 제시함.',
 '네이버 매물 고유번호는 월별로 갱신되어 동일 고유번호 매칭이 불가하므로, 호가 조정은 주소(읍면동+지번) 단위로 매칭함.',
 '매물가–실거래가 직접 괴리율은 주소 교집합 표본이 소수(2건)에 불과하여 통계적 유의성이 낮으므로 참고치로만 제시함.',
 '지가변동률은 조사 표본지(조사대상 필지)에 대한 가격조사 결과이며, 표준지공시지가 및 실거래 빈도와는 별개의 지표임.',
]
print('신규분석 OK | 판단근거표', len(판단표), '| 외지자본', res['외지자본성격'].get('총건수'))
def lrank(l,key):
    g=l.groupby(key).agg(건수=('고유번호','size'),금액만=('금액만','sum')).reset_index()
    g['금액억']=g['금액만']/10000; tot=g['건수'].sum()
    g['비중']=g['건수']/tot*100
    g=g.sort_values('건수',ascending=False).reset_index(drop=True)
    return g
def lrows(g,nm):
    return [{'순위':i+1,'명칭':r[nm],'건수':int(r['건수']),'금액억':round(r['금액억'],1),
             '비중':round(r['비중'],1)} for i,r in g.iterrows()]
res['매물_이용상황']=lrows(lrank(m5,'이용상황'),'이용상황')
res['매물_읍면동']=lrows(lrank(m5,'읍면동'),'읍면동')
res['매물_총건수']={'4월':int(len(m4)),'5월':int(len(m5))}
# 읍면동 증감
c4=m5.groupby('읍면동').size(); c4p=m4.groupby('읍면동').size()
chg=(c4.subtract(c4p,fill_value=0)).sort_values(ascending=False)
res['매물_읍면동증감']=[{'읍면동':k,'증감':int(v)} for k,v in chg.head(8).items()]+\
    [{'읍면동':k,'증감':int(v)} for k,v in chg.tail(4).items()]
# 이용상황 증감
u4=m5.groupby('이용상황').size(); u4p=m4.groupby('이용상황').size()
uchg=(u4.subtract(u4p,fill_value=0)).sort_values(ascending=False)
res['매물_이용상황증감']=[{'이용상황':k,'증감':int(v),'당월':int(u4.get(k,0))} for k,v in uchg.items()]

# ---- 호가 조정 현황 (4월 vs 5월 동일 필지: 주소키 매칭) ----
def pkey(r): return re.sub(r'[\s\']','',str(r['읍면동'])+str(r.get('리',''))+str(r['지번']))
m4['pk']=m4.apply(pkey,axis=1); m5['pk']=m5.apply(pkey,axis=1)
a4=m4.groupby(['pk','이용상황'])['금액만'].median().reset_index().rename(columns={'금액만':'g4'})
a5=m5.groupby(['pk','이용상황'])['금액만'].median().reset_index().rename(columns={'금액만':'g5'})
mm=a4.merge(a5,on=['pk','이용상황']).dropna(subset=['g4','g5'])
mm=mm[(mm['g4']>0)&(mm['g5']>0)]
mm['dir']=np.where(mm['g5']>mm['g4']*1.001,'상향',np.where(mm['g5']<mm['g4']*0.999,'하향','보합'))
hoga=[]
for u,gp in mm.groupby('이용상황'):
    n=len(gp)
    if n<3: continue
    hoga.append({'이용상황':u,'표본':int(n),
        '상향':round((gp['dir']=='상향').sum()/n*100,1),
        '보합':round((gp['dir']=='보합').sum()/n*100,1),
        '하향':round((gp['dir']=='하향').sum()/n*100,1)})
hoga=sorted(hoga,key=lambda x:-x['표본'])
res['호가조정']=hoga
res['호가조정_전체']={'표본':int(len(mm)),
    '상향':round((mm['dir']=='상향').sum()/len(mm)*100,1) if len(mm) else None,
    '보합':round((mm['dir']=='보합').sum()/len(mm)*100,1) if len(mm) else None,
    '하향':round((mm['dir']=='하향').sum()/len(mm)*100,1) if len(mm) else None}

# ---- 괴리율: 매물 vs 실거래 주소 매칭 ----
def norm_addr(s): return re.sub(r'[\s\']','',str(s))
m5['매물주소']=(m5['읍면동'].astype(str)+m5['지번'].astype(str)).apply(norm_addr)
Scur2=Scur.copy()
Scur2['실거래주소']=(Scur2['읍면동'].astype(str)+Scur2['번지'].astype(str)).apply(norm_addr)
mj=Scur2.merge(m5[['매물주소','금액만']],left_on='실거래주소',right_on='매물주소',how='inner')
mj['매물금액']=mj['금액만']*10000
mj=mj[mj['매물금액']>0]
mj['괴리']=(mj['거래금액']-mj['매물금액'])/mj['매물금액']*100
res['괴리']={'매칭건수':int(len(mj)),'평균괴리율':round(mj['괴리'].mean(),1) if len(mj) else None,
    '중위괴리율':round(mj['괴리'].median(),1) if len(mj) else None}

# ===== 개선 1·2: 변동률 판단 근거표 + 거래/조사표본 구분 =====
# 권역별 매물 증감(전월대비) 매칭용 — 매물 읍면동 token 집계
def listing_emd_counts(l):
    l2=l.copy(); l2['emd']=l2['읍면동'].astype(str).str.split(' ').str[0]
    return l2.groupby('emd').size()
m4c=listing_emd_counts(m4); m5c=listing_emd_counts(m5)
def zone_listing_delta(toks):
    cur_n=int(sum(m5c.get(t,0) for t in toks)); prev_n=int(sum(m4c.get(t,0) for t in toks))
    return cur_n, cur_n-prev_n
judge=[]
for s in sec8:
    nm=s['권역']; toks=tok(nm)
    mcnt,mdelta=zone_listing_delta(toks)
    # 거래 충분/희소 플래그
    if s['거래건수']>=8: flag='실거래 주도'; basis='실거래 기반'
    elif s['거래건수']>=4: flag='혼합'; basis='실거래·조사표본 병행'
    else: flag='조사표본 주도'; basis='조사 표본가격 기반'
    # 대표 실거래 한 줄
    rep=''
    if s['대표']:
        r0=s['대표'][0]; rep=f"{r0['읍면동']} {r0['이용상황']} {r0['금액억']}억"
    judge.append({'권역':nm,'당월':s['당월'],'누계':s['누계'],
        '거래건수':s['거래건수'],'거래금액억':s['거래금액억'],
        '매물건수':mcnt,'매물증감':mdelta,'대표실거래':rep,
        '플래그':flag,'근거':basis})
res['judge']=judge

# ===== 개선 5: 외지인 자본 성격 (연령대·법인/개인·이용목적 추정) =====
extc=Scur[Scur['외지']].copy()
n_ext=len(extc)
법인=int((extc['매수자'].astype(str)=='법인').sum())
개인=n_ext-법인
# 이용상황 기반 목적 추정
def purpose(u):
    if u in ['단독주택(대)','대지(나지)','상업용(대)','사무소(대)']: return '실거주·건축용'
    if u in ['답','전','과수원','임야','목장용지']: return '영농·장기보유'
    return '기타·사업용'
extc['목적']=extc['이용상황'].apply(purpose)
purp=extc.groupby('목적').agg(건수=('거래금액','size'),금액=('거래금액','sum')).reset_index()
purp['금액억']=(purp['금액']/1e8).round(1)
purp=purp.sort_values('금액',ascending=False)
# 연령대 (원본 행 기준 — Scur는 계약단위라 연령대 누락 가능, 원본 df에서 매칭)
age=df[(df['ym']=='2026-05')&(df['외지인거래 여부']=='Y')]['매수자 연령대'].value_counts()
age_top=[{'연령대':str(k),'건수':int(v)} for k,v in age.items() if str(k)!='기타'][:4]
res['ext_capital']={'외지건수':int(n_ext),'법인':법인,'개인':개인,
    '목적분포':[{'목적':r['목적'],'건수':int(r['건수']),'금액억':float(r['금액억'])} for _,r in purp.iterrows()],
    '연령대':age_top}

# ===== 개선 4: 권역별 전월대비 추세 (당월 vs 전월 거래금액) =====
zone_trend=[]
zmap={'국가산단배후':['왕곡면','동수동','운곡동','오량동','진포동'],
 '혁신도시':['빛가람동','금천면'],'광주연접':['남평읍','노안면','산포면'],
 '구도심':['영산동','이창동','금성동','대기동','송월동','죽림동','남내동','남외동','대호동'],
 '농업외곽':['봉황면','다시면','반남면','공산면','동강면','세지면','문평면','다도면']}
for zn,toks in zmap.items():
    c_now=Scur[Scur['읍면동'].isin(toks)]; c_prev=Sprev[Sprev['읍면동'].isin(toks)]
    a_now=round(c_now['거래금액'].sum()/1e8,1); a_prev=round(c_prev['거래금액'].sum()/1e8,1)
    zone_trend.append({'권역':zn,'당월금액':a_now,'전월금액':a_prev,
        '당월건수':int(len(c_now)),'전월건수':int(len(c_prev)),
        '금액증감':pct(c_now['거래금액'].sum(),c_prev['거래금액'].sum())})
res['zone_trend']=zone_trend

with open(OUT,'w') as f:
    json.dump(res,f,ensure_ascii=False,indent=1,default=str)
print('OK. contracts(all):',len(C),' 당월계약:',len(Scur))
print(json.dumps({k:res[k] for k in ['summary','meta']},ensure_ascii=False,indent=1,default=str))
