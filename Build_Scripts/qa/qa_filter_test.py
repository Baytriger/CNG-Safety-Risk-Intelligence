import sys,subprocess,shutil,os,datetime as dt
import os, sys, subprocess, tempfile, shutil
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.dirname(HERE))
WB = os.path.join(ROOT, "Data_and_Excel", "CNG_Safety_Risk_Assessment_Workbook.xlsx")
TMP = tempfile.mkdtemp(prefix="cng_qa_")
shutil.copy(WB, os.path.join(TMP, "in.xlsx"))
from lo_recalc import recalc_convert
recalc_convert(os.path.join(TMP, "in.xlsx"), os.path.join(TMP, "out"))
import build_data
from openpyxl import load_workbook
D=build_data.build();H={h['hazard_id']:h for h in D['hazards']};A=D['actions'];R=D['records']
combos=[dict(cat='Loading & Offloading'),dict(level='High'),dict(role='Station Supervisor'),dict(status='Overdue'),
 dict(d1=dt.date(2026,6,1),d2=dt.date(2026,8,31)),dict(cat='Maintenance & Contractors',level='Extreme',status='All'),dict(cat='Transportation & Fleet',role='Fleet Manager',d1=dt.date(2026,5,1),d2=dt.date(2026,10,31))]
allok=True
for k,cmb in enumerate(combos):
    wb=load_workbook(os.path.join(TMP,'in.xlsx')); ws=wb['Dashboard']
    ws['B5']=cmb.get('cat','All');ws['D5']=cmb.get('level','All');ws['F5']=cmb.get('role','All');ws['H5']=cmb.get('status','All')
    d1=cmb.get('d1',dt.date(2026,1,1));d2=cmb.get('d2',dt.date(2026,12,31));ws['J5']=dt.datetime.combine(d1,dt.time());ws['L5']=dt.datetime.combine(d2,dt.time())
    wb.save(os.path.join(TMP,f't{k}.xlsx'))
    recalc_convert(os.path.join(TMP,f't{k}.xlsx'),os.path.join(TMP,'o'))
    v=load_workbook(os.path.join(TMP,'o',f't{k}.xlsx'),data_only=True)['KPI Dashboard Data']
    got=[v[f'C{r}'].value for r in range(13,24)]
    def ok(x,c,key): return c.get(key) in (None,'All') or x==c[key]
    hz=[h for h in H.values() if ok(h['category'],cmb,'cat') and ok(h['init_cat'],cmb,'level') and ok(h['role'],cmb,'role') and d1<=h['assess_date']<=d2]
    ac=[a for a in A if ok(a['category'],cmb,'cat') and ok(a['init_cat'],cmb,'level') and ok(a['role'],cmb,'role') and ok(a['status'],cmb,'status') and d1<=a['identified']<=d2]
    rc=[r for r in R if ok(r['category'],cmb,'cat') and ok(H[r['hazard_id']]['init_cat'],cmb,'level') and ok(H[r['hazard_id']]['role'],cmb,'role') and d1<=r['date']<=d2]
    n=len(ac);cl=sum(a['status']=='Closed (Verified)' for a in ac)
    exp=[len(hz),sum(h['init_cat']=='Extreme' for h in hz),sum(h['init_cat']=='High' for h in hz),n,n-cl,sum(a['status']=='Overdue' for a in ac),sum(a['status']=='Completed (Awaiting Verification)' for a in ac),cl,(cl/n if n else 0),sum(r['record_type']=='Near miss' for r in rc),len(rc)]
    good=all(abs((g or 0)-e)<1e-9 for g,e in zip(got,exp));allok&=good
    print(k,cmb,'OK' if good else 'MISMATCH',got if not good else '',exp if not good else '')
print('ALL FILTER TESTS PASS' if allok else 'FAIL')
