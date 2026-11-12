import sys,re,csv,os,datetime as dt
import os, sys, subprocess, tempfile, shutil
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(os.path.dirname(HERE))
sys.path.insert(0, os.path.dirname(HERE))
WB = os.path.join(ROOT, "Data_and_Excel", "CNG_Safety_Risk_Assessment_Workbook.xlsx")
TMP = tempfile.mkdtemp(prefix="cng_qa_")
shutil.copy(WB, os.path.join(TMP, "in.xlsx"))
from lo_recalc import recalc_convert
recalc_convert(os.path.join(TMP, "in.xlsx"), os.path.join(TMP, "out"))
import build_data, config as C
from openpyxl import load_workbook
D=build_data.build();H=D['hazards'];A=D['actions'];R=D['records'];CL=D['checklist']
res=[]
def chk(name,ok,detail=''): res.append((name,bool(ok),detail)); print(('PASS ' if ok else 'FAIL ')+name,detail)
chk('Hazard IDs unique',len({h['hazard_id'] for h in H})==len(H),f'{len(H)} hazards')
chk('Action IDs unique',len({a['action_id'] for a in A})==len(A),f'{len(A)} actions')
chk('Record IDs unique',len({r['record_id'] for r in R})==len(R),f'{len(R)} records')
chk('Checklist IDs unique',len({c['item_id'] for c in CL})==len(CL),f'{len(CL)} items')
chk('Scores = L x S',all(h['init_score']==h['L']*h['S'] and h['res_score']==h['rL']*h['rS'] for h in H))
chk('Band boundaries consistent',all(h['init_cat']==C.band(h['init_score']) and h['res_cat']==C.band(h['res_score']) for h in H))
chk('Residual <= initial (L and S)',all(h['rL']<=h['L'] and h['rS']<=h['S'] for h in H))
chk('Residual within hierarchy credit limits',all(h['credit_ok'] for h in H))
ids={h['hazard_id'] for h in H}
chk('Action->hazard references valid',all(a['hazard_id'] in ids for a in A))
chk('Record->hazard references valid',all(r['hazard_id'] in ids for r in R))
aid={a['action_id'] for a in A}
chk('Record->action references valid',all((not r['action_id']) or r['action_id'] in aid for r in R))
chk('Every hazard has >=1 action',all(any(a['hazard_id']==h['hazard_id'] for a in A) for h in H))
sd=C.STATUS_DATE
def st(a):
    if a['completion'] and a['evidence'] and a['verified_on']: return 'Closed (Verified)'
    if a['completion']: return 'Completed (Awaiting Verification)'
    if a['target']<sd: return 'Overdue'
    return 'In Progress' if a['recorded_status']=='In Progress' else 'Not Started'
chk('Action status = independent recomputation',all(st(a)==a['status'] for a in A))
chk('No overdue action has a completion date',all(not a['completion'] for a in A if a['status']=='Overdue'))
chk('No action closed without date+evidence+verification',all(a['completion'] and a['evidence'] and a['verified_on'] for a in A if a['status']=='Closed (Verified)'))
chk('Past-target open actions are never Closed',all(a['status']!='Closed (Verified)' for a in A if (not a['completion']) and a['target']<sd))
chk('No future-dated completions/records',all((not a['completion'] or a['completion']<=sd) for a in A) and all(r['date']<=sd for r in R))
chk('Record types limited to 5 defined types',{r['record_type'] for r in R}<=set(C.RECORD_TYPES))
chk('>=35 hazards, >=50 records',len(H)>=35 and len(R)>=50)
chk('All rows tagged SIMULATED',all('SIMULATED' in h['data_status'].upper() for h in H) and all('SIMULATED' in a['data_status'].upper() for a in A) and all('SIMULATED' in r['data_status'].upper() for r in R))
# text scan for attribution
txt=' '.join(str(v) for r in R for v in r.values())+' '.join(str(v) for h in H for v in h.values())
chk('No record/hazard text mentions Powergas as source',('powergas' not in txt.lower()))
# workbook vs python
wb=load_workbook(os.path.join(TMP,'out','in.xlsx'),data_only=True)
ws=wb['Hazard Register'];bad=0
hm={h['hazard_id']:h for h in H}
hdr=[c.value for c in ws[5]]
for row in ws.iter_rows(min_row=6,max_row=45,values_only=True):
    if not row[0]: continue
    h=hm[row[0]]
    if not(row[12]==h['init_score'] and row[13]==h['init_cat'] and row[18]==h['res_score'] and row[19]==h['res_cat'] and row[21]=='OK'): bad+=1
chk('Workbook Hazard Register formulas = Python (scores, categories, residual check)',bad==0,f'{bad} mismatches')
ws=wb['Corrective Action Tracker'];bad=0;am={a['action_id']:a for a in A}
for row in ws.iter_rows(min_row=6,max_row=59,values_only=True):
    if not row[0]: continue
    a=am[row[0]]
    if not(row[20]==a['status'] and (row[21] or 0)==a['days_overdue'] and row[24]=='OK'): bad+=1
chk('Workbook tracker status/overdue formulas = Python; Status Check all OK',bad==0,f'{bad} mismatches')
# reference urls present
import sources
chk('Every reference has a URL and status',all(s['url'].startswith('http') and s['status'] for s in sources.SOURCES),f'{len(sources.SOURCES)} sources')
open(os.path.join(TMP,'audit_results.txt'),'w').write('\n'.join(('PASS' if o else 'FAIL')+' | '+n+' | '+d for n,o,d in res))
print('SUMMARY',sum(o for _,o,_ in res),'/',len(res))
