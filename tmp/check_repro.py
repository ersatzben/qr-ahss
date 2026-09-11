import xlrd, collections, re
# QR
b=xlrd.open_workbook('data_sources/resdata1011.xls'); s=b.sheet_by_name('QR1011data')
def g_uoa(u):
 u=int(u)
 if u<=13 or u==15: return 'Health'
 if u==14 or 16<=u<=32: return 'STEM'
 if 34<=u<=46: return 'Social Sciences'
 if u==33 or 47<=u<=67: return 'Arts and Humanities'
 return 'OTHER'
qr=collections.defaultdict(lambda:[0]*4)
for r in range(6,s.nrows):
 g=g_uoa(s.cell_value(r,2))
 for i,c in enumerate(range(17,21)): qr[g][i]+=float(s.cell_value(r,c) or 0)/1000
# HESA external
b=xlrd.open_workbook('tmp/hesa1011/tables/Finance_Plus_1011_Table_5b.xls')
cc_group={}
for x in [1,2,4,5,6,7,8]:cc_group[x]='Health'
for x in [3,10,11,12,13,14,16,17,18,19,20,21,23,24,25]:cc_group[x]='STEM'
for x in [26,27,28,29,34,38,41]:cc_group[x]='Social Sciences'
for x in [30,31,33,35,37]:cc_group[x]='Arts and Humanities'
ext=collections.defaultdict(lambda:[0]*7)
industry_by_inst_group=collections.defaultdict(float)
# RC, char, govt, industry, overseas, other source, total raw
for s in b.sheets():
 for r in range(6,s.nrows):
  reg=str(s.cell_value(r,2)).strip()
  if reg in ['WALE','SCOT','NIRE','']: continue
  for c in range(4,s.ncols,14):
   name=str(s.cell_value(3,c)).strip()
   m=re.match(r'(\d{2}) ',name)
   if not m: continue
   cc=int(m.group(1)); g=cc_group.get(cc)
   if not g: continue
   v=[float(s.cell_value(r,c+j) or 0) for j in range(14)]
   vals=[v[0],v[1]+v[2],v[3],v[4],sum(v[5:12]),v[12],v[13]]
   for j,z in enumerate(vals): ext[g][j]+=z
   inst=re.sub(r"\D", "", str(s.cell_value(r,0))).zfill(4)[-4:]
   industry_by_inst_group[(inst,g)]+=v[4]
# Split institution-level business QR across disciplines using its 2010/11 UK-industry income.
business=collections.defaultdict(float)
bb=xlrd.open_workbook("data_sources/business1011.xls").sheet_by_name("business")
for r in range(5,bb.nrows):
 inst=re.sub(r"\D", "", str(bb.cell_value(r,0))).zfill(4)[-4:]
 grant=float(bb.cell_value(r,5) or 0)/1000
 weights={g:industry_by_inst_group[(inst,g)] for g in ["STEM","Social Sciences","Arts and Humanities","Health"]}
 total=sum(weights.values())
 if total:
  for g,w in weights.items(): business[g]+=grant*w/total
for g in ['STEM','Social Sciences','Arts and Humanities','Health']:
 main,lon,rdp,chqr=qr[g]
 oth=lon+rdp+chqr+business[g]
 rc,ch,gov,ind,ov,other,total=ext[g]
 denom=main+oth+rc+ch+gov+ind+ov
 print('\n',g)
 print('qr',main,oth,qr[g], 'business',business[g], 'external',ext[g], 'denom',denom)
 print('pct main other rc gov char ind ov QR',*[round(x/denom*100,3) for x in [main,oth,rc,gov,ch,ind,ov,main+oth]])
