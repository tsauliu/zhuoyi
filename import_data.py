from openpyxl import load_workbook
from operator import itemgetter
import csv
import numpy

ivols=[]

for sheet in range(1,5,1):
    count=0
    wb = load_workbook(filename='./data/ivol'+str(sheet)+'.xlsx', read_only=True)
    ws = wb['Sheet1'] # ws is now an IterableWorksheet
    for row in ws.rows:
        if count >=1:
            ivols.append({"stkcd":row[0].value,"date":row[1].value,"ivol":row[3].value})
        count+=1

ivols=sorted(ivols,key=itemgetter("date"))

for i in range(1,5,1):
    print ivols[i]

reader = csv.reader(open('./data/ret1.csv', 'rU'))
count=0
atrributes=[]
for row in reader:
    if count>3:
        atrributes.append({"stkcd":row[0],"year":row[1],"mon":row[2],"mktvalue":row[3],"Dturn":row[4],"ret":row[6]})
    count+=1

for i in range(1,5,1):
    print atrributes[i]

rongzi=csv.reader(open('./data/rongzi.csv', 'rU'))
rzlst=[]
for item in rongzi:
    rzlst.append(item[0])

rongquan=csv.reader(open('./data/rongquan.csv', 'rU'))
rqlst=[]
for item in rongquan:
    rqlst.append(item[0])

#restructring
ivolsr={}
count=0
for item in ivols:
    id=item["stkcd"]+str(item['date'].year)+str(item['date'].month)
    ivolsr.update({str(id):{"year":item['date'].year,"month":item['date'].month
                       ,'ivol':item['ivol'],'stkcd':item["stkcd"]}})
    # print id
    # count=count+1
    # if count>500:
    #     break
atrributesstr={}
idlist=[]
for item in atrributes:
    id=item['stkcd']+str(item['year'])+str(item['mon'])
    id=id[0:10]+id[11:-1]
    if id not in idlist:
        idlist.append(id)
        atrributesstr.update({str(id):{'year':item['year'],'month':item['mon'],'Dturn': item['Dturn'],
                                  'ret': item['ret'], 'mktvalue': item['mktvalue'], 'stkcd': item['stkcd']}})
# i=0
# for item in atrributesstr:
#     print item,atrributesstr[item]
#     i=i+1
#     if i>50:
#         break
#
#start calculating
rzdata=[]
rzym=[]
for item in ivolsr:
    ym=str(ivolsr[item]['year'])+str(ivolsr[item]['month'])
    if ivolsr[item]['stkcd'] in rzlst and ym not in rzym:
        rzym.append(ym)

for ym in rzym:
    for item in ivolsr:
        if ym[0:4]==str(ivolsr[item]['year']) and ym[4:]==str(ivolsr[item]['month']) and ivolsr[item]['stkcd'] in rzlst:
            rzdata.append([item,ym,ivolsr[item]['ivol']])

# for i in rzdata:
#     print i

# rzym=['20081']
ranks={}
for ym in rzym:
    lst=[]
    rank=[]
    for i in rzdata:
        if i[1]==ym:
            lst.append({'stkcd':i[0][0:6],'ivol':i[2]})
    lst=sorted(lst,key=itemgetter("ivol"))
    for i in range(1,6,1):
        start=int(round(float(len(lst)*(i-1)/5),0))
        end=int(round(float(len(lst)*i/5),0))
        rank.append([d['stkcd'] for d in lst[start:end]])
    ranks.update({ym:rank})

atrbuts=['ym','ivolrank','ivol_lag',"v_weight_ilag","avg_return","v_weight_return","crate","crate_lag","v_weight_crate","v_weight_crate_lag","avg_totalv"]

with open("stkcds.csv","wb") as f1:
    for item in ranks:
        for i in range(0,5,1):
            f1.write(item+","+str(i)+",")
            for stk in ranks[item][i]:
                f1.write(stk+",")
            f1.write("\n")
f1.close()

print ranks['20081'][0]
ymlst=[]
for year in [2008,2009,2010,2013,2014]:
    for month in range(1,13,1):
        ym=str(year)+str(month)
        if ym in rzym:
            ymlst.append(ym)
#ivol_lag
ivollagdict={}
print ymlst
for ym in ymlst:
    for rank in range(0,5,1):
        lst=ranks[ym][rank]
        ivollaglst=[]
        for stk in lst:
            id=stk+str(ym)
            idpo=ymlst.index(ym)
            idlag=stk+ymlst[idpo-1]
            # print id,idlag
            if not ym=='20081' or '20132':
                ivollaglst.append(ivolsr[id]['ivol'])
        ivollag=sum(ivollaglst)/len(ivollaglst)
        ymrank=str(ym)+str(rank)
        ivollagdict.update({ymrank:ivollag})
print ivollagdict

ymranklst=[]

#ret
retdict={}
print ymlst
for ym in ymlst:
    if not str(ym)==('20081' or '20132'):
        for rank in range(0,5,1):
            lst=ranks[ym][rank]
            ret=[]
            for stk in lst:
                id=str(stk)+str(ym)
                # print id,idlag
                if not atrributesstr[id]['ret']=="":
                    ret.append(float(atrributesstr[id]['ret']))
            retavg=sum(ret)/len(ret)
            ymrank=str(ym)+str(rank)
            ymranklst.append(ymrank)
            retdict.update({ymrank:retavg})
print retdict

with open('result.csv','wb') as filere:
    for ymrank in ymranklst:
        filere.write(ymrank[0:-1]+','+ymrank[-1]+','+str(ivollagdict[ymrank])+','+str(retdict[ymrank])+'\n')
filere.close()