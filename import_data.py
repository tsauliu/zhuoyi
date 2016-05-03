from openpyxl import load_workbook
ivols=[]

for sheet in range(1,5,1):
    count=0
    wb = load_workbook(filename='./data/ivol'+str(sheet)+'.xlsx', read_only=True)
    ws = wb['Sheet1'] # ws is now an IterableWorksheet
    for row in ws.rows:
        if count >=1:
            ivols.append([row[0].value,row[1].value,row[2].value,row[3].value,row[4].value])
        count+=1
        # if i>10:
        #     break
for list in ivols:
    print list
