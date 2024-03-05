import xlrd
import xlwt

workbook = xlrd.open_workbook(r"C:\Users\zms\Desktop\气象数据\呼和浩特2.xls")
table = workbook.sheet_by_name("test")
new_workbook = xlwt.Workbook()
worksheet = new_workbook.add_sheet('test')
maxspring = -0.1
maxautumn = -0.1
maxsummer = -0.1
maxwinter = -0.1
year = 2013
summer = [0,1,2,3,4,5,6,7,8,9,10,11,12]
winter = [0,1,2,3,4,5,6,7,8,9,10,11,12]
summer0 = 182
winter0 = 366
maxhot = -50
hotdate = 0
colddate = 0
maxcold = 40
year0 = 0
for year0 in range(11):
    maxhot = table.cell_value(summer0, 10)
    hotdate = summer0
    maxcold = table.cell_value(winter0, 10)
    colddate = winter0
    for i in range(61):
        summer0 = summer0 + 1
        if(table.cell_value(summer0,10)>maxhot):
            maxhot = table.cell_value(summer0,10)
            hotdate = summer0
        winter0 = winter0 + 1
        if(winter0>4060):
            break
        if(table.cell_value(winter0,10) < maxcold):
            maxcold = table.cell_value(winter0, 10)
            colddate = winter0
    summer1 = hotdate - 4
    winter1 = colddate - 4
    maxhot = table.cell_value(summer1, 5)
    hotdate = summer1
    maxcold = table.cell_value(winter1, 5)
    colddate = winter1
    for i in range(4):
        summer1 = summer1 + 1
        if (table.cell_value(summer1, 5) > maxhot):
            maxhot = table.cell_value(summer1, 5)
            hotdate = summer1
        winter1 = winter1 + 1
        if (table.cell_value(winter1, 5) < maxcold):
            maxcold = table.cell_value(winter1, 5)
            colddate = winter1
    summer[year0] = hotdate
    winter[year0] = colddate
    print(summer[year0])
    print(winter[year0])
    print(year0)
    summer0 = summer0 + 308
    winter0 = winter0 + 308
    if(year0%4 == 3):
        summer0 = summer0 + 1
        winter0 = winter0 + 1

t=211
yearnumber = 0
maxtime = 0
while True:
    a = table.cell_value(t, 7)
    ave = table.cell_value(t, 8)
    time = table.cell_value(t,0)
    t = t + 1
    if((t+210)%365==0):
        yearnumber = yearnumber + 1
        print(maxtime)
        worksheet.write(yearnumber,0,maxtime)
        worksheet.write(yearnumber,1,maxwinter)
        worksheet.write(yearnumber,2,maxave)
        print(maxwinter)
        print(maxave)
        maxwinter = -0.1
    if (t == 3993):
        break
    if(ave>=10):
        continue
    elif(ave<10):
        if(a>maxwinter):
            maxwinter = a
            maxtime = time
            maxave = ave
yearnumber = 0
print("冬季输入完毕")
flag = True
t=31
i=0
while True:
    a = table.cell_value(t, 7)
    ave = table.cell_value(t, 8)
    time = table.cell_value(t,0)
    min = table.cell_value(t, 9)
    t = t + 1
    if((t-15)==summer[yearnumber]):
        yearnumber = yearnumber + 1
        print(maxtime)
        print(maxspring)
        print(maxave)
        worksheet.write(yearnumber, 4, maxtime)
        worksheet.write(yearnumber, 5, maxspring)
        worksheet.write(yearnumber, 6, maxave)
        maxspring = -0.1
        flag = False
    if(t-15==winter[yearnumber-1]):
        flag = True
        continue
    if (t == 3993):
        break
    if(ave<10):
        continue
    elif(min>=18):
        continue
    elif(flag == False):
        continue
    elif(ave>=10):
        if(a>maxspring):
            maxspring = a
            maxtime = time
            maxave = ave
print("春季输入完毕")
yearnumber = 0
flag = True
t=180
while True:
    a = table.cell_value(t, 7)
    ave = table.cell_value(t, 8)
    time = table.cell_value(t,0)
    min = table.cell_value(t,9)
    t = t + 1
    if((t-15)==winter[yearnumber]):
        yearnumber = yearnumber + 1
        print(maxtime)
        print(maxautumn)
        print(maxave)
        worksheet.write(yearnumber, 8, maxtime)
        worksheet.write(yearnumber, 9, maxautumn)
        worksheet.write(yearnumber, 10, maxave)
        maxautumn = -0.1
        flag = True
    if(t-15==summer[yearnumber]):
        flag = False
        continue
    if (t == 3993):
        break
    if(ave<10):
        continue
    elif(min>=18):
        continue
    elif(flag == True):
        continue
    elif(ave>=10):
        if(a>maxautumn):
            maxautumn = a
            maxtime = time
            maxave = ave
print("秋季输入完毕")
print(summer)
print(winter)
new_workbook.save(r"C:\Users\zms\Desktop\气象数据\呼和浩特3.xls")
