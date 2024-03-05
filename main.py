import xlrd
import xlwt

workbook = xlrd.open_workbook(r"C:\Users\zms\Desktop\气象数据\长沙1.xls")
table = workbook.sheet_by_name("sheet")
t = 0
new_workbook = xlwt.Workbook()
worksheet = new_workbook.add_sheet('test')
worksheet.write(0,1,'2')
worksheet.write(0,2,'8')
worksheet.write(0,3,'14')
worksheet.write(0,4,'20')
while True:
    a = table.cell_value(t,0)
    b = table.cell_value(t,1)
    i = 0
    for text in a:
        i = i+1
        if(i == 1):
            day = text
        elif(i == 2):
            day = day +text
        elif (i == 4):
            month = text
        elif (i == 5):
            month =  month + text
        elif (i == 9):
            year = text
        elif (i == 10):
            year =  year + text
        elif (i == 12):
            hour = text
        elif (i == 13):
            hour =  hour + text
    year = int(year)
    month = int(month)
    day = int(day)
    hour = int(hour)
    if ((year==16 and month >= 3) or year>=17):
        day = day + 1
    if ((year == 20 and month >= 3) or year >= 21):
        day = day + 1
    year = year - 13
    day = day + 365*year
    if(month == 2):
        day = day + 31
    elif(month == 3):
        day = day + 59
    elif (month == 4):
        day = day + 90
    elif (month == 5):
        day = day + 120
    elif (month == 6):
        day = day + 151
    elif (month == 7):
        day = day + 181
    elif (month == 8):
        day = day + 212
    elif (month == 9):
        day = day + 243
    elif (month == 10):
        day = day + 273
    elif (month == 11):
        day = day + 304
    elif (month == 12):
        day = day + 334
    t = t+1
    if(hour == 2):
        worksheet.write(day,1,b)
    elif(hour == 8):
        worksheet.write(day, 2, b)
    elif(hour == 14):
        worksheet.write(day, 3, b)
    elif (hour == 20):
        worksheet.write(day, 4, b)
    if(t%100 == 0):
        print(t)
    if(t == 31133):
        break
print(t)
new_workbook.save(r"C:\Users\zms\Desktop\气象数据\长沙2.xls")
