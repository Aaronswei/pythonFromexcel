#-*-coding:utf-8-*-

import os
import sys
import xlwt
import xlrd

holidays = []
comm_year = [31,28,31,30,31,30,31,31,30,31,30,31]
leap_year = [31,29,31,30,31,30,31,31,30,31,30,31]

curpath = os.getcwd()

except_members = ['郭健','殷宏杰']


def whether_workdays(datetime):
    ret = 0
    datetime_year = datetime.split("年")
    year = int(datetime_year[0])
    datetime_month = datetime_year[1].split("月")
    month = int(datetime_month[0])
    date = int(datetime_month[1].split("日")[0])

    if (year % 100 == 0 or year % 400 == 0 or year % 4 == 0):
        for i in range(month - 1):
            ret = ret + leap_year[i]
    else:
        for i in range(month - 1):
            ret = ret + comm_year[i]

    S = (year + (year-1) // 4 - (year-1) // 100 + (year-1) // 400) % 7
    days = ((ret + date) + S - 1) % 7

    if days == 0 or days == 6:
        return 0
    else:
        return 1
    

def timeToSecond(nowTime):
    time = nowTime.strip().split('外勤')[0]

    hour, minute = time.strip().split(":")
    return int(hour) * 3600 + int(minute) * 60

def secondTotime(nowTime):
    hour = int(nowTime / 3600)
    minute = int((nowTime - hour * 3600) / 60)
    second = nowTime - hour * 3600 - minute * 60

    return hour, minute, second


def set_style(name, height, bold=False):
    style = xlwt.XFStyle()

    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height

    alignment = xlwt.Alignment()
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER

    style.font = font
    style.alignment = alignment

    return style    

def readAndWrite_excel(year, month, excel_filename, out_filename):
    workbook = xlrd.open_workbook(excel_filename)
    outbook = xlwt.Workbook(style_compression=2)
    mingxi = outbook.add_sheet('sheet1')
    print(workbook.sheet_names())

    sheet3 = workbook.sheet_by_index(3)
    print(sheet3.name, sheet3.nrows, sheet3.ncols)
    sheet1 = workbook.sheet_by_index(1)

    mingxi.write_merge(2, 4, 1, 1, "日期", set_style('Arial',220,True))

    for j in range(5, sheet3.ncols):
        datetime = month + "/"+ str(j-4) + "/" + year
        mingxi.write(j, 1, str(datetime), set_style('Arial',220,True))

    temp_rows = 0
    if sheet3.nrows >= 50:
        temp_rows = 50
    elif sheet3.nrows < 50:
        temp_rows = sheet3.nrows
    
    for i in range(3, temp_rows):
        name = sheet3.cell(i, 0).value

        mingxi.write_merge(2, 2, (i-3)*5 + 2, (i-3)*5 + 6, name, set_style('Arial',220,True))
        

        mingxi.write_merge(3, 4, (i-3)*5 + 2, (i-3)*5 + 2, "上班", set_style('Arial',220,True))
        mingxi.write_merge(3, 4, (i-3)*5 + 3, (i-3)*5 + 3, "下班", set_style('Arial',220,True))
        mingxi.write_merge(3, 4, (i-3)*5 + 6, (i-3)*5 + 6, "备注", set_style('Arial',220,True))
        mingxi.write_merge(3, 3, (i-3)*5 + 4, (i-3)*5 + 5,"当日在岗", set_style('Arial',220,True))
        mingxi.write(4, (i-3)*5 + 4, "分钟", set_style('Arial',220,True))
        mingxi.write(4, (i-3)*5 + 5, "小时", set_style('Arial',220,True))
        print("========================")
        print(name)
        if name in except_members:
            print("不参加考勤记录")
        else:
            if len(name.split("（")) == 2:
#                print("员工离职，考勤另外记录")
                pass
            else:
                for j in range(5, sheet3.ncols):
                    datetime = year + "年" + month + "月"+ str(j-4) + "日"
                    timevalue = sheet3.cell(i, j).value
                    timelong = timevalue.split('\n')
                    if len(timelong) > 1:
                        starttime = timelong[0]
                        endtime = timelong[-1]
                        mingxi.write(j, (i-3)*5 + 2, starttime, set_style('Arial',220))
                        mingxi.write(j, (i-3)*5 + 3, endtime, set_style('Arial',220))
                        worktime = timeToSecond(endtime) - timeToSecond(starttime)
                        workhour, workmin, _ = secondTotime(worktime)

#                        print(datetime, " : ", starttime, endtime)
#                        print("Total time %02d:%02d"%(workhour, workmin))
                        mintime = str(workhour) + "时" + str(workmin) + "分"
                        hourtime = round(worktime / 3600, 2)
                        mingxi.write(j, (i-3)*5 + 4, mintime, set_style('Arial',220))
                        mingxi.write(j, (i-3)*5 + 5, hourtime, set_style('Arial',220))

                    elif len(timelong) == 1:
                        if whether_workdays(datetime):
                            mingxi.write(j, (i-3)*5 + 6, "An Error Here", set_style('Arial',220))
                        else:
                            print(datetime, ": this day is weekends")

    outbook.save(out_filename)


if __name__ == '__main__':
    year = '2018'
    month = '11'
    input_file = '苏州安智汽车零部件有限公司_考勤报表_20181101-20181130.xlsx'
    output_file = '苏州安智11月考勤表_1.csv'

    excel_filename = os.path.join(curpath, input_file)
    out_filename = os.path.join(curpath, output_file)
    readAndWrite_excel(year, month, excel_filename, out_filename)
