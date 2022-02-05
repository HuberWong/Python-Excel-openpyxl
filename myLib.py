import datetime
import random
import openpyxl

"""
    写入每个sheet的指定日期指定内容
"""


def writeSheetDate(book, startTime):
    sheet = book.active
    timeDelta = startTime - datetime.datetime.strptime('2022-01-14', '%Y-%m-%d')
    # print('与01月14日时间差为：', timeDelta)
    for row in sheet.iter_rows(min_row=3, min_col=6):
        # for cell in row:
        #     print(cell.value)
        # print('------ End of Row ------')
        row[timeDelta.days].value = random.randint(2, 7) * 0.1 + 36
    # book.save('./TestFolder/output.xlsx')


"""
    分日起保存副本与Output文件夹中
    文件名：xx财xx班学生体温检测表-xx月xx日
"""


def saveFileByTimeToOutput(basalPlate, startTime, endTime):
    book = openpyxl.load_workbook('./BasalPlate/' + basalPlate + '.xlsx')
    # print(type(startTime.day))
    while (endTime - startTime).days >= 0:
        writeSheetDate(book, startTime)
        book.save('./Output/' + basalPlate + '-' + str(startTime.month) + '月' + str(startTime.day) + '日' + '.xlsx')
        print('文件'+'./Output/' + basalPlate + '-' + str(startTime.month) + '月' + str(startTime.day) + '日' + '.xlsx'+'构建完成')
        print(startTime)
        startTime = startTime + datetime.timedelta(1)
    return 1
