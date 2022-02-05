import myLib
import datetime

startTime = datetime.datetime.strptime('2022-01-14', '%Y-%m-%d')
endTime = datetime.datetime.strptime('2022-02-28', '%Y-%m-%d')
basalPlates = ['20财14班学生体温检测表', '20财15班学生体温检测表', '21财8班学生体温检测表', '21财9班学生体温检测表']
for basalPlate in basalPlates:
    myLib.saveFileByTimeToOutput(basalPlate, startTime, endTime)
time = datetime.datetime.strptime('2022-01-01', '%Y-%m-%d')

# for i in range(1,33):
#     time = time + datetime.timedelta(1)
#     print(time)

# print(startTime - endTime)