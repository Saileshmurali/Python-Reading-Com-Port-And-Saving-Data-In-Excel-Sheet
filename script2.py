import serial
import numpy as np
import time
import datetime as dt
import xlsxwriter

workbook = xlsxwriter.Workbook('give path of excel sheet')
worksheet = workbook.add_worksheet("Timestamp data")
worksheet.write(0,0,"Timestamp")
worksheet.write(0,1,"Voltage")
ser = serial.Serial('COM8',9600)
ser.close()
ser.open()
i=0
tick=time.time()
tock=tick
w=tick-tock
while w<10:
    data=ser.readline()
    data=str(data.decode())
    volt=float(data)
    ct = dt.datetime.now()
    ct=str(ct)
    i=i+1
    worksheet.write(i,0,ct)
    worksheet.write(i,1,volt)
    tock=time.time()
    w=tock-tick
workbook.close()
