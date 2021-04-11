import serial
import matplotlib.pyplot as plt
import numpy as np
import time
from scipy.fft import fft
import matplotlib.animation as animation
import datetime as dt
import xlsxwriter

workbook = xlsxwriter.Workbook('D:\\111SaveFiles\\python\\Virtual_Com/Example1.xlsx')
worksheet = workbook.add_worksheet("Timestamp data")
inpt=workbook.add_worksheet("Interruption data")
worksheet.write(0,0,"Timestamp")
worksheet.write(0,1,"Voltage")
worksheet.write(0,2,"Current")
ser = serial.Serial('COM8',9600)
ser.close()
ser.open()
i=0
j=0
x=[]
y=[]
tick=time.time()
tock=tick
w=tick-tock
while w<1:
    data=ser.readline()
    data=str(data.decode())
    p=data.split()
    volt=p[0]
    volt=float(volt)
    b=p[1]
    Iamp=float(b)
    ct = dt.datetime.now()
    ct=str(ct)
    #data=int(data)
    #print(type(data))
    #print(ct,a,b)
    i=i+1
    worksheet.write(i,0,ct)
    worksheet.write(i,1,volt)
    worksheet.write(i,2,Iamp)
    tock=time.time()
    w=tock-tick
workbook.close()
