from sympy import symbols,Eq,solve,nsolve
import numpy as np
import pandas as pd
import xlwt

book = xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet = book.add_sheet('teu',cell_overwrite_ok=True)

p = 34000
tr = 14
wf = 6000*970
wice = 466060
vm = 1484
vf = 6000
vice = 2968
s = 300*45.6
vload = 37

sheet.write(0,1,'内燃机成本')
sheet.write(0,2,'电费成本')
sheet.write(0,3,'集装箱成本')
sheet.write(0,4,'电池成本')
sheet.write(0,5,'充电设施成本')
sheet.write(0,6,'EV总成本')

for i in range(22000):
    sheet.write(i+1,0,i+1)

for i in range(22000):
        #计算燃油船成本
    l = i+1

    ehfo =2 * p * l * (1 / vload) * 0.8#内燃机能量计算

        #n_co2 = e_hfo * 0.515
        #c_ice_co2 = n_co2 * 0.043 / l
        #c_ice_nox = e_hfo * 0.1157 / l
        #c_ice_so2 = e_hfo * 0.0576 / l  # 内燃机污染计算

    c_hfo_hfo = ehfo * 0.075 / l #0.048
    c_hfo_sfo = ehfo * 0.005 / l
    c_hfo_om = 64 * (7650/ 7650 ) * 0.035  #内燃机基础成本


    c_hfo = c_hfo_om  + c_hfo_hfo +c_hfo_sfo
    c_hfo_km = c_hfo #内燃机总成本


        #计算电池成本
    e,dt,wb,dteu,vb,g = symbols('e dt wb dteu vb g')
    eqs =  [Eq(e-1.10803*p*l*(1/vload)*0.8*pow(g,2/3)-0.22*p*l*(1/vload),0),#e-1.10803*pow(vt,3)*p*(100*i/vload)*(pow(g,(2/3)))-0.22*p*(100*i/vload)
            Eq(wb-wf-0.5*wice+dteu*28200-1026*dt*s,0),
            Eq(wb-e/(1.69*0.8),0),#0.26质量密度
            Eq(vf+vice-vb-vm-38.064*dteu,0),
            Eq(vb-e/(1200*0.76*0.8),0),#470体积密度
            Eq(g-(1+dt/tr),0)]
    X0=[300000,-1,30000,20,3000,1]
    res =  nsolve(eqs,[e,dt,wb,dteu,vb,g],X0)

        #基础费用
    c_ev_e = res[0] * 0.035 / 0.8  # 电费=能量*单价/放电深度
    c_ev_battery = res[0] * 1.25 * 100 / (175200 / (86.2 + (l / 37)))  # 电池费用，需要折旧
    c_ev_frame = res[0] * 1.25 * 0.029  # 充电设施
    c_ev_container = res[3] * (-104)  # 集装箱费用 根据文章的结果反推的
    c_ev_om = c_hfo_om / 2
    c_ev_so2_nox = res[0]*0.35*24*0.001/l+res[0]*13*0.47*0.001/l
    c_ev_co2  = res[0]*413*0.043*0.001/l



    c_ev = c_ev_container+c_ev_e+c_ev_battery+c_ev_frame
    c_ev_km = c_ev / l + c_ev_om


    #dc_hfo_ev = c_hfo_km - c_ev_km
    '''sheet.write(i+1,1, float(c_hfo_km))
    sheet.write(i+1,2, float(c_ev_e/l))
    sheet.write(i+1,3, float(c_ev_container/l))
    sheet.write(i+1,4, float(c_ev_battery/l))
    sheet.write(i+1,5, float(c_ev_frame/l))
    sheet.write(i+1,6, float(c_ev_km))'''

    '''sheet.write(0,1,'内燃机成本')
       sheet.write(0,2,'电费成本')
       sheet.write(0,3,'集装箱成本')
       sheet.write(0,4,'电池成本')
       sheet.write(0,5,'充电设施成本')
       sheet.write(0,6,'EV总成本')'''


'''savepath = '/Users/taohanqi/Desktop/Python/fig3.xls'
book.save(savepath)'''
