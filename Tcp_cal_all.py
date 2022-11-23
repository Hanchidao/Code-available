from sympy import symbols,Eq,solve,nsolve
import math
import numpy as np
import pandas as pd
import xlwt

book = xlwt.Workbook(encoding='utf-8',style_compression=0)
sheet = book.add_sheet('noname',cell_overwrite_ok=True)
df = pd.read_excel('/Users/taohanqi/Desktop/Python/data_teu_2.xlsx')

for i in range(221):
    sheet.write(0,i,(i+1)*100)

for j in range(16):
    if j<8:
        d2=0.26
        d=470
        phfo = 0.048
    else:
        d2=1.69
        d = 1200
        phfo = 0.075
    p = df.iloc[j, 1]
    tr = df.iloc[j, 2]
    wf = df.iloc[j, 3]
    wice = df.iloc[j, 4]
    vm = df.iloc[j, 5]
    vf = df.iloc[j, 6]
    vice = df.iloc[j, 7]
    s = df.iloc[j, 8]
    vmax= df.iloc[j, 9] * 1.85
    vload = df.iloc[j,12] * 1.85
    vt = vload/vmax




    for i in range(221):
        #计算燃油船成本
        l = 100 * (i + 1)

        ehfo = p * l * (1 / vload)* pow(vt,3)#内燃机能量计算

        #n_co2 = e_hfo * 0.515
        #c_ice_co2 = n_co2 * 0.043 / l
        #c_ice_nox = e_hfo * 0.1157 / l
        #c_ice_so2 = e_hfo * 0.0576 / l  # 内燃机污染计算

        c_hfo_hfo = ehfo * phfo / l
        c_hfo_sfo = ehfo * 5 / l
        c_hfo_om = 64 * (df.iloc[j,0] / 7650 ) * 0.035  #内燃机基础成本



        c_hfo = c_hfo_om + c_hfo_sfo + c_hfo_hfo
        c_hfo_km = c_hfo / l #内燃机总成本


        #计算电池成本
        e,dt,wb,dteu,vb,g = symbols('e dt wb dteu vb g')
        eqs =  [Eq(e-1.10803*p*l*(1/vload)*pow(vt,3)*pow(g,2/3)-0.22*p*l*(1/vload),0),#e-1.10803*pow(vt,3)*p*(100*i/vload)*(pow(g,(2/3)))-0.22*p*(100*i/vload)
                Eq(wb-wf-0.5*wice+dteu*28200-1026*dt*s,0),
                Eq(wb-e/(d2*0.8),0),
                Eq(vf+vice-vb-vm-38.064*dteu,0),
                Eq(vb-e/(d*0.76*0.8),0),
                Eq(g-(1+dt/tr),0)]
        X0=[300000,-1,30000,20,3000,1]
        res =  nsolve(eqs,[e,dt,wb,dteu,vb,g],X0)

        #基础费用
        c_ev_e = res[0] * 0.035 / 0.8  # 电费=能量*单价/放电深度
        c_ev_battery = res[0] * 1.25 * 50 / (175200 / (86.2 + (l / 37)))  # 电池费用，需要折旧
        c_ev_frame = res[0] * 1.25 * 0.021  # 充电设施
        c_ev_container = res[3] * (-104)  # 集装箱费用 根据文章的结果反推的
        c_ev_om = 0.5 * c_hfo_om



        c_ev = c_ev_container+c_ev_e+c_ev_battery+c_ev_frame
        c_ev_km = c_ev / l + c_ev_om

        dc_hfo_ev = c_hfo_km - c_ev_km


        #sheet.write(i+1 , j+1 ,res[3].df.iloc[j,0] ) # 记录集装箱体积变化

        #sheet.write(j, i, float(dc_hfo_ev))   # 记录成本差异


        # sheet.write(i + 1, 1, float(c_hfo_km))
        # sheet.write(i + 1, 2, float(c_ev_e / l))
        # sheet.write(i + 1, 3, float(c_ev_container / l))
        # sheet.write(i + 1, 4, float(c_ev_battery / l))
        # sheet.write(i + 1, 5, float(c_ev_frame / l))
        # sheet.write(i + 1, 6, float(c_ev_km))  #记录成本部分




savepath = '/Users/taohanqi/Desktop/Python/all_teu_tcp.xls'
book.save(savepath)




