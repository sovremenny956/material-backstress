import pandas as pd
# import openpyxl
import numpy as np
import matplotlib.pyplot as plt
# import matplotlib
# import scipy.signal
import xlwt
import math
test = ('C:\\Users\\Lenovo\\Desktop\\实验\\包申格\\origin.xlsx')
do_test = pd.read_excel(test)

number_original = do_test['Column1'].values
strain_original = do_test['Column2'].values  #4
stress_original = do_test['Column3'].values  #8

number1 = number_original
strain1 = np.array([100*(math.log((1+strain_original[i]/100), math.e)) for i in range(0, len(strain_original))])
stress1 = np.array([stress_original[i]*(1+strain_original[i]/100) for i in range(0, len(stress_original))])

number = [number1[x] for x in range(0, len(number1))]
strain = [strain1[x] for x in range(0, len(strain1))]
stress = [stress1[x] for x in range(0, len(stress1))]

max_str = 0.00
peak_strain = []
min_str = 100.00
valley_strain = []
i = 0


while i < len(number):
    if strain[i] > max_str:
        max_str = strain[i]
    else:
        pass

    if strain[i] < min_str:
        min_str = strain[i]
    else:
        pass
    if abs(strain[i]-min_str) > abs(0.1):
        valley_strain.append(strain.index(min_str))
        valley_strain = list(set(valley_strain))
    else:
        pass
    if abs(max_str-strain[i]) > abs(0.1):
        if not strain.index(max_str) in peak_strain:
            peak_strain.append(strain.index(max_str))
            # peak_strain=list(set(peak_strain))
            min_str = max_str
    else:
        pass
    i += 1
peaks = peak_strain
peaks.sort()
valleys = valley_strain
valleys.sort()

plt.figure(figsize=(10., 3.))
plt.scatter(number, strain, s=0.1, c='b')
#plt.scatter(number1[peak_strain], strain1[peak_strain], s=0.3, c='r')
#plt.scatter(number1[valley_strain], strain1[valley_strain], s=0.3, c='k')
plt.savefig('strain_peak_valley.png', dpi=600)
plt.clf()


# 切片圆环

def rings():
    i = 0
    group_ring = []
    while i < len(peaks)-1:
        group_ring.append(stress1[peaks[i]:peaks[i+1]])
        i += 1
    return list(group_ring)


def rings1():
    j = 1
    group_rings1 = []
    while j < len(valleys)-1:
        group_rings1.append(stress1[valleys[j]:valleys[j+1]])
        j += 1
    return group_rings1


# 找出拟合点的位置
def back_stress(a):
    j = 0
    back_stress2 = []
    while j < len(rings()):
        e = float(rings()[j][0])*a
        f = peaks[j]
        while stress1[f] > e:
            f += 1
        back_stress2.append(f)
        j += 1
    return back_stress2


def back_stress_valleys(a):
    j = 0
    back_stress1 = []
    while j < len((rings())):
        e = float(rings()[j][0])*a
        f = valleys[j+1]
        while stress1[f] < e:
            f += 1
        back_stress1.append(f)
        j += 1
    return back_stress1


def line_fitting(a, b, color1, color2):   # 拟合直线，偏移，求交点
    q = 0
    x_p = []
    x_v = []
    y_p = []
    y_v = []
    y_up = []
    x_up = []
    y_down = []
    x_down = []
    strain_backstress = []
    stress_backstress = []
    backstress_download = []
    backstress_load = []
    while q < len(back_stress(a)):

        x_p_1 = [strain1[p] for p in range(peaks[q], back_stress(a)[q])]
        # 按应变峰值将真应力应变曲线分段，选出应变峰值所对应的应力到a*应变峰值所对应的应力之间的所有点
        y_p_1 = [stress1[p] for p in range(peaks[q], back_stress(a)[q])]
        x_v_1 = [strain1[p] for p in range(valleys[q+1], back_stress_valleys(b)[q])]
        # 按应变谷值将真应力应变曲线分段
        y_v_1 = [stress1[p] for p in range(valleys[q+1], back_stress_valleys(b)[q])]
        x_s_up = [strain1[p] for p in range(valleys[q+1], valleys[q+1]+int(0.6*(peaks[q+1]-valleys[q+1])))]
        # 截取出加载段，即应变谷值和后一个应变峰值之间的点，因存在塑性变形区域（应力平台）因此取0.6乘总点数
        y_s_up = [stress1[p] for p in range(valleys[q+1], valleys[q+1]+int(0.6*(peaks[q+1]-valleys[q+1])))]
        x_s_down = [strain1[p] for p in range(peaks[q], valleys[q+1])]  # 截取出卸载段，即应变峰值和后一个应变谷值之间的点
        y_s_down = [stress1[p] for p in range(peaks[q], valleys[q+1])]

        # plt.scatter(x,y,c='r')
   
        parameter_p = np.polyfit(x_p_1, y_p_1, 1)  # 拟合卸载弹性段（直线）的参数
        parameter_v = np.polyfit(x_v_1, y_v_1, 1)
        parameter_s_down = np.polyfit(x_s_down, y_s_down, 3)  # 拟合卸载曲线的参数
        parameter_s_up = np.polyfit(x_s_up, y_s_up, 3)  # 拟合加载曲线的参数
        # print(np.poly1d(parameter), '\n')
        # 四条拟合曲线，分别为：y_p_line:卸载段弹性区域拟合直线，y_v_line:加载段弹性区域拟合直线
        # y_up_line:加载段拉伸曲线拟合曲线,y_down_line:卸载段拉伸曲线拟合曲线
        y_p_line = np.poly1d(parameter_p)+0.1*parameter_p[0]  # 拟合卸载弹性段（直线）偏移
        y_v_line = np.poly1d(parameter_v)-0.1*parameter_v[0]
        y_up_line = np.poly1d(parameter_s_up)  # 拟合加载曲线
        y_down_line = np.poly1d(parameter_s_down)  # 拟合卸载曲线
        plt.figure(figsize=(3., 5.))
        for x1 in range(0, len(x_s_down)):
            y_p.append(y_p_line(x_s_down[x1]))
            y_down.append(y_down_line(x_s_down[x1]))
        x_down = x_s_down
        plt.plot(x_down, y_down, linewidth=0.2, color='b')
        y_down = []
        x_p = x_s_down
        plt.plot(x_p, y_p, linewidth=0.2, c=color1)
        y_p = []
        root_down = np.roots(y_p_line-y_down_line)
        # 输出卸载段直线与卸载曲线的交点
        for x in range(0, len(root_down)):
            if min(x_s_down) < root_down[x] < max(x_s_down):
                backstress_download.append(round(abs(y_p_line(root_down[x])), 4))
        # 结尾
        for x1 in range(0, len(x_s_up)):
            y_v.append(y_v_line(x_s_up[x1]))
            y_up.append(y_up_line(x_s_up[x1]))
        x_v = x_s_up
        # plt.plot(x_v, y_v, linewidth=0.2, c=color2)
        y_v = []
        x_up = x_s_up
        plt.plot(x_up, y_up, linewidth=0.2, color='b')
        y_up = []
        root_up = np.roots(y_v_line-y_up_line)
        for x in range(0, len(root_up)):
            if min(x_s_down) < root_up[x] < max(x_s_down):
                backstress_load.append(round(abs(y_v_line(root_up[x])), 4))
        # 结尾
        # x_p += x_p_1
        # x_v += x_v_1
        strain_backstress.append(round(abs(strain1[peaks[q]]), 5))
        stress_backstress.append(round(abs(stress1[peaks[q]]), 5))
        plt.savefig(f'rings{q}.png', dpi=600)
        plt.clf()
        # plt.xticks(ticks = x_s_up)
        # x_up += x_s_up
        # x_down += x_s_down
        q += 1

    # plt.plot(x_up,y_up, linewidth=0.2, c='b')
    # plt.plot(x_down, y_down, linewidth=0.2, c='b')
    # plt.show()
    # plt.savefig('abc.png', dpi=600)
    print(len(rings1()))
    load = [stress[x] for x in peaks]
    print('加载段的背应力列表为', load)
    print('卸载段的背应力列表为', backstress_download)
    finish = [(-backstress_download[x]+load[x])/2 for x in range(0, len(backstress_download))]
    print(finish)

    excel = xlwt.Workbook('encoding = utf-8')  # 设置工作簿编码
    sheet1 = excel.add_sheet('sheet1', cell_overwrite_ok=True)  # 创建sheet工作表
    for i in range(0, len(strain_backstress)):
        sheet1.write(i, 0, strain_backstress[i])  # 写入数据参数对应 行, 列, 值
    for i in range(0, len(load)):
        sheet1.write(i, 1, load[i])
    for i in range(0, len(finish)):
        sheet1.write(i, 4, finish[i])
    excel.save('D://cjy//origin.xls')  # 保存.xls到当前工作目录

line_fitting(0.4, 0.7, color1='k', color2='y')
# line_fitting(0.7, 0.5, color1='r', color2='g')
# code用法
# save(obj, file)
