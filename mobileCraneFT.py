"""2020-7-14"""
"""行走塔机轮压计算"""
# 轮距m
a = 11
# 轨距m
b = 13.6
# 工作状态最大不平衡力矩t.m
mw = 1335
# 非工作状态最大不平衡力矩t.m
mn = 677
# 塔机自重t
gt = 365
# 行走机构自重t
gx = 90
# 最大吊重t
g = 70
# 计算最不利方向长度
l = a * b / ((a * a + b * b) ** 0.5)
print(f'{l=:.2f}')
"""工作状态轮压计算"""
fworkmax = (gt + gx + g) / 4 + mw / (2 * l)
fworkmin = (gt + gx + g) / 4 - mw / (2 * l)
print(f'工作状态最大轮压为{round(fworkmax, 1)}t')
print(f'工作状态最小轮压为{round(fworkmin, 1)}t')
# 稳定性校核
mst = (gt + gx + g) * min(a, b) / 2
print(f'工作状态稳定力矩{round(mst, 1)}, 不平衡力矩{mw}, 稳定系数{round(mst/mw, 1)}')
"""非工作状态轮压计算"""
fnworkmax = (gt + gx) / 4 + mn / (2 * l)
fnworkmin = (gt + gx) / 4 - mn / (2 * l)
print(f'非工作状态最大轮压为{round(fnworkmax, 1)}t')
print(f'非工作状态最小轮压为{round(fnworkmin, 1)}t')
# 稳定性校核
mnst = (gt + gx) * min(a, b) / 2
print(f'非工作状态稳定力矩{round(mnst, 1)}, 不平衡力矩{mn}, 稳定系数{round(mnst/mn, 1)}')
