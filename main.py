import xlwings as xw
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import argparse
from Libs import data_process as dp

# 创建一个ArgumentParser对象 
parser = argparse.ArgumentParser(description='Process some files.') 
# 添加一个参数，用于指定文件名 
parser.add_argument('filename', help='the name of the file to process') 
# parser.add_argument('opt', help='option') 
# 解析命令行参数 
args = parser.parse_args()  

# 使用指定的文件名来操作文件 
# with open(args.filename, 'r') as f:  

path = r'./'+args.filename
# print(path)

app = xw.App(visible=True, add_book=False) # 程序可见，只打开不新建工作薄
app.display_alerts = False # 警告关闭
app.screen_updating = False # 屏幕更新关闭
wb = app.books.open(path)
# print(wb)

sheet = wb.sheets.active

shape = sheet.used_range.shape
""" print(shape)
print(sheet.used_range.row)
print(sheet.used_range.column)
print(sheet.range(sheet.used_range.row,sheet.used_range.column).value)
print(sheet.used_range.last_cell.value)
print(sheet.used_range.rows.count)

for i in range(sheet.used_range.column, sheet.used_range.column+shape[1]):
    print(sheet.range(sheet.used_range.row,i).value) """
fst_row = sheet.range((sheet.used_range.row,sheet.used_range.column),
                  (sheet.used_range.row,sheet.used_range.column+shape[1]-1))

""" print(sheet.range((sheet.used_range.row,sheet.used_range.column),
                  (sheet.used_range.row+shape[0]-1,sheet.used_range.column)).value)

clm1 = sheet.range((sheet.used_range.row,sheet.used_range.column),
                  (sheet.used_range.row+shape[0]-1,sheet.used_range.column)).value

clm2 = sheet.range(sheet.used_range.row,sheet.used_range.column).expand('down') """
tbl = sheet.range(sheet.used_range.row,sheet.used_range.column).expand('table')
my_pd = pd.DataFrame(tbl.value)
# print(sheet.used_range.column, shape[1], my_pd.shape[1])

wb.save() # 保存文件
wb.close() # 关闭文件
app.quit() # 关闭程序
""" my_pd2 = my_pd.iloc[1:,1:]
my_pd2.columns = ['zhangsan','lisi','wangwu','zhaoliu','sunqi']
print(my_pd2)
my_pd2.plot(kind='line')
plt.show() """

max_sum = max_col = min_col = listidx = 0
min_sum = 999999999
mlist = []
for col in range(1,shape[1]):
    tmp_rtn = dp.data_wash(my_pd,col)
    mlist.append(tmp_rtn)
    listidx = listidx + 1
    if tmp_rtn > max_sum:
        max_sum = tmp_rtn
        max_col = col
    if tmp_rtn < min_sum:
        min_sum = tmp_rtn
        min_col = col
    
    my_pd2 = my_pd.iloc[1:,col]
    # print(my_pd2)
    clr = ['red','blue','yellow','green','pink']
    my_pd2.plot(kind='bar',title=col,color=clr[col-1])
    plt.show(block=False) #block=False保证窗口关闭
    plt.pause(2)
    plt.close('all')

print('max name %s, min name %s' 
      % (my_pd.loc[0,max_col],my_pd.loc[0,min_col]))

for i in range(1):
    my_pd2 = my_pd.iloc[1:,[max_col,min_col]]
    my_pd2.columns = ['max','min']
    # print(my_pd2)
    my_pd2.plot(kind='line',title='max')
    plt.show(block=False) #block=False保证窗口关闭
    plt.pause(2)
    plt.close('all')

    my_pd2 = my_pd.iloc[1:,[min_col,max_col]]
    my_pd2.columns = ['min','max']
    # print(my_pd2)
    my_pd2.plot(kind='line',title='min')
    plt.show(block=False) #block=False保证窗口关闭
    plt.pause(2)
    plt.close('all')

""" print(mlist, max(mlist), mlist.index(max(mlist)), min(mlist), mlist.index(min(mlist)))
mlist.sort()
nlist = mlist
print(mlist, nlist)

my_pd1 = pd.DataFrame({6:['魏八',1,2,3,4],7:['陈九',5,6,7,8]})
my_pd = pd.concat([my_pd,my_pd1],axis=1)
print(my_pd1)
print(my_pd) """

""" sum = my_pd.loc[1:,1].sum()
avg = my_pd.loc[1:,1].mean()
min = my_pd.loc[1:,1].min()
min_idx = my_pd.loc[1:,1].idxmin()
max = my_pd.loc[1:,1].max()
max_idx = my_pd.loc[1:,1].idxmax()
# print(my_pd.loc[1:,1],sum,avg,max,max_idx, min, min_idx)
print('sum: %d, avg: %f, max: %d, idx: %d, min: %d, idx: %d' % (sum,avg,max,max_idx, min, min_idx)) """

# print(fst_row.value)
""" for itm in fst_row:
    if itm.value == '李四':
        clm2 = itm.offset(1,0).expand('down')
        break """


# print(clm1[1].year, clm1[1].month)

