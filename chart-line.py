import openpyxl
from openpyxl import Workbook
from openpyxl.chart import Reference,LineChart
def set_line_chart(): 
    wb = Workbook()
    sh = wb.active
    sh.title = 'MySheet' 

    # 初始化数据
    rows = [['Date','Batch1','Batch2','Batch3'],
    [1,40,30,25],
    [2,43,49,34],
    [3,41,35,28],
    [4,37,33,41],
    [5,43,40,55],
    [6,47,34,29],
    [7,46,39,21],
    [8,39,42,35]
    ]
    for row in rows:
        sh.append(row)
    # 初始化折线图
    line = LineChart()
    line.title = '折线图'
    line.style = 2
    line.x_axis.title = '横坐标显示标题'
    line.y_axis.title = '纵坐标显示标题'
    data = Reference(sh,min_row=1,max_row=9,min_col=2,max_col=4)
    # titles_from_data是否启用标题
    line.add_data(data,titles_from_data=True)
    # 添加到图表中
    sh.add_chart(line,'I10')
    # 保存Excel 文件
    wb.save('./data/line-chart.xlsx')

if __name__ == '__main__':
    set_line_chart()
