import openpyxl
# 01 创建一个新的工作簿对象
wb = openpyxl.Workbook()
print(wb)
# 02 获取工作簿名称
# sheetname = wb.sheetnames
# print(sheetname)
# 03 给工作表设置名称
sheet = wb.active
sheet.title = "跟进记录"

# 04 保存工作簿
wb.save('./data/demo-test.xlsx')
