---
title: "Python多级表头处理"
date: 2024-07-31T15:55:23+08:00
lastmod: 2024-07-31T15:55:23+08:00
author: ["Cyan Feather"]
keywords: 
- 合并单元格
categories: 
- Redis
tags: 
- Excel
- Python
description: "Python多级表头处理"
weight:
slug: ""
draft: false # 是否为草稿
comments: true
reward: true # 打赏
mermaid: false # 是否开启mermaid
showToc: true # 显示目录
TocOpen: true # 自动展开目录
hidemeta: false # 是否隐藏文章的元信息，如发布日期、作者等
disableShare: true # 底部不显示分享栏
showbreadcrumbs: true # 顶部显示路径
cover:
    image: "" # 图片路径：posts/tech/123/123.png
    caption: "" # 图片底部描述
    alt: ""
    relative: false
---
# Python多级表头处理

功能：将合并单元格拆分并填充原数据

```python
import xlwings as xw
 
# 启动Excel应用程序


with xw.App(visible=True,add_book=False) as app: # 这样写就不用再写app.kill()了，会自动关闭Excel软件
    # wb = xw.Book(r'C:\Users\青羽\Desktop\work\expy\1.xlsx')
    wb = app.books.open(r'C:\Users\青羽\Desktop\work\expy\1.xlsx')

    # 选择工作表
    sheet = wb.sheets['Sheet2']
    
    # 定位多级表头范围
    range = sheet.used_range

    headet_rows = 0
    headet_cols = len(range.columns)

    color1 = sheet.range('A1').color
    for col in range.columns[0]:
        if col.color == color1:
            headet_rows = headet_rows+1
            continue;
        else:
            break;
    
    print(headet_rows)
    print(headet_cols)

    headet_range = sheet[:headet_rows,:headet_cols]

    for i in headet_range:
        v = i.value
        print(v)
        a = i.merge_area
        r1 = sheet.range(a.get_address(False, False))
        c = a.count
        print(c)
        if c > 1:
            a.unmerge()
            r1.value = v
        print('1')

    wb.save()
```

xlwings对Linux支持不友好，Linux中可使用openpyxl

```python
from openpyxl import load_workbook
 
def header_handle(filename1, filename2):
 
    # 加载Excel文件
    wb = load_workbook(filename1)
    
    # 选择工作簿中的工作表
    ws = wb['Sheet2']  # 或者使用 wb.get_sheet_by_name('Sheet1')

    header_rows = 0
    header_cols = ws.max_column

    cell = ws['A1']
    bg_color = cell.fill.fgColor.rgb
    for each_cell in ws['A']:
        print(any(each_cell.coordinate in merged for merged in ws.merged_cells.ranges))
        if each_cell.fill.fgColor.rgb == bg_color or any(each_cell.coordinate in merged for merged in ws.merged_cells.ranges):
            header_rows = header_rows+1
            continue;
        else:
            break;

    merge_cell_list = [x.bounds for x in ws.merged_cells]

    # 复制合并单元格的值并取消合并
    for merged_cell in merge_cell_list:
        # 获取合并单元格的起始和结束坐标 
        # col_1, row_1, col_2, row_2对就(2, 4, 3, 4) 即 B4:C4
        col_1, row_1, col_2, row_2 = merged_cell
        
        # 读取合并单元格的值
        merge_cell_value = ws.cell(row=row_1, column=col_1).value
        merge_cell_style = ws.cell(row=row_1, column=col_1).style
        
        # 取消合并
        ws.unmerge_cells(start_row=row_1, start_column=col_1, end_row=row_2, end_column=col_2)
        
        # 将原合并单元格的值填入新的单元格
        for row in range(row_1, row_2 + 1):
            for col in range(col_1, col_2 + 1):
                ws.cell(row=row, column=col, value=merge_cell_value)
                ws.cell(row=row, column=col, value=merge_cell_value).style = merge_cell_style
    
    # 保存工作簿
    wb.save(filename2)
    return header_rows

```
