# -*- coding: UTF-8 -*-
from operator import itemgetter
from itertools import groupby
import ReadFromFile

def get_data_list_preparative(sheet_data):
    list_title = {}
    list_subproject, list_year_month, dict_statistics_data = ReadFromFile.read_statistics_data(sheet_data)
    dict_item_cost = {}
    table_used = ReadFromFile.read_preparative_data('预统计数据清单（正式账户）.csv', '数据清单（正式账户）.csv')
    table_used.sort(key=itemgetter('子产品名称'))
    for name, items_groupby_name in groupby(table_used, key=itemgetter('子产品名称')):
        #print(' ', name)#子产品名称
        item_cost = {} 
        list_subproject.append(name) 
        list_groupby_date = list(items_groupby_name)
        list_groupby_date.sort(key=itemgetter('结束使用时间'))
        for year_month, items_groupby_date in groupby(list_groupby_date, key=itemgetter('结束使用时间')):
            #print(year_month)#年月
            list_year_month.append(year_month)
            item_cost[year_month] = 0
            for item in items_groupby_date:
                item_cost[year_month] = item_cost[year_month] + float(item['现金账户支出(元)'])
        dict_item_cost[name] = item_cost
            #print(item_cost)#子产品每月花费
        #print(dict_item_cost)

    list_year_month = list(set(list_year_month))
    list_subproject = list(set(list_subproject))
    list_year_month.sort()
    list_subproject.sort()
    #print(dict_item_cost)
    #print(list_year_month)
    #print(list_name)
    #ReadFromFile.WriteToFile.write_statistics_data(sheet_data, list_subproject, list_year_month, dict_item_cost)
    return list_subproject, list_year_month, dict_item_cost


sheet_data = '正式账户'
list_subproject_statistics, list_year_month_statistics, dict_statistics_data = ReadFromFile.read_statistics_data(sheet_data)
list_subproject_preparative, list_year_month_preparative, dict_preparative_data = get_data_list_preparative(sheet_data)
list_subproject_preparative.extend(list_subproject_statistics)
list_subproject = list(set(list_subproject_preparative))
list_subproject.sort()
list_year_month_preparative.extend(list_year_month_statistics)
list_year_month = list(set(list_year_month_preparative))
list_year_month.sort()
for preparative_key,preparative_value in dict_preparative_data.items():
    print(preparative_key in dict_statistics_data.keys())
    if dict_statistics_data[preparative_key] is not None:
        print(preparative_value)
        for preparative_value_key,preparative_value_value in preparative_value.items():
            print(preparative_value_key in dict_statistics_data[preparative_key].keys())
##            if (preparative_value_key in dict_statistics_data[preparative_key].keys()):
##                if dict_statistics_data[preparative_key][preparative_value_key] is not None:
##            continue
##                else:
##                    dict_statistics_data[preparative_key][preparative_value_key] = preparative_value_value
##            if (preparative_value_key in dict_statistics_data[preparative_key].keys()):
##                dict_statistics_data[preparative_key][preparative_value_key] = preparative_value_value

##print(dict_statistics_data)
##ReadFromFile.WriteToFile.write_statistics_data(sheet_data, list_subproject, list_year_month, dict_statistics_data)
##     
##wb = openpyxl.Workbook()
##ws = wb.active
##
##rows = [
##    ['Number', 'Batch 1', 'Batch 2'],
##    [2, 30, 40],
##    [3, 25, 40],
##    [4 ,30, 50],
##    [5 ,10, 30],
##    [6,  5, 25],
##    [7 ,10, 50],
##]
##
##for row in rows:
##    ws.append(row)
##
##chart = openpyxl.chart.AreaChart3D()
##chart.title = "Area Chart"
##chart.style = 13
##chart.x_axis.title = 'Test'
##chart.y_axis.title = 'Percentage'
##chart.legend = None
##
##cats = openpyxl.chart.Reference(ws, min_col=1, min_row=1, max_row=7)
##data = openpyxl.chart.Reference(ws, min_col=2, min_row=1, max_col=3, max_row=7)
##chart.add_data(data, titles_from_data=True)
##chart.set_categories(cats)
##
##ws.add_chart(chart, "A10")
##
##wb.save("area3D.xlsx")
