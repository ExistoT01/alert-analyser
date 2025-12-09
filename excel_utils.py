import zipfile
import os
import pandas as pd
import xlwings as xw
from datetime import datetime, timedelta

import get_config
from path_utils import PathUtils
from logger import Logger


path_utils = PathUtils()
logger = Logger(True, path_utils.log_path)


'''
将列号转换为 Excel 列字母
'''
def col_letter(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


class ExcelUtils:
    def __init__(self):
        super().__init__()
        self.cities = get_config.read_cfg()['cities']

    # temp func
    def unzip(self, file_name, day):
        logger.log(f'正在解压文件：{file_name}')

        with zipfile.ZipFile(os.path.join(path_utils.zips_path, file_name)) as f:
            target_path = os.path.join(path_utils.data_path, day.strftime("%m-%d"))
            os.makedirs(target_path, exist_ok=True)

            # if path is not empty, clear all outdated files
            for file_name in os.listdir(target_path):
                file_path = os.path.join(target_path, file_name)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception as e:
                    logger.log(f'删除 {file_path} 失败！原因：{e}')
                    return

            f.extractall(target_path)

        logger.log("文件解压完毕")


    # def unzip(self, file_name):
        # logger.log(f'正在解压文件：{file_name}')

        # with zipfile.ZipFile(os.path.join(path_utils.zips_path, file_name)) as f:
        #     target_path = os.path.join(path_utils.data_path, datetime.today().strftime("%m-%d"))
        #     os.makedirs(target_path, exist_ok=True)

        #     # if path is not empty, clear all outdated files
        #     for file_name in os.listdir(target_path):
        #         file_path = os.path.join(target_path, file_name)
        #         try:
        #             if os.path.isfile(file_path):
        #                 os.remove(file_path)
        #         except Exception as e:
        #             logger.log(f'删除 {file_path} 失败！原因：{e}')

        #     f.extractall(target_path)

        # logger.log("文件解压完毕")


    # test func
    def concat(self, files, day):
        logger.log("开始合并Excel文件")

        total = len(files)
        data_frames = []
        for i, file in enumerate(files):
            logger.log(f"正在读取第 {i + 1}/{total} 个Excel文件")

            df = pd.read_excel(file)
            data_frames.append(df)

        logger.log("读取完毕，开始合并")
        merged_df = pd.concat(data_frames, ignore_index=True)


        # “网络类型”列剔除 “一干”
        # “告警标题”列剔除 “GPON ONT掉电(DGi)”
        # “专业”列剔除 5GC、VIMS、骨干云池、固网、信令网、增值平台、MEC、业务平台
        drop_prof = ["5GC", "VIMS", "骨干云池", "固网", "信令网", "增值平台", "MEC", "业务平台"]

        filtered_df = merged_df[
            (merged_df["网络类型"].fillna("") != "一干") &
            (merged_df["告警标题"].fillna("") != "GPON ONT掉电(DGi)") &
            (~merged_df["专业"].fillna("").isin(drop_prof))
        ]

        target_path = os.path.join(path_utils.data_path, day.strftime("%m-%d"))
        os.makedirs(target_path, exist_ok=True)

        file_name = "merged_" + day.strftime("%m_%d") + ".xlsx"
        filtered_df.to_excel(os.path.join(target_path, file_name), index=False)

        logger.log("Excel文件合并完毕")
        return os.path.join(target_path, file_name)


    # def concat(self, files):
        # logger.log("开始合并Excel文件")
 
        # total = len(files)
        # data_frames = []
        # for i, file in enumerate(files):
        #     logger.log(f"正在读取第 {i + 1}/{total} 个Excel文件")

        #     df = pd.read_excel(file)
        #     data_frames.append(df)

        # logger.log("读取完毕，开始合并")
        # merged_df = pd.concat(data_frames, ignore_index=True)
        # target_path = os.path.join(path_utils.data_path, datetime.today().strftime("%m-%d"))
        # os.makedirs(target_path, exist_ok=True)

        # file_name = "merged_" + datetime.today().strftime("%m_%d") + ".xlsx"
        # merged_df.to_excel(os.path.join(target_path, file_name), index=False)

        # logger.log("Excel文件合并完毕")
        # return os.path.join(target_path, file_name)


    # test func
    # def gen_pivot_table(self, file_path, day):
    #     logger.log("开始生成数据透视表")

    #     wb = xw.Book(file_path)
    #     ws_data = wb.sheets['Sheet1']

    #     yesterday = day - timedelta(days=1)
    #     ws_data.name = f"{yesterday.strftime("%#m月%#d日")}告警"

    #     data_range = ws_data.range("A1").expand()

    #     ws_pivot = wb.sheets.add(before=ws_data)
    #     ws_pivot.name = f"{yesterday.strftime("%#m月%#d日")}分析"

    #     source_str = f"{ws_data.name}!{data_range.address.replace('$', '')}"
    #     pivot_cache = wb.api.PivotCaches().Create(SourceType=1, SourceData=source_str)

    #     pivot_table = pivot_cache.CreatePivotTable(
    #         TableDestination=ws_pivot.range("A1").api,
    #         TableName="PivotTable1"
    #     )

    #     # 获取透视表字段集合
    #     pivot_fields = pivot_table.PivotFields()

    #     # 设置行字段和列字段
    #     # 1）将“（资源）地市名称”设置为行字段（xlRowField=1）
    #     row_field = pivot_fields.Item("（资源）地市名称")
    #     row_field.Orientation = 1  # xlRowField

    #     # 2）将“专业”设置为列字段（xlColumnField=2）
    #     col_field = pivot_fields.Item("专业")
    #     col_field.Orientation = 2  # xlColumnField

    #     # 3）将“（资源）地市名称”设置为数据字段并做计数（-4112=xlCount）
    #     pivot_table.AddDataField(
    #         pivot_fields.Item("告警标题"),   # 计数哪个字段
    #         "计数项:（资源）地市名称",         # 在透视表中显示的名字
    #         -4112                             # xlCount
    #     )

    #     # -------------------------------
    #     # 筛选：只允许甘肃省的城市出现
    #     # -------------------------------
    #     # 遍历 row_field 的所有 PivotItems
    #     for item in row_field.PivotItems():
    #         if item.Name not in self.cities:
    #             item.Visible = False

    #     # -------------------------------
    #     # 筛选：隐藏列字段中不需要的项
    #     # -------------------------------
    #     hide_list = ["5GC", "VIMS", "骨干云池", "固网", "信令网", "增值平台", "MEC", "业务平台"]
    #     for item in col_field.PivotItems():
    #         if item.Name in hide_list:
    #             item.Visible = False

    #     target_path = os.path.join(path_utils.data_path, day.strftime("%m-%d"), "new_pivot.xlsx")
    #     wb.save(target_path)
    #     wb.close()

    #     logger.log("数据透视表生成完毕")
    #     return target_path


    # test func v2
    def gen_pivot_table(self, file_path, day):
        logger.log("开始生成数据透视表")

        wb = xw.Book(file_path)
        ws_data = wb.sheets['Sheet1']

        yesterday = day - timedelta(days=1)
        ws_data.name = f"{yesterday.strftime("%#m月%#d日")}告警"

        # data_range = ws_data.range("A1").expand()
        data_range = ws_data.used_range

        ws_pivot = wb.sheets.add(before=ws_data)
        ws_pivot.name = f"{yesterday.strftime("%#m月%#d日")}分析"

        source_str = f"{ws_data.name}!{data_range.address.replace('$', '')}"
        pivot_cache = wb.api.PivotCaches().Create(SourceType=1, SourceData=source_str)

        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=ws_pivot.range("A1").api,
            TableName="PivotTable1"
        )

        # 获取透视表字段集合
        pivot_fields = pivot_table.PivotFields()

        # 设置行字段和列字段
        # 1）将“（资源）地市名称”设置为行字段（xlRowField=1）
        row_field = pivot_fields.Item("（资源）地市名称")
        row_field.Orientation = 1  # xlRowField

        # 2）将“专业”设置为列字段（xlColumnField=2）
        col_field = pivot_fields.Item("专业")
        col_field.Orientation = 2  # xlColumnField

        # 3）将“（资源）地市名称”设置为数据字段并做计数（-4112=xlCount）
        pivot_table.AddDataField(
            pivot_fields.Item("告警标题"),   # 计数哪个字段
            "计数项:（资源）地市名称",         # 在透视表中显示的名字
            -4112                             # xlCount
        )

        # -------------------------------
        # 筛选：只允许甘肃省的城市出现
        # -------------------------------
        # 遍历 row_field 的所有 PivotItems
        for item in row_field.PivotItems():
            if item.Name not in self.cities:
                item.Visible = False

        # -------------------------------
        # 筛选：隐藏列字段中不需要的项
        # -------------------------------
        hide_list = ["5GC", "VIMS", "骨干云池", "固网", "信令网", "增值平台", "MEC", "业务平台"]
        for item in col_field.PivotItems():
            if item.Name in hide_list:
                item.Visible = False
                


        # =============================
        # 新增：按「专业 + 网络类型」的透视表
        # 列：专业（第一列层级）、网络类型（第二列层级）
        # 行：（资源）地市名称
        # 值：计数（资源）地市名称
        # =============================
        # ws_pivot2 = wb.sheets.add(before=ws_data)
        # ws_pivot2.name = f"{yesterday.strftime('%#m月%#d日')}分析-网络类型"


        # pivot_table2 = pivot_cache.CreatePivotTable(
        #     TableDestination=ws_pivot2.range("A1").api,
        #     TableName="PivotTable2"
        # )

        # 字段集合
        # pf2 = pivot_table2.PivotFields()

        # 行字段：（资源）地市名称
        # row_field2 = pf2.Item("（资源）地市名称")
        # row_field2.Orientation = 1  # xlRowField

        # 列字段：「网络类型」
        # col_nettype = pf2.Item("网络类型")
        # col_nettype.Orientation = 2 # xlColumnField

        # 数据字段：对「（资源）地市名称」计数（-4112=xlCount）
        # pivot_table2.AddDataField(
        #     pf2.Item("告警标题"),
        #     "计数项:（资源）地市名称",
        #     -4112
        # )

        # 行字段城市筛选（仅保留甘肃省城市）
        # for item in row_field2.PivotItems():
        #     if item.Name not in self.cities:
        #         item.Visible = False


        # -------------------------------
        # 列字段「网络类型」：仅保留「一干」「二干」
        # -------------------------------
        # target_set = {"一干", "二干"}

        # 批量更新时先关闭自动刷新，防止中间态为空导致WPS/Excel清空显示
        # pivot_table2.ManualUpdate = True

        # names = [it.Name for it in col_nettype.PivotItems()]
        # 如果目标项一个都不在，就放弃筛选，避免把表清空
        # if any(n in target_set for n in names):
        #     for it in col_nettype.PivotItems():
        #         it.Visible = (it.Name in target_set)
        # else:
            # 什么都不做（或视情况全部设为可见）
        #     for it in col_nettype.PivotItems():
        #         it.Visible = True

        # pivot_table2.ManualUpdate = False
        # pivot_table2.PivotCache().Refresh()
        

        target_path = os.path.join(path_utils.data_path, day.strftime("%m-%d"), "new_pivot.xlsx")
        wb.save(target_path)
        wb.close()

        logger.log("数据透视表生成完毕")
        return target_path


    # Temp func
    # def update_chart(self, file_path, chart_path, day):
    #     logger.log("开始更新数据图表")

    #     wb_pivot = xw.Book(file_path)
    #     wb_chart = xw.Book(chart_path)

    #     yesterday = day - timedelta(days=1)

    #     chart_sheet_name = f"{day.month}月分析"
    #     # chart_sheet_name = "3月分析"
    #     pivot_sheet_name = f"{yesterday.strftime('%#m月%#d')}日分析"
    #     data_sheet_name = f"{yesterday.strftime("%#m月%#d")}日告警"

    #     ws_pivot = wb_pivot.sheets[pivot_sheet_name]
    #     ws_chart = wb_chart.sheets[chart_sheet_name]
    #     ws_data = wb_pivot.sheets[data_sheet_name]

    #     logger.log("读入表正常")

    #     # 从 pivot 表读取"总计"列数据
    #     total_col = ws_pivot.cells(2, ws_pivot.cells.last_cell.column).end('left').column
    #     pivot_total = ws_pivot.range(ws_pivot.cells(3, total_col), ws_pivot.cells(17, total_col)).options(ndim=1).value
    #     # pivot_total = ws_pivot.range("M3:M17").options(ndim=1).value

    #     # 找到新的一列的位置
    #     last_col = ws_chart.cells(1, ws_chart.cells.last_cell.column).end('left').column + 1

    #     # 填入新数据
    #     ws_chart.range(ws_chart.cells(2, last_col), ws_chart.cells(len(pivot_total), last_col)).value = [[value] for value in pivot_total]

    #     # 填入新日期
    #     ws_chart.range(ws_chart.cells(1, last_col)).value = yesterday.strftime("%#m月%#d日")

    #     # 转换新列对应字母
    #     last_col_index = col_letter(last_col)

    #     new_source_str = f"{chart_sheet_name}!$DH$1:${last_col_index}$1,{chart_sheet_name}!$DH$16:${last_col_index}$16"

    #     chart = ws_chart.api.ChartObjects(1).Chart

    #     chart.SetSourceData(ws_chart.range(new_source_str).api)

    #     logger.log("写入数据正常")

    #     # 复制图表 sheet 到 pivot 文件
    #     ws_chart.api.Copy(Before=wb_pivot.sheets[0].api)

    #     file_name = f"{yesterday.strftime("%Y%m%d")}告警日分析.xlsx"
    #     # wb_pivot.save(os.path.join(path_utils.data_path, yesterday.strftime("%m-%d"), file_name))
    #     wb_pivot.save(os.path.join(path_utils.backbone_data_path, file_name))
    #     wb_chart.close()
    #     wb_pivot.close()

    #     logger.log("数据图表更新完毕！")    



    # def update_chart(self, file_path, chart_path):
        # logger.log("开始更新数据图表")

        # wb_pivot = xw.Book(file_path)
        # wb_chart = xw.Book(chart_path)

        # yesterday = datetime.today() - timedelta(days=1)

        # chart_sheet_name = f"{datetime.today().month}月分析"

        # # 如果不是同一个月（月初和上月末）
        # if yesterday.month != datetime.today().month:
        #     last_month_sheet = ws_chart.sheets[0]
        #     last_month_sheet.name = chart_sheet_name


        # print(f"sheet name: [{chart_sheet_name}]")

        # # chart_sheet_name = "3月分析"
        # pivot_sheet_name = f"{yesterday.strftime('%#m月%#d')}日分析"
        # data_sheet_name = f"{yesterday.strftime("%#m月%#d")}日告警"

        # ws_pivot = wb_pivot.sheets[pivot_sheet_name]
        # ws_chart = wb_chart.sheets[chart_sheet_name]
        # ws_data = wb_pivot.sheets[data_sheet_name]

        # print("[step check]: ok before reading data")

        # # 从 pivot 表读取"总计"列数据
        # total_col = ws_pivot.cells(2, ws_pivot.cells.last_cell.column).end('left').column
        # pivot_total = ws_pivot.range(ws_pivot.cells(3, total_col), ws_pivot.cells(17, total_col)).options(ndim=1).value
        # # pivot_total = ws_pivot.range("M3:M17").options(ndim=1).value

        # # 找到新的一列的位置
        # last_col = ws_chart.cells(1, ws_chart.cells.last_cell.column).end('left').column + 1

        # # 填入新数据
        # ws_chart.range(ws_chart.cells(2, last_col), ws_chart.cells(len(pivot_total), last_col)).value = [[value] for value in pivot_total]

        # # 填入新日期
        # ws_chart.range(ws_chart.cells(1, last_col)).value = yesterday.strftime("%#m月%#d日")

        # # 转换新列对应字母
        # last_col_index = col_letter(last_col)

        # new_source_str = f"{chart_sheet_name}!$DH$1:${last_col_index}$1,{chart_sheet_name}!$DH$16:${last_col_index}$16"

        # chart = ws_chart.api.ChartObjects(1).Chart

        # chart.SetSourceData(ws_chart.range(new_source_str).api)

        # # 复制图表 sheet 到 pivot 文件
        # ws_chart.api.Copy(Before=wb_pivot.sheets[0].api)

        # file_name = f"{yesterday.strftime("%Y%m%d")}告警日分析.xlsx"
        # # wb_pivot.save(os.path.join(path_utils.data_path, yesterday.strftime("%m-%d"), file_name))
        # wb_pivot.save(os.path.join(path_utils.backbone_data_path, file_name))
        # wb_chart.close()
        # wb_pivot.close()

        # logger.log("数据图表更新完毕！")





    # def gen_pivot_table(self, file_path):
    #     logger.log("开始生成数据透视表")

    #     wb = xw.Book(file_path)
    #     ws_data = wb.sheets['Sheet1']

    #     yesterday = datetime.today() - timedelta(days=1)
    #     ws_data.name = f"{yesterday.strftime("%#m月%#d日")}告警"

    #     # data_range = ws_data.range("A1").expand()
    #     data_range = ws_data.used_range

    #     ws_pivot = wb.sheets.add(before=ws_data)
    #     ws_pivot.name = f"{yesterday.strftime("%#m月%#d日")}分析"

    #     source_str = f"{ws_data.name}!{data_range.address.replace('$', '')}"
    #     pivot_cache = wb.api.PivotCaches().Create(SourceType=1, SourceData=source_str)

    #     pivot_table = pivot_cache.CreatePivotTable(
    #         TableDestination=ws_pivot.range("A1").api,
    #         TableName="PivotTable1"
    #     )

    #     # 获取透视表字段集合
    #     pivot_fields = pivot_table.PivotFields()

    #     # 设置行字段和列字段
    #     # 1）将“（资源）地市名称”设置为行字段（xlRowField=1）
    #     row_field = pivot_fields.Item("（资源）地市名称")
    #     row_field.Orientation = 1  # xlRowField

    #     # 2）将“专业”设置为列字段（xlColumnField=2）
    #     col_field = pivot_fields.Item("专业")
    #     col_field.Orientation = 2  # xlColumnField

    #     # 3）将“（资源）地市名称”设置为数据字段并做计数（-4112=xlCount）
    #     pivot_table.AddDataField(
    #         pivot_fields.Item("告警标题"),   # 计数哪个字段
    #         "计数项:（资源）地市名称",         # 在透视表中显示的名字
    #         -4112                             # xlCount
    #     )

    #     # -------------------------------
    #     # 筛选：只允许甘肃省的城市出现
    #     # -------------------------------
    #     # 遍历 row_field 的所有 PivotItems
    #     for item in row_field.PivotItems():
    #         if item.Name not in self.cities:
    #             item.Visible = False

    #     # -------------------------------
    #     # 筛选：隐藏列字段中不需要的项
    #     # -------------------------------
    #     hide_list = ["5GC", "VIMS", "骨干云池", "固网", "信令网", "增值平台", "MEC", "业务平台"]
    #     for item in col_field.PivotItems():
    #         if item.Name in hide_list:
    #             item.Visible = False
                


    #     # =============================
    #     # 新增：按「专业 + 网络类型」的透视表
    #     # 列：专业（第一列层级）、网络类型（第二列层级）
    #     # 行：（资源）地市名称
    #     # 值：计数（资源）地市名称
    #     # =============================
    #     ws_pivot2 = wb.sheets.add(before=ws_data)
    #     ws_pivot2.name = f"{yesterday.strftime('%#m月%#d日')}分析-网络类型"


    #     pivot_table2 = pivot_cache.CreatePivotTable(
    #         TableDestination=ws_pivot2.range("A1").api,
    #         TableName="PivotTable2"
    #     )

    #     # 字段集合
    #     pf2 = pivot_table2.PivotFields()

    #     # 行字段：（资源）地市名称
    #     row_field2 = pf2.Item("（资源）地市名称")
    #     row_field2.Orientation = 1  # xlRowField

    #     # 列字段：「网络类型」
    #     col_nettype = pf2.Item("网络类型")
    #     col_nettype.Orientation = 2 # xlColumnField

    #     # 数据字段：对「（资源）地市名称」计数（-4112=xlCount）
    #     pivot_table2.AddDataField(
    #         pf2.Item("告警标题"),
    #         "计数项:（资源）地市名称",
    #         -4112
    #     )

    #     # 行字段城市筛选（仅保留甘肃省城市）
    #     for item in row_field2.PivotItems():
    #         if item.Name not in self.cities:
    #             item.Visible = False


    #     # -------------------------------
    #     # 列字段「网络类型」：仅保留「一干」「二干」
    #     # -------------------------------
    #     target_set = {"一干", "二干"}

    #     # 批量更新时先关闭自动刷新，防止中间态为空导致WPS/Excel清空显示
    #     pivot_table2.ManualUpdate = True

    #     names = [it.Name for it in col_nettype.PivotItems()]
    #     # 如果目标项一个都不在，就放弃筛选，避免把表清空
    #     if any(n in target_set for n in names):
    #         for it in col_nettype.PivotItems():
    #             it.Visible = (it.Name in target_set)
    #     else:
    #         # 什么都不做（或视情况全部设为可见）
    #         for it in col_nettype.PivotItems():
    #             it.Visible = True

    #     pivot_table2.ManualUpdate = False
    #     pivot_table2.PivotCache().Refresh()
        

    #     target_path = os.path.join(path_utils.data_path, datetime.today().strftime("%m-%d"), "new_pivot.xlsx")
    #     wb.save(target_path)
    #     wb.close()

    #     logger.log("数据透视表生成完毕")
    #     return target_path


    # # Temp func
    # # def update_chart(self, file_path, chart_path, day):
    # #     logger.log("开始更新数据图表")

    # #     wb_pivot = xw.Book(file_path)
    # #     wb_chart = xw.Book(chart_path)

    # #     yesterday = day - timedelta(days=1)

    # #     chart_sheet_name = f"{day.month}月分析"
    # #     # chart_sheet_name = "3月分析"
    # #     pivot_sheet_name = f"{yesterday.strftime('%#m月%#d')}日分析"
    # #     data_sheet_name = f"{yesterday.strftime("%#m月%#d")}日告警"

    # #     ws_pivot = wb_pivot.sheets[pivot_sheet_name]
    # #     ws_chart = wb_chart.sheets[chart_sheet_name]
    # #     ws_data = wb_pivot.sheets[data_sheet_name]

    # #     logger.log("读入表正常")

    # #     # 从 pivot 表读取"总计"列数据
    # #     total_col = ws_pivot.cells(2, ws_pivot.cells.last_cell.column).end('left').column
    # #     pivot_total = ws_pivot.range(ws_pivot.cells(3, total_col), ws_pivot.cells(17, total_col)).options(ndim=1).value
    # #     # pivot_total = ws_pivot.range("M3:M17").options(ndim=1).value

    # #     # 找到新的一列的位置
    # #     last_col = ws_chart.cells(1, ws_chart.cells.last_cell.column).end('left').column + 1

    # #     # 填入新数据
    # #     ws_chart.range(ws_chart.cells(2, last_col), ws_chart.cells(len(pivot_total), last_col)).value = [[value] for value in pivot_total]

    # #     # 填入新日期
    # #     ws_chart.range(ws_chart.cells(1, last_col)).value = yesterday.strftime("%#m月%#d日")

    # #     # 转换新列对应字母
    # #     last_col_index = col_letter(last_col)

    # #     new_source_str = f"{chart_sheet_name}!$DH$1:${last_col_index}$1,{chart_sheet_name}!$DH$16:${last_col_index}$16"

    # #     chart = ws_chart.api.ChartObjects(1).Chart

    # #     chart.SetSourceData(ws_chart.range(new_source_str).api)

    # #     logger.log("写入数据正常")

    # #     # 复制图表 sheet 到 pivot 文件
    # #     ws_chart.api.Copy(Before=wb_pivot.sheets[0].api)

    # #     file_name = f"{yesterday.strftime("%Y%m%d")}告警日分析.xlsx"
    # #     # wb_pivot.save(os.path.join(path_utils.data_path, yesterday.strftime("%m-%d"), file_name))
    # #     wb_pivot.save(os.path.join(path_utils.backbone_data_path, file_name))
    # #     wb_chart.close()
    # #     wb_pivot.close()

    # #     logger.log("数据图表更新完毕！")    



    def update_chart(self, file_path, chart_path):
        logger.log("开始更新数据图表")
        
        app = xw.App(visible=False, add_book=False)
        wb_pivot = app.books.open(file_path)
        wb_chart = app.books.open(chart_path)

        # wb_pivot = xw.Book(file_path)
        # wb_chart = xw.Book(chart_path)

        # 测试xlwings环境
        # app = wb_pivot.app
        # print(app.api.Name)

        yesterday = datetime.today() - timedelta(days=1)

        chart_sheet_name = f"{datetime.today().month}月分析"

        # 如果不是同一个月（月初和上月末）
        if yesterday.month != datetime.today().month:
            last_month_sheet = ws_chart.sheets[0]
            last_month_sheet.name = chart_sheet_name


        print(f"sheet name: [{chart_sheet_name}]")

        # chart_sheet_name = "3月分析"
        pivot_sheet_name = f"{yesterday.strftime('%#m月%#d')}日分析"
        data_sheet_name = f"{yesterday.strftime("%#m月%#d")}日告警"

        ws_pivot = wb_pivot.sheets[pivot_sheet_name]
        ws_chart = wb_chart.sheets[chart_sheet_name]
        ws_data = wb_pivot.sheets[data_sheet_name]

        print("[step check]: ok before reading data")

        # 从 pivot 表读取"总计"列数据
        total_col = ws_pivot.cells(2, ws_pivot.cells.last_cell.column).end('left').column
        pivot_total = ws_pivot.range(ws_pivot.cells(3, total_col), ws_pivot.cells(17, total_col)).options(ndim=1).value
        # pivot_total = ws_pivot.range("M3:M17").options(ndim=1).value

        # 找到新的一列的位置
        last_col = ws_chart.cells(1, ws_chart.cells.last_cell.column).end('left').column + 1

        # 填入新数据
        ws_chart.range(ws_chart.cells(2, last_col), ws_chart.cells(len(pivot_total), last_col)).value = [[value] for value in pivot_total]

        # 填入新日期
        ws_chart.range(ws_chart.cells(1, last_col)).value = yesterday.strftime("%#m月%#d日")

        # 转换新列对应字母
        last_col_index = col_letter(last_col)

        new_source_str = f"{chart_sheet_name}!$DH$1:${last_col_index}$1,{chart_sheet_name}!$DH$16:${last_col_index}$16"

        chart = ws_chart.api.ChartObjects(1).Chart

        chart.SetSourceData(ws_chart.range(new_source_str).api)

        # 复制图表 sheet 到 pivot 文件
        # ws_chart.api.Copy(Before=wb_pivot.sheets[0].api) 
        ws_chart.copy(before=wb_pivot.sheets[0])

        file_name = f"{yesterday.strftime("%Y%m%d")}告警日分析.xlsx"
        # wb_pivot.save(os.path.join(path_utils.data_path, yesterday.strftime("%m-%d"), file_name))
        wb_pivot.save(os.path.join(path_utils.backbone_data_path, file_name))
        wb_chart.close()
        wb_pivot.close()

        logger.log("数据图表更新完毕！")


    # 带有新数据透视表
    # def update_chart(self, file_path, chart_path):
    #     logger.log("开始更新数据图表")

    #     # return

    #     wb_pivot = xw.Book(file_path)
    #     wb_chart = xw.Book(chart_path)

    #     yesterday = datetime.today() - timedelta(days=1)

    #     chart_sheet_name = f"{datetime.today().month}月分析"
    #     pivot_sheet_name = f"{yesterday.strftime('%#m月%#d')}日分析"
    #     pivot_sheet_name2 = f"{yesterday.strftime('%#m月%#d')}日分析-网络类型"  # 第二张
    #     data_sheet_name = f"{yesterday.strftime('%#m月%#d')}日告警"

    #     ws_pivot = wb_pivot.sheets[pivot_sheet_name]
    #     ws_pivot2 = wb_pivot.sheets[pivot_sheet_name2]
    #     ws_chart = wb_chart.sheets[chart_sheet_name]
    #     ws_data = wb_pivot.sheets[data_sheet_name]

    #     print("[step check]: ok before reading data")

    #     # ------------------------
    #     # 第一张透视表：读取总计列
    #     # ------------------------
    #     total_col = ws_pivot.cells(2, ws_pivot.cells.last_cell.column).end('left').column
    #     pivot_total = ws_pivot.range(
    #         ws_pivot.cells(3, total_col),
    #         ws_pivot.cells(17, total_col)
    #     ).options(ndim=1).value

    #     # ------------------------
    #     # 第二张透视表：取列标签为「一干」「二干」在“总计”行的值
    #     # ------------------------
    #     used = ws_pivot2.api.UsedRange
    #     last_row = used.Rows.Count
    #     last_col = used.Columns.Count

    #     # 1) 找到“列标签”所在行（通常在第1行），其下一行是具体列项
    #     col_labels_row = None
    #     for r in range(1, min(5, last_row + 1)):  # 前几行里找
    #         val = ws_pivot2.cells(r, 2).value  # B列一般会写“列标签”
    #         if isinstance(val, str) and "列标签" in val:
    #             col_labels_row = r
    #             break
    #     if not col_labels_row:
    #         col_labels_row = 1  # 兜底
    #     items_row = col_labels_row + 1  # 这一行放的是各列项（如 一干/二干/总计）

    #     # 2) 找到“一干”“二干”的列号
    #     col_yigan = col_ergan = None
    #     for c in range(2, last_col + 1):
    #         v = ws_pivot2.cells(items_row, c).value
    #         if v == "一干":
    #             col_yigan = c
    #         elif v == "二干":
    #             col_ergan = c

    #     # 3) 找到“总计”行（在A列从下往上找）
    #     grand_row = None
    #     for r in range(last_row, 1, -1):
    #         v = ws_pivot2.cells(r, 1).value
    #         if v == "总计":
    #             grand_row = r
    #             break

    #     # 4) 交叉取值
    #     x1 = x2 = 0
    #     if grand_row:
    #         if col_yigan:
    #             x1 = ws_pivot2.cells(grand_row, col_yigan).value or 0  # 一干
    #         if col_ergan:
    #             x2 = ws_pivot2.cells(grand_row, col_ergan).value or 0  # 二干

    #     print(f"一干总计={x1}, 二干总计={x2}")


    #     # 拼接新数据
    #     if isinstance(pivot_total, (list, tuple)):
    #         pivot_total = list(pivot_total)
    #     else:
    #         pivot_total = [pivot_total]

    #     pivot_total.extend([x1, x2])

    #     print(f"附加的二干/一干总计值: {x2}, {x1}")

    #     # ------------------------
    #     # 在图表文件中填充新列
    #     # ------------------------
    #     last_col = ws_chart.cells(1, ws_chart.cells.last_cell.column).end('left').column + 1
    #     ws_chart.range(
    #         ws_chart.cells(2, last_col),
    #         ws_chart.cells(len(pivot_total) + 1, last_col)
    #     ).value = [[v] for v in pivot_total]

    #     ws_chart.range(ws_chart.cells(1, last_col)).value = yesterday.strftime("%#m月%#d日")

    #     # ------------------------
    #     # 更新图表引用
    #     # ------------------------
    #     last_col_index = col_letter(last_col)
    #     new_source_str = f"{chart_sheet_name}!$DH$1:${last_col_index}$1,{chart_sheet_name}!$DH$16:${last_col_index}$16"

    #     chart = ws_chart.api.ChartObjects(1).Chart
    #     chart.SetSourceData(ws_chart.range(new_source_str).api)

    #     # ------------------------
    #     # 保存并复制到透视文件
    #     # ------------------------
    #     # ws_chart.api.Copy(Before=wb_pivot.sheets[0].api)
    #     file_name = f"{yesterday.strftime('%Y%m%d')}告警日分析.xlsx"
    #     wb_pivot.save(os.path.join(path_utils.backbone_data_path, file_name))

    #     wb_chart.close()
    #     wb_pivot.close()

    #     logger.log("数据图表更新完毕！")
