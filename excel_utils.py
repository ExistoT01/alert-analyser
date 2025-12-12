import zipfile
import os
import pandas as pd
import xlwings as xw
from datetime import datetime, timedelta

import config_utils
from config_utils import Config_Utils
from logger import Logger


config = Config_Utils()
logger = Logger(True, config.log_path)


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
        self.cities = config.cities

    # 解压文件
    def unzip(self, file_name, day):
        logger.log(f'正在解压文件：{file_name}')

        with zipfile.ZipFile(os.path.join(config.zips_path, file_name)) as f:
            target_path = os.path.join(config.data_path, day.strftime("%m-%d"))
            os.makedirs(target_path, exist_ok=True)

            # 如果目录不为空，清空目录
            for file_name in os.listdir(target_path):
                file_path = os.path.join(target_path, file_name)
                try:
                    if os.path.isfile(file_path):
                        os.remove(file_path)
                except Exception as e:
                    logger.log(f'删除 {file_path} 失败！原因：{e}')
                    raise

            f.extractall(target_path)

        logger.log("文件解压完毕")


    # 合并多张Excel表
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


        # "网络类型" 列剔除 "一干"
        # "告警标题" 列剔除 "GPON ONT掉电(DGi)"
        # "专业" 列剔除 "5GC、VIMS、骨干云池、固网、信令网、增值平台、MEC、业务平台"
        drop_prof = ["5GC", "VIMS", "骨干云池", "固网", "信令网", "增值平台", "MEC", "业务平台"]

        filtered_df = merged_df[
            (merged_df["网络类型"].fillna("") != "一干") &
            (merged_df["告警标题"].fillna("") != "GPON ONT掉电(DGi)") &
            (~merged_df["专业"].fillna("").isin(drop_prof))
        ]

        target_path = os.path.join(config.data_path, day.strftime("%m-%d"))
        os.makedirs(target_path, exist_ok=True)

        file_name = "merged_" + day.strftime("%m_%d") + ".xlsx"
        filtered_df.to_excel(os.path.join(target_path, file_name), index=False)

        logger.log("Excel文件合并完毕")
        return os.path.join(target_path, file_name)
    

    # 生成数据透视表
    def gen_pivot_table(self, file_path, day):
        logger.log("开始生成数据透视表")

        wb = xw.Book(file_path)
        ws_data = wb.sheets['Sheet1']

        yesterday = day - timedelta(days=1)
        ws_data.name = f"{yesterday.strftime("%#m月%#d日")}告警"

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
        # 将 "（资源）地市名称" 设置为行字段（xlRowField=1）
        row_field = pivot_fields.Item("（资源）地市名称")
        row_field.Orientation = 1  # xlRowField

        # 将 "专业" 设置为列字段（xlColumnField=2）
        col_field = pivot_fields.Item("专业")
        col_field.Orientation = 2  # xlColumnField

        # 将 "（资源）地市名称" 设置为数据字段并做计数（-4112=xlCount）
        pivot_table.AddDataField(
            pivot_fields.Item("告警标题"),   # 计数哪个字段
            "计数项:（资源）地市名称",         # 在透视表中显示的名字
            -4112                             # xlCount
        )

        # 遍历 row_field 的所有 PivotItems，筛选出甘肃地市
        for item in row_field.PivotItems():
            if item.Name not in self.cities:
                item.Visible = False

        target_path = os.path.join(config.data_path, day.strftime("%m-%d"), "new_pivot.xlsx")
        wb.save(target_path)
        wb.close()

        logger.log("数据透视表生成完毕")
        return target_path


    # 更新数据图表
    def update_chart(self, file_path, chart_path, day):
        logger.log("开始更新数据图表")
        
        app = xw.App(visible=False, add_book=False)
        wb_pivot = app.books.open(file_path)
        wb_chart = app.books.open(chart_path)

        yesterday = day - timedelta(days=1)

        chart_sheet_name = f"{datetime.today().month}月分析"

        # 如果不是同一个月（月初和上月末）
        if yesterday.month != datetime.today().month:
            last_month_sheet = ws_chart.sheets[0]
            last_month_sheet.name = chart_sheet_name

        pivot_sheet_name = f"{yesterday.strftime('%#m月%#d')}日分析"
        data_sheet_name = f"{yesterday.strftime("%#m月%#d")}日告警"

        ws_pivot = wb_pivot.sheets[pivot_sheet_name]
        ws_chart = wb_chart.sheets[chart_sheet_name]

        # 从 pivot 表读取"总计"列数据
        total_col = ws_pivot.cells(2, ws_pivot.cells.last_cell.column).end('left').column
        pivot_total = ws_pivot.range(ws_pivot.cells(3, total_col), ws_pivot.cells(17, total_col)).options(ndim=1).value

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
        ws_chart.copy(before=wb_pivot.sheets[0])

        file_name = f"{yesterday.strftime("%Y%m%d")}告警日分析.xlsx"
        wb_pivot.save(os.path.join(config.backbone_data_path, file_name))
        wb_chart.close()
        wb_pivot.close()

        logger.log("数据图表更新完毕！")