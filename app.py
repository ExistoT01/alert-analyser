import os.path
import threading

import customtkinter as ctk
import time
import schedule
from PIL import Image

from alert_utils import AlertUtils
from excel_utils import ExcelUtils
from path_utils import PathUtils
from pathlib import Path
from datetime import datetime, timedelta
from logger import Logger


path_utils = PathUtils()
logger = Logger(True, path_utils.log_path)



class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Analyse Generator")
        self.geometry("400x400")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        font_facon = ctk.CTkFont("facon")
        font_lovelo = ctk.CTkFont("Lovelo-Linelight")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        self.status_prefix = "当前任务："
        self.status_nothing = "当前无任务正在执行"

        self.alert_utils = AlertUtils()
        self.excel_utils = ExcelUtils()

        # logo
        self.icon_img = ctk.CTkImage(light_image=Image.open(path_utils.light_icon_path), dark_image=Image.open(path_utils.dark_icon_path), size=(40, 40))
        self.icon_label = ctk.CTkLabel(self, image=self.icon_img, text="")
        self.icon_label.grid(row=0, column=0, padx=20, pady=10)

        # title
        self.title_label = ctk.CTkLabel(self, font=font_lovelo, text="Alert Analyzer")
        self.title_label.cget("font").configure(size=36)
        self.title_label.grid(row=1, column=0, padx=20, pady=0)

        # version
        self.version_label = ctk.CTkLabel(self, text="v1.0")
        self.version_label.grid(row=2, column=0, padx=20, pady=0)

        # rebuild btn
        self.rebuild_btn = ctk.CTkButton(self, text="重做当日报表", command=self.rebuild)
        self.rebuild_btn.grid(row=3, column=0, padx=20, pady=20)

        # status
        self.status_label = ctk.CTkLabel(self, text=self.status_prefix + self.status_nothing)
        self.status_label.grid(row=4, column=0, padx=20, pady=20)


        # self.export_files_btn = ctk.CTkButton(self, text="click", command=self.export_files)
        # self.export_files_btn = ctk.CTkButton(self, text="click", command=self.test_func)
        # self.export_files_btn.grid(row=2, column=0, padx=20, pady=20, columnspan=2)

        # self.merge_files_btn = ctk.CTkButton(self, text="merge", command=self.merge_excels)
        # self.merge_files_btn.grid(row=3, column=0, padx=20, pady=20)

        self.test_func()
        # self.setup_schedule()

    # rebuild btn function
    def rebuild(self):
        logger.log("开始重做当日报表")
        self.export_files()

    # def export_files(self):
    #     self.update_app_status("设置session id")
    #     self.alert_utils.set_session_id()

    #     self.update_app_status("从系统导出告警文件")
    #     self.alert_utils.export_csv_files()

    #     file_name = ""
    #     data_object = None

    #     while True:
    #         data_object = self.alert_utils.check_export_progress()
    #         if data_object and data_object['progress'] != 100:
    #             logger.log(f"当前导出进度：{self.alert_utils.check_export_progress()['progress']}")
    #             time.sleep(5)
    #             continue
    #         break

    #     file_src = data_object['fileSrc']
    #     logger.log(f"导出完成！文件地址：{file_src}")

    #     file_name = file_src.split('/')[-1]

    #     self.update_app_status("下载告警文件压缩包")
    #     self.alert_utils.download_files(file_src)

    #     self.merge_excels(file_name)


    def export_files(self, day):
        self.update_app_status("设置session id")
        self.alert_utils.set_session_id(day)

        self.update_app_status("从系统导出告警文件")
        self.alert_utils.export_csv_files()

        file_name = ""
        data_object = None

        while True:
            data_object = self.alert_utils.check_export_progress()
            if data_object and data_object['progress'] != 100:
                logger.log(f"当前导出进度：{self.alert_utils.check_export_progress()['progress']}")
                time.sleep(5)
                continue
            break

        file_src = data_object['fileSrc']
        logger.log(f"导出完成！文件地址：{file_src}")

        file_name = file_src.split('/')[-1]

        self.update_app_status("下载告警文件压缩包")
        self.alert_utils.download_files(file_src)

        self.merge_excels(file_name, day)




    # def merge_excels(self, file_name):
    #     self.update_app_status("解压文件")
    #     self.excel_utils.unzip(file_name)

    #     target_path = os.path.join(path_utils.data_path, datetime.today().strftime("%m-%d"))
    #     path = Path(target_path)

    #     self.update_app_status("合并文件")
    #     merged_file = self.excel_utils.concat(list(path.rglob("*")))

    #     self.update_app_status("生成数据透视表")
    #     pivot_table_path = self.excel_utils.gen_pivot_table(merged_file)

    #     self.update_app_status("更新数据图表")
    #     the_day_before_yesterday = datetime.today() - timedelta(days=2)
    #     self.excel_utils.update_chart(pivot_table_path, os.path.join(path_utils.backbone_data_path, f"{the_day_before_yesterday.strftime("%Y%m%d")}告警日分析.xlsx"))

    #     self.update_app_status(self.status_nothing)

    def merge_excels(self, file_name, day):
        self.update_app_status("解压文件")
        self.excel_utils.unzip(file_name, day)

        target_path = os.path.join(path_utils.data_path, day.strftime("%m-%d"))
        path = Path(target_path)

        self.update_app_status("合并文件")
        merged_file = self.excel_utils.concat(list(path.rglob("*")), day)

        self.update_app_status("生成数据透视表")
        pivot_table_path = self.excel_utils.gen_pivot_table(merged_file, day)

        self.update_app_status("更新数据图表")
        the_day_before_yesterday = day - timedelta(days=2)
        self.excel_utils.update_chart(pivot_table_path, os.path.join(path_utils.backbone_data_path, f"{the_day_before_yesterday.strftime("%Y%m%d")}告警日分析.xlsx"))

        self.update_app_status(self.status_nothing)


    def test_func(self):

        start_date = datetime(2025, 12, 8)
        end_date = datetime(2025, 12, 8)

        day = start_date
        while day <= end_date:
            try:
                self.export_files(day)

            except Exception as e:
                logger.log(f"❌ {day.strftime('%m-%d')} 日出错：{e}")

            day += timedelta(days=1)

        # self.excel_utils.unzip("alarmExport20250404123100.zip")
        #
        # target_path = os.path.join(path_utils.data_path, datetime.today().strftime("%m-%d"))
        # path = Path(target_path)
        #
        # print(target_path)
        # print(list(path.rglob("*")))
        # merged_file = self.excel_utils.concat(list(path.rglob("*")))
        #
        # pivot_table_path = self.excel_utils.gen_pivot_table(merged_file)

        # the_day_before_yesterday = datetime.today() - timedelta(days=2)
        # self.excel_utils.update_chart("D:\\Perry\\daily-analyse\\gui-app\\data\\04-04\\new_pivot.xlsx", os.path.join(path_utils.data_path, the_day_before_yesterday.strftime("%m-%d"), "20250402告警日分析.xlsx"))

        # self.alert_utils.download_files("/static/alarmExport20250820085537.zip")


        # 替换名字(例如: "alarmExport20250824081857.zip")
        # self.merge_excels("alarmExport20251031003158.zip")


    def update_app_status(self, msg):
        self.status_label.configure(text=self.status_prefix + msg)


    def setup_schedule(self):
        def job():
            logger.log("开始每日任务")
            self.export_files()

        def run_schedule():
            # schedule.every(1).minutes.do(job)
            schedule.every().day.at("06:00").do(job)
            while True:
                # schedule.run_pending()
                self.test_func()
                # time.sleep(30)

        thread = threading.Thread(target=run_schedule, daemon=True)
        thread.start()


app = App()
app.mainloop()