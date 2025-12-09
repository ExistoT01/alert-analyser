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

        # 若是首次运行，检查目录下是否存在zips与datas文件夹，创建如果不存在
        os.makedirs(path_utils.zips_path, exist_ok=True)
        os.makedirs(path_utils.data_path, exist_ok=True)

        self.title("Analyse Generator")
        self.geometry("400x400")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        font_facon = ctk.CTkFont("facon")
        font_lovelo = ctk.CTkFont("Lovelo-Linelight")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        self.status_prefix = "【当前任务】："
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

        # open folder
        self.open_folder_btn = ctk.CTkButton(self, text="打开告警分析目录", command=self.open_folder)
        self.open_folder_btn.grid(row=3, column=0, padx=20, pady=10)

        # rebuild btn
        self.rebuild_btn = ctk.CTkButton(self, text="重做当日报表", command=self.rebuild)
        self.rebuild_btn.grid(row=4, column=0, padx=20, pady=10)

        # status
        self.status_label = ctk.CTkLabel(self, text=self.status_prefix + self.status_nothing)
        self.status_label.grid(row=5, column=0, padx=20, pady=10)

        # progress bar
        self.progress_bar = ctk.CTkProgressBar(self, orientation='horizontal', mode='determinate')
        self.progress_bar.grid(row=6, column=0, padx=20, pady=20)
        self.progress_bar.set(0)
        self.progress_bar.grid_remove()
        self.current_step = 0

        # 测试代码 or 实际应用
        # self.test_func()
        self.setup_schedule()

    
    # 打开告警分析目录
    def open_folder(self):
        os.startfile(path_utils.backbone_data_path)


    # 重做每日报表
    def rebuild(self):
        logger.log("开始重做当日报表")

        self.rebuild_btn.configure(state="disabled")

        def task():
            try:
                self.export_files()
            finally:
                self.after(0, lambda: self.rebuild_btn.configure(state="normal"))

        thread = threading.Thread(target=task, daemon=True)
        thread.start()


    # progress bar 封装函数
    def show_progress(self):
        def _show():
            self.current_step = 0
            self.progress_bar.set(0)
            self.progress_bar.grid()
        self.after(0, _show)


    def hide_progress(self):
        def _hide():
            self.progress_bar.grid_remove()
        self.after(0, _hide)


    def set_progress_step(self, step: int):
        def _update():
            self.current_step = step
            self.progress_bar.set(self.current_step / 7.0)
        self.after(0, _update)

    
    # 更新状态封装函数
    def set_status(self, msg: str):
        def _update():
            self.status_label.configure(text=self.status_prefix + msg)
        self.after(0, _update)


    # 导出文件/任务开始
    def export_files(self, day=None):
        self.show_progress()

        # 如果未指定日期，则为正常每日报表生成
        if day is None:
            day = datetime.today()

        self.set_status("设置session id (1/7)")
        self.set_progress_step(1)
        self.alert_utils.set_session_id(day)

        self.set_status("从系统导出告警文件 (2/7)")
        self.set_progress_step(2)
        self.alert_utils.export_csv_files()

        file_name = ""
        data_object = None

        while True:
            data_object = self.alert_utils.check_export_progress()
            if data_object and data_object['progress'] != 100:
                logger.log(f"当前导出进度：{self.alert_utils.check_export_progress()['progress']}")
                time.sleep(10)
                continue
            break

        file_src = data_object['fileSrc']
        logger.log(f"导出完成！文件地址：{file_src}")

        file_name = file_src.split('/')[-1]

        self.set_status("下载告警文件压缩包 (3/7)")
        self.set_progress_step(3)
        self.alert_utils.download_files(file_src)

        self.merge_excels(file_name, day)


    # excel操作函数
    def merge_excels(self, file_name, day):
        self.set_status("解压文件 (4/7)")
        self.set_progress_step(4)
        self.excel_utils.unzip(file_name, day)

        target_path = os.path.join(path_utils.data_path, day.strftime("%m-%d"))
        path = Path(target_path)

        self.set_status("合并文件 (5/7)")
        self.set_progress_step(5)
        merged_file = self.excel_utils.concat(list(path.rglob("*")), day)

        self.set_status("生成数据透视表 (6/7)")
        self.set_progress_step(6)
        pivot_table_path = self.excel_utils.gen_pivot_table(merged_file, day)

        self.set_status("更新数据图表 (7/7)")
        self.set_progress_step(7)
        the_day_before_yesterday = day - timedelta(days=2)
        self.excel_utils.update_chart(pivot_table_path, os.path.join(path_utils.backbone_data_path, f"{the_day_before_yesterday.strftime("%Y%m%d")}告警日分析.xlsx"), day)

        self.hide_progress()
        self.set_status(self.status_nothing)


    # 测试函数，目前用于月告警总结分析
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


    def setup_schedule(self):
        def job():
            logger.log("开始每日任务")
            self.export_files()

        def run_schedule():
            # schedule.every(1).minutes.do(job)
            schedule.every().day.at("06:00").do(job)
            while True:
                schedule.run_pending()

        thread = threading.Thread(target=run_schedule, daemon=True)
        thread.start()


app = App()
app.mainloop()