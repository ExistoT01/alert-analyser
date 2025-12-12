import os.path
import threading

from tkinter import filedialog
from CTkMessagebox import CTkMessagebox
import customtkinter as ctk
import time
import schedule
from PIL import Image

from alert_utils import AlertUtils
from excel_utils import ExcelUtils
from config_utils import Config_Utils
from pathlib import Path
from datetime import datetime, timedelta
from logger import Logger


config = Config_Utils()
logger = Logger(True, config.log_path)

BG_COLOR = "#292929"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 若是首次运行，检查目录下是否存在zips与datas文件夹，创建如果不存在
        os.makedirs(config.zips_path, exist_ok=True)
        os.makedirs(config.data_path, exist_ok=True)

        # 界面设置
        self.title("Analyse Generator")
        self.geometry("400x400")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # fonts
        self.font_lovelo = ctk.CTkFont("Lovelo Line light")
        self.font_microsoft_yahei_normal = ctk.CTkFont("Microsoft YaHei UI", weight="normal")
        self.font_microsoft_yahei_bold = ctk.CTkFont("Microsoft YaHei UI", weight="bold")

        self.status_prefix = "【当前任务】："
        self.status_nothing = "当前无任务正在执行"

        self.alert_utils = AlertUtils()
        self.excel_utils = ExcelUtils()

        # icons
        self.setting_icon = ctk.CTkImage(light_image=Image.open(config.setting_icon_path), dark_image=Image.open(config.setting_icon_path), size=(20, 20))
        self.setting_page_icon = ctk.CTkImage(light_image=Image.open(config.setting_icon_path), dark_image=Image.open(config.setting_icon_path), size=(60, 60))
        self.icon_img = ctk.CTkImage(light_image=Image.open(config.logo_icon_path), dark_image=Image.open(config.logo_icon_path), size=(60, 60))
        self.folder_icon = ctk.CTkImage(light_image=Image.open(config.folder_icon_path), dark_image=Image.open(config.folder_icon_path), size=(20, 20))
        self.excel_icon = ctk.CTkImage(light_image=Image.open(config.excel_icon_path), dark_image=Image.open(config.excel_icon_path), size=(20, 20))
        self.back_icon = ctk.CTkImage(light_image=Image.open(config.back_icon_path), dark_image=Image.open(config.back_icon_path), size=(20, 20))

        # topbar
        top_bar = ctk.CTkFrame(self, fg_color=BG_COLOR)
        top_bar.pack(fill="x", pady=(5, 0))

        ## nav button
        self.nav_btn = ctk.CTkButton(top_bar, image=self.setting_icon, text="", width=20, height=20, fg_color="transparent", hover_color="gray75", command=self.show_setting_page)
        self.nav_btn.pack(side="left", padx=10)

        # page container
        self.page_container = ctk.CTkFrame(self)
        self.page_container.pack(expand=True, fill="both")
        self.page_container.grid_rowconfigure(0, weight=1)
        self.page_container.grid_columnconfigure(0, weight=1)

        # Main page
        self.main_page = ctk.CTkFrame(self.page_container)
        self.main_page.grid(row=0, column=0, sticky="nsew")
        self._build_main_page()

        # setting page
        self.setting_page = ctk.CTkFrame(self.page_container)
        self.setting_page.grid(row=0, column=0, sticky="nsew")
        self._build_setting_page()

        # show the main page at initial
        self.show_main_page()

        # 测试代码 or 实际应用
        # self.test_func()
        self.setup_schedule()


    # 主界面布局
    def _build_main_page(self):
        page = self.main_page

        # logo
        self.icon_label = ctk.CTkLabel(page, image=self.icon_img, text="")
        self.icon_label.pack(pady=10)

        # title
        self.title_label = ctk.CTkLabel(page, font=self.font_lovelo, text="Alert Analyzer")
        self.title_label.cget("font").configure(size=36)
        self.title_label.pack(pady=10)

        # version
        self.version_label = ctk.CTkLabel(page, font=self.font_microsoft_yahei_bold, text="v" + config.version)
        self.version_label.pack(pady=10)

        # open folder
        self.open_folder_btn = ctk.CTkButton(page, font=self.font_microsoft_yahei_normal, text="打开告警分析目录", image=self.folder_icon, compound="left", command=self.open_folder, width=220)
        self.open_folder_btn.pack(pady=10)

        # rebuild btn
        self.rebuild_btn = ctk.CTkButton(page, font=self.font_microsoft_yahei_normal, text="重做当日报表", image=self.excel_icon, compound="left", command=self.rebuild, width=220)
        self.rebuild_btn.pack(pady=10)

        # status
        self.status_label = ctk.CTkLabel(page, font=self.font_microsoft_yahei_bold, text=self.status_prefix + self.status_nothing)
        self.status_label.pack(pady=10)

        # progress bar
        self.progress_bar = ctk.CTkProgressBar(page, orientation='horizontal', mode='determinate')
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)
        self.progress_bar.pack_forget()
        self.current_step = 0


    # 设置界面布局
    def _build_setting_page(self):
        page = self.setting_page

        # setting page icon
        self.setting_label = ctk.CTkButton(page, image=self.setting_page_icon, width=60, height=60, hover_color=BG_COLOR, text="", fg_color=BG_COLOR, command=self._show_about_dialog)
        # self.setting_label = ctk.CTkLabel(page, image=self.setting_page_icon, text="")
        self.setting_label.pack()

        # start time label
        self.start_time_label = ctk.CTkLabel(page, font=self.font_microsoft_yahei_bold, text="任务开始时间")
        self.start_time_label.pack(pady=20)

        # time select frame
        time_frame = ctk.CTkFrame(page, fg_color="transparent")
        time_frame.pack()

        # hour combo
        hours = [f"{i:02d}" for i in range(24)]
        self.hour_combo = ctk.CTkComboBox(time_frame, values=hours, width=80, state="readonly", justify="center")
        self.hour_combo.set(config.start_time_hour)
        self.hour_combo.pack(side="left", padx=(0, 40))

        # minute combo
        minutes = [f"{i:02d}" for i in range(60)]
        self.minute_combo = ctk.CTkComboBox(time_frame, values=minutes, width=80, state="readonly", justify="center")
        self.minute_combo.set(config.start_time_minute)
        self.minute_combo.pack(side="right")

        # output folder label
        self.output_folder_label = ctk.CTkLabel(page, font=self.font_microsoft_yahei_bold, text="告警分析输出目录")
        self.output_folder_label.pack(pady=20)

        # output folder path frame
        output_folder_frame = ctk.CTkFrame(page, fg_color="transparent")
        output_folder_frame.pack()

        # output folder path
        self.output_folder_path = ctk.CTkButton(output_folder_frame, font=self.font_microsoft_yahei_normal, width=200, text=config.backbone_data_path, state="disabled")
        self.output_folder_path.pack(side="left")

        # output folder path change btn
        self.output_folder_change_btn = ctk.CTkButton(output_folder_frame, image=self.folder_icon, text="", width=20, height=20, hover_color="gray75", command=self._select_output_folder)
        self.output_folder_change_btn.pack(side="right")

        # save setting changes btn
        self.save_setting_changes_btn = ctk.CTkButton(page, font=self.font_microsoft_yahei_bold, width=80, text="保存修改", command=self._save_setting_changes)
        self.save_setting_changes_btn.pack(pady=40)


    def _show_about_dialog(self):
        CTkMessagebox(
            title="About",
            message="每日告警分析自动生成程序 v2.0\nAuthor: Perry Ye\n如有问题请联系我",
            icon="info",
            font=self.font_microsoft_yahei_normal
        )


    def _select_output_folder(self):
        new_dir = filedialog.askdirectory()

        if new_dir:
            self.output_folder_path.configure(text=new_dir)


    def _save_setting_changes(self):
        msg = CTkMessagebox(
            title="确认保存",
            message="确认要保存当前设置吗？",
            icon="question",
            option_1="no",
            option_2="yes"
        )

        user_choice = msg.get()

        if user_choice == "no":
            return

        hour = self.hour_combo.get()
        minute = self.minute_combo.get()
        out_dir = self.output_folder_path.cget("text")

        config.set_start_time(hour, minute)
        config.set_backbone_data_path(out_dir)

        self._setting_snapshot = {
            "hour": self.hour_combo.get(),
            "minute": self.minute_combo.get(),
            "out_path": self.output_folder_path.cget("text")
        }


    # 展示主界面
    def show_main_page(self):
        if self.setting_page.winfo_ismapped() and self._setting_is_ditry():
            msg = CTkMessagebox(
                title="未保存的设置",
                message="你有已修改但未保存的设置\n是否保存后返回?",
                icon="question",
                option_1="取消(继续编辑)",
                option_2="恢复之前设置",
                option_3="保存并返回主界面"
            )

            choice = msg.get()

            if choice == "取消(继续编辑)":
                return
            
            elif choice == "恢复之前设置":
                snapshot = self._setting_snapshot
                self.hour_combo.set(snapshot["hour"])
                self.minute_combo.set(snapshot["minute"])
                self.output_folder_path.configure(text=snapshot["out_path"])
                return
            
            elif choice == "保存并返回主界面":
                self._save_setting_changes()

        self.nav_btn.configure(image=self.setting_icon, command=self.show_setting_page)
        self.main_page.tkraise()


    # 展示设置页面
    def show_setting_page(self):
        self.nav_btn.configure(image=self.back_icon, command=self.show_main_page)
        
        # 保存设置快照, 判断用户是否修改设置
        self._setting_snapshot = {
            "hour": self.hour_combo.get(),
            "minute": self.minute_combo.get(),
            "out_path": self.output_folder_path.cget("text")
        }
        
        self.setting_page.tkraise()

    
    # 检查是否有未保存的设置
    def _setting_is_ditry(self) -> bool:
        if not hasattr(self, "_setting_snapshot"):
            return False

        cur = {
            "hour": self.hour_combo.get(),
            "minute": self.minute_combo.get(),
            "out_path": self.output_folder_path.cget("text")
        }

        return cur != self._setting_snapshot

    
    # 打开告警分析目录
    def open_folder(self):
        os.startfile(config.backbone_data_path)


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
            self.progress_bar.pack(pady=10)
        self.after(0, _show)


    def hide_progress(self):
        def _hide():
            self.progress_bar.pack_forget()
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


    # 预先检查，检查不过不开始任务
    def pre_check(self, day):
        # 前一天的告警分析必须存在
        the_day_before_yesterday = day - timedelta(days=2)
        date_str = the_day_before_yesterday.strftime("%Y%m%d")
        last_day_file_path = os.path.join(config.backbone_data_path, f"{date_str}告警日分析.xlsx")

        if not os.path.exists(last_day_file_path):
            return False, f"{date_str}告警日分析不存在"
        
        return True, None


    # 导出文件/任务开始
    def export_files(self, day=None):
        # 如果未指定日期，则为正常每日报表生成
        if day is None:
            day = datetime.today()

        ok, msg = self.pre_check(day)
        
        if not ok:
            def _warn():
                CTkMessagebox(
                    title="无法开始任务",
                    message=msg,
                    icon="warning",
                    option_1="确认"
                )

                self.set_status(f"生成不成功, 原因: {msg}")
            
            self.after(0, _warn)
            return

        self.show_progress()

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

        target_path = os.path.join(config.data_path, day.strftime("%m-%d"))
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
        self.excel_utils.update_chart(pivot_table_path, os.path.join(config.backbone_data_path, f"{the_day_before_yesterday.strftime("%Y%m%d")}告警日分析.xlsx"), day)

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
            schedule.every().day.at(config.start_time_hour + ":" + config.start_time_minute).do(job)
            while True:
                schedule.run_pending()
                time.sleep(40)

        thread = threading.Thread(target=run_schedule, daemon=True)
        thread.start()


app = App()
app.mainloop()