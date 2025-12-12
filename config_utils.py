import json
import os
import tempfile

CONFIG_PATH = './config/cfg.json'

class Config_Utils:
    def __init__(self):
        self._config = self._load_config()

        self.version = self._config["version"]

        self.zips_path = self._config["zips_path"]
        self.data_path = self._config["data_path"]

        self.setting_icon_path = self._config["setting_icon_path"]
        self.logo_icon_path = self._config["logo_icon_path"]
        self.folder_icon_path = self._config["folder_icon_path"]
        self.excel_icon_path = self._config["excel_icon_path"]
        self.back_icon_path = self._config["back_icon_path"]

        self.font_path = self._config["font_path"]
        self.log_path = self._config["log_path"]
        self.backbone_data_path = self._config["backbone_data_path"]

        self.cities = self._config["cities"]

        self.start_time_hour = self._config["start_time_hour"]
        self.start_time_minute = self._config["start_time_minute"]


    def _load_config(self):
        # 如果配置文件不存在，直接报错
        if not os.path.exists(CONFIG_PATH):
            raise FileNotFoundError("配置文件不存在，请检查！")
        
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
        
    
    def _save_config(self):
        try:
            # 生成临时文件
            dir_name = os.path.dirname(CONFIG_PATH) or "."
            fd, tmp_path = tempfile.mkstemp(prefix="config_", suffix=".tmp", dir=dir_name)

            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(self._config, f, ensure_ascii=False, indent=4)
                f.flush()
                os.fsync(f.fileno())

            os.replace(tmp_path, CONFIG_PATH)

        except Exception as e:
            try:
                os.remove(tmp_path)
            except:
                pass
            raise Exception(f"保存配置文件时出错! 原因: {e}")


    def set_backbone_data_path(self, new_path: str):
        old_path = self.backbone_data_path
        self.backbone_data_path = new_path
        self._config["backbone_data_path"] = new_path

        self._save_config()

    
    def set_start_time(self, hour: str, minute: str):
        self.start_time_hour = hour
        self.start_time_minute = minute
        
        self._config["start_time_hour"] = hour
        self._config["start_time_minute"] = minute
        self._save_config()

