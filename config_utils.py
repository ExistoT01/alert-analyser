import json
import os, sys
import tempfile
import shutil

# CONFIG_PATH = './config/cfg.json'
APP_NAME = "AlertAnalyser"


def _bundle_base_dir() -> str:
    if getattr(sys, "_MEIPASS", None):
        return sys._MEIPASS
    return os.path.abspath(os.path.dirname(__file__))


def _storage_base_dir() -> str:
    appdata = os.getenv("APPDATA") or os.path.expanduser("~")
    base = os.path.join(appdata, APP_NAME)
    os.makedirs(base, exist_ok=True)
    return base


def _normalize(p: str) -> str:
    p = (p or "").replace("/", os.sep)
    if p.startswith("." + os.sep):
        p = p[2:]
    return p


def resolve_resource_path(p: str) -> str:
    p = _normalize(p)
    if os.path.isabs(p):
        return p
    return os.path.join(_bundle_base_dir(), p)


def resolve_storage_path(p: str) -> str:
    p = _normalize(p)
    if os.path.isabs(p):
        return p
    return os.path.join(_storage_base_dir(), p)


def user_config_path() -> str:
    return os.path.join(_storage_base_dir(), "config", "cfg.json")


def default_config_template_path() -> str:
    return resolve_resource_path(os.path.join("config", "cfg.json"))


class Config_Utils:
    def __init__(self):
        self.config_path = user_config_path()
        os.makedirs(os.path.dirname(self.config_path), exist_ok=True)

        if not os.path.exists(self.config_path):
            tpl = default_config_template_path()
            if not os.path.exists(tpl):
                raise FileNotFoundError(f"默认配置模板不存在: {tpl}")
            shutil.copyfile(tpl, self.config_path)
            
        self._config = self._load_config()


        # ====== 写入类（持久化）======
        self.zips_path = resolve_storage_path(self._config["zips_path"])
        self.data_path = resolve_storage_path(self._config["data_path"])
        self.log_path  = resolve_storage_path(self._config["log_path"])

        # 确保目录存在
        os.makedirs(self.zips_path, exist_ok=True)
        os.makedirs(self.data_path, exist_ok=True)
        os.makedirs(os.path.dirname(self.log_path), exist_ok=True)

        # ====== 资源类（只读）======
        self.app_icon_path     = resolve_resource_path(self._config["app_icon_path"])
        self.light_icon_path   = resolve_resource_path(self._config["light_icon_path"])
        self.dark_icon_path    = resolve_resource_path(self._config["dark_icon_path"])
        self.setting_icon_path = resolve_resource_path(self._config["setting_icon_path"])
        self.logo_icon_path    = resolve_resource_path(self._config["logo_icon_path"])
        self.folder_icon_path  = resolve_resource_path(self._config["folder_icon_path"])
        self.excel_icon_path   = resolve_resource_path(self._config["excel_icon_path"])
        self.back_icon_path    = resolve_resource_path(self._config["back_icon_path"])

        self.version = self._config["version"]
       
        self.font_path = resolve_resource_path(self._config["font_path"])
        self.backbone_data_path = self._config["backbone_data_path"]

        self.cities = self._config["cities"]

        self.start_time_hour = self._config["start_time_hour"]
        self.start_time_minute = self._config["start_time_minute"]


    def _load_config(self):
        # 如果配置文件不存在，直接报错
        if not os.path.exists(self.config_path):
            raise FileNotFoundError("配置文件不存在，请检查！")
        
        with open(self.config_path, "r", encoding="utf-8") as f:
            return json.load(f)
        
    
    def _save_config(self):
        try:
            # 生成临时文件
            dir_name = os.path.dirname(self.config_path) or "."
            fd, tmp_path = tempfile.mkstemp(prefix="config_", suffix=".tmp", dir=dir_name)

            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(self._config, f, ensure_ascii=False, indent=4)
                f.flush()
                os.fsync(f.fileno())

            os.replace(tmp_path, self.config_path)

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

