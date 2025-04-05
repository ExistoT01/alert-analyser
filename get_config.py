import json

def read_cfg():
    cfg_path = "./config/cfg.json"
    with open(cfg_path, "r", encoding="utf-8") as f:
        return json.load(f)
