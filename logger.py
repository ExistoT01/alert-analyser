import os
from datetime import datetime

class Logger:
    def __init__(self, log_to_file=False, log_path="log.txt"):
        self.log_to_file = log_to_file
        self.log_path = log_path

        if log_to_file:
            os.makedirs(os.path.dirname(log_path), exist_ok=True)


    def log(self, msg):
        log_time_prefix = datetime.now().strftime("[%y-%m-%d %H:%M:%S]")
        log_text = f"{log_time_prefix} {msg}"

        if self.log_to_file:
            with open(self.log_path, "a", encoding="utf-8") as f:
                f.write(log_text + "\n")

        print(log_text)

