import os.path

import requests
import time
from datetime import datetime, timedelta
from config_utils import Config_Utils
from logger import Logger

config = Config_Utils()
logger = Logger(True, config.log_path)


def safe_post(url, headers=None, json=None, retries=5, backoff=2):
    for attempt in range(1, retries + 1):
        try:
            # logger.log(f"[{attempt}/{retries}] 正在请求: {url}")
            response = requests.post(url, headers=headers, json=json, timeout=10)
            response.raise_for_status()
            return response
        except requests.RequestException as e:
            logger.log(f"[{attempt}/{retries}] 请求失败: {e}")
            if attempt == retries:
                logger.log(f"请求连续失败 {retries} 次，程序终止")
                raise
            time.sleep(backoff)

class AlertUtils:
    def __init__(self):
        super().__init__()

        self.data = config._config
        self.session_id = None

    # test func
    def set_session_id(self, day):
        logger.log("开始设置session id")

        data = self.data['set_session_id']

        url = data['url']
        headers = data['headers']
        payload = data['payload']

        session_id = int(time.time())
        today = day
        yesterday = today - timedelta(days=1)

        payload['sessionId'] = session_id
        payload['dateBean']['startTime'] = yesterday.strftime("%Y-%m-%d") + " 00:00:00"
        payload['dateBean']['endTime'] = yesterday.strftime("%Y-%m-%d") + " 23:59:59"

        safe_post(url, headers, payload)

        self.session_id = session_id
        logger.log(f"session id: {session_id}")


    # export csv file
    def export_csv_files(self):
        logger.log("正在从系统导出告警文件")

        data = self.data['export_csv_files']

        url = data['url']
        headers = data['headers']
        payload = data['payload']

        payload['sessionId'] = self.session_id

        res = safe_post(url, headers, payload)

        logger.log(res.json()['message'])


    def check_export_progress(self):
        data = self.data['check_export_progress']

        url = data['url']
        headers = data['headers']
        payload = data['payload']

        res = safe_post(url, headers, payload)

        return res.json()['dataObject']


    def download_files(self, file_src):
        logger.log(f"开始下载告警文件压缩包，收到的文件地址: {file_src}")

        data = self.data['download_files']

        url = data['url'] + file_src
        headers = data['headers']

        logger.log(f"文件下载链接: {url}")

        res = requests.get(url, headers=headers, stream=True)

        zip_name = file_src.split('/')[-1]
        with open(os.path.join(config.zips_path, zip_name), 'wb') as f:
            f.write(res.content)

        logger.log("文件下载完毕，请在文件夹中查看！")

