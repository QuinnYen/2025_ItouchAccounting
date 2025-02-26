import os
import sys
import re
import subprocess
import platform
import zipfile
import shutil
import requests
import logging
from datetime import datetime, timedelta

class ChromeDriverManager:
    def __init__(self, cache_valid_days=7):
        # 判斷是否為執行檔環境
        if getattr(sys, 'frozen', False):
            self.base_path = os.path.dirname(sys.executable)
        else:
            self.base_path = os.path.dirname(os.path.abspath(__file__))
            
        # 建立驅動程式目錄
        self.driver_dir = os.path.join(self.base_path, 'drivers')
        os.makedirs(self.driver_dir, exist_ok=True)
        
        # 緩存資訊檔案
        self.cache_info_file = os.path.join(self.driver_dir, 'driver_info.txt')
        
        # 緩存有效期（天）
        self.cache_valid_days = cache_valid_days

    def get_chrome_version(self):
        """獲取當前系統已安裝的 Chrome 版本"""
        try:
            system = platform.system()
            if system == "Windows":
                # Windows：從註冊表獲取
                try:
                    from winreg import HKEY_CURRENT_USER, OpenKey, QueryValueEx
                    key = OpenKey(HKEY_CURRENT_USER, r'Software\Google\Chrome\BLBeacon')
                    version, _ = QueryValueEx(key, 'version')
                    return version
                except:
                    # 嘗試從安裝路徑獲取
                    paths = [
                        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
                        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
                        os.path.expanduser(r'~\AppData\Local\Google\Chrome\Application\chrome.exe')
                    ]
                    for path in paths:
                        if os.path.exists(path):
                            version_info = subprocess.check_output(f'wmic datafile where name="{path.replace("\\", "\\\\")}" get Version /value', shell=True)
                            match = re.search(r'Version=(\d+\.\d+\.\d+\.\d+)', version_info.decode('utf-8'))
                            if match:
                                return match.group(1)
            elif system == "Darwin":  # macOS
                process = subprocess.Popen(['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '--version'], stdout=subprocess.PIPE)
                version = process.communicate()[0].decode('UTF-8').replace('Google Chrome ', '').strip()
                return version
            elif system == "Linux":
                process = subprocess.Popen(['google-chrome', '--version'], stdout=subprocess.PIPE)
                version = process.communicate()[0].decode('UTF-8').replace('Google Chrome ', '').strip()
                return version
        except Exception as e:
            logging.warning(f"無法獲取 Chrome 版本: {str(e)}")
        
        # 如果無法獲取版本，返回最新版本的下載鏈接
        return None

    def get_compatible_driver_url(self, chrome_version):
        """獲取與 Chrome 版本兼容的 ChromeDriver 下載 URL"""
        if chrome_version:
            major_version = chrome_version.split('.')[0]
            
            # 使用新的 Chrome for Testing 下載 API
            # 參考: https://googlechromelabs.github.io/chrome-for-testing/
            api_url = f"https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_{major_version}"
            try:
                response = requests.get(api_url)
                response.raise_for_status()
                driver_version = response.text.strip()
                
                # 準備下載 URL
                system = platform.system().lower()
                if system == "windows":
                    platform_name = "win32"
                elif system == "darwin":  # macOS
                    platform_name = "mac-x64" if platform.machine() != "arm64" else "mac-arm64"
                else:  # Linux
                    platform_name = "linux64"
                
                download_url = f"https://storage.googleapis.com/chrome-for-testing-public/{driver_version}/{platform_name}/chromedriver-{platform_name}.zip"
                return download_url
            except Exception as e:
                logging.warning(f"使用 Chrome for Testing API 獲取 URL 失敗: {str(e)}")
        
        # 如果無法獲取特定版本，使用穩定版
        try:
            # 獲取最新穩定版本
            response = requests.get("https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE")
            response.raise_for_status()
            stable_version = response.text.strip()
            
            system = platform.system().lower()
            if system == "windows":
                platform_name = "win32"
            elif system == "darwin":  # macOS
                platform_name = "mac-x64" if platform.machine() != "arm64" else "mac-arm64"
            else:  # Linux
                platform_name = "linux64"
            
            download_url = f"https://storage.googleapis.com/chrome-for-testing-public/{stable_version}/{platform_name}/chromedriver-{platform_name}.zip"
            return download_url
        except Exception as e:
            logging.error(f"獲取 ChromeDriver 下載鏈接失敗: {str(e)}")
            return None
    
    def is_cache_valid(self):
        """檢查緩存的驅動程式是否仍然有效"""
        if not os.path.exists(self.cache_info_file):
            return False
            
        try:
            with open(self.cache_info_file, 'r') as f:
                lines = f.readlines()
                if len(lines) >= 2:
                    # 檢查日期
                    cache_date = datetime.strptime(lines[0].strip(), '%Y-%m-%d')
                    if datetime.now() - cache_date > timedelta(days=self.cache_valid_days):
                        return False
                    
                    # 檢查 Chrome 版本是否匹配
                    cached_chrome_version = lines[1].strip()
                    current_chrome_version = self.get_chrome_version()
                    
                    # 只比較主版本號
                    if current_chrome_version:
                        current_major = current_chrome_version.split('.')[0]
                        cached_major = cached_chrome_version.split('.')[0]
                        return current_major == cached_major
                        
            return False
        except Exception:
            return False
    
    def update_cache_info(self, chrome_version):
        """更新緩存信息"""
        with open(self.cache_info_file, 'w') as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d')}\n")
            f.write(f"{chrome_version}\n")
    
    def download_driver(self, url):
        """下載並解壓 ChromeDriver"""
        try:
            # 下載 ZIP 文件
            response = requests.get(url, stream=True)
            response.raise_for_status()
            
            zip_path = os.path.join(self.driver_dir, 'chromedriver.zip')
            with open(zip_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            # 解壓 ZIP 文件
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(self.driver_dir)
            
            # 刪除 ZIP 文件
            os.remove(zip_path)
            
            # 找到解壓後的驅動程式路徑
            driver_name = 'chromedriver.exe' if platform.system() == 'Windows' else 'chromedriver'
            
            # 搜索解壓文件夾找到 chromedriver
            for root, dirs, files in os.walk(self.driver_dir):
                for file in files:
                    if file == driver_name or file == driver_name.replace(".exe", ""):
                        extracted_path = os.path.join(root, file)
                        target_path = os.path.join(self.driver_dir, driver_name)
                        
                        # 如果目標路徑存在，先刪除
                        if os.path.exists(target_path):
                            os.remove(target_path)
                            
                        # 移動檔案到目標位置
                        shutil.move(extracted_path, target_path)
                        
                        # 設置執行權限 (Linux/Mac)
                        if platform.system() != 'Windows':
                            os.chmod(target_path, 0o755)
                        
                        # 清理其他解壓出來的文件夾
                        if root != self.driver_dir:
                            shutil.rmtree(root)
                            
                        return target_path
            
            raise FileNotFoundError(f"無法在解壓後的檔案中找到 {driver_name}")
            
        except Exception as e:
            logging.error(f"下載或解壓 ChromeDriver 失敗: {str(e)}")
            return None
    
    def install(self):
        """安裝與當前 Chrome 版本兼容的 ChromeDriver"""
        driver_path = os.path.join(self.driver_dir, 'chromedriver.exe' if platform.system() == 'Windows' else 'chromedriver')
        
        # 檢查緩存是否有效
        if self.is_cache_valid() and os.path.exists(driver_path):
            logging.info("使用緩存的 ChromeDriver")
            return driver_path
        
        # 獲取 Chrome 版本
        chrome_version = self.get_chrome_version()
        logging.info(f"檢測到 Chrome 版本: {chrome_version}")
        
        # 獲取下載 URL
        download_url = self.get_compatible_driver_url(chrome_version)
        if not download_url:
            raise Exception("無法獲取 ChromeDriver 下載 URL")
        
        logging.info(f"下載 ChromeDriver: {download_url}")
        
        # 下載並解壓驅動程式
        driver_path = self.download_driver(download_url)
        if not driver_path:
            raise Exception("下載或解壓 ChromeDriver 失敗")
        
        # 更新緩存信息
        self.update_cache_info(chrome_version or "unknown")
        
        return driver_path