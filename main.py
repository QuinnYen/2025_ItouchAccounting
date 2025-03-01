import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
import keyring
import time
import os, sys
from excel_exporter import ExcelExporter
import logging
import threading
from concurrent.futures import ThreadPoolExecutor
from selenium.webdriver.common.action_chains import ActionChains
from chrome_manager import ChromeDriverManager

class ErrorLogger:
    def __init__(self, log_dir='logs'):
        # 判斷是否為執行檔
        if getattr(sys, 'frozen', False):
            # 如果是執行檔，使用執行檔所在目錄
            base_path = os.path.dirname(sys.executable)
        else:
            # 如果是一般 Python 腳本，使用腳本所在目錄
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        # 建立完整的日誌目錄路徑
        self.log_dir = os.path.join(base_path, log_dir)
        os.makedirs(self.log_dir, exist_ok=True)
        
        # 設置日誌檔案
        today = time.strftime('%Y%m%d')
        log_file = os.path.join(self.log_dir, f'error_{today}.log')
        
        # 配置日誌記錄器
        self.logger = logging.getLogger('itouch_crawler')
        self.logger.setLevel(logging.ERROR)
        
        # 移除所有既有的處理器
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)
        
        # 檔案處理器
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.ERROR)
        
        # 格式化
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(formatter)
        
        self.logger.addHandler(file_handler)
        
        # 清理舊的日誌檔案（保留最近30天的紀錄）
        self.cleanup_old_logs(30)
    
    def cleanup_old_logs(self, days_to_keep):
        """清理超過指定天數的舊日誌檔案"""
        try:
            current_time = time.time()
            for filename in os.listdir(self.log_dir):
                if filename.startswith('error_') and filename.endswith('.log'):
                    filepath = os.path.join(self.log_dir, filename)
                    file_time = os.path.getmtime(filepath)
                    
                    # 如果檔案超過指定天數就刪除
                    if current_time - file_time > days_to_keep * 86400:  # 86400 = 24 * 60 * 60 秒
                        os.remove(filepath)
        except Exception as e:
            self.logger.error(f"清理舊日誌檔案時發生錯誤: {str(e)}")
    
    def log_error(self, error_message, exception=None):
        """記錄錯誤到日誌檔案"""
        if exception:
            self.logger.error(f"{error_message}: {str(exception)}", exc_info=True)
        else:
            self.logger.error(error_message)

class ItouchCrawler:
    def __init__(self, root):
        # 加入開發人員控制變數(開發人員模式：True=顯示瀏覽器，False=無頭模式)
        self.DEVELOPER_MODE = False 
        
        self.root = root
        self.root.title('iTouch-會計帳目自動抓取程式 v3')
        
        # 顯示載入提示
        self.loading_label = ttk.Label(root, text="正在初始化...", font=('Helvetica', 12))
        self.loading_label.place(relx=0.5, rely=0.5, anchor='center')
        
        # 初始化錯誤記錄器
        self.error_logger = ErrorLogger()
        
        # 在背景執行初始化
        self.initialization_thread = threading.Thread(target=self.initialize_background)
        self.initialization_thread.start()
        
        # 初始化變數
        self.driver = None
        self.is_logged_in = False
        self.plan_codes = []
        
        # 服務名稱和金鑰名稱
        self.service_id = 'itouch_crawler'
        self.username_key = 'saved_username'
        
        # 綁定初始化完成事件
        self.root.bind('<<InitializationComplete>>', lambda e: self.on_initialization_complete())
        self.root.bind('<<InitializationError>>', lambda e: self.on_initialization_error())
        
        # 定期檢查初始化狀態
        self.check_initialization()

    def initialize_background(self):
        """在背景執行初始化工作"""
        try:
            # 預先準備瀏覽器選項
            self.prepare_browser_options()
            
            # 完成後通知主線程
            self.root.event_generate('<<InitializationComplete>>', when='tail')
            
        except Exception as e:
            self.initialization_error = str(e)
            self.root.event_generate('<<InitializationError>>', when='tail')

    def prepare_browser_options(self):
        """預先準備瀏覽器選項"""
        self.options = webdriver.ChromeOptions()
        
        # 根據開發人員模式決定是否使用無頭模式
        if not self.DEVELOPER_MODE:
            self.options.add_argument('--headless=new')
            
        self.options.add_argument('--disable-gpu')
        self.options.add_argument('--no-sandbox')
        self.options.add_argument('--disable-dev-shm-usage')
        self.options.add_argument('--window-size=1920,1080')
        
        # 優化性能的額外選項
        self.options.add_argument('--disable-extensions')
        self.options.add_argument('--disable-notifications')
        self.options.add_argument('--disable-logging')
        self.options.add_argument('--log-level=3')
        self.options.page_load_strategy = 'eager'  # 加快頁面載入

    def check_initialization(self):
        """檢查初始化狀態"""
        if not self.initialization_thread.is_alive():
            self.loading_label.destroy()
            self.setup_gui()
        else:
            self.root.after(100, self.check_initialization)

    def on_initialization_complete(self):
        """初始化完成的處理"""
        self.loading_label.destroy()
        self.setup_gui()
        self.update_status("程式初始化完成")

    def on_initialization_error(self):
        """初始化錯誤的處理"""
        self.loading_label.configure(text=f"初始化失敗: {self.initialization_error}")
        self.error_logger.log_error(f"初始化失敗: {self.initialization_error}")

    def initialize_driver(self):
        """使用自定義 ChromeDriver 管理器初始化瀏覽器"""
        if not self.driver:
            try:
                # 使用自定義管理器獲取驅動程式路徑
                driver_manager = ChromeDriverManager()
                driver_path = driver_manager.install()
                
                # 添加更多兼容性選項
                self.options.add_argument('--disable-dev-shm-usage')
                self.options.add_argument('--no-sandbox')
                self.options.add_argument('--disable-web-security')
                self.options.add_argument('--allow-running-insecure-content')
                self.options.add_argument('--ignore-certificate-errors')
                
                # 避免被檢測為自動化軟體
                self.options.add_argument('--disable-blink-features=AutomationControlled')
                self.options.add_experimental_option('excludeSwitches', ['enable-automation'])
                self.options.add_experimental_option('useAutomationExtension', False)
                
                # 設定 chromedriver 執行環境
                service = Service(driver_path)
                
                # 重試機制
                max_retries = 3
                retry_count = 0
                
                while retry_count < max_retries:
                    try:
                        self.driver = webdriver.Chrome(service=service, options=self.options)
                        
                        # 隱藏自動化控制特徵
                        self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
                            'source': '''
                                Object.defineProperty(navigator, 'webdriver', {
                                    get: () => undefined
                                });
                            '''
                        })
                        
                        # 設定頁面加載超時
                        self.driver.set_page_load_timeout(30)
                        
                        # 當在無頭模式時才禁用登入按鈕
                        if not self.DEVELOPER_MODE:
                            self.login_button.config(state=tk.DISABLED)
                            
                        # 成功初始化，跳出迴圈
                        break
                        
                    except Exception as e:
                        retry_count += 1
                        if retry_count >= max_retries:
                            raise e
                        else:
                            self.update_status(f"初始化瀏覽器失敗，正在重試 ({retry_count}/{max_retries})...")
                            time.sleep(2)  # 等待2秒後重試
                    
            except Exception as e:
                self.error_logger.log_error("瀏覽器初始化失敗", e)
                self.update_status("瀏覽器初始化失敗，請參考錯誤日誌", True)

    def login_and_query(self):
        """優化的登入查詢流程"""
        def background_login():
            if self.login():
                if self.navigate_to_query():
                    self.root.event_generate('<<LoginSuccess>>', when='tail')
                else:
                    self.root.event_generate('<<LoginError>>', when='tail')
            else:
                self.root.event_generate('<<LoginError>>', when='tail')

        # 在背景執行登入
        threading.Thread(target=background_login).start()

    def update_status(self, message, is_error=False):
        """更新狀態訊息"""
        timestamp = time.strftime('%H:%M:%S')
        # 錯誤訊息
        if is_error and not message.startswith(("請先選擇學年", "請至少選擇一個計畫編號", "登入失敗", "請先登入系統")):
            formatted_message = f'[{timestamp}] 發生錯誤，詳細資訊請查看錯誤記錄檔\n'
        else:
            formatted_message = f'[{timestamp}] {message}\n'
            
        self.message_text.insert(tk.END, formatted_message)
        if is_error:
            # 計算標籤的起始和結束位置
            last_line_start = self.message_text.index("end-1c linestart")
            last_line_end = self.message_text.index("end-1c")
            
            # 只將紅色標記套用在這一行
            self.message_text.tag_add('error', last_line_start, last_line_end)
            self.message_text.tag_config('error', foreground='red')
            
        self.message_text.see(tk.END)
        self.root.update()

    def setup_gui(self):
        """設置GUI介面"""
        # 定義大型按鈕樣式
        style = ttk.Style()
        style.configure('Large.TButton', font=('Helvetica', 12, 'bold'), padding=5)
        
        # 設定視窗大小
        window_width = 1000
        window_height = 600

        # 計算視窗置中的位置
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)

        # 設定視窗大小和位置
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        
        # 建立左側框架
        self.left_frame = ttk.Frame(self.root, padding="10")
        self.left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 建立右側框架
        self.right_frame = ttk.Frame(self.root, padding="10")
        self.right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 建立登入框架 (左側框架)
        self.login_frame = ttk.LabelFrame(self.left_frame, text="登入資訊", padding="10")
        self.login_frame.grid(row=0, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # 建立GUI元件
        ttk.Label(self.login_frame, text='帳號:').grid(row=0, column=0, padx=5, pady=5)
        self.username = ttk.Entry(self.login_frame)
        self.username.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(self.login_frame, text='密碼:').grid(row=1, column=0, padx=5, pady=5)
        self.password = ttk.Entry(self.login_frame, show='*')
        self.password.grid(row=1, column=1, padx=5, pady=5)
        
        # 初始化記住帳密的變數
        self.remember_var = tk.BooleanVar()
        
        # 記住帳密的核取方塊
        self.remember_checkbox = ttk.Checkbutton(self.login_frame, text='記住帳密', 
                                              variable=self.remember_var)
        self.remember_checkbox.grid(row=2, column=0, columnspan=2, pady=5)
        
        # 修改登入按鈕的宣告
        self.login_button = ttk.Button(self.login_frame, text='登入帳號 並 執行查詢', 
                                command=self.login_and_query,
                                style='Large.TButton')
        self.login_button.grid(row=3, column=0, columnspan=2, padx=5)
        
        # 新增執行中標籤（初始隱藏）
        self.running_label = ttk.Label(self.login_frame, text="爬蟲程式執行中...", 
                                foreground='blue')
        self.running_label.grid(row=3, column=0, columnspan=2, padx=5)
        self.running_label.grid_remove()
        
        # 在 login_frame 中添加重啟按鈕（初始時隱藏）
        self.restart_button = ttk.Button(self.login_frame, text='關閉爬蟲程式', 
                                   command=self.restart_program,
                                   style='Large.TButton')
        self.restart_button.grid(row=3, column=2, padx=5)
        self.restart_button.grid_remove()
        
        # 學年選擇框架 (左側框架)
        self.year_frame = ttk.LabelFrame(self.left_frame, text="查詢設定", padding="10")
        self.year_frame.grid(row=1, column=0, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        ttk.Label(self.year_frame, text='學年:').grid(row=0, column=0, padx=5, pady=5)
        self.year_select = ttk.Combobox(self.year_frame, state='readonly')
        self.year_select.grid(row=0, column=1, padx=5, pady=5)
        
        # 選擇提示標籤和查詢按鈕的容器框架
        self.select_frame = ttk.Frame(self.year_frame)
        self.select_frame.grid(row=5, column=0, columnspan=3, padx=5, pady=5)
        
        # 選擇提示標籤
        self.select_label = ttk.Label(self.select_frame, text="選擇'學年度'和'計畫編號'後，按下右側'查詢按鈕'", 
                                foreground='blue', font=('TkDefaultFont', 12))
        self.select_label.grid(row=0, column=0, padx=5)
        self.select_label.grid_remove()
        
        # 計畫編號框架
        self.plan_frame = ttk.LabelFrame(self.year_frame, text="計畫編號設定", padding="5")
        self.plan_frame.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky=(tk.W, tk.E))
        
        # 計畫編號輸入區
        ttk.Label(self.plan_frame, text='計畫編號:').grid(row=0, column=0, padx=5, pady=5)
        self.plan_code_entry = ttk.Entry(self.plan_frame)
        self.plan_code_entry.grid(row=0, column=1, padx=5, pady=5)
        self.plan_code_entry.bind('<Return>', lambda e: self.add_plan_code())
        
        # 計畫按鈕
        self.add_plan_button = ttk.Button(self.plan_frame, text='加入計畫', command=self.add_plan_code)
        self.add_plan_button.grid(row=0, column=2, padx=5, pady=5)
        
        # 計畫編號清單(使用 Listbox 支援多重選擇)
        self.plan_codes_list = tk.Listbox(self.plan_frame, selectmode=tk.MULTIPLE, height=10)
        self.plan_codes_list.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E))
        
        # 計畫編號捲軸
        scrollbar = ttk.Scrollbar(self.plan_frame, orient="vertical", command=self.plan_codes_list.yview)
        scrollbar.grid(row=1, column=3, sticky=(tk.N, tk.S))
        self.plan_codes_list.configure(yscrollcommand=scrollbar.set)
        
        # 移除選取計畫按鈕
        self.remove_plan_button = ttk.Button(self.plan_frame, text='移除選取計畫', 
                                           command=self.remove_selected_plans)
        self.remove_plan_button.grid(row=2, column=2, padx=5, pady=5)
        
        # 全選按鈕
        self.select_all_button = ttk.Button(self.plan_frame, text='全選', 
                                      command=self.select_all_plans)
        self.select_all_button.grid(row=2, column=0, padx=5, pady=5)

        # 取消全選按鈕
        self.deselect_all_button = ttk.Button(self.plan_frame, text='取消全選', 
                                        command=self.deselect_all_plans)
        self.deselect_all_button.grid(row=2, column=1, padx=5, pady=5)

        # 按鈕框架 (左側框架)
        self.button_frame = ttk.Frame(self.left_frame)
        self.button_frame.grid(row=2, column=0, pady=10)
        
        # 修改右側框架佈局
        self.right_container = ttk.Frame(self.right_frame)
        self.right_container.grid(row=0, column=0, pady=10, padx=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 訊息文字框
        self.message_text = scrolledtext.ScrolledText(self.right_container, height=25, width=60)
        self.message_text.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 按鈕框架
        button_container = ttk.Frame(self.right_container)
        button_container.grid(row=1, column=0, columnspan=2, pady=(0, 5), sticky=(tk.W, tk.E))
        
        # 查詢按鈕 (左側)
        self.query_button = ttk.Button(button_container, text='查詢計畫 並 匯出報表',
                                   command=self.select_year_and_report,
                                   style='Large.TButton')
        self.query_button.grid(row=0, column=0, padx=(0, 5), sticky=(tk.W))
        self.query_button.grid_remove()  # 初始時隱藏按鈕
        
        # 開啟報表位置按鈕 (右側)
        self.open_export_button = ttk.Button(button_container, text='報表位置',
                                         command=self.open_export_folder,
                                         style='Large.TButton')
        self.open_export_button.grid(row=0, column=1, sticky=(tk.E))

        # 設定列/行權重
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.right_frame.grid_columnconfigure(0, weight=1)
        self.right_frame.grid_rowconfigure(0, weight=1)
        self.right_container.grid_columnconfigure(0, weight=1)
        self.right_container.grid_rowconfigure(0, weight=1)
        
        # 載入儲存的帳密和計畫編號
        self.load_credentials()
        self.load_plan_codes()
        self.refresh_plan_codes_list()
        self.excel_exporter = ExcelExporter()

    def login(self):
        """執行登入操作"""
        try:
            # 禁用帳密輸入
            self.username.config(state='disabled')
            self.password.config(state='disabled')
            self.remember_checkbox.config(state='disabled')
            
            self.login_button.grid_remove()  # 隱藏登入按鈕
            self.running_label.grid()  # 顯示執行中標籤
            self.restart_button.grid()  # 顯示重啟按鈕
            self.root.update()  # 立即更新界面
            
            self.initialize_driver()
            self.driver.get('https://itouch.cycu.edu.tw/home/')
            
            self.update_status("正在登入系統...")
            
            # 等待登入表單出現
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "UserNm"))
            )
            
            # 輸入帳號密碼
            username_input = self.driver.find_element(By.NAME, "UserNm")
            password_input = self.driver.find_element(By.NAME, "UserPasswd")
            
            username_input.send_keys(self.username.get())
            password_input.send_keys(self.password.get())
            
            # 點擊登入按鈕
            login_button = self.driver.find_element(By.NAME, "Submit")
            login_button.click()
            
            # 等待登入後的元素出現
            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "app-header__logo"))
                )
                self.update_status("登入成功")
                self.is_logged_in = True
                
                # 如果登入成功,根據核取方塊狀態儲存認證資訊
                self.save_credentials()
                
                # 登入成功時重置重試計數
                if hasattr(self, 'retry_count'):
                    self.retry_count = 0
                    
                return True
                
            except TimeoutException:
                self.update_status("登入失敗，請檢查帳號密碼", True)
                self.is_logged_in = False
                # 登入失敗時恢復登入按鈕
                self.running_label.grid_remove()
                self.login_button.grid()
                self.restart_button.grid_remove()  # 發生錯誤時隱藏重啟按鈕
                return False
                
        except Exception as e:
            self.error_logger.log_error("登入過程發生錯誤", e)
            self.update_status("登入失敗", True)
            self.is_logged_in = False
            # 發生錯誤時恢復登入按鈕
            self.running_label.grid_remove()
            self.login_button.grid()
            self.restart_button.grid_remove()  # 發生錯誤時隱藏重啟按鈕
            return False

    def save_credentials(self):
        """儲存認證資訊"""
        if self.remember_var.get():
            username = self.username.get()
            password = self.password.get()
            keyring.set_password(self.service_id, self.username_key, username)
            keyring.set_password(self.service_id, username, password)
        else:
            try:
                saved_username = keyring.get_password(self.service_id, self.username_key)
                if saved_username:
                    keyring.delete_password(self.service_id, self.username_key)
                    keyring.delete_password(self.service_id, saved_username)
            except keyring.errors.PasswordDeleteError:
                pass

    def load_credentials(self):
        """載入儲存的認證資訊"""
        try:
            saved_username = keyring.get_password(self.service_id, self.username_key)
            if saved_username:
                saved_password = keyring.get_password(self.service_id, saved_username)
                if saved_password:
                    self.username.insert(0, saved_username)
                    self.password.insert(0, saved_password)
                    self.remember_var.set(True)
        except keyring.errors.PasswordDeleteError:
            pass

    def load_plan_codes(self):
        """讀取本地儲存的計畫編號清單"""
        # 判斷是否為執行檔
        if getattr(sys, 'frozen', False):
            # 如果是執行檔，使用執行檔所在目錄
            base_path = os.path.dirname(sys.executable)
        else:
            # 如果是一般 Python 腳本，使用腳本所在目錄
            base_path = os.path.dirname(__file__)
            
        plans_path = os.path.join(base_path, 'plan_codes.txt')
        self.plan_codes = []
        
        # 如果檔案不存在，建立一個空的檔案
        if not os.path.exists(plans_path):
            try:
                with open(plans_path, 'w', encoding='utf-8') as f:
                    pass  # 建立空檔案
            except Exception as e:
                self.error_logger.log_error("建立計畫編號檔案失敗", e)
                return

        try:
            with open(plans_path, 'r', encoding='utf-8') as f:
                for line in f:
                    # 去除空白並檢查是否為註解或空行
                    code = line.strip()
                    if code and not code.startswith('#') and code not in self.plan_codes:
                        self.plan_codes.append(code)
        except Exception as e:
            self.error_logger.log_error("讀取計畫編號檔案失敗", e)
            self.update_status("讀取計畫編號檔案失敗", True)

    def save_plan_codes(self):
        """儲存計畫編號清單到本地檔案"""
        # 判斷是否為執行檔
        if getattr(sys, 'frozen', False):
            # 如果是執行檔，使用執行檔所在目錄
            base_path = os.path.dirname(sys.executable)
        else:
            # 如果是一般 Python 腳本，使用腳本所在目錄
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        plans_path = os.path.join(base_path, 'plan_codes.txt')
        
        try:
            with open(plans_path, 'w', encoding='utf-8') as f:
                for code in self.plan_codes:
                    f.write(f"{code}\n")
            self.update_status("計畫編號已儲存")
        except Exception as e:
            self.error_logger.log_error("儲存計畫編號檔案失敗", e)
            self.update_status("儲存計畫編號檔案失敗", True)

    def add_plan_code(self):
        """新增計畫編號到清單"""
        code = self.plan_code_entry.get().strip().upper()
        if code and code not in self.plan_codes:
            self.plan_codes.append(code)
            self.plan_codes_list.insert(tk.END, code)
            self.save_plan_codes()
        self.plan_code_entry.delete(0, tk.END)

    def refresh_plan_codes_list(self):
        """更新計畫編號清單顯示"""
        self.plan_codes_list.delete(0, tk.END)
        for code in self.plan_codes:
            self.plan_codes_list.insert(tk.END, code)

    def get_selected_plan_codes(self):
        """獲取已選擇的計畫編號"""
        return [self.plan_codes_list.get(idx) for idx in self.plan_codes_list.curselection()]

    def remove_selected_plans(self):
        """移除選取的計畫編號"""
        selected_indices = self.plan_codes_list.curselection()
        if not selected_indices:
            return
        # 由大到小排序，以避免刪除時索引變化的問題
        selected_indices = sorted(selected_indices, reverse=True)
        # 移除選取的項目
        for idx in selected_indices:
            code = self.plan_codes_list.get(idx)
            self.plan_codes.remove(code)
            self.plan_codes_list.delete(idx)
        self.save_plan_codes()

    def select_all_plans(self):
        """全選所有計畫編號"""
        self.plan_codes_list.select_set(0, tk.END)
        self.update_status("已全選所有計畫編號")

    def deselect_all_plans(self):
        """取消全選所有計畫編號"""
        self.plan_codes_list.selection_clear(0, tk.END)
        self.update_status("已取消全選所有計畫編號")

    def navigate_to_query(self):
        """導航到會計經費查詢系統"""
        if not self.is_logged_in:
            self.update_status("請先登入系統", True)
            return False
            
        try:
            # 確保重試計數器已初始化
            if not hasattr(self, 'retry_count'):
                self.retry_count = 0
                
            # 檢查重試次數是否超過限制
            if self.retry_count >= 2:  # 最多重試2次
                self.update_status("導航重試次數已達上限，請手動重新啟動程式", True)
                # 重置重試計數
                self.retry_count = 0
                # 清理瀏覽器資源
                if self.driver:
                    self.driver.quit()
                    self.driver = None
                # 重置登入狀態
                self.is_logged_in = False
                # 恢復介面狀態
                self.running_label.grid_remove()
                self.login_button.grid()
                self.restart_button.grid_remove()
                self.username.config(state='normal')
                self.password.config(state='normal')
                self.remember_checkbox.config(state='normal')
                return False

            # 確保頁面載入完成
            wait = WebDriverWait(self.driver, 10)
            
            # ====== 處理網站地圖按鈕 ======
            # 嘗試多種可能的網站地圖按鈕選擇器
            map_selectors = [
                (By.CLASS_NAME, "info"),
                (By.XPATH, "//button[contains(., '網站地圖')]"),
                (By.XPATH, "//a[contains(., '網站地圖')]"),
                (By.CSS_SELECTOR, ".info")
            ]
            
            # 依序嘗試不同的選擇器
            map_button = None
            for selector in map_selectors:
                try:
                    map_button = wait.until(EC.element_to_be_clickable(selector))
                    if map_button:
                        break
                except:
                    continue
                    
            if not map_button:
                raise Exception("找不到網站地圖按鈕")
                
            # 安全點擊地圖按鈕
            try:
                map_button.click()
            except:
                try:
                    # 嘗試使用JavaScript點擊
                    self.driver.execute_script("arguments[0].click();", map_button)
                except:
                    # 如果仍然失敗，嘗試第三種方法
                    action = webdriver.ActionChains(self.driver)
                    action.move_to_element(map_button).click().perform()
            
            self.update_status("點擊網站地圖")
            
            # ====== 等待選單出現並點擊會計室 ======
            # 給予足夠時間讓選單出現
            time.sleep(1)
            
            # 嘗試多種可能的會計室選擇器
            accounting_selectors = [
                "//li[contains(@class, 'menuLevel1Folder') and contains(., '會計室')]",
                "//div[contains(@class, 'menuItem')]//a[contains(text(), '會計室')]",
                "//span[contains(text(), '會計室')]",
                "//a[contains(text(), '會計室')]"
            ]
            
            # 重新尋找每次尋找元素，避免stale element
            for selector in accounting_selectors:
                try:
                    # 等待元素可見
                    accounting_element = wait.until(
                        EC.visibility_of_element_located((By.XPATH, selector))
                    )
                    # 嘗試點擊
                    try:
                        accounting_element.click()
                        self.update_status("點擊會計室")
                        break
                    except:
                        try:
                            # JavaScript點擊
                            self.driver.execute_script("arguments[0].click();", accounting_element)
                            self.update_status("點擊會計室 (JS)")
                            break
                        except Exception as e:
                            self.error_logger.log_error(f"點擊會計室失敗 (使用選擇器 {selector}): {str(e)}")
                            continue
                except:
                    continue
            else:
                # 如果所有選擇器都失敗，報告錯誤
                raise Exception("無法找到或點擊會計室選單")
                
            # ====== 等待經費請款系統選單出現 ======
            # 給予足夠時間讓下拉選單顯示
            time.sleep(1)
            
            payment_selectors = [
                "//li[contains(@class, 'menuLevel2Folder') and contains(., '經費請款系統')]",
                "//li[contains(text(), '經費請款系統')]",
                "//a[contains(text(), '經費請款系統')]"
            ]
            
            # 嘗試多個選擇器
            for selector in payment_selectors:
                try:
                    payment_element = wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    try:
                        payment_element.click()
                        self.update_status("點擊經費請款系統")
                        break
                    except:
                        try:
                            self.driver.execute_script("arguments[0].click();", payment_element)
                            self.update_status("點擊經費請款系統 (JS)")
                            break
                        except:
                            continue
                except:
                    continue
            else:
                raise Exception("無法找到或點擊經費請款系統選單")
            
            # ====== 等待請款查詢系統選項出現 ======
            # 給予足夠時間讓選單顯示
            time.sleep(1)
            
            query_selectors = [
                "//a[contains(., '請款.授權.查詢系統')]",
                "//a[contains(text(), '請款') and contains(text(), '授權') and contains(text(), '查詢')]",
                "//a[contains(text(), '請款.授權.查詢')]"
            ]
            
            # 嘗試多個選擇器
            for selector in query_selectors:
                try:
                    query_element = wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    try:
                        query_element.click()
                        self.update_status("點擊請款.授權.查詢系統")
                        break
                    except:
                        try:
                            self.driver.execute_script("arguments[0].click();", query_element)
                            self.update_status("點擊請款.授權.查詢系統 (JS)")
                            break
                        except:
                            continue
                except:
                    continue
            else:
                raise Exception("無法找到或點擊請款.授權.查詢系統選單")
            
            # 處理新分頁
            WebDriverWait(self.driver, 10).until(lambda d: len(d.window_handles) > 1)
            self.driver.switch_to.window(self.driver.window_handles[-1])
            
            # ====== 點擊會計經費查詢 ======
            finance_selectors = [
                "//a[contains(., '會計經費查詢')]",
                "//a[text()='會計經費查詢']"
            ]
            
            # 嘗試多個選擇器
            for selector in finance_selectors:
                try:
                    finance_element = wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    try:
                        finance_element.click()
                        self.update_status("點擊會計經費查詢")
                        break
                    except:
                        try:
                            self.driver.execute_script("arguments[0].click();", finance_element)
                            self.update_status("點擊會計經費查詢 (JS)")
                            break
                        except:
                            continue
                except:
                    continue
            else:
                raise Exception("無法找到或點擊會計經費查詢")
            
            # 處理最後的分頁切換
            WebDriverWait(self.driver, 10).until(lambda d: len(d.window_handles) > 2)
            self.driver.switch_to.window(self.driver.window_handles[-1])
            
            self.update_status("成功進入會計經費查詢系統")
            self.load_year_options()
            return True
            
        except Exception as e:
            self.error_logger.log_error("導航過程發生錯誤", e)
            self.update_status("導航失敗，嘗試重新登入...", True)
            
            # 關閉瀏覽器
            if self.driver:
                self.driver.quit()
                self.driver = None
                
            # 重置登入狀態
            self.is_logged_in = False
            
            # 增加重試計數
            self.retry_count += 1
            
            # 重新登入並導航
            if self.retry_count < 2:  # 只有在未達到最大重試次數時才重試
                if self.login():
                    return self.navigate_to_query()
            else:
                self.update_status("重試次數已達上限，請手動重新啟動程式", True)
                # 重置重試計數
                self.retry_count = 0
                # 恢復介面狀態
                self.running_label.grid_remove()
                self.login_button.grid()
                self.restart_button.grid_remove()
                self.username.config(state='normal')
                self.password.config(state='normal')
                self.remember_checkbox.config(state='normal')
            return False

    def load_year_options(self):
        """從網頁載入可用學年"""
        try:
            # 等待頁面完全載入
            time.sleep(1)
            # 重新獲取學年選擇元素
            year_select_el = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, "swYear"))
            )
            options = year_select_el.find_elements(By.TAG_NAME, "option")
            
            # 過濾掉無效值
            year_list = [opt.get_attribute("value") for opt in options 
                        if opt.get_attribute("value").isdigit()]
            
            self.year_select['values'] = year_list
            if year_list:
                self.year_select.set(year_list[0])
                self.select_label.grid()  # 顯示選擇提示
            self.update_status(f"成功載入學年選單: {year_list}")
            self.show_select_year_and_report_button()
        except Exception as e:
            self.error_logger.log_error("載入學年選單失敗", e)
            self.update_status(f"載入學年選單失敗: {str(e)}", True)

    def show_select_year_and_report_button(self):
        """在載入學年後顯示選擇學年和報表的按鈕"""
        self.select_label.grid()  # 顯示提示標籤
        self.query_button.grid()  # 顯示查詢按鈕

    def select_year_and_report(self):
        try:
            # 禁用查詢按鈕和開啟報表按鈕
            self.query_button.config(state=tk.DISABLED)
            self.open_export_button.config(state=tk.DISABLED)
            
            selected_year = self.year_select.get()
            if not selected_year:
                self.update_status("請先選擇學年", True)
                # 重新啟用按鈕
                self.query_button.config(state=tk.NORMAL)
                self.open_export_button.config(state=tk.NORMAL)
                return False

            selected_plans = self.get_selected_plan_codes()
            if not selected_plans:
                self.update_status("請至少選擇一個計畫編號", True)
                # 重新啟用按鈕
                self.query_button.config(state=tk.NORMAL)
                self.open_export_button.config(state=tk.NORMAL)
                return False

            try:
                # 檢查是否在明細帳頁面並返回
                try:
                    # 尋找返回連結並點擊
                    back_link = self.driver.find_element(By.XPATH, "//a[contains(text(), '年度與報表選擇')]")
                    back_link.click()
                    self.update_status("返回年度選擇頁面")
                    
                    # 等待年度選擇頁面載入
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.NAME, "swYear"))
                    )
                except:
                    pass  # 如果找不到返回連結，表示已經在正確頁面
                
                # 導航到明細帳頁面
                self.navigate_to_project_input_page(selected_year)
                input_page_handle = self.driver.current_window_handle
                
                # 初始化Excel匯出器
                self.excel_exporter = ExcelExporter()
                
                # 依序處理每個計畫編號
                for plan_code in selected_plans:
                    try:
                        # 確保在輸入頁面
                        self.driver.switch_to.window(input_page_handle)
                        
                        # 輸入並送出計畫編號
                        self.input_and_submit_plan(plan_code)
                        
                        # 等待新的結果分頁開啟
                        time.sleep(1)
                        
                        # 找到新開啟的結果分頁
                        result_window = [h for h in self.driver.window_handles 
                                       if h != input_page_handle][-1]
                        
                        # 切換到結果分頁
                        self.driver.switch_to.window(result_window)
                        
                        # 等待結果頁面完全載入
                        WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.TAG_NAME, "table"))
                        )
                        
                        # 檢查是否為無搜尋結果頁面
                        if "沒查詢到任何結果" in self.driver.page_source:
                            self.update_status(f"計畫 {plan_code} 查無資料")
                        else:
                            # 在結果頁面獲取HTML內容並存入匯出器
                            html_content = self.driver.page_source
                            self.excel_exporter.add_data(plan_code, html_content)
                            self.update_status(f"計畫 {plan_code} 查詢完成")
                        
                        # 關閉結果分頁
                        time.sleep(1)
                        self.driver.close()

                    except Exception as e:
                        self.error_logger.log_error(f"處理計畫 {plan_code} 時發生錯誤", e)
                        self.update_status(f"處理計畫 {plan_code} 時發生錯誤: {str(e)}", True)
                        continue
                
                # 修改這部分：匯出Excel時使用執行檔所在路徑
                if getattr(sys, 'frozen', False):
                    # 如果是 exe 執行檔
                    application_path = os.path.dirname(sys.executable)
                else:
                    # 如果是一般 Python 腳本
                    application_path = os.path.dirname(os.path.abspath(__file__))
                    
                output_folder = 'Exports' # 輸出資料夾名稱
                try:
                    output_file = self.excel_exporter.export_excel(output_folder)
                    if output_file:
                        # 等待檔案系統完成寫入
                        time.sleep(1)
                        self.update_status(f"已匯出Excel檔案: {output_file}")
                except Exception as e:
                    self.error_logger.log_error("匯出Excel檔案時發生錯誤", e)
                    self.update_status(f"匯出Excel檔案時發生錯誤: {str(e)}", True)
                
                # 切回輸入頁面
                self.driver.switch_to.window(input_page_handle)
                self.update_status("爬蟲完成")
                # 成功完成後重新啟用按鈕
                self.query_button.config(state=tk.NORMAL)
                self.open_export_button.config(state=tk.NORMAL)
                return True
                
            except Exception as e:
                self.error_logger.log_error("查詢報表過程發生錯誤", e)
                self.update_status(f"查詢過程發生錯誤: {str(e)}", True)
                # 發生錯誤時重新啟用按鈕
                self.query_button.config(state=tk.NORMAL)
                self.open_export_button.config(state=tk.NORMAL)
                return False

        except Exception as e:
            self.error_logger.log_error("查詢報表過程發生錯誤", e)
            self.update_status(f"查詢過程發生錯誤: {str(e)}", True)
            # 發生錯誤時重新啟用按鈕
            self.query_button.config(state=tk.NORMAL)
            self.open_export_button.config(state=tk.NORMAL)
            return False

    def safe_click(self, locator, wait_time=10, retries=3):
        """
        安全地點擊元素，處理可能的 StaleElementReferenceException
        
        Args:
            locator: 元素定位器，格式為 (定位方法, 定位值) 如 (By.ID, "myId")
            wait_time: 等待元素出現的秒數
            retries: 重試次數
        
        Returns:
            bool: 操作是否成功
        """
        for attempt in range(retries):
            try:
                # 每次重新找元素，避免 stale element
                element = WebDriverWait(self.driver, wait_time).until(
                    EC.element_to_be_clickable(locator)
                )
                
                # 嘗試三種點擊方法
                try:
                    # 1. 常規點擊
                    element.click()
                    return True
                except:
                    try:
                        # 2. JavaScript 點擊
                        self.driver.execute_script("arguments[0].click();", element)
                        return True
                    except:
                        try:
                            # 3. ActionChains 點擊
                            action = ActionChains(self.driver)
                            action.move_to_element(element).click().perform()
                            return True
                        except Exception as e:
                            if attempt == retries - 1:  # 最後一次嘗試仍然失敗
                                self.error_logger.log_error(f"點擊元素失敗 ({str(locator)})", e)
                                return False
                            # 短暫等待後重試
                            time.sleep(1)
            except Exception as e:
                if attempt == retries - 1:  # 最後一次嘗試仍然失敗
                    self.error_logger.log_error(f"找不到元素 ({str(locator)})", e)
                    return False
                # 短暫等待後重試
                time.sleep(1)
        
        return False

    def safe_send_keys(self, locator, text, wait_time=10, retries=3):
        """
        安全地向元素輸入文字，處理可能的 StaleElementReferenceException
        
        Args:
            locator: 元素定位器，格式為 (定位方法, 定位值) 如 (By.ID, "myId")
            text: 要輸入的文字
            wait_time: 等待元素出現的秒數
            retries: 重試次數
        
        Returns:
            bool: 操作是否成功
        """
        for attempt in range(retries):
            try:
                # 每次重新找元素，避免 stale element
                element = WebDriverWait(self.driver, wait_time).until(
                    EC.visibility_of_element_located(locator)
                )
                
                # 先清空欄位再輸入
                try:
                    element.clear()
                    element.send_keys(text)
                    return True
                except Exception as e:
                    if attempt == retries - 1:  # 最後一次嘗試仍然失敗
                        self.error_logger.log_error(f"輸入文字至元素失敗 ({str(locator)})", e)
                        return False
                    # 短暫等待後重試
                    time.sleep(1)
            except Exception as e:
                if attempt == retries - 1:  # 最後一次嘗試仍然失敗
                    self.error_logger.log_error(f"找不到元素 ({str(locator)})", e)
                    return False
                # 短暫等待後重試
                time.sleep(1)
        
        return False

    def safe_get_text(self, locator, wait_time=10, retries=3, default=""):
        """
        安全地獲取元素文字，處理可能的 StaleElementReferenceException
        
        Args:
            locator: 元素定位器，格式為 (定位方法, 定位值) 如 (By.ID, "myId")
            wait_time: 等待元素出現的秒數
            retries: 重試次數
            default: 如果無法獲取文字時的預設值
        
        Returns:
            str: 元素文字或預設值
        """
        for attempt in range(retries):
            try:
                # 每次重新找元素，避免 stale element
                element = WebDriverWait(self.driver, wait_time).until(
                    EC.visibility_of_element_located(locator)
                )
                
                try:
                    return element.text
                except:
                    try:
                        # 嘗試使用 JavaScript 獲取文字
                        return self.driver.execute_script("return arguments[0].textContent;", element)
                    except Exception as e:
                        if attempt == retries - 1:  # 最後一次嘗試仍然失敗
                            self.error_logger.log_error(f"獲取元素文字失敗 ({str(locator)})", e)
                            return default
                        # 短暫等待後重試
                        time.sleep(1)
            except Exception as e:
                if attempt == retries - 1:  # 最後一次嘗試仍然失敗
                    self.error_logger.log_error(f"找不到元素 ({str(locator)})", e)
                    return default
                # 短暫等待後重試
                time.sleep(1)
        
        return default

    def navigate_to_project_input_page(self, selected_year):
        """導航到計畫編號輸入頁面"""
        try:
            # 選擇學年
            year_select_locator = (By.NAME, "swYear")
            if not self.safe_click(year_select_locator):
                raise Exception("無法點擊學年選擇下拉選單")
            
            # 選擇學年選項
            year_option_locator = (By.XPATH, f"//option[@value='{selected_year}']")
            if not self.safe_click(year_option_locator):
                raise Exception("無法選擇指定學年")
            
            self.update_status(f"選擇學年: {selected_year}")

            # 等待學年更新
            WebDriverWait(self.driver, 10).until(
                EC.text_to_be_present_in_element((By.ID, "lblYear"), selected_year)
            )

            # 點擊經費申請明細帳(科目)
            detail_btn_locator = (By.XPATH, "//td[contains(., '經費申請明細帳(科目)')]")
            if not self.safe_click(detail_btn_locator):
                raise Exception("無法點擊經費申請明細帳按鈕")
            
            # 等待新分頁開啟
            WebDriverWait(self.driver, 10).until(lambda d: len(d.window_handles) > 1)
            
            # 切換到新分頁
            new_window = self.driver.window_handles[-1]
            self.driver.switch_to.window(new_window)
            
            self.update_status("進入經費申請明細帳(科目)頁面")

            # 等待新頁面載入，確認有計畫編號輸入欄位
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "pjNoFrom"))
            )

        except Exception as e:
            self.error_logger.log_error(f"導航到計畫編號頁面時發生錯誤 (學年: {selected_year})", e)
            self.update_status(f"導航到計畫編號頁面時發生錯誤: {str(e)}", True)
            raise

    def input_and_submit_plan(self, plan_code):
        """在計畫編號頁面輸入並送出查詢"""
        try:
            # 使用改進的安全元素操作
            if not self.safe_send_keys((By.ID, "pjNoFrom"), plan_code):
                raise Exception("無法輸入計畫編號到第一個欄位")

            # 找到第二個計畫編號欄位並點擊，觸發自動填入
            if not self.safe_click((By.ID, "pjNoTo")):
                raise Exception("無法點擊第二個計畫編號欄位")

            # 確保第二個欄位有正確填入值
            for attempt in range(5):  # 重試5次
                try:
                    # 每次重新獲取元素
                    input_to = self.driver.find_element(By.ID, "pjNoTo")
                    # 檢查值是否已填入
                    if input_to.get_attribute('value') == plan_code:
                        break
                    # 等待短暫時間後再次檢查
                    time.sleep(0.5)
                except:
                    time.sleep(0.5)
                    continue
            else:
                # 如果無法自動填入，嘗試手動輸入
                if not self.safe_send_keys((By.ID, "pjNoTo"), plan_code):
                    raise Exception("無法確認第二個計畫編號欄位已填入正確值")
            
            self.update_status(f"輸入計畫編號: {plan_code}")

            # 點擊送出按鈕
            if not self.safe_click((By.NAME, "Submit")):
                # 嘗試使用其他選擇器
                if not self.safe_click((By.XPATH, "//input[@value='查詢']")):
                    if not self.safe_click((By.XPATH, "//input[@type='submit']")):
                        raise Exception("無法點擊送出按鈕")
                        
            self.update_status("送出查詢")
            
            # 等待查詢處理
            time.sleep(2)

        except Exception as e:
            self.error_logger.log_error(f"輸入計畫編號時發生錯誤 (計畫編號: {plan_code})", e)
            self.update_status(f"輸入計畫編號時發生錯誤: {str(e)}", True)
            raise

    def restart_program(self):
        """重啟程式功能"""
        try:
            # 關閉瀏覽器
            if self.driver:
                try:
                    self.driver.close()  # 先關閉視窗
                except:
                    pass
                try:
                    self.driver.quit()  # 再關閉驅動
                except:
                    pass
                self.driver = None
            
            # 重置登入狀態
            self.is_logged_in = False
            
            # 更新界面
            self.running_label.grid_remove()
            self.restart_button.grid_remove()
            self.login_button.grid()
            self.login_button.config(state=tk.NORMAL)  # 確保登入按鈕為啟用狀態
            self.query_button.grid_remove()  # 隱藏查詢按鈕
            
            for widget in self.button_frame.winfo_children():
                widget.destroy()
            
            # 啟用帳密輸入
            self.username.config(state='normal')
            self.password.config(state='normal')
            self.remember_checkbox.config(state='normal')
            
            # 隱藏選擇提示
            self.select_label.grid_remove()
            
            self.update_status("程式已重啟，請重新登入")
            
        except Exception as e:
            self.error_logger.log_error("重啟程式時發生錯誤", e)
            self.update_status(f"重啟程式時發生錯誤: {str(e)}", True)

    def open_export_folder(self):
        """開啟報表輸出資料夾"""
        try:
            # 使用執行檔所在目錄
            if getattr(sys, 'frozen', False):
                # 如果是 exe 執行檔
                application_path = os.path.dirname(sys.executable)
            else:
                # 如果是一般 Python 腳本
                application_path = os.path.dirname(os.path.abspath(__file__))
                
            output_folder = os.path.join(application_path, 'exports')
            os.makedirs(output_folder, exist_ok=True)
            
            if os.name == 'nt':  # Windows
                os.startfile(output_folder)
            elif os.name == 'posix':  # macOS 和 Linux
                import subprocess
                subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', output_folder])
                
            self.update_status("開啟報表輸出資料夾")
        except Exception as e:
            self.error_logger.log_error("開啟報表資料夾失敗", e)
            self.update_status(f"開啟報表資料夾失敗: {str(e)}", True)

    def __del__(self):
        """清理瀏覽器資源"""
        if hasattr(self, 'driver') and self.driver:
            try:
                self.driver.close()
            except:
                pass
            try:
                self.driver.quit()
            except:
                pass

def get_resource_path(relative_path):
    """獲取資源檔案的絕對路徑"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

if __name__ == '__main__':
    root = tk.Tk()
    app = ItouchCrawler(root)
    def on_closing():
        # 關閉程式前進行清理
        if hasattr(app, 'driver') and app.driver:
            try:
                app.driver.close()
            except:
                pass
            try:
                app.driver.quit()
            except:
                pass
        
        # 強制進行垃圾回收
        import gc
        gc.collect()
        
        root.destroy()
    
    # 綁定關閉視窗事件
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    root.mainloop()
