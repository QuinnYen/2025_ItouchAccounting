from bs4 import BeautifulSoup
import pandas as pd
import os, sys
from datetime import datetime

class ExcelExporter:
    def __init__(self):
        self.projects_data = []
        
    def extract_project_info(self, soup):
        """提取計畫基本資訊"""
        table1 = soup.find('table', {'id': 'table1'})
        rows = table1.find_all('tr')
        
        # 修改獲取學年度的方式
        header_text = rows[0].get_text(strip=True)
        import re
        year_match = re.search(r'(\d+)\s*學年度', header_text)
        academic_year = f"{year_match.group(1)}學年度" if year_match else "未知學年度"
        
        # 提取學年度和部門資訊（包含學校名稱）
        dept_info = rows[1].find_all('td')[0].get_text(strip=True)
        department = dept_info.split('：')[1].split()[0]  # 獲取部門代碼
        
        # 提取計畫編號和名稱
        project_info = rows[1].get_text(strip=True)
        project_code = project_info.split('計畫編號：')[1].split('計畫名稱：')[0].strip()
        project_name = project_info.split('計畫名稱：')[1].strip()
        
        # 提取預算金額
        budget = rows[2].find_all('td')[1].get_text(strip=True)
        
        # 提取可用餘額
        last_table = soup.find_all('table', {'id': 'table1'})[-1]
        available = last_table.find_all('td')[1].get_text(strip=True)
        
        return {
            '學年度': academic_year,
            '計畫編號': project_code,
            '計畫名稱': project_name,
            '目前預算': budget,
            '可用餘額': available
        }
        
    def extract_subtotals(self, soup):
        table2 = soup.find('table', {'id': 'table2'})
        rows = table2.find_all('tr')
        subtotals = {}
        for row in rows:
            cells = row.find_all('td')
            # 如果其中一個儲存格含有"小計"文字，則從該儲存格取完整文字
            if any('小計' in c.get_text() for c in cells) and '預算收支' not in row.get_text() and '非預算收支' not in row.get_text():
                # 假設科目與"小計"在同一個儲存格
                for c in cells:
                    if '小計' in c.get_text():
                        subject_text = c.get_text(strip=True)
                        # 保留完整的小計文字
                        subject_code = subject_text.split('&')[0].strip()  # 移除可能的 &nbsp;
                        
                # 金額通常在下一個含有 <strong> 的儲存格
                for c in cells:
                    strong_el = c.find('strong')
                    if strong_el and strong_el.get_text(strip=True).replace(',', '').isdigit():
                        amount = strong_el.get_text(strip=True).replace(',', '')
                        subtotals[subject_code] = amount
                        break
        return subtotals

    def add_data(self, plan_code, html_content):
        """處理 HTML 內容並提取所需資料"""
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 提取基本資訊
        project_info = self.extract_project_info(soup)
        
        # 提取科目小計
        subtotals = self.extract_subtotals(soup)
        
        # 儲存所有資料
        project_data = {
            'info': project_info,
            'subtotals': subtotals
        }
        
        # print(f"已處理計畫 {plan_code}:")
        # print(f"基本資訊: {project_info}")
        # print(f"科目小計: {subtotals}")
        
        self.projects_data.append(project_data)

    def export_excel(self, output_folder):
        """匯出資料到 Excel 檔案"""
        if not self.projects_data:
            return None
            
        # 判斷是否為 exe 執行環境
        if getattr(sys, 'frozen', False):
            # 如果是 exe 執行檔
            base_path = os.path.dirname(sys.executable)
        else:
            # 如果是一般 Python 腳本
            base_path = os.path.dirname(os.path.abspath(__file__))
            
        # 使用基礎路徑建立完整的輸出路徑
        output_path = os.path.join(base_path, output_folder)
        os.makedirs(output_path, exist_ok=True)
        
        # 生成檔案名稱
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = os.path.join(output_path, f'計畫經費報表_{timestamp}.xlsx')
        
        try:
            # 建立 Excel 寫入器
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                worksheet = workbook.add_worksheet('經費報表')
                
                # 設定格式
                header_format = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1,
                    'bg_color': '#FFA500',  # 橘色背景
                    'font_color': 'white',   # 白色文字
                    'font_size': 11,         # 字體大小
                    'text_wrap': True        # 自動換行
                })
                
                data_format = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1,
                    'font_size': 10,
                    'text_wrap': True
                })
                
                # 金額格式（包含千分位）
                money_format = workbook.add_format({
                    'align': 'right',
                    'valign': 'vcenter',
                    'border': 1,
                    'font_size': 10,
                    'num_format': '#,##0'
                })
                
                # 間隔列的格式（淺灰色背景）
                alt_row_format = workbook.add_format({
                    'align': 'center',
                    'valign': 'vcenter',
                    'border': 1,
                    'font_size': 10,
                    'text_wrap': True,
                    'bg_color': '#F5F5F5'  # 淺灰色背景
                })
                
                alt_row_money_format = workbook.add_format({
                    'align': 'right',
                    'valign': 'vcenter',
                    'border': 1,
                    'font_size': 10,
                    'num_format': '#,##0',
                    'bg_color': '#F5F5F5'  # 淺灰色背景
                })
                
                # 設定起始行
                current_row = 0
                
                # 處理每個計畫
                for project in self.projects_data:
                    # 準備基本欄位
                    base_headers = ['學年度', '計畫編號', '計畫名稱', '目前預算', '可用餘額']
                    
                    # 獲取此計畫的科目代碼
                    subject_codes = sorted(project['subtotals'].keys())
                    
                    # 合併所有標題
                    headers = base_headers + subject_codes
                    
                    # 設定欄寬
                    worksheet.set_column('A:A', 12)  # 學年度
                    worksheet.set_column('B:B', 15)  # 計畫編號
                    worksheet.set_column('C:C', 40)  # 計畫名稱
                    worksheet.set_column('D:E', 15)  # 預算和餘額
                    worksheet.set_column(5, len(headers), 15)  # 其他科目欄位
                    
                    # 寫入標題
                    for col, header in enumerate(headers):
                        worksheet.write(current_row, col, header, header_format)
                    
                    # 準備資料行
                    row_data = [
                        project['info']['學年度'],
                        project['info']['計畫編號'],
                        project['info']['計畫名稱'],
                        int(project['info']['目前預算'].replace(',', '')),
                        int(project['info']['可用餘額'].replace(',', ''))
                    ]
                    
                    # 添加各科目小計
                    for code in subject_codes:
                        row_data.append(int(project['subtotals'][code]))
                    
                    # 寫入資料（使用間隔列格式）
                    for col, value in enumerate(row_data):
                        # 判斷是否為金額欄位（從第4列開始的都是金額）
                        if col >= 3:
                            format_to_use = money_format if (current_row + 1) % 6 != 0 else alt_row_money_format
                        else:
                            format_to_use = data_format if (current_row + 1) % 6 != 0 else alt_row_format
                        
                        worksheet.write(current_row + 1, col, value, format_to_use)
                    
                    # 設定第一欄的行高
                    worksheet.set_row(current_row, 30)     # 標題列高度
                    worksheet.set_row(current_row + 1, 25) # 資料列高度
                    
                    # 新增空白行
                    current_row += 3
                
            return output_file
            
        except Exception as e:
            # 如果發生錯誤，確保資源被釋放
            if os.path.exists(output_file):
                try:
                    # 嘗試刪除可能未完成的文件
                    os.unlink(output_file)
                except:
                    pass
            raise e