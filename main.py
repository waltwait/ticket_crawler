import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException
import time
import pandas as pd
import re
from webdriver_manager.chrome import ChromeDriverManager
import datetime

class KKdayFlightScraper:
    def __init__(self, url):
        self.url = url
        self.options = Options()
        # 可選: 無頭模式，不顯示瀏覽器
        # self.options.add_argument('--headless')
        self.options.add_argument('--no-sandbox')
        self.options.add_argument('--disable-dev-shm-usage')
        self.options.add_argument('--window-size=1920,1080')  # 設置較大的窗口尺寸
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=self.options)
        self.all_product_data = []  # 存儲所有產品的數據
        
    def open_page(self):
        """打開網頁"""
        self.driver.get(self.url)
        print("頁面已打開")
        time.sleep(5)  # 等待頁面完全加載
        
    def get_product_options(self):
        """獲取所有產品選項"""
        try:
            # 等待產品選項加載
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.option-head"))
            )
            
            # 獲取所有產品選項
            product_options = self.driver.find_elements(By.CSS_SELECTOR, "div.option-head")
            print(f"找到 {len(product_options)} 個產品選項")
            return product_options
        except TimeoutException:
            print("找不到產品選項")
            return []
            
    def get_product_info(self, product_element):
        """提取產品基本信息"""
        try:
            # 提取產品標題
            title_element = product_element.find_element(By.CSS_SELECTOR, "span.kk-u-text-h6")
            title = title_element.text if title_element else "未知產品"
            
            # 提取價格
            price_element = product_element.find_element(By.CSS_SELECTOR, "div.kk-price-local__normal")
            base_price = price_element.text.strip() if price_element else "未知價格"
            
            return {
                "title": title,
                "base_price": base_price
            }
        except (NoSuchElementException, StaleElementReferenceException) as e:
            print(f"提取產品信息時發生錯誤: {e}")
            return {"title": "未知產品", "base_price": "未知價格"}
            
    def click_select_button(self, product_element):
        """點擊特定產品的選擇按鈕"""
        try:
            # 找到該產品的選擇按鈕
            select_button = product_element.find_element(By.CSS_SELECTOR, "button.kk-button.select-option")
            
            # 滾動到按鈕位置
            self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_button)
            time.sleep(1)  # 等待滾動完成
            
            # 點擊按鈕
            select_button.click()
            print("已點擊'選擇'按鈕")
            time.sleep(2)  # 等待彈出窗口加載
            return True
        except (NoSuchElementException, StaleElementReferenceException) as e:
            print(f"找不到'選擇'按鈕或點擊失敗: {e}")
            return False
            
    def extract_available_dates_and_prices(self):
        """提取可用日期和價格"""
        dates_prices = []
        
        try:
            # 等待日曆表格加載
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.date-table"))
            )
            
            # 解析頁面內容
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            
            # 找到所有可選擇的日期單元格（有價格的才是可選的）
            date_cells = soup.select("td.cell-date.selectable")
            
            for cell in date_cells:
                date_num = cell.select_one("div.date-num").text.strip()
                price_elem = cell.select_one("div.price")
                price = price_elem.text.strip() if price_elem else "無價格"
                
                # 獲取當前月份和年份
                current_month_elem = soup.select_one("div.current-month")
                current_month = current_month_elem.text.strip() if current_month_elem else "未知月份"
                
                dates_prices.append({
                    'date': f"{current_month} {date_num}日",
                    'price': price
                })
                
            return dates_prices
        except Exception as e:
            print(f"提取日期和價格時發生錯誤: {e}")
            return []
    
    def navigate_through_months(self, product_info, months_ahead=3):
        """瀏覽接下來幾個月的數據"""
        all_dates_prices = []
        
        # 獲取當前月份的數據
        current_month_data = self.extract_available_dates_and_prices()
        for item in current_month_data:
            item.update(product_info)  # 將產品信息添加到日期價格數據中
        all_dates_prices.extend(current_month_data)
        
        # 遍歷接下來的幾個月
        for _ in range(months_ahead):
            try:
                # 檢查下個月按鈕是否可用
                next_month_buttons = self.driver.find_elements(By.CSS_SELECTOR, "div.change-month.next-month")
                if not next_month_buttons or "disabled" in next_month_buttons[0].get_attribute("class"):
                    print("沒有更多月份可瀏覽")
                    break
                
                # 點擊下個月按鈕
                next_month_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "div.change-month.next-month"))
                )
                next_month_button.click()
                print("已點擊下個月按鈕")
                time.sleep(1.5)  # 等待日曆更新
                
                # 獲取新月份的數據
                month_data = self.extract_available_dates_and_prices()
                for item in month_data:
                    item.update(product_info)  # 將產品信息添加到日期價格數據中
                all_dates_prices.extend(month_data)
                
            except Exception as e:
                print(f"瀏覽下個月時發生錯誤: {e}")
                break
                
        return all_dates_prices
    
    def close_booking_modal(self):
        """關閉預訂彈窗，返回到產品列表頁"""
        try:
            # 嘗試找到關閉按鈕或返回按鈕
            close_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button.modal-close, button.kk-button--ghost:not(.select-option)")
            if close_buttons:
                close_buttons[0].click()
                print("已關閉預訂彈窗")
                time.sleep(1)
                return True
            
            # 如果找不到關閉按鈕，嘗試點擊頁面空白處
            self.driver.execute_script("document.body.click();")
            time.sleep(1)
            return True
        except Exception as e:
            print(f"關閉彈窗時發生錯誤: {e}")
            # 如果關閉失敗，嘗試刷新頁面
            self.driver.refresh()
            time.sleep(3)
            return False
    
    def save_to_excel(self, data, filename=None):
        """將數據保存到Excel文件"""
        if not filename:
            # 使用當前日期時間生成文件名
            current_datetime = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"kkday_flight_prices_{current_datetime}.xlsx"
        
        if data:
            df = pd.DataFrame(data)
            # 重新排列列的順序，使產品信息在前面
            if 'title' in df.columns and 'base_price' in df.columns:
                cols = ['title', 'base_price', 'date', 'price']
                df = df[cols + [c for c in df.columns if c not in cols]]
            
            # 確保價格列是數字格式
            try:
                df['price'] = df['price'].replace('', '0')  # 替換空字符串為0
                df['price'] = pd.to_numeric(df['price'].str.replace(',', ''), errors='coerce')
                df['base_price'] = pd.to_numeric(df['base_price'].str.replace(',', ''), errors='coerce')
            except Exception as e:
                print(f"轉換價格格式時發生錯誤: {e}")
            
            # 創建Excel writer對象
            writer = pd.ExcelWriter(filename, engine='xlsxwriter')
            
            # 將DataFrame寫入Excel
            df.to_excel(writer, sheet_name='航班價格', index=False)
            
            # 獲取workbook和worksheet對象以進行格式化
            workbook = writer.book
            worksheet = writer.sheets['航班價格']
            
            # 定義格式
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'bg_color': '#D7E4BC',
                'border': 1
            })
            
            price_format = workbook.add_format({'num_format': '#,##0'})
            date_format = workbook.add_format({'num_format': 'yyyy年mm月dd日'})
            
            # 設置列寬
            worksheet.set_column('A:A', 40)  # 產品標題列
            worksheet.set_column('B:B', 12, price_format)  # 基本價格列
            worksheet.set_column('C:C', 20)  # 日期列
            worksheet.set_column('D:D', 12, price_format)  # 價格列
            
            # 添加表頭格式
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # 添加自動篩選
            worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
            
            # 添加凍結窗格
            worksheet.freeze_panes(1, 0)
            
            # 創建一個圖表
            chart = workbook.add_chart({'type': 'line'})
            
            # 設置圖表數據範圍
            last_row = len(df)
            chart.add_series({
                'name': '價格趨勢',
                'categories': ['航班價格', 1, 2, last_row, 2],  # 日期列
                'values': ['航班價格', 1, 3, last_row, 3],      # 價格列
                'line': {'color': 'red'},
            })
            
            # 設置圖表標題和軸標籤
            chart.set_title({'name': '航班價格趨勢'})
            chart.set_x_axis({'name': '日期'})
            chart.set_y_axis({'name': '價格 (NT$)'})
            
            # 將圖表插入到新的工作表
            chart_sheet = workbook.add_worksheet('價格趨勢圖')
            chart_sheet.insert_chart('B2', chart, {'x_scale': 2, 'y_scale': 1.5})
            
            # 保存Excel文件
            writer.close()
            
            print(f"數據已保存到 {filename}")
        else:
            print("沒有數據可保存")
    
    def run(self, months_to_scrape=3):
        """執行完整的爬蟲流程"""
        try:
            self.open_page()
            
            # 獲取所有產品選項
            product_options = self.get_product_options()
            
            if not product_options:
                print("未找到產品選項，爬蟲終止")
                return
                
            for i, product_element in enumerate(product_options):
                try:
                    print(f"\n正在處理第 {i+1}/{len(product_options)} 個產品選項")
                    
                    # 提取產品基本信息
                    product_info = self.get_product_info(product_element)
                    print(f"產品標題: {product_info['title']}")
                    print(f"基本價格: {product_info['base_price']}")
                    
                    # 點擊選擇按鈕
                    if self.click_select_button(product_element):
                        # 獲取該產品的日期和價格
                        dates_prices = self.navigate_through_months(product_info, months_to_scrape)
                        
                        if dates_prices:
                            print(f"共找到 {len(dates_prices)} 個可用日期")
                            # 存儲該產品的數據
                            self.all_product_data.extend(dates_prices)
                        else:
                            print("未找到該產品的可用日期")
                        
                        # 關閉預訂彈窗，返回到產品列表頁
                        self.close_booking_modal()
                    
                    # 重新獲取產品元素列表，避免StaleElementReferenceException
                    if i < len(product_options) - 1:  # 如果不是最後一個產品
                        product_options = self.get_product_options()
                
                except Exception as e:
                    print(f"處理產品 {i+1} 時發生錯誤: {e}")
                    # 嘗試關閉彈窗並繼續下一個產品
                    self.close_booking_modal()
                    # 重新獲取產品元素列表
                    if i < len(product_options) - 1:
                        product_options = self.get_product_options()
            
                            # 顯示並保存所有產品的數據
            if self.all_product_data:
                print(f"\n所有產品共找到 {len(self.all_product_data)} 個可用日期")
                self.save_to_excel(self.all_product_data)
            else:
                print("未找到任何可用日期")
                
        except Exception as e:
            print(f"爬蟲過程中發生錯誤: {e}")
        finally:
            # 關閉瀏覽器
            self.driver.quit()
            print("瀏覽器已關閉")


if __name__ == "__main__":
    # 替換為你要爬取的KKday產品頁面URL
    url = "https://www.kkday.com/zh-tw/product/137240"  # 將PRODUCT_ID替換為實際的產品ID
    
    scraper = KKdayFlightScraper(url)
    
    # 可自定義參數
    months_to_scrape = 3  # 每個產品向後查詢的月數
    
    # 執行爬蟲
    print(f"開始爬取URL: {url}")
    print(f"將查詢每個產品的 {months_to_scrape} 個月數據")
    scraper.run(months_to_scrape)