import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import pandas as pd
import random
from webdriver_manager.chrome import ChromeDriverManager
from fake_useragent import UserAgent
import undetected_chromedriver as uc
from selenium.webdriver.common.action_chains import ActionChains
import logging
import concurrent.futures

class KKdayFlightScraper:
    def __init__(self, url):
        self.url = url
        self.setup_logging()
        self.setup_browser()
        self.all_product_data = []

    def setup_logging(self):
        """設置日誌系統"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('scraper.log'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)

    def setup_browser(self):
        """設置瀏覽器配置"""
        try:
            # 使用 undetected_chromedriver 來避免檢測
            options = uc.ChromeOptions()
            
            # 添加瀏覽器指紋隨機化
            self.add_browser_fingerprint_randomization(options)
            
            # 使用 undetected_chromedriver
            self.driver = uc.Chrome(options=options)
            
            # 設置視窗大小隨機化
            self.randomize_window_size()
            
            # 添加 CDP 指令來修改 webdriver 標記
            self.modify_navigator_webdriver()
            
        except Exception as e:
            self.logger.error(f"設置瀏覽器時發生錯誤: {e}")
            raise

    def add_browser_fingerprint_randomization(self, options):
        # 使用 fake-useragent 生成隨機 User-Agent
        ua = UserAgent()
        options.add_argument(f'user-agent={ua.random}')
        
        # 添加隨機語言設置
        languages = ['en-US,en;q=0.9', 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7']
        options.add_argument(f'--lang={random.choice(languages)}')
        
        # 隨機設置 viewport 大小
        options.add_argument('--start-maximized')
        
        # 禁用自動化控制特徵
        options.add_argument('--disable-blink-features=AutomationControlled')
        
        # 添加其他隨機參數
        if random.random() > 0.5:
            options.add_argument('--disable-gpu')
        if random.random() > 0.5:
            options.add_argument('--disable-extensions')

    def modify_navigator_webdriver(self):
        """修改 navigator.webdriver 標記"""
        self.driver.execute_cdp_cmd('Page.addScriptToEvaluateOnNewDocument', {
            'source': '''
                Object.defineProperty(navigator, 'webdriver', {
                    get: () => undefined
                });
                Object.defineProperty(navigator, 'language', {
                    get: () => ['en-US', 'zh-TW'][Math.floor(Math.random() * 2)]
                });
                Object.defineProperty(navigator, 'platform', {
                    get: () => ['Win32', 'MacIntel'][Math.floor(Math.random() * 2)]
                });
            '''
        })

    def randomize_window_size(self):
        """隨機化視窗大小"""
        widths = [1366, 1440, 1536, 1920]
        heights = [768, 900, 864, 1080]
        width = random.choice(widths)
        height = random.choice(heights)
        self.driver.set_window_size(width, height)

    def simulate_human_behavior(self):
        """模擬人類行為"""
        # 隨機滾動
        for _ in range(random.randint(2, 5)):
            scroll_amount = random.randint(100, 700)
            self.driver.execute_script(f"window.scrollBy(0, {scroll_amount});")
            time.sleep(random.uniform(0.5, 2))
            
            # 偶爾向上滾動
            if random.random() > 0.7:
                self.driver.execute_script(f"window.scrollBy(0, -{scroll_amount//2});")
                time.sleep(random.uniform(0.3, 1))
        
        # 隨機鼠標移動
        if random.random() > 0.5:
            elements = self.driver.find_elements(By.TAG_NAME, "div")
            if elements:
                element = random.choice(elements)
                try:
                    ActionChains(self.driver).move_to_element(element).perform()
                    time.sleep(random.uniform(0.3, 1))
                except:
                    pass

    def add_random_delay(self, min_delay=1, max_delay=3):
        """添加智能隨機延遲"""
        base_delay = random.uniform(min_delay, max_delay)
        # 添加微小的隨機變化
        noise = random.gauss(0, 0.1)
        delay = max(0.1, base_delay + noise)
        time.sleep(delay)

    def open_page(self):
        """打開頁面並模擬真實用戶行為"""
        try:
            # 添加頁面加載超時處理
            self.driver.set_page_load_timeout(30)
            self.driver.get(self.url)
            
            # 模擬真實用戶行為
            self.simulate_human_behavior()
            
            # 隨機等待
            self.add_random_delay(3, 7)
            
            self.logger.info("頁面已成功打開")
        except Exception as e:
            self.logger.error(f"打開頁面時發生錯誤: {e}")
            raise

    def click_element(self, element):
        """智能點擊元素"""
        try:
            # 先確保元素在視圖中
            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
            self.add_random_delay(0.5, 1.5)
            
            # 隨機選擇點擊方式
            if random.random() > 0.5:
                # 使用 ActionChains
                ActionChains(self.driver).move_to_element(element).pause(random.uniform(0.1, 0.3)).click().perform()
            else:
                # 使用 JavaScript
                self.driver.execute_script("arguments[0].click();", element)
            
            self.add_random_delay(0.5, 1.5)
            
        except Exception as e:
            self.logger.error(f"點擊元素時發生錯誤: {e}")
            raise

    def get_product_info(self):
        """獲取產品基本信息"""
        try:
            # 只獲取產品標題
            try:
                title_element = self.driver.find_element(By.CSS_SELECTOR, "span.kk-u-text-h6")
                title = title_element.text
            except:
                try:
                    title_element = self.driver.find_element(By.CSS_SELECTOR, "h1.product-name")
                    title = title_element.text
                except:
                    title_element = self.driver.find_element(By.CSS_SELECTOR, "div.kk-product-name")
                    title = title_element.text
            
            return {
                "title": title if title else "未知產品",
                "base_price": ""  # 返回空字符串作為價格
            }
        except Exception as e:
            print(f"提取產品信息時發生錯誤: {e}")
            return {"title": "未知產品", "base_price": ""}

    def check_page_type(self):
        """檢查頁面類型，判斷是否直接顯示日曆"""
        try:
            # 檢查是否有選擇按鈕
            select_buttons = self.driver.find_elements(By.CSS_SELECTOR, "button.kk-button.select-option")
            if select_buttons:
                return "has_select_button"
            
            # 檢查是否直接顯示日曆
            calendar = self.driver.find_elements(By.CSS_SELECTOR, "div.option-booking")
            if calendar:
                return "direct_calendar"
            
            return "unknown"
        except Exception as e:
            print(f"檢查頁面類型時發生錯誤: {e}")
            return "unknown"

    def process_with_select_button(self):
        """處理有選擇按鈕的頁面"""
        try:
            # 等待產品選項加載
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.option-head"))
            )
            
            # 模擬人類行為
            self.simulate_human_behavior()
            self.add_random_delay(2, 4)
            
            # 先獲取所有產品選項
            product_options = self.driver.find_elements(By.CSS_SELECTOR, "div.option-head")
            if not product_options:
                print("未找到產品選項")
                return
            
            total_options = len(product_options)
            print(f"找到 {total_options} 個產品選項")
            
            for i in range(total_options):
                try:
                     # 添加人性化延遲
                    self.add_random_delay(2, 4)
                    self.logger.info(f"\n正在處理第 {i+1}/{total_options} 個產品選項")
                    
                    # 重新獲取最新的產品元素列表
                    product_options = self.driver.find_elements(By.CSS_SELECTOR, "div.option-head")
                    if i >= len(product_options):
                        print(f"找不到第 {i+1} 個產品選項")
                        continue
                    
                    current_option = product_options[i]
                    
                    # 模擬真實滾動行為
                    self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", current_option)
                    self.add_random_delay(1, 2)
                    self.simulate_human_behavior()
                
                    
                    # 獲取產品信息
                    try:
                        title_element = current_option.find_element(By.CSS_SELECTOR, "span.kk-u-text-h6")
                        title = title_element.text
                        product_info = {"title": title if title else "未知產品", "base_price": ""}
                        print(f"產品標題: {product_info['title']}")
                    except Exception as e:
                        print(f"提取產品標題時發生錯誤: {e}")
                        continue
                    
                    # 尋找並點擊選擇按鈕
                    try:
                        # 使用 WebDriverWait 等待按鈕可點擊
                        select_button = WebDriverWait(current_option, 10).until(
                            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.kk-button.select-option"))
                        )
                        
                        # 確保按鈕在視圖中且可點擊
                        self.driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", select_button)
                        time.sleep(1)
                        
                        # 使用 JavaScript 點擊按鈕
                        self.driver.execute_script("arguments[0].click();", select_button)
                        print("已點擊'選擇'按鈕")
                        time.sleep(2)
                        
                        # 獲取日期和價格
                        dates_prices = self.navigate_through_months(product_info)
                        if dates_prices:
                            self.all_product_data.extend(dates_prices)
                        
                        # 關閉彈窗
                        self.close_booking_modal()
                        time.sleep(1)  # 等待彈窗完全關閉
                        
                    except Exception as e:
                        print(f"點擊選擇按鈕時發生錯誤: {e}")
                        self.close_booking_modal()
                        time.sleep(1)
                        continue
                        
                except Exception as e:
                    print(f"處理產品選項時發生錯誤: {e}")
                    continue
                    
        except Exception as e:
            print(f"處理選擇按鈕頁面時發生錯誤: {e}")

    def process_direct_calendar(self):
        """處理直接顯示日曆的頁面"""
        try:
            # 獲取產品信息
            product_info = self.get_product_info()
            print(f"產品標題: {product_info['title']}")
            print(f"基本價格: {product_info['base_price']}")
            
            # 直接獲取日期和價格
            dates_prices = self.navigate_through_months(product_info)
            if dates_prices:
                self.all_product_data.extend(dates_prices)
                
        except Exception as e:
            print(f"處理直接日曆頁面時發生錯誤: {e}")

    def extract_available_dates_and_prices(self):
        """提取可用日期和價格"""
        dates_prices = []
        try:
            # 等待日曆表格加載
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "table.date-table"))
            )
            
            # 只添加最小延遲，不進行滾動
            self.add_random_delay(0.5, 1)
            
            # 解析頁面內容
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            
            # 找到所有可選擇的日期單元格
            date_cells = soup.select("td.cell-date.selectable")
            
            # 獲取當前月份
            current_month_elem = soup.select_one("div.current-month")
            current_month = current_month_elem.text.strip() if current_month_elem else "未知月份"
            
            for cell in date_cells:
                dates_prices.append({
                    'date': f"{current_month} {cell.select_one('div.date-num').text.strip()}日",
                    'price': cell.select_one('div.price').text.strip() if cell.select_one('div.price') else "無價格"
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
            item.update(product_info)
        all_dates_prices.extend(current_month_data)
        
        # 遍歷接下來的幾個月
        for _ in range(months_ahead):
            try:
                # 只添加最小延遲
                self.add_random_delay(1, 1.5)
                
                # 檢查下個月按鈕
                next_month_buttons = self.driver.find_elements(By.CSS_SELECTOR, "div.change-month.next-month")
                if not next_month_buttons or "disabled" in next_month_buttons[0].get_attribute("class"):
                    print("沒有更多月份可瀏覽")
                    break
                
                # 簡單點擊下個月按鈕
                next_month_button = next_month_buttons[0]
                self.driver.execute_script("arguments[0].click();", next_month_button)
                self.add_random_delay(1, 1.5)
                
                # 獲取新月份的數據
                month_data = self.extract_available_dates_and_prices()
                for item in month_data:
                    item.update(product_info)
                all_dates_prices.extend(month_data)
                
            except Exception as e:
                print(f"瀏覽下個月時發生錯誤: {e}")
                break
                
        return all_dates_prices

    def close_booking_modal(self):
        """改進的關閉預訂彈窗方法"""
        try:
            # 先等待一小段時間讓頁面穩定
            self.add_random_delay(1, 2)
            
            # 嘗試多個可能的選擇器
            modal_selectors = [
                "div.modal-dialog",
                "div.modal-content",
                "div.modal",
                "div.booking-modal",
                "div.modal-wrapper"
            ]
            
            modal_found = False
            for selector in modal_selectors:
                try:
                    # 使用較短的超時時間來快速檢查每個選擇器
                    modal = WebDriverWait(self.driver, 3).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    modal_found = True
                    self.logger.info(f"找到彈窗元素: {selector}")
                    break
                except:
                    continue
            
            if not modal_found:
                self.logger.warning("未找到彈窗元素，嘗試其他關閉方法")
                # 嘗試按 ESC 鍵關閉
                try:
                    ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                    self.add_random_delay(1, 2)
                except:
                    self.logger.warning("ESC鍵關閉失敗")
                return
            
            # 嘗試找到並點擊關閉按鈕
            close_button_selectors = [
                "button.modal-close",
                "button.kk-button--ghost:not(.select-option)",
                "button.close",
                "i.close-icon",
                "div.modal-close"
            ]
            
            for selector in close_button_selectors:
                try:
                    close_buttons = self.driver.find_elements(By.CSS_SELECTOR, selector)
                    if close_buttons:
                        # 使用JavaScript點擊第一個可見的關閉按鈕
                        for button in close_buttons:
                            if button.is_displayed():
                                self.driver.execute_script("arguments[0].click();", button)
                                self.logger.info(f"成功點擊關閉按鈕: {selector}")
                                self.add_random_delay(1, 2)
                                return
                except Exception as e:
                    self.logger.warning(f"嘗試點擊關閉按鈕失敗: {selector}, {str(e)}")
                    continue
            
            # 如果找不到關閉按鈕，嘗試點擊彈窗外部
            try:
                self.driver.execute_script("""
                    var elements = document.getElementsByClassName('modal-dialog');
                    if(elements.length > 0) {
                        elements[0].parentElement.click();
                    }
                """)
                self.logger.info("已嘗試點擊彈窗外部關閉")
                self.add_random_delay(1, 2)
            except Exception as e:
                self.logger.warning(f"點擊彈窗外部失敗: {str(e)}")
            
        except Exception as e:
            self.logger.error(f"關閉彈窗時發生錯誤: {str(e)}")
            # 如果所有方法都失敗，嘗試刷新頁面
            try:
                self.driver.refresh()
                self.add_random_delay(2, 4)
                self.logger.info("已刷新頁面")
            except:
                self.logger.error("刷新頁面失敗")
            
    def save_to_excel(self, data, filename):
        """保存數據到Excel"""
        if not data:
            print("沒有數據可保存")
            return
            
        try:
            df = pd.DataFrame(data)
            if 'title' in df.columns:
                cols = ['title', 'date', 'price']  # 移除 base_price
                df = df[cols + [c for c in df.columns if c not in cols]]
            
            # 只處理 price 列
            df['price'] = df['price'].replace('', '0')
            df['price'] = pd.to_numeric(df['price'].str.replace(',', ''), errors='coerce')
            
            with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='航班價格', index=False)
                
                workbook = writer.book
                worksheet = writer.sheets['航班價格']
                
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'bg_color': '#D7E4BC',
                    'border': 1
                })
                
                price_format = workbook.add_format({'num_format': '#,##0'})
                
                worksheet.set_column('A:A', 40)  # 標題列
                worksheet.set_column('B:B', 20)  # 日期列
                worksheet.set_column('C:C', 12, price_format)  # 價格列
                
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)
                worksheet.freeze_panes(1, 0)
                
            print(f"數據已保存到 {filename}")
        except Exception as e:
            print(f"保存Excel時發生錯誤: {e}")

    def run(self, months_to_scrape=3):
        """執行爬蟲"""
        try:
            self.open_page()
            
            # 檢查頁面類型
            page_type = self.check_page_type()
            print(f"檢測到頁面類型: {page_type}")
            
            # 根據頁面類型選擇處理方式
            if page_type == "has_select_button":
                self.process_with_select_button()
            elif page_type == "direct_calendar":
                self.process_direct_calendar()
            else:
                print("無法識別的頁面類型")
            
            # 保存數據
            if self.all_product_data:
                url_id = self.url.rstrip('/').split('/')[-1]
                filename = f"kkday_{url_id}.xlsx"
                self.save_to_excel(self.all_product_data, filename)
                print(f"共找到 {len(self.all_product_data)} 個可用日期")
            else:
                print("未找到任何可用日期")
                
        except Exception as e:
            print(f"爬蟲過程中發生錯誤: {e}")
        finally:
            self.driver.quit()
            print("瀏覽器已關閉")

class KKdayMultiScraper:
    def __init__(self, urls):
        self.urls = urls if isinstance(urls, list) else [urls]

    def scrape_url(self, url):
        """爬取單個URL"""
        print(f"\n開始爬取 URL: {url}")
        scraper = KKdayFlightScraper(url)
        try:
            scraper.run(months_to_scrape=3)
        except Exception as e:
            print(f"處理 URL {url} 時發生錯誤: {e}")
        finally:
            if hasattr(scraper, 'driver'):
                scraper.driver.quit()

    def run(self):
        """依序執行爬蟲"""
        print(f"開始爬取 {len(self.urls)} 個URLs")
        
        for url in self.urls:
            self.scrape_url(url)
            # 在每個URL之間添加延遲，避免過於頻繁的請求
            time.sleep(random.uniform(3, 5))

def main():
    # 要爬取的URL列表
    urls = [
        "https://www.kkday.com/zh-tw/product/137240",
        "https://www.kkday.com/zh-tw/product/139665",
        "https://www.kkday.com/zh-tw/product/146962"
    ]
    
    # 初始化多URL爬蟲器（設置同時爬取的最大URL數量為3）
    scraper = KKdayMultiScraper(urls)
    
    # 執行爬蟲
    scraper.run()

if __name__ == "__main__":
    main()