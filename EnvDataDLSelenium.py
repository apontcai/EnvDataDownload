# login env.nem.com.hk:10027 to download data
# by NEM (Novox E&M Limited)
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from datetime import datetime, timedelta
import os
import re
import subprocess
import platform
import time
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# Hide console window when running as exe
if hasattr(sys, '_MEIPASS'):
    # Running as PyInstaller bundle
    import ctypes

    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)


class DataDownloader:
    def __init__(self, root):
        self.root = root
        self.root.title("Equipment Data Downloader")
        self.root.geometry("600x400")

        # Variables
        self.excel_file_path = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Excel file selection
        ttk.Label(main_frame, text="Select Excel File:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.excel_file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(main_frame, text="Browse", command=self.browse_excel_file).grid(row=0, column=2, padx=5, pady=5)

        # Download folder info and open button
        download_info_frame = ttk.Frame(main_frame)
        download_info_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(download_info_frame, text=f"Downloads will be saved to: {self.get_default_download_folder()}").pack(
            side=tk.LEFT)
        ttk.Button(download_info_frame, text="Open Download Folder", command=self.open_download_folder).pack(
            side=tk.RIGHT, padx=5)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)

        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready to start")
        self.status_label.grid(row=3, column=0, columnspan=3, pady=5)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)

        ttk.Button(button_frame, text="Preview Data", command=self.preview_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Start Download", command=self.start_download).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Exit", command=self.root.quit).pack(side=tk.LEFT, padx=5)

        # Text area for logs
        self.log_text = tk.Text(main_frame, height=15, width=70)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=10)

        # Scrollbar for text area
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=5, column=3, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

    def get_default_download_folder(self):
        """Get the system default download folder"""
        if platform.system() == "Windows":
            return os.path.join(os.path.expanduser("~"), "Downloads")
        elif platform.system() == "Darwin":  # macOS
            return os.path.join(os.path.expanduser("~"), "Downloads")
        else:  # Linux
            return os.path.join(os.path.expanduser("~"), "Downloads")

    def open_download_folder(self):
        """Open the default download folder in file explorer"""
        download_folder = self.get_default_download_folder()
        try:
            if platform.system() == "Windows":
                os.startfile(download_folder)
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", download_folder])
            else:  # Linux
                subprocess.run(["xdg-open", download_folder])
            self.log_message(f"Opened download folder: {download_folder}")
        except Exception as e:
            self.log_message(f"Error opening download folder: {str(e)}")

    def browse_excel_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.excel_file_path.set(filename)

    def log_message(self, message):
        """Add message to log text area"""
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.root.update()

    def clear_log(self):
        """Clear the log text area"""
        self.log_text.delete(1.0, tk.END)

    def update_status(self, status):
        """Update status label"""
        self.status_label.config(text=status)
        self.root.update()

    def cell_to_string(self, cell_value):
        """Convert Excel cell value to string safely"""
        if cell_value is None:
            return ""
        return str(cell_value).strip()

    def parse_excel_date(self, cell_value):
        """Parse Excel date which might be a formula or direct date"""
        if cell_value is None:
            return datetime.now().date()

        # Convert to string first
        cell_str = str(cell_value).strip()

        if isinstance(cell_value, str) and cell_str.startswith('='):
            # Handle Excel functions like =today()-1
            formula = cell_str.lower()
            if 'today()' in formula:
                today = datetime.now().date()
                # Extract number after today()
                match = re.search(r'today\(\)\s*([+-])\s*(\d+)', formula)
                if match:
                    operator, days = match.groups()
                    days = int(days)
                    if operator == '+':
                        return today + timedelta(days=days)
                    else:
                        return today - timedelta(days=days)
                return today
        elif isinstance(cell_value, datetime):
            return cell_value.date()
        else:
            # Try to parse as date string
            try:
                # Handle different date formats
                if len(cell_str) == 10 and '-' in cell_str:  # YYYY-MM-DD
                    return datetime.strptime(cell_str, '%Y-%m-%d').date()
                elif len(cell_str) == 8 and cell_str.isdigit():  # YYYYMMDD
                    return datetime.strptime(cell_str, '%Y%m%d').date()
                else:
                    # If it's a number, treat as Excel date serial
                    if cell_str.replace('.', '').isdigit():
                        # Excel date serial number
                        excel_date = float(cell_str)
                        # Excel epoch starts from 1900-01-01, but has a leap year bug
                        base_date = datetime(1899, 12, 30)  # Adjusted for Excel's leap year bug
                        return (base_date + timedelta(days=excel_date)).date()
                    else:
                        return datetime.now().date()
            except:
                return datetime.now().date()

    def read_excel_data(self):
        """Read and parse Excel file"""
        try:
            workbook = openpyxl.load_workbook(self.excel_file_path.get())
            sheet = workbook.active

            # Read configuration - convert all to strings safely
            website = self.cell_to_string(sheet['B1'].value)
            username = self.cell_to_string(sheet['B2'].value)
            password = self.cell_to_string(sheet['B3'].value)

            # Read dates
            start_date_cell = sheet['F1'].value
            end_date_cell = sheet['F2'].value

            start_date = self.parse_excel_date(start_date_cell)

            if end_date_cell:
                end_date = self.parse_excel_date(end_date_cell)
            else:
                end_date = start_date

            # Read equipment SNs
            equipment_sns = []
            row = 6
            while True:
                sn_cell = sheet[f'A{row}'].value
                if sn_cell is None:
                    break

                sn = self.cell_to_string(sn_cell)
                if sn == '':
                    break

                equipment_sns.append(sn)
                row += 1

            return {
                'website': website,
                'username': username,
                'password': password,
                'start_date': start_date,
                'end_date': end_date,
                'equipment_sns': equipment_sns
            }

        except Exception as e:
            raise Exception(f"Error reading Excel file: {str(e)}")

    def preview_data(self):
        """Preview the data that will be processed and display in log area"""
        if not self.excel_file_path.get():
            messagebox.showerror("Error", "Please select an Excel file first")
            return

        try:
            data = self.read_excel_data()

            # Clear log and show preview data
            self.clear_log()
            self.log_message("=" * 50)
            self.log_message("EXCEL DATA PREVIEW")
            self.log_message("=" * 50)

            # Safe password masking
            password_display = '*' * len(data['password']) if data['password'] else 'None'

            self.log_message(f"Website: {data['website']}")
            self.log_message(f"Username: {data['username']}")
            self.log_message(f"Password: {password_display}")
            self.log_message(f"Start Date: {data['start_date']}")
            self.log_message(f"End Date: {data['end_date']}")
            self.log_message(f"Equipment SNs ({len(data['equipment_sns'])}):")

            for i, sn in enumerate(data['equipment_sns'], 1):
                self.log_message(f"  {i:2d}. {sn}")

            self.log_message("=" * 50)

            if not data['website'] or not data['username'] or not data['password']:
                self.log_message("⚠️  WARNING: Website, username, or password is missing!")

            if not data['equipment_sns']:
                self.log_message("⚠️  WARNING: No equipment SNs found!")
            else:
                self.log_message(f"✓ Ready to process {len(data['equipment_sns'])} equipment(s)")

            self.log_message(f"✓ Downloads will be saved to: {self.get_default_download_folder()}")

        except Exception as e:
            self.log_message(f"❌ Error reading Excel file: {str(e)}")

    def get_export_url(self, base_url):
        """Generate the export URL from base URL"""
        # Extract the base domain from the login URL
        if 'env.nem.com.hk:10027' in base_url:
            return 'https://env.nem.com.hk:10027/syntheticSystem/dataAnalysis/export'
        else:
            # For other domains, try to construct the URL
            from urllib.parse import urlparse
            parsed = urlparse(base_url)
            return f"{parsed.scheme}://{parsed.netloc}/syntheticSystem/dataAnalysis/export"

    def find_input_by_selectors(self, driver, selectors, timeout=5):
        """Try to find input element using multiple selectors"""
        wait = WebDriverWait(driver, timeout)

        for selector in selectors:
            try:
                if selector.startswith('input['):
                    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    return element
                elif selector.startswith('#'):
                    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    return element
                else:
                    element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, selector)))
                    return element
            except (TimeoutException, NoSuchElementException):
                continue

        return None

    def download_data_for_sn(self, driver, sn, start_date, end_date):
        """Download data for a specific equipment SN"""
        try:
            self.log_message(f"Processing SN: {sn}")

            # Set date range
            start_date_str = start_date.strftime('%Y-%m-%d')
            end_date_str = end_date.strftime('%Y-%m-%d')

            self.log_message(f"  Date range: {start_date_str} to {end_date_str}")

            # Wait for page to load completely
            time.sleep(2)

            # Select real-time values (實時值) radio button
            try:
                radio_selectors = [
                    'label:has-text("實時值")',
                    'input[value="實時值"]',
                    'label.is-active > span:has-text("實時值")',
                    'text=實時值'
                ]

                radio_selected = False
                for selector in radio_selectors:
                    try:
                        if selector == 'text=實時值':
                            element = driver.find_element(By.XPATH, "//*[contains(text(), '實時值')]")
                            element.click()
                        elif 'has-text' in selector:
                            if 'label' in selector and 'span' in selector:
                                element = driver.find_element(By.XPATH,
                                                              "//label[@class='is-active']//span[contains(text(), '實時值')]")
                            else:
                                element = driver.find_element(By.XPATH, "//label[contains(text(), '實時值')]")
                            element.click()
                        else:
                            element = driver.find_element(By.CSS_SELECTOR, selector)
                            element.click()
                        radio_selected = True
                        self.log_message("  ✓ Selected 實時值 option")
                        break
                    except:
                        continue

                if not radio_selected:
                    self.log_message("  ⚠️  Could not select 實時值 option, continuing anyway...")

                time.sleep(0.5)
            except Exception as e:
                self.log_message(f"  ⚠️  Error selecting real-time values: {str(e)}")

            # Fill start date - IMPROVED APPROACH
            try:
                start_date_selectors = [
                    'input[placeholder*="開始時間"]',
                    'input[placeholder*="开始时间"]',
                    'input[aria-label*="開始時間"]',
                    'input[aria-label*="开始时间"]',
                    'div.flex-wrap > div:nth-of-type(2) input:nth-of-type(1)',
                    'input[type="text"]'
                ]

                start_filled = False
                for selector in start_date_selectors:
                    try:
                        element = driver.find_element(By.CSS_SELECTOR, selector)

                        # Method 1: Click to focus, clear, type, and confirm
                        element.click()  # Focus the element first
                        time.sleep(0.3)

                        # Clear field using multiple methods
                        element.clear()
                        element.send_keys(Keys.CONTROL + "a")  # Select all
                        element.send_keys(Keys.DELETE)  # Delete selected
                        time.sleep(0.2)

                        # Type date slowly
                        for char in start_date_str:
                            element.send_keys(char)
                            time.sleep(0.05)

                        time.sleep(0.3)
                        element.send_keys(Keys.ENTER)
                        time.sleep(0.5)

                        # Verify the value was set
                        current_value = element.get_attribute('value')
                        if start_date_str in current_value:
                            start_filled = True
                            self.log_message(f"  ✓ Filled start date: {start_date_str}")
                            break
                        else:
                            self.log_message(
                                f"  ⚠️  Start date verification failed. Expected: {start_date_str}, Got: {current_value}")

                            # Method 2: Try JavaScript if normal method failed
                            try:
                                driver.execute_script(f"arguments[0].value = '{start_date_str}';", element)
                                driver.execute_script(
                                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", element)
                                driver.execute_script(
                                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", element)
                                time.sleep(0.3)

                                current_value = element.get_attribute('value')
                                if start_date_str in current_value:
                                    start_filled = True
                                    self.log_message(f"  ✓ Filled start date via JavaScript: {start_date_str}")
                                    break
                            except:
                                continue

                    except:
                        continue

                if not start_filled:
                    self.log_message(f"  ❌ Could not fill start date: {start_date_str}")

                time.sleep(0.5)
            except Exception as e:
                self.log_message(f"  ❌ Error filling start date: {str(e)}")

            # Fill end date - IMPROVED APPROACH
            try:
                end_date_selectors = [
                    'input[placeholder*="結束時間"]',
                    'input[placeholder*="结束时间"]',
                    'input[aria-label*="結束時間"]',
                    'input[aria-label*="结束时间"]',
                    'div.flex-wrap > div:nth-of-type(2) input:nth-of-type(2)',
                    'input[type="text"]:nth-of-type(2)'
                ]

                end_filled = False
                for selector in end_date_selectors:
                    try:
                        element = driver.find_element(By.CSS_SELECTOR, selector)

                        # Method 1: Click to focus, clear, type, and confirm
                        element.click()  # Focus the element first
                        time.sleep(0.3)

                        # Clear field using multiple methods
                        element.clear()
                        element.send_keys(Keys.CONTROL + "a")  # Select all
                        element.send_keys(Keys.DELETE)  # Delete selected
                        time.sleep(0.2)

                        # Type date slowly
                        for char in end_date_str:
                            element.send_keys(char)
                            time.sleep(0.05)

                        time.sleep(0.3)
                        element.send_keys(Keys.ENTER)
                        time.sleep(0.5)

                        # Verify the value was set
                        current_value = element.get_attribute('value')
                        if end_date_str in current_value:
                            end_filled = True
                            self.log_message(f"  ✓ Filled end date: {end_date_str}")
                            break
                        else:
                            self.log_message(
                                f"  ⚠️  End date verification failed. Expected: {end_date_str}, Got: {current_value}")

                            # Method 2: Try JavaScript if normal method failed
                            try:
                                driver.execute_script(f"arguments[0].value = '{end_date_str}';", element)
                                driver.execute_script(
                                    "arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", element)
                                driver.execute_script(
                                    "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", element)
                                time.sleep(0.3)

                                current_value = element.get_attribute('value')
                                if end_date_str in current_value:
                                    end_filled = True
                                    self.log_message(f"  ✓ Filled end date via JavaScript: {end_date_str}")
                                    break
                            except:
                                continue

                    except:
                        continue

                if not end_filled:
                    self.log_message(f"  ❌ Could not fill end date: {end_date_str}")

                time.sleep(0.5)
            except Exception as e:
                self.log_message(f"  ❌ Error filling end date: {str(e)}")

            # Enter equipment SN
            try:
                sn_selectors = [
                    'input[placeholder*="設備號"]',
                    'input[placeholder*="设备号"]',
                    'input[placeholder*="請輸入設備號"]',
                    'input[placeholder*="请输入设备号"]',
                    'input[aria-label*="設備號"]',
                    'input[aria-label*="设备号"]',
                    '#el-id-215-53'
                ]

                sn_filled = False
                for selector in sn_selectors:
                    try:
                        element = driver.find_element(By.CSS_SELECTOR, selector)
                        element.clear()
                        time.sleep(0.2)
                        element.send_keys(sn)
                        element.send_keys(Keys.ENTER)
                        sn_filled = True
                        self.log_message(f"  ✓ Filled equipment SN: {sn}")
                        break
                    except:
                        continue

                if not sn_filled:
                    self.log_message(f"  ❌ Could not fill equipment SN: {sn}")
                    return False

                time.sleep(1)
            except Exception as e:
                self.log_message(f"  ❌ Error filling equipment SN {sn}: {str(e)}")
                return False

            # Click query button
            try:
                query_selectors = [
                    'button:has-text("查詢")',
                    'button:has-text("查询")',
                    'text=查詢',
                    'text=查询',
                    'button.el-button--warning',
                    '[role="button"]:has-text("查詢")',
                    '[role="button"]:has-text("查询")'
                ]

                query_clicked = False
                for selector in query_selectors:
                    try:
                        if selector.startswith('text='):
                            text = selector.replace('text=', '')
                            element = driver.find_element(By.XPATH, f"//*[contains(text(), '{text}')]")
                            element.click()
                        elif 'has-text' in selector:
                            if 'button' in selector:
                                text = selector.split('"')[1]
                                element = driver.find_element(By.XPATH, f"//button[contains(text(), '{text}')]")
                            else:
                                text = selector.split('"')[1]
                                element = driver.find_element(By.XPATH,
                                                              f"//*[@role='button'][contains(text(), '{text}')]")
                            element.click()
                        else:
                            element = driver.find_element(By.CSS_SELECTOR, selector)
                            element.click()
                        query_clicked = True
                        self.log_message("  ✓ Clicked query button")
                        break
                    except:
                        continue

                if not query_clicked:
                    self.log_message("  ❌ Could not click query button")
                    return False

                # Wait for data to load
                time.sleep(5)
                self.log_message("  ⏳ Waiting for data to load...")

            except Exception as e:
                self.log_message(f"  ❌ Error clicking query button: {str(e)}")
                return False

            # Check if data exists and click download
            try:
                download_selectors = [
                    'button:has-text("導出文件")',
                    'button:has-text("导出文件")',
                    'text=導出文件',
                    'text=导出文件',
                    '[role="button"]:has-text("導出文件")',
                    '[role="button"]:has-text("导出文件")',
                    'div:has-text("導出文件")'
                ]

                download_clicked = False
                for selector in download_selectors:
                    try:
                        if selector.startswith('text='):
                            text = selector.replace('text=', '')
                            element = driver.find_element(By.XPATH, f"//*[contains(text(), '{text}')]")
                            element.click()
                        elif 'has-text' in selector:
                            if 'button' in selector:
                                text = selector.split('"')[1]
                                element = driver.find_element(By.XPATH, f"//button[contains(text(), '{text}')]")
                            elif 'div' in selector:
                                text = selector.split('"')[1]
                                element = driver.find_element(By.XPATH, f"//div[contains(text(), '{text}')]")
                            else:
                                text = selector.split('"')[1]
                                element = driver.find_element(By.XPATH,
                                                              f"//*[@role='button'][contains(text(), '{text}')]")
                            element.click()
                        else:
                            element = driver.find_element(By.CSS_SELECTOR, selector)
                            element.click()
                        download_clicked = True
                        self.log_message(f"  ✓ Download initiated for SN: {sn}")
                        break
                    except:
                        continue

                if download_clicked:
                    time.sleep(3)  # Wait for download to start
                    return True
                else:
                    self.log_message(f"  ⚠️  No data available or could not find download button for SN: {sn}")
                    return False

            except Exception as e:
                self.log_message(f"  ❌ Error during download for SN {sn}: {str(e)}")
                return False

        except Exception as e:
            self.log_message(f"❌ Error processing SN {sn}: {str(e)}")
            return False

    def run_automation(self, data):
        """Run the web automation process"""
        # Setup Chrome options
        chrome_options = Options()

        # Set download directory
        download_folder = self.get_default_download_folder()
        prefs = {
            "download.default_directory": download_folder,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)

        # Optional: run in headless mode (uncomment the line below)
        # chrome_options.add_argument("--headless")

        # Launch browser
        driver = None
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.maximize_window()

            # Navigate to login page
            self.log_message(f"🌐 Navigating to {data['website']}")
            driver.get(data['website'])
            time.sleep(3)

            # Login
            self.log_message("🔐 Logging in...")

            # Fill username
            try:
                username_selectors = [
                    'input[placeholder*="賬號"]',
                    'input[placeholder*="账号"]',
                    'input[placeholder*="用户名"]',
                    'input[aria-label*="賬號"]',
                    'input[aria-label*="账号"]',
                    'input[type="text"]',
                    '#el-id-215-31'
                ]

                username_element = self.find_input_by_selectors(driver, username_selectors)
                if username_element:
                    username_element.clear()
                    username_element.send_keys(data['username'])
                    username_element.send_keys(Keys.TAB)
                    self.log_message(f"✓ Filled username: {data['username']}")
                else:
                    raise Exception("Could not fill username")

            except Exception as e:
                self.log_message(f"❌ Error filling username: {str(e)}")
                return

            # Fill password
            try:
                password_selectors = [
                    'input[placeholder*="密碼"]',
                    'input[placeholder*="密码"]',
                    'input[type="password"]',
                    'input[aria-label*="密碼"]',
                    'input[aria-label*="密码"]',
                    '#el-id-215-32'
                ]

                password_element = self.find_input_by_selectors(driver, password_selectors)
                if password_element:
                    password_element.clear()
                    password_element.send_keys(data['password'])
                    password_element.send_keys(Keys.ENTER)
                    self.log_message("✓ Password filled successfully")
                    time.sleep(5)
                else:
                    raise Exception("Could not fill password")

            except Exception as e:
                self.log_message(f"❌ Error filling password: {str(e)}")
                return

            self.log_message("✅ Login successful")

            # Navigate directly to export page
            export_url = self.get_export_url(data['website'])
            self.log_message(f"📊 Navigating to export page: {export_url}")
            driver.get(export_url)
            time.sleep(3)

            # Process each equipment SN
            total_sns = len(data['equipment_sns'])
            successful_downloads = 0

            self.log_message(f"🚀 Starting to process {total_sns} equipment(s)...")

            for i, sn in enumerate(data['equipment_sns']):
                progress = (i / total_sns) * 100 if total_sns > 0 else 0
                self.progress_var.set(progress)

                self.log_message(f"\n📍 Processing {i + 1}/{total_sns}")
                success = self.download_data_for_sn(
                    driver, sn, data['start_date'], data['end_date']
                )

                if success:
                    successful_downloads += 1

                time.sleep(2)  # Wait between requests

            # Wait for all downloads to complete
            self.log_message("⏳ Waiting for downloads to complete...")
            time.sleep(5)  # Give time for downloads to complete

            self.progress_var.set(100)
            self.log_message(f"\n🎉 Process completed!")
            self.log_message(f"📊 Success rate: {successful_downloads}/{total_sns} downloads successful")
            self.log_message(f"📁 Files saved to: {download_folder}")

        except Exception as e:
            self.log_message(f"❌ Automation error: {str(e)}")
        finally:
            if driver:
                time.sleep(2)  # Give final downloads time to complete
                driver.quit()

    def start_download(self):
        """Start the download process"""
        if not self.excel_file_path.get():
            messagebox.showerror("Error", "Please select an Excel file first")
            return

        try:
            # Read Excel data
            self.update_status("Reading Excel file...")
            data = self.read_excel_data()

            # Validate data
            if not data['website'] or not data['username'] or not data['password']:
                messagebox.showerror("Error", "Website, username, and password are required")
                return

            if not data['equipment_sns']:
                messagebox.showerror("Error", "No equipment SNs found in the Excel file")
                return

            self.log_message("\n" + "=" * 50)
            self.log_message("STARTING DOWNLOAD PROCESS")
            self.log_message("=" * 50)
            self.update_status("Running automation...")

            # Run automation
            self.run_automation(data)

            self.update_status("Process completed")
            messagebox.showinfo("Success", "Download process completed!")

        except Exception as e:
            self.log_message(f"❌ Error: {str(e)}")
            self.update_status("Error occurred")
            messagebox.showerror("Error", str(e))


def main():
    root = tk.Tk()
    app = DataDownloader(root)
    root.mainloop()


if __name__ == "__main__":
    main()