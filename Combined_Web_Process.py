from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import time
import os
import subprocess
import shutil
import re
import pandas as pd


class UnifiedAutomation:
    def __init__(
        self,
        login_url: str,
        login_id: str,
        password: str,
        use_vpn: bool = True,
        vpn_exe: str = r"C:\Program Files\OpenVPN Connect\OpenVPNConnect.exe",
        vpn_shortcut_id: str = "1752582150336",
        download_dir: str | None = None,
    ):
        self.login_url = login_url
        self.login_id = login_id
        self.password = password
        self.use_vpn = use_vpn
        self.vpn_exe = vpn_exe
        self.vpn_shortcut_id = vpn_shortcut_id
        self.download_dir = download_dir or os.path.dirname(os.path.abspath(__file__))
        self.driver = None

    # ========== Helper Methods ==========
    def _vpn(self, action: str, wait: int = 0):
        if not self.use_vpn: return
        arg = f"--{action}-shortcut={self.vpn_shortcut_id}"
        subprocess.Popen([self.vpn_exe, arg])
        if wait: time.sleep(wait)

    def _make_driver(self, download_mode=False):
        chrome_opts = webdriver.ChromeOptions()
        if download_mode:
            prefs = {
                "download.default_directory": self.download_dir,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_settings.popups": 0
            }
            chrome_opts.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(options=chrome_opts)
        self.driver.implicitly_wait(5)

    def _login(self):
        d = self.driver
        print("Navigating to login page…")
        d.get(self.login_url)
        WebDriverWait(d, 15).until(EC.presence_of_element_located((By.NAME, "username")))
        d.find_element(By.NAME, "username").send_keys(self.login_id)
        d.find_element(By.NAME, "password").send_keys(self.password)
        d.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
        print("Submitting credentials…")
        WebDriverWait(d, 15).until(EC.url_changes(self.login_url))
        print("Logged in successfully.")

    @staticmethod
    def _get_first_of_current_month() -> str:
        return datetime.now().replace(day=1).strftime("%Y-%m-%d")

    @staticmethod
    def _get_previous_month_dates():
        current_date = datetime.now()
        previous_month_date = current_date - relativedelta(months=1)
        start_date = previous_month_date.replace(day=1)
        end_date = (start_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        return start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")

    @staticmethod
    def _build_target_url(base_url: str, start_date: str, end_date: str,
                          start_param: str = "ApptStartDate", end_param: str = "ApptEndDate") -> str:
        sep = '&' if '?' in base_url else '?'
        return f"{base_url}{sep}{start_param}={start_date}&{end_param}={end_date}"

    # ========== Download Methods ==========
    def _start_report_download(self, base_url, label, start_date=None, end_date=None,
                               start_param="ApptStartDate", end_param="ApptEndDate",
                               start_timeout=15):
        d = self.driver
        
        # Build URL with dates if provided
        if start_date and end_date:
            target = self._build_target_url(base_url, start_date, end_date, start_param, end_param)
        else:
            target = base_url
            
        print(f"[{label}] Opening report: {target}")
        d.get(target)

        WebDriverWait(d, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "svg.Icon-download")))
        d.find_element(By.CSS_SELECTOR, "svg.Icon-download").click()

        # Wait specifically for XLSX button
        print(f"[{label}] Looking for XLSX option...")
        WebDriverWait(d, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "button.text-white-hover.bg-brand-hover"))
        )
        buttons = d.find_elements(By.CSS_SELECTOR, "button.text-white-hover.bg-brand-hover")
        xlsx_btn = None
        for b in buttons:
            try:
                b.find_element(By.CSS_SELECTOR, "svg.Icon-xlsx")
                xlsx_btn = b
                break
            except:
                pass
        if not xlsx_btn:
            raise RuntimeError(f"[{label}] XLSX option not found.")

        before = set(os.listdir(self.download_dir))
        click_time = time.time()
        xlsx_btn.click()
        print(f"[{label}] XLSX download clicked.")

        expected_final_name = None
        final_path = None
        t0 = time.time()
        while time.time() - t0 < start_timeout:
            after = set(os.listdir(self.download_dir))
            new_files = list(after - before)
            if new_files:
                newest = max(new_files, key=lambda f: os.path.getmtime(os.path.join(self.download_dir, f)))
                if newest.endswith(".crdownload"):
                    expected_final_name = newest[:-len(".crdownload")]
                    print(f"[{label}] Download started (temp): {newest} -> expecting {expected_final_name}")
                    break
                elif newest.endswith(".xlsx"):
                    final_path = os.path.join(self.download_dir, newest)
                    print(f"[{label}] Download completed quickly: {final_path}")
                    break
            time.sleep(0.5)

        return {
            "label": label,
            "expected_final_name": expected_final_name,
            "final_path": final_path,
            "click_time": click_time
        }

    def _wait_for_download(self, handle, timeout=300):
        print(f"Waiting for download to complete...")
        end = time.time() + timeout
        
        while time.time() < end:
            expected = handle.get("expected_final_name")
            if expected:
                candidate = os.path.join(self.download_dir, expected)
                if os.path.exists(candidate) and not os.path.exists(candidate + ".crdownload"):
                    print(f"[{handle['label']}] Download complete: {candidate}")
                    return candidate
            
            for f in os.listdir(self.download_dir):
                if not f.endswith(".xlsx"):
                    continue
                path = os.path.join(self.download_dir, f)
                if os.path.getmtime(path) >= handle["click_time"] - 1:
                    if not os.path.exists(path + ".crdownload"):
                        print(f"[{handle['label']}] Download complete: {path}")
                        return path
            
            time.sleep(1)
        
        raise TimeoutError(f"Download timed out for {handle['label']}")

    def _rename_downloaded_file(self, original_path: str, label: str) -> str:
        clean_label = label.replace(" ", "_").replace("/", "-")
        new_filename = f"{clean_label}.xlsx"
        new_path = os.path.join(self.download_dir, new_filename)
        
        if os.path.exists(new_path):
            counter = 1
            while os.path.exists(new_path):
                new_filename = f"{clean_label}_{counter}.xlsx"
                new_path = os.path.join(self.download_dir, new_filename)
                counter += 1
        
        try:
            shutil.move(original_path, new_path)
            print(f"[{label}] Renamed file to: {new_filename}")
            return new_path
        except Exception as e:
            print(f"[{label}] Warning: Could not rename file. Error: {e}")
            return original_path

    def _wait_for_all_downloads(self, handles, timeout=300):
        print("Waiting for all downloads to complete…")
        end = time.time() + timeout
        results = {h["label"]: h.get("final_path") for h in handles}
        assigned_files = set(os.path.basename(p) for p in results.values() if p)

        while time.time() < end:
            for h in handles:
                label = h["label"]
                if results[label]:
                    continue

                expected = h.get("expected_final_name")
                if expected:
                    candidate = os.path.join(self.download_dir, expected)
                    if os.path.exists(candidate) and not os.path.exists(candidate + ".crdownload"):
                        if os.path.basename(candidate) not in assigned_files:
                            results[label] = candidate
                            assigned_files.add(os.path.basename(candidate))
                            print(f"[{label}] Finished: {candidate}")
                            continue

                for f in os.listdir(self.download_dir):
                    if not f.endswith(".xlsx") or f in assigned_files:
                        continue
                    path = os.path.join(self.download_dir, f)
                    if os.path.getmtime(path) >= h["click_time"] - 1:
                        if not os.path.exists(path + ".crdownload"):
                            results[label] = path
                            assigned_files.add(f)
                            print(f"[{label}] Finished (fallback): {path}")
                            break

            if all(results.values()):
                any_cr = any(name.endswith(".crdownload") for name in os.listdir(self.download_dir))
                if not any_cr:
                    break

            time.sleep(1)

        unresolved = [lbl for lbl, p in results.items() if not p]
        if unresolved:
            raise TimeoutError(f"Timed out waiting for: {', '.join(unresolved)}")

        # Rename all downloaded files
        for h in handles:
            label = h["label"]
            if results[label]:
                new_path = self._rename_downloaded_file(results[label], label)
                results[label] = new_path

        return results
    
    def save_results_to_excel(self, results, filename="automation_summary.xlsx"):
        """Save scraped data to an Excel file"""
        excel_path = os.path.join(self.download_dir, filename)
        
        # Prepare data for Excel - all in one list
        all_data = []
        
        for key, value in results.items():
            # Check if it's a download task by looking for specific keywords in the key
            if any(keyword in key for keyword in ['Report', 'Transactions', 'TV4', 'Enrollments', 'Memberships']) and isinstance(value, str) and '.xlsx' in value:
                all_data.append({
                    'Type': 'Downloaded File',
                    'Name': key,
                    'Value/Path': value,
                    'Timestamp': datetime.now()
                })
            else:
                all_data.append({
                    'Type': 'Scraped Data',
                    'Name': key,
                    'Value/Path': value,
                    'Timestamp': datetime.now()
                })
        
        # Create Excel with single sheet
        df = pd.DataFrame(all_data)
        df.to_excel(excel_path, sheet_name='Summary', index=False)
        
        print(f"\nSummary saved to: {excel_path}")
        return excel_path

    # ========== Task Implementations ==========
    def instacart_downloads(self):
        """Download all Instacart reports"""
        start_date, end_date = self._get_previous_month_dates()
        
        reports = [
            {"url": "https://metabase.caradvise.com/question/1499", "label": "TV4", 
             "start_param": "ApptStartDate", "end_param": "ApptEndDate"},
            {"url": "https://metabase.caradvise.com/question/1262", "label": "Instacart_Canada_Memberships",
             "start_param": "ApptStartDate", "end_param": "ApptEndDate"},
            {"url": "https://metabase.caradvise.com/question/1764?Affiliate=Instacart", "label": "Instacart_US_Enrollments",
             "start_param": "EnrolledStart", "end_param": "EnrolledEnd"},
            {"url": "https://metabase.caradvise.com/question/1764?Affiliate=Instacart%20Canada", "label": "Instacart_Canada_Enrollments",
             "start_param": "EnrolledStart", "end_param": "EnrolledEnd"},
        ]
        
        handles = []
        for report in reports:
            handle = self._start_report_download(
                report["url"], report["label"], start_date, end_date,
                report["start_param"], report["end_param"]
            )
            handles.append(handle)
        
        results = self._wait_for_all_downloads(handles, timeout=900)
        return results

    def clearcover_download(self):
        """Download Clearcover report"""
        handle = self._start_report_download("https://metabase.caradvise.com/question/1058", "Clearcover_Report")
        downloaded_path = self._wait_for_download(handle)
        renamed_path = self._rename_downloaded_file(downloaded_path, "Clearcover_Report")
        return {"Clearcover_Report": renamed_path}

    def grubhub_scrape(self):
        """Scrape Grubhub Premium & Elite memberships"""
        date_str = self._get_first_of_current_month()
        dashboard_url = "https://metabase.caradvise.com/dashboard/58?id=723&id=1769&id=4509&id=4511&id=4512&id=4513"
        target_url = f"{dashboard_url}&date_filter=~{date_str}"
        
        print(f"Navigating to dashboard: {target_url}")
        self.driver.get(target_url)

        results = {}

        # Scrape Premium Memberships
        premium_xpath = "//div[contains(@class, 'Card') and .//div[normalize-space()='Premium Memberships']]//h1[contains(@class, 'ScalarValue')]"
        print("Waiting for the 'Premium Memberships' card to load...")
        try:
            value_element = WebDriverWait(self.driver, 45).until(
                EC.visibility_of_element_located((By.XPATH, premium_xpath))
            )
            results["Grubhub_Premium_Memberships"] = int(value_element.text.replace(",", ""))
        except:
            results["Grubhub_Premium_Memberships"] = None

        # Scrape Elite Memberships
        elite_xpath = "//div[contains(@class, 'Card') and .//div[normalize-space()='Elite Memberships']]//h1[contains(@class, 'ScalarValue')]"
        try:
            value_element = self.driver.find_element(By.XPATH, elite_xpath)
            results["Grubhub_Elite_Memberships"] = int(value_element.text.replace(",", ""))
        except:
            results["Grubhub_Elite_Memberships"] = None

        return results

    def shipt_scrape(self):
        """Scrape Shipt Premium & Elite memberships"""
        date_str = self._get_first_of_current_month()
        dashboard_url = "https://metabase.caradvise.com/dashboard/58?id=3966&id=3973&id=3969&id=3970&id=3972"
        target_url = f"{dashboard_url}&date_filter=~{date_str}"
        
        print(f"Navigating to dashboard: {target_url}")
        self.driver.get(target_url)

        results = {}

        # Scrape Premium Memberships
        premium_xpath = "//div[contains(@class, 'Card') and .//div[normalize-space()='Premium Memberships']]//h1[contains(@class, 'ScalarValue')]"
        print("Waiting for the 'Premium Memberships' card to load...")
        try:
            value_element = WebDriverWait(self.driver, 45).until(
                EC.visibility_of_element_located((By.XPATH, premium_xpath))
            )
            results["Shipt_Premium_Memberships"] = int(value_element.text.replace(",", ""))
        except:
            results["Shipt_Premium_Memberships"] = None

        # Scrape Elite Memberships
        elite_xpath = "//div[contains(@class, 'Card') and .//div[normalize-space()='Elite Memberships']]//h1[contains(@class, 'ScalarValue')]"
        try:
            value_element = self.driver.find_element(By.XPATH, elite_xpath)
            results["Shipt_Elite_Memberships"] = int(value_element.text.replace(",", ""))
        except:
            results["Shipt_Elite_Memberships"] = None

        return results

    def sunland_scrape(self):
        """Scrape Sunland row count"""
        target_url = "https://metabase.caradvise.com/question/909?ActiveOnly=1&FleetAffiliate=Sunland"
        
        print(f"Navigating to: {target_url}")
        self.driver.get(target_url)

        results = {}
        rows_xpath = "//span[contains(text(), 'Showing') and contains(text(), 'rows')]"
        
        print("Waiting for the row count to load...")
        try:
            value_element = WebDriverWait(self.driver, 45).until(
                EC.visibility_of_element_located((By.XPATH, rows_xpath))
            )
            full_text = value_element.text
            match = re.search(r'Showing (\d+) rows', full_text)
            if match:
                results["Sunland_Row_Count"] = int(match.group(1))
            else:
                results["Sunland_Row_Count"] = None
        except:
            results["Sunland_Row_Count"] = None

        return results

    def sunland_download(self):
        """Download Sunland transactions report"""
        start_date, end_date = self._get_previous_month_dates()
        handle = self._start_report_download(
            "https://metabase.caradvise.com/question/1276", 
            "Sunland_Transactions", 
            start_date, 
            end_date
        )
        downloaded_path = self._wait_for_download(handle)
        renamed_path = self._rename_downloaded_file(downloaded_path, "Sunland_Transactions")
        return {"Sunland_Transactions": renamed_path}

    def uber_scrape(self):
        """Scrape Uber Paid Memberships"""
        date_str = self._get_first_of_current_month()
        dashboard_url = "https://metabase.caradvise.com/dashboard/58?id=485&id=486&id=487&id=488"
        target_url = f"{dashboard_url}&date_filter=~{date_str}"
        
        print(f"Navigating to dashboard: {target_url}")
        self.driver.get(target_url)

        results = {}
        paid_xpath = "//div[contains(@class, 'Card') and .//div[normalize-space()='Paid Memberships']]//h1[contains(@class, 'ScalarValue')]"
        
        print("Waiting for the 'Paid Memberships' card to load...")
        try:
            value_element = WebDriverWait(self.driver, 45).until(
                EC.visibility_of_element_located((By.XPATH, paid_xpath))
            )
            results["Uber_Paid_Memberships"] = int(value_element.text.replace(",", ""))
        except:
            results["Uber_Paid_Memberships"] = None

        return results

    # ========== Task Configuration ==========
    def tasks_config(self):
        return {
            "1": {"name": "Instacart - Download All Reports", "method": self.instacart_downloads, "type": "download"},
            "2": {"name": "Clearcover - Download Report", "method": self.clearcover_download, "type": "download"},
            "3": {"name": "Grubhub - Scrape Memberships", "method": self.grubhub_scrape, "type": "scrape"},
            "4": {"name": "Shipt - Scrape Memberships", "method": self.shipt_scrape, "type": "scrape"},
            "5": {"name": "Sunland - Scrape Row Count", "method": self.sunland_scrape, "type": "scrape"},
            "6": {"name": "Sunland - Download Transactions", "method": self.sunland_download, "type": "download"},
            "7": {"name": "Uber - Scrape Paid Memberships", "method": self.uber_scrape, "type": "scrape"},
        }

    # ========== Public Runner ==========
    def run(self, selected_tasks: list) -> dict:
        self._vpn("connect", wait=10)
        all_results = {}
        
        try:
            # Determine if we need download capabilities
            tasks = self.tasks_config()
            needs_download = any(tasks[task_id]["type"] == "download" for task_id in selected_tasks)
            
            self._make_driver(download_mode=needs_download)
            self._login()
            
            # Execute selected tasks
            for task_id in selected_tasks:
                if task_id in tasks:
                    task = tasks[task_id]
                    print(f"\n--- Executing: {task['name']} ---")
                    try:
                        results = task["method"]()
                        all_results.update(results)
                    except Exception as e:
                        print(f"Error in {task['name']}: {e}")
                        # Continue with other tasks
            
            return all_results
            
        finally:
            if self.driver:
                print("\nClosing browser…")
                self.driver.quit()
            self._vpn("disconnect")


# ==================== MAIN EXECUTION ====================
if __name__ == "__main__":
    automation = UnifiedAutomation(
        login_url="https://metabase.caradvise.com/auth/login",
        login_id= input("Enter the Login ID : "),
        password= input("Enter the Password : "),
        use_vpn=True,
    )

    # Show menu
    tasks = automation.tasks_config()
    print("\n" + "="*50)
    print("Complete Web Scraping")
    print("="*50)
    print("\nAvailable tasks:")
    for task_id, task in tasks.items():
        print(f"{task_id}. {task['name']}")
    
    print("\nEnter task numbers separated by commas (e.g., 1,3,5)")
    print("Or press Enter to run ALL tasks")
    choice = input("\nYour selection: ").strip()
    
    if not choice:
        selected = list(tasks.keys())
        print("\nRunning ALL tasks...")
    else:
        selected = [t.strip() for t in choice.split(",") if t.strip() in tasks]
        if not selected:
            print("No valid selection. Exiting.")
            exit(1)
    
    try:
        results = automation.run(selected)
        if results:
            excel_path = automation.save_results_to_excel(results)
        
        print("\n" + "="*50)
        print("RESULTS")
        print("="*50)
        
        # Display results organized by type
        if results:
            # Downloads
            download_results = {k: v for k, v in results.items() if k.endswith('.xlsx') or 'Report' in k or 'Transactions' in k}
            if download_results:
                print("\nDownloaded Files:")
                for name, path in download_results.items():
                    print(f"  • {name}: {path}")
            
            # Scraped data
            scrape_results = {k: v for k, v in results.items() if k not in download_results}
            if scrape_results:
                print("\nScraped Data:")
                for name, value in scrape_results.items():
                    print(f"  • {name}: {value}")
        else:
            print("\nNo results to display.")
                
    except Exception as e:
        print(f"\n--- AUTOMATION FAILED ---")
        print(f"Error: {e}")