import os
import time
import re
import urllib.parse
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException

from webdriver_manager.chrome import ChromeDriverManager


class WhatsAppSeleniumSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Automated Sender")

        # UI
        tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        self.file_entry = tk.Entry(root, width=50)
        self.file_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(root, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(root, text="Country Code (+):").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        self.cc_entry = tk.Entry(root, width=10)
        self.cc_entry.insert(0, "+2")
        self.cc_entry.grid(row=1, column=1, sticky='w')

        tk.Label(root, text="Chrome profile folder:").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        self.profile_entry = tk.Entry(root, width=50)
        default_profile = os.path.join(os.getcwd(), "selenium_profile")
        self.profile_entry.insert(0, default_profile)
        self.profile_entry.grid(row=2, column=1, padx=5, pady=5)
        tk.Button(root, text="Browse", command=self.browse_profile).grid(row=2, column=2, padx=5, pady=5)

        tk.Label(root, text="Message:").grid(row=3, column=0, padx=5, pady=5, sticky='ne')
        self.message_entry = tk.Text(root, height=6, width=50)
        self.message_entry.grid(row=3, column=1, padx=5, pady=5, columnspan=2)

        tk.Button(root, text="Send Messages", command=self.send_messages, bg="green", fg="white").grid(row=4, column=1, pady=10)

        self.output = scrolledtext.ScrolledText(root, height=18, width=85)
        self.output.grid(row=5, column=0, columnspan=3, padx=5, pady=5)

    def browse_file(self):
        fp = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if fp:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, fp)

    def browse_profile(self):
        folder = filedialog.askdirectory()
        if folder:
            self.profile_entry.delete(0, tk.END)
            self.profile_entry.insert(0, folder)

    def log(self, msg):
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        self.output.insert(tk.END, f"[{ts}] {msg}\n")
        self.output.see(tk.END)
        self.root.update()

    def normalize_number(self, raw, country_code):
        if raw is None:
            return None
        s = str(raw).strip()
        s = re.sub(r"[^\d]", "", s)
        if not s:
            return None
        # remove leading zero and apply country code if needed
        s = s.lstrip("0")
        cc = country_code.strip().lstrip("+")
        if not s.startswith(cc):
            s = cc + s
        if len(s) < 8 or len(s) > 15:
            return None
        return s

    def start_driver(self, profile_path):
        opts = webdriver.ChromeOptions()
        opts.add_argument(f"--user-data-dir={profile_path}")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_argument("--start-maximized")
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=opts)
        driver.set_page_load_timeout(60)
        return driver

    def check_send_status(self, driver, message_text, timeout=10):
        """
        Attempts to determine if the last outgoing message (matching message_text)
        was sent or failed. Returns tuple (bool sent, str reason).
        Reasons: 'msg_found', 'icon_found', 'popup_text', 'timeout', 'no_match'
        """
        start = time.time()
        short = (message_text[:40]).strip().lower()

        # xpaths to find outgoing message text nodes (try many fallbacks)
        outgoing_xpaths = [
            "//div[contains(@class,'message-out')]//span[contains(@class,'selectable-text') or contains(@class,'copyable-text')]",
            "//div[contains(@data-testid,'msg-container') and contains(@class,'message-out')]//span[contains(@class,'selectable-text')]",
            "//div[@id='main']//div[contains(@class,'message-out')]//span"
        ]

        # xpaths inside a message bubble to detect failure icons / retry buttons
        bubble_icon_xpaths = [
            ".//span[@data-icon='msg-error']",
            ".//svg[@data-icon='msg-error']",
            ".//span[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'not sent')]",
            ".//span[contains(translate(@title,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'not sent')]",
            ".//button[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'retry') or contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'resend')]",
            ".//button[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'retry')]"
        ]

        # global phrases that sometimes appear in page source/popups
        global_failure_phrases = [
            "message not sent", "couldn't send", "could not send", "failed to send",
            "couldn't deliver", "not sent", "try again"
        ]

        while time.time() - start < timeout:
            try:
                page = driver.page_source.lower()
            except WebDriverException:
                page = ""

            # quick global text check
            for ph in global_failure_phrases:
                if ph in page:
                    return False, "popup_text"

            # search for outgoing message elements and try to match the last one
            for xpath in outgoing_xpaths:
                try:
                    elems = driver.find_elements(By.XPATH, xpath)
                except Exception:
                    elems = []

                if not elems:
                    continue

                # take the last outgoing text element
                last = elems[-1]
                try:
                    txt = last.text.strip().lower()
                except Exception:
                    txt = ""

                # if the short snippet exists in the found reply => treat as that outgoing message
                if short and short in txt:
                    # try to locate the surrounding bubble then search within it for failure icons/buttons
                    bubble = None
                    try:
                        bubble = last.find_element(By.XPATH, "ancestor::div[contains(@class,'message-out')]")
                    except Exception:
                        try:
                            bubble = last.find_element(By.XPATH, "ancestor::div[contains(@data-testid,'msg-container')]")
                        except Exception:
                            bubble = last

                    # check for failure icons/buttons inside this bubble
                    for bix in bubble_icon_xpaths:
                        try:
                            icons = bubble.find_elements(By.XPATH, bix)
                        except Exception:
                            icons = []
                        if icons:
                            return False, "icon_found"

                    # no failure icons -> consider sent
                    return True, "msg_found"

            time.sleep(0.4)

        # timed out without confirmation
        return False, "timeout"

    def attempt_resend(self, driver):
        """
        Try to find and click a 'Retry' / 'Resend' button (if it appears)
        Returns True if clicked something, False otherwise.
        """
        retry_xpaths = [
            "//button[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'retry')]",
            "//button[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'resend')]",
            "//div[contains(@role,'alert')]//button",
            "//div//button[contains(translate(@aria-label,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'retry')]"
        ]
        for xp in retry_xpaths:
            try:
                elems = driver.find_elements(By.XPATH, xp)
            except Exception:
                elems = []
            if elems:
                for e in elems:
                    try:
                        driver.execute_script("arguments[0].click();", e)
                        time.sleep(0.8)
                        return True
                    except Exception:
                        continue
        return False

    def send_messages(self):
        file_path = self.file_entry.get().strip()
        message_text = self.message_entry.get("1.0", tk.END).strip()
        country_code = self.cc_entry.get().strip()
        profile_path = self.profile_entry.get().strip()

        if not file_path or not message_text or not country_code:
            messagebox.showerror("Error", "Select Excel file, message, and country code.")
            return

        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Can't read Excel: {e}")
            return

        if 'Phone' not in df.columns:
            messagebox.showerror("Error", "Excel must contain a 'Phone' column.")
            return

        # start browser
        try:
            self.log("Starting Chrome... If not logged in, scan the QR.")
            driver = self.start_driver(profile_path)
            driver.get("https://web.whatsapp.com")
            WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']")))
            self.log("WhatsApp Web is ready.")
        except Exception as e:
            self.log(f"❌ Failed to start WhatsApp Web: {e}")
            return

        success_list = []
        failed_list = []

        for raw in df['Phone']:
            phone_raw = str(raw).strip()
            digits = self.normalize_number(phone_raw, country_code)
            if not digits:
                self.log(f"❌ Invalid phone format: {phone_raw}")
                failed_list.append(phone_raw)
                continue

            url = f"https://web.whatsapp.com/send?phone={digits}&text={urllib.parse.quote(message_text)}"
            self.log(f"Opening chat for +{digits}...")
            try:
                driver.get(url)
            except Exception:
                pass

            # wait for chat open or for invalid-number message
            try:
                WebDriverWait(driver, 20).until(
                    lambda d: ("phone number shared via url is invalid" in d.page_source.lower())
                              or ("not on whatsapp" in d.page_source.lower())
                              or d.find_elements(By.XPATH, "//div[@contenteditable='true']")
                )
            except TimeoutException:
                self.log(f"❌ Timeout opening chat for +{digits}")
                failed_list.append(f"+{digits}")
                continue

            page = driver.page_source.lower()
            if ("phone number shared via url is invalid" in page) or ("not on whatsapp" in page) or ("doesn't have whatsapp" in page) or ("does not have whatsapp" in page):
                self.log(f"❌ Number not on WhatsApp: +{digits}")
                failed_list.append(f"+{digits}")
                continue

            # send the message (try click send button, fallback to Enter)
            try:
                try:
                    send_el = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR,
                            "span[data-icon='send'], button[data-testid='compose-btn-send'], button[aria-label='Send']"))
                    )
                    driver.execute_script("arguments[0].click();", send_el)
                except TimeoutException:
                    inp = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//div[@contenteditable='true']"))
                    )
                    inp.send_keys(Keys.ENTER)
            except Exception as e:
                self.log(f"⚠️ Error while attempting to click send for +{digits}: {e}")

            # check status
            ok, reason = self.check_send_status(driver, message_text, timeout=8)
            if ok:
                self.log(f"✅ Sent to +{digits} (reason: {reason})")
                success_list.append(f"+{digits}")
            else:
                # try one automatic resend if possible (to handle transient network glitch)
                self.log(f"⚠️ Initial send check failed for +{digits} (reason: {reason}). Trying automatic retry if available...")
                tried = False
                try:
                    tried = self.attempt_resend(driver)
                except Exception:
                    tried = False

                if tried:
                    # re-check after retry
                    ok2, reason2 = self.check_send_status(driver, message_text, timeout=6)
                    if ok2:
                        self.log(f"✅ Sent to +{digits} after retry (reason: {reason2})")
                        success_list.append(f"+{digits}")
                    else:
                        self.log(f"❌ Failed to send to +{digits} after retry (reason: {reason2})")
                        failed_list.append(f"+{digits}")
                else:
                    self.log(f"❌ Failed to send to +{digits} (reason: {reason}). No retry button found or retry failed.")
                    failed_list.append(f"+{digits}")

            time.sleep(2)  # polite delay

        # save report
        try:
            with pd.ExcelWriter("whatsapp_message_report.xlsx") as writer:
                pd.DataFrame(success_list, columns=["Successful Numbers"]).to_excel(writer, sheet_name="Success", index=False)
                pd.DataFrame(failed_list, columns=["Failed Numbers"]).to_excel(writer, sheet_name="Failed", index=False)
                pd.DataFrame({"Success Count": [len(success_list)],
                              "Failure Count": [len(failed_list)]}).to_excel(writer, sheet_name="Summary", index=False)
            self.log("✅ Report saved as whatsapp_message_report.xlsx")
            messagebox.showinfo("Done", f"Sent: {len(success_list)} | Failed: {len(failed_list)}")
        except Exception as e:
            self.log(f"❌ Could not save report: {e}")

        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppSeleniumSenderApp(root)
    root.mainloop()
