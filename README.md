# WhatsApp Automated Sender (GUI)

A Python GUI tool to send personalized WhatsApp messages from an Excel sheet using Selenium automation.  
Supports selecting a sheet, entering your message, and tracking send status in a log file.

---

## Features

- Select Excel file and sheet
- Automatically detect phone numbers
- Enter and customize your message
- Sends messages via WhatsApp Web
- Logs both successful and failed sends
- No manual ChromeDriver setup — works automatically

---

## Requirements

- Python 3.9 or newer
- Google Chrome installed and updated with  WhatsApp account logged in
- WhatsApp account

---

## Installation

in your command line

Install the latest versions of required Python packages:
````markdown

pip install pandas openpyxl selenium webdriver-manager
````
or using requirements.txt file:

````markdown

pip install -r requirements.txt
````
---

## Usage

1. **Run the script:**

   ```bash
   python main.py
   ```

2. **Steps in the GUI:**

   * Select your Excel file
   * select country code
   * Enter your message
   * Click **Send Messages**

3. **WhatsApp Web** will open in Chrome:

   * Scan the QR code (first time only if you are not logged-in)
   * Messages will be sent automatically
   * A status log will be saved in `whatsapp_message_report.xlsx`

---

## Excel Format

Your Excel file should contain:

* **Column Name:** `Phone`
* Example:

  | Phone        |
  | ------------ |
  | 202357890.. |
  | 211223334.. |

---

## Troubleshooting

* **Chrome update issue:**
  This script uses `webdriver-manager` which auto-updates the driver to match Chrome’s latest version.

* **Numbers without WhatsApp:**
  Such numbers will be marked as failed in the log.

* **Script not starting:**
  Make sure Python and Chrome are both installed and added to your system PATH.

---

## 🔐 Disclaimer

This script is for educational and personal use only. Automated WhatsApp messaging through unofficial means may violate WhatsApp's terms of service. Use responsibly.

