````markdown
# 📤 WhatsApp Message Automation using Python

This Python script allows you to send WhatsApp messages to a list of phone numbers loaded from an Excel file (`numbers.xlsx`).  
It logs successful and failed attempts and exports a full report to a new Excel file.

---

## ✅ Features

- 📄 Read phone numbers from an Excel sheet.
- 💬 Send messages using WhatsApp Web via `pywhatkit`.
- 🧾 No need to save contacts on your phone.
- ✅ Track successful and failed messages.
- 📊 Generate a report with:
  - `Success` sheet: Numbers that received messages.
  - `Failed` sheet: Numbers where sending failed.
  - `Summary` sheet: Total sent vs. failed counts.

---

## 🧩 Requirements

Make sure you have the following installed:

```bash
pip install pywhatkit pandas openpyxl
````

* Python 3.x
* WhatsApp Web must be open and logged in on your default browser
* The script uses `pywhatkit` to simulate sending messages

---

## 📄 Input File Format

The script expects an Excel file named `numbers.xlsx` with a column labeled `Phone` containing full international phone numbers.

### Example `nums.xlsx`:

| Phone         |
| ------------- |
| +2013234567890 |
| +2011412223334 |

---

## ▶️ How to Use

1. Open WhatsApp Web in your browser and stay logged in.
2. Place the `numbers.xlsx` file in the same folder as the script.
3. Run the script:

```bash
python send_whatsapp_messages.py
```

4. Wait while the messages are sent one by one.
5. After completion, check the generated file:

```
📄 whatsapp_message_report.xlsx
├── Success  → Successfully sent numbers
├── Failed   → Failed numbers
└── Summary  → Count of success/failure
```

---

## ⚠️ Important Notes

* The script sends messages using your browser, not the WhatsApp API.
* **Delay between messages** is currently set to `20 seconds`. You can adjust it in the script.
* To reduce the chance of getting banned:

  * Don't send more than 50–100 messages per day.
  * Use real, non-spammy content.
  * Avoid sending too fast — increase delay if necessary.

---

## ✅ Example Console Output

```
✅ Message sent to +2012346567890
❌ Failed to send message to +2011125223334: Error details...

📄 Report saved as whatsapp_message_report.xlsx
Total Sent: 1, Total Failed: 1
```

---

## 🔐 Disclaimer

This script is for **educational and personal use only**.
Using it for bulk or promotional messaging may violate WhatsApp's Terms of Service.
Use responsibly.

---


