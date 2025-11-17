📌 WhatsApp Excel Automation Bot

This project automates sending messages on WhatsApp Web using Selenium, pulling data from an Excel sheet and messaging each person based on the calculated values in the sheet.

The script is built to:

✔ Load Excel data
✔ Open WhatsApp Web through Microsoft Edge
✔ Search for each contact
✔ Paste a pre-generated message
✔ Send messages one-by-one with human-like timing
✔ Handle both positive and negative balances
✔ Skip incomplete rows safely

🚀 Features

💬 Automatic WhatsApp messaging

📄 Reads data directly from Excel (.xlsx)

🔍 Smart search system to find each contact

🤖 Human-like delays to avoid detection

🛡️ Exception handling (missing names, failed searches, invalid numbers)

🔄 Message templates for positive & negative balances

🔧 Configurable Excel path, start row, and Edge driver path

📂 Project Structure
/whatsapp-accounting-sender
│── README.md
│── main.py   # Your automation script
│── requirements.txt

🧩 Requirements

Install dependencies:

pip install pandas selenium pyperclip


You must also:

Install Microsoft Edge

Download the correct msedgedriver.exe version

Update the path inside the script:

edgedriver_path = r"C:\Users\Cyber\Desktop\محاسبة\edgedriver_win64\msedgedriver.exe"

📝 Configuration

Edit these values at the top of the script:

excel_path = r"C:\Users\Cyber\Downloads\29-4_9 (1).xlsx"
start_row = 1
edgedriver_path = r"C:\Users\Cyber\Desktop\محاسبة\edgedriver_win64\msedgedriver.exe"

▶️ How to Run

Open terminal

Navigate to your folder

Run:

python main.py


WhatsApp Web will open

If first run → scan QR

Script will start sending messages from Excel automatically

📦 requirements.txt (paste into your repo)
pandas
selenium
pyperclip

⚠️ Important Notes

WhatsApp Web UI changes sometimes → XPaths may need updating

Keep your Edge browser version matching the driver

Use this responsibly — do not spam

📄 License

MIT License — free to use and modify.
