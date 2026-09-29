# SE-AI-B12-2-W1-D3

Week 1 – Day 3 assignment: three Python automation and web scraping mini-projects, built with **Playwright**, **PyAutoGUI**, and **Excel reporting**.
## 📁 Repository Structure

```
SE-AI-B12-2-W1-D3/
│
├── Cricketscore_playwright/
│   ├── .gitignore
│   ├── cricbuzz_matches.xlsx
│   ├── cricbuzz_screenshot.png
│   └── playwrightexample.py
│
├── Playwright_assignment/
│   ├── whatsapp_screenshots/
│   ├── .gitattributes
│   ├── .gitignore
│   ├── contacts.xlsx
│   ├── playwright_assign.py
│   ├── poster.jpeg
│   ├── whatsapp_report_2026-09-29.json
│   └── whatsapp_report_2026-09-29.xlsx
│
├── pyautogui_webscraping/
│   ├── daily_report_2026-09-30.png
│   ├── daily_report_2026-09-30.xls
│   └── daily_report_bot.py
│
├── README.md
└── requirements.txt
```

---

## 📌 1. Cricket Score Scraping — `Cricketscore_playwright/`

**Goal:** Scrape live cricket scores from Cricbuzz and save them to Excel, with a screenshot of the page as proof.

**Script:** `playwrightexample.py`
**Tech stack:** `playwright`, `openpyxl`

**What it does:**
- Launches a browser and navigates to Cricbuzz
- Scrapes current match scores from the page
- Saves the parsed data into `cricbuzz_matches.xlsx`
- Captures a full-page screenshot as `cricbuzz_screenshot.png`

**How to run:**
```bash
cd Cricketscore_playwright
pip install -r ../requirements.txt
playwright install chromium
python playwrightexample.py
```

---

## 📌 2. WhatsApp Automation — `Playwright_assignment/`

**Goal:** Send personalized WhatsApp messages (text and image posters) to a list of consented contacts, and generate a delivery report. Numbers masked for security reasons.

**Script:** `playwright_assign.py`
**Tech stack:** `playwright`, `openpyxl`

**What it does:**
- Reads contact details (name, phone, message, poster) from `contacts.xlsx`
- Logs into WhatsApp Web via a persistent browser session (QR scan once)
- Sends a personalized text message and/or the `poster.jpeg` image to each contact
- Saves a screenshot of each conversation to `whatsapp_screenshots/`
- Reads back the last 3 messages in each chat for confirmation
- Saves a full delivery report as `whatsapp_report_<date>.json` and `whatsapp_report_<date>.xlsx`

**How to run:**
```bash
cd Playwright_assignment
pip install -r ../requirements.txt
playwright install chromium
python playwright_assign.py
```

**⚠️ Responsible use:** This tool only messages contacts who have consented to receive messages. Built-in random delays are used between actions to avoid spam-like behavior.

---

## 📌 3. Daily Report Automation — `pyautogui_webscraping/`

**Goal:** Automate an on-screen/browser workflow using PyAutoGUI to capture a daily report and save it as an Excel file and screenshot.

**Script:** `daily_report_bot.py`
**Tech stack:** `pyautogui`, `openpyxl`

**What it does:**
- Uses PyAutoGUI to interact with the screen and gather the daily report data
- Parses and structures the extracted data
- Saves the results into `daily_report_<date>.xls`
- Saves a screenshot of the report as `daily_report_<date>.png`

**How to run:**
```bash
cd pyautogui_webscraping
pip install -r ../requirements.txt
python daily_report_bot.py
```

**Note:** Keep the target window in the foreground and avoid moving the mouse/keyboard while the script runs, since PyAutoGUI automates the screen directly.

---

## ⚙️ General Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/strangefirm/SE-AI-B12-2-W1-D3.git
   cd SE-AI-B12-2-W1-D3
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv venv
   venv\Scripts\activate        # Windows
   source venv/bin/activate     # macOS/Linux
   ```

3. Install all dependencies:
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

4. Run any project from its respective folder as shown above.

---

## 🛠️ Tools & Libraries Used

| Library | Purpose |
|---|---|
| `playwright` | Browser automation for Cricbuzz scraping and WhatsApp Web |
| `pyautogui` | Screen-level automation for the daily report bot |
| `openpyxl` | Reading and writing Excel (.xlsx / .xls) reports |

---

## 👤 Author

**Fathima Suhaila**
Week 1 – Day 3 Assignment | Python Automation & Web Scraping

---

## 📄 License

This project is for educational purposes as part of a training assignment.
