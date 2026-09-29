# Week 1 – Day 3 Assignment

This repository contains three automation and web scraping mini-projects built as part of the Week 1, Day 3 assignment.
Each project automates a real-world, repetitive task using Python, and stores the results in an easy-to-review format (Excel and/or screenshots).

---

## 📁 Project Structure

```
week1-day3-assignment/
│
├── 1_cricket_score_scraper/
│   ├── cricket_score.py
│   └── cricket_scores.xlsx        (generated after running)
│
├── 2_whatsapp_automation/
│   ├── whatsapp_sender.py
│   ├── contacts.xlsx              (input file you provide)
│   └── whatsapp_screenshots/      (generated after running)
│
├── 3_trending_topics_scraper/
│   ├── trending_scraper.py
│   ├── trending_topics.xlsx       (generated after running)
│   └── trending_screenshot.png    (generated after running)
│
├── requirements.txt
└── README.md
```

---

## 🧰 Prerequisites

- Python 3.9 or above
- Google Chrome installed
- A virtual environment (recommended)

Set up the environment:

```bash
python -m venv venv
venv\Scripts\activate        # Windows
source venv/bin/activate     # macOS/Linux

pip install -r requirements.txt
playwright install chromium
```

---

## 1️⃣ Cricket Score Web Scraping (PyAutoGUI)

**Goal:** Fetch live/recent cricket scores from a public cricket website and save them into an Excel sheet, using PyAutoGUI to automate on-screen browser interactions.

**How it works:**
- Opens the target cricket website in the default browser.
- Uses `pyautogui` to simulate mouse clicks and keyboard actions to navigate to the scores page.
- Extracts match details (teams, score, overs, status).
- Saves the structured data into `cricket_scores.xlsx` using `openpyxl`.

**Run:**
```bash
cd 1_cricket_score_scraper
python cricket_score.py
```

**Note:** Since PyAutoGUI controls the real mouse and keyboard, keep the browser window active and avoid touching the mouse/keyboard while the script runs.

**Output:** `cricket_scores.xlsx` containing match-wise score data.

---

## 2️⃣ WhatsApp Automation (Playwright)

**Goal:** Send personalized WhatsApp messages (with an optional poster image) to a list of consenting contacts, and log delivery status automatically.

**How it works:**
- Reads contacts (Name, Phone, Message, Poster) from `contacts.xlsx`.
- Launches WhatsApp Web via Playwright with a persistent browser profile, so the QR code only needs to be scanned once.
- Sends a personalized text message, or an image with a caption, to each contact.
- Takes a screenshot of each conversation as proof of delivery.
- Reads the last 3 messages in the chat to confirm the conversation state.
- Saves a full report (`whatsapp_report_<date>.xlsx` and `.json`) with status, screenshot path, and any errors per contact.

**Run:**
```bash
cd 2_whatsapp_automation
python whatsapp_sender.py
```

**Note:** Only message contacts who have explicitly agreed to receive messages. Scan the WhatsApp Web QR code when prompted, then press Enter in the terminal to begin sending.

**Output:** `whatsapp_screenshots/`, `whatsapp_report_<date>.xlsx`, `whatsapp_report_<date>.json`

---

## 3️⃣ Trending Topics Scraper (Playwright)

**Goal:** Scrape currently trending topics/keywords from a public trends source, parse and sort the data, store it in Excel, and capture a screenshot of the source page as evidence.

**How it works:**
- Launches a headless/headed browser using Playwright and navigates to the trends page.
- Parses the trend names, ranks, and related metrics (e.g., tweet/search volume) from the page content.
- Sorts the parsed data (e.g., by rank or volume).
- Writes the sorted data into `trending_topics.xlsx` with formatted headers.
- Captures a full-page screenshot of the trends page as `trending_screenshot.png`.

**Run:**
```bash
cd 3_trending_topics_scraper
python trending_scraper.py
```

**Output:** `trending_topics.xlsx`, `trending_screenshot.png`

---

## 📦 requirements.txt

```
pyautogui
playwright
openpyxl
pandas
```

---

## ⚠️ Disclaimer

These scripts are built for **educational purposes** as part of a training assignment. Web scraping and automation should always comply with the target website's Terms of Service and applicable data/privacy laws (e.g., only messaging contacts who have given consent). Selectors used for scraping may need updates if the source websites change their page structure.

---

## 👤 Author

Fathima
Week 1 – Day 3 Assignment
