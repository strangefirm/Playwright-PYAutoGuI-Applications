
import os
import re
import time
from datetime import date
from pathlib import Path

import pyautogui
import pyperclip
import xlwt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

REPORT_HEADERS = ("Date", "Trend", "Search Volume", "Comments")
VOLUME_PATTERN = re.compile(r"^([\d,]+(?:\.\d+)?)\s*([KMB])?\+$", re.IGNORECASE)


def volume_to_number(value):
	match = VOLUME_PATTERN.fullmatch(value.strip())
	if match is None:
		return None

	multiplier = {"K": 1_000, "M": 1_000_000, "B": 1_000_000_000}.get(
		(match.group(2) or "").upper(), 1
	)
	return int(float(match.group(1).replace(",", "")) * multiplier)


def comment_for_volume(volume):
	if volume >= 1_000_000:
		return "Very high search interest"
	if volume >= 100_000:
		return "High search interest"
	if volume >= 10_000:
		return "Moderate search interest"
	return "Lower search interest"


def parse_trends(page_text):
	lines = [line.strip() for line in page_text.splitlines() if line.strip()]
	trends = []
	for index, line in enumerate(lines):
		volume = volume_to_number(line)
		if volume is not None and index > 0:
			trends.append((lines[index - 1], volume))
	return trends


def read_existing_rows(worksheet):
	headers = tuple(cell.value for cell in worksheet[1])
	rows = []
	if headers[:4] == REPORT_HEADERS:
		for row in worksheet.iter_rows(min_row=2, max_col=4, values_only=True):
			if row[1] is not None and row[2] is not None:
				report_row = list(row)
				if not report_row[3]:
					report_row[3] = comment_for_volume(report_row[2])
				rows.append(report_row)
	elif headers[:3] == ("Date", "Copied Data", "Comments"):
		for row_date, page_text, comments in worksheet.iter_rows(
			min_row=2, max_col=3, values_only=True
		):
			if not page_text:
				continue
			for trend, volume in parse_trends(str(page_text)):
				rows.append(
					[row_date, trend, volume, comments or comment_for_volume(volume)]
				)
	elif any(headers):
		raise ValueError("The existing workbook has an unsupported column layout.")
	return rows


def deduplicate_report_rows(rows):
	unique_rows = {}
	for row in rows:
		key = (str(row[0])[:10], " ".join(str(row[1]).split()).casefold())
		unique_rows[key] = row
	return list(unique_rows.values())


def export_xls(path, report_rows):
	workbook = xlwt.Workbook()
	worksheet = workbook.add_sheet("Daily Report")
	header_style = xlwt.easyxf("font: bold on;")
	number_style = xlwt.easyxf(num_format_str="#,##0")

	for column, header in enumerate(REPORT_HEADERS):
		worksheet.write(0, column, header, header_style)

	for row_index, row in enumerate(report_rows, start=1):
		row_date = row[0].strftime("%Y-%m-%d") if hasattr(row[0], "strftime") else str(row[0])[:10]
		worksheet.write(row_index, 0, row_date)
		worksheet.write(row_index, 1, str(row[1]))
		worksheet.write(row_index, 2, row[2], number_style)
		worksheet.write(row_index, 3, str(row[3] or ""))

	worksheet.col(0).width = 14 * 256
	worksheet.col(1).width = 50 * 256
	worksheet.col(2).width = 18 * 256
	worksheet.col(3).width = 40 * 256
	workbook.save(str(path))


def windows_with_title_fragment(fragment):
	fragment = fragment.casefold()
	return [
		window
		for window in pyautogui.getAllWindows()
		if fragment in window.title.casefold()
	]


report_path = Path(__file__).with_name("trends_report.xlsx")
today = date.today()
daily_report_path = Path(__file__).with_name(
	f"daily_report_{today.isoformat()}.xls"
)
excel_lock_path = report_path.with_name(f"~${report_path.name}")
daily_excel_lock_path = daily_report_path.with_name(f"~${daily_report_path.name}")
if excel_lock_path.exists() or daily_excel_lock_path.exists():
	print("Close the report workbook in Excel before running this report.")
	raise SystemExit

print("Daily Report Bot is running!")
time.sleep(0.5)
pyautogui.hotkey("win", "r")
time.sleep(0.5)
pyautogui.typewrite("https://trends.google.com/trending?geo=IN&hl=en-US&sort=search-volume&status=active")

pyautogui.press("enter")
time.sleep(10)
pyperclip.copy("")
pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "c")
time.sleep(0.5)
copied_data = pyperclip.paste().strip()
browser_windows = windows_with_title_fragment("Google Trends")
if browser_windows:
	browser_windows[0].activate()
	time.sleep(0.5)
	active_window = pyautogui.getActiveWindow()
	if active_window and "google trends" in active_window.title.casefold():
		pyautogui.hotkey("ctrl", "w")
		deadline = time.monotonic() + 3
		while time.monotonic() < deadline:
			if not windows_with_title_fragment("Google Trends"):
				break
			time.sleep(0.25)
		if windows_with_title_fragment("Google Trends"):
			print("Google Trends tab is still open; close it manually.")
	else:
		print("Could not focus the Google Trends tab to close it.")
else:
	print("Google Trends browser window was not found to close.")

if not copied_data:
	print("No page text was copied. The report was not saved.")
	raise SystemExit

trends = parse_trends(copied_data)
if not trends:
	print("No trend names and search volumes were found on the page.")
	raise SystemExit

if report_path.exists():
	workbook = load_workbook(report_path)
	worksheet = workbook.active
	existing_rows = read_existing_rows(worksheet)
else:
	workbook = Workbook()
	worksheet = workbook.active
	existing_rows = []

existing_rows.extend(
	[today, trend, volume, comment_for_volume(volume)]
	for trend, volume in trends
)
existing_rows = deduplicate_report_rows(existing_rows)
existing_rows.sort(key=lambda row: row[2], reverse=True)
worksheet.delete_rows(1, worksheet.max_row)
worksheet.append(REPORT_HEADERS)
for report_row in existing_rows:
	worksheet.append(report_row)
	data_row = worksheet.max_row
	worksheet.cell(data_row, 1).number_format = "yyyy-mm-dd"
	worksheet.cell(data_row, 2).alignment = Alignment(wrap_text=True, vertical="top")
worksheet.freeze_panes = "A2"
worksheet.auto_filter.ref = f"A1:D{worksheet.max_row}"
worksheet.column_dimensions["A"].width = 14
worksheet.column_dimensions["B"].width = 50
worksheet.column_dimensions["C"].width = 18
worksheet.column_dimensions["D"].width = 40
workbook.save(report_path)

today_rows = [row for row in existing_rows if str(row[0])[:10] == today.isoformat()]
export_xls(daily_report_path, today_rows)

print(f"Report saved to {daily_report_path.name}. Excel will open it next.")
os.startfile(str(daily_report_path))
window_title = daily_report_path.stem
deadline = time.monotonic() + 15
excel_windows = []
while time.monotonic() < deadline:
	excel_windows = windows_with_title_fragment(window_title)
	if excel_windows:
		break
	time.sleep(0.5)

if excel_windows:
	excel_windows[0].activate()
	time.sleep(0.5)
	active_window = pyautogui.getActiveWindow()
	if active_window and window_title.casefold() in active_window.title.casefold():
		screenshot_path = Path(__file__).with_name(f"daily_report_{today.isoformat()}.png")
		pyautogui.screenshot(str(screenshot_path))
		pyautogui.hotkey("alt", "f4")
		deadline = time.monotonic() + 5
		while time.monotonic() < deadline:
			if not windows_with_title_fragment(window_title):
				break
			time.sleep(0.25)
		if windows_with_title_fragment(window_title):
			print("Excel report window is still open; close it manually.")
	else:
		print("Excel report window could not be focused; screenshot and close were skipped.")
else:
	print("Excel report window was not found; screenshot and close were skipped.")
print(f"Daily report bot finished. Report saved to {daily_report_path.name}.")