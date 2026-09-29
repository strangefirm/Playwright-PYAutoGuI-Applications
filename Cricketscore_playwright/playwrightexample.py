# Import Playwright's synchronous API for controlling the browser step by step.
from playwright.sync_api import sync_playwright
# Import date and time utilities for the log and spreadsheet date column.
from datetime import datetime
# Import Path for building output-file paths beside this script.
from pathlib import Path
# Import regular expressions to recognize cricket score lines.
import re
# Import the Excel workbook creator.
from openpyxl import Workbook
# Import cell formatting tools for readable spreadsheet text and headers.
from openpyxl.styles import Alignment, Font, PatternFill
# Import Excel table tools for filters and alternating row colors.
from openpyxl.worksheet.table import Table, TableStyleInfo

# Print a short message so the user knows the automation has started.
print("starting playwright automation script")
# Print the local start time for a simple run log.
print("Starting the script at:", datetime.now())


# Define a helper that converts Cricbuzz's visible text into Excel-ready rows.
def extract_match_rows(page_text, match_date):
	# Split the page into trimmed lines and discard blank lines.
	lines = [line.strip() for line in page_text.splitlines() if line.strip()]
	try:
		# Find where the live-score section begins.
		section_start = lines.index("Live Cricket Score") + 1
	except ValueError:
		# Return no matches if Cricbuzz changes the page and this heading is missing.
		return []

	try:
		# Stop before videos so unrelated page content is not treated as match data.
		section_end = lines.index("FEATURED VIDEOS", section_start)
	except ValueError:
		# If there is no video heading, use the rest of the page.
		section_end = len(lines)

	# Keep only the lines that belong to the live-score section.
	match_lines = lines[section_start:section_end]
	# Find bullet-separated match-detail lines; each starts a new match card.
	detail_indices = [index for index, line in enumerate(match_lines) if "\u2022" in line]
	# Match common score formats, including wickets, overs, and multiple innings.
	score_pattern = re.compile(r"^\d+(?:[-/]\d+)?(?:\s*&\s*\d+(?:[-/]\d+)?)?(?:\s+\(\d+(?:\.\d+)?\))?$")
	# Store each match as a three-cell row: match, score, and capture date.
	rows = []

	# Process each match card, using the next card's detail line as its boundary.
	for position, detail_index in enumerate(detail_indices):
		# Find the end of this card's text block.
		segment_end = detail_indices[position + 1] if position + 1 < len(detail_indices) else len(match_lines)
		# Get the lines after this card's heading and before the next card.
		segment = match_lines[detail_index + 1:segment_end]
		# Record the line numbers that look like team scores.
		score_indices = [index for index, line in enumerate(segment) if score_pattern.fullmatch(line)]
		# Skip cards that do not contain a score.
		if not score_indices:
			continue

		# The first two non-score lines in a card are its team names.
		team_positions = [
			(index, line)
			for index, line in enumerate(segment)
			if not score_pattern.fullmatch(line)
		][:2]
		if len(team_positions) < 2:
			continue
		team_names = [line for _, line in team_positions]

		# A score belongs to the closest team name before it; the other side may not have batted yet.
		scores_by_team = {}
		for score_index in score_indices:
			previous_teams = [
				line for index, line in team_positions if index < score_index
			]
			if previous_teams:
				scores_by_team[previous_teams[-1]] = segment[score_index]

		# Take the match type (such as "1st ODI") from before the bullet separator.
		match_type = match_lines[detail_index].split("\u2022", 1)[0].strip()
		# The line before the match type names its tournament or series.
		series = match_lines[detail_index - 1] if detail_index > 0 else "Cricbuzz match"
		# Take the first line after the scores and team names as the result, if present.
		result = next(
			(
				line for line in segment[max(score_indices) + 1:]
				if line not in team_names
				and line not in {"Live Score", "|", "Scorecard", "Full Commentary", "News"}
			),
			"",
		)
		# Include a placeholder when the other team has not posted a score.
		score = "; ".join(
			f"{team}: {scores_by_team.get(team, 'No score shown')}" for team in team_names
		)
		# Add the result to the score cell when Cricbuzz provides one.
		if result:
			score = f"{score}; {result}"
		# Add one row; the date is when this page was captured.
		rows.append([f"{team_names[0]} vs {team_names[1]} - {series} ({match_type})", score, match_date])

	# Give the completed list of spreadsheet rows back to the caller.
	return rows


# Start Playwright and keep its browser-control session open for this block.
with sync_playwright() as p:
	# Launch the installed Google Chrome visibly rather than using headless mode.
	browser = p.chromium.launch(channel="chrome", headless=False)
	try:
		# Open a fresh browser tab for Cricbuzz.
		page = browser.new_page()
		# Navigate to the live scores page and wait for its initial HTML to load.
		page.goto("https://www.cricbuzz.com/cricket-match/live-scores", wait_until="domcontentloaded", timeout=60000)
		# Wait until the page body exists before reading its text.
		page.locator("body").wait_for()

		# Read the text currently visible on the page.
		page_text = page.locator("body").inner_text()
		# Use the script's folder as the destination for all generated files.
		output_dir = Path(__file__).resolve().parent
		# Set the output path for the page screenshot.
		screenshot_path = output_dir / "cricbuzz_screenshot.png"
		# Set the output path for the formatted Excel workbook.
		excel_path = output_dir / "cricbuzz_matches.xlsx"
		# Get the date used in the spreadsheet's Date column.
		capture_time = datetime.now().astimezone()

		# Capture the whole page and save it as an image.
		page.screenshot(path=str(screenshot_path), full_page=True)
		# Parse the match cards into rows and use the capture date as the date value.
		match_rows = extract_match_rows(page_text, capture_time.date())
		# Stop with a clear error rather than creating an empty workbook.
		if not match_rows:
			page_text_lower = page_text.lower()
			if "access denied" in page_text_lower or "don't have permission to access" in page_text_lower:
				raise RuntimeError(
					f"Cricbuzz denied this browser request, so no scores were available. "
					f"Check the saved screenshot: {screenshot_path}"
				)
			raise RuntimeError("No match score rows were found in Cricbuzz's visible page text.")

		# Create a new Excel workbook and use its first worksheet.
		workbook = Workbook()
		worksheet = workbook.active
		# Name the worksheet shown at the bottom of Excel.
		worksheet.title = "Match Scores"
		# Add the requested column headings as the first row.
		worksheet.append(["Match", "Score", "Date"])
		# Add one worksheet row for every extracted match.
		for row in match_rows:
			worksheet.append(row)

		# Apply a dark green fill, white bold text, and vertical alignment to the headings.
		for cell in worksheet[1]:
			cell.fill = PatternFill(fill_type="solid", fgColor="145A46")
			cell.font = Font(color="FFFFFF", bold=True)
			cell.alignment = Alignment(vertical="center")
		# Wrap long match and score text so it stays inside its cell.
		for row in worksheet.iter_rows(min_row=2):
			for cell in row:
				cell.alignment = Alignment(vertical="top", wrap_text=True)
			# Display the date in a human-readable day-month-year format.
			row[2].number_format = "dd mmm yyyy"
		# Set column widths so the match and score text is easier to read.
		worksheet.column_dimensions["A"].width = 56
		worksheet.column_dimensions["B"].width = 58
		worksheet.column_dimensions["C"].width = 16
		# Keep the heading visible while scrolling through match rows.
		worksheet.freeze_panes = "A2"
		# Add filter dropdowns to the table's headings.
		worksheet.auto_filter.ref = worksheet.dimensions
		# Create an Excel table covering the headings and all match rows.
		table = Table(displayName="CricbuzzMatches", ref=worksheet.dimensions)
		# Apply a built-in table style with alternating row colors.
		table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium4", showRowStripes=True, showColumnStripes=False)
		# Attach the styled table to the worksheet.
		worksheet.add_table(table)
		# Write the completed workbook to the Excel output file.
		workbook.save(excel_path)

		# Print the locations of the Excel workbook and screenshot.
		print(f"Screenshot saved to: {screenshot_path}")
		print(f"Excel file saved to: {excel_path} ({len(match_rows)} matches)")
	finally:
		# Close Chrome even if navigation, extraction, or file saving fails.
		browser.close()

    

