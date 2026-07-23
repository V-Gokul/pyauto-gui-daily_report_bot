"""Weather Automation Bot - Complete Real-World Project

This script demonstrates a COMPLETE automation workflow:
1. Open Chrome browser
2. Navigate to AccuWeather website
3. Wait for page to load
4. Take screenshot for visual verification
5. Copy page content to clipboard
6. Parse weather data using regex patterns
7. Extract Hourly Weather forecast
8. Extract 10-Day forecast
9. Write data to Excel workbook
10. Save Excel file

LEARNING CONCEPTS:
- Browser automation (open, navigate, wait)
- Page content extraction (clipboard)
- Regular expressions (regex) for data parsing
- Excel file creation with openpyxl
- Time management (strategic delays for page load)

WARNING: This script is for learning purposes only.
Do not use for commercial data scraping without permission.
"""

import pyautogui
import time
import pyperclip      # Library to access system clipboard
import re             # Regular expressions for pattern matching
from openpyxl import Workbook  # Excel file creation


# ============================================================
# CONFIGURATION & SAFETY
# ============================================================
pyautogui.FAILSAFE = True  # Emergency stop at top-left corner
pyautogui.PAUSE = 0.5      # Wait 0.5 seconds between actions


# ============================================================
# STEP 1: Open Chrome Browser
# ============================================================
print("Step 1: Opening Chrome...")
time.sleep(3)  # Wait 3 seconds (let user prepare)

# Press Windows key to open Start menu
pyautogui.press("win")
time.sleep(1)  # Wait for Start menu to appear

# Type "chrome" to search for Chrome
pyautogui.write("chrome", interval=0.1)
time.sleep(1)  # Wait for search results

# Press Enter to launch Chrome
pyautogui.press("enter")

# Wait for Chrome to fully load (5 seconds)
time.sleep(5)


# ============================================================
# STEP 2: Navigate to AccuWeather
# ============================================================
print("Step 2: Opening the weather page...")

# Ctrl+L focuses on address bar (works in most browsers)
pyautogui.hotkey("ctrl", "l")
time.sleep(1)  # Wait for address bar to be ready

# Type the weather URL (interval=0.03 means 30ms between chars)
pyautogui.write(
    "https://www.accuweather.com/en/in/chennai/206671/weather-forecast/206671",
    interval=0.03,
)

# Press Enter to navigate
pyautogui.press("enter")

# Wait 8 seconds for page to fully load
time.sleep(8)


# ============================================================
# STEP 3: Take Screenshot (for visual verification)
# ============================================================
print("Step 3: Taking screenshot...")

# Screenshot is useful for:
# - Debugging automation issues
# - Verifying page loaded correctly
# - Manual review of extracted data
screenshot = pyautogui.screenshot()  # Capture entire screen
screenshot.save("screenshot.png")    # Save to file

print("Screenshot saved.")


# ============================================================
# STEP 4: Extract Page Content via Clipboard
# ============================================================
print("Step 4: Copying page content...")

# Select all text on page (Ctrl+A)
pyautogui.hotkey("ctrl", "a")
time.sleep(1)  # Wait for selection

# Copy selected text to clipboard (Ctrl+C)
pyautogui.hotkey("ctrl", "c")
time.sleep(2)  # Wait for clipboard operation

# Read the clipboard content into a variable
# This gives us the page text to parse
text = pyperclip.paste()

# Create new Excel workbook
wb = Workbook()


# ============================================================
# STEP 5: Parse Hourly Weather Data
# ============================================================
print("Extracting hourly weather data...")

# Sheet 1: Hourly Weather
ws1 = wb.active
ws1.title = "Hourly Weather"

# Add header row
ws1.append(['Time', 'Temperature', 'Rain Chance'])

# REGEX PATTERN EXPLANATION:
# Find "Hourly Weather" section, extract until "10-Day Weather Forecast"
# re.S = dot matches newlines (multiline mode)
hourly_match = re.search(
    r"Hourly Weather(.*?)10-Day Weather Forecast",
    text,
    re.S,
)

if hourly_match:
    # Extract just the hourly section
    hourly_text = hourly_match.group(1)

    # PATTERN: Extract time, temperature, and rain chance
    # Example matches:
    # 9 PM\n31°\n39%
    # This pattern finds:
    # (\d+\s(?:AM|PM)) = Time like "9 PM"
    # (\d+°) = Temp like "31°"
    # (\d+%) = Rain chance like "39%"
    pattern = r"(\d+\s(?:AM|PM))\s+(\d+°)\s+(\d+%)"

    # Find all matches and add each row to Excel
    for row in re.findall(pattern, hourly_text):
        ws1.append(row)


# ============================================================
# STEP 6: Parse 10-Day Forecast Data
# ============================================================
print("Extracting 10-day forecast data...")

# Sheet 2: 10-Day Forecast
ws2 = wb.create_sheet("10-Day Forecast")

ws2.append([
    "Day",
    "Date",
    "High",
    "Low",
    "Summary",
    "Night Summary",
    "Rain Chance",
])

forecast_match = re.search(
    r"10-Day Weather Forecast(.*?)Sun & Moon",
    text,
    re.S,
)

if forecast_match:
    forecast = forecast_match.group(1)

    pattern = re.compile(
        r"(Tonight|Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+"
        r"(\d+/\d+)\s+"
        r"(\d+)°\s+"
        r"(?:Lo\s+)?(\d+)°?\s+"
        r"(.+?)\s+"
        r"(.+?)\s+"
        r"(\d+)%",
        re.S,
    )

    for item in pattern.findall(forecast):
        ws2.append(item)

wb.save("Weather_Data.xlsx")

print("Weather_Data.xlsx saved successfully.")

