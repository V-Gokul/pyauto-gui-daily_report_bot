# PyAutoGUI Daily Weather Report Bot

A desktop automation bot that opens Chrome, pulls the live weather forecast for Chennai from AccuWeather, and saves it into a formatted Excel report — no APIs, just GUI automation, clipboard scraping, and regex.

<img width="1914" height="753" alt="image" src="https://github.com/user-attachments/assets/0b54f2fd-21ec-46d9-b8e3-6fd755c2ffb9" />

> ⚠️ **Learning project only.** This automates keyboard/mouse actions on a real browser window — it is not a scraping API and should not be used for commercial data collection without permission from the site.

## What it does

Running the script performs a full end-to-end automation workflow:

1. **Opens Chrome** via the Windows Start menu.
2. **Navigates** to the AccuWeather forecast page for Chennai.
3. **Takes a screenshot** of the loaded page for visual verification (`screenshot.png`).
4. **Selects and copies** the entire page text to the clipboard (`Ctrl+A`, `Ctrl+C`).
5. **Parses the clipboard text** with regex to extract:
   - Hourly forecast (time, temperature, rain chance)
   - 10-day forecast (day, date, high/low, summary, night summary, rain chance)
6. **Writes the results to Excel**, saving two sheets — `Hourly Weather` and `10-Day Forecast` — into `Weather_Data.xlsx`.

## How it works

The bot doesn't use any weather API — it drives the real UI:

| Step | Technique |
|------|-----------|
| Launch Chrome & navigate | `pyautogui` simulates key presses (`win`, typing, `Ctrl+L`, `Enter`) |
| Timing | Fixed `time.sleep()` delays give the OS/browser/page time to load |
| Extract page text | `Ctrl+A` → `Ctrl+C` copies visible text, then `pyperclip.paste()` reads the clipboard |
| Parse weather data | `re` (regex) patterns match time/temperature/forecast text blocks |
| Save report | `openpyxl` writes the parsed rows into a multi-sheet `.xlsx` workbook |

`pyautogui.FAILSAFE` is enabled, so moving the mouse to the top-left corner of the screen at any time immediately aborts the script.

## Packages to install

Requires **Python 3.8+**. Install the dependencies with pip:

```bash
pip install pyautogui pyperclip openpyxl
```

| Package | Purpose |
|---------|---------|
| [`pyautogui`](https://pypi.org/project/PyAutoGUI/) | Simulates keyboard/mouse input to control Chrome |
| [`pyperclip`](https://pypi.org/project/pyperclip/) | Reads the system clipboard |
| [`openpyxl`](https://pypi.org/project/openpyxl/) | Creates and writes the `.xlsx` Excel report |

## Usage

1. Make sure **Google Chrome** is installed and available from the Start menu.
2. Install the dependencies above (a `venv` is recommended).
3. Run the script:
   ```bash
   python daily_report_bot.py
   ```
4. Don't touch the mouse/keyboard while it runs — it takes over control of Chrome for several seconds. There's a 3-second countdown at the start to let you switch away from the terminal.
5. When it finishes, check the project folder for:
   - `screenshot.png` — screenshot of the loaded weather page
   - `Weather_Data.xlsx` — the extracted hourly and 10-day forecast data

## Notes & limitations

- Uses fixed `time.sleep()` delays, so it's sensitive to slow internet or page load times.
- The regex patterns are tailored to AccuWeather's current page layout and may break if the site changes.
- The target city/URL is hardcoded (`Chennai`) inside [daily_report_bot.py](daily_report_bot.py) — edit the URL in **Step 2** to target a different location.
