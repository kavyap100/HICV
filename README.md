# HICV
Holiday Inn Club Vacations (HICV) Availability Automation is a Python + Playwright project that signs into the HICV member portal, searches availability by destination/date/unit size, and exports structured results to CSV/XLSX. It’s built for repeatable checks, quick analysis, and easy integration into automations (e.g., Make.com or n8n).

## Features
Headless Playwright flow with robust selectors and retry logic

Secure login via .env (HICV_USERNAME, HICV_PASSWORD)

Configurable filters: destination/state/resort, unit sizes, guests, date range, nights

Exports results to CSV/XLSX (filename configurable)

Saves screenshots and raw HTML for troubleshooting

Designed to plug into external automations/pipelines

## Quick start
1.cp .env.example .env and set HICV_USERNAME/HICV_PASSWORD

2.pip install -r requirements.txt and playwright install

3.Adjust config at the top of the script (dates, unit sizes, headless, slow_mo, output paths)

4.python hicv.py → Produces member_availability.csv and member_availability.xlsx (or your configured names)

## Output
A structured table of matches (e.g., resort, location, unit type, check-in, check-out, nights, availability/price if present), ready for spreadsheets or downstream tooling.
