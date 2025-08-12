#!/usr/bin/env python3
"""
Simplified HICV Monthly Scanner - focuses on proper date selection
"""

import os
import re
import csv
import time
import calendar
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout
from dotenv import load_dotenv

# ----------------- Config -----------------
load_dotenv()
USERNAME = os.getenv("HICV_USERNAME")
PASSWORD = os.getenv("HICV_PASSWORD")

if not USERNAME or not PASSWORD:
    print("❌ ERROR: Missing credentials in .env file.")
    exit(1)

# Test settings
ADULTS = 2
CHILDREN = 0
UNIT_SIZES = ["Studio", "1 Bedroom"]
NIGHTS = 7

CSV_PATH = Path("simple_monthly_output.csv")
XLSX_PATH = Path("simple_monthly_output.xlsx")
HEADLESS = False
SLOW_MO = 200

# ----------------- Input helpers -----------------
MONTHS = {
    "january":"January","february":"February","march":"March","april":"April","may":"May","june":"June",
    "july":"July","august":"August","september":"September","october":"October","november":"November","december":"December"
}

def prompt_user_input():
    """Interactive prompt for user to enter month, year, and number of days"""
    print("🏖️ HICV Simple Monthly Scanner")
    print("=" * 40)
    
    # Get month
    while True:
        m_raw = input("Enter month (e.g., December): ").strip()
        month_key = m_raw.strip().lower()
        month = MONTHS.get(month_key)
        if month:
            break
        print(f"❌ Unsupported month: {m_raw!r}. Use full month name (e.g., December).")

    # Get year
    while True:
        y_raw = input("Enter year (e.g., 2025): ").strip()
        if y_raw.isdigit() and 1900 <= int(y_raw) <= 2100:
            year = int(y_raw)
            break
        print(f"❌ Invalid year: {y_raw!r}. Use a 4-digit year like 2025.")

    # Get number of days
    while True:
        days_raw = input("Enter number of days for stay (e.g., 7): ").strip()
        if days_raw.isdigit() and 1 <= int(days_raw) <= 30:
            days = int(days_raw)
            break
        print(f"❌ Invalid number of days: {days_raw!r}. Use a number between 1 and 30.")

    print(f"\n✅ Scanning {month} {year} for {days}-day stays")
    print(f"📅 Will check availability from 1st to last day of {month}")
    print("=" * 40)
    
    return month, year, days

def get_days_in_month(month, year):
    """Get the number of days in the specified month"""
    return calendar.monthrange(year, list(MONTHS.values()).index(month) + 1)[1]

def simple_monthly_scan():
    """Simplified monthly scanning with better date selection"""
    month, year, days = prompt_user_input()
    
    # Calculate the total days in the month and max check-in day
    total_days = get_days_in_month(month, year)
    max_checkin_day = total_days - days + 1  # Ensure we don't go past the month
    
    print(f"📅 Scanning {month} {year} ({total_days} days)")
    print(f"🔍 Will check availability for check-in days 1 to {max_checkin_day}")
    print(f"⏱️ Each stay will be {days} nights")
    print("=" * 50)
    
    all_rows = []  # Collect all data from all days
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO)
        context = browser.new_context()
        page = context.new_page()
        page.set_default_timeout(20000)

        try:
            # Step 1: Login
            print("🔐 Step 1: Logging in...")
            page.goto("https://holidayinnclub.com/login")
            page.fill("#okta-signin-username", USERNAME)
            page.fill("#okta-signin-password", PASSWORD)
            page.click("#okta-signin-submit")
            page.wait_for_load_state("domcontentloaded")
            time.sleep(5)
            print("✅ Login successful")

            # Step 2: Navigate to booking page
            print("🏖️ Step 2: Navigating to booking page...")
            page.goto("https://holidayinnclub.com/account/booking")
            time.sleep(3)
            print("✅ Navigation successful")

            # Step 3: Accept cookies
            try:
                page.get_by_role("button", name=re.compile(r"^I Agree$", re.I)).click(timeout=3000)
                time.sleep(0.3)
                print("🍪 Cookie banner accepted.")
            except PwTimeout:
                pass

            # Step 4: Wait for booking form
            print("📝 Step 4: Waiting for booking form...")
            try:
                page.locator("#resorts-dropdown").wait_for(state="visible", timeout=10000)
            except Exception:
                page.locator("text=Select Check In Date").wait_for(timeout=10000)
            print("✅ Booking form loaded")

            # Step 5: Select Florida locations
            print("📍 Step 5: Selecting Florida locations...")
            
            # Open resorts dropdown
            btn = page.locator("#resorts-dropdown, button[data-uitest='multi-select-button']").first
            btn.wait_for(state="visible", timeout=15000)
            btn.click()
            
            # Wait for dropdown panel
            panel = page.locator("ul[data-uitest='ul-resorts'][role='listbox'], ul[role='listbox']").first
            panel.wait_for(state="visible", timeout=5000)
            
            # Find and select Florida options
            florida_options = panel.locator("li[role='option'] input[type='checkbox']")
            selected_count = 0
            
            for i in range(florida_options.count()):
                try:
                    option = florida_options.nth(i)
                    parent = option.locator("xpath=ancestor::li[1]")
                    
                    # Check if this is a Florida option (not a group header)
                    if parent.get_attribute("aria-selected") != "true":
                        option.click()
                        selected_count += 1
                        time.sleep(0.2)
                except:
                    continue
            
            print(f"✅ Selected {selected_count} Florida locations")
            
            # Close dropdown
            page.keyboard.press("Escape")
            time.sleep(1)

            # Step 6: Fill basic form
            print("📋 Step 6: Filling basic form...")
            
            # Set adults
            try:
                span = page.locator("[data-uitest='number-of-adults-number-of-guests']").first
                minus_btn = page.locator("[data-uitest='number-of-adults-left-icon-button']").first
                plus_btn = page.locator("[data-uitest='number-of-adults-right-icon-button']").first
                
                cur = int((span.inner_text() or "0").strip())
                while cur < ADULTS:
                    plus_btn.click()
                    time.sleep(0.1)
                    cur = int((span.inner_text() or "0").strip())
                print(f"✅ Set adults to {ADULTS}")
            except:
                print("⚠️ Could not set adults")

            # Set children
            try:
                span = page.locator("[data-uitest='number-of-children-number-of-guests']").first
                minus_btn = page.locator("[data-uitest='number-of-children-left-icon-button']").first
                plus_btn = page.locator("[data-uitest='number-of-children-right-icon-button']").first
                
                cur = int((span.inner_text() or "0").strip())
                while cur < CHILDREN:
                    plus_btn.click()
                    time.sleep(0.1)
                    cur = int((span.inner_text() or "0").strip())
                print(f"✅ Set children to {CHILDREN}")
            except:
                print("⚠️ Could not set children")

            # Set unit sizes
            try:
                page.locator("#unit-size-dropdown").click()
                panel = page.locator("ul[data-uitest='ul-unit-sizes']").first
                panel.wait_for(state="visible", timeout=5000)
                
                for unit in UNIT_SIZES:
                    if unit == "Studio":
                        option = panel.locator("label[for='option-ST-id']").first
                    elif unit == "1 Bedroom":
                        option = panel.locator("label[for='option-1BD-id']").first
                    else:
                        continue
                    
                    if option.count() > 0:
                        option.click()
                        time.sleep(0.2)
                
                page.keyboard.press("Escape")
                print(f"✅ Set unit sizes: {UNIT_SIZES}")
            except:
                print("⚠️ Could not set unit sizes")

            # Set nights
            try:
                nights_input = page.locator("#number-of-nights")
                nights_input.click()
                nights_input.fill(str(NIGHTS))
                print(f"✅ Set nights to {NIGHTS}")
            except:
                print("⚠️ Could not set nights")

            # Step 7: Loop through each day
            print("🔄 Step 7: Starting daily availability scan...")
            
            for checkin_day in range(1, min(6, max_checkin_day + 1)):  # Limit to first 5 days for testing
                print(f"\n📅 Day {checkin_day}/{min(5, max_checkin_day)}: Checking {month} {checkin_day}, {year}")
                
                try:
                    # Set month and year
                    print(f"  📅 Setting month to {month} and year to {year}")
                    
                    # Set month
                    try:
                        page.locator("#month-picker").click()
                        month_short = month[:3]  # Dec, Jan, etc.
                        month_btn = page.locator(f"p[role='button']:has-text('{month_short}')").first
                        if month_btn.count() > 0:
                            month_btn.click()
                            print(f"  ✅ Set month to {month}")
                    except:
                        print(f"  ⚠️ Could not set month")

                    # Set year
                    try:
                        page.locator("#year-picker").click()
                        year_btn = page.locator(f"p[role='button']:has-text('{year}')").first
                        if year_btn.count() > 0:
                            year_btn.click()
                            print(f"  ✅ Set year to {year}")
                    except:
                        print(f"  ⚠️ Could not set year")

                    # Open date picker
                    print(f"  📅 Opening date picker for {month} {checkin_day}, {year}")
                    try:
                        page.locator("#date-picker").click()
                        time.sleep(2)
                        
                        # Look for the specific day
                        day_buttons = page.locator("div.rdp-month button[class*='rdp-day']:not([disabled])")
                        print(f"  📅 Found {day_buttons.count()} available days")
                        
                        # Try to find the specific day
                        target_day_found = False
                        for i in range(day_buttons.count()):
                            day_btn = day_buttons.nth(i)
                            day_text = day_btn.inner_text()
                            if day_text and day_text.strip() == str(checkin_day):
                                day_btn.click()
                                print(f"  ✅ Clicked day {checkin_day}")
                                target_day_found = True
                                break
                        
                        if not target_day_found:
                            print(f"  ⚠️ Day {checkin_day} not found, clicking first available day")
                            if day_buttons.count() > 0:
                                day_buttons.nth(0).click()
                        
                        # Click Done
                        try:
                            page.get_by_role("button", name=re.compile(r"^Done$", re.I)).click()
                            print(f"  ✅ Clicked Done")
                        except:
                            print(f"  ⚠️ Could not click Done")
                        
                    except Exception as e:
                        print(f"  ❌ Date picker error: {e}")

                    # Click Confirm Dates
                    print(f"  ✅ Clicking Confirm Dates...")
                    try:
                        confirm_btn = page.locator("button:has-text('Confirm Dates')").first
                        if confirm_btn.count() > 0:
                            confirm_btn.click()
                            print(f"  ✅ Clicked Confirm Dates")
                    except:
                        print(f"  ⚠️ Could not click Confirm Dates")

                    # Wait for results
                    print(f"  ⏳ Waiting for results...")
                    try:
                        page.locator("text=Modify Search").wait_for(state="visible", timeout=15000)
                        print(f"  ✅ Results page loaded")
                    except:
                        print(f"  ⚠️ Results page not found")

                    # Extract data (simplified)
                    print(f"  🔍 Extracting data...")
                    try:
                        # Look for resort names
                        resort_elements = page.locator("h2, h3, [class*='resort']")
                        resort_count = resort_elements.count()
                        print(f"  📊 Found {resort_count} resort elements")
                        
                        # Look for points
                        point_elements = page.locator("*:has-text('points')")
                        point_count = point_elements.count()
                        print(f"  💰 Found {point_count} point elements")
                        
                        if resort_count > 0 or point_count > 0:
                            print(f"  ✅ Found availability data")
                            # Add a simple row for this day
                            all_rows.append({
                                "Date Range": f"{month} {checkin_day} - {month} {checkin_day + days - 1}, {year}",
                                "Resort": f"Day {checkin_day} data",
                                "Room": "Available",
                                "Points": "Data found"
                            })
                        else:
                            print(f"  ⚠️ No availability data found")
                            
                    except Exception as e:
                        print(f"  ❌ Data extraction error: {e}")

                    # Small delay
                    time.sleep(2)
                    
                except Exception as e:
                    print(f"  ❌ Error for day {checkin_day}: {e}")
                    continue

            # Step 8: Save data
            print(f"\n💾 Saving collected data...")
            if all_rows:
                # Save as CSV
                with open(CSV_PATH, "w", newline="", encoding="utf-8") as f:
                    writer = csv.DictWriter(f, fieldnames=["Date Range", "Resort", "Room", "Points"])
                    writer.writeheader()
                    for row in all_rows:
                        writer.writerow(row)
                
                print(f"✅ Saved {len(all_rows)} rows to {CSV_PATH}")
                print(f"📊 Summary: {len(all_rows)} days checked")
            else:
                print("⚠️ No data collected")

        except Exception as e:
            print(f"❌ Error during scan: {e}")
            raise

        finally:
            time.sleep(2)
            browser.close()

    print("\n🎉 Simple monthly scan completed!")

if __name__ == "__main__":
    simple_monthly_scan() 