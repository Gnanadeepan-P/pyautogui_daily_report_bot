from playwright.sync_api import sync_playwright
from datetime import datetime
import os

with sync_playwright() as p:

    # Open Chrome
    browser = p.chromium.launch(
        headless=False
    )

    # Create page
    page = browser.new_page()

    # Open Cricbuzz Live Scores directly
    page.goto(
        "https://www.cricbuzz.com/cricket-match/live-scores",
        wait_until="domcontentloaded"
    )

    # Wait for the page to load
    page.wait_for_timeout(5000)

    print("Page title:", page.title())

    # Display some page content in terminal
    print("\n--- Cricbuzz Scores ---")

    text = page.locator("body").inner_text()

    print(text[:5000])

    # Create screenshots folder
    os.makedirs("screenshots", exist_ok=True)

    # Create today's filename
    today = datetime.now().strftime("%Y-%m-%d")

    screenshot_path = f"screenshots/cricbuzz_{today}.png"

    # Take screenshot
    page.screenshot(
        path=screenshot_path,
        full_page=True
    )

    print("\nScreenshot saved:")
    print(screenshot_path)

    # Keep browser open for 5 seconds
    page.wait_for_timeout(5000)

    # Close browser
    browser.close()