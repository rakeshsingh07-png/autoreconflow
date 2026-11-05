import sys
from playwright.sync_api import sync_playwright

APP_URL = "https://autoreconflow.streamlit.app"

def wake_and_ping_app():
    print(f"🤖 [AI Worker] Launching browser to ping {APP_URL}...")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()

        try:
            # Navigate to the app with a generous timeout
            page.goto(APP_URL, wait_until="networkidle", timeout=60000)
            print("🌐 Page loaded.")

            # Check if the sleeping button exists
            # Streamlit sleeping button typically contains "Yes, get this app back up!"
            wake_button = page.locator("button:has-text('Yes, get this app back up!')")
            
            if wake_button.count() > 0 and wake_button.is_visible():
                print("⚠️ App was sleeping! Clicking 'Yes, get this app back up!'...")
                wake_button.click()
                # Wait 20 seconds for the container to reload
                page.wait_for_timeout(20000)
                print("✅ Wake-up button triggered successfully.")
            else:
                print("✅ App is already awake and active!")

            # Take a small delay to ensure Streamlit WebSocket session registers
            page.wait_for_timeout(5000)

        except Exception as e:
            print(f"❌ Error during worker run: {e}")
            browser.close()
            sys.exit(1)

        browser.close()
        print("🎯 Worker completed run successfully.")

if __name__ == "__main__":
    wake_and_ping_app()
