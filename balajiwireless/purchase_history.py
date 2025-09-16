import json
import os
import time
import requests
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from product_page import extract_product_data
from create_html_table import create_html_table

# PRICING_TIERS - Rules to drive List Price from Cost Price 
# If Cost ≤ $2 → 24.99
# If $2 < Cost ≤ $5 → 34.99
# If $5 < Cost ≤ $7.5 → 39.99
# If $7.5 < Cost ≤ $10 → 44.99
# If $10 < Cost ≤ $12 → 49.99
# If $12 < Cost ≤ $14 → 54.99
# If $14 < Cost ≤ $20 → 69.99
# Else (Cost > $20) → (Cost * 4) - 0.01 #rounded to nearest .99

PRICING_TIERS = [
    (0, 2, 24.99),
    (2, 5, 34.99),
    (5, 7.5, 39.99),
    (7.5, 10, 44.99),
    (10, 12, 49.99),
    (12, 14, 54.99),
    (14, 20, 69.99),
]

def cost_to_list(cost_price: float) -> float:
    """
    Convert cost price to list price using tiered rules
    stored in PRICING_TIERS.
    """

    for min_c, max_c, list_price in PRICING_TIERS:
        if min_c < cost_price <= max_c:
            return list_price

    # Fallback: ~4x markup, rounded to nearest .99
    markup = cost_price * 4
    return round(markup) - 0.01

def extract_purchase_history_data():
    """
    Extract purchase history data from website
    and save to an Excel file.
    """
    # === Load login credentials ===
    with open("credentials.json") as f:
        creds = json.load(f)

    EMAIL = creds["email"]
    PASSWORD = creds["password"]
    WEB_URL = creds["web_url"]

    # === Setup Selenium driver ===
    options = Options()
    options.add_argument("--start-maximized")
    #options.add_argument("--headless")  # Add this line for headless mode
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    # === Step 1: Go to website and login ===
    # Step 1: Go to site and click sign in
    driver.get(WEB_URL)

    # Click the "Sign in" link in the header
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.header-profile-login-link"))).click()

    # Wait until login form is fully visible by waiting for the login button
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button.login-register-login-submit")))

    # Fill in email and password using exact IDs
    driver.find_element(By.ID, "login-email").send_keys(EMAIL)
    driver.find_element(By.ID, "login-password").send_keys(PASSWORD)

    # Click the login button
    driver.find_element(By.CSS_SELECTOR, "button.login-register-login-submit").click()


    # === Step 1.5: Ensure on account page ===
    ACCOUNT_URL = "https://www.balajiwireless.com/sca-dev-2019-1/my_account.ssp"
    time.sleep(10)  # Wait for cart to load after login
    if driver.current_url != ACCOUNT_URL:
        driver.get(ACCOUNT_URL)
        time.sleep(5)  # Wait for page to load

    # === Step 2: Navigate to purchase history ===

    # Wait for the "Purchases" side nav item to appear
    purchases_nav_selector = 'a.menu-tree-node-item-anchor[data-id="orders"]'
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, purchases_nav_selector)))

    # Click the "Purchases" side nav item
    driver.find_element(By.CSS_SELECTOR, purchases_nav_selector).click()


    # Wait for the "Purchase History" link to appear after expanding "Purchases"
    purchase_history_selector = 'a.menu-tree-node-item-anchor[data-id="purchases"]'
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, purchase_history_selector)))
    driver.find_element(By.CSS_SELECTOR, purchase_history_selector).click()


    # Wait for the order history table to appear
    order_history_table_selector = 'table.order-history-list-recordviews-actionable-table'
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, order_history_table_selector)))
    time.sleep(2)  # Extra wait to ensure table is fully loaded

    # Wait for the order history table to appear
    order_history_table_selector = 'table.order-history-list-recordviews-actionable-table'
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, order_history_table_selector)))
    order_history_table = driver.find_element(By.CSS_SELECTOR, order_history_table_selector)

    # Find the first row and click its link
    first_row = order_history_table.find_element(By.CSS_SELECTOR, 'tr.recordviews-actionable')
    #print(f"Found first order row: {first_row.get_attribute('outerHTML')}")
    first_link = first_row.find_element(By.TAG_NAME, 'a')
    #print(f"Found link in first order row: {first_link.get_attribute('outerHTML')}")
    first_link.click()
    #print("Clicked on first order row: Navigating to order details page.")
    #print("Clicked on order: Navigating to order details page.")

    # === Step 3: Scrape order details and save to Excel ===
    # Wait for the order detail section to appear
    order_packages_selector = 'div[data-view="OrderPackages"]'
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, order_packages_selector)))
    order_packages = driver.find_element(By.CSS_SELECTOR, order_packages_selector)

    # Find all package dividers (one for each shipping date)
    package_dividers = order_packages.find_elements(By.CSS_SELECTOR, 'div.order-history-packages-acordion-divider')

    all_items = []

    for divider_idx, divider in enumerate(package_dividers):
        # Find the accordion body inside this divider
        accordion_body = divider.find_element(By.CSS_SELECTOR, 'div.order-history-packages-accordion-body')
        # If not expanded, expand it by clicking the header (if needed)
        if "in" not in accordion_body.get_attribute("class"):
            # Find the header that toggles this accordion
            header = divider.find_element(By.CSS_SELECTOR, '[data-toggle="collapse"]')
            driver.execute_script("arguments[0].scrollIntoView(true);", header)
            header.click()
            # Wait for expansion
            wait.until(lambda d: "in" in accordion_body.get_attribute("class"))

        # Now, get all order item rows inside this expanded accordion
        item_rows = accordion_body.find_elements(By.CSS_SELECTOR, 'tr[data-type="order-item"]')
        for row_idx, row in enumerate(item_rows):
            print(f"Processing divider {divider_idx}, row {row_idx}")
            # Product Name and Link
            product_name = ""
            product_page_link = ""
            try:
                # Try to get product name and link from <a>
                name_link = row.find_element(By.CSS_SELECTOR, 'a.transaction-line-views-cell-actionable-name-link')
                product_name = name_link.text.strip()
                product_page_link = name_link.get_attribute("href")
            except Exception:
                try:
                    # Fallback for out-of-stock/no-link items: get from <span>
                    name_span = row.find_element(By.CSS_SELECTOR, 'span.transaction-line-views-cell-actionable-name-viewonly')
                    product_name = name_span.text.strip()
                    product_page_link = ""
                except Exception as e:
                    print(f"Could not find product name in divider {divider_idx}, row {row_idx}: {e}")
                    print("Row HTML:")
                    print(row.get_attribute("outerHTML"))
                    #continue  # Skip this row

            # Price
            price = row.find_element(By.CSS_SELECTOR, 'span.transaction-line-views-price-lead').text.strip()
            # SKU
            sku = row.find_element(By.CSS_SELECTOR, 'span.product-line-sku-value').text.strip()
            # Color
            color = row.find_element(By.CSS_SELECTOR, 'li.transaction-line-views-selected-option-color-text').text.strip()
            # Quantity
            quantity = row.find_element(By.CSS_SELECTOR, 'span.transaction-line-views-quantity-amount-value').text.strip()
            # Total Amount
            total_amount = row.find_element(By.CSS_SELECTOR, 'span.transaction-line-views-quantity-amount-item-amount').text.strip()

            all_items.append({
                "SKU": sku,
                "Product name": product_name,
                "Color": color,
                "Price": price,
                "Quantity": quantity,
                "Total Amount": total_amount,
                "Product Link": product_page_link
            })

    # Save to Excel
    folderpath = "data/purchase_history"
    os.makedirs(folderpath, exist_ok=True)
    output_path = Path(f"{folderpath}/purchase_items.xlsx")
    
    df = pd.DataFrame(all_items)

    # Add List Price column using cost_to_list
    def price_to_float(price_str):
        try:
            return float(price_str.replace("$", "").replace(",", "").strip())
        except Exception:
            return 0.0

    df["List Price"] = df["Price"].apply(price_to_float).apply(cost_to_list)

    df.to_excel(output_path, index=False)
    print(f"Saved {len(all_items)} items to {output_path}")
    driver.close()
    driver.quit()
    return output_path

if __name__ == "__main__":
    filepath = Path("data/purchase_history/purchase_items.xlsx")
    if not os.path.isfile(filepath):
        print("File not found. Extracting purchase history data...")
        filepath = extract_purchase_history_data()
    print(f"purchase_history: Using purchase history file: {filepath}")
    filepath = extract_product_data(filepath)
    print(f"purchase_history: Enriched purchase history data saved to: {filepath}")
    create_html_table(filepath)
    print(f"purchase_history: HTML tables added to purchase history data in: {filepath}")


