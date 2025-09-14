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
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, purchases_nav_selector))).click()

# Wait for the "Purchase History" link to appear after expanding "Purchases"
purchase_history_selector = 'a.menu-tree-node-item-anchor[data-id="purchases"]'
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, purchase_history_selector))).click()

# Wait for the order history table to appear
order_history_table_selector = 'table.order-history-list-recordviews-actionable-table'
wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, order_history_table_selector)))
time.sleep(5)  # Extra wait to ensure table is fully loaded

# Click on the specific order number link after the table appears
order_no_selector = 'a[href="#/purchases/view/salesorder/54455818"]'
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, order_no_selector))).click()

# === Step 3: Scrape order details and save to Excel ===

# Wait for the order detail section to appear
order_packages_selector = 'div[data-view="OrderPackages"]'
wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, order_packages_selector)))
order_packages = driver.find_element(By.CSS_SELECTOR, order_packages_selector)

# Find all package dividers (one for each shipping date)
package_dividers = order_packages.find_elements(By.CSS_SELECTOR, 'div.order-history-packages-acordion-divider')

all_items = []

for divider in package_dividers:
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
    for row in item_rows:
        # Product name and link
        name_link = row.find_element(By.CSS_SELECTOR, 'a.transaction-line-views-cell-actionable-name-link')
        product_name = name_link.text.strip()
        product_page_link = name_link.get_attribute("href")

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
df = pd.DataFrame(all_items)
output_path = Path("order_items.xlsx")
df.to_excel(output_path, index=False)
print(f"Saved {len(all_items)} items to {output_path}")


