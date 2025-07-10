import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import requests
import os
from pathlib import Path

# Read Excel file with product links and headings
df = pd.read_excel('amazon.xlsx')
data = df.to_dict(orient='records')
headings = df.columns.tolist()

options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

results = []

for index, item in enumerate(data):
    product_link = item.get('Product Link') or item.get('Link')  # Adjust column name as needed
    if not product_link or str(product_link).strip() in ["", "nan"]:
        print(f"Skipping index {index} due to missing link.")
        continue

    #print(f"Processing: {product_link}")
    driver.get(product_link)
    time.sleep(3)  # Let page load

    result = {
        "Title": "",
        "Price": "",
        "Description": "",
        "Product Link": product_link
    }

    # --- Extract Title ---
    try:
        #time.sleep(500)  # Let page load
        wait.until(EC.presence_of_element_located((By.ID, "productTitle")))
        title = driver.find_element(By.ID, "productTitle").text.strip()
        print(f"Extracted Title: {title}")
        result['Title'] = title
    except:
        print("Title not found, trying alternative selectors.")
        result['Title'] = ""

    # --- Extract Price ---
    try:
        price_elements = driver.find_elements(By.CSS_SELECTOR, "span.a-offscreen")
        price = ""
        for elem in price_elements:
            print(f"Checking price element: {elem.get_attribute('innerHTML')}")
            text = elem.get_attribute('innerHTML').strip()
            print(f"Checking price text: {text}")

            if text.startswith("$") and len(text) > 1:
                price = text
                break
        result['Price'] = price
        print(f"Extracted Price: {price}")
    except Exception as e:
        print(f"Price not found, trying alternative selectors. Exception: {e}")
        result['Price'] = ""

    # --- Extract Description (table + feature bullets) ---
    try:
        # Extract table content
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.a-normal.a-spacing-micro")))
        desc_rows = driver.find_elements(By.CSS_SELECTOR, "table.a-normal.a-spacing-micro tr")
        desc_list = []
        for row in desc_rows:
            tds = row.find_elements(By.TAG_NAME, "td")
            if len(tds) == 2:
                key = tds[0].text.strip()
                value = tds[1].text.strip()
                desc_list.append(f"{key}: {value}")
        table_desc = "; ".join(desc_list)
        #print(f"Extracted Table Description: {table_desc}")
        
        # Extract feature bullets
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#feature-bullets ul.a-unordered-list")))
        bullets = driver.find_elements(By.CSS_SELECTOR, "#feature-bullets ul.a-unordered-list li span.a-list-item")
        bullet_list = [b.text.strip() for b in bullets if b.text.strip()]
        bullets_desc = " | ".join(bullet_list)
        #print(f"Extracted Bullets Description: {bullets_desc}")
        
        # Combine both
        if table_desc and bullets_desc:
            desc = f"{table_desc} || {bullets_desc}"
        elif table_desc:
            desc = table_desc
        else:
            desc = bullets_desc
        #print(f"Combined Description: {desc}")
        result['Description'] = desc
    except:
        result['Description'] = ""

    results.append(result)
    #print(f"Done: {result['Title']}")
    if index == 1: 
        break

driver.quit()

# Save to Excel
output_df = pd.DataFrame(results)
output_df.to_excel("amazon_extracted.xlsx", index=False)
print("Amazon data extraction complete. Results saved to 'amazon_extracted.xlsx'.")