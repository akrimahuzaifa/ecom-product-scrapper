import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import requests
import time
import os
from pathlib import Path
import json

OUTPUT_FILE = "amazon_extracted.xlsx"

# Check if output file exists to determine where to resume from
processed_links = set()
if os.path.exists(OUTPUT_FILE):
    try:
        existing_df = pd.read_excel(OUTPUT_FILE)
        processed_links = set(existing_df["Product Link"].dropna().tolist())
        # Load existing results
        results = existing_df.to_dict(orient='records')
        print(f"Resuming from existing file. Already processed {len(processed_links)} products.")
    except Exception as e:
        print(f"Error loading existing file: {e}")
        results = []
else:
    results = []

# Read Excel file with product links and headings
df = pd.read_excel('amazon.xlsx')
data = df.to_dict(orient='records')
headings = df.columns.tolist()

options = Options()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

try:
    for index, item in enumerate(data):
        product_link = item.get('Product Link') or item.get('Link')  # Adjust column name as needed
        if not product_link or str(product_link).strip() in ["", "nan"]:
            print(f"Skipping index {index} due to missing link.")
            continue
            
        # Skip if already processed
        if product_link in processed_links:
            print(f"Skipping already processed: {product_link}")
            continue

        print(f"Processing {index+1}/{len(data)}: {product_link}")
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
            wait.until(EC.presence_of_element_located((By.ID, "productTitle")))
            title = driver.find_element(By.ID, "productTitle").text.strip()
            print(f"Extracted Title: {title}")
            result['Title'] = title
        except:
            print("Title not found, trying alternative selectors.")
            result['Title'] = ""
            title = f"Product_{index+1}"

        # --- Extract Price ---
        try:
            price_elements = driver.find_elements(By.CSS_SELECTOR, "span.a-offscreen")
            price = ""
            for elem in price_elements:
                text = elem.get_attribute('innerHTML').strip()
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
            
            # Extract feature bullets
            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#feature-bullets ul.a-unordered-list")))
            bullets = driver.find_elements(By.CSS_SELECTOR, "#feature-bullets ul.a-unordered-list li span.a-list-item")
            bullet_list = [b.text.strip() for b in bullets if b.text.strip()]
            bullets_desc = " | ".join(bullet_list)
            
            # Combine both
            if table_desc and bullets_desc:
                desc = f"{table_desc} || {bullets_desc}"
            elif table_desc:
                desc = table_desc
            else:
                desc = bullets_desc
            result['Description'] = desc
        except:
            result['Description'] = ""

        # Add to results and save progress immediately
        results.append(result)
        output_df = pd.DataFrame(results)
        output_df.to_excel(OUTPUT_FILE, index=False)
        processed_links.add(product_link)
        print(f"✅ Saved data for product {index+1} to {OUTPUT_FILE}")

        # --- Now download images ---
        try:
            words = title.split(" ")
            short_title = " ".join(words[:12])
            folder_name = short_title.replace(" ", "_").replace("/", "-")
            print(f"Creating folder: {folder_name}")
            folder_path = Path("amazon_images") / folder_name
            folder_path.mkdir(parents=True, exist_ok=True)
            
            # Click on the main image to open the image viewer popup
            main_img = driver.find_element(By.ID, "landingImage")
            main_img.click()
            time.sleep(2)  # Wait for popup to fully load
            
            # Make sure the image viewer is loaded
            wait.until(EC.presence_of_element_located((By.ID, "ivLargeImage")))
            
            # Find all thumbnails in the image viewer
            thumbnails = driver.find_elements(By.CSS_SELECTOR, "div.ivThumb")
            print(f"Found {len(thumbnails)} thumbnails in the image viewer")
            
            img_urls = set()
            
            # For each thumbnail
            for i, thumb in enumerate(thumbnails):
                try:
                    # Click on the thumbnail to display the image
                    thumb.click()
                    time.sleep(1)  # Wait for image to load
                    
                    # Extract the image URL from the large image
                    large_img = driver.find_element(By.CSS_SELECTOR, "#ivLargeImage img")
                    img_url = large_img.get_attribute("src")
                    
                    if img_url and img_url not in img_urls:
                        img_urls.add(img_url)
                        
                        # Download the image
                        image_name = os.path.basename(img_url.split("?")[0])
                        image_path = folder_path / image_name
                        if not image_path.exists():
                            try:
                                with open(image_path, "wb") as f:
                                    f.write(requests.get(img_url, timeout=10).content)
                                print(f"✅ Downloaded: {image_path}")
                            except Exception as e:
                                print(f"❌ Failed to download {img_url}: {e}")
                except Exception as e:
                    print(f"Error processing thumbnail #{i}: {e}")
            
            # Close the image viewer by pressing Escape
            actions = ActionChains(driver)
            actions.send_keys(Keys.ESCAPE).perform()
            time.sleep(1)
            
            print(f"Downloaded {len(img_urls)} unique images for {title}")

        except Exception as e:
            print(f"⚠️ Error downloading images for {title}: {e}")
            # Try to close popup if it's still open
            try:
                actions = ActionChains(driver)
                actions.send_keys(Keys.ESCAPE).perform()
            except:
                pass

except Exception as e:
    print(f"Script interrupted: {e}")
finally:
    driver.quit()
    print(f"Amazon data extraction complete. Results saved to '{OUTPUT_FILE}'.")