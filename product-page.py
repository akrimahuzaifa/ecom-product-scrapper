from datetime import date
import json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import requests
from pathlib import Path
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# === Load login credentials ===
with open("credentials.json") as f:
    creds = json.load(f)

EMAIL = creds["email"]
PASSWORD = creds["password"]
WEB_URL = creds["web_url"]


# Read Excel file
df = pd.read_excel('cart_items-link-july-17-2025.xlsx')
data = df.to_dict(orient='records')

options = Options()
options.add_argument("--start-maximized")
#options.add_argument(r'--user-data-dir=C:\Users\Akrima\AppData\Local\Google\Chrome\User Data')
#options.add_argument('--profile-directory=Profile 3')

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)

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

time.sleep(10)  # Wait for page to load after login

results = []

for index, item in enumerate(data):
    product_link = item.get('Product Link')
    if not product_link or not str(product_link).strip() or product_link == "nan" or product_link == "":
        print(f"Skipping product at index {index} due to missing or invalid link => {product_link}")
        continue
    print(f"Opening product link: {product_link}")
    driver.execute_script("window.open(arguments[0]);", product_link)
    driver.switch_to.window(driver.window_handles[-1])
    #driver.get(product_link)
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "h1.product-details-full-content-header-title")))

    # 1. Extract Title
    try:
        title_elem = driver.find_element(By.CSS_SELECTOR, "h1.product-details-full-content-header-title")
        title = title_elem.text.strip()
    except:
        title = f"Product_{index+1}"

    # 2. Extract Overview
    try:
        overview_elem = driver.find_element(By.ID, "product-details-information-tab-content-container-0")
        overview = overview_elem.text.strip()
    except:
        overview = ""

    # 3. Extract Features & Specs (if present)
    features = {}
    try:
        features_tab = driver.find_element(By.CSS_SELECTOR, "li.tab-heading-title a[data-id='1']")
        features_tab.click()
        #scroll to the features tab content
        driver.execute_script("arguments[0].scrollIntoView();", features_tab)
        print(f"Switched to Features tab for {title}")
        # Wait for the features tab content to load
        time.sleep(2)
        rows = driver.find_elements(By.CSS_SELECTOR, ".product-details-information-tab-content-panel.active tr")
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) == 2:
                key = cols[0].text.strip()
                value = cols[1].text.strip()
                features[key] = value
    except:
        pass

    # Prepare folder for images

    #folder_name = re.sub(r'[^\w\- ]', '_', title)  # Replace non-alnum (except _-) with _
    folder_name = title.replace(" ", "_").replace("/", "-")#.replace("__", "_")  # Original logic
    
    full_folder_path = Path("product_images") / folder_name
    full_folder_path.mkdir(parents=True, exist_ok=True)

    # 4. Loop through color options and download images
    try:
        option_inputs = driver.find_elements(By.CSS_SELECTOR, "div.product-views-option-color-container input.product-views-option-color-picker-input")
        print(f"For Index: {index} Name: {title}\nFound {len(option_inputs)} color options.\nProcessing...")

        if not option_inputs or len(option_inputs) == 1:
            # No multi options — download main gallery
            print(f"Only 1 or No color options found for {title}. Downloading main gallery images.")
            images = driver.find_elements(By.CSS_SELECTOR, ".bx-custom-pager img")
            img_urls = [img.get_attribute("src") for img in images]

            for url in img_urls:
                image_name = os.path.basename(url.split("?")[0])
                image_path = full_folder_path / image_name

                if image_path.exists():
                    #print(f"⏩ Skipping download, already exists: {image_path}")
                    continue

                try:
                    with open(image_path, "wb") as f:
                        f.write(requests.get(url).content)
                    print(f"✅ Downloaded: {image_path}")
                except Exception as e:
                    print(f"❌ Failed to download {url}: {e}")

            # Extract SKU, Price, Option Name (main/default)
            try:
                sku_elem = driver.find_element(By.CSS_SELECTOR, "span.product-line-sku-value")
                sku = sku_elem.text.strip()
            except:
                sku = ""

            try:
                price_elem = driver.find_element(By.CSS_SELECTOR, "span.product-views-price-lead")
                #print(f"Price element found: {price_elem.get_attribute('outerHTML')}")
                price = price_elem.text.strip()
            except:
                price = ""

            results.append({
                "Product Link": product_link,
                "Title": title,
                "Overview": overview,
                "Features": features,
                "SKU": sku,
                "Price": price,
                "Option Name": "",
                "Image Folder": str(full_folder_path)
            })

        else:
            # === Loop through each color option
            last_img_set = set()
            for j, input_el in enumerate(option_inputs):
                try:
                    # Switch the color option using JavaScript
                    driver.execute_script("arguments[0].click();", input_el)
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "bx-custom-pager"))
                    )
                    time.sleep(2)  # Let images visually update

                    # Scrape current gallery image URLs
                    images = driver.find_elements(By.CSS_SELECTOR, ".bx-custom-pager img")
                    img_urls = [img.get_attribute("src") for img in images]
                    current_img_set = set(img_urls)

                    # Try to get color name from label
                    label_imgs = input_el.find_elements(By.XPATH, "./following-sibling::img | ../img")
                    if label_imgs:
                        color_alt = label_imgs[0].get_attribute("alt") or ""
                    else:
                        color_alt = ""

                    # If color_alt is empty, try to extract from image filename
                    if not color_alt and img_urls:
                        first_img_url = img_urls[0]
                        image_name = os.path.basename(first_img_url.split("?")[0])
                        # Example: N9AL-SAMGS25_media-Space%20Blue-1.jpg
                        # Extract the part before the last '-' and after the last space
                        match = re.search(r'-([^-]+)-(\d+)\.jpg$', image_name)
                        if match:
                            color_part = match.group(1)
                            # Replace %20 or underscores with space
                            color_alt = color_part.replace('%20', ' ').replace('_', ' ').strip()
                        else:
                            # Fallback: try to get the word before the last '-'
                            parts = image_name.rsplit('-', 2)
                            if len(parts) >= 2:
                                color_alt = parts[-2].replace('%20', ' ').replace('_', ' ').strip()
                            else:
                                color_alt = f"Unknown_Color_{j+1}"

                    if not color_alt:
                        color_alt = f"Unknown_Color_{j+1}"

                    color_folder = full_folder_path / color_alt.strip().replace(" ", "_")
                    color_folder.mkdir(parents=True, exist_ok=True)

                    # Save each image if not already downloaded
                    for url in img_urls:
                        image_name = os.path.basename(url.split("?")[0])
                        #print(f"Processing image: {image_name} for color: {color_alt}")
                        image_path = color_folder / image_name

                        if image_path.exists():
                            #print(f"⏩ Already exists: {image_path}")
                            continue

                        try:
                            with open(image_path, "wb") as f:
                                f.write(requests.get(url).content)
                            print(f"✅ Saved: {image_path}")
                        except Exception as e:
                            print(f"❌ Failed to download {url}: {e}")

                    # Extract SKU, Price, Option Name for this color
                    try:
                        sku_elem = driver.find_element(By.CSS_SELECTOR, "span.product-line-sku-value")
                        #print(f"SKU element found: {sku_elem.get_attribute('outerHTML')}")
                        sku = sku_elem.text.strip()
                    except:
                        sku = ""

                    try:
                        price_elem = driver.find_element(By.CSS_SELECTOR, "div.product-views-price span.product-views-price-lead")
                        #print(f"Price element found: {price_elem.get_attribute('outerHTML')}")
                        price = price_elem.text.strip()
                    except Exception as e:
                        price = ""
                        print(f"⚠️ Price element not found, using empty string.\nError: {e}")

                    results.append({
                        "Product Link": product_link,
                        "Title": title,
                        "Overview": overview,
                        "Features": features,
                        "SKU": sku,
                        "Price": price,
                        "Option Name": color_alt,
                        "Image Folder": str(color_folder)
                    })

                except Exception as e:
                    print(f"⚠️ Error on color variant for {index}th product '{title}': {e}")

    except Exception as e:
        print(f"⚠️ Error processing images for {title}: {e}")
        # If no options, still save the main info
        results.append({
            "Product Link": product_link,
            "Title": title,
            "Overview": overview,
            "Features": features,
            "SKU": "",
            "Price": "",
            "Option Name": "",
            "Image Folder": str(full_folder_path)
        })

    # Add an empty row after each product
    empty_row = {key: "" for key in results[0].keys()} if results else {}
    if empty_row:
        results.append(empty_row)
    print(f"Processed features: {features}")
    # Close the product detail tab and switch back to the cart
    driver.close()
    driver.switch_to.window(driver.window_handles[0])

    # Remove break to process all products
    # if index >= 1:  # For testing, remove this line to process all products
    #     print("Stopping after 2 products for testing.")
    #     break

driver.quit()

# Convert features dict to string for Excel
for r in results:
    r["Features"] = str(r["Features"])

# Save to Excel
output_df = pd.DataFrame(results)
output_df.to_excel(f"extracted_products_{date.today()}.xlsx", index=False)
print(f"Data extraction complete. Results saved to 'extracted_products_{date.today()}.xlsx'.")