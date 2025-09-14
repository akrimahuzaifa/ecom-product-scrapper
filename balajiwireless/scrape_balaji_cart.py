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

time.sleep(10)  # Wait for cart to load after login

# === Step 2: Wait for cart icon and click it ===
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.header-mini-cart-menu-cart-link[data-type='mini-cart']"))).click()

time.sleep(10)  # Wait for cart to load after login

# Step 3: Click "Cart" in mini cart popup
# Wait for the "View Cart" button in the mini-cart and click it
wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "a.header-mini-cart-button-view-cart"))).click()


# Step 4: Wait for the page to load and find cart items
# Wait for the cart page to load by checking for a footer element
wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.footer-links.home")))

# Step 5: Scrape cart items
items_container = driver.find_element(By.CSS_SELECTOR, "div[data-view='Item.ListNavigable']")
item_divs = items_container.find_elements(By.CSS_SELECTOR, "div.cart-lines-row")

# Prepare a list to hold cart data
cart_data = []

# Loop through each item and extract details
for index, item in enumerate(item_divs):
    try:
        # === Get product name and link (if any) ===
        name_el = item.find_elements(By.CSS_SELECTOR, "a.cart-lines-name-link")
        if name_el:
            name_el = name_el[0]
            product_name = name_el.text.strip()
            detail_url = name_el.get_attribute("href")
            has_link = True
        else:
            # Fallback to span-based name (non-clickable item)
            name_span = item.find_elements(By.CSS_SELECTOR, "span.cart-lines-name-viewonly")
            if not name_span:
                print(f"⚠️ Item index {index} has no recognizable product name structure. Skipping.")
                continue
            product_name = name_span[0].text.strip()
            detail_url = None
            has_link = False

        # name_el = item.find_element(By.CSS_SELECTOR, "a.cart-lines-name-link")
        # product_name = name_el.text.strip()
        # detail_url = name_el.get_attribute("href")

        image_el = item.find_element(By.CSS_SELECTOR, "div.cart-lines-thumbnail img")
        image_url = image_el.get_attribute("src")
        
        price_el = item.find_element(By.CSS_SELECTOR, "span.transaction-line-views-price-lead")
        print(f"Price Element in cart list HTML: {price_el.get_attribute("outerHTML")}")
        price = price_el.text.strip()
        print(f"Processing item index {index} =>Name: '{product_name}' =>Price text: '{price}'")

        quantity_el = item.find_element(By.CSS_SELECTOR, "input[data-type='cart-item-quantity-input']")
        quantity = quantity_el.get_attribute("value").strip()
        
        sku_el = item.find_element(By.CSS_SELECTOR, "span.product-line-sku-value")
        sku = sku_el.text.strip()
        
        # Try UPC (if exists)
        upc_els = item.find_elements(By.XPATH, ".//span[contains(text(), 'UPC:')]/following-sibling::span")
        upc = upc_els[0].text.strip() if upc_els else "N/A"
        #print(f"Processing item index {index}:\nName: {product_name}\nLink: {detail_url}\nImage URL: {image_url}\nPrice: {price}\nSKU: {sku}\nUPC: {upc}\nQuantity: {quantity}")
        
        # Step 5.1: Download product images
        if has_link:
            # Visit detail page in new tab
            driver.execute_script("window.open(arguments[0]);", detail_url)
            driver.switch_to.window(driver.window_handles[-1])
            
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CLASS_NAME, "bx-custom-pager"))
            )

            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.product-views-price-lead")))

            if not price:
                detail_price_el = driver.find_element(By.CSS_SELECTOR, "span.product-views-price-lead")
                print(f"Price Element in product detail page, its HTML: {detail_price_el.get_attribute("outerHTML")}")
                price = detail_price_el.text.strip()
                print(f"💲 Got price from product page: {price}")

            # Create a folder for the product images
            folder_name = product_name.replace(" ", "_").replace("/", "-")
            full_folder_path = f"product_images/{folder_name}"
            os.makedirs(f"{full_folder_path}", exist_ok=True)

            # === Find color option radio buttons
            option_inputs = driver.find_elements(By.CSS_SELECTOR, "div.product-views-option-color-container input.product-views-option-color-picker-input")
            print(f"For Index: {index} Name: {product_name}\nFound{len(option_inputs)} color options.")
            
            # If no color options, just download the main gallery
            if not option_inputs or len(option_inputs) == 1:
                # No multi options — download main gallery
                print(f"Only 1 or No color options found for {product_name}. Downloading main gallery images.")
                images = driver.find_elements(By.CSS_SELECTOR, ".bx-custom-pager img")
                img_urls = [img.get_attribute("src") for img in images]

                for url in img_urls:
                    image_name = os.path.basename(url.split("?")[0])  # Strip query params
                    image_path = Path(full_folder_path) / image_name

                    if image_path.exists():
                        print(f"⏩ Skipping download, already exists: {image_path}")
                        continue

                    try:
                        with open(image_path, "wb") as f:
                            f.write(requests.get(url).content)
                        print(f"✅ Downloaded: {image_path}")
                    except Exception as e:
                        print(f"❌ Failed to download {url}: {e}")
            else:
                # === Loop through each color option
                # Track last image set to compare
                last_img_set = set()

                for j, input_el in enumerate(option_inputs):
                    try:
                        # Find color name from associated label img or span
                        label_imgs = input_el.find_elements(By.XPATH, "./following-sibling::img | ../img")
                        if label_imgs:
                            color_alt = label_imgs[0].get_attribute("alt") or "Unknown"
                        else:
                            color_alt = f"Unknown_Color_{j+1}"

                        color_folder = Path(full_folder_path) / color_alt.strip().replace(" ", "_")
                        color_folder.mkdir(parents=True, exist_ok=True)

                        # === Switch the color option using JavaScript
                        driver.execute_script("arguments[0].click();", input_el)

                        # Wait for images to update
                        WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CLASS_NAME, "bx-custom-pager"))
                        )
                        time.sleep(2)  # Let images visually update

                        # Scrape current gallery image URLs
                        images = driver.find_elements(By.CSS_SELECTOR, ".bx-custom-pager img")
                        img_urls = [img.get_attribute("src") for img in images]
                        current_img_set = set(img_urls)

                        if current_img_set == last_img_set:
                            print(f"⚠️ Skipping '{color_alt}' — images same as previous")
                            continue
                        last_img_set = current_img_set

                        # Save each image if not already downloaded
                        for url in img_urls:
                            image_name = os.path.basename(url.split("?")[0])
                            image_path = color_folder / image_name

                            if image_path.exists():
                                print(f"⏩ Already exists: {image_path}")
                                continue

                            try:
                                with open(image_path, "wb") as f:
                                    f.write(requests.get(url).content)
                                print(f"✅ Saved: {image_path}")
                            except Exception as e:
                                print(f"❌ Failed to download {url}: {e}")

                    except Exception as e:
                        print(f"⚠️ Error on color variant for {index}th product '{product_name}': {e}")

            # Close the product detail tab and switch back to the cart
            driver.close()
            driver.switch_to.window(driver.window_handles[0])
        else:
            print(f"ℹ️  Skipping image download for '{product_name}' — no detail link. (OUT OF STOCK?)")
            print(f"Setting Price from — data-rate attribute.")
            if not price:  # fallback if visible text is empty
                price = f"${price_el.get_attribute('data-rate')}"
                print(f"🔁 Used data-rate instead: {price}")
            else:
                print(f"✅ Got price from text: {price}")
        
        # Append item data to cart_data
        cart_data.append({
            "name": product_name,
            "link": detail_url,
            "image": image_url,
            "price": price,
            "quantity": quantity,
            "sku": sku,
            "upc": upc
        })
    except Exception as e:
        print(f"❌ Error processing item index {index}:\nitem: {item.get_attribute("outerHTML")}\n\nError:{e}")


# Save cart_data to Excel
df = pd.DataFrame(cart_data)

# Rename columns if needed
df.columns = ["Product Name", "Product Link", "Image URL", "Price", "Quantity", "SKU", "UPC"]

# Save to Excel file
df.to_excel("cart_items.xlsx", index=False)
print("✅ Saved to cart_items.xlsx")


# === Done ===
driver.quit()
print("✅ All done. Data and images saved.")
