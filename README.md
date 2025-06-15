# Wireless Product Scraper

A Python-based tool to automate the extraction of product details and images from the shopping cart of an e-commerce website (e.g., [balajiwireless.com](https://www.balajiwireless.com/)). The script logs in, navigates to the cart, scrapes product information, and downloads product images, saving all data to an Excel file.

## Features

- **Automated Login:** Uses credentials from a JSON file to log in securely.
- **Cart Scraping:** Extracts product name, link, image URL, price, quantity, SKU, and UPC from the cart.
- **Image Downloading:** Downloads all product images, preserving original filenames and skipping already-downloaded images.
- **Variant Handling:** Supports products with multiple color options, saving images for each variant in separate folders.
- **Robust Error Handling:** Skips items with missing or unexpected structures and logs issues for review.
- **Excel Export:** Saves all scraped data to `cart_items.xlsx`.

## Requirements

- Python 3.7+
- Google Chrome browser
- ChromeDriver (compatible with your Chrome version)

### Python Packages

Install dependencies with:

```bash
pip install selenium pandas openpyxl requests
```

## Setup

1. Clone the repository and navigate to the project directory.
2. Create a `credentials.json` file in the project root with your login details:

```bash
{
    "web_url": "https://www.balajiwireless.com/",
    "email": "your-email@example.com",
    "password": "your-password"
}
```

3. Download ChromeDriver and ensure it is in your PATH or the project directory.

## Usage

- Run the scraper script:

```bash
python scrape_balaji_cart.py
```

- The script will open Chrome, log in, and process your cart.
- Product images will be saved in the `product_images/` directory, organized by product and color variant.
- Scraped data will be saved to `cart_items.xlsx`.

## Configuration

- To run the scraper in headless mode (no browser window), uncomment the `--headless` line in `scrape_balaji_cart.py`:

```bash
options.add_argument("--headless")
```

## Notes

- The script is tailored for the structure of balajiwireless.com. For other sites, selectors may need adjustment.
- Sensitive files like `credentials.json`, product images, and Excel exports are excluded from version control via `.gitignore`.

### License

This project is for internal use. Please respect the terms of service of any website you scrape.

## Author:

AKRIMA HUZAIFA AKHTAR
