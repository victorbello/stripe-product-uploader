# Stripe Product Tools

This repository contains Node.js scripts for synchronizing product data between Excel files and the Stripe product catalog.

## Stripe Product Uploader

A script that processes product data from an Excel file and synchronizes it with the Stripe product catalog. The script creates products in Stripe and updates the Excel file with the corresponding Stripe Product and Price IDs.

## Stripe Product Downloader

A script that fetches product data from Stripe and creates an Excel file with the same structure as the one used for uploading. The script downloads product images and saves them to the productImages folder.

## Features

### Uploader Features
- Reads product data from an Excel file
- Creates products and prices in Stripe
- Updates the Excel file with Stripe IDs
- Supports dry run mode for testing
- Provides detailed logging
- Handles errors gracefully

### Downloader Features
- Fetches products and prices from Stripe
- Downloads product images to the productImages folder
- Creates an Excel file with the same structure as the one used for uploading
- Saves Excel files to the downloads folder
- Names the Excel file with the current date and time
- Supports dry run mode for testing
- Provides detailed logging
- Handles errors gracefully

## Prerequisites

- Node.js v18 or newer
- A Stripe account with API access
- An Excel file with product data

## Required Stripe API Permissions

The Stripe API key used for this script requires the following permissions:

- **Products**: `write` - To create products in your Stripe catalog
- **Prices**: `write` - To create prices for products
- **Files**: `write` - To upload product images

When creating a restricted API key in your Stripe Dashboard, make sure to enable these permissions.

## Installation

1. Clone or download this repository
2. Install dependencies:

```bash
npm install
```

3. Create a `.env` file based on the provided `.env.example`:

```bash
cp .env.example .env
```

4. Add your Stripe API key to the `.env` file:

```
STRIPE_API_KEY=your_restricted_api_key_here
```

## Excel File Format

Your Excel file should have the following columns:

- `CODE`: A unique identifier for the product
- `NAME`: The product name
- `DESCRIPTION`: The product description
- `PRICE`: The product price in dollars (will be converted to cents for Stripe)
- `IMAGE`: The filename of the product image (must exist in the `productImages` folder)

The script will add two new columns to the Excel file:

- `STRIPE_PRODUCT_ID`: The Stripe Product ID
- `STRIPE_PRICE_ID`: The Stripe Price ID

## Product Images

The script looks for product images in the `productImages` folder. Each product in the Excel file must have a corresponding image specified in the `IMAGE` column. Products without images will be skipped.

When creating products in Stripe:

1. The image is uploaded to Stripe using the Files API with purpose 'product_image'
2. A FileLink is created to make the image publicly accessible
3. The public FileLink URL is used when creating the product
4. The product code is added as metadata to the product
5. The product code is used as the nickname for the price

This approach ensures that product images are properly displayed in the Stripe dashboard and on customer-facing pages without requiring authentication.

## Usage

### Uploader Usage

Run the uploader script with the path to your Excel file:

```bash
node stripe-product-uploader.js --file=StripeProducts.xlsx
```

#### Uploader Command-line Options

- `--file`, `-f`: Path to the Excel file (required)
- `--output`, `-o`: Path to save the updated Excel file (defaults to overwriting the input file)
- `--dryRun`, `-d`: Perform a dry run without making changes to Stripe or the Excel file
- `--help`, `-h`: Show help
- `--version`, `-v`: Show version

#### Uploader Examples

Process products from an Excel file:

```bash
node stripe-product-uploader.js --file=StripeProducts.xlsx
```

Save the updated data to a new file:

```bash
node stripe-product-uploader.js --file=StripeProducts.xlsx --output=updated_products.xlsx
```

Perform a dry run to test without making changes:

```bash
node stripe-product-uploader.js --file=StripeProducts.xlsx --dryRun
```

### Downloader Usage

Run the downloader script to fetch products from Stripe:

```bash
node stripe-product-downloader.js
```

#### Downloader Command-line Options

- `--limit`, `-l`: Maximum number of products to fetch (default: 100)
- `--dryRun`, `-d`: Perform a dry run without downloading images or creating the Excel file
- `--help`, `-h`: Show help
- `--version`, `-v`: Show version

#### Downloader Examples

Download all products from Stripe (up to default limit):

```bash
node stripe-product-downloader.js
```

Download up to 500 products:

```bash
node stripe-product-downloader.js --limit=500
```

Perform a dry run:

```bash
node stripe-product-downloader.js --dryRun
```

## Error Handling

The script includes error handling for common issues:

- Missing or invalid Excel file
- Missing required columns in the Excel file
- Invalid price values
- Stripe API errors

If an error occurs during processing, the script will immediately stop and exit with an error message. This prevents partial uploads and helps identify issues quickly.

## Troubleshooting

### Missing Required Columns

If you see an error about missing required columns, make sure your Excel file has columns named exactly: `CODE`, `NAME`, `DESCRIPTION`, `PRICE`, and `IMAGE`.

### Stripe API Key Issues

If you encounter authentication errors, check that your Stripe API key is correctly set in the `.env` file and that it has the necessary permissions.

### Excel File Access Issues

If the script cannot read or write to the Excel file, make sure the file is not open in another program and that you have the necessary permissions.

## License

This project is licensed under the MIT License.
