#!/usr/bin/env node

/**
 * Stripe Product Downloader
 * 
 * This script fetches product data from Stripe and creates an Excel file.
 * It downloads product images and saves them to the productImages folder.
 * The Excel file is saved to the downloads folder.
 * 
 * Usage:
 *   node stripe-product-downloader.js
 * 
 * Environment variables:
 *   STRIPE_API_KEY - Your Stripe API key (required)
 *   STRIPE_API_VERSION - Stripe API version (optional)
 */

// Import dependencies
const fs = require('fs');
const path = require('path');
const https = require('https');
const Excel = require('exceljs');
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const dotenv = require('dotenv');
const chalk = require('chalk');
const Stripe = require('stripe');

// Load environment variables
dotenv.config();

// Parse command-line arguments
const argv = yargs(hideBin(process.argv))
  .option('limit', {
    alias: 'l',
    describe: 'Maximum number of products to fetch (default: 100)',
    type: 'number',
    default: 100
  })
  .option('dryRun', {
    alias: 'd',
    describe: 'Perform a dry run without downloading images or creating the Excel file',
    type: 'boolean',
    default: false
  })
  .help()
  .alias('help', 'h')
  .version()
  .alias('version', 'v')
  .example('$0', 'Download all products from Stripe (up to default limit)')
  .example('$0 --limit=500', 'Download up to 500 products')
  .example('$0 --dryRun', 'Perform a dry run')
  .argv;

// Validate environment variables
if (!process.env.STRIPE_API_KEY) {
  console.error(chalk.red('Error: STRIPE_API_KEY environment variable is required.'));
  console.error(chalk.yellow('Please set it in a .env file or as an environment variable.'));
  process.exit(1);
}

// Initialize Stripe client
const stripe = new Stripe(process.env.STRIPE_API_KEY, {
  apiVersion: process.env.STRIPE_API_VERSION || null,
});

// Required columns for the Excel file
const REQUIRED_COLUMNS = ['CODE', 'NAME', 'DESCRIPTION', 'PRICE', 'IMAGE'];
const STRIPE_PRODUCT_ID_COLUMN = 'STRIPE_PRODUCT_ID';
const STRIPE_PRICE_ID_COLUMN = 'STRIPE_PRICE_ID';

/**
 * Main function
 */
async function main() {
  try {
    console.log(chalk.blue('Starting Stripe Product Downloader'));
    
    if (argv.dryRun) {
      console.log(chalk.yellow('DRY RUN MODE: No images will be downloaded and no Excel file will be created'));
    }
    
    // Create productImages directory if it doesn't exist
    const productImagesDir = path.join(process.cwd(), 'productImages');
    if (!fs.existsSync(productImagesDir)) {
      if (!argv.dryRun) {
        console.log(chalk.blue('Creating productImages directory...'));
        fs.mkdirSync(productImagesDir);
      } else {
        console.log(chalk.yellow('[DRY RUN] Would create productImages directory'));
      }
    }
    
    // Fetch products from Stripe
    console.log(chalk.blue(`Fetching products from Stripe (limit: ${argv.limit})...`));
    const products = await fetchAllProducts(argv.limit);
    console.log(chalk.green(`Found ${products.length} products in Stripe`));
    
    // Create a new workbook
    const workbook = new Excel.Workbook();
    const worksheet = workbook.addWorksheet('Products');
    
    // Add headers
    const headers = [...REQUIRED_COLUMNS, STRIPE_PRODUCT_ID_COLUMN, STRIPE_PRICE_ID_COLUMN];
    worksheet.columns = headers.map(header => ({ header, key: header.toLowerCase() }));
    
    // Process each product
    console.log(chalk.blue(`Processing ${products.length} products...`));
    
    for (const product of products) {
      console.log(chalk.gray(`Processing product: ${product.id} - ${product.name}`));
      
      // Get product code from metadata
      const productCode = product.metadata?.product_code || '';
      
      // Get product prices
      const prices = await fetchProductPrices(product.id);
      const price = prices.length > 0 ? prices[0] : null;
      const priceInDollars = price ? (price.unit_amount / 100).toFixed(2) : '';
      const priceId = price ? price.id : '';
      
      // Process image
      let imageFileName = '';
      if (product.images && product.images.length > 0) {
        const imageUrl = product.images[0];
        imageFileName = await downloadProductImage(imageUrl, productCode, argv.dryRun);
      }
      
      // Add row to worksheet
      worksheet.addRow({
        code: productCode,
        name: product.name,
        description: product.description || '',
        price: priceInDollars,
        image: imageFileName,
        stripe_product_id: product.id,
        stripe_price_id: priceId
      });
    }
    
    // Format the worksheet
    worksheet.getRow(1).font = { bold: true };
    worksheet.columns.forEach(column => {
      column.width = 20;
    });
    
    // Create downloads directory if it doesn't exist
    const downloadsDir = path.join(process.cwd(), 'downloads');
    if (!fs.existsSync(downloadsDir)) {
      if (!argv.dryRun) {
        console.log(chalk.blue('Creating downloads directory...'));
        fs.mkdirSync(downloadsDir);
      } else {
        console.log(chalk.yellow('[DRY RUN] Would create downloads directory'));
      }
    }
    
    // Generate filename with current date and time
    const now = new Date();
    const dateStr = now.toISOString()
      .replace(/T/, '_')
      .replace(/\..+/, '')
      .replace(/:/g, '-');
    const fileName = `import_${dateStr}.xlsx`;
    const filePath = path.join(downloadsDir, fileName);
    
    // Save the workbook
    if (!argv.dryRun) {
      console.log(chalk.blue(`Saving Excel file as ${filePath}...`));
      await workbook.xlsx.writeFile(filePath);
      console.log(chalk.green(`Excel file created successfully: ${filePath}`));
    } else {
      console.log(chalk.yellow(`[DRY RUN] Would save Excel file as ${filePath}`));
    }
    
    console.log(chalk.green('Process completed successfully!'));
    
  } catch (error) {
    console.error(chalk.red(`Error: ${error.message}`));
    process.exit(1);
  }
}

/**
 * Fetch all products from Stripe
 * @param {number} limit - Maximum number of products to fetch
 * @returns {Promise<Array>} - Array of Stripe product objects
 */
async function fetchAllProducts(limit) {
  const products = [];
  let hasMore = true;
  let startingAfter = null;
  const pageSize = Math.min(limit, 100); // Stripe's max page size is 100
  
  while (hasMore && products.length < limit) {
    const params = {
      limit: pageSize,
      active: true,
    };
    
    if (startingAfter) {
      params.starting_after = startingAfter;
    }
    
    const response = await stripe.products.list(params);
    
    products.push(...response.data);
    hasMore = response.has_more;
    
    if (response.data.length > 0) {
      startingAfter = response.data[response.data.length - 1].id;
    }
    
    console.log(chalk.gray(`Fetched ${response.data.length} products (total: ${products.length})`));
    
    if (products.length >= limit) {
      console.log(chalk.yellow(`Reached product limit of ${limit}`));
      break;
    }
  }
  
  return products;
}

/**
 * Fetch prices for a specific product
 * @param {string} productId - Stripe product ID
 * @returns {Promise<Array>} - Array of Stripe price objects
 */
async function fetchProductPrices(productId) {
  const response = await stripe.prices.list({
    product: productId,
    active: true,
    limit: 100,
  });
  
  return response.data;
}

/**
 * Download a product image from URL
 * @param {string} imageUrl - URL of the image
 * @param {string} productCode - Product code to use in the filename
 * @param {boolean} dryRun - Whether this is a dry run
 * @returns {Promise<string>} - Filename of the downloaded image
 */
async function downloadProductImage(imageUrl, productCode, dryRun) {
  // Check if this is a Stripe FileLink URL
  const isFileLink = imageUrl.includes('files.stripe.com/links/');
  
  // Extract file extension from URL or default to .jpg
  let fileExt = '.jpg';
  if (!isFileLink) {
    const urlParts = imageUrl.split('?')[0].split('.');
    fileExt = urlParts.length > 1 ? `.${urlParts.pop().toLowerCase()}` : '.jpg';
  }
  
  // Generate a filename based on product code or a random string
  const baseFileName = productCode ? 
    `${productCode.replace(/[^a-zA-Z0-9]/g, '_')}${fileExt}` : 
    `product_${Date.now()}${fileExt}`;
  
  const filePath = path.join('productImages', baseFileName);
  
  if (dryRun) {
    console.log(chalk.yellow(`[DRY RUN] Would download image from ${imageUrl} to ${filePath}`));
    return baseFileName;
  }
  
  console.log(chalk.gray(`Downloading image from ${imageUrl} to ${filePath}`));
  
  // Ensure the productImages directory exists
  const productImagesDir = path.join(process.cwd(), 'productImages');
  if (!fs.existsSync(productImagesDir)) {
    fs.mkdirSync(productImagesDir, { recursive: true });
  }
  
  return new Promise((resolve, reject) => {
    const file = fs.createWriteStream(filePath);
    
    https.get(imageUrl, (response) => {
      if (response.statusCode !== 200) {
        file.close();
        fs.unlink(filePath, () => {}); // Delete the file if there was an error
        reject(new Error(`Failed to download image: HTTP status ${response.statusCode}`));
        return;
      }
      
      response.pipe(file);
      
      file.on('finish', () => {
        file.close();
        console.log(chalk.green(`Image downloaded successfully: ${filePath}`));
        resolve(baseFileName);
      });
      
      file.on('error', (err) => {
        file.close();
        fs.unlink(filePath, () => {}); // Delete the file if there was an error
        reject(new Error(`Failed to write image file: ${err.message}`));
      });
    }).on('error', (err) => {
      file.close();
      fs.unlink(filePath, () => {}); // Delete the file if there was an error
      reject(new Error(`Failed to download image: ${err.message}`));
    });
  });
}

// Run the main function
main().catch(error => {
  console.error(chalk.red(`Unhandled error: ${error.message}`));
  process.exit(1);
});
