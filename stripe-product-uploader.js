#!/usr/bin/env node

/**
 * Stripe Product Uploader
 * 
 * This script reads product data from an Excel file and creates products in Stripe.
 * It then updates the Excel file with the Stripe Product and Price IDs.
 * 
 * Usage:
 *   node stripe-product-uploader.js --file=products.xlsx
 * 
 * Environment variables:
 *   STRIPE_API_KEY - Your Stripe API key (required)
 *   STRIPE_API_VERSION - Stripe API version (optional)
 */

// Import dependencies
const fs = require('fs');
const path = require('path');
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
  .option('file', {
    alias: 'f',
    describe: 'Path to the Excel file containing product data',
    type: 'string',
    demandOption: true
  })
  .option('output', {
    alias: 'o',
    describe: 'Path to save the updated Excel file (defaults to overwriting the input file)',
    type: 'string'
  })
  .option('dryRun', {
    alias: 'd',
    describe: 'Perform a dry run without making changes to Stripe or the Excel file',
    type: 'boolean',
    default: false
  })
  .help()
  .alias('help', 'h')
  .version()
  .alias('version', 'v')
  .example('$0 --file=products.xlsx', 'Process products from products.xlsx')
  .example('$0 --file=products.xlsx --output=updated_products.xlsx', 'Save results to a new file')
  .example('$0 --file=products.xlsx --dryRun', 'Perform a dry run')
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

// Required columns in the Excel file
const REQUIRED_COLUMNS = ['CODE', 'NAME', 'DESCRIPTION', 'PRICE', 'IMAGE'];
const STRIPE_PRODUCT_ID_COLUMN = 'STRIPE_PRODUCT_ID';
const STRIPE_PRICE_ID_COLUMN = 'STRIPE_PRICE_ID';

/**
 * Main function
 */
async function main() {
  try {
    const inputFilePath = argv.file;
    const outputFilePath = argv.output || inputFilePath;
    
    console.log(chalk.blue('Starting Stripe Product Uploader'));
    console.log(chalk.gray(`Input file: ${inputFilePath}`));
    console.log(chalk.gray(`Output file: ${outputFilePath}`));
    
    if (argv.dryRun) {
      console.log(chalk.yellow('DRY RUN MODE: No changes will be made to Stripe or the Excel file'));
    }
    
    // Validate input file exists
    if (!fs.existsSync(inputFilePath)) {
      throw new Error(`Input file not found: ${inputFilePath}`);
    }
    
    // Read the Excel file
    console.log(chalk.blue('Reading Excel file...'));
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(inputFilePath);
    
    // Get the first worksheet
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error('No worksheet found in the Excel file');
    }
    
    // Validate worksheet structure
    validateWorksheetStructure(worksheet);
    
    // Add Stripe ID columns if they don't exist
    const headerRow = worksheet.getRow(1);
    let productIdColIndex = -1;
    let priceIdColIndex = -1;
    
    headerRow.eachCell((cell, colNumber) => {
      if (cell.value === STRIPE_PRODUCT_ID_COLUMN) {
        productIdColIndex = colNumber;
      } else if (cell.value === STRIPE_PRICE_ID_COLUMN) {
        priceIdColIndex = colNumber;
      }
    });
    
    // Add columns if they don't exist
    if (productIdColIndex === -1) {
      productIdColIndex = headerRow.cellCount + 1;
      headerRow.getCell(productIdColIndex).value = STRIPE_PRODUCT_ID_COLUMN;
    }
    
    if (priceIdColIndex === -1) {
      priceIdColIndex = headerRow.cellCount + 1;
      headerRow.getCell(priceIdColIndex).value = STRIPE_PRICE_ID_COLUMN;
    }
    
    // Get column indices for required columns
    const columnIndices = {};
    headerRow.eachCell((cell, colNumber) => {
      if (REQUIRED_COLUMNS.includes(cell.value)) {
        columnIndices[cell.value] = colNumber;
      }
    });
    
    // Process each row
    const rowCount = worksheet.rowCount;
    console.log(chalk.blue(`Processing ${rowCount - 1} products...`));
    
    for (let rowNumber = 2; rowNumber <= rowCount; rowNumber++) {
      const row = worksheet.getRow(rowNumber);
      
      // Skip empty rows
      if (row.getCell(columnIndices.CODE).value === null) {
        continue;
      }
      
      const productCode = row.getCell(columnIndices.CODE).value.toString();
      const productName = row.getCell(columnIndices.NAME).value.toString();
      const productDescription = row.getCell(columnIndices.DESCRIPTION).value?.toString() || '';
      // Convert price from dollars to cents (Stripe requires amounts in smallest currency unit)
      const priceInDollars = parseFloat(row.getCell(columnIndices.PRICE).value);
      const productPrice = Math.round(priceInDollars * 100); // Round to avoid floating point issues
      const imageFileName = row.getCell(columnIndices.IMAGE).value?.toString();
      
      // Skip products without images
      if (!imageFileName) {
        console.warn(chalk.yellow(`Warning: Product ${productCode} has no image specified, skipping...`));
        continue;
      }
      
      // Validate price
      if (isNaN(productPrice) || productPrice <= 0) {
        console.warn(chalk.yellow(`Warning: Invalid price for product ${productCode}, skipping...`));
        continue;
      }
      
      // Check if image file exists
      const imagePath = path.join('productImages', imageFileName);
      if (!fs.existsSync(imagePath)) {
        console.warn(chalk.yellow(`Warning: Image file not found for product ${productCode}: ${imagePath}, skipping...`));
        continue;
      }
      
      console.log(chalk.gray(`Processing product: ${productCode} - ${productName}`));
      
      // Check if product already has Stripe IDs
      const existingProductId = row.getCell(productIdColIndex).value;
      const existingPriceId = row.getCell(priceIdColIndex).value;
      
      if (existingProductId && existingPriceId) {
        console.log(chalk.yellow(`Product ${productCode} already has Stripe IDs, skipping...`));
        continue;
      }
      
      if (!argv.dryRun) {
        // Create product in Stripe with image
        const product = await createStripeProduct(productName, productDescription, productCode, imagePath);
        console.log(chalk.green(`Created Stripe product: ${product.id}`));
        
        // Create price in Stripe with product code as description
        const price = await createStripePrice(product.id, productPrice, productCode);
        console.log(chalk.green(`Created Stripe price: ${price.id}`));
        
        // Update Excel row with Stripe IDs
        row.getCell(productIdColIndex).value = product.id;
        row.getCell(priceIdColIndex).value = price.id;
      } else {
        console.log(chalk.yellow(`[DRY RUN] Would upload image to Stripe and create FileLink: ${imagePath}`));
        console.log(chalk.yellow(`[DRY RUN] Would create Stripe product for ${productCode} with public image URL`));
        console.log(chalk.yellow(`[DRY RUN] Would create Stripe price for ${productCode} with nickname: ${productCode} (${priceInDollars} USD = ${productPrice} cents)`));
      }
    }
    
    // Save the updated Excel file
    if (!argv.dryRun) {
      console.log(chalk.blue(`Saving updated Excel file to ${outputFilePath}...`));
      await workbook.xlsx.writeFile(outputFilePath);
      console.log(chalk.green('Excel file updated successfully!'));
    } else {
      console.log(chalk.yellow('[DRY RUN] Would save updated Excel file'));
    }
    
    console.log(chalk.green('Process completed successfully!'));
    
  } catch (error) {
    console.error(chalk.red(`Error: ${error.message}`));
    process.exit(1);
  }
}

/**
 * Validate the worksheet structure
 * @param {Excel.Worksheet} worksheet - The worksheet to validate
 */
function validateWorksheetStructure(worksheet) {
  const headerRow = worksheet.getRow(1);
  const missingColumns = [];
  
  // Check for required columns
  REQUIRED_COLUMNS.forEach(columnName => {
    let found = false;
    headerRow.eachCell((cell) => {
      if (cell.value === columnName) {
        found = true;
      }
    });
    
    if (!found) {
      missingColumns.push(columnName);
    }
  });
  
  if (missingColumns.length > 0) {
    throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
  }
}

/**
 * Upload an image to Stripe and create a public FileLink
 * @param {string} imagePath - Path to the image file
 * @returns {Promise<string>} - Public URL for the image
 */
async function uploadImageToStripe(imagePath) {
  console.log(chalk.gray(`Uploading image: ${imagePath}`));
  
  // Validate that the file exists
  if (!fs.existsSync(imagePath)) {
    throw new Error(`Image file not found: ${imagePath}`);
  }
  
  // Get file stats to check size
  const stats = fs.statSync(imagePath);
  if (stats.size === 0) {
    throw new Error(`Image file is empty: ${imagePath}`);
  }
  
  // Determine MIME type based on file extension
  const ext = path.extname(imagePath).toLowerCase();
  let mimeType = 'application/octet-stream';
  if (ext === '.jpg' || ext === '.jpeg') {
    mimeType = 'image/jpeg';
  } else if (ext === '.png') {
    mimeType = 'image/png';
  } else if (ext === '.gif') {
    mimeType = 'image/gif';
  }
  
  try {
    // 1. Upload the file to Stripe
    const fileData = fs.readFileSync(imagePath);
    console.log(chalk.gray(`Successfully read image file (${fileData.length} bytes)`));
    
    const file = await stripe.files.create({
      purpose: 'product_image',
      file: {
        data: fileData,
        name: path.basename(imagePath),
        type: mimeType,
      },
    });
    
    console.log(chalk.green(`Successfully uploaded image to Stripe: ${file.id}`));
    
    // 2. Create a FileLink to make the file publicly accessible
    const fileLink = await stripe.fileLinks.create({
      file: file.id,
    });
    
    console.log(chalk.green(`Created public FileLink: ${fileLink.url}`));
    
    return fileLink.url;
  } catch (error) {
    throw new Error(`Failed to upload image to Stripe: ${error.message}`);
  }
}

/**
 * Create a product in Stripe
 * @param {string} name - Product name
 * @param {string} description - Product description
 * @param {string} productCode - Product code for metadata
 * @param {string} imagePath - Path to the product image
 * @returns {Promise<Object>} - Stripe product object
 */
async function createStripeProduct(name, description, productCode, imagePath) {
  // First upload the image and get a public URL
  const imageUrl = await uploadImageToStripe(imagePath);
  
  console.log(chalk.gray(`Using public image URL for product: ${imageUrl}`));
  
  // Then create the product with the image
  try {
    const product = await stripe.products.create({
      name,
      description,
      metadata: {
        product_code: productCode
      },
      images: [imageUrl]
    });
    
    // Verify the product was created with the image
    if (!product.images || product.images.length === 0) {
      console.warn(chalk.yellow(`Warning: Product created but no images were attached. Product ID: ${product.id}`));
    } else {
      console.log(chalk.gray(`Product created with ${product.images.length} images: ${product.images.join(', ')}`));
    }
    
    return product;
  } catch (error) {
    throw new Error(`Failed to create product in Stripe: ${error.message}`);
  }
}

/**
 * Create a price in Stripe
 * @param {string} productId - Stripe product ID
 * @param {number} amount - Price amount in cents
 * @param {string} productCode - Product code to use as nickname
 * @returns {Promise<Object>} - Stripe price object
 */
async function createStripePrice(productId, amount, productCode) {
  console.log(chalk.gray(`Creating price for product ${productId}: ${amount} cents (${amount/100} USD)`));
  
  try {
    const price = await stripe.prices.create({
      product: productId,
      unit_amount: amount,
      currency: 'usd',
      nickname: productCode
    });
    
    console.log(chalk.gray(`Price created successfully: ${price.id}, amount: ${price.unit_amount} cents`));
    return price;
  } catch (error) {
    throw new Error(`Failed to create price in Stripe: ${error.message}`);
  }
}

// Run the main function
main().catch(error => {
  console.error(chalk.red(`Unhandled error: ${error.message}`));
  process.exit(1);
});
