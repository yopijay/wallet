# Product Management Excel File

## Overview
This Excel file (`Product_Management.xlsx`) is designed for managing products, their sizes/colors, and orders with automated features for efficiency and ease of use.

## File Structure

### Sheet 1: Products
**Purpose**: Master list of products and their prices

**Columns**:
- `Product Name`: Name of the product
- `Price`: Price of the product in USD

**Sample Data**:
- Product A - $29.99
- Product B - $49.99
- Product C - $19.99
- Product D - $39.99
- Product E - $59.99

### Sheet 2: Sizes and Colors
**Purpose**: Define available sizes and colors for each product

**Columns**:
- `Product Name`: Reference to product from Products sheet
- `Size`: Available size options (e.g., Small, Medium, Large, XS, S, M, L, XL, One Size)
- `Color`: Available color options (e.g., Red, Blue, Black, White, Green, etc.)

**Sample Data**:
The sheet contains 20 different size/color combinations across all products, such as:
- Product A: Small/Medium/Large in Red and Blue
- Product B: Small/Medium/Large in Black, Small/Medium in White
- Product C: One Size in Green and Yellow
- And more...

### Sheet 3: Orders
**Purpose**: Order entry form with automated features

**Columns**:
- `Order ID`: Auto-generated, starts from 1001
- `Product Name`: Dropdown selection from Products sheet
- `Size`: Dropdown selection from Sizes and Colors sheet
- `Color`: Dropdown selection from Sizes and Colors sheet
- `Price`: Auto-populated based on selected product

## Features

### 1. Dropdown Lists (Data Validation)
- **Product Name**: Click the dropdown arrow in column B to select from available products
- **Size**: Click the dropdown arrow in column C to select from available sizes
- **Color**: Click the dropdown arrow in column D to select from available colors

All dropdowns are linked to their respective source sheets, ensuring data consistency.

### 2. Auto-Population

#### Order ID
- **Formula**: `=IF(B{row}="","",1000+ROW()-1)`
- **Behavior**: Automatically generates a unique order ID when a product is selected
- **Starting value**: 1001 for the first order

#### Price
- **Formula**: `=VLOOKUP(B{row},Products!A:B,2,FALSE)`
- **Behavior**: Automatically looks up and displays the price when a product is selected
- **Source**: Retrieves price from the Products sheet

### 3. Documentation and Comments
- Each column header in the Orders sheet has an attached comment explaining its functionality
- To view comments, hover over cells with red corner indicators
- A comprehensive documentation section is included at the bottom of the Orders sheet (rows 103+)

### 4. Formatting
- Professional header styling with blue background and white text
- Bordered cells for clear data separation
- Currency formatting for price columns ($#,##0.00)
- Frozen header rows for easy navigation
- Appropriate column widths for readability

## How to Use

### Creating a New Order
1. Open the Excel file
2. Navigate to the **Orders** sheet
3. Click on a cell in the **Product Name** column (column B)
4. Select a product from the dropdown list
5. Select a **Size** from the dropdown in column C
6. Select a **Color** from the dropdown in column D
7. The **Order ID** and **Price** will automatically populate

### Adding New Products
1. Go to the **Products** sheet
2. Add the new product name in column A and price in column B
3. Go to the **Sizes and Colors** sheet
4. Add all available size/color combinations for the new product
5. The new product will automatically appear in the Orders sheet dropdown

### Modifying Existing Products
- To change a price: Edit the price in the **Products** sheet (column B)
- To add/remove sizes or colors: Modify the **Sizes and Colors** sheet
- Changes will automatically reflect in the Orders sheet

## Technical Details

### Formulas Used
- **Order ID**: `=IF(B2="","",1000+ROW()-1)`
  - Generates sequential IDs starting from 1001
  - Only shows ID if a product is selected

- **Price Lookup**: `=VLOOKUP(B2,Products!A:B,2,FALSE)`
  - Searches for the product name in the Products sheet
  - Returns the corresponding price
  - Shows blank if no product is selected

### Data Validation Ranges
- Product Name dropdown: `=Products!$A$2:$A$6`
- Size dropdown: `='Sizes and Colors'!$B$2:$B$21`
- Color dropdown: `='Sizes and Colors'!$C$2:$C$21`

**Note**: If you add more products or size/color combinations beyond these ranges, you may need to extend the validation formulas in the Orders sheet data validation settings.

## Tips and Best Practices

1. **Data Entry Order**: Always select Product Name first, then Size and Color
2. **Validation**: Verify that the size and color you select are actually available for the chosen product by checking the Sizes and Colors sheet
3. **Backup**: Keep a backup copy before making extensive changes
4. **Extending Ranges**: If you add many new products, update the data validation ranges in the Orders sheet to include all rows
5. **Consistency**: Keep product names consistent across all sheets (exact matches required for formulas to work)

## Troubleshooting

### Price not showing up
- Ensure the product name in Orders matches exactly with the Products sheet
- Check that the product exists in the Products sheet
- Verify the VLOOKUP formula hasn't been deleted

### Dropdown not working
- Check that data validation is still applied to the cell
- Verify the source sheets (Products, Sizes and Colors) contain data
- Ensure the validation range includes the new data if you've added items

### Order ID not generating
- Ensure a product name is selected in column B
- Check that the formula in column A hasn't been overwritten

## Support
For questions or issues with this Excel file, please refer to the comments in the Orders sheet or check the documentation section at the bottom of that sheet.
