# Excel Sheets Documentation

This document describes the Excel sheets used for the product ordering system in this repository.

## Files

### 1. products.xlsx
Contains the master product catalog with the following columns:
- **Product ID**: Unique identifier for each product
- **Product Name**: Display name of the product
- **Description**: Detailed description of the product
- **Price**: Unit price
- **One Size**: Indicates if the product is available in only one size (Yes/No)
- **One Color**: Indicates if the product is available in only one color (Yes/No)

### 2. size_and_colors.xlsx
Contains all available size and color combinations for each product:
- **Product ID**: Links to the product in products.xlsx
- **Product Name**: Product name for reference
- **Available Sizes**: Size option for this variant
- **Available Colors**: Color option for this variant

**Note**: As of the latest update, Product A no longer requires a red color entry in every size. Colors are now flexible and can vary by size based on availability.

### 3. orders.xlsx
Contains customer orders with the following columns:
- **Order ID**: Unique order identifier
- **Customer Name**: Name of the customer
- **Product ID**: Links to the product being ordered
- **Product Name**: Product name for reference
- **Size**: Selected size (or "N/A" if product has One Size = Yes)
- **Color**: Selected color (or "N/A" if product has One Color = Yes)
- **Quantity**: Number of units ordered
- **Order Date**: Date of the order
- **Notes**: Additional order information

## Ordering Rules

When creating an order:

1. **For products with "One Size" = Yes**: 
   - The Size field should be set to "N/A" (in an automated system, this field would be disabled)
   
2. **For products with "One Color" = Yes**: 
   - The Color field should be set to "N/A" (in an automated system, this field would be disabled)
   
3. **For products with both "One Size" and "One Color" = Yes**: 
   - Both Size and Color fields should be set to "N/A"

4. **For standard products (both flags = No)**:
   - Size and Color must be selected from the available options in size_and_colors.xlsx

## Recent Changes

### Version 2.0 (January 2026)
1. **Removed red color requirement**: Product A no longer requires a red color entry in every size. Color availability now varies by size based on actual inventory.

2. **Added restriction flags**: The products sheet now includes "One Size" and "One Color" columns to clearly indicate products with limited size or color options.

3. **Streamlined ordering**: The orders sheet now automatically handles products with size/color restrictions by using "N/A" for disabled fields, making the ordering process more intuitive.

## Sample Data

The sheets include sample data demonstrating all scenarios:
- **Product A**: Standard product with multiple sizes and colors (no restrictions)
- **Product B**: One size fits all (One Size = Yes)
- **Product C**: Available in one color only (One Color = Yes)
- **Product D**: Single size and color variant (both flags = Yes)

## Implementation Notes

For integration with an automated system:
- Use the "One Size" and "One Color" flags to dynamically disable/enable form fields
- Validate that size/color selections exist in the size_and_colors sheet
- Automatically populate "N/A" for restricted fields during order creation
- Consider implementing data validation to prevent invalid size/color combinations
