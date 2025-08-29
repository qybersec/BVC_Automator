# TMS Input File Format

This document describes the expected Excel file format for the TMS Data Processor.

## File Structure

### Header Information (Rows 1-7)
- **Row 2, Column B**: Report Title
- **Row 4, Column B**: Company Name  
- **Row 6, Column B**: Date Range

### Data Headers (Row 9)
The system expects the following column structure starting from Column C:

| Column | Header Name |
|--------|-------------|
| C | Load No. |
| D | Ship Date |
| E | Origin City |
| F | Origin State |
| G | Origin Postal |
| H | Destination City |
| I | Destination State |
| J | Destination Postal |
| K | Selected Carrier |
| L | Selected Service Type |
| M | Selected Transit Days |
| N | Selected Freight Cost |
| O | Selected Accessorial Cost |
| P | Selected Total Cost |
| Q | Least Cost Carrier |
| R | Least Cost Service Type |
| S | Least Cost Transit Days |
| T | Least Cost Freight Cost |
| U | Least Cost Accessorial Cost |
| V | Least Cost Total Cost |
| W | Potential Savings |

### Data Rows (Starting Row 11)
Each row represents one shipment with the corresponding data in the columns above.

## Sample Data Example

```
Load No.: A12345
Ship Date: 01/15/2024
Origin: New York, NY 10001
Destination: Los Angeles, CA 90210
Selected Carrier: ABC Transport
Selected Service: Standard
Selected Cost: $1,100.00
Least Cost Carrier: XYZ Logistics  
Least Cost Service: Express
Least Cost Cost: $1,070.00
Potential Savings: $30.00
```

## Business Logic Applied

1. **Same Carrier Rule**: If Selected Carrier = Least Cost Carrier, Potential Savings = $0.00
2. **Empty Data Rule**: If Least Cost data is missing, copy Selected Carrier data and set Savings = $0.00  
3. **Negative Savings Rule**: If Potential Savings < $0.00, copy Selected data to Least Cost and set Savings = $0.00

## Output Features

- Color-coded sections (Selected = Blue, Least Cost = Orange, Savings = Green)
- Professional formatting with borders
- Dynamic row heights for long carrier names
- Financial totals row with emphasized Potential Savings total
- Auto-filtering enabled
- Summary statistics calculated

## File Requirements

- Excel format (.xlsx or .xls)
- Maximum file size: 100MB
- Minimum 10 rows of data
- At least 5 columns with data