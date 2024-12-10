# Stock Analysis VBA Script

## Overview

This VBA script performs stock analysis on a dataset, calculating quarterly changes, percentage changes, and total stock volumes for unique stock tickers. It highlights positive and negative changes using colors and formats percentage values in the results.

## How It Works

1. **Unique Tickers Extraction**:
   - The script extracts unique stock tickers in the input sheet and places them in column `I`.

2. **Data Analysis**:
   - For each unique ticker, the script calculates:
     - **Quarterly Change**: The difference between the close price and the open price.
     - **Percentage Change**: The percentage change relative to the open price.
     - **Total Stock Volume**: The sum of all volumes for the ticker.

3. **Conditional Formatting**:
   - Highlights the **Quarterly Change** in column `J`:
     - **Green** for positive values.
     - **Red** for negative values.
     - **White** for zero values.

4. **Percentage Formatting**:
   - Formats the **Percentage Change** values in column `K` as percentages with two decimal places.

## VBA Code Functionality

### Key Functions

- **Extract Unique Tickers**:
  Uses `AdvancedFilter` to copy unique stock tickers from column `A` to column `I`.

- **Calculate Metrics**:
  Loops through the dataset to compute:
  - **Quarterly Change**: `Close Price - Open Price`
  - **Percentage Change**: `(Quarterly Change / Open Price)`
  - **Total Volume**: Sum of volumes for each ticker.

- **Highlight Quarterly Change**:
  Applies conditional formatting to column `J`:
  - **Green**: `> 0`
  - **Red**: `< 0`
  - **White**: `= 0`

- **Format Percentage Change**:
  Formats column `K` values as percentages with two decimal places.

### Key Loops

- **Outer Loop**:
  Iterates through unique tickers in column `I`.

- **Inner Loop**:
  Iterates through all rows in the dataset to compute metrics for the current ticker.

- **Percentage Formatting Loop**:
  Formats the `K` column to display percentage values correctly.
