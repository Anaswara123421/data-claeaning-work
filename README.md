# Data Cleaning in Excel

## Overview
This project demonstrates various data cleaning and preprocessing techniques in Microsoft Excel using a product sales dataset. The objective is to improve data quality by identifying and correcting inconsistencies, handling missing values, removing duplicate records, and preparing the dataset for analysis.

## Dataset
The dataset contains the following fields:

- Product ID
- Product Name
- Brand Name
- Price ($)
- Quantity
- Category

## Tasks Performed

### 1. Handling Missing Values
- Identified missing values in the **Price** column.
- Replaced missing prices using the average price.
- Handled missing values in the **Category** column by assigning appropriate values.

### 2. Standardizing Text
- Corrected inconsistent product names (e.g., "laptop" → "Laptop").
- Fixed spelling mistakes in the **Category** column (e.g., "Electroni" → "Electronics").

### 3. Find and Replace
- Used Excel's **Find & Replace** feature to standardize text values.
- Corrected inconsistent category names.

### 4. Removing Duplicate Records
- Identified duplicate rows.
- Removed duplicate entries using **Data → Remove Duplicates** while preserving unique records.

### 5. Creating Product Brand
- Combined **Brand Name** and **Product Name** into a new **Product Brand** column using:

```excel
=CONCAT(C2," ",B2)
```

or

```excel
=C2&" "&B2
```

### 6. Conditional Formatting
Applied conditional formatting to highlight:

- Electronics products
- Missing values
- Duplicate values

Formula used:

```excel
=F2="Electronics"
```

### 7. Data Validation
- Verified numeric fields.
- Checked for blank cells.
- Ensured consistent formatting.

## Excel Functions Used

- IF()
- AVERAGE()
- COUNTBLANK()
- COUNTIF()
- CONCAT()
- LEFT()
- MID()
- TRIM()
- PROPER()
- SEARCH()
- ISNUMBER()

## Excel Features Used

- Find & Replace
- Conditional Formatting
- Remove Duplicates
- Data Validation
- Sorting
- Filtering
- Flash Fill

## Learning Outcomes

After completing this project, I learned how to:

- Handle missing data effectively.
- Standardize inconsistent text values.
- Remove duplicate records safely.
- Use Excel formulas for data transformation.
- Apply conditional formatting for data visualization.
- Prepare datasets for further analysis.

## Tools Used

- Microsoft Excel

## Project Structure

```
data-cleaning-excel/
│
├── data cleaning work.xlsx
└── README.md
```

## Author

Anaswara
