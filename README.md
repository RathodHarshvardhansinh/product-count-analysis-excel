# Product Count Analysis (Excel)

**Veda Technology Data Analytics Internship, Task 15 (Level 1, Day 15)**

Count products by category and find the largest category, using `COUNTIF` and `COUNTIFS` in Excel.

## Objective

Practice `COUNTIF` / `COUNTIFS` by building a product count table and identifying the top category.

## Tools

- Microsoft Excel (formulas, Table, chart)
- Python (`openpyxl`) was used only to generate the sample data and assemble the workbook

## Dataset

Superstore-style dataset (1,500 order lines, 213 unique products, 3 categories, 17 sub-categories, 4 regions).
It has the same columns as the public Superstore dataset, so you can paste the real file into the `Data` sheet and the analysis works the same way.

> Note: the file in `data/superstore_sample.csv` is generated sample data (see `scripts/generate_sample_data.py`), so the numbers below belong to this sample, not to the original Kaggle file.

## Approach

1. Loaded the data into the `Data` sheet and turned it into an Excel Table.
2. Built a `Unique Products` list (one row per Product ID) so unique products can be counted without double counting.
3. **Product Count sheet**
   - `COUNTIF` for order lines per category: `=COUNTIF(Category, A5)`
   - `COUNTIF` on the unique list for unique products per category
   - `COUNTIFS` for the sub-category breakdown: `=COUNTIFS(Category, A13, SubCategory, B13)`
   - `COUNTIFS` for a Category x Region matrix
   - A check cell confirms the total equals the number of data rows
4. **Top Category sheet** uses `INDEX` + `MATCH` + `MAX` so the answer updates by itself if the data changes.
5. Added a bar chart for the category counts.

## Results

| Category | Order lines | Unique products | Share of lines |
|---|---|---|---|
| Office Supplies | 916 | 120 | 61.1% |
| Furniture | 300 | 48 | 20.0% |
| Technology | 284 | 45 | 18.9% |
| **Total** | **1,500** | **213** | **100%** |

**Top category: Office Supplies** (61.1% of all order lines, 120 unique products).
Largest sub-category: **Art** with 122 order lines.

## Deliverables

- Product count table: `Product Count` sheet
- Top category: `Top Category` sheet

## Interview questions

**COUNTIF vs COUNTIFS?**
`COUNTIF` checks one condition on one range. `COUNTIFS` checks several range/condition pairs and counts only rows that meet all of them (for example Category = Furniture AND Region = West).

**Why count unique products?**
Counting rows only tells you how often something was sold. A category can look big because a few products are ordered over and over. Counting unique products shows how wide the range really is, and avoids counting one product many times.

## Project structure

```
product-count-analysis-excel/
├── data/
│   └── superstore_sample.csv
├── excel/
│   └── Product_Count_Analysis.xlsx
├── scripts/
│   ├── generate_sample_data.py
│   └── build_workbook.py
└── README.md
```

## How to run

Open `excel/Product_Count_Analysis.xlsx` in Excel. To rebuild it from scratch:

```bash
pip install openpyxl
python scripts/generate_sample_data.py
python scripts/build_workbook.py
```

## Author

Harshvardhansinh Rathod, Data Analytics Intern, Veda Technology
