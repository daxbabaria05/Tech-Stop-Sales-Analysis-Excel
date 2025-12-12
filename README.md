This Excel workbook contains a complete analysis of sales data for Tech Stop, a small IT retail chain on the West Coast of the South Island. The file demonstrates practical Excel skills including data formatting, calculations, conditional formatting, sorting, filtering, Pivot Tables, and charts. It is designed to help visualize business insights and automate calculations for better decision-making.

- Features
Data Formatting & Styling

Merged and centered the title across columns with Cambria 16pt font and bold formatting.

Bolded column headers and aligned text to the left for clarity.

Added borders around all data cells.

Applied dollar formatting with two decimal places where appropriate.

Highlighted specific cells (K4:L5) in yellow for emphasis.

- Calculated Columns

Gross Income (Column M): Total Sales – Cost of Goods Sold.

Gross Margin (Column N): (Total Sales – Cost of Goods Sold) ÷ Total Sales, formatted as a percentage.

Sales Year (Column O): Extracted year from Invoice Date using the YEAR function.

Tax (Column K): Cost of Goods Sold × Tax Rate (from cell L4).

Price Category (Column P): Assigned based on Unit Price using nested IF statements:

1 = Economy (< $400)

2 = Normal ($400–$899.99)

3 = Premium (≥ $900)

- Conditional Formatting

Total Sales values below $1,000 are highlighted in green for quick analysis.

- Data Management

Sorted data first by Sales Year (ascending), then by Total Sales (descending).

Filter applied to show only sales records for Greymouth.

- Pivot Table & Visualization

Pivot Table summarizing Total Sales by Manufacturer for each Branch.

3D Clustered Column Chart visualizing Total Sales by Manufacturer for each Branch.

Chart includes a descriptive title and legend positioned at the top.

- File Structure

Sales Data Worksheet: Raw data with all calculations and formatting applied.

Sales Summary Worksheet: Pivot Table and chart summarizing Total Sales by Manufacturer and Branch.

- Skills Demonstrated

Advanced Excel formatting and styling

Formulas and cell referencing (relative & absolute)

Conditional formatting for data insights

Sorting, filtering, and data management

Pivot Tables and data visualization with charts

- Usage

Open the workbook in Microsoft Excel (recommended version 2016 or later).

Review the Sales worksheet for detailed sales data and calculated metrics.

Explore the Sales Summary worksheet for Pivot Table insights and the 3D Clustered Column chart.

Modify filters or data as needed to analyze other branches or sales periods.

- License

This repository is for educational purposes and does not contain proprietary data.
