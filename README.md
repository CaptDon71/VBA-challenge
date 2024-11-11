# Visual Basic Scripting Challenge
Please see my completed XLSM document, screenshot of the results, separate VBA script file, and README file found in the linked GitHub repository above.

**StockMarketAnalysis Macro**
This macro performs a comprehensive stock market analysis across multiple worksheets within an Excel workbook. It calculates and outputs various financial metrics, including the greatest percent increase, greatest percent decrease, and greatest total volume.

**Instructions**
To use the StockMarketAnalysis macro, follow these steps:

**Prepare Your Workbook:**
Ensure that your workbook contains multiple worksheets with stock market data.
Each worksheet should have the following columns:
Column A: Ticker
Column C: Opening Price
Column F: Closing Price
Column G: Volume
Any columns with dates should be formatted as mm/dd/yyyy.


**Running the Macro:**
Open the Excel workbook containing the macro.
Press ALT + F11 to open the VBA Editor.
In the VBA Editor, go to Insert > Module to create a new module.
Copy and paste the StockMarketAnalysis macro code into the new module.
Close the VBA Editor.
To run the macro, press ALT + F8, select StockMarketAnalysis, and click Run.

**Output Results:**
The macro processes each worksheet and calculates the following for each ticker:
Quarterly Change
Percentage Change
Total Stock Volume
Results are output in columns I to M of each worksheet.

After processing all worksheets, the macro outputs the greatest values across all worksheets:
Greatest Percent Increase
Greatest Percent Decrease
Greatest Total Volume
These results are displayed in columns N to P of the first worksheet.

**Formatting:**
Columns containing percentages (Percentage Change) should be manually formatted as percentages:
Select the cells or columns.
Right-click and choose Format Cells.
Select Percentage and set the desired number of decimal places.

**Additional Notes:**
Ensure that all worksheets follow the same column structure for accurate analysis.
The macro processes all worksheets in the workbook, so additional worksheets will also be analyzed.
The greatest values are updated only on the first worksheet after all worksheets have been processed.

**Troubleshooting**
If the macro does not run, ensure that macros are enabled in Excel.
Verify that the column headers and data align with the expected format.
Ensure that date columns are formatted correctly and that there are no inconsistencies in data or formatting.

**Contact**
For further assistance, please contact Don Doggett at dondoggett@att.net


