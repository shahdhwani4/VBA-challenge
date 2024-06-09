# VBA-challenge

### Description

Create a script that loops through all the stocks for each quarter and outputs the following information:

- The ticker symbol
- Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
- The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
- The total stock volume of the stock.
- Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
- Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.


### File Structure
VBA Scripts: VBA scripts for the challenge are under "VBA-scripts" folder.
Screenshots of the results: screenshots of results running VBA script on alphabetical_testing.xlsx & Multi_year_stock_data.xlsx are under "Screenshots_vba" folder.
Excel Files: Output excel files are included under "excel-files" folder.

### Instructions
Steps to execute VBA file:
1. Open excel file you want to analyze
2. Go to developer tab in excel
3. Open "Visual Basics"
4. Import file "VBA-scripts/stock-report-vba-script.bas"
5. Run "GenerateQuarterlyReportForWorksheets"
