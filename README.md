# VBA-challenge

Here is the VBA challenge for module 2.  This project analyzes an Excel file that displays stock purchases and price changes through multiple years.  My script will output the yearly change, percent change and total stock volume for each stock (ticker name) on each page for each worksheet within the workbook.  As a bonus, it will calculate the greatest % increase, greates % decrease, and greatest total volume from the output columns.

# Installation
Run the VBS script 'stockupdate.vbs' while 'Multiple_year_stock_data.xlsm' is open. The sheet should be just the initial data in it, as in the calculated fields should still be empty. In the Excel document, press the button on the first sheet in order to run the macro and output all calculations.

# Repository Details
Within this file, you should find the Excel file with the stock data, the vbs file for my macro, and the screenshots of a successful run of the script.  Within the Excel file, there will be a button on the first sheet to run the macro as intended.

# Resources and comments

In this script there was research needed to understand how to write various parts of the it. Most of this file is modeled from the activites done in module 2. The links below are sources used to research various code commands needed to preform the outputs in the file and they are also commented within the script itself.

'https://www.extendoffice.com/documents/excel/5333-excel-run-macro-multiple-sheets.html
    Was researched to learn how to execute a code over multiple worksheets at once.

'https://learn.microsoft.com/en-us/office/troubleshoot/excel/loop-through-data-using-macro
    Was researched in order to learn how to iterate through a column with an unknown # of rows

'https://www.skillsyouneed.com/num/percent-change.html#:~:text=First%3A%20work%20out%20the%20difference,two%20numbers%20you%20are%20comparing.&text=Then%3A%20divide%20the%20increase%20by,this%20is%20a%20percentage%20decrease
    Was researched in order to know how to calculate percentage change between two values

https://superuser.com/questions/452832/turn-off-scientific-notation-in-excel#:~:text=Unfortunately%20excel%20does%20not%20allow,your%20data%20to%20scientific%20notation.
    Was researched to find out how to format a cell so that it would not output in scientific notation.
