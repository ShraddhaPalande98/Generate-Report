Report Generation with Excel and PDF Export
This project allows users to generate reports from a data set and export them in various formats, specifically Excel and PDF. The export options provide users with an easy way to save, share, and manipulate data in widely-used formats.

Key Features:
Pagination Support: Data is displayed in pages with a specified number of items per page. Only the current page of data is exported to ensure that large datasets are managed efficiently.

Search Functionality: Users can search for specific products using a search bar. This search filter is also applied when exporting the data, ensuring that only relevant records are included in the report.

Sorting: The report can be sorted by various columns like Product Name, Category, and Price. This ensures that the exported data reflects the same sort order as the data displayed on the page.

Export to Excel:

Excel export generates a .xlsx file containing the data displayed on the current page. This file includes column headers (ID, Product Name, Description, Category, Price, and Quantity).

The generated Excel file can be opened in spreadsheet software like Microsoft Excel or Google Sheets for further analysis or sharing.

Export to PDF:

PDF export creates a well-structured, paginated PDF document with the current page of data.

This PDF file can be downloaded and printed or shared for offline reading.

Dynamic Report Generation: The report is generated based on the current page and search criteria. This allows for quick and focused report creation based on the user's needs.

Customizable Column Display: Users have the option to choose which columns to include in the export. This gives users flexibility when exporting only the relevant data fields.

Technologies Used:
ASP.NET Core MVC for backend data processing and report generation.

EPPlus for Excel file generation (creating .xlsx files).

PdfSharpCore for PDF file generation.

Bootstrap for responsive and clean front-end design.

How to Use:
Navigate to the Product List page where the data is displayed in a paginated table.

Use the search bar to filter the data.

Select the page you want to export and click on either the "Export to Excel" or "Export to PDF" button.

The file will be downloaded with the relevant data based on your current view and search settings.
