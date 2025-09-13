# Google Sheets Apps Script Project

This repository contains the code and links for two Google Sheets: Workbook A (form input) and Workbook B (database with web app). Workbook A collects data via a form and posts it to Workbook B using Apps Script. Workbook B has a deployed web app to handle incoming data.

## Apps Script Code
- Form script for Workbook A: [FormScript.gs](Userformscript.gs)
- Web app script for Workbook B: [WebApp.gs](Webappscript.gs)

## View the Sheets
- Workbook A (Form): [View Link](https://docs.google.com/spreadsheets/d/1XDsfI5EPZsa5frLXL4JTmWJIcmed0X2Q59iwye6hSHw/edit?usp=sharing)
- Workbook B (Database): [View Link](https://docs.google.com/spreadsheets/d/1nR1XPxWwaCvaRCSZS6jI8TJZ6hJOK-lRO5mX59FV81s/edit?usp=sharing)

## Embed the Sheets
### Workbook A (Form Input)
<iframe src="https://docs.google.com/spreadsheets/d/1XDsfI5EPZsa5frLXL4JTmWJIcmed0X2Q59iwye6hSHw/edit?usp=sharing" width="800" height="600"></iframe>

### Workbook B (Database - Online Sales Data)
<iframe src="https://docs.google.com/spreadsheets/d/1nR1XPxWwaCvaRCSZS6jI8TJZ6hJOK-lRO5mX59FV81s/edit?usp=sharing" width="800" height="600"></iframe>

## How It Works
- **Workbook A**: Contains a form where users input data. An Apps Script (e.g., `FormScript.gs`) posts this data to Workbook B.
- **Workbook B**: Acts as the database, storing data received from Workbook A. It includes a deployed web app (e.g., `WebApp.gs`) to handle incoming data.
- **Data Example**: Workbook B contains sales data with columns like Transaction ID, Date, Product Category, Product Name, Units Sold, Unit Price, Total Revenue, Region, and Payment Method.

To use:
1. Open the view links above to see the sheets.
2. Copy the Apps Script code to your own Google Sheets project.
3. Ensure sharing settings allow access (set to "Anyone with the link" as Viewer).
4. Deploy the web app in Workbook B as described in the script comments.

**Note**: Ensure no sensitive data is shared in the sheets. Use dummy data for public sharing.
