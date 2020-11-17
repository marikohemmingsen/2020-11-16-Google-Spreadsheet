# 2020-11-16-Google-Spreadsheet
Google Apps Script function run by buttons on a sheet

A google spreadsheet has two tabs, one behaves as a data entry sheet with "Submit", "Fetch" and "Delete" buttons. The other tab collects the data entered in the entry sheet as a "database.

Here are the button behaviours:

1) Enter a value in the ID cell in the data entry sheet then press the "Fetch" button, the data of the same ID will be copied from the database sheet.

2）Enter a value in the ID cell in the data entry sheet then press the "Delete" button, the data of the same ID will be deleted from the database sheet.

3）Enter a value in the ID cell in the data entry sheet then press the "Fetch" button, the data of the same ID will be copied from the database sheet. Then change some values in the cells, click "Register" button. The data of the same ID will be updated in the database sheet.

4）Enter values in the data entry sheet without an ID, a new row will be added to the database sheet and the new ID will be assigned (auto-increment).

As the data entry sheet has some cells with formulae, those cells will be skipped when the data is copied from the database sheet, so that the formulae will not be removed.
