# 2020-11-16-Google-Spreadsheet
Google Apps Script function run by buttons on a sheet

A google spreadsheet has two tabs, one behaves as a data entry sheet with "Submit", "Fetch" and "Delete" buttons. The other tab collects the data entered in the entry sheet as a "database.

Here are the button behaviours:

1）Enter a value in the ID cell in the data entry sheet then press the "Fetch" button, the data of the same ID will be copied from the database sheet.

2）Enter a value in the ID cell in the data entry sheet then press the "Delete" button, the data of the same ID will be deleted from the database sheet.

3）Enter a value in the ID cell in the data entry sheet then press the "Fetch" button, the data of the same ID will be copied from the database sheet. Then change some values in the cells, click "Register" button. The data of the same ID will be updated in the database sheet.

4）Enter values in the data entry sheet without an ID, a new row will be added to the database sheet and the new ID will be assigned (auto-increment).

As the data entry sheet has some cells with formulae, those cells will be skipped when the data is copied from the database sheet, so that the formulae will not be removed.


2つのタブ（データエントリーとデータベース）があるGoogle Spreadsheetで、エントリーシート内にボタンを設置し、関数を呼ぶ設定をします。

1）登録IDに数字を入れて、「参照」を押すと、その数字のデータが「データベースシート」から呼ばれます。

2）登録IDに数字を入れて、「削除」を押すと、その数字のデータが「データベースシート」から削除されます。

3）登録IDに数字を入れて、「登録／更新」を押すと、その数字のデータが「データベースシート」で上書きされます。（その場合はまず「参照」を押してから、データの更新を行い、更新ボタンを押してください）

4）登録IDに数字を入れず、「登録／更新」を押すと、新しい行としてデータが「データベースシート」に追加されます。新しいIDが自動採番で割り当てられます。
