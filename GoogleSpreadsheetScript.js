/* 動作の流れ　

1)　新規データ入力
1．データ入力シートに新しいデータが入力され、「登録」ボタンが押される。（ボタンから関数を呼ぶ）
2．入力されたデータをデータ一覧用タブにコピーする。IDは自動採番で新しいIDを作る。

2)　すでにあるデータの呼び出し
1．データ入力シートのID欄に、すでにあるデータのIDが入力され、「参照」ボタンが押される。（ボタンから関数を呼ぶ）
2．データ一覧タブから同IDのデータを入力シートにコピーする。この際、関数の入っているセルにはコピーしない。

3)　すでにあるデータの消去
1．上記2)の後で、「削除」ボタンが押される。（ボタンから関数を呼ぶ）
2．同IDのデータの行をデータ一覧タブから消去する。

4)　既存データ更新
1．＃2の「参照」の後、「登録」が押されたら、データベースシートの同じIDの行を更新する。


*/

/* */
const thisSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

/* データ入力用のタブの名前　*/
const dataEntrySheetName = '登録シート';

/* データ一覧用のタブの名前　*/
const databaseSheetName = '登録データベース';

/* データ入力シートのセル番号一覧　
*/
const dataEntryCells = [
    'B2', // index 0 - No
    'B6', // index 1 - 日付
    'F6', // index 2 - 会場名
    'F8', // index 3 - 市区町村 【関数】
    'B8', // index 4 - イベント名
    'C10', // index 5 - 前売料金（A）
    'C12', // index 6 - 前売料金（B）
    'C14', // index 7 - 前売料金（C）
    'C16', // index 8 - 前売配信料金
    'C18', // index 9 - 当日料金（A）
    'C20', // index 10 - 当日料金（B）
    'C22', // index 11 - 当日料金（C）
    'C24', // index 12 - 当日配信料金
    'E10', // index 13 - 前売り備考
    'E18', // index 14 - 当日備考
    'B26', // index 15 - ドリンク代
    'E26', // index 16 - 部数
    'A36', // index 17 - メモ
    'J6', // index 18 - アーティスト名
    'M6', // index 19 - RK 【関数】
    'K8', // index 20 - 前売集客（A）
    'K10', // index 21 - 前売集客（B）
    'K12', // index 22 - 前売集客（C）
    'K14', // index 23 - 前売集客（配信）
    'K16', // index 24 - 当日集客（A）
    'K18', // index 25 - 当日集客（B）
    'K20', // index 26 - 当日集客（C）
    'K22', // index 27 - 当日集客（配信）
    'M10', // index 28 - 出演料固定
    'M12', // index 29 - 出演料歩合（CB・%）
    'M14', // index 30 - 出演料歩合（CB・定額）
    'M16', // index 31 - 出演料歩合（配信・％）
    'M18', // index 32 - 出演料歩合（配信・定額）
    'M20', // index 33 - その他出演料
    'K24', // index 34 - 前売来場（A）収益 【関数】
    'K26', // index 35 - 前売来場（B）収益 【関数】
    'K28', // index 36 - 前売来場（C）収益 【関数】
    'K30', // index 37 - 前売配信収益 【関数】
    'K32', // index 38 - 当日来場（A）収益 【関数】
    'K34', // index 39 - 当日来場（B）収益 【関数】
    'K36', // index 40 - 当日来場（C）収益 【関数】
    'K38', // index 41 - 当日配信収益 【関数】
    'K40', // index 42 - その他収益
    'M40', // index 43 - その他収益備考
    'K42', // index 44 - 収益合計 【関数】
    'M24', // index 45 - 固定費用 【関数】
    'M26', // index 46 - 歩合（CB・％）費用 【関数】
    'M28', // index 47 - 歩合（CB・定額）費用 【関数】
    'M30', // index 48 - 歩合（配信・％） 【関数】
    'M32', // index 49 - 歩合（配信・定額） 【関数】
    'M34', // index 50 - その他費用 【関数】
    'M42', // index 51 - 費用合計 【関数】
    'M36' // index 52 - 収益対費用 【関数】
];

/* データベースシートの列一覧　
（入力シートからデータベースに値をコピーするときに使用する)
*/
const databaseColumnsAll = [
    'A', // index 0 - No
    'B', // index 1 - 日付
    'C', // index 2 - 会場名
    'D', // index 3 - 市区町村 【関数】
    'E', // index 4 - イベント名
    'F', // index 5 - 前売料金（A）
    'G', // index 6 - 前売料金（B）
    'H', // index 7 - 前売料金（C）
    'I', // index 8 - 前売配信料金
    'J', // index 9 - 当日料金（A）
    'K', // index 10 - 当日料金（B）
    'L', // index 11 - 当日料金（C）
    'M', // index 12 - 当日配信料金
    'N', // index 13 - 前売り備考
    'O', // index 14 - 当日備考
    'P', // index 15 - ドリンク代
    'Q', // index 16 - 部数
    'R', // index 17 - メモ
    'S', // index 18 - アーティスト名
    'T', // index 19 - RK 【関数】
    'U', // index 20 - 前売集客（A）
    'V', // index 21 - 前売集客（B）
    'W', // index 22 - 前売集客（C）
    'X', // index 23 - 前売集客（配信）
    'Y', // index 24 - 当日集客（A）
    'Z', // index 25 - 当日集客（B）
    'AA', // index 26 - 当日集客（C）
    'AB', // index 27 - 当日集客（配信）
    'AC', // index 28 - 出演料固定
    'AD', // index 29 - 出演料歩合（CB・%）
    'AE', // index 30 - 出演料歩合（CB・定額）
    'AF', // index 31 - 出演料歩合（配信・％）
    'AG', // index 32 - 出演料歩合（配信・定額）
    'AH', // index 33 - その他出演料
    'AI', // index 34 - 前売来場（A）収益 【関数】
    'AJ', // index 35 - 前売来場（B）収益 【関数】
    'AK', // index 36 - 前売来場（C）収益 【関数】
    'AL', // index 37 - 前売配信収益 【関数】
    'AM', // index 38 - 当日来場（A）収益 【関数】
    'AN', // index 39 - 当日来場（B）収益 【関数】
    'AO', // index 40 - 当日来場（C）収益 【関数】
    'AP', // index 41 - 当日配信収益 【関数】
    'AQ', // index 42 - その他収益
    'AR', // index 43 - その他収益備考
    'AS', // index 44 - 収益合計 【関数】
    'AT', // index 45 - 固定費用 【関数】
    'AU', // index 46 - 歩合（CB・％）費用 【関数】
    'AV', // index 47 - 歩合（CB・定額）費用 【関数】
    'AW', // index 48 - 歩合（配信・％） 【関数】
    'AX', // index 49 - 歩合（配信・定額） 【関数】
    'AY', // index 50 - その他費用 【関数】
    'AZ', // index 51 - 費用合計 【関数】
    'BA' // index 52 - 収益対費用 【関数】    
];

/* データベースシートの列一覧　関数除く
（データベースから入力シートに値をコピーするときに使用する)
*/
const databaseColumnsExcFormulae = [
    'A', // index 0 - No
    'B', // index 1 - 日付
    'C', // index 2 - 会場名
    '', // index 3 - 市区町村 【関数】
    'E', // index 4 - イベント名
    'F', // index 5 - 前売料金（A）
    'G', // index 6 - 前売料金（B）
    'H', // index 7 - 前売料金（C）
    'I', // index 8 - 前売配信料金
    'J', // index 9 - 当日料金（A）
    'K', // index 10 - 当日料金（B）
    'L', // index 11 - 当日料金（C）
    'M', // index 12 - 当日配信料金
    'N', // index 13 - 前売り備考
    'O', // index 14 - 当日備考
    'P', // index 15 - ドリンク代
    'Q', // index 16 - 部数
    'R', // index 17 - メモ
    'S', // index 18 - アーティスト名
    '', // index 19 - RK 【関数】
    'U', // index 20 - 前売集客（A）
    'V', // index 21 - 前売集客（B）
    'W', // index 22 - 前売集客（C）
    'X', // index 23 - 前売集客（配信）
    'Y', // index 24 - 当日集客（A）
    'Z', // index 25 - 当日集客（B）
    'AA', // index 26 - 当日集客（C）
    'AB', // index 27 - 当日集客（配信）
    'AC', // index 28 - 出演料固定
    'AD', // index 29 - 出演料歩合（CB・%）
    'AE', // index 30 - 出演料歩合（CB・定額）
    'AF', // index 31 - 出演料歩合（配信・％）
    'AG', // index 32 - 出演料歩合（配信・定額）
    'AH', // index 33 - その他出演料
    '', // index 34 - 前売来場（A）収益 【関数】
    '', // index 35 - 前売来場（B）収益 【関数】
    '', // index 36 - 前売来場（C）収益 【関数】
    '', // index 37 - 前売配信収益 【関数】
    '', // index 38 - 当日来場（A）収益 【関数】
    '', // index 39 - 当日来場（B）収益 【関数】
    '', // index 40 - 当日来場（C）収益 【関数】
    '', // index 41 - 当日配信収益 【関数】
    'AQ', // index 42 - その他収益
    'AR', // index 43 - その他収益備考
    '', // index 44 - 収益合計 【関数】
    '', // index 45 - 固定費用 【関数】
    '', // index 46 - 歩合（CB・％）費用 【関数】
    '', // index 47 - 歩合（CB・定額）費用 【関数】
    '', // index 48 - 歩合（配信・％） 【関数】
    '', // index 49 - 歩合（配信・定額） 【関数】
    '', // index 50 - その他費用 【関数】
    '', // index 51 - 費用合計 【関数】
    '' // index 52 - 収益対費用 【関数】   
];



/* データベースシートの列一覧　関数セルのindex
上記のデータベースシートの列一覧に対応する。
*/
const databaseFormulaeIndex = [
    3, // index 3 - 市区町村 【関数】
    19, // index 19 - RK 【関数】
    34, // index 34 - 前売来場（A）収益 【関数】
    35, // index 35 - 前売来場（B）収益 【関数】
    36, // index 36 - 前売来場（C）収益 【関数】
    37, // index 37 - 前売配信収益 【関数】
    38, // index 38 - 当日来場（A）収益 【関数】
    39, // index 39 - 当日来場（B）収益 【関数】
    40, // index 40 - 当日来場（C）収益 【関数】
    41, // index 41 - 当日配信収益 【関数】
    44, // index 44 - 収益合計 【関数】
    45, // index 45 - 固定費用 【関数】
    46, // index 46 - 歩合（CB・％）費用 【関数】
    47, // index 47 - 歩合（CB・定額）費用 【関数】
    48, // index 48 - 歩合（配信・％） 【関数】
    49, // index 49 - 歩合（配信・定額） 【関数】
    50, // index 50 - その他費用 【関数】
    51, // index 51 - 費用合計 【関数】
    52 // index 52 - 収益対費用 【関数】   
]



/*　「登録」ボタンが押されたときのfunction。
 */
function registerDataToDatabasesheet() {

    var entrySheet = thisSpreadsheet.getSheetByName(dataEntrySheetName);

    var databaseSheet = thisSpreadsheet.getSheetByName(databaseSheetName);

    var entryValueArray = getEntrySheetValues(entrySheet);

    var idInEntrySheet = entryValueArray[0];

    Logger.log("idInEntrySheet=" + idInEntrySheet);

    var rowNumToEnter = 0;

    // エントリーデータにIDなし。新IDを付ける。
    if (idInEntrySheet == '') {

        var newId = getNextId();
        entryValueArray[0] = newId;

        var currentLastRow = databaseSheet.getLastRow();

        rowNumToEnter = currentLastRow + 1;
    }
    // エントリーデータにIDあり。既存データかどうかチェックし、データを更新する。
    else {
        rowNumToEnter = getRowOfIdInDatabase(idInEntrySheet);

        // もしIDがデータベースから見つからなければ、新しい行に入れる。IDはそのまま。
        if (rowNumToEnter === 0) {

            var currentLastRow = databaseSheet.getLastRow();
            rowNumToEnter = currentLastRow + 1;
        }
    }

    Logger.log("rowNumToEnter=" + rowNumToEnter);

    // データをデータベースに入れる

    var cellRange = databaseColumnsAll[0] + rowNumToEnter + ":" + databaseColumnsAll[databaseColumnsAll.length - 1] + rowNumToEnter;
    Logger.log('cellRange:' + cellRange);

    range = databaseSheet.getRange(cellRange);
    var nestedEntryValueArray = [entryValueArray];
    range.setValues(nestedEntryValueArray);

}



/*　「参照」ボタンが押されたときのfunction。
 */
function pullDataFromDatabasesheet() {

    var entrySheet = thisSpreadsheet.getSheetByName(dataEntrySheetName);

    var databaseSheet = thisSpreadsheet.getSheetByName(databaseSheetName);

    var idInEntrySheet = getCellValue(entrySheet, dataEntryCells[0]);

    var databaseRowNum = getRowOfIdInDatabase(idInEntrySheet);

    var databaseValueArray = getDatabaseSheetValues(databaseSheet, databaseRowNum);


    databaseValueArray.forEach((databaseCellValue, index) => {

        // 関数の入っているセルは飛ばすので、関数Indexのarrayに同じindexがないことをチェックする。
        if (databaseFormulaeIndex.indexOf(index) < 0) {
            var cellRange = dataEntryCells[index];

            if (index == 16) {
                databaseCellValue = databaseCellValue + "部";
            }
            var cell = entrySheet.getRange(cellRange);
            cell.setValue(databaseCellValue);
        }

    });
}



/*　「削除」ボタンが押されたときのfunction。
 */
function deleteRowFromDatabasesheet() {

    var entrySheet = thisSpreadsheet.getSheetByName(dataEntrySheetName);

    var databaseSheet = thisSpreadsheet.getSheetByName(databaseSheetName);

    var idInEntrySheet = getCellValue(entrySheet, dataEntryCells[0]);

    Logger.log("idInEntrySheet=" + idInEntrySheet);

    Logger.log("getRowOfIdInDatabase(idInEntrySheet)=" + getRowOfIdInDatabase(idInEntrySheet));

    databaseSheet.deleteRow(getRowOfIdInDatabase(idInEntrySheet));
}



/*　エントリーシートのデータを取得する。
 */
function getEntrySheetValues(sheet) {
    var valueArray = [];

    dataEntryCells.forEach((entryCell, index) => {
        var value = getCellValue(sheet, entryCell);

        // エントリーシートの【index 16 - 部数】を数字だけにする
        if (index == 16) {
            value = value.replace("部", "");
        }

        valueArray.push(value);
    });

    valueArray.forEach((valueContents, index) => {
        Logger.log('valueContents index=' + index + ": " + valueContents);
    });

    return valueArray;
}



/*　データベースシートのデータを取得する。
 */
function getDatabaseSheetValues(sheet, rowNum) {
    var valueArray = [];

    databaseColumnsExcFormulae.forEach((databaseColumn, index) => {

        var cell = databaseColumn + rowNum;

        var value = '';

        if (databaseColumn !== '') {
            value = getCellValue(sheet, cell);
        }

        valueArray.push(value);
    });

    valueArray.forEach((valueContents, index) => {
        Logger.log('valueContents index=' + index + ": " + valueContents);
    });

    return valueArray;
}



/* シート名とセルを指定して、値を返す
*/
function getCellValue(sheet, cell) {
    var value = sheet.getRange(cell).getValue();
    return value;
}



/*　現在の一番大きいIDを探し、次のIDを返す
 */
function getNextId() {
    var column = 0;
    var HEADER_ROW_COUNT = 2;

    var worksheet = thisSpreadsheet.getSheetByName(databaseSheetName);
    var rows = worksheet.getDataRange().getNumRows();
    var vals = worksheet.getSheetValues(1, 1, rows, 1);
    // getSheetValues(startRow, startColumn, numRows, numColumns)

    var max = 0;

    for (var row = HEADER_ROW_COUNT; row < vals.length; row++) {
        var id = vals[row][column];
        if (id > max) {
            max = id;
        }
    }

    Logger.log("getNextId return:" + (max + 1));
    return max + 1;
}



/*　IDがデータベースシートの何行目かを検索して返す
 */
function getRowOfIdInDatabase(id) {

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(databaseSheetName);
    var column = 1; //column Index   
    var columnValues = sheet.getRange(2, column, sheet.getLastRow()).getValues(); //First 2 rows are header rows

    Logger.log("getRowOfIdInDatabase sheet.getLastRow()=" + sheet.getLastRow());
    var rowNum = 0;

    for (i = 0; i < columnValues.length; i++) {
        if (columnValues[i][0] == id) {
            rowNum = i + 2;
        }
    }

    Logger.log("search ID:" + id + " Returned rowNum:" + rowNum);

    return rowNum;
}

