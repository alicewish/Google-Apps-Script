/**
 * The getGoogleSpreadsheetAsExcel() method will convert the current Google Spreadsheet to Excel XLSX format and then emails the file as an attachment to the specified user.
 */
function getGoogleSpreadsheetAsExcel() {

    try {

        var ss = SpreadsheetApp.getActive();

        var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + ss.getId() + "&exportFormat=xlsx";

        var params = {
            method: "get",
            headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
            muteHttpExceptions: true
        };

        var blob = UrlFetchApp.fetch(url, params).getBlob();

        blob.setName(ss.getName() + ".xlsx");

        MailApp.sendEmail("amit@labnol.org", "Google Sheet to Excel", "The XLSX file is attached", {attachments: [blob]});

    } catch (f) {
        Logger.log(f.toString());
    }
}
