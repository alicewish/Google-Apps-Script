// Credit Stéphane Giron
// 将谷歌文档保存为HTML

function exportAsHTML(documentId) {
    var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + documentId + "&exportFormat=html";
    var param = {
        method: "get",
        headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true,
    };
    var html = UrlFetchApp.fetch(url, param).getContentText();
    var file = DriveApp.createFile(documentId + ".html", html);
    return file.getUrl();
}

// Credit: Eric Koleda
// 将谷歌表格导出为Microsoft Excel

function exportAsExcel(spreadsheetId) {
    var file = Drive.Files.get(spreadsheetId);
    var url = file.exportLinks['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'];
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
        headers: {
            'Authorization': 'Bearer ' + token
        }
    });
    return response.getBlob();
}


function runThis() {
    var documentId = "1bB3RGo1tHFTnYgQ9M56d42AtYq9x3fLYVBR4qHdmrzY"
    var forDriveScope = DriveApp.getStorageUsed(); //needed to get Drive Scope requested
    var url = "https://docs.google.com/feeds/download/documents/export/Export?id=" + documentId + "&exportFormat=html";
    var param = {
        method: "get",
        headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true,
    };
    var html = UrlFetchApp.fetch(url, param).getContentText();
    var file = DriveApp.createFile(documentId + ".html", html);
    return file.getUrl();
}

//测试功能gapps upload