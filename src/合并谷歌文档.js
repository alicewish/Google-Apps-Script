/**
 * 在表格启动时运行主函数
 */
function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [
        {name: "合并", functionName: "mergeDocuments"}
    ];
    ss.addMenu("合并文档", menuEntries);
}

/**
 * 合并谷歌文档
 */
function mergeDocuments() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var idsRange = sheet.getRange("A2:A");
    var idsArray = idsRange.getValues();
    // 创建新的谷歌文档
    var mergedDoc = DocumentApp.create("合并文档: " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'"));
    var mergedDocId = mergedDoc.getId();
    // Cycle the ID array
    for(var i = 0; i < idsArray.length; i++) {
        // 查看单元格是否非空
        if(idsArray[i][0] != "" && idsArray[i][0] != undefined) {
            var currentId = idsArray[i][0];
            mergeThem(mergedDocId, currentId);
        }
    }

}

/**
 * Merge two Google Documents
 *
 * @param {string} firstDocumentId the first document id, second document will be merged here
 * @param {string} secondDocumentId the second document id that will be merge to the first document
 */
function mergeThem(firstDocumentId, secondDocumentId) {
    var baseDoc = DocumentApp.openById(firstDocumentId);
    var body = baseDoc.getActiveSection();
    var otherBody = DocumentApp.openById(secondDocumentId).getActiveSection();
    var totalElements = otherBody.getNumChildren();
    for( var j = 0; j < totalElements; ++j ) {
        var element = otherBody.getChild(j).copy();
        var type = element.getType();
        if( type == DocumentApp.ElementType.PARAGRAPH ) {
            body.appendParagraph(element);
        } else if( type == DocumentApp.ElementType.TABLE ) {
            body.appendTable(element);
        } else if( type == DocumentApp.ElementType.LIST_ITEM ) {
            body.appendListItem(element);
        } else {
            throw new Error("According to the doc this type couldn't appear in the body: "+type);
        }
    }

    // 添加空段落以分隔文档
    body.appendParagraph("");

}