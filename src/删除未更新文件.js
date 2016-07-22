// Trash every untitled spreadsheet that hasn't been updated in a week.
var files = DriveApp.getFilesByName('Untitled spreadsheet');
while (files.hasNext()) {
    var file = files.next();
    if (new Date() - file.getLastUpdated() > 7 * 24 * 60 * 60 * 1000) {
        file.setTrashed(true);
    }
}/**
 * Created by alicewish on 16/7/17.
 */
