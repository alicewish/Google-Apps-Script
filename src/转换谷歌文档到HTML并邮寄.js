/* Credit: https://gist.github.com/erickoledadevrel/11143648 */
/**
 * Sends an email using the contents of a Google Document as the body.
 */
function sendDocument(documentId, recipient, subject) {
    var html = convertToHtml(documentId);
    html = inlineCss(html);
    GmailApp.sendEmail(recipient, subject, null, {
        htmlBody: html
    });
}

/**
 * Converts a file to HTML. The Advanced Drive service must be enabled to use
 * this function.
 */
function convertToHtml(fileId) {
    var file = Drive.Files.get(fileId);
    var htmlExportLink = file.exportLinks['text/html'];
    if (!htmlExportLink) {
        throw 'File cannot be converted to HTML.';
    }
    var oAuthToken = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(htmlExportLink, {
        headers:{
            'Authorization': 'Bearer ' + oAuthToken
        },
        muteHttpExceptions: true
    });
    if (!response.getResponseCode() == 200) {
        throw 'Error converting to HTML: ' + response.getContentText();
    }
    return response.getContentText();
}

/**
 * Inlines CSS within an HTML file using the MailChimp API. To use the API you must
 * register for an account and then copy your API key into the script property
 * "mailchimp.apikey". See here for information on how to find your API key:
 * http://kb.mailchimp.com/article/where-can-i-find-my-api-key/.
 */
function inlineCss(html) {
    var apikey = CacheService.getPublicCache().get('mailchimp.apikey');
    if (!apikey) {
        apikey = PropertiesService.getScriptProperties().getProperty('mailchimp.apikey');
        CacheService.getPublicCache().put('mailchimp.apikey', apikey);
    }
    var datacenter = apikey.split('-')[1];
    var url = Utilities.formatString('https://%s.api.mailchimp.com/2.0/helper/inline-css', datacenter);
    var response = UrlFetchApp.fetch(url, {
        method: 'post',
        payload: {
            apikey: apikey,
            html: html,
            strip_css: true
        }
    });
    var output = JSON.parse(response.getContentText());
    if (!response.getResponseCode() == 200) {
        throw 'Error inlining CSS: ' + output['error'];
    }
    return output['html'];
}
 