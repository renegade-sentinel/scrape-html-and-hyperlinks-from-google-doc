// Get spreadsheet
var sheet = SpreadsheetApp.openById('1LsDAmU_ROs0R4RnQDU5MvbxObSfy4cLLIzHn5wPiLn0');
var dbSheet = sheet.getSheetByName('Internal Links')
var htmlSheet = sheet.getSheetByName('Doc HTML')
var currentDateinMonthDayYear = currentDateinMonthDayYear()

function currentDateinMonthDayYear() {
  const currentDate = new Date();
  const formattedDate = (currentDate.getMonth() + 1).toString().padStart(2, '0') + '/' + currentDate.getDate().toString().padStart(2, '0') + '/' + currentDate.getFullYear();
  return formattedDate
}

var links = ["https://docs.google.com/document/d/1Qk8IYoFM2W1KTATXMIa8-HC9CISyU49D7E4Q0oXKAr4/edit"]

function scrapeDocsAndAddInternalLinkTargets() {
  // clearRange(dbSheet)
  // addHeadersToSheet(dbSheet,currentDateinMonthDayYear,"Doc Name", "Doc Link", "Anchor", "Href")
  removeExtraColumns(dbSheet)
  removeExtraRows(dbSheet)
  for (var i = 0; i < links.length; i++) {
    var doc = DocumentApp.openByUrl(links[i]);
    var docId = doc.getId();
    var html = getContent(docId)
    var anchors = parseHtmlAnchors(html)[0]
    var urls = parseHtmlAnchors(html)[1]
    for (var z = 0; z < anchors.length; z++) {
      appendRowToSheet(dbSheet, currentDateinMonthDayYear, doc.getName(), doc.getUrl(), anchors[z], urls[z])
    }
  }
  removeAnchorDuplicates()
  removeExtraRows(dbSheet)
}

function scrapeHTMLAndAddHTMLToSheet() {
  clearRange(htmlSheet)
  addHeadersToSheet(htmlSheet, currentDateinMonthDayYear, "Doc Name", "Doc Link", "HTML")
  removeExtraColumns(htmlSheet)
  removeExtraRows(htmlSheet)
  for (var i = 0; i < links.length; i++) {
    var doc = DocumentApp.openByUrl(links[i]);
    var docId = doc.getId();
    var html = getContent(docId)
    appendRowToSheet(htmlSheet, currentDateinMonthDayYear, doc.getName(), doc.getUrl(), html)
  }
}

/**
 * Returns the HTML content of a Google Doc.
 *
 * @param {string} docId - The ID of the Google Doc.
 * @param {boolean} [useCaching=true] - Enable or disable caching.
 * @return {string} The HTML content of the Google Doc.
 * @throws {Error} If the docId is not provided or is invalid.
 * @throws {Error} If the useCaching parameter is provided and is not a boolean.
 */
function getContent(docId, useCaching = true) {
  if (!docId) {
    throw new Error("Please call this API with a valid Google Doc ID.");
  }

  if (typeof useCaching !== "boolean") {
    throw new Error("If you're going to specify useCaching, it must be a boolean.");
  }

  const cache = CacheService.getScriptCache();
  const cached = cache.get(docId);
  let html;

  if (cached && useCaching) {
    html = cached;
    console.log("Pulling doc html from cache...");
  } else {
    console.log("Grabbing and parsing fresh html from the doc...");

    try {
      const doc = DriveApp.getFileById(docId);
      const docName = doc.getName();
      const forDriveScope = DriveApp.getStorageUsed();
      const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${docId}&exportFormat=html`;
      const options = {
        method: "get",
        headers: { "Authorization": `Bearer ${ScriptApp.getOAuthToken()}` },
        muteHttpExceptions: true,
      };

      html = UrlFetchApp.fetch(url, options).getContentText();
      html = html.replace(/<head>.*<\/head>/, "");
      html = html.replace(/ (id|class|style|start|colspan|rowspan)="[^"]*"/g, "");
      html = html.replace(/<(span|\/span|body|\/body|html|\/html)>/g, "");
      return html
    }
    catch (err) {
      throw err.message;
    }
  }
}

/**
* @param {SpreadsheetApp.Sheet} sheetName
* @param {string[]} cols
*/
function addHeadersToSheet(sheetName, ...cols) {
  sheetName.appendRow([...cols]);
}

/**
* @param {SpreadsheetApp.Sheet} htmlSheet
* @param {string[]} cols
*/
function appendRowToSheet(htmlSheet, ...cols) {
  htmlSheet.appendRow([...cols]);
}

function clearRange(sheetName) {
  sheetName.getRange("A1:D").clearContent();
}

function parseHtmlAnchors(html) {
  var regex = /<a.*?>(.*?)<\/a>/g;
  var matches = html.match(regex);
  var anchors = [];
  var hrefs = [];
  for (var i = 0; i < matches.length; i++) {
    let anchor = matches[i].replace(/<.*?>/g, "")
    if (extractUrl(matches[i]) && anchor) {
      if (anchor != '&nbsp;') {
        if (exactMatchesAnyNegativeKeyword(anchor)) {
          if (containsAnyNegativeKeyword(anchor)) {
            anchor = replaceHTMLEntities(anchor)
            anchors.push(anchor);
            hrefs.push(extractUrl(matches[i]))
          }
        }
      }
    }
    continue
  }
  return [anchors, hrefs]
}

function extractUrl(input) {
  if (input.includes('https://www.google.com/url?q=') && input.includes('www.youraspire.com')) {
    var extractedUrl = input.match(/q=(.*?)&/)[1];
    return extractedUrl;
  } else {
    return ''
  }
}

function removeAnchorDuplicates() {
  dbSheet.getDataRange().removeDuplicates([4])
}

// Remove All Extra Columns In One Sheet 
function removeExtraColumns(sheetName) {
  var maxColumns = sheetName.getMaxColumns();
  var lastColumn = sheetName.getLastColumn();
  if (maxColumns - lastColumn != 0) {
    sheetName.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }

}

//Remove All Extra Rows in One Sheet 
function removeExtraRows(sheetName) {
  var maxRows = sheetName.getMaxRows();
  var lastRow = sheetName.getLastRow();
  if (maxRows - lastRow != 0) {
    sheetName.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

function replaceHTMLEntities(inputString) {
  var outputString = inputString.replace(/&nbsp;/g, "");
  outputString = outputString.replace(/&amp;/g, "&");
  outputString = outputString.replace(/&rsquo;/g, "'");
  outputString = outputString.replace(/&#39;/g, "'");
  return outputString;
}

function exactMatchesAnyNegativeKeyword(str) {
  let negativeKeywords = ["L", "C", "RM", "landscape", "Flexible s", "ERP integrations for popular b", "cheduling", "Case Study", "s all-in-one", "Management", "software", "ookkeeping", "anscape business", "software", "form"]
  for (let i = 0; i < negativeKeywords.length; i++) {
    const e = negativeKeywords[i];
    if (e == str) return false
  }
  return true
}
function containsAnyNegativeKeyword(str) {
  let negativeKeywords = ["softwater"]
  for (let i = 0; i < negativeKeywords.length; i++) {
    const e = negativeKeywords[i];
    if (e.includes(str)) return false
  }
  return true
}
