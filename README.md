
# Scrape HTML and Hyperlinks from Google Doc

This is a Google Apps Script that allows you to scrape the HTML and hyperlinks from a Google Doc and append them to a spreadsheet. The script has two separate functions that you can run:

1.  `ScrapeDocHtml()`: Scrapes all the HTML from a Google Doc and appends it as a row in the sheet to the Doc HTML sheet tab.
2.  `ScrapeDocLinks()`: Scrapes all current hyperlinks and anchor text that are in a Google Doc and appends them to the Internal Links sheet tab.

## Getting Started

To get started, follow these steps:

1.  Open the Google Doc that you want to scrape.
2.  Click on **Tools** > **Script editor**.
3.  Copy and paste the code from `scrape-html-and-hyperlinks-from-google-doc.js` into the script editor.
4.  Save the script.
5.  In the script editor, click on **Run** > **ScrapeDocHtml** or **ScrapeDocLinks** to run the desired function.
