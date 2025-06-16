/**
 * Creates a custom menu in the spreadsheet to run the script.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fidelity Parser')
    .addItem('Process & Download Trades...', 'showDatePickerDialog')
    .addToUi();
}

/**
 * Displays an HTML dialog with date input fields.
 */
function showDatePickerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DatePicker')
    .setWidth(400)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Process Trades by Date');
}

/**
 * The main processing function. It now appends to the active sheet,
 * creates a new archive sheet, and generates a CSV for download.
 */
function processTrades(startDate, endDate) {
  Logger.log(`--- Starting Full Process from ${startDate} to ${endDate} ---`);
  
  const startParts = startDate.split('-').map(Number);
  const endParts = endDate.split('-').map(Number);
  
  const afterDate = new Date(startParts[0], startParts[1] - 1, startParts[2]);
  const beforeDate = new Date(endParts[0], endParts[1] - 1, endParts[2]);
  beforeDate.setDate(beforeDate.getDate() + 1);

  const afterQuery = Utilities.formatDate(afterDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const beforeQuery = Utilities.formatDate(beforeDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  const searchQuery = `label:"Fidelity Trades" after:${afterQuery} before:${beforeQuery}`;
  const threads = GmailApp.search(searchQuery);
  Logger.log(`Found ${threads.length} email threads.`);
  
  const dataToWrite = [];

  threads.forEach(thread => {
    thread.getMessages().forEach(message => {
      const htmlBody = message.getBody();
      if (!htmlBody) return;

      const tradeSectionMatch = htmlBody.match(/Your order to.*?Order Number:.*?<\/td>/is);
      if (!tradeSectionMatch) return;

      let tradeText = tradeSectionMatch[0].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
      const messageDate = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
      
      let tradeData = null;

      const stockRegex = /Your order to (BUY|SELL): ([\d,.]+) shares of ([A-Z]+) was.*?Filled: ([\d,.]+) shares @ \$([\d,.]+)/i;
      const stockMatch = tradeText.match(stockRegex);

      if (stockMatch) {
        tradeData = { transactionType: stockMatch[1], ticker: stockMatch[3], quantity: stockMatch[4], price: stockMatch[5] };
      } else {
        const optionsRegex = /Your order to (BUY|SELL) (PUT|CALL): ([\d,.]+) of -([A-Z0-9]+) was.*?Filled: ([\d,.]+) @ \$([\d,.]+)/i;
        const optionsMatch = tradeText.match(optionsRegex);
        if (optionsMatch) {
          tradeData = { transactionType: `${optionsMatch[1]} ${optionsMatch[2]}`, ticker: optionsMatch[4], quantity: optionsMatch[5], price: optionsMatch[6] };
        }
      }

      if (tradeData) {
        const orderNumberMatch = tradeText.match(/Order Number: (\S+)/i);
        dataToWrite.push([
          messageDate,
          tradeData.transactionType.toUpperCase(),
          tradeData.ticker,
          tradeData.quantity.replace(/,/g, ''),
          tradeData.price.replace(/,/g, ''),
          orderNumberMatch ? orderNumberMatch[1] : 'N/A'
        ]);
      }
    });
  });

  if (dataToWrite.length > 0) {
    // Action 1: Append to the CURRENT active sheet
    const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    activeSheet.getRange(activeSheet.getLastRow() + 1, 1, dataToWrite.length, 6).setValues(dataToWrite);

    // Action 2: Create a NEW Google Sheet for archive
    const ss = SpreadsheetApp.create('Fidelity Trade Archive ' + new Date().toLocaleString());
    const archiveSheet = ss.getActiveSheet();
    archiveSheet.appendRow(['Date', 'Transaction Type', 'Ticker', 'Shares', 'Price', 'Order Number']);
    archiveSheet.getRange(2, 1, dataToWrite.length, 6).setValues(dataToWrite);

    // Action 3: Generate CSV data for download
    const header = ['Date', 'Transaction Type', 'Ticker', 'Shares', 'Price', 'Order Number'];
    let csvContent = header.join(',') + '\n';
    dataToWrite.forEach(row => {
      csvContent += row.map(value => `"${value}"`).join(',') + '\n';
    });

    Logger.log(`Processed ${dataToWrite.length} trades.`);
    return { 
      status: 'SUCCESS',
      csvData: csvContent, 
      archiveUrl: ss.getUrl(),
      tradeCount: dataToWrite.length
    };
  } else {
    return { status: 'NO_TRADES_FOUND' };
  }
}
