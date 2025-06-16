<!DOCTYPE html>
<html>
  <head>
    <base target="_top">
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      #date-container { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
      label { display: block; margin-bottom: 5px; font-weight: bold; }
      input[type="date"] { padding: 8px; border: 1px solid #ccc; border-radius: 4px; width: 130px; }
      button { background-color: #4285F4; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-size: 14px; width: 100%; }
      button:disabled { background-color: #ccc; cursor: not-allowed; }
      #status { margin-top: 15px; text-align: center; font-style: italic; color: #555; }
      #results-container a { display: block; text-align: center; margin-top: 10px; padding: 10px; color: white; text-decoration: none; border-radius: 4px; font-weight: bold; }
      .csv-link { background-color: #34A853; } /* Green */
      .archive-link { background-color: #FBBC05; } /* Yellow */
    </style>
  </head>
  <body>
    <div id="date-container">
      <div>
        <label for="start-date">Start Date</label>
        <input type="date" id="start-date">
      </div>
      <div>
        <label for="end-date">End Date</label>
        <input type="date" id="end-date">
      </div>
    </div>
    
    <button id="process-button" onclick="processAll()">Process & Download</button>
    <div id="status"></div>
    <div id="results-container"></div>

    <script>
      document.getElementById('end-date').valueAsDate = new Date();
      let sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
      document.getElementById('start-date').valueAsDate = sevenDaysAgo;

      function processAll() {
        const startDate = document.getElementById('start-date').value;
        const endDate = document.getElementById('end-date').value;
        const button = document.getElementById('process-button');
        const statusDiv = document.getElementById('status');

        if (!startDate || !endDate) {
          statusDiv.textContent = "Please select both a start and end date.";
          return;
        }

        button.disabled = true;
        button.textContent = 'Processing...';
        statusDiv.textContent = 'This may take a minute...';

        google.script.run
          .withSuccessHandler(onSuccess)
          .withFailureHandler(onFailure)
          .processTrades(startDate, endDate);
      }

      function onSuccess(result) {
        const statusDiv = document.getElementById('status');
        const resultsContainer = document.getElementById('results-container');
        document.getElementById('process-button').style.display = 'none';

        if (result.status === 'NO_TRADES_FOUND') {
          statusDiv.textContent = 'No trades were found in the selected date range.';
          return;
        }

        statusDiv.textContent = `Success! Appended ${result.tradeCount} trades to your sheet.`;

        // Create CSV Download Link
        const blob = new Blob([result.csvData], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const csvLink = document.createElement("a");
        const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, "-");
        csvLink.setAttribute("href", url);
        csvLink.setAttribute("download", `Fidelity_Report_${timestamp}.csv`);
        csvLink.textContent = "Download CSV Report";
        csvLink.classList.add('csv-link');
        resultsContainer.appendChild(csvLink);
        
        // Create Archive Link
        const archiveLink = document.createElement("a");
        archiveLink.setAttribute("href", result.archiveUrl);
        archiveLink.setAttribute("target", "_blank"); // Open in new tab
        archiveLink.textContent = "Open Archive Sheet";
        archiveLink.classList.add('archive-link');
        resultsContainer.appendChild(archiveLink);
      }

      function onFailure(error) {
        document.getElementById('status').textContent = 'Error: ' + error.message;
        document.getElementById('process-button').disabled = false;
        document.getElementById('process-button').textContent = 'Process & Download';
      }
    </script>
  </body>
</html>
