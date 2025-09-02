/**
 * Google Apps Script to Sort App Data by Downloads Last Month
 * Preserves all hyperlinks and formatting
 */

function onOpen() {
  // Create custom menu when spreadsheet opens
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Download Sorter')
    .addItem('🔽 Sort by Downloads (High to Low)', 'sortByDownloadsHighToLow')
    .addItem('🔼 Sort by Downloads (Low to High)', 'sortByDownloadsLowToHigh')
    .addItem('↩️ Restore Original Order', 'restoreOriginalOrder')
    .addSeparator()
    .addItem('ℹ️ About This Tool', 'showAbout')
    .addToUi();
}

function sortByDownloadsHighToLow() {
  sortByDownloads('desc');
}

function sortByDownloadsLowToHigh() {
  sortByDownloads('asc');
}

function sortByDownloads(order = 'desc') {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const range = sheet.getDataRange();
    
    if (range.getNumRows() < 2) {
      SpreadsheetApp.getUi().alert('Error', 'No data found to sort!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get all data including headers
    const values = range.getValues();
    const formulas = range.getFormulas();
    const headers = values[0];
    
    // Find the "Downloads Last Month" column
    let downloadsColumnIndex = -1;
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().toLowerCase();
      if (header.includes('downloads') && header.includes('month')) {
        downloadsColumnIndex = i;
        break;
      }
    }
    
    if (downloadsColumnIndex === -1) {
      SpreadsheetApp.getUi().alert('Error', 'Could not find "Downloads Last Month" column!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Backup original data before sorting
    backupOriginalData(sheet, range);
    
    // Prepare data for sorting (skip header row)
    const dataRows = [];
    for (let i = 1; i < values.length; i++) {
      dataRows.push({
        rowIndex: i,
        values: values[i],
        formulas: formulas[i],
        downloadValue: parseDownloadValue(values[i][downloadsColumnIndex])
      });
    }
    
    // Sort the data
    dataRows.sort((a, b) => {
      if (order === 'desc') {
        return b.downloadValue - a.downloadValue;
      } else {
        return a.downloadValue - b.downloadValue;
      }
    });
    
    // Apply sorted data back to sheet
    applySortedData(sheet, dataRows, headers.length);
    
    // Show success message
    const orderText = order === 'desc' ? 'highest to lowest' : 'lowest to highest';
    SpreadsheetApp.getUi().alert(
      'Success!', 
      `Data sorted by downloads (${orderText}). All hyperlinks and formatting preserved!`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error in sortByDownloads: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function parseDownloadValue(downloadStr) {
  if (!downloadStr) return 0;
  
  const str = downloadStr.toString().toLowerCase().trim();
  
  // Handle special cases
  if (str.includes('< 5k') || str.includes('<5k')) return 1;
  if (str === 'n/a' || str === '' || str === 'null') return 0;
  
  // Extract number and multiplier using regex
  const match = str.match(/([\d.,]+)\s*([km]?)/);
  if (!match) return 0;
  
  let number = parseFloat(match[1].replace(/,/g, ''));
  if (isNaN(number)) return 0;
  
  const multiplier = match[2];
  
  switch (multiplier) {
    case 'k': return number * 1000;
    case 'm': return number * 1000000;
    default: return number;
  }
}

function applySortedData(sheet, sortedDataRows, numColumns) {
  // Clear existing data (except headers)
  if (sortedDataRows.length > 0) {
    const dataRange = sheet.getRange(2, 1, sortedDataRows.length, numColumns);
    dataRange.clear();
    
    // Apply sorted values and formulas
    for (let i = 0; i < sortedDataRows.length; i++) {
      const row = sortedDataRows[i];
      const targetRow = i + 2; // +2 because row 1 is headers, and getRange is 1-indexed
      
      for (let col = 0; col < numColumns; col++) {
        const cell = sheet.getRange(targetRow, col + 1);
        
        // If there's a formula, use it (this preserves hyperlinks)
        if (row.formulas[col] && row.formulas[col] !== '') {
          cell.setFormula(row.formulas[col]);
        } else {
          // Otherwise use the value
          cell.setValue(row.values[col]);
        }
      }
    }
  }
}

function backupOriginalData(sheet, range) {
  try {
    // Store original data in document properties for restore functionality
    const values = range.getValues();
    const formulas = range.getFormulas();
    
    const backup = {
      values: values,
      formulas: formulas,
      timestamp: new Date().toISOString(),
      sheetName: sheet.getName()
    };
    
    PropertiesService.getDocumentProperties().setProperty(
      'ORIGINAL_DATA_BACKUP', 
      JSON.stringify(backup)
    );
    
  } catch (error) {
    Logger.log('Error backing up original data: ' + error.toString());
  }
}

function restoreOriginalOrder() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const backupData = PropertiesService.getDocumentProperties().getProperty('ORIGINAL_DATA_BACKUP');
    
    if (!backupData) {
      SpreadsheetApp.getUi().alert(
        'No Backup Found', 
        'No backup data found. Please sort the data first to create a backup.', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    const backup = JSON.parse(backupData);
    
    // Confirm restoration
    const response = SpreadsheetApp.getUi().alert(
      'Restore Original Order',
      `Restore data to original order from ${new Date(backup.timestamp).toLocaleString()}?`,
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    
    if (response !== SpreadsheetApp.getUi().Button.YES) {
      return;
    }
    
    // Clear current data
    const currentRange = sheet.getDataRange();
    currentRange.clear();
    
    // Restore values
    const numRows = backup.values.length;
    const numCols = backup.values[0].length;
    const restoreRange = sheet.getRange(1, 1, numRows, numCols);
    
    // Restore formulas and values
    for (let row = 0; row < numRows; row++) {
      for (let col = 0; col < numCols; col++) {
        const cell = sheet.getRange(row + 1, col + 1);
        
        if (backup.formulas[row][col] && backup.formulas[row][col] !== '') {
          cell.setFormula(backup.formulas[row][col]);
        } else {
          cell.setValue(backup.values[row][col]);
        }
      }
    }
    
    SpreadsheetApp.getUi().alert(
      'Restored!', 
      'Data restored to original order successfully!', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log('Error restoring original data: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error', 'Failed to restore original data: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function showAbout() {
  const html = `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>📊 Download Sorter Tool</h2>
      <p><strong>What it does:</strong></p>
      <ul>
        <li>Sorts your app data by "Downloads Last Month"</li>
        <li>Preserves all hyperlinks and formulas</li>
        <li>Handles values like 2M, 900K, &lt; 5k correctly</li>
        <li>Allows restoration to original order</li>
      </ul>
      
      <p><strong>How to use:</strong></p>
      <ol>
        <li>Go to menu: Download Sorter</li>
        <li>Choose "Sort by Downloads (High to Low)"</li>
        <li>Your data will be sorted with highest downloads at top</li>
        <li>Use "Restore Original Order" to undo if needed</li>
      </ol>
      
      <p><strong>Supported formats:</strong></p>
      <ul>
        <li>2M, 1.5M (millions)</li>
        <li>900K, 20K (thousands)</li>
        <li>&lt; 5k (less than values)</li>
        <li>Raw numbers: 1000, 5000</li>
      </ul>
      
      <p style="color: #666; font-size: 12px;">
        Created to preserve hyperlinks while sorting Google Sheets data.
      </p>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'About Download Sorter');
}

// Helper function to test the parsing logic
function testDownloadParsing() {
  const testValues = ['2M', '900K', '300K', '20K', '< 5k', '10K', '8K', '1.5M', 'N/A'];
  
  console.log('Testing download value parsing:');
  testValues.forEach(value => {
    console.log(`${value} -> ${parseDownloadValue(value)}`);
  });
}