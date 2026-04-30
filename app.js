function doGet(e) {
  const action = e.parameter.action;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (action === 'test') {
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Connected to P3K Stock Manager'
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'getAll') {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1).map(row => ({
      id: row[0],
      name: row[1],
      category: row[2],
      quantity: row[3],
      expiry: row[4],
      notes: row[5],
      lastUpdate: row[6]
    }));
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      data: rows
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: false,
    message: 'Invalid action'
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  try {
    if (data.action === 'add') {
      sheet.appendRow([
        data.data.id,
        data.data.name,
        data.data.category,
        data.data.quantity,
        data.data.expiry,
        data.data.notes,
        data.data.lastUpdate
      ]);
    } else if (data.action === 'update') {
      // Find row by ID and update
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === data.data.id) {
          sheet.getRange(i + 1, 2, 1, 6).setValues([[
            data.data.name,
            data.data.category,
            data.data.quantity,
            data.data.expiry,
            data.data.notes,
            data.data.lastUpdate
          ]]);
          break;
        }
      }
    } else if (data.action === 'delete') {
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (allData[i][0] === data.id) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
