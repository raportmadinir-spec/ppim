/**
 * FILE: Database.js
 * FUNGSI: Koneksi dan manipulasi Data Spreadsheet
 */

const DB = {
  getSheet: function(sheetName) {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  },
  
  getAll: function(sheetName) {
    const sheet = this.getSheet(sheetName);
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 
    
    const headers = data.shift(); 
    return data.map(row => {
      let obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index];
      });
      return obj;
    });
  }
};