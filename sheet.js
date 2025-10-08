class Formulation{
  constructor(sheetName = null, headersRow=3) {
    if(sheetName){
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    }
    else{
      this.sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    }
    this.headers = this.getHeaders(headersRow);
  }

  getSheet(){
    return this.sheet;
  }

  getSheetName(){
    return this.sheet.getSheetName();
  }

  countColumns(){
    return this.sheet.getLastColumn();
  }

  countRows(){
    return;
  }

  getRow(rowNumber){
    const numColumns = this.countColumns(); 
    return this.sheet.getRange(rowNumber, 1, 1, numColumns).getValues()[0];
  }

getHeaders(headersRow=1) {
    const lastColumn = this.sheet.getLastColumn();
    if (lastColumn === 0) {
        return [];
    }
    const headers = this.sheet.getRange(headersRow, 1, 1, lastColumn).getValues()[0];
    console.log(headers)
    return headers;
}

  getColumnIndexByHeader(header){
    const index = this.headers.indexOf(header);
    if(index == -1){
      throw new Error(`La columna '${header}' no se encontró.`);
    }
    return index;
  }

  getCellValue(columnName, rowNumber){
    const columIndex = this.getColumnIndexByHeader(columnName);
    const row = this.getRow(rowNumber);
    return row[columIndex];
  }

  setCellValue(columnName, rowNumber, value) {
    try {
        // Get column index from the mapped header name
        const columnIndex = this.getColumnIndexByHeader(columnName);
        
        // Validate inputs
        if (columnIndex === -1) {
            throw new Error(`Column ${columnName} not found`);
        }
        if (rowNumber < 1) {
            throw new Error('Row number must be greater than 0');
        }

        // Set the value (+1 because Sheets is 1-based)
        this.sheet.getRange(rowNumber, columnIndex + 1).setValue(value);
    } catch (error) {
        console.error(`Error setting cell value: ${error.message}`);
        throw error;
    }
}

  getSelectedRows(){
    const rangeList = this.sheet.getActiveRangeList();
    if (!rangeList) {
      throw new Error(`No hay filas seleccionadas.`);
    }

    const ranges = rangeList.getRanges();
    const selectedRows = new Set();

    ranges.forEach(range => {
      const startRow = range.getRow();
      const numRows = range.getNumRows();
      for (let i = 0; i < numRows; i++) {
        selectedRows.add(startRow + i);
      }
    });

    const selectedRowNumbers = Array.from(selectedRows).sort((a, b) => a - b);
    console.info("selectedRowNumbers "+selectedRowNumbers)
    return selectedRowNumbers;
  }

  getRowContent(rowNumber) {
    const row = this.getRow(rowNumber);
    const rowContent = {};
    
    this.headers.forEach((header, index) => {
      rowContent[header] = row[index];
    });
    
    return JSON.stringify(rowContent, null, 2); 
  }
  
  createColumn(requiredColumns){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = this.getHeaders(3);
  requiredColumns.forEach((colName) => {
    if (!headers.includes(colName)) {
      sheet.getRange(3, headers.length + 1).setValue(colName);
      headers.push(colName);
    }
  });
  }
  
saveResultToSheet(jsonResult, rowNumber) {
        const jsonString = jsonResult.copy
            .replace(/```json\n/, '')
            .replace(/\n```/, '')
            .trim();
        const data = jsonString//JSON.parse(jsonString);
        console.log("=========",data);
        this.setCellValue("Generado", rowNumber, data);
}

saveResultToSheetRow(jsonResult, rowNumber, column) {
        //const jsonString = JSON.parse(jsonResult.copy).replace(/```json\n/, '').replace(/\n```/, '').trim();
                // Remove curly braces and split into key-value pairs
        const text = jsonResult.copy
                                .replace(/```json\n/, '')
                                .replace(/```json/g, '')
                                .replace(/```/g, '')
                                .replace(/\n```/, '')
                                .replace(/\n/g, ' ')
                                .replace(/\s+/g, ' ')
                                .trim();

        // Create object from entries
        console.log(typeof text)
        console.log(text)
        const obj = JSON.parse(text);//Object.fromEntries(entries);
    
        console.log("=========",obj);
        this.setCellValue(column, rowNumber, obj[column]);
} 

toString(json){
    // Convert to string if not already a string
    const jsonStr = typeof json === 'string' ? json : JSON.stringify(json);
    
    const jsonString = jsonStr
        .replace(/```json\n/, '')
        .replace(/\n```/, '')
        .trim();
    return jsonString;
}

saveResultToSheetColumn(celda,data) {
    // Save only to B2
    this.sheet.getRange(celda).setValue(data);
}

jsonToFormattedString(jsonData) {
    try {
        if (typeof jsonData === 'string') {
            jsonData = JSON.parse(jsonData);
        }
        
        return jsonData.map(item => {
            return `(${item.Name})\n${item.Description}\nCTA: ${item.CTA}`;
        }).join('\n\n');
    } catch (error) {
        console.error('Error converting JSON:', error);
        return jsonData;
    }
} 
saveExpenses(jsonResult, rowNumber, sheet) {
      //const jsonString = JSON.parse(jsonResult.copy).replace(/```json\n/, '').replace(/\n```/, '').trim();
            // Remove curly braces and split into key-value pairs
    const text = jsonResult.copy
                            .replace(/```json\n/, '')
                            .replace(/\n```/, '')
                            .replace(/\n/g, ' ')
                            .replace(/\s+/g, ' ')
                            .trim();
    const entries = text.split(',').map(pair => {
        const [key, ...valueParts] = pair.split('=');
        // Join value parts back together in case the value contains commas
        const value = valueParts.join('=');
        return [key.trim(), value.trim()];
    });

    // Create object from entries
    const obj = JSON.parse(text);//Object.fromEntries(entries);
    let result = obj;
    try {
        // Extract keys as a list
        let keys = Object.keys(result.Description)
            .join('\n');

        // Convert Description object to formatted string
        let descriptionText = Object.entries(result.Description)
            .map(([key, value]) => {
                return `${key}:\n` +
                       `- Descripción: ${value.Descripción}\n` +
                       `- Cantidad: ${value.Cantidad}\n` +
                       `- Valor Unitario: $${value.ValorUnitario}\n` +
                       `- Total: $${value.Total}`;
            })
            .join('\n\n');

        // Save Keys to a separate column
        sheet.setCellValue('Keys', rowNumber, keys);
        
        // Save Description to second column
        sheet.setCellValue('DescriptionOK', rowNumber, descriptionText);
        
        // Save Presupuesto status to third column
        sheet.setCellValue('Presupuesto', rowNumber, result.Presupuesto);

    } catch (error) {
        console.error('Error saving admin expenses:', error);
        SpreadsheetApp.getUi().alert('Error saving administrative expenses: ' + error.message);
    }
}

}






































