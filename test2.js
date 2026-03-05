function batchDataProcessing() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 1. 确定数据范围 (假设从 A2 开始，动态获取最后一行)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // 没有数据
  
  // 2. 【关键】一次性读取所有数据到内存 (得到二维数组 [[val], [val], ...])
  // 范围：A2 到 C(lastRow)
  const dataRange = sheet.getRange(2, 1, lastRow - 1, 3); 
  const values = dataRange.getValues(); 
  
  // 3. 在内存中处理 (使用 JS 强大的数组方法，比 VBA 循环快且简洁)
  const processedData = values.map((row, index) => {
    const originalText = row[0]; // A 列
    
    // 业务逻辑：转大写
    const upperText = typeof originalText === 'string' ? originalText.toUpperCase() : originalText;
    
    // 业务逻辑：B 列标记 (假设原 B 列为空，填入处理时间)
    const status = "Processed at " + new Date().toLocaleTimeString();
    
    // 返回新行：[A列新值, B列新值, C列原值保持不变]
    return [upperText, status, row[2]]; 
  });
  
  // 4. 【关键】一次性写回
  // 注意：setValues 要求二维数组的行列数必须与 Range 完全一致
  dataRange.setValues(processedData);
  
  SpreadsheetApp.getUi().alert(`成功处理 ${values.length} 行数据！`);
}