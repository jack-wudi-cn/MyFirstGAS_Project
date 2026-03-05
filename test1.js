function testBasicOperations() {
  // 1. 获取活跃表格和工作表
  // from vscode
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 安全检查：如果没有打开的表格，直接停止，防止报错
  if (!ss) {
    Logger.log("错误：未找到活动的电子表格。请确保脚本绑定在正确的表格上。");
    return;
  }
  
  const sheet = ss.getActiveSheet(); 
  
  // 2. 写入数据
  sheet.getRange("A1").setValue("Hello Google Sheets!"); 
  
  // 3. 读取数据
  const value = sheet.getRange("A1").getValue();
  Logger.log("读取到的值: " + value); 
  
  // 4. 【关键修改】安全地处理 UI 弹窗
  try {
    // 尝试获取 UI 对象
    const ui = SpreadsheetApp.getUi();
    // 如果成功，则弹出对话框
    ui.alert("提示", "操作已完成！\n读取值: " + value, ui.ButtonSet.OK);
  } catch (e) {
    // 如果失败（比如在自动触发器中运行），捕获错误并只记录日志，不中断程序
    Logger.log("注意：当前环境不支持弹窗 (getUi)，已跳过弹窗步骤。错误详情：" + e.message);
    
    // 可选：如果你希望在非交互模式下也能看到结果，可以用 Logger 代替弹窗
    Logger.log("=== 任务完成 ===");
  }
}