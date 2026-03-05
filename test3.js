function fetchMockData() {
  try {
    // 1. 【模拟】构建一个假的 API 响应 (实际开发中这里会是 UrlFetchApp.fetch)
    // 我们手动构造一个 JSON 字符串，模拟从服务器拿到的数据
    const mockJsonString = `{
      "status": "success",
      "data": {
        "id": 888,
        "product_name": "Google Sheets 高级教程",
        "price": 99.50,
        "currency": "CNY",
        "stock": 120,
        "last_updated": "2023-10-27T10:00:00Z"
      }
    }`;
    
    Logger.log("📡 正在请求数据... (模拟)");
    Utilities.sleep(500); // 模拟网络延迟 0.5 秒，让你感觉像在请求
    
    // 2. 【核心】解析 JSON (这和解析真实 API 返回的数据一模一样)
    const jsonResponse = JSON.parse(mockJsonString);
    
    // 3. 【核心】提取数据 (像操作 VBA 的 Dictionary 或 Collection 一样)
    if (jsonResponse.status !== "success") {
      throw new Error("API 返回状态异常");
    }
    
    const data = jsonResponse.data;
    const productName = data.product_name;
    const price = data.price;
    const stock = data.stock;
    
    Logger.log(`✅ 解析成功: ${productName}, 价格: ${price}`);
    
    // 4. 【核心】写入 Sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Mock_API_Test");
    
    if (!sheet) {
      sheet = ss.insertSheet("Mock_API_Test");
      // 初始化表头
      sheet.getRange("A1:E1").setValues([["时间戳", "商品名称", "价格", "库存", "原始 ID"]]);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
    }
    
    // 追加数据
    sheet.appendRow([
      new Date(),
      productName,
      price,
      stock,
      data.id
    ]);
    
    // 5. 弹窗提示
    SpreadsheetApp.getUi().alert(
      "🎉 测试成功！", 
      `已获取数据：\n商品：${productName}\n价格：¥${price}\n库存：${stock}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (e) {
    Logger.log("❌ 发生错误: " + e.toString());
    SpreadsheetApp.getUi().alert("运行失败", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}