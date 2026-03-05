/**
 * 主函数：一键执行汇总与分发 (列宽优化版 - 填充整个页面)
 * 修复点：
 * 1. 调整四列宽度，使总宽度≈750px，填满A4页面
 * 2. 保留所有美化效果（边框、背景色、页脚等）
 * 3. Logo 问题暂不处理（按你要求）
 */
function runWeeklyReportProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  const templateSheetName = "Report_Template";
  
  if (!configSheet || !ss.getSheetByName(templateSheetName)) {
    SpreadsheetApp.getUi().alert("错误", "未找到 'Config' 或 'Report_Template' 工作表！", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastConfigRow = configSheet.getLastRow();
  if (lastConfigRow < 2) {
    SpreadsheetApp.getUi().alert("提示", "Config 表中没有数据。", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const configs = configSheet.getRange(2, 1, lastConfigRow - 1, 3).getValues();
  const ui = SpreadsheetApp.getUi();
  
  let successCount = 0;
  let failCount = 0;
  let errorLog = "";

  Logger.log("开始处理报表任务...");

  // === 【可选】Logo URL（当前不使用，但保留结构）===
  const LOGO_URL = "https://lh3.googleusercontent.com/d/1ABCxyz123"; // ← 可留空或注释

  configs.forEach((row, index) => {
    const regionName = row[0];
    const managerName = row[1];
    const email = row[2];

    if (!regionName || !email || String(email).trim() === "") {
      return; 
    }

    try {
      Logger.log(`正在处理第 ${index + 1} 个区域：${regionName}`);

      const sourceSheet = ss.getSheetByName(regionName);
      if (!sourceSheet) {
        throw new Error(`找不到源数据表：${regionName}`);
      }
      
      const lastRow = sourceSheet.getLastRow();
      if (lastRow < 2) {
        Logger.log(`${regionName} 无数据，跳过。`);
        return; 
      }

      const rawData = sourceSheet.getRange(2, 1, lastRow - 1, 5).getValues();

      let totalSales = 0;
      let totalTarget = 0;
      
      const detailData = rawData.map(r => {
        const sales = Number(r[3]) || 0;
        const target = Number(r[4]) || 0;
        totalSales += sales;
        totalTarget += target;
        return [r[0], r[1], r[2], sales]; 
      });

      const achievementRate = totalTarget > 0 ? (totalSales / totalTarget) : 0;
      
      const startDate = rawData[0][0];
      const endDate = rawData[rawData.length - 1][0];
      const dateRange = `${startDate} 至 ${endDate}`;

      const timestamp = new Date().getTime();
      const tempSheetName = `Temp_${regionName}_${timestamp}`;
      
      const template = ss.getSheetByName(templateSheetName);
      const newSheet = template.copyTo(ss).setName(tempSheetName);

      // === 【关键】立即删除多余列（E~Z），从根源消除右侧空白 ===
      for (let col = 26; col >= 5; col--) {
        newSheet.deleteColumn(col);
      }
      Logger.log(`已删除 E~Z 列，当前最大列为 D 列`);

      // === 填充动态内容 ===
      newSheet.getRange("A1").setValue(`【周报】${regionName} 销售业绩汇报`);
      newSheet.getRange("B3").setValue(dateRange);
      newSheet.getRange("B5").setValue(totalSales);
      newSheet.getRange("B7").setValue(achievementRate);
      
      newSheet.getRange("B5").setNumberFormat("¥#,##0");
      newSheet.getRange("B7").setNumberFormat("0.0%");

      // === 表头设置 ===
      const headerRow = 10;
      const headers = ["日期", "销售员", "产品", "销售额"];
      const headerRange = newSheet.getRange(headerRow, 1, 1, 4);
      headerRange.setValues([headers]);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#4CAF50");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      headerRange.setVerticalAlignment("middle");
      headerRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
      headerRange.setFontSize(11);

      // === 明细数据设置 ===
      let dataStartRow = headerRow + 1;
      if (detailData.length > 0) {
        const dataRange = newSheet.getRange(dataStartRow, 1, detailData.length, 4);
        dataRange.setValues(detailData);
        dataRange.setHorizontalAlignment("center");
        dataRange.setVerticalAlignment("middle");
        dataRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
        dataRange.setFontSize(10);
        
        // === 【核心修改】精确设置每列宽度，填满A4页面 ===
        newSheet.setColumnWidth(1, 180);  // 日期列 → 更宽，适应长日期
        newSheet.setColumnWidth(2, 150);  // 销售员列 → 适中
        newSheet.setColumnWidth(3, 200);  // 产品列 → 最宽，适应长产品名
        newSheet.setColumnWidth(4, 220);  // 销售额列 → 右对齐+货币符号需空间
        
        // 设置行高
        newSheet.setRowHeights(dataStartRow, detailData.length, 22);
      }

      // === 关键指标背景色 ===
      newSheet.getRange("B5").setBackground("#E8F5E9");
      newSheet.getRange("B7").setBackground("#E8F5E9");

      // === 【可选】插入 Logo（当前跳过，避免干扰）===
      // 如果需要，取消下方注释并配置 LOGO_URL
      /*
      try {
        const response = UrlFetchApp.fetch(LOGO_URL, { muteHttpExceptions: true, timeout: 5000 });
        if (response.getResponseCode() === 200) {
          const logoBlob = response.getBlob();
          if (logoBlob && logoBlob.getBytes().length > 0) {
            const image = newSheet.insertImage(logoBlob, 1, 1);
            image.setAnchorCell(newSheet.getRange("A1"));
            image.setPosition(5, 5, 0, 0);
          }
        }
      } catch (e) {
        Logger.log("⚠️ Logo 插入失败：" + e.message);
      }
      */

      // === 添加页脚 ===
      const lastDataRow = detailData.length > 0 ? dataStartRow + detailData.length - 1 : headerRow;
      const footerRow = lastDataRow + 2;
      
      const generatedTime = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
      const footerText = `第 1 页 / 共 1 页 | 生成时间：${generatedTime}`;
      
      newSheet.getRange(footerRow, 4).setValue(footerText);
      newSheet.getRange(footerRow, 4).setFontStyle("italic");
      newSheet.getRange(footerRow, 4).setFontSize(9);
      newSheet.getRange(footerRow, 4).setFontColor("#666666");

      // === 添加整体边框 ===
      const borderRange = newSheet.getRange(1, 1, footerRow, 4);
      borderRange.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);

      // === 清除下方多余行内容 ===
      for (let row = footerRow + 1; row <= 100; row++) {
        newSheet.getRange(row, 1, 1, 4).clearContent();
      }

      // ==========================================
      // 【核心修复】强制刷新与等待
      // ==========================================
      SpreadsheetApp.flush(); 
      Utilities.sleep(2000); 
      Logger.log(`数据已填充并刷新，准备生成 PDF: ${tempSheetName}`);
      // ==========================================

      // --- 步骤 D: 导出 PDF ---
      const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=pdf&gid=${newSheet.getSheetId()}&portrait=true&fitw=true&size=A4&top_margin=0.3&bottom_margin=0.3&left_margin=0.3&right_margin=0.3`;
      
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        throw new Error(`PDF 生成失败 (HTTP ${response.getResponseCode()}) - Response: ${response.getContentText()}`);
      }
      
      const fileName = `${regionName}_周报_${new Date().toLocaleDateString()}.pdf`;
      const pdfBlob = response.getBlob().setName(fileName);

      // --- 步骤 E: 发送邮件 ---
      const subject = `【自动报表】${regionName} 本周销售汇报 (${managerName})`;
      const body = `尊敬的 ${managerName}：\n\n您好！\n\n附件是 ${regionName} 的本周销售详细报表。\n\n📊 关键指标概览：\n- 总销售额：¥${totalSales.toLocaleString()}\n- 目标达成率：${(achievementRate * 100).toFixed(1)}%\n- 统计周期：${dateRange}\n\n请注意查收。\n\n------------------\n系统自动发送，请勿回复。`;
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: body,
        attachments: [pdfBlob]
      });

      Logger.log(`邮件已成功发送至：${email}`);

      // --- 步骤 F: 清理临时 Sheet ---
      ss.deleteSheet(newSheet);
      Logger.log(`临时表 ${tempSheetName} 已删除。`);

      successCount++;

    } catch (e) {
      failCount++;
      const errMsg = `区域 [${regionName}] 处理失败：${e.message}`;
      Logger.log(errMsg);
      Logger.log(e.stack);
      errorLog += `❌ ${errMsg}\n`;
    }
  });

  let alertMsg = `✅ 报表处理完成！\n\n🎉 成功：${successCount} 份\n⚠️ 失败：${failCount} 份`;
  
  if (errorLog) {
    alertMsg += `\n\n--- 错误详情 ---\n${errorLog}`;
  }
  
  ui.alert("📊 报表分发结果", alertMsg, ui.ButtonSet.OK);
  Logger.log("所有任务结束。");
}

/**
 * 辅助功能：打开表格时自动添加自定义菜单
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 报表自动化')
      .addItem('🚀 运行本周报表分发', 'runWeeklyReportProcess')
      .addToUi();
}