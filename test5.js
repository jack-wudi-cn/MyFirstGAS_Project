function onOpen() {
//   const ui = SpreadsheetApp.getUi();
//   ui.createMenu('🚀 我的工具箱')
//       .addItem('运行数据清洗', 'batchDataProcessing')
//       .addItem('获取币价', 'fetchCryptoPrice')
//       .addSeparator()
//       .addItem('关于本工具', 'showAbout')
//       .addToUi();
}

function showAbout() {
  SpreadsheetApp.getUi().alert("版本: 1.0\n开发者: 前 VBA 专家");
}