// 这是一个简单的触发器函数，当用户编辑表格时自动运行
function onEdit(e) {
  // e.range 获取编辑的单元格，e.value 获取新值
  const range = e.range;
  const sheet = range.getSheet();
  
  // 场景：如果在 "订单" 表的 C 列输入 "完成"，自动在 D 列填入时间
  if (sheet.getName() === "订单" && range.getColumn() === 3 && e.value === "完成") {
    // 偏移一行一列写入时间
    range.offset(0, 1).setValue(new Date());
  }
}