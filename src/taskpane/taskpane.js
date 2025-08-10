/* global Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // 初始化事件绑定
    document.getElementById("sendBtn").onclick = sendMessage;
    document.getElementById("userInput").addEventListener("keydown", e => {
      if (e.key === "Enter") sendMessage();
    });
  }
});

const BACKEND_URL = "http://localhost:3001/analyze";

const chat = document.getElementById("chat");
const loading = document.getElementById("loading");
const userInput = document.getElementById("userInput");
const previewArea = document.getElementById("previewArea");
const previewContent = document.getElementById("previewContent");
const previewActions = document.getElementById("previewActions");

function appendMessage(text, from = "bot") {
  const div = document.createElement("div");
  div.className = "message " + (from === "user" ? "userMsg" : "botMsg");
  div.textContent = text;
  chat.appendChild(div);
  chat.scrollTop = chat.scrollHeight;
}

async function sendMessage() {
  const msg = userInput.value.trim();
  if (!msg) return;

  appendMessage(msg, "user");
  userInput.value = "";
  previewArea.style.display = "none";
  previewContent.innerHTML = "";
  previewActions.innerHTML = "";

  loading.style.display = "block";
  disableInput(true);

  try {
    const resp = await fetch(BACKEND_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message: msg }),
    });
    const data = await resp.json();

    loading.style.display = "none";
    disableInput(false);

    appendMessage(data.description || "无响应");

    switch (data.action) {
      case "sort_sales_desc":
        await showSortPreview();
        break;
      case "scatter_sales_costs":
        await showScatterPreview();
        break;
      case "insert_profits":
        await showInsertProfitsPreview();
        break;
      default:
        // 不支持其他操作，隐藏预览区
        previewArea.style.display = "none";
        previewContent.innerHTML = "";
        previewActions.innerHTML = "";
        break;
    }
  } catch (error) {
    loading.style.display = "none";
    disableInput(false);
    appendMessage("请求失败：" + error.message);
  }
}

function disableInput(disable) {
  userInput.disabled = disable;
  document.getElementById("sendBtn").disabled = disable;
}

async function showSortPreview() {
  try {
    const tableData = await getTableData();
    renderTable(tableData);

    previewArea.style.display = "block";
    previewActions.innerHTML = "";

    const btn = document.createElement("button");
    btn.textContent = "应用操作";
    btn.onclick = async () => {
      await applySort();
      alert("排序已应用到Excel表！");
      previewArea.style.display = "none";
    };
    previewActions.appendChild(btn);
  } catch (e) {
    appendMessage("读取Excel表格失败：" + e.message);
  }
}

async function showScatterPreview() {
  try {
    const scatterData = await getScatterData();
    renderScatterPlot(scatterData);

    previewArea.style.display = "block";
    previewActions.innerHTML = "";

    const btn = document.createElement("button");
    btn.textContent = "插入";
    btn.onclick = async () => {
      await insertScatterChart();
      alert("散点图已插入到Excel！");
      previewArea.style.display = "none";
    };
    previewActions.appendChild(btn);
  } catch (e) {
    appendMessage("读取Excel数据失败：" + e.message);
  }
}

async function showInsertProfitsPreview() {
  previewArea.style.display = "block";
  previewActions.innerHTML = "";

  previewContent.innerHTML = `
    <p>即将插入利润列，公式为：<code>=[@Sales]-[@Costs]</code></p>
  `;

  const btn = document.createElement("button");
  btn.textContent = "插入公式";
  btn.onclick = async () => {
    try {
      await insertProfitColumn();
      alert("利润列公式已插入Excel表！");
      previewArea.style.display = "none";
    } catch (e) {
      appendMessage("插入公式失败：" + e.message);
    }
  };
  previewActions.appendChild(btn);
}

// ========== Excel API 具体操作 ==========

// 读取表格，返回表头+数据
async function getTableData() {
  return await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItemAt(0);

    const headerRange = table.getHeaderRowRange();
    const dataRange = table.getDataBodyRange();

    headerRange.load("values");
    dataRange.load("values");
    await context.sync();

    return {
      header: headerRange.values[0],
      rows: dataRange.values,
      table,
      sheet,
    };
  });
}

// 渲染表格到预览区
function renderTable({ header, rows }) {
  let html = "<table><thead><tr>";
  header.forEach(h => {
    html += `<th>${h}</th>`;
  });
  html += "</tr></thead><tbody>";
  rows.forEach(row => {
    html += "<tr>";
    row.forEach(cell => {
      html += `<td>${cell}</td>`;
    });
    html += "</tr>";
  });
  html += "</tbody></table>";
  previewContent.innerHTML = html;
}

// 应用排序（按Sales降序）
async function applySort() {
  await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItemAt(0);

    const columns = table.columns;
    columns.load("items/name");
    await context.sync();

    const salesIndex = columns.items.findIndex(c => c.name === "Sales");
    if (salesIndex === -1) throw new Error("表中没有Sales列");

    table.sort.apply([{ key: salesIndex, ascending: false }]);
    await context.sync();
  });
}


// 获取散点图数据：Sales和Costs两列数字
async function getScatterData() {
  return await Excel.run(async context => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItemAt(0);

    const columns = table.columns;
    columns.load("items/name");
    await context.sync();

    const salesIndex = columns.items.findIndex(c => c.name === "Sales");
    const costsIndex = columns.items.findIndex(c => c.name === "Costs");
    if (salesIndex === -1 || costsIndex === -1) {
      throw new Error("表中缺少 Sales 或 Costs 列");
    }

    const dataRange = table.getDataBodyRange();
    dataRange.load("values");
    await context.sync();

    const series = dataRange.values.map(row => ({
      x: Number(row[salesIndex]) || 0,
      y: Number(row[costsIndex]) || 0,
    }));

    return { series, xLabel: "Sales", yLabel: "Costs", sheet, table };
  });
}

// 渲染散点图预览（简单文本展示）
function renderScatterPlot({ series, xLabel, yLabel }) {
  let html = `<p>散点图预览 (${xLabel} vs ${yLabel}):</p>`;
  html += `<table><thead><tr><th>${xLabel}</th><th>${yLabel}</th></tr></thead><tbody>`;
  series.forEach(p => {
    html += `<tr><td>${p.x}</td><td>${p.y}</td></tr>`;
  });
  html += "</tbody></table>";
  previewContent.innerHTML = html;
}

// 插入Excel散点图表
// 插入 Excel 散点图（X轴 = Sales，Y轴 = Costs）
// - 会在工作簿中新建或重写名为 __chart_data_temp 的临时表格工作表（隐藏）
// - 将 Sales/Costs 转为数字写入临时表，再基于该临时表创建图表，保证数值型轴
async function insertScatterChart() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItemAt(0);

    // load table columns to find Sales / Costs
    table.columns.load("items/name");
    await context.sync();

    const colNames = table.columns.items.map(c => String(c.name).trim());
    const salesExists = colNames.some(n => n.toLowerCase() === "sales");
    const costsExists = colNames.some(n => n.toLowerCase() === "costs");
    if (!salesExists || !costsExists) {
      throw new Error("表格中必须包含 'Sales' 和 'Costs' 两列。");
    }

    // 获取 Sales/Costs 的 DataBodyRange 并读取值
    const salesRange = table.columns.getItem("Sales").getDataBodyRange();
    const costsRange = table.columns.getItem("Costs").getDataBodyRange();
    salesRange.load(["rowCount", "values"]);
    costsRange.load(["rowCount", "values"]);
    await context.sync();

    if (salesRange.rowCount !== costsRange.rowCount) {
      throw new Error("Sales 与 Costs 列行数不一致。");
    }
    const n = salesRange.rowCount;
    if (n === 0) throw new Error("表格没有数据行。");

    // 清洗并构建写入临时sheet的数据（第一行为标题）
    const rows = [["Sales", "Costs"]];
    for (let i = 0; i < n; i++) {
      const rawS = salesRange.values[i][0];
      const rawC = costsRange.values[i][0];
      // 把可能的 $ , 空格 等干扰字符移除，确保 parseFloat 正确
      const s = parseFloat(String(rawS).replace(/[^0-9.\-]+/g, "")) || 0;
      const c = parseFloat(String(rawC).replace(/[^0-9.\-]+/g, "")) || 0;
      rows.push([s, c]);
    }

    // 创建或清空临时表 sheet "__chart_data_temp"
    const tempName = "__chart_data_temp";
    try {
      const existing = context.workbook.worksheets.getItem(tempName);
      existing.load("name");
      await context.sync();
      // 如果存在就删除（先删除再创建，确保干净）
      existing.delete();
      await context.sync();
    } catch (e) {
      // 不存在则忽略
    }
    const tempSheet = context.workbook.worksheets.add(tempName);

    // 写数据到临时 sheet，从 A1 开始（包含表头）
    const start = tempSheet.getRange("A1");
    const dataRange = start.getResizedRange(rows.length - 1, rows[0].length - 1);
    dataRange.values = rows;
    // 自动调整列宽
    dataRange.format.autofitColumns();
    await context.sync();

    // 现在在目标 sheet 上添加一个占位散点图，然后添加真正的 series
    // 使用一个最小的占位范围，随后添加 series 并删除占位 series
    const placeholder = tempSheet.getRange("A1:A2");
    const chart = sheet.charts.add(Excel.ChartType.xyscatter, placeholder, Excel.ChartSeriesBy.columns);

    const seriesCollection = chart.series;
    // 新建 series 并设置 X/Y
    const series = seriesCollection.add("Sales vs Costs");
    const xRange = tempSheet.getRange(`A2:A${rows.length}`); // Sales 数值列
    const yRange = tempSheet.getRange(`B2:B${rows.length}`); // Costs 数值列
    series.setXAxisValues(xRange);
    series.setValues(yRange);

    // 删除默认的第一个占位 series （通常是 index 0）
    try {
      seriesCollection.getItemAt(0).delete();
    } catch (e) {
      // 如果不存在占位不影响
    }

    // 设置图表位置与样式
    chart.title.text = "Sales (X) vs Costs (Y)";
    chart.setPosition(sheet.getRange("H5"), sheet.getRange("M25"));
    chart.legend.position = Excel.ChartLegendPosition.right;

    // 隐藏临时数据工作表（保留数据以便图表引用不会断）
    tempSheet.visibility = Excel.SheetVisibility.hidden;

    await context.sync();
    console.log("散点图已插入（X=Sales，Y=Costs），临时表名：", tempName);
  }).catch(err => {
    console.error("insertScatterChart 错误:", err);
    throw err;
  });
}






// 插入利润列并优先写入结构化公式，若失败则回退写数值
async function insertProfitColumn() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const table = sheet.tables.getItemAt(0);

    // 先加载列名，避免重复插入
    table.columns.load("items/name");
    await context.sync();

    if (table.columns.items.some(c => String(c.name).trim() === "Profits")) {
      console.log("表中已存在 Profits 列");
      return;
    }

    // 插入新列（列名是 Profits）
    table.columns.add(null, ["Profits"]);
    await context.sync();

    // 获取新列的数据区域
    const profitsRange = table.columns.getItem("Profits").getDataBodyRange();

    // 直接一次性赋结构化公式给整列
    profitsRange.formulas = [["=[@Sales]-[@Costs]"]];
    // 注意：这里的 [["公式"]] 是因为 profitsRange 是多行一列，
    // Excel 会自动将结构化引用复制到每一行

    await context.sync();
    console.log("利润列已插入并填充公式");
  }).catch((error) => {
    console.error("执行失败:", error);
  });
}