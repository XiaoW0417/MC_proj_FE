const express = require("express");
const cors = require("cors");

const app = express();
app.use(cors());
app.use(express.json());

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

app.post("/analyze", async (req, res) => {
  const { message } = req.body;
  if (!message) return res.status(400).json({ error: "missing message" });

  await sleep(3000);

  const m = message.toLowerCase();

  if (m.includes("sort") && m.includes("sales")) {
    return res.json({
      action: "sort_sales_desc",
      description: "请预览并确认排序，点击“应用操作”即可将销售额降序排序应用到Excel表。",
    });
  }
  if (m.includes("scatter") && (m.includes("sales") || m.includes("costs"))) {
    return res.json({
      action: "scatter_sales_costs",
      description: "请预览散点图，点击“插入”即可将散点图插入到Excel表中。",
    });
  }
  if (
    m.includes("profit") ||
    (m.includes("profits") && m.includes("insert")) ||
    (m.includes("sales") && m.includes("costs") && m.includes("insert"))
  ) {
    return res.json({
      action: "insert_profits",
      description: "即将插入利润列，公式为 Profits = Sales - Costs。点击“插入公式”即可完成。",
    });
  }

  return res.json({
    action: "unsupported",
    description: "目前暂不支持，请重新输入",
  });
});

app.listen(3001, () => {
  console.log("Backend listening on http://localhost:3001");
});
