const FIELDS = [
  { key: "order_no", label: "订单号", required: true },
  { key: "status", label: "交易状态", required: false },
  { key: "product_name", label: "商品名称", required: true },
  { key: "sales_amount", label: "销售额", required: true },
  { key: "cost_amount", label: "成本", required: false },
];

const HINTS = {
  order_no: ["订单号", "订单编号", "单号", "快手订单编号"],
  status: ["交易状态", "订单状态", "状态", "发货状态"],
  product_name: ["商品名称", "产品名称", "标题"],
  sales_amount: ["销售额", "订单金额", "成交金额", "支付金额", "实付"],
  cost_amount: ["成本", "采购价", "成本价", "成本金额"],
};

const officialFile = document.getElementById("officialFile");
const serviceFile = document.getElementById("serviceFile");
const inspectBtn = document.getElementById("inspectBtn");
const fileHint = document.getElementById("fileHint");
const mappingCard = document.getElementById("mappingCard");
const officialMapping = document.getElementById("officialMapping");
const serviceMapping = document.getElementById("serviceMapping");
const compareBtn = document.getElementById("compareBtn");
const resultCard = document.getElementById("resultCard");
const summary = document.getElementById("summary");
const resultTableBody = document.querySelector("#resultTable tbody");
const pageInfo = document.getElementById("pageInfo");
const prevBtn = document.getElementById("prevBtn");
const nextBtn = document.getElementById("nextBtn");
const exportBtn = document.getElementById("exportBtn");
const defaultCostInput = document.getElementById("defaultCost");

let officialColumns = [];
let serviceColumns = [];
let reportId = "";
let currentPage = 1;
let totalPages = 1;
const pageSize = 50;

inspectBtn.addEventListener("click", async () => {
  if (!officialFile.files[0] || !serviceFile.files[0]) {
    fileHint.textContent = "请先上传官方和客服两个文件。";
    return;
  }

  try {
    const [officialRes, serviceRes] = await Promise.all([
      inspectSingleFile(officialFile.files[0]),
      inspectSingleFile(serviceFile.files[0]),
    ]);

    officialColumns = officialRes.columns;
    serviceColumns = serviceRes.columns;

    renderMapping("official", officialColumns, officialMapping, true);
    renderMapping("service", serviceColumns, serviceMapping, false);

    mappingCard.classList.remove("hidden");
    fileHint.textContent = `文件读取完成：官方 ${officialRes.rows} 行，客服 ${serviceRes.rows} 行。`;
  } catch (err) {
    fileHint.textContent = err.message;
  }
});

compareBtn.addEventListener("click", async () => {
  if (!officialFile.files[0] || !serviceFile.files[0]) {
    alert("请先上传文件。");
    return;
  }

  const mappingForm = buildMappingForm();
  if (!mappingForm.ok) {
    alert(mappingForm.message);
    return;
  }

  try {
    const formData = new FormData();
    formData.append("official_file", officialFile.files[0]);
    formData.append("service_file", serviceFile.files[0]);

    Object.entries(mappingForm.mapping).forEach(([k, v]) => formData.append(k, v));
    formData.append("default_cost", defaultCostInput.value || "0");

    const resp = await fetch("/api/compare", { method: "POST", body: formData });
    const data = await resp.json();
    if (!data.ok) {
      throw new Error(data.message || "处理失败");
    }

    reportId = data.report_id;
    currentPage = 1;
    renderSummary(data.summary);
    await loadPage();
    resultCard.classList.remove("hidden");
  } catch (err) {
    alert(err.message);
  }
});

prevBtn.addEventListener("click", async () => {
  if (currentPage > 1) {
    currentPage -= 1;
    await loadPage();
  }
});

nextBtn.addEventListener("click", async () => {
  if (currentPage < totalPages) {
    currentPage += 1;
    await loadPage();
  }
});

exportBtn.addEventListener("click", () => {
  if (!reportId) {
    alert("请先进行对比。")
    return;
  }
  window.open(`/api/export/${reportId}`, "_blank");
});

async function inspectSingleFile(file) {
  const fd = new FormData();
  fd.append("file", file);

  const resp = await fetch("/api/inspect", { method: "POST", body: fd });
  const data = await resp.json();
  if (!data.ok) {
    throw new Error(data.message || "文件读取失败");
  }
  return data;
}

function renderMapping(prefix, columns, mountNode, isOfficial) {
  mountNode.innerHTML = "";
  FIELDS.forEach((field) => {
    if (!isOfficial && field.key === "status") {
      return;
    }

    const row = document.createElement("div");
    row.className = "mapping-item";

    const title = document.createElement("span");
    title.textContent = `${field.label}${field.required ? " *" : ""}`;

    const select = document.createElement("select");
    select.id = `${prefix}_${field.key}`;

    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "请选择列";
    select.appendChild(emptyOption);

    columns.forEach((col) => {
      const option = document.createElement("option");
      option.value = col;
      option.textContent = col;
      select.appendChild(option);
    });

    const guessed = guessField(field.key, columns);
    if (guessed) select.value = guessed;

    row.appendChild(title);
    row.appendChild(select);
    mountNode.appendChild(row);
  });
}

function guessField(key, columns) {
  const hints = HINTS[key] || [];
  for (const col of columns) {
    const lowerCol = String(col).toLowerCase();
    if (hints.some((h) => lowerCol.includes(h.toLowerCase()))) {
      return col;
    }
  }
  return "";
}

function buildMappingForm() {
  const mapping = {
    official_order_no: getVal("official_order_no"),
    official_status: getVal("official_status"),
    official_product_name: getVal("official_product_name"),
    official_sales_amount: getVal("official_sales_amount"),
    official_cost_amount: getVal("official_cost_amount"),
    service_order_no: getVal("service_order_no"),
    service_product_name: getVal("service_product_name"),
    service_sales_amount: getVal("service_sales_amount"),
    service_cost_amount: getVal("service_cost_amount"),
    service_status: "",
  };

  const requiredChecks = [
    ["official_order_no", "官方订单号"],
    ["official_status", "官方交易状态"],
    ["official_product_name", "官方商品名称"],
    ["official_sales_amount", "官方销售额"],
    ["service_order_no", "客服订单号"],
    ["service_product_name", "客服商品名称"],
    ["service_sales_amount", "客服销售额"],
  ];

  for (const [key, label] of requiredChecks) {
    if (!mapping[key]) {
      return { ok: false, message: `请完成字段映射：${label}` };
    }
  }

  return { ok: true, mapping };
}

function getVal(id) {
  const node = document.getElementById(id);
  return node ? node.value : "";
}

async function loadPage() {
  if (!reportId) return;

  const resp = await fetch(`/api/report/${reportId}?page=${currentPage}&page_size=${pageSize}`);
  const data = await resp.json();
  if (!data.ok) {
    alert(data.message || "分页读取失败");
    return;
  }

  totalPages = data.total_pages;
  pageInfo.textContent = `第 ${data.page} / ${data.total_pages} 页，共 ${data.total_rows} 条`;

  renderRows(data.records);
}

function renderRows(records) {
  resultTableBody.innerHTML = "";

  records.forEach((row) => {
    const tr = document.createElement("tr");
    if (row["利润"] < 0) tr.classList.add("loss-row");
    if (row["交易状态"] === "客服漏记" || row["交易状态"] === "异常订单") tr.classList.add("warn-row");

    tr.innerHTML = `
      <td>${row["类序号"] ?? ""}</td>
      <td>${row["订单号"] ?? ""}</td>
      <td>${row["商品名称"] ?? ""}</td>
      <td>${toFixed2(row["销售额"])}</td>
      <td>${toFixed2(row["成本"])}</td>
      <td>${toFixed2(row["利润"])}</td>
      <td>${row["交易状态"] ?? ""}</td>
    `;
    resultTableBody.appendChild(tr);
  });
}

function renderSummary(s) {
  summary.innerHTML = `
    <strong>汇总统计：</strong>
    总销售额 ${toFixed2(s.total_sales)} ｜ 总成本 ${toFixed2(s.total_cost)} ｜ 总利润 ${toFixed2(s.total_profit)} ｜
    订单总数 ${s.order_count} ｜ 客服漏记 ${s.missing_count} ｜ 异常订单 ${s.abnormal_count} ｜ 亏损订单 ${s.loss_count}
    <br/>
    数据清洗：官方去空订单号 ${s.official_stats.empty_order_removed}，官方重复 ${s.official_stats.duplicate_rows}，官方状态过滤 ${s.official_stats.status_filtered_rows}；
    客服去空订单号 ${s.service_stats.empty_order_removed}，客服重复 ${s.service_stats.duplicate_rows}。
  `;
}

function toFixed2(v) {
  const n = Number(v || 0);
  return Number.isFinite(n) ? n.toFixed(2) : "0.00";
}
