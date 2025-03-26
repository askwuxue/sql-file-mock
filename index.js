const { faker, zh_CN, Faker } = require("@faker-js/faker");
// const csvWriter = require('csv-writer').createObjectCsvWriter;
const ExcelJS = require("exceljs");

// 强制设置中文环境
faker.locale = "zh_CN";

const resorts = [
  "太舞滑雪小镇",
  "万龙滑雪场",
  "长白山万达滑雪场",
  "万科松花湖滑雪场",
  "云顶滑雪场",
  "将军山滑雪场",
  "怀北滑雪场",
  "富龙滑雪场",
  "亚布力滑雪场",
  "多乐美地滑雪场",
];

async function generateData() {
  const data = [];
  const startDate = new Date("2025-03-04");
  const endDate = new Date("2025-05-04");
  const twoMonthsInMs = endDate - startDate; // 62天间隔

  for (let i = 0; i < 100; i++) {
    // 用户名称（2-20字）
    let name;
    do {
      name = `${faker.person.firstName()} ${faker.person.lastName()}`.replace(
        /\s+/g,
        ""
      );
    } while (name.length < 2 || name.length > 20);

    // 操作类型
    const operation = faker.helpers.arrayElement(["purchase", "sale"]);

    // 雪场名称
    const resort = faker.helpers.arrayElement(resorts);

    // 时间（精确到秒，2025-03-04后两个月内）
    const randomDay = Math.floor(Math.random() * 62) + 1; // 包含起始日
    const date = new Date(startDate.getTime() + randomDay * 24 * 60 * 60);
    const time = new Date(date.getTime() + Math.random() * (24 * 60 * 60 - 1)); // 当天随机时间

    // 雪票数量
    const quantity = Math.floor(Math.random() * 100) + 1;

    // 价格（≥0，保留两位小数）
    // ✅ 使用新版 API 生成价格
    const price = faker.commerce.price({
      min: 0,
      max: 1000,
      precision: 2,
      locale: "zh_CN",
    });

    // 备注信息（200字内中文）
    let remark = faker.lorem.paragraphs(2); // 修改这里
    if (remark.length > 200) {
      remark = remark.slice(0, 200);
    }

    data.push({
      用户: name,
      操作类型: operation,
      雪场名称: resort,
      时间: time
        .toISOString()
        .replace("T", " ")
        .replace(/\.\d{3}/, ""),
      雪票数量: quantity,
      价格: price, // 强制保留两位小数显示
      备注信息: remark,
    });
  }
  return data;
}

async function exportExcel(data) {
  // 创建工作簿和工作表
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("模拟数据");

  // 添加标题行
  const headers = [
    "用户",
    "操作类型",
    "雪场名称",
    "时间",
    "雪票数量",
    "价格",
    "备注信息",
  ];
  worksheet.addRow(headers);

  // 添加数据行
  data.forEach((item) => {
    worksheet.addRow([
      item.用户,
      item.操作类型,
      item.雪场名称,
      item.时间,
      item.雪票数量,
      item.价格,
      item.备注信息,
    ]);
  });

  // 自动调整列宽
  worksheet.columns.forEach((column) => {
    column.width = Math.max(
      column.values?.map((v) => v.toString().length || 1),
      15
    );
  });

  // 保存文件
  await workbook.xlsx.writeFile("simulation_data.xlsx");
  console.log("📊 数据已生成至 simulation_data.xlsx 文件！");
}

generateData().then(exportExcel);
