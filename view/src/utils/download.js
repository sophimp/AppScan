import FileSaver from "file-saver";
import { utils, write } from "xlsx";
import ExcelJS from "exceljs";
import moment from "moment";

export function exportJson2Excel(data, fileName, sheetName) {
  const sheet = utils.json_to_sheet(data);
  return FileSaver.saveAs(
    new Blob([sheet2Blob(sheet, sheetName)], {
      type: "application/octet-stream;charset=utf-8",
    }),
    fileName
  );
}

// ============ 专业合规检测报告导出 ============

// 检测依据标准列表
const DETECTION_STANDARDS = [
  "全国人大常委会《中华人民共和国网络安全法》",
  "全国人大常委会《个人信息保护法》",
  "全国信息安全标准化技术委员会《GB/T 35273-2020-信息安全技术 个人信息安全规范》",
  "全国信息安全标准化技术委员会《GBT 41391-2022-信息安全技术 移动互联网应用程序（App）收集个人信息基本要求》",
  "全国信息安全标准化技术委员会《GB/T 42574—2023 信息安全技术 个人信息处理中告知和同意的实施指南》",
  "全国信息安全标准化技术委员会《移动互联网应用程序（App）收集使用个人信息自评估指南》",
  "全国信息安全标准化技术委员会《网络安全标准实践指南—移动互联网应用程序（App）个人信息保护常见问题及处置指南》",
  "全国信息安全标准化技术委员会《移动互联网应用程序（APP）系统权限申请使用指南》",
  "App专项治理工作组《App违法违规收集使用个人信息自评估指南》",
  "APP专项治理工作组《App申请安卓系统权限机制分析与建议》",
  "四部委《App违法违规收集使用个人信息行为认定方法》",
  "四部委《常见类型移动互联网应用程序必要个人信息范围规定》",
  "工业和信息化部《关于开展APP侵害用户权益专项整治工作的通知（工信部信管函〔2019〕337号）》",
  "工业和信息化部《关于开展纵深推进APP侵害用户权益专项整治行动的通知（工信部信管函〔2020〕164号）》",
  "国家互联网信息办公室《数据出境安全评估办法》",
  "国家互联网信息办公室《儿童个人信息网络保护规定》",
  "电信终端产业协会-工信部团体标《移动智能终端与应用软件用户个人信息保护实施指南》",
  "电信终端产业协会-工信部团体标《APP收集使用个人信息最小必要评估规范》",
  "电信终端产业协会-工信部团体标《APP用户权益保护测评规范》",
  "全国信息安全标准化技术委员会《移动互联网应用程序(App接入软件开发工具包(SDK)个人信息安全指南》（T/CSAC 006—2023）",
  "全国网络安全标准化技术委员会《信息安全技术 移动互联应用程序（App）个人信息安全测评规范》",
  "工业和信息化部《关于进一步提升移动互联网应用服务能力的通知（工信部信管函〔2023〕26号）》",
];

// 规则类型 → 检测大类映射
const TYPE_CATEGORY_MAP = {
  useRules: "隐私政策配置合规",
  saveRules: "数据使用合规",
  externalRules: "数据对外提供合规",
  collectionRules: "数据采集合规",
  expressConsent: "明示同意合规",
  termStatus: "隐私政策条款合规",
  userPermissions: "用户权益保障合规",
};

// 风险等级映射
function riskLevelText(level) {
  switch (level) {
    case 0: return "低风险";
    case 1: return "中风险";
    case 2: return "高风险";
    default: return "暂未发现";
  }
}

// 风险颜色
function riskColor(level) {
  switch (level) {
    case 0: return { argb: "FF4CAF50" }; // 绿色-低
    case 1: return { argb: "FFFF9800" }; // 橙色-中
    case 2: return { argb: "FFF44336" }; // 红色-高
    default: return { argb: "FF9E9E9E" }; // 灰色
  }
}

// 检测结果文本
function checkResultText(option) {
  if (option === 0) return "不通过";
  if (option === 1) return "通过";
  return "待检测";
}

// 通用样式
const BORDER_THIN = {
  top: { style: "thin" },
  left: { style: "thin" },
  bottom: { style: "thin" },
  right: { style: "thin" },
};
const HEADER_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1A237E" } };
const SECTION_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE3F2FD" } };
const LIGHT_FILL = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
const HEADER_FONT = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
const SECTION_FONT = { bold: true, color: { argb: "FF1A237E" }, size: 11 };
const TITLE_FONT = { bold: true, color: { argb: "FF1A237E" }, size: 18 };
const SUBTITLE_FONT = { bold: true, color: { argb: "FF1A237E" }, size: 14 };

function applyHeaderRow(row) {
  row.eachCell(cell => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.border = BORDER_THIN;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  });
  row.height = 28;
}

function applyDataRow(row, fillBg) {
  row.eachCell(cell => {
    cell.border = BORDER_THIN;
    cell.alignment = { vertical: "middle", wrapText: true };
    if (fillBg) cell.fill = LIGHT_FILL;
  });
}

// 写入KV信息表（两列：键、值）
function writeKVTable(sheet, startRow, kvPairs) {
  const headerRow = sheet.getRow(startRow);
  headerRow.getCell(1).value = "项目";
  headerRow.getCell(2).value = "内容";
  applyHeaderRow(headerRow);

  let r = startRow + 1;
  kvPairs.forEach(([key, val], i) => {
    const row = sheet.getRow(r);
    row.getCell(1).value = key;
    row.getCell(2).value = val || "/";
    applyDataRow(row, i % 2 === 0);
    row.height = 20;
    r++;
  });
  return r;
}

// ============ 主导出函数 ============
export async function exportProfessionalReport(params) {
  const {
    appName,
    appPackage,
    version,
    createTime,
    marks,     // getRiskData 返回的检测项数组
    methods,   // 隐私行为原始数据
  } = params;

  const wb = new ExcelJS.Workbook();
  wb.creator = "AppScan";
  wb.created = new Date();

  const reportDate = moment(createTime).format("YYYY-MM-DD HH:mm:ss");
  const exportDate = moment().format("YYYY年MM月DD日");

  // ==================== Sheet 1: 报告概览 ====================
  const sheet1 = wb.addWorksheet("报告概览");
  sheet1.columns = [
    { width: 20 },
    { width: 50 },
  ];
  sheet1.views = [{ showGridLines: false }];

  // 标题
  sheet1.mergeCells("A1:B1");
  const titleCell = sheet1.getCell("A1");
  titleCell.value = "App隐私合规检测报告";
  titleCell.font = TITLE_FONT;
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  sheet1.getRow(1).height = 50;

  // 副标题
  sheet1.mergeCells("A2:B2");
  const subCell = sheet1.getCell("A2");
  subCell.value = `报告生成时间：${exportDate}`;
  subCell.font = { color: { argb: "FF616161" }, size: 11 };
  subCell.alignment = { horizontal: "center", vertical: "middle" };
  sheet1.getRow(2).height = 28;

  // 空行
  sheet1.getRow(3).height = 10;

  // 应用基本信息
  sheet1.mergeCells("A4:B4");
  const sec1 = sheet1.getCell("A4");
  sec1.value = "应用详情 — 基本信息";
  sec1.font = SUBTITLE_FONT;
  sec1.fill = SECTION_FILL;
  sec1.alignment = { vertical: "middle" };
  sheet1.getRow(4).height = 30;

  let nextRow = writeKVTable(sheet1, 5, [
    ["应用名称", appName],
    ["应用包名", appPackage],
    ["文件类型", "Android"],
    ["应用版本", version],
    ["检测时间", reportDate],
  ]);

  // 空行
  sheet1.getRow(nextRow).height = 10;
  nextRow++;

  // 检测概览统计
  sheet1.mergeCells(nextRow, 1, nextRow, 2);
  const sec2 = sheet1.getCell(`A${nextRow}`);
  sec2.value = "检测概览";
  sec2.font = SUBTITLE_FONT;
  sec2.fill = SECTION_FILL;
  sec2.alignment = { vertical: "middle" };
  sheet1.getRow(nextRow).height = 30;
  nextRow++;

  // 统计
  const totalItems = marks.length;
  let highCount = 0, midCount = 0, lowCount = 0, passCount = 0, failCount = 0, pendingCount = 0;
  marks.forEach(m => {
    if (m.riskLevel === 2) highCount++;
    else if (m.riskLevel === 1) midCount++;
    else if (m.riskLevel === 0) lowCount++;
    if (m.option === 1) passCount++;
    else if (m.option === 0) failCount++;
    else pendingCount++;
  });

  const statsHeader = sheet1.getRow(nextRow);
  statsHeader.values = ["风险等级", "数量（占比）"];
  applyHeaderRow(statsHeader);
  nextRow++;

  const riskStats = [
    ["高风险", `${highCount}（${totalItems ? Math.round(highCount / totalItems * 100) : 0}%）`],
    ["中风险", `${midCount}（${totalItems ? Math.round(midCount / totalItems * 100) : 0}%）`],
    ["低风险", `${lowCount}（${totalItems ? Math.round(lowCount / totalItems * 100) : 0}%）`],
    ["通过", `${passCount}（${totalItems ? Math.round(passCount / totalItems * 100) : 0}%）`],
    ["不通过", `${failCount}（${totalItems ? Math.round(failCount / totalItems * 100) : 0}%）`],
    ["待检测", `${pendingCount}（${totalItems ? Math.round(pendingCount / totalItems * 100) : 0}%）`],
    ["总检测项", `${totalItems}`],
  ];
  riskStats.forEach(([label, val], i) => {
    const row = sheet1.getRow(nextRow);
    row.values = [label, val];
    applyDataRow(row, i % 2 === 0);
    // 颜色标识
    if (label === "高风险") row.getCell(1).font = { bold: true, color: { argb: "FFF44336" } };
    if (label === "中风险") row.getCell(1).font = { bold: true, color: { argb: "FFFF9800" } };
    if (label === "低风险") row.getCell(1).font = { bold: true, color: { argb: "FF4CAF50" } };
    row.height = 22;
    nextRow++;
  });

  // ==================== Sheet 2: 检测依据 ====================
  const sheet2 = wb.addWorksheet("检测依据");
  sheet2.columns = [
    { width: 8 },
    { width: 80 },
  ];

  const s2Title = sheet2.getRow(1);
  s2Title.values = ["序号", "检测依据"];
  applyHeaderRow(s2Title);

  DETECTION_STANDARDS.forEach((std, i) => {
    const row = sheet2.getRow(i + 2);
    row.values = [i + 1, std];
    applyDataRow(row, i % 2 === 0);
    row.height = 22;
  });

  // ==================== Sheet 3: 检测结论 ====================
  const sheet3 = wb.addWorksheet("检测结论");
  sheet3.columns = [
    { width: 8 },
    { width: 25 },
    { width: 35 },
    { width: 15 },
    { width: 15 },
    { width: 40 },
  ];

  // 标题
  sheet3.mergeCells("A1:F1");
  const s3Title = sheet3.getCell("A1");
  s3Title.value = "检测结论";
  s3Title.font = SUBTITLE_FONT;
  s3Title.fill = SECTION_FILL;
  s3Title.alignment = { vertical: "middle" };
  sheet3.getRow(1).height = 30;

  const s3Header = sheet3.getRow(2);
  s3Header.values = ["序号", "检测项", "评估标准", "风险等级", "检测结果", "整改建议"];
  applyHeaderRow(s3Header);

  marks.forEach((mark, i) => {
    const row = sheet3.getRow(i + 3);
    const riskText = riskLevelText(mark.riskLevel);
    const resultText = checkResultText(mark.option);
    row.values = [
      i + 1,
      mark.project,
      mark.standard,
      riskText,
      resultText,
      mark.suggest || "/",
    ];
    applyDataRow(row, i % 2 === 0);
    // 风险等级颜色
    row.getCell(4).font = { bold: true, color: riskColor(mark.riskLevel) };
    // 检测结果颜色
    if (mark.option === 0) row.getCell(5).font = { bold: true, color: { argb: "FFF44336" } };
    else if (mark.option === 1) row.getCell(5).font = { bold: true, color: { argb: "FF4CAF50" } };
    else row.getCell(5).font = { color: { argb: "FF9E9E9E" } };
    row.height = 50;
  });

  // ==================== Sheet 4: 检测详情 ====================
  const sheet4 = wb.addWorksheet("检测详情");
  sheet4.columns = [
    { width: 8 },
    { width: 18 },
    { width: 30 },
    { width: 35 },
    { width: 12 },
    { width: 12 },
    { width: 40 },
  ];

  // 标题
  sheet4.mergeCells("A1:G1");
  const s4Title = sheet4.getCell("A1");
  s4Title.value = "检测详情";
  s4Title.font = SUBTITLE_FONT;
  s4Title.fill = SECTION_FILL;
  s4Title.alignment = { vertical: "middle" };
  sheet4.getRow(1).height = 30;

  // 按大类分组
  const grouped = {};
  marks.forEach(mark => {
    const category = TYPE_CATEGORY_MAP[mark.type] || "其他";
    if (!grouped[category]) grouped[category] = [];
    grouped[category].push(mark);
  });

  let s4Row = 2;
  const categoryOrder = [
    "隐私政策配置合规",
    "明示同意合规",
    "数据采集合规",
    "数据使用合规",
    "数据对外提供合规",
    "隐私政策条款合规",
    "用户权益保障合规",
  ];

  categoryOrder.forEach(category => {
    const items = grouped[category];
    if (!items || items.length === 0) return;

    // 分类标题行
    const catRow = sheet4.getRow(s4Row);
    sheet4.mergeCells(s4Row, 1, s4Row, 7);
    catRow.getCell(1).value = category;
    catRow.getCell(1).font = SECTION_FONT;
    catRow.getCell(1).fill = SECTION_FILL;
    catRow.getCell(1).alignment = { vertical: "middle" };
    catRow.height = 28;
    s4Row++;

    // 表头
    const hRow = sheet4.getRow(s4Row);
    hRow.values = ["序号", "风险等级", "检测项", "评估标准", "检测结果", "合规状态", "整改建议"];
    applyHeaderRow(hRow);
    s4Row++;

    items.forEach((mark, i) => {
      const row = sheet4.getRow(s4Row);
      const riskText = riskLevelText(mark.riskLevel);
      const resultText = checkResultText(mark.option);
      const complianceText = mark.option === 1 ? "合规" : mark.option === 0 ? "不合规" : "待确认";
      row.values = [
        i + 1,
        riskText,
        mark.project,
        mark.standard,
        resultText,
        complianceText,
        mark.suggest || "/",
      ];
      applyDataRow(row, i % 2 === 0);
      row.getCell(2).font = { bold: true, color: riskColor(mark.riskLevel) };
      if (mark.option === 0) row.getCell(5).font = { bold: true, color: { argb: "FFF44336" } };
      else if (mark.option === 1) row.getCell(5).font = { bold: true, color: { argb: "FF4CAF50" } };
      row.getCell(6).font = {
        bold: true,
        color: complianceText === "不合规" ? { argb: "FFF44336" } :
               complianceText === "合规" ? { argb: "FF4CAF50" } : { argb: "FF9E9E9E" }
      };
      row.height = 60;
      s4Row++;
    });

    // 空行分隔
    sheet4.getRow(s4Row).height = 8;
    s4Row++;
  });

  // ==================== Sheet 5: 隐私行为数据 ====================
  const sheet5 = wb.addWorksheet("隐私行为数据");
  sheet5.columns = [
    { width: 8 },
    { width: 30 },
    { width: 30 },
    { width: 20 },
    { width: 50 },
  ];

  const s5Header = sheet5.getRow(1);
  s5Header.values = ["序号", "调用主体", "操作行为", "行为类型", "调用堆栈/详情"];
  applyHeaderRow(s5Header);

  if (methods && methods.length > 0) {
    methods.forEach((method, i) => {
      const row = sheet5.getRow(i + 2);
      row.values = [
        i + 1,
        method.main || "",
        method.action || "",
        method.type || "",
        method.data || "",
      ];
      applyDataRow(row, i % 2 === 0);
      row.height = 28;
    });
  } else {
    const emptyRow = sheet5.getRow(2);
    sheet5.mergeCells("A2:E2");
    emptyRow.getCell(1).value = "暂无隐私行为数据";
    emptyRow.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
    emptyRow.height = 30;
  }

  // ==================== Sheet 6: 声明 ====================
  const sheet6 = wb.addWorksheet("声明");
  sheet6.columns = [{ width: 100 }];
  sheet6.views = [{ showGridLines: false }];

  const statement = [
    "声明",
    "",
    "本报告为AppScan隐私合规检测平台根据相关法律法规对APP进行的隐私合规自动化检测后出具的专业版《隐私合规检测报告》。",
    "",
    "本报告为专业型的机器自动检测，部分涉及人工评估的条款均未涵盖，需要您根据自身情况继续进行自查。",
    "",
    "由于检测方法的局限，本报告中的部分评估项检测结果可能存在偏差，必要时请进行人工核验结果。",
    "",
    "本报告不属于隐私合规法律意见，仅作为隐私合规检测的建议性参考文件。对被检测APP收集使用个人信息情况的任何改进建议均不保证一定能满足监管部门隐私合规要求，也不保证您一定会通过监管部门的评估，需要您具体考虑监管部门的要求进行改进。",
  ];

  const stmtTitle = sheet6.getCell("A1");
  stmtTitle.value = "声明";
  stmtTitle.font = SUBTITLE_FONT;
  stmtTitle.alignment = { horizontal: "center", vertical: "middle" };
  sheet6.getRow(1).height = 40;

  statement.slice(1).forEach((text, i) => {
    const row = sheet6.getRow(i + 2);
    row.getCell(1).value = text;
    row.getCell(1).alignment = { wrapText: true, vertical: "middle" };
    row.getCell(1).font = { color: { argb: "FF424242" }, size: 11 };
    row.height = text ? 40 : 10;
  });

  // 导出
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const fileName = `${appName || "应用"}_${appPackage || ""}_隐私合规检测报告_${moment().format("YYYY-MM-DD_HH_mm_ss")}.xlsx`;
  return FileSaver.saveAs(blob, fileName);
}

function sheet2Blob(sheet, sheetName) {
  sheetName = sheetName || "sheet1";
  const workbook = {
    SheetNames: [sheetName],
    Sheets: {},
  };
  workbook.Sheets[sheetName] = sheet;
  // 生成excel的配置项
  const wopts = {
    bookType: "xlsx", // 要生成的文件类型
    bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
    type: "binary",
  };
  const wbout = write(workbook, wopts);
  const blob = new Blob([string2ArrayBuffer(wbout)], {
    type: "application/octet-stream",
  });
  return blob;
}

// 字符串转ArrayBuffer
function string2ArrayBuffer(s) {
  const buf = new ArrayBuffer(s.length);
  const view = new Uint8Array(buf);
  for (let i = 0; i < s.length; ++i) {
    view[i] = s.charCodeAt(i) & 0xff;
  }
  return buf;
}

export function formatSize(limit) {
  if (!limit || Number(limit) == 0) return "";
  limit = Number(limit);
  // 将size B转换成 M
  var size = "";
  if (limit < 1 * 1024) {
    //小于1KB，则转化成B
    size = limit.toFixed(2) + "B";
  } else if (limit < 1 * 1024 * 1024) {
    //小于1MB，则转化成KB
    size = (limit / 1024).toFixed(2) + "KB";
  } else if (limit < 1 * 1024 * 1024 * 1024) {
    //小于1GB，则转化成MB
    size = (limit / (1024 * 1024)).toFixed(2) + "MB";
  } else {
    //其他转化成GB
    size = (limit / (1024 * 1024 * 1024)).toFixed(2) + "GB";
  }

  var sizeStr = size + ""; //转成字符串
  var index = sizeStr.indexOf("."); //获取小数点处的索引
  var dou = sizeStr.substr(index + 1, 2); //获取小数点后两位的值
  if (dou == "00") {
    //判断后两位是否为00，如果是则删除00
    return sizeStr.substring(0, index) + sizeStr.substr(index + 3, 2);
  }
  return size;
}
