// ======================================================
// ZiiZiiMoo - Google Apps Script
// 貼到：試算表 → 擴充功能 → Apps Script
// ======================================================

const SHEET_NAME_PRODUCTS = "Products";
const SHEET_NAME_ORDERS   = "Orders";
const NOTIFY_EMAIL        = "你的Gmail@gmail.com"; // ← 改成你的信箱

// ── 主入口 ────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  if (action === "getProducts") return getProducts();
  return jsonResponse({ error: "unknown action" });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === "submitOrder") return submitOrder(data);
  return jsonResponse({ error: "unknown action" });
}

// ── 取得商品列表 ──────────────────────────────────────
function getProducts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME_PRODUCTS);
  const rows  = sheet.getDataRange().getValues();
  // 自動找到表頭列（任一格 trim 後為 "id" 的那一列）
  const headerIdx = rows.findIndex(r => r.some(cell => String(cell).trim().toLowerCase() === "id"));
  if (headerIdx === -1) return jsonResponse({ products: [], debug: "no header found", firstRows: rows.slice(0,4).map(r => r.map(c => String(c).trim())) });
  const headers = rows[headerIdx].map(h => String(h).trim().toLowerCase());
  const products = rows.slice(headerIdx + 1)
    .filter(r => {
      const avail = String(r[headers.indexOf("available")]).trim().toUpperCase();
      return avail === "TRUE";
    })
    .map(r => ({
      id:          r[headers.indexOf("id")],
      name:        r[headers.indexOf("name")],
      description: r[headers.indexOf("description")],
      price:       r[headers.indexOf("price")],
      image_url:   r[headers.indexOf("image_url")],
      available:   r[headers.indexOf("available")]
    }));
  return jsonResponse({ products });
}

// ── 寫入訂單 + 寄信 ───────────────────────────────────
function submitOrder(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME_ORDERS);

  const timestamp  = new Date();
  const orderId    = "ORD-" + timestamp.getTime();
  const itemsText  = data.items.map(i => `${i.name} x${i.qty}`).join("、");

  // 寫入試算表
  sheet.appendRow([
    Utilities.formatDate(timestamp, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss"),
    orderId,
    data.customerName,
    data.phone,
    data.pickupDate,
    itemsText,
    data.total,
    data.note || ""
  ]);

  // 寄通知信給店主
  sendNotificationEmail(orderId, data, itemsText, timestamp);

  return jsonResponse({ success: true, orderId });
}

// ── 寄信功能 ──────────────────────────────────────────
function sendNotificationEmail(orderId, data, itemsText, timestamp) {
  const dateStr = Utilities.formatDate(timestamp, "Asia/Taipei", "yyyy/MM/dd HH:mm");
  const subject = `🧁 ZiiZiiMoo 新訂單 ${orderId}`;
  const body = `
【ZiiZiiMoo 新訂單通知】

訂單編號：${orderId}
下單時間：${dateStr}

─────────────────
👤 客戶資訊
姓名：${data.customerName}
電話：${data.phone}
取貨日期：${data.pickupDate}
備註：${data.note || "（無）"}

🧁 訂購品項
${data.items.map(i => `• ${i.name}  $${i.price} × ${i.qty} = $${i.price * i.qty}`).join("\n")}

💰 訂單總計：$${data.total}
─────────────────

請至 Google Sheets 查看所有訂單。
`.trim();

  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

// ── 工具函式 ──────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ======================================================
// 【初始化】第一次使用：執行這個函式建立表頭
// ======================================================
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Products 分頁
  let ps = ss.getSheetByName(SHEET_NAME_PRODUCTS);
  if (!ps) ps = ss.insertSheet(SHEET_NAME_PRODUCTS);
  if (ps.getLastRow() === 0) {
    ps.appendRow(["id","name","description","price","image_url","available"]);
    // 範例商品
    ps.appendRow([1,"抹茶磅蛋糕","濃郁日式抹茶，口感紮實",280,"","TRUE"]);
    ps.appendRow([2,"草莓塔","季節新鮮草莓，酥脆塔皮",220,"","TRUE"]);
    ps.appendRow([3,"伯爵奶油卷","英式伯爵茶香，每日限量",180,"","TRUE"]);
  }

  // Orders 分頁
  let os = ss.getSheetByName(SHEET_NAME_ORDERS);
  if (!os) os = ss.insertSheet(SHEET_NAME_ORDERS);
  if (os.getLastRow() === 0) {
    os.appendRow(["timestamp","order_id","姓名","電話","取貨日期","品項","總金額","備註"]);
  }

  SpreadsheetApp.getUi().alert("✅ 初始化完成！Products 和 Orders 分頁已建立。");
}
