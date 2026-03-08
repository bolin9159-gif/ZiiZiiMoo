// ======================================================
// ZiiZiiMoo - Google Apps Script
// 貼到：試算表 → 擴充功能 → Apps Script
// ======================================================

const SHEET_NAME_PRODUCTS = "Products";
const SHEET_NAME_ORDERS   = "Orders";
const SHEET_NAME_MEMBERS  = "Members";
const NOTIFY_EMAIL        = "你的Gmail@gmail.com"; // ← 改成你的信箱

// ── 主入口 ────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action;
  if (action === "getProducts") return getProducts();
  return jsonResponse({ error: "unknown action" });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === "submitOrder")    return submitOrder(data);
  if (data.action === "sendVerifyCode") return sendVerifyCode(data);
  if (data.action === "verifyCode")     return verifyCode(data);
  if (data.action === "checkMember")    return checkMember(data);
  return jsonResponse({ error: "unknown action" });
}

// ── 取得商品列表 ──────────────────────────────────────
function getProducts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME_PRODUCTS);
  const rows  = sheet.getDataRange().getValues();
  const headerIdx = rows.findIndex(r => r.some(cell => String(cell).trim().toLowerCase() === "id"));
  if (headerIdx === -1) return jsonResponse({ products: [] });
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

// ── 會員相關 ─────────────────────────────────────────

// 取得 Members 分頁資料
function getMembersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME_MEMBERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_MEMBERS);
    sheet.appendRow(["email", "name", "phone", "verified", "verify_code", "code_expires", "status", "registered_at"]);
  }
  return sheet;
}

function getMemberHeaders(sheet) {
  const rows = sheet.getDataRange().getValues();
  const headerIdx = rows.findIndex(r => r.some(cell => String(cell).trim().toLowerCase() === "email"));
  if (headerIdx === -1) return { headerIdx: -1, headers: [], rows: [] };
  const headers = rows[headerIdx].map(h => String(h).trim().toLowerCase());
  return { headerIdx, headers, rows };
}

// 檢查會員狀態（前端用來判斷是否已驗證、是否有未完成訂單）
function checkMember(data) {
  const email = String(data.email || "").trim().toLowerCase();
  if (!email) return jsonResponse({ error: "請輸入 Email" });

  const sheet = getMembersSheet();
  const { headerIdx, headers, rows } = getMemberHeaders(sheet);
  if (headerIdx === -1) return jsonResponse({ exists: false });

  const emailCol = headers.indexOf("email");
  const verifiedCol = headers.indexOf("verified");
  const statusCol = headers.indexOf("status");

  for (let i = headerIdx + 1; i < rows.length; i++) {
    if (String(rows[i][emailCol]).trim().toLowerCase() === email) {
      const verified = String(rows[i][verifiedCol]).trim().toUpperCase() === "TRUE";
      const status = String(rows[i][statusCol]).trim();
      if (status === "blocked") {
        return jsonResponse({ exists: true, verified: false, blocked: true });
      }
      // 檢查是否有未完成訂單
      const hasActiveOrder = checkActiveOrder(email);
      return jsonResponse({ exists: true, verified, blocked: false, hasActiveOrder });
    }
  }
  return jsonResponse({ exists: false });
}

// 檢查是否有未完成訂單
function checkActiveOrder(email) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_ORDERS);
  if (!sheet) return false;
  const rows = sheet.getDataRange().getValues();
  const headerIdx = rows.findIndex(r => r.some(cell => String(cell).trim().toLowerCase() === "order_id"));
  if (headerIdx === -1) return false;
  const headers = rows[headerIdx].map(h => String(h).trim().toLowerCase());
  const emailCol = headers.indexOf("email");
  const statusCol = headers.indexOf("order_status");
  if (emailCol === -1 || statusCol === -1) return false;

  const completedStatuses = ["已取貨", "已取消"];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const orderEmail = String(rows[i][emailCol]).trim().toLowerCase();
    const orderStatus = String(rows[i][statusCol]).trim();
    if (orderEmail === email && !completedStatuses.includes(orderStatus)) {
      return true;
    }
  }
  return false;
}

// 發送驗證碼
function sendVerifyCode(data) {
  const email = String(data.email || "").trim().toLowerCase();
  const name  = String(data.name || "").trim();
  const phone = String(data.phone || "").trim();

  if (!email) return jsonResponse({ error: "請輸入 Email" });
  if (!email.includes("@")) return jsonResponse({ error: "Email 格式不正確" });

  const sheet = getMembersSheet();
  const { headerIdx, headers, rows } = getMemberHeaders(sheet);

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const expires = new Date(Date.now() + 10 * 60 * 1000); // 10 分鐘後過期

  if (headerIdx !== -1) {
    const emailCol = headers.indexOf("email");
    const statusCol = headers.indexOf("status");
    const codeCol = headers.indexOf("verify_code");
    const expiresCol = headers.indexOf("code_expires");

    for (let i = headerIdx + 1; i < rows.length; i++) {
      if (String(rows[i][emailCol]).trim().toLowerCase() === email) {
        if (String(rows[i][statusCol]).trim() === "blocked") {
          return jsonResponse({ error: "此 Email 已被停用，請聯繫店家" });
        }
        // 更新驗證碼
        sheet.getRange(i + 1, codeCol + 1).setValue(code);
        sheet.getRange(i + 1, expiresCol + 1).setValue(Utilities.formatDate(expires, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss"));
        sendCodeEmail(email, code);
        return jsonResponse({ success: true, message: "驗證碼已寄出" });
      }
    }
  }

  // 新會員：寫入一筆
  sheet.appendRow([
    email, name, phone, "FALSE", code,
    Utilities.formatDate(expires, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss"),
    "inactive",
    Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy/MM/dd HH:mm:ss")
  ]);
  sendCodeEmail(email, code);
  return jsonResponse({ success: true, message: "驗證碼已寄出" });
}

// 驗證碼確認
function verifyCode(data) {
  const email = String(data.email || "").trim().toLowerCase();
  const code  = String(data.code || "").trim();

  if (!email || !code) return jsonResponse({ error: "請輸入 Email 和驗證碼" });

  const sheet = getMembersSheet();
  const { headerIdx, headers, rows } = getMemberHeaders(sheet);
  if (headerIdx === -1) return jsonResponse({ error: "系統錯誤" });

  const emailCol = headers.indexOf("email");
  const codeCol = headers.indexOf("verify_code");
  const expiresCol = headers.indexOf("code_expires");
  const verifiedCol = headers.indexOf("verified");

  for (let i = headerIdx + 1; i < rows.length; i++) {
    if (String(rows[i][emailCol]).trim().toLowerCase() === email) {
      const savedCode = String(rows[i][codeCol]).trim();
      const expiresStr = String(rows[i][expiresCol]).trim();
      const expiresDate = new Date(expiresStr);

      if (new Date() > expiresDate) {
        return jsonResponse({ error: "驗證碼已過期，請重新發送" });
      }
      if (savedCode !== code) {
        return jsonResponse({ error: "驗證碼不正確" });
      }
      // 驗證成功：更新 verified + status
      sheet.getRange(i + 1, verifiedCol + 1).setValue("TRUE");
      const statusCol = headers.indexOf("status");
      if (statusCol !== -1) sheet.getRange(i + 1, statusCol + 1).setValue("active");
      sheet.getRange(i + 1, codeCol + 1).setValue(""); // 清除驗證碼
      return jsonResponse({ success: true, verified: true });
    }
  }
  return jsonResponse({ error: "找不到此 Email，請先發送驗證碼" });
}

// 寄驗證碼信
function sendCodeEmail(email, code) {
  const subject = "ZiiZiiMoo 訂購驗證碼";
  const body = `您好！

您的 ZiiZiiMoo 訂購驗證碼為：

    ${code}

此驗證碼將在 10 分鐘後失效。
如果這不是您本人的操作，請忽略此信件。

ZiiZiiMoo Handmade Bakery`;

  MailApp.sendEmail(email, subject, body);
}

// ── 寫入訂單 + 寄信 ───────────────────────────────────
function submitOrder(data) {
  const email = String(data.email || "").trim().toLowerCase();

  // 驗證會員身份
  const memberSheet = getMembersSheet();
  const { headerIdx: mhIdx, headers: mHeaders, rows: mRows } = getMemberHeaders(memberSheet);
  if (mhIdx === -1) return jsonResponse({ error: "系統錯誤" });

  const mEmailCol = mHeaders.indexOf("email");
  const mVerifiedCol = mHeaders.indexOf("verified");
  const mStatusCol = mHeaders.indexOf("status");
  let isMember = false;

  for (let i = mhIdx + 1; i < mRows.length; i++) {
    if (String(mRows[i][mEmailCol]).trim().toLowerCase() === email) {
      if (String(mRows[i][mStatusCol]).trim() === "blocked") {
        return jsonResponse({ error: "此帳號已被停用" });
      }
      if (String(mRows[i][mVerifiedCol]).trim().toUpperCase() !== "TRUE") {
        return jsonResponse({ error: "請先完成 Email 驗證" });
      }
      isMember = true;
      break;
    }
  }
  if (!isMember) return jsonResponse({ error: "請先完成 Email 驗證" });

  // 檢查是否有未完成訂單
  if (checkActiveOrder(email)) {
    return jsonResponse({ error: "您有未完成的訂單，結單後才能再次訂購" });
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(SHEET_NAME_ORDERS);

  const timestamp  = new Date();
  const orderId    = "ORD-" + timestamp.getTime();
  const itemsText  = data.items.map(i => `${i.name} x${i.qty}`).join("、");

  // 寫入試算表（新增 email、訂單狀態、訂金、尾款欄位）
  sheet.appendRow([
    Utilities.formatDate(timestamp, "Asia/Taipei", "yyyy/MM/dd HH:mm:ss"),
    orderId,
    email,
    data.customerName,
    data.phone,
    data.pickupDate,
    itemsText,
    data.total,
    data.note || "",
    "待確認",  // order_status
    0,         // deposit (訂金)
    0,         // balance (尾款)
    ""         // payment_note
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
Email：${data.email}
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
