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
  const firstLine = c => String(c).split("\n")[0].trim().toLowerCase();
  const headerIdx = rows.findIndex(r => r.some(cell => firstLine(cell) === "id"));
  if (headerIdx === -1) return jsonResponse({ products: [] });
  const headers = rows[headerIdx].map(h => firstLine(h));
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
  // 取第一行文字（換行前）做比對，例如 "email\n電子信箱" → "email"
  const firstLine = c => String(c).split("\n")[0].trim().toLowerCase();
  const headerIdx = rows.findIndex(r => r.some(cell => firstLine(cell) === "email"));
  if (headerIdx === -1) return { headerIdx: -1, headers: [], rows: [] };
  const headers = rows[headerIdx].map(h => firstLine(h));
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
  const firstLine = c => String(c).split("\n")[0].trim().toLowerCase();
  const headerIdx = rows.findIndex(r => r.some(cell => firstLine(cell) === "order_id"));
  if (headerIdx === -1) return false;
  const headers = rows[headerIdx].map(h => firstLine(h));
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
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return jsonResponse({ error: "Email 格式不正確，請確認格式（例如：name@example.com）" });

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
  styleNewRow(sheet, SHEET_NAME_MEMBERS);
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
      // 更新該列樣式（verified → 綠色, status → 綠色）
      applyStatusStyle(sheet, i + 1, verifiedCol + 1, "true");
      if (statusCol !== -1) applyMemberStatusStyle(sheet, i + 1, statusCol + 1, "active");
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
  styleNewRow(sheet, SHEET_NAME_ORDERS);

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

// ══════════════════════════════════════════════════════
// ── 統一樣式系統（完全依據現有 Sheet 設計）──────────────
// ══════════════════════════════════════════════════════

// ── 色票定義（從現有 Sheet 擷取）─────────────────────
const COLORS = {
  // 共用
  white:     "#FFFCF8",
  cream:     "#FDF6EE",
  warm:      "#F5E6D3",
  brown:     "#7A4F3A",
  text:      "#4A3428",
  mauve:     "#C97D7D",
  muted:     "#9C7B6E",

  // 狀態色
  greenBg:   "#D6F0D3",  greenText:   "#3A7A4F",  // TRUE, active
  redBg:     "#FAD4D4",  redText:     "#7A3A3A",  // FALSE, blocked
  yellowBg:  "#FFF8E1",  yellowText:  "#856404",  // inactive, 待確認
  blueBg:    "#E3F2FD",  blueText:    "#1565C0",  // 已收訂金
  tealBg:    "#EAF4F2",  tealText:    "#4A8A7A",  // Orders 說明列

  // Products 專用
  prodHeader:    "#C97D7D",

  // Orders 專用
  orderHeader:   "#E8A0A0",
  orderPayHeader:"#4A8A7A",
  orderNoteBg:   "#EAF4F2",

  // Members 專用
  memberTitle:   "#4A6FA5",
  memberHeader:  "#4A6FA5",
  memberNoteBg:  "#EAF0FB"
};

// ── 開啟試算表時加入自訂選單 ─────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("ZiiZiiMoo")
    .addItem("格式化所有分頁", "formatAllSheets")
    .addItem("格式化目前分頁", "formatCurrentSheet")
    .addToUi();
}

// ── 格式化所有分頁 ──────────────────────────────────
function formatAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHEET_NAME_PRODUCTS, SHEET_NAME_ORDERS, SHEET_NAME_MEMBERS].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) formatSheet(sheet, name);
  });
  SpreadsheetApp.getUi().alert("✅ 所有分頁格式化完成！");
}

// ── 格式化目前分頁 ──────────────────────────────────
function formatCurrentSheet() {
  const sheet = SpreadsheetApp.getActiveSheet();
  formatSheet(sheet, sheet.getName());
  SpreadsheetApp.getUi().alert("✅ " + sheet.getName() + " 格式化完成！");
}

// ── 格式化單一分頁 ──────────────────────────────────
function formatSheet(sheet, sheetName) {
  const lastRow = Math.max(sheet.getLastRow(), 4);
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  // 全頁基礎：Arial、垂直置中
  const allRange = sheet.getRange(1, 1, lastRow, lastCol);
  allRange.setFontFamily("Arial");
  allRange.setVerticalAlignment("middle");

  // ── Row 1：標題列 ──
  const titleBg = (sheetName === SHEET_NAME_MEMBERS) ? COLORS.memberTitle : COLORS.brown;
  const r1 = sheet.getRange(1, 1, 1, lastCol);
  if (lastCol > 1) r1.merge();
  r1.setBackground(titleBg)
    .setFontColor(COLORS.white)
    .setFontSize(14)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(1, 40);

  // ── Row 2：說明列 ──
  if (lastRow >= 2) {
    const r2 = sheet.getRange(2, 1, 1, lastCol);
    if (lastCol > 1) r2.merge();
    r2.setBackground(COLORS.warm)
      .setFontColor(COLORS.mauve)
      .setFontSize(9)
      .setFontWeight("normal")
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
    sheet.setRowHeight(2, 25);
  }

  // ── Row 3：表頭列 ──
  if (lastRow >= 3) {
    let headerBg = COLORS.prodHeader;
    if (sheetName === SHEET_NAME_ORDERS) headerBg = COLORS.orderHeader;
    if (sheetName === SHEET_NAME_MEMBERS) headerBg = COLORS.memberHeader;

    const r3 = sheet.getRange(3, 1, 1, lastCol);
    r3.setBackground(headerBg)
      .setFontColor(COLORS.white)
      .setFontSize(10)
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setWrap(true);
    sheet.setRowHeight(3, 45);

    // Orders: J~M 欄（付款區）用青綠色
    if (sheetName === SHEET_NAME_ORDERS && lastCol >= 10) {
      const payCols = Math.min(lastCol - 9, 4); // J=10, K=11, L=12, M=13
      sheet.getRange(3, 10, 1, payCols)
        .setBackground(COLORS.orderPayHeader);
    }
  }

  // ── Row 4：說明列（Orders / Members 才有）──
  if (lastRow >= 4 && (sheetName === SHEET_NAME_ORDERS || sheetName === SHEET_NAME_MEMBERS)) {
    const r4 = sheet.getRange(4, 1, 1, lastCol);
    if (lastCol > 1) r4.merge();
    const noteBg = (sheetName === SHEET_NAME_ORDERS) ? COLORS.orderNoteBg : COLORS.memberNoteBg;
    const noteColor = (sheetName === SHEET_NAME_ORDERS) ? COLORS.tealText : COLORS.memberHeader;
    r4.setBackground(noteBg)
      .setFontColor(noteColor)
      .setFontSize(9)
      .setFontWeight("normal")
      .setFontStyle("italic")
      .setHorizontalAlignment("center");
  }

  // ── Row 4/5+：資料列 ──
  const dataStartRow = (sheetName === SHEET_NAME_ORDERS || sheetName === SHEET_NAME_MEMBERS) ? 5 : 4;
  for (let row = dataStartRow; row <= lastRow; row++) {
    formatDataRow(sheet, sheetName, row, lastCol);
  }

  // 凍結前 3 列
  sheet.setFrozenRows(3);

  // 自動調整欄寬（80~250）
  for (let col = 1; col <= lastCol; col++) {
    sheet.autoResizeColumn(col);
    const w = sheet.getColumnWidth(col);
    if (w < 80) sheet.setColumnWidth(col, 80);
    if (w > 250) sheet.setColumnWidth(col, 250);
  }
}

// ── 格式化單一資料列 ─────────────────────────────────
function formatDataRow(sheet, sheetName, rowNum, lastCol) {
  const isAlt = (rowNum % 2 === 0);
  const baseBg = isAlt ? COLORS.cream : COLORS.white;
  const rowRange = sheet.getRange(rowNum, 1, 1, lastCol);

  // 基礎樣式
  rowRange.setBackground(baseBg)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setFontWeight("normal")
    .setFontStyle("normal")
    .setFontColor(COLORS.text)
    .setVerticalAlignment("middle")
    .setWrap(true);

  // ── 各分頁的欄位特殊樣式 ──

  if (sheetName === SHEET_NAME_PRODUCTS) {
    // A 欄（id）：置中、粗體、棕色
    sheet.getRange(rowNum, 1).setHorizontalAlignment("center").setFontWeight("bold").setFontColor(COLORS.brown);
    // B 欄（name）：靠左
    if (lastCol >= 2) sheet.getRange(rowNum, 2).setHorizontalAlignment("left");
    // C 欄（description）：靠左
    if (lastCol >= 3) sheet.getRange(rowNum, 3).setHorizontalAlignment("left");
    // D 欄（price）：靠右、玫瑰色
    if (lastCol >= 4) sheet.getRange(rowNum, 4).setHorizontalAlignment("right").setFontColor(COLORS.mauve);
    // F 欄（available）：條件色
    if (lastCol >= 6) {
      const val = String(sheet.getRange(rowNum, 6).getValue()).trim().toUpperCase();
      applyStatusStyle(sheet, rowNum, 6, val === "TRUE" ? "true" : "false");
    }
  }

  if (sheetName === SHEET_NAME_ORDERS) {
    // A 欄（timestamp）：置中、棕色
    sheet.getRange(rowNum, 1).setHorizontalAlignment("center").setFontColor(COLORS.brown);
    // B 欄（order_id）：置中、棕色
    if (lastCol >= 2) sheet.getRange(rowNum, 2).setHorizontalAlignment("center").setFontColor(COLORS.brown);
    // C~G 欄：靠左
    for (let c = 3; c <= Math.min(7, lastCol); c++) {
      sheet.getRange(rowNum, c).setHorizontalAlignment("left");
    }
    // H 欄（總金額）：靠右、玫瑰色
    if (lastCol >= 8) sheet.getRange(rowNum, 8).setHorizontalAlignment("right").setFontColor(COLORS.mauve);
    // I 欄（備註）：靠左
    if (lastCol >= 9) sheet.getRange(rowNum, 9).setHorizontalAlignment("left");
    // J 欄（order_status）：條件色
    if (lastCol >= 10) {
      const status = String(sheet.getRange(rowNum, 10).getValue()).trim();
      applyOrderStatusStyle(sheet, rowNum, 10, status);
    }
    // K 欄（deposit）：靠右、棕色
    if (lastCol >= 11) sheet.getRange(rowNum, 11).setHorizontalAlignment("right").setFontColor(COLORS.brown);
    // L 欄（balance）：靠右、棕色
    if (lastCol >= 12) sheet.getRange(rowNum, 12).setHorizontalAlignment("right").setFontColor(COLORS.brown);
    // M 欄（payment_note）：靠左
    if (lastCol >= 13) sheet.getRange(rowNum, 13).setHorizontalAlignment("left");
  }

  if (sheetName === SHEET_NAME_MEMBERS) {
    // A 欄（email）：置中、棕色
    sheet.getRange(rowNum, 1).setHorizontalAlignment("center").setFontColor(COLORS.brown);
    // B 欄（name）：靠左
    if (lastCol >= 2) sheet.getRange(rowNum, 2).setHorizontalAlignment("left");
    // C 欄（phone）：靠左
    if (lastCol >= 3) sheet.getRange(rowNum, 3).setHorizontalAlignment("left");
    // D 欄（verified）：條件色
    if (lastCol >= 4) {
      const val = String(sheet.getRange(rowNum, 4).getValue()).trim().toUpperCase();
      applyStatusStyle(sheet, rowNum, 4, val === "TRUE" ? "true" : "false");
    }
    // E~F 欄：靠左
    if (lastCol >= 5) sheet.getRange(rowNum, 5).setHorizontalAlignment("left");
    if (lastCol >= 6) sheet.getRange(rowNum, 6).setHorizontalAlignment("left");
    // G 欄（status）：條件色
    if (lastCol >= 7) {
      const val = String(sheet.getRange(rowNum, 7).getValue()).trim().toLowerCase();
      applyMemberStatusStyle(sheet, rowNum, 7, val);
    }
    // H 欄（registered_at）：置中、棕色
    if (lastCol >= 8) sheet.getRange(rowNum, 8).setHorizontalAlignment("center").setFontColor(COLORS.brown);
  }
}

// ── 狀態樣式：TRUE / FALSE ──────────────────────────
function applyStatusStyle(sheet, row, col, type) {
  const cell = sheet.getRange(row, col);
  cell.setHorizontalAlignment("center").setFontWeight("bold");
  if (type === "true") {
    cell.setBackground(COLORS.greenBg).setFontColor(COLORS.greenText);
  } else {
    cell.setBackground(COLORS.redBg).setFontColor(COLORS.redText);
  }
}

// ── 狀態樣式：訂單狀態 ──────────────────────────────
function applyOrderStatusStyle(sheet, row, col, status) {
  const cell = sheet.getRange(row, col);
  cell.setHorizontalAlignment("center").setFontWeight("bold");
  if (status === "待確認") {
    cell.setBackground(COLORS.yellowBg).setFontColor(COLORS.yellowText);
  } else if (status === "已收訂金") {
    cell.setBackground(COLORS.blueBg).setFontColor(COLORS.blueText);
  } else if (status === "已付清") {
    cell.setBackground("#E8F5E9").setFontColor("#2E7D32");
  } else if (status === "已取貨") {
    cell.setBackground(COLORS.greenBg).setFontColor(COLORS.greenText);
  } else if (status === "已取消") {
    cell.setBackground(COLORS.redBg).setFontColor(COLORS.redText);
  }
}

// ── 狀態樣式：會員狀態 ──────────────────────────────
function applyMemberStatusStyle(sheet, row, col, status) {
  const cell = sheet.getRange(row, col);
  cell.setHorizontalAlignment("center").setFontWeight("bold");
  if (status === "active") {
    cell.setBackground(COLORS.greenBg).setFontColor(COLORS.greenText);
  } else if (status === "inactive") {
    cell.setBackground(COLORS.yellowBg).setFontColor(COLORS.yellowText);
  } else if (status === "blocked") {
    cell.setBackground(COLORS.redBg).setFontColor(COLORS.redText);
  }
}

// ── 新資料列自動格式化（appendRow 後呼叫）────────────
function styleNewRow(sheet, sheetName) {
  const rowNum = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  formatDataRow(sheet, sheetName, rowNum, lastCol);
}
