// --- CONFIGURATION ---
const SPREADSHEET_ID = '10UrESr-Qnl43PLGFGvvb4tMczCp5JPq71BnwdultrWI'; // Spreadsheet ID ของคุณ
const UPLOAD_FOLDER_ID = '1MIAxWlFog4zHecxm5u2EGNE_peRCNnVk'; // โฟลเดอร์เก็บรูปภาพ
const SHEET_INCOME = 'Income';
const SHEET_WITHDRAW = 'Withdrawals';
const ADMIN_PHONE_NUMBER = '0811606998'; // เบอร์แอดมิน

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setTitle('Connect Center App')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- SETUP DATABASE ---
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const incomeHeaders = ["Date Time", "Name", "Phone", "Email", "Service Type", "Amount", "Commission (10%)", "Channel", "Proof URL", "Status", "Transaction ID"];
  createSheetIfNotExists(ss, SHEET_INCOME, incomeHeaders, 10);
  const withdrawHeaders = ["Date Time", "Name", "Phone", "Email", "Bank", "Account No", "Amount", "Remaining Balance", "Status", "Transaction ID"];
  createSheetIfNotExists(ss, SHEET_WITHDRAW, withdrawHeaders, 9);
}

function createSheetIfNotExists(ss, sheetName, headers, statusColIndex) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#4285F4").setFontColor("#FFFFFF").setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, statusColIndex, 999, 1).setDataValidation(rule);
    sheet.getRange("A2:A").setNumberFormat("d/M/yyyy, H:mm:ss"); 
    if (sheetName === SHEET_INCOME) sheet.getRange("F2:G").setNumberFormat("#,##0.00");
    else sheet.getRange("G2:H").setNumberFormat("#,##0.00");
  }
}

function formatTransactionId(dateObj) {
  if (!dateObj) return "-";
  try {
    const dateStr = Utilities.formatDate(new Date(dateObj), "GMT+7", "d/M/yyyy, H:mm:ss");
    const cleanStr = dateStr.replace(/[\/\,\:\s]/g, '');
    return "TX" + cleanStr;
  } catch (e) { return "ERROR"; }
}

// --- EMAIL SYSTEM (Updated to support status) ---
function sendNotificationEmail(data, type, statusText) {
  const recipient = data.email;
  const adminEmail = Session.getActiveUser().getEmail(); 
  const subject = `แจ้งผลการทำรายการ${type} (${statusText})`;
  const phoneStr = String(data.phone || "").replace(/'/g, '');
  const maskedPhone = phoneStr.length > 6 ? phoneStr.substring(0, 2) + "-XXXX-" + phoneStr.substring(phoneStr.length - 4) : phoneStr;
  const now = Utilities.formatDate(new Date(), "GMT+7", "dd/01/yyyy HH:mm:ss");

  // Determine color based on status
  const statusColor = statusText.includes("สำเร็จ") ? "#10B981" : "#F59E0B"; // Green or Orange

  const htmlBody = `
    <div style="font-family: 'Prompt', sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 10px;">
      <h2 style="color: #4285F4;">Connect Center Notification</h2>
      <p>เรียน ผู้ใช้โทรศัพท์มือถือหมายเลข ${maskedPhone}</p>
      <p><strong>เรื่อง</strong> แจ้งผลการทำรายการ${type} (${statusText})</p>
      <hr style="border: 0; border-top: 1px solid #eee;">
      <p>ตามที่ คุณได้ทำรายการบันทึกข้อมูลผ่านบริการ C SHOP โดยมีรายละเอียด ดังนี้</p>
      <table style="width: 100%; border-collapse: collapse; margin-top: 10px;">
        <tr><td style="padding: 5px; color: #555;">วันที่ทำรายการ:</td><td style="padding: 5px; font-weight: bold;">${now}</td></tr>
        <tr><td style="padding: 5px; color: #555;">ชื่อ-สกุล:</td><td style="padding: 5px; font-weight: bold;">${data.name}</td></tr>
        <tr><td style="padding: 5px; color: #555;">อีเมล์ที่ทำรายการ:</td><td style="padding: 5px; font-weight: bold;">${recipient || '-'}</td></tr>
        <tr><td style="padding: 5px; color: #555;">ทำรายการ:</td><td style="padding: 5px; font-weight: bold;">${type}</td></tr>
        <tr><td style="padding: 5px; color: #555;">จำนวนเงิน:</td><td style="padding: 5px; font-weight: bold;">${parseFloat(data.amount).toLocaleString('en-US', {minimumFractionDigits: 2})} บาท</td></tr>
        <tr><td style="padding: 5px; color: #555;">สถานะ:</td><td style="padding: 5px; font-weight: bold; color: ${statusColor};">${statusText}</td></tr>
      </table>
      <div style="background-color: #f9f9f9; padding: 15px; margin-top: 20px; border-radius: 5px; font-size: 14px;">
        ทางบริการขอเรียนให้ทราบว่า ระบบได้ดำเนินการ${type}เรียบร้อยแล้ว (${statusText}) ทั้งนี้ คุณสามารถตรวจสอบผลของการทำรายการได้ ที่เมนูตรวจสอบข้อมูล
      </div>
      <br>
      <p>ขอแสดงความนับถือ<br>บริการ : C SHOP<br>บริหารโดย Connect Center</p>
    </div>
  `;

  try {
    if(recipient) MailApp.sendEmail({ to: recipient, cc: adminEmail, subject: subject, htmlBody: htmlBody });
    else MailApp.sendEmail({ to: adminEmail, subject: subject + " (Admin Only - No User Email)", htmlBody: htmlBody });
  } catch (e) { Logger.log("Email Error: " + e.toString()); }
}

// --- IMAGE UPLOAD HELPER ---
function uploadImageToDrive(base64Data, fileName) {
  if (!base64Data || base64Data === "") return "";
  try {
    const folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
    const split = base64Data.split('base64,');
    const type = split[0].split(':')[1].split(';')[0];
    const data = Utilities.base64Decode(split[1]);
    const blob = Utilities.newBlob(data, type, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/uc?export=view&id=" + file.getId();
  } catch (e) {
    return "Error: " + e.toString();
  }
}

// --- INCOME LOGIC ---
function saveIncomeData(formObject) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_INCOME);
    const amount = parseFloat(formObject.amount) || 0;
    const commission = amount * 0.10;
    const txId = formatTransactionId(new Date());
    
    // Upload Image
    let imageUrl = "";
    if (formObject.imageFile) {
      const fileName = `Slip_${txId}_${formObject.phone}.jpg`;
      imageUrl = uploadImageToDrive(formObject.imageFile, fileName);
    }

    const rowData = [new Date(), `${formObject.firstName} ${formObject.lastName}`, "'" + formObject.phone, formObject.email, formObject.serviceType, amount, commission, formObject.channel, imageUrl, false, txId];
    sheet.appendRow(rowData);
    sheet.getRange(sheet.getLastRow(), 10).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()).setValue(false);
    
    // Send Email "Waiting for Inspection"
    let balanceInfo = getUserBalanceInfo(formObject.phone); // Get current balance
    const emailData = {
      name: `${formObject.firstName} ${formObject.lastName}`,
      phone: formObject.phone,
      email: formObject.email,
      amount: commission,
      balance: balanceInfo.remainingBalance
    };
    sendNotificationEmail(emailData, 'บันทึกรายได้', 'รอตรวจสอบ');

    return { success: true };
  } catch (e) { return { success: false, message: e.toString() }; }
}

// --- WITHDRAW LOGIC ---
function getUserBalanceInfo(phone) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const incomeSheet = ss.getSheetByName(SHEET_INCOME);
  const withdrawSheet = ss.getSheetByName(SHEET_WITHDRAW);
  const incomeData = incomeSheet.getDataRange().getValues();
  const withdrawData = withdrawSheet.getDataRange().getValues();

  let totalCommission = 0, totalWithdrawn = 0, userName = "", withdrawCountToday = 0, pendingTransaction = null;
  const todayStart = new Date(); todayStart.setHours(0,0,0,0);

  for (let i = 1; i < incomeData.length; i++) {
    let rowPhone = String(incomeData[i][2]).replace(/'/g, '');
    let status = incomeData[i][9];
    if (rowPhone === phone) {
      userName = incomeData[i][1];
      if (status === true || status === "TRUE" || status === "ดำเนินการแล้ว") totalCommission += parseFloat(incomeData[i][6] || 0);
    }
  }
  for (let i = 1; i < withdrawData.length; i++) {
    let rowPhone = String(withdrawData[i][2]).replace(/'/g, '');
    let rowDate = new Date(withdrawData[i][0]);
    let status = withdrawData[i][8];
    if (rowPhone === phone) {
      let amount = parseFloat(withdrawData[i][6] || 0);
      totalWithdrawn += amount;
      if (rowDate >= todayStart) withdrawCountToday++;
      if (status !== true && status !== "TRUE" && status !== "โอนแล้ว") pendingTransaction = { date: formatDate(rowDate), id: formatTransactionId(rowDate), amount: amount };
    }
  }
  return { found: userName !== "", name: userName || "ไม่ระบุ", remainingBalance: totalCommission - totalWithdrawn, withdrawCountToday: withdrawCountToday, pendingTransaction: pendingTransaction };
}

function saveWithdrawalData(formObject) {
  try {
    const balanceInfo = getUserBalanceInfo(formObject.withdrawPhone);
    const requestAmount = parseFloat(formObject.withdrawAmount);
    
    // Custom Error for Pending
    if (balanceInfo.pendingTransaction) {
      const msg = `ไม่สามารถดำเนินการถอนได้<br>เนื่องจากมีรายการถอน ครั้งก่อนหน้านี้ค้างอยู่ในระบบ<br>หากรายการก่อนหน้า อนุมัติแล้ว ระบบจะแจ้งเตือนผ่านอีเมล์อีกครั้งค่ะ<br>เลขที่รายการ ${balanceInfo.pendingTransaction.id}`;
      return { success: false, message: msg };
    }
    
    if (requestAmount < 300) return { success: false, message: "ยอดถอนขั้นต่ำ 300 บาท" };
    if (requestAmount > 20000) return { success: false, message: "ยอดถอนเกิน 20,000 บาท" };
    if (requestAmount > balanceInfo.remainingBalance) return { success: false, message: "ยอดเงินคงเหลือไม่เพียงพอ" };
    if (balanceInfo.withdrawCountToday >= 3) return { success: false, message: "ครบโควตาถอน 3 ครั้งต่อวันแล้ว" };

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_WITHDRAW);
    const txId = formatTransactionId(new Date());
    let userEmail = "";
    const incomeData = ss.getSheetByName(SHEET_INCOME).getDataRange().getValues();
    for(let i=1; i<incomeData.length; i++) if(String(incomeData[i][2]).replace(/'/g, '') == formObject.withdrawPhone) { userEmail = incomeData[i][3]; break; }

    sheet.appendRow([new Date(), balanceInfo.name, "'" + formObject.withdrawPhone, userEmail, formObject.bankSelect, "'" + formObject.accountNumber, requestAmount, balanceInfo.remainingBalance - requestAmount, false, txId]);
    sheet.getRange(sheet.getLastRow(), 9).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()).setValue(false);
    
    // Send Email "Waiting for Inspection"
    const emailData = {
      name: balanceInfo.name,
      phone: formObject.withdrawPhone,
      email: userEmail,
      amount: requestAmount,
      balance: balanceInfo.remainingBalance - requestAmount
    };
    sendNotificationEmail(emailData, 'ถอนรายได้', 'รอตรวจสอบ');

    return { success: true, message: "ทำรายการสำเร็จ" };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getHistoryData(contactInfo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const incomeData = ss.getSheetByName(SHEET_INCOME).getDataRange().getValues();
  const withdrawData = ss.getSheetByName(SHEET_WITHDRAW).getDataRange().getValues();
  let incomeList = [], withdrawList = [], totalApprovedIncome = 0, totalWithdrawn = 0, withdrawCountToday = 0, pendingCount = 0;
  let serviceStats = {}, dailyIncome = {};
  const todayStart = new Date(); todayStart.setHours(0,0,0,0);

  for (let i = 1; i < incomeData.length; i++) {
    let rowPhone = String(incomeData[i][2]).replace(/'/g, '');
    let rowEmail = incomeData[i][3];
    let status = incomeData[i][9];
    if (rowPhone === contactInfo || rowEmail === contactInfo) {
      let isApproved = (status === true || status === "TRUE" || status === "ดำเนินการแล้ว");
      let comm = parseFloat(incomeData[i][6] || 0);
      let dateRaw = new Date(incomeData[i][0]);
      let dateKey = Utilities.formatDate(dateRaw, "GMT+7", "dd/MM");
      incomeList.push({ date: formatDate(dateRaw), service: incomeData[i][4], commission: comm, status: isApproved ? "ดำเนินการแล้ว" : "รอตรวจสอบ" });
      if (isApproved) { totalApprovedIncome += comm; dailyIncome[dateKey] = (dailyIncome[dateKey] || 0) + comm; }
      if (incomeData[i][4]) serviceStats[incomeData[i][4]] = (serviceStats[incomeData[i][4]] || 0) + 1;
    }
  }
  for (let i = 1; i < withdrawData.length; i++) {
    let rowPhone = String(withdrawData[i][2]).replace(/'/g, '');
    let rowEmail = withdrawData[i][3];
    let status = withdrawData[i][8];
    let rowDate = new Date(withdrawData[i][0]);
    if (rowPhone === contactInfo || rowEmail === contactInfo) {
      let isPaid = (status === true || status === "TRUE" || status === "โอนแล้ว");
      withdrawList.push({ date: formatDate(rowDate), bank: withdrawData[i][4], amount: withdrawData[i][6], status: isPaid ? "โอนแล้ว" : "รอดำเนินการ", id: formatTransactionId(rowDate) });
      totalWithdrawn += parseFloat(withdrawData[i][6] || 0);
      if (!isPaid) pendingCount++;
      if (rowDate >= todayStart) withdrawCountToday++;
    }
  }
  return { incomeList: incomeList.reverse(), withdrawList: withdrawList.reverse(), summary: { totalIncome: totalApprovedIncome, totalWithdrawn: totalWithdrawn, remaining: totalApprovedIncome - totalWithdrawn, withdrawCountToday: withdrawCountToday, pendingCount: pendingCount }, charts: { services: serviceStats, daily: dailyIncome } };
}

// --- ADMIN SYSTEM ---
function adminLogin(phone, email) {
  if (phone === ADMIN_PHONE_NUMBER) return { success: true, message: "Login Successful" };
  return { success: false, message: "Access Denied" };
}

function getAdminDashboardData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const incomeData = ss.getSheetByName(SHEET_INCOME).getDataRange().getValues().slice(1);
  const withdrawData = ss.getSheetByName(SHEET_WITHDRAW).getDataRange().getValues().slice(1);

  let totalRevenue = 0, totalOrders = incomeData.length;
  let uniqueCustomers = new Set();
  let dailyStats = {}; 
  let allIncome = [], allWithdraw = [];

  incomeData.forEach((row, index) => {
    let status = row[9];
    let amount = parseFloat(row[5] || 0);
    let comm = parseFloat(row[6] || 0);
    if (status === true || status === "TRUE") totalRevenue += amount;
    uniqueCustomers.add(row[2]);
    
    let dateStr = formatDateKey(new Date(row[0]));
    if(!dailyStats[dateStr]) dailyStats[dateStr] = { income: 0, withdraw: 0 };
    dailyStats[dateStr].income += amount;

    allIncome.push({
      id: index + 2, type: 'Income', date: formatDate(row[0]), name: row[1], phone: row[2], email: row[3], service: row[4], amount: amount, comm: comm, channel: row[7], proof: row[8], status: status
    });
  });

  withdrawData.forEach((row, index) => {
    let amount = parseFloat(row[6] || 0);
    let dateStr = formatDateKey(new Date(row[0]));
    if(!dailyStats[dateStr]) dailyStats[dateStr] = { income: 0, withdraw: 0 };
    dailyStats[dateStr].withdraw += amount;

    allWithdraw.push({
      id: index + 2, type: 'Withdraw', date: formatDate(row[0]), name: row[1], phone: row[2], email: row[3], bank: row[4], acc: row[5], amount: amount, status: row[8]
    });
  });

  let sortedDates = Object.keys(dailyStats).sort();
  let chartLabels = sortedDates.map(d => d.substring(5)); 
  let chartIncome = sortedDates.map(d => dailyStats[d].income);
  let chartWithdraw = sortedDates.map(d => dailyStats[d].withdraw);

  return {
    stats: { revenue: totalRevenue, orders: totalOrders, customers: uniqueCustomers.size, conversion: "3.2%" },
    charts: { labels: chartLabels, income: chartIncome, withdraw: chartWithdraw },
    tables: { income: allIncome.reverse(), withdraw: allWithdraw.reverse() }
  };
}

function updateTransactionStatus(type, rowId, newStatus, currentData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = type === 'Income' ? SHEET_INCOME : SHEET_WITHDRAW;
  const sheet = ss.getSheetByName(sheetName);
  const statusCol = type === 'Income' ? 10 : 9;
  
  sheet.getRange(rowId, statusCol).setValue(newStatus === 'Approve');
  
  if (newStatus === 'Approve') {
    let phone = String(currentData.phone).replace(/'/g, '');
    let balanceInfo = getUserBalanceInfo(phone);
    let emailData = { name: currentData.name, phone: phone, email: currentData.email, amount: type === 'Income' ? currentData.comm : currentData.amount, balance: balanceInfo.remainingBalance };
    let transactionType = type === 'Income' ? 'บันทึกรายได้' : 'ถอนรายได้';
    sendNotificationEmail(emailData, transactionType, 'สำเร็จ');
  }
  return { success: true };
}

function formatDate(date) { try { return Utilities.formatDate(new Date(date), "GMT+7", "dd/MM/yyyy HH:mm"); } catch (e) { return ""; } }
function formatDateKey(date) { try { return Utilities.formatDate(new Date(date), "GMT+7", "yyyy-MM-dd"); } catch (e) { return "0000-00-00"; } }
