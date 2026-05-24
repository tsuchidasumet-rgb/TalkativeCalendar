const scriptProperties = PropertiesService.getScriptProperties();

const CONFIG = {
  LINE_ACCESS_TOKEN: scriptProperties.getProperty('LINE_ACCESS_TOKEN'),
  MY_USER_ID: scriptProperties.getProperty('MY_USER_ID'),
  SHEETS: { 
    GENERAL: "GENERAL", 
    PAYMENTS: "PAYMENTS" 
  }
};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const event = data.events[0];
    if (!event || !event.message || event.message.type !== 'text') return;

    const replyToken = event.replyToken;
    const fullMsg = event.message.text.trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const lines = fullMsg.split('\n');
    let results = [];

    lines.forEach(line => {
      const msg = line.trim();
      if (!msg) return;

      if (msg.includes("สรุป") || msg.includes("เดือนหน้า") || msg.includes("มีไรบ้าง") || msg.includes("มีอะไรบ้าง") || msg.includes("นัด") || msg.includes("จ่าย")) {
        results.push(handleSummary(msg, ss));
      }
      else if (msg.startsWith("ยกเลิก")) {
        results.push(handleCancel(msg, ss));
      }
      // เพิ่มให้รองรับการตรวจจับวันในสัปดาห์
      else if (msg.match(/ทุก(วันที่|วัน)?\s*\d+/) || msg.includes("เดือน") || msg.includes("สิ้นเดือน") || msg.match(/ทุกวัน(จันทร์|อังคาร|พุธ|พฤหัส|ศุกร์|เสาร์|อาทิตย์)/)) {
        results.push(handlePaymentTask(msg, ss));
      }
      else if (msg.match(/(\d{4}\/)?(\d{1,2})\/(\d{1,2})/)) {
        results.push(handleGeneralTask(msg, ss));
      }
      else {
        results.push(`❓ ไม่เข้าใจ: "${msg.substring(0, 10)}..."`);
      }
    });

    replyLine(replyToken, "✅ ดำเนินการเรียบร้อย:\n\n" + results.join('\n\n'));
  } catch (err) { console.error(err.message); }
}

function formatDateToYYYYMMDD(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}

function handleSummary(msg, ss) {
  const now = new Date();
  const todayNoTime = new Date(); todayNoTime.setHours(0,0,0,0);
  let targetMonth = now.getMonth();
  let targetYear = now.getFullYear();

  if (msg.includes("เดือนหน้า")) {
    targetMonth++;
    if (targetMonth > 11) { targetMonth = 0; targetYear++; }
  } else {
    const mMatch = msg.match(/เดือน\s*(\d+)/);
    if (mMatch) targetMonth = parseInt(mMatch[1]) - 1;
  }

  const isShowGeneral = msg.includes("นัด") || msg.includes("ไปไหน") || msg.includes("สรุป") || msg.includes("มีไรบ้าง") || msg.includes("มีอะไรบ้าง");
  const isShowPayment = msg.includes("ค่า") || msg.includes("จ่าย") || msg.includes("ชำระ") || msg.includes("สรุป") || msg.includes("มีไรบ้าง") || msg.includes("มีอะไรบ้าง");

  let summary = `📊 สรุปรายการ [${targetMonth + 1}/${targetYear}]\n`;

  // --- ส่วนนัดหมาย (เหมือนเดิมแต่เรียงเวลา) ---
  if (isShowGeneral) {
    const genData = ss.getSheetByName(CONFIG.SHEETS.GENERAL).getDataRange().getValues();
    let genList = [];
    for (let i = 1; i < genData.length; i++) {
      const d = new Date(genData[i][2]);
      if (d.getMonth() === targetMonth && d.getFullYear() === targetYear && genData[i][4] === "Pending") {
        genList.push({ date: d, task: genData[i][1] });
      }
    }
    genList.sort((a, b) => a.date - b.date);
    let genStrings = genList.map(item => {
      const timeStr = (item.date.getHours() === 0 && item.date.getMinutes() === 0) ? "" : ` [${String(item.date.getHours()).padStart(2, '0')}:${String(item.date.getMinutes()).padStart(2, '0')}]`;
      return `• วันที่ ${item.date.getDate()}${timeStr}: ${item.task}`;
    });
    summary += `\n🗓️ นัดหมายทั่วไป:\n${genStrings.length > 0 ? genStrings.join('\n') : "- ไม่มีนัด"}\n`;
  }

  // --- ส่วนชำระเงิน (เพิ่มยอดคงเหลือ) ---
  if (isShowPayment) {
    const payData = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS).getDataRange().getValues();
    let payList = [];
    let totalYen = 0, totalBaht = 0;
    let remYen = 0, remBaht = 0; // ยอดคงเหลือที่ต้องจ่าย

    for (let i = 1; i < payData.length; i++) {
      const status = payData[i][5];
      if (status === "Cancelled" || status === "Done") continue;

      const savedDate = new Date(payData[i][3]);
      let rawValue = payData[i][2].toString();
      let amt = parseFloat(rawValue.replace(/[^\d.]/g, ''));
      let symbol = rawValue.includes("฿") ? "฿" : "￥";
      const term = payData[i][4] !== "Infinite" ? ` ${payData[i][4]}` : "";

      if (status === "Weekly") {
        let dayOfWeek = savedDate.getDay();
        for (let d = 1; d <= 31; d++) {
          let check = new Date(targetYear, targetMonth, d);
          if (check.getMonth() !== targetMonth) break;
          if (check.getDay() === dayOfWeek) {
            if (symbol === "฿") totalBaht += amt; else totalYen += amt;
            // คำนวณยอดที่เหลือ (ตั้งแต่วันนี้เป็นต้นไป)
            if (check >= todayNoTime) {
              if (symbol === "฿") remBaht += amt; else remYen += amt;
            }
            payList.push({ date: new Date(check), name: payData[i][1], amt: amt, symbol: symbol, term: term });
          }
        }
      } else {
        let d;
        if (status === "Recurring") {
          let lastDayOfTarget = new Date(targetYear, targetMonth + 1, 0).getDate();
          let targetDay = Math.min(savedDate.getDate(), lastDayOfTarget);
          d = new Date(targetYear, targetMonth, targetDay);
        } else { d = savedDate; }

        if (d.getMonth() === targetMonth && d.getFullYear() === targetYear) {
          if (symbol === "฿") totalBaht += amt; else totalYen += amt;
          // คำนวณยอดที่เหลือ (ตั้งแต่วันนี้เป็นต้นไป)
          if (d >= todayNoTime) {
            if (symbol === "฿") remBaht += amt; else remYen += amt;
          }
          payList.push({ date: d, name: payData[i][1], amt: amt, symbol: symbol, term: term });
        }
      }
    }
    payList.sort((a, b) => a.date - b.date);
    let payStrings = payList.map(item => `• วันที่ ${item.date.getDate()}: ${item.name} [${item.amt.toLocaleString()} ${item.symbol}]${item.term}`);

    summary += `\n💰 รายการชำระเงิน:\n${payStrings.length > 0 ? payStrings.join('\n') : "- ไม่มีรายการ"}\n`;
    
    if (totalYen > 0 || totalBaht > 0) {
      summary += `\n💸 ยอดรวมทั้งหมด:`;
      if (totalYen > 0) summary += `\n🇯🇵 เยน: ${totalYen.toLocaleString()} ￥ \n   └ เหลือจ่าย: ${remYen.toLocaleString()} ￥`;
      if (totalBaht > 0) summary += `\n🇹🇭 บาท: ${totalBaht.toLocaleString()} ฿ \n   └ เหลือจ่าย: ${remBaht.toLocaleString()} ฿`;
    }
  }
  return summary;
}

function handlePaymentTask(msg, ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
  const amountMatch = msg.match(/([\d,]+)\s*([a-zA-Z\u0E00-\u0E7F]+)?/);
  const monthMatch = msg.match(/(\d+)\s*เดือน/);
  const isEndOfMonth = msg.includes("สิ้นเดือน");
  const forceNextMonth = msg.includes("เริ่มเดือนหน้า");
  
  // ตรวจจับวันในสัปดาห์
  const weekDayMatch = msg.match(/ทุกวัน(จันทร์|อังคาร|พุธ|พฤหัส|ศุกร์|เสาร์|อาทิตย์)/);
  const dayMap = {'อาทิตย์':0, 'จันทร์':1, 'อังคาร':2, 'พุธ':3, 'พฤหัส':4, 'ศุกร์':5, 'เสาร์':6};
  
  let dayMatch = msg.match(/ทุก(?:วันที่|วัน)?\s*(\d+)/);
  let explicitDateMatch = msg.match(/(\d{4}\/)?(\d{1,2})\/(\d{1,2})/);

  if (!amountMatch) return "❌ ไม่พบยอดเงิน";

  let rawAmount = amountMatch[1].replace(/,/g, "");
  let unit = amountMatch[2] ? amountMatch[2].toLowerCase() : ""; 
  let symbol = (unit === "b" || unit === "บาท" || unit === "thb") ? "฿" : "￥";
  const amountWithSymbol = rawAmount + symbol;

  const name = msg.split(amountMatch[1])[0].replace(/ทุกวันที่|ทุกวัน|ทุก/g, "").trim();
  const now = new Date();
  let startDay, startMonth = now.getMonth(), startYear = now.getFullYear();

  if (weekDayMatch) {
    let targetDayNum = dayMap[weekDayMatch[1]];
    let start = new Date();
    start.setDate(now.getDate() + (targetDayNum + 7 - now.getDay()) % 7);
    if (start <= now && !forceNextMonth) start.setDate(start.getDate() + 7);
    if (forceNextMonth) start.setDate(start.getDate() + 7);
    
    sheet.appendRow([new Date(), name, amountWithSymbol, formatDateToYYYYMMDD(start), "Infinite", "Weekly"]);
    return `♾️ ${name} (ทุกวัน${weekDayMatch[1]}) [${amountWithSymbol}]`;
  }

  // --- Logic เดิมสำหรับรายเดือน/สิ้นเดือน ---
  if (explicitDateMatch) {
    const parts = explicitDateMatch[0].split('/');
    if (parts.length === 3) { startYear = parseInt(parts[0]); startMonth = parseInt(parts[1]) - 1; startDay = parseInt(parts[2]); }
    else { startMonth = parseInt(parts[0]) - 1; startDay = parseInt(parts[1]); if (startMonth < now.getMonth()) startYear++; }
  } else if (isEndOfMonth) {
    startDay = new Date(startYear, startMonth + 1, 0).getDate();
    if (forceNextMonth || new Date(startYear, startMonth, startDay) <= now) startMonth++;
  } else if (dayMatch) {
    startDay = parseInt(dayMatch[1]);
    if (forceNextMonth || new Date(startYear, startMonth, startDay) < now) startMonth++;
  }

  if (monthMatch) {
    const total = parseInt(monthMatch[1]);
    for (let i = 0; i < total; i++) {
      let currentM = startMonth + i;
      let currentY = startYear;
      if (currentM > 11) { currentY += Math.floor(currentM / 12); currentM = currentM % 12; }
      let day = isEndOfMonth ? new Date(currentY, currentM + 1, 0).getDate() : startDay;
      let deadline = new Date(currentY, currentM, Math.min(day, new Date(currentY, currentM + 1, 0).getDate()));
      sheet.appendRow([new Date(), name, amountWithSymbol, formatDateToYYYYMMDD(deadline), `[${i + 1}/${total}]`, "Pending"]);
    }
    return `💰 ${name} (${total} งวด) [${amountWithSymbol}]`;
  } else {
    let day = isEndOfMonth ? new Date(startYear, startMonth + 1, 0).getDate() : startDay;
    let infiniteDate = new Date(startYear, startMonth, day);
    sheet.appendRow([new Date(), name, amountWithSymbol, formatDateToYYYYMMDD(infiniteDate), "Infinite", "Recurring"]);
    return `♾️ ${name} (ทุกเดือน) [${amountWithSymbol}]`;
  }
}

// --- ฟังก์ชันช่วยเหลืออื่นๆ เหมือนเดิม ---
function handleGeneralTask(msg, ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEETS.GENERAL);
  const dateMatch = msg.match(/(\d{4}\/)?(\d{1,2})\/(\d{1,2})/);
  const timeMatch = msg.match(/(\d{1,2})[\.:](\d{2})/);
  const placeMatch = msg.match(/ที่\s*([^\s]+)/);
  const targetDate = getSmartDate(dateMatch[0], timeMatch ? timeMatch[0] : null);
  const place = placeMatch ? placeMatch[1] : "-";
  const task = msg.split(/ที่|\d{1,2}\/|(\d{1,2})[\.:]/)[0].trim();
  const formattedDateTime = Utilities.formatDate(targetDate, "GMT+9", "yyyy/MM/dd HH:mm");
  sheet.appendRow([new Date(), task, formattedDateTime, place, "Pending"]);
  return `🗓️ ${task} ${timeMatch ? '(' + timeMatch[0] + ')' : ''}`;
}

function getSmartDate(dateStr, timeStr) {
  const now = new Date();
  let year = now.getFullYear(), month, day;
  const parts = dateStr.split('/');
  if (parts.length === 3) { year = parseInt(parts[0]); month = parseInt(parts[1]) - 1; day = parseInt(parts[2]); }
  else { month = parseInt(parts[0]) - 1; day = parseInt(parts[1]); if (month < now.getMonth()) year += 1; }
  const date = new Date(year, month, day);
  if (timeStr) { const t = timeStr.replace('.', ':').split(':'); date.setHours(parseInt(t[0]), parseInt(t[1] || 0)); }
  else date.setHours(0, 0, 0, 0);
  return date;
}

function handleCancel(msg, ss) {
  const targetName = msg.replace("ยกเลิก", "").trim();
  if (!targetName) return "❌ ลืมพิมพ์ชื่อรายการครับ";
  let count = 0;
  Object.values(CONFIG.SHEETS).forEach(shName => {
    const sheet = ss.getSheetByName(shName);
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const statusIdx = (shName === CONFIG.SHEETS.GENERAL) ? 4 : 5;
      if (data[i][1].toString().includes(targetName) && (data[i][statusIdx] === "Pending" || data[i][statusIdx] === "Recurring" || data[i][statusIdx] === "Weekly")) {
        sheet.getRange(i + 1, statusIdx + 1).setValue("Cancelled");
        count++;
      }
    }
  });
  return count > 0 ? `🚫 ยกเลิก [${targetName}] เรียบร้อย (${count})` : `🔍 หาชื่อ [${targetName}] ไม่พบครับ`;
}

function checkAndNotify() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  let logs = [];
  
  const paySheet = ss.getSheetByName(CONFIG.SHEETS.PAYMENTS);
  if (paySheet) {
    const payData = paySheet.getDataRange().getValues();
    for (let i = 1; i < payData.length; i++) {
      const status = payData[i][5];
      if (status === "Cancelled" || status === "Done") continue;
      let targetDate;
      const termInfo = payData[i][4] !== "Infinite" ? payData[i][4] : "";
      const savedDate = new Date(payData[i][3]);

      if (status === "Weekly") {
        let dayOfWeek = savedDate.getDay();
        targetDate = new Date(today);
        targetDate.setDate(today.getDate() + (dayOfWeek + 7 - today.getDay()) % 7);
      } else if (status === "Recurring") {
        let lastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
        let targetDay = Math.min(savedDate.getDate(), lastDay);
        targetDate = new Date(today.getFullYear(), today.getMonth(), targetDay);
      } else {
        targetDate = savedDate;
      }
      // ใส่ false ท้ายสุดเพราะไม่ใช่ General Task
      processNotification(payData[i][1], payData[i][2], targetDate, today, logs, termInfo, false);
    }
  }

  const genSheet = ss.getSheetByName(CONFIG.SHEETS.GENERAL);
  if (genSheet) {
    const genData = genSheet.getDataRange().getValues();
    for (let i = 1; i < genData.length; i++) {
      if (genData[i][4] !== "Pending") continue;
      // ใส่ true ท้ายสุดเพราะเป็น General Task
      processNotification(genData[i][1], null, new Date(genData[i][2]), today, logs, "", true);
    }
  }
  if (logs.length > 0) sendPush(logs.join("\n\n"));
}

function processNotification(name, amount, targetDate, today, logs, termInfo, isGeneral) {
  const checkDate = new Date(targetDate);
  checkDate.setHours(0, 0, 0, 0);
  const diff = Math.ceil((checkDate - today) / (1000 * 60 * 60 * 24));
  let displayName = name;

  // 1. แสดงเวลาเฉพาะนัดหมายทั่วไป
  if (isGeneral) {
    const hh = String(targetDate.getHours()).padStart(2, '0');
    const mm = String(targetDate.getMinutes()).padStart(2, '0');
    if (!(hh === "00" && mm === "00")) displayName += ` [${hh}:${mm}]`;
  }

  if (amount) {
    let rawStr = amount.toString();
    let symbol = rawStr.includes("฿") ? "฿" : "￥";
    let amt = parseFloat(rawStr.replace(/[^\d.]/g, ''));
    displayName += ` [${amt.toLocaleString()} ${symbol}]`;
  }

  if (termInfo) displayName += ` ${termInfo}`;

  if (diff === 7) logs.push(`🔔 อีก 7 วัน: ${displayName.trim()}`);
  else if (diff === 1) logs.push(`⚠️ พรุ่งนี้แล้ว!: ${displayName.trim()}`);
  else if (diff === 0) logs.push(`🚨 วันนี้!!: ${displayName.trim()}`);
}

function replyLine(token, text) { callLineAPI("reply", { "replyToken": token, "messages": [{ "type": "text", "text": text }] }); }
function sendPush(text) { callLineAPI("push", { "to": CONFIG.MY_USER_ID, "messages": [{ "type": "text", "text": "📢 แจ้งเตือนตารางงาน\n\n" + text }] }); }
function callLineAPI(endpoint, payload) {
  UrlFetchApp.fetch(`https://api.line.me/v2/bot/message/${endpoint}`, {
    "method": "post",
    "headers": { "Authorization": "Bearer " + CONFIG.LINE_ACCESS_TOKEN, "Content-Type": "application/json" },
    "payload": JSON.stringify(payload)
  });
}