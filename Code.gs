// --- CONFIGURATION ---
const GEMINI_API_KEY = "YOUR_GEMINI_API_KEY_HERE"; // ใส่ Gemini API Key ที่นี่
const SPREADSHEET_ID = "18XEyKLNW1mguv0ukaz714u9jg1k87bWlUzDJ58ILnJg"; // ปล่อยว่างไว้ถ้าสคริปต์ผูกกับ Sheet นี้อยู่แล้ว หรือใส่ ID ถ้าแยกไฟล์

// --- MAIN ---
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบเบิกพัสดุโรงเรียน (GAS)')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- INITIAL SETUP ---
function setupSheets() {
  const ss = getSS();
  const sheets = [
    { name: 'Inventory', headers: ['id', 'code', 'name', 'unit', 'price', 'quantity', 'min_quantity', 'updatedAt'] },
    { name: 'Requests', headers: ['id', 'userId', 'userName', 'department', 'items_json', 'reason', 'status', 'date', 'actionBy'] },
    { name: 'Users', headers: ['id', 'username', 'password', 'name', 'role', 'department', 'uid', 'lastLogin'] },
    { name: 'Settings', headers: ['key', 'value'] }
  ];

  sheets.forEach(s => {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      
      // Add Default Data
      if(s.name === 'Users') {
        sheet.appendRow([uuid(), 'admin', '1234', 'Admin Demo', 'admin', 'IT', '', new Date()]);
        sheet.appendRow([uuid(), 'teacher', '1234', 'ครูสมศรี', 'teacher', 'ภาษาไทย', '', new Date()]);
      }
      if(s.name === 'Inventory') {
        sheet.appendRow([uuid(), 'P001', 'กระดาษ A4', 'รีม', 120, 100, 10, new Date()]);
        sheet.appendRow([uuid(), 'P002', 'ปากกาแดง', 'ด้าม', 15, 50, 5, new Date()]);
      }
      if(s.name === 'Settings') {
        sheet.appendRow(['config', JSON.stringify({ schoolName: 'โรงเรียนตัวอย่าง (GAS)', theme: 'teal', darkMode: 'system' })]);
      }
    }
  });
}

// --- DATA HANDLERS ---

function getData() {
  try {
    const ss = getSS();
    
    // ตรวจสอบว่ามี Sheet ครบไหม ถ้าไม่มีให้ Setup ก่อน
    if (!ss.getSheetByName('Inventory')) {
      setupSheets();
    }

    const data = {
      inventory: sheetToJSON(ss.getSheetByName('Inventory')),
      requests: sheetToJSON(ss.getSheetByName('Requests')),
      users: sheetToJSON(ss.getSheetByName('Users')),
      settings: getSettings(ss)
    };
    
    // สำคัญ: แปลงเป็น Text ก่อนส่ง เพื่อป้องกัน error เรื่อง Date object หรือ format
    return JSON.stringify(data); 
    
  } catch (e) {
    // ถ้า Error ให้ส่งข้อความ Error กลับไป
    Logger.log("Error in getData: " + e.toString());
    return JSON.stringify({ error: e.toString() });
  }
}
// ฟังก์ชันสำหรับ Test ดู Log ฝั่ง Server (กด Run ตัวนี้ในหน้า Editor เพื่อเช็ค)
function testGetData() {
  const result = getData();
  console.log(result); // ดูค่าใน Execution Log ด้านล่าง
}

// --- CRUD OPERATIONS ---

// 1. Inventory
function saveInventoryItem(item) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Inventory');
  const data = sheet.getDataRange().getValues();
  
  let savedItem = { ...item };
  savedItem.updatedAt = new Date();

  if (item.id) { // Update
    const rowIndex = data.findIndex(r => r[0] === item.id);
    if (rowIndex > 0) {
      const row = rowIndex + 1;
      // ['id', 'code', 'name', 'unit', 'price', 'quantity', 'min_quantity', 'updatedAt']
      sheet.getRange(row, 2, 1, 7).setValues([[
        item.code, item.name, item.unit, item.price, item.quantity, item.min_quantity, savedItem.updatedAt
      ]]);
    }
  } else { // Create
    savedItem.id = uuid();
    sheet.appendRow([
      savedItem.id, item.code, item.name, item.unit, item.price, item.quantity, item.min_quantity, savedItem.updatedAt
    ]);
  }
  
  // Return only the saved item
  return JSON.stringify(savedItem);
}

function deleteInventoryItem(id) {
  deleteRow('Inventory', id);
  return JSON.stringify({ id: id, status: 'deleted' });
}

// 2. Requests
function saveRequest(req) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Requests');
  const newId = uuid();
  const now = new Date();
  
  const savedReq = { ...req, id: newId, status: 'pending', date: now, actionBy: '' };
  
  // ['id', 'userId', 'userName', 'department', 'items_json', 'reason', 'status', 'date', 'actionBy']
  sheet.appendRow([
    newId, req.userId, req.userName, req.department, JSON.stringify(req.items), req.reason, 'pending', now, ''
  ]);
  
  return JSON.stringify(savedReq);
}

function updateRequestStatus(id, status, actionBy) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Wait up to 30 seconds
    
    const ss = getSS();
    const sheet = ss.getSheetByName('Requests');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(r => r[0] === id);
    
    if (rowIndex > 0) {
      const row = rowIndex + 1;
      
      // If approved, deduct stock with validation
      if (status === 'approved') {
        const items = JSON.parse(data[rowIndex][4]); // items_json
        const invSheet = ss.getSheetByName('Inventory');
        const invData = invSheet.getDataRange().getValues();
        
        // 1. Validate Stock First
        for (const item of items) {
          const invRowIndex = invData.findIndex(r => r[0] === item.itemId);
          if (invRowIndex > 0) {
            const currentQty = Number(invData[invRowIndex][5]);
            if (currentQty < item.qty) {
              throw new Error(`สินค้า "${invData[invRowIndex][2]}" ไม่พอ (เหลือ ${currentQty}, ต้องการ ${item.qty})`);
            }
          } else {
             throw new Error(`ไม่พบสินค้า ID: ${item.itemId}`);
          }
        }
        
        // 2. Deduct Stock (Only if validation passed)
        items.forEach(item => {
          const invRowIndex = invData.findIndex(r => r[0] === item.itemId);
          if (invRowIndex > 0) {
            const currentQty = Number(invData[invRowIndex][5]);
            const newQty = currentQty - item.qty;
            invSheet.getRange(invRowIndex + 1, 6).setValue(newQty);
          }
        });
      }
      
      // Update Status
      sheet.getRange(row, 7).setValue(status); // Status Column
      sheet.getRange(row, 9).setValue(actionBy); // ActionBy Column
    }
    
    return JSON.stringify({ id, status, actionBy });
    
  } catch (e) {
    Logger.log("Error in updateRequestStatus: " + e.toString());
    throw e; // Re-throw to client
  } finally {
    lock.releaseLock();
  }
}

// 3. Users
function saveUser(user) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Users');
  
  const savedUser = { ...user };
  if(!savedUser.id) savedUser.id = uuid();
  savedUser.lastLogin = new Date(); // Just for init
  
  // ['id', 'username', 'password', 'name', 'role', 'department', 'uid', 'lastLogin']
  sheet.appendRow([
    savedUser.id, user.username, user.password, user.name, user.role, user.department, uuid(), savedUser.lastLogin
  ]);
  
  return JSON.stringify(savedUser);
}

function deleteUser(id) {
  deleteRow('Users', id);
  return JSON.stringify({ id: id, status: 'deleted' });
}

function updateUserLogin(id, uid) {
  // In GAS, we might not strictly need UID binding like Firebase, but let's keep logic simple
  // This is a placeholder for compatibility
  return true; 
}

// 4. Settings
function saveSettings(config) {
  const ss = getSS();
  const sheet = ss.getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[0] === 'config');
  
  if (rowIndex >= 0) {
    sheet.getRange(rowIndex + 1, 2).setValue(JSON.stringify(config));
  } else {
    sheet.appendRow(['config', JSON.stringify(config)]);
  }
  return JSON.stringify(config);
}

// --- GEMINI AI ---
function callGeminiAPI(prompt) {
  if (!GEMINI_API_KEY || GEMINI_API_KEY === "YOUR_GEMINI_API_KEY_HERE") {
    return "กรุณาตั้งค่า API Key ในไฟล์ Code.gs ก่อนใช้งาน AI";
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }]
  };

  try {
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload)
    };
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    return data.candidates[0].content.parts[0].text;
  } catch (e) {
    return "เกิดข้อผิดพลาด: " + e.toString();
  }
}

// --- HELPERS ---
function getSS() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function uuid() {
  return Utilities.getUuid();
}

function getInventoryData() {
  const data = getSheetData('Inventory');
  return data.slice(1).map(r => ({
    id: r[0], name: r[1], category: r[2], unit: r[3], price: r[4], quantity: r[5], minStock: r[6], image: r[7]
  }));
}

function getRequestsData() {
  const data = getSheetData('Requests');
  return data.slice(1).map(r => ({
    id: r[0], userId: r[1], userName: r[2], userRole: r[3], items: JSON.parse(r[4]), date: r[5], status: r[6], note: r[7], actionBy: r[8]
  })).reverse();
}

function getUsersData() {
  const data = getSheetData('Users');
  return data.slice(1).map(r => ({
    id: r[0], username: r[1], password: r[2], name: r[3], role: r[4], department: r[5]
  }));
}

function getSheetData(sheetName) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet.getDataRange().getValues() : [];
}

function deleteRow(sheetName, id) {
  const ss = getSS();
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  const rowIndex = data.findIndex(r => r[0] === id);
  if (rowIndex > 0) {
    sheet.deleteRow(rowIndex + 1);
  }
}

function sheetToJSON(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      // Parse JSON columns automatically
      if (h.includes('_json') && row[i]) {
        try { obj[h.replace('_json', '')] = JSON.parse(row[i]); } catch(e) { obj[h] = []; }
        // Keep original too for safety? No, simplify.
      } else {
        obj[h] = row[i];
      }
    });
    // Convert Dates to ISO String for React
    if (obj.date instanceof Date) obj.date = { seconds: obj.date.getTime() / 1000 };
    if (obj.updatedAt instanceof Date) obj.updatedAt = { seconds: obj.updatedAt.getTime() / 1000 };
    return obj;
  });
}

function getSettings(ss) {
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const configRow = data.find(r => r[0] === 'config');
  if (configRow && configRow[1]) {
    try { return JSON.parse(configRow[1]); } catch(e) { return {}; }
  }
  return { schoolName: 'School System', theme: 'teal', darkMode: 'system' };
}