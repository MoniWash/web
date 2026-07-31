
const CONFIG = {
  SPREADSHEET_ID: '1StVCQyVb-kXne0TjPpslXDifQBW4IeacS-K2yFGERSA',
  TZ: 'Asia/Bangkok',
  SESSION_HOURS: 24,
  SHEETS: {
    USERS: 'Users',
    MACHINES: 'Machines',
    BOOKINGS: 'Bookings',
    SETTINGS: 'Settings',
    NOTIFICATIONS: 'Notifications',
    PAYMENTS: 'Payments',
    LOGS: 'Logs',
    TIME_SLOTS: 'TimeSlots'
  }
};

function doGet() {
  setupSystem();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('ระบบจองเครื่องซักผ้า')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ss_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function sh_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต ' + name);
  return sh;
}

function setupSystem() {
  const ss = ss_();
  const defs = [
    [CONFIG.SHEETS.USERS, ['id','name','email','passwordHash','role','phone','active','createdAt','profilePhotoUrl']],
    [CONFIG.SHEETS.MACHINES, ['id','name','status','floor','zone','price','duration','active','note','updatedAt']],
    [CONFIG.SHEETS.BOOKINGS, ['id','userId','userName','machineId','machineName','date','startTime','endTime','price','paymentMethod','paymentStatus','bookingStatus','slipUrl','note','createdAt','updatedAt']],
    [CONFIG.SHEETS.SETTINGS, ['key','value','label','updatedAt']],
    [CONFIG.SHEETS.NOTIFICATIONS, ['id','userId','title','message','type','isRead','createdAt']],
    [CONFIG.SHEETS.PAYMENTS, ['id','bookingId','userId','method','amount','status','slipUrl','verifiedBy','verifiedAt','createdAt']],
    [CONFIG.SHEETS.LOGS, ['id','userId','action','detail','createdAt']],
    [CONFIG.SHEETS.TIME_SLOTS, ['id','startTime','endTime','active','sortOrder','note','createdAt']]
  ];

  defs.forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#0d6efd').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, headers.length);
    }
  });

  ensureColumn_(CONFIG.SHEETS.USERS, 'profilePhotoUrl');
  ensureMachineSchema_();
  ensureSettingsSchema_();
  repairSettings_();
  seedSettings_();
  seedAdmin_();
  seedMachines_();
  repairMachines_();
  seedTimeSlots_();
  migrateMachineNames_();
  return {ok:true};
}

function ensureColumn_(sheetName, headerName) {
  const sh = sh_(sheetName);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
  if (!headers.includes(headerName)) {
    sh.getRange(1,lastCol+1).setValue(headerName).setFontWeight('bold').setBackground('#0d6efd').setFontColor('#ffffff');
  }
}




function ensureSettingsSchema_() {
  const sh = sh_(CONFIG.SHEETS.SETTINGS);
  const required = ['key','value','label','updatedAt'];
  const lastCol = Math.max(sh.getLastColumn(), required.length);
  const current = sh.getRange(1,1,1,lastCol).getValues()[0].map(v => clean_(v));
  required.forEach((h, i) => {
    if (current[i] !== h) sh.getRange(1, i + 1).setValue(h);
  });
  sh.getRange(1,1,1,required.length)
    .setFontWeight('bold').setBackground('#0d6efd').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function settingKey_(v) {
  return String(v == null ? '' : v).trim();
}

function repairSettings_() {
  const sh = sh_(CONFIG.SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return;

  const byKey = {};
  values.slice(1).forEach((r, index) => {
    const key = settingKey_(r[0]);
    if (!key) return;
    const value = r[1] == null ? '' : r[1];
    const label = clean_(r[2]) || key;
    const updatedAt = r[3] || '';
    // เก็บค่าที่มีข้อมูลล่าสุด ป้องกันแถวซ้ำที่ว่างมาทับค่าที่แอดมินบันทึกไว้
    if (!byKey[key] || String(value).trim() !== '') {
      byKey[key] = [key, value, label, updatedAt];
    }
  });

  const rows = Object.keys(byKey).map(k => byKey[k]);
  const expectedRows = values.length - 1;
  const hasBadKey = values.slice(1).some(r => String(r[0] == null ? '' : r[0]) !== settingKey_(r[0]));
  if (rows.length !== expectedRows || hasBadKey) {
    sh.getRange(2,1,Math.max(sh.getMaxRows()-1,1),4).clearContent();
    if (rows.length) sh.getRange(2,1,rows.length,4).setValues(rows);
  }
}

function publicImageUrl_(url) {
  const raw = clean_(url);
  if (!raw) return '';
  // รองรับลิงก์แชร์ Google Drive และแปลงเป็นลิงก์รูปที่ <img> เปิดได้โดยตรง
  let m = raw.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (!m) m = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w1000';
  return raw;
}

function ensureMachineSchema_() {
  const sh = sh_(CONFIG.SHEETS.MACHINES);
  const required = ['id','name','status','floor','zone','price','duration','active','note','updatedAt'];
  const lastCol = Math.max(sh.getLastColumn(), required.length);
  const current = sh.getRange(1,1,1,lastCol).getValues()[0].map(v => clean_(v));

  // ชีตเก่าที่มีหัวตารางคนละชื่อ ให้ยึดตำแหน่งคอลัมน์เดิมและเปลี่ยนหัวตารางเป็นมาตรฐาน
  required.forEach((h, i) => {
    if (current[i] !== h) sh.getRange(1, i + 1).setValue(h);
  });
  sh.getRange(1,1,1,required.length)
    .setFontWeight('bold').setBackground('#0d6efd').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function machineActive_(v) {
  if (v === '' || v == null) return true;
  if (v === true || v === 1) return true;
  if (v === false || v === 0) return false;
  const t = String(v).trim().toLowerCase();
  if (['false','0','no','n','off','disabled','inactive','ปิด','ปิดใช้งาน','ไม่ใช้งาน'].includes(t)) return false;
  if (['true','1','yes','y','on','active','enabled','เปิด','เปิดใช้งาน','ใช้งาน'].includes(t)) return true;
  // ข้อมูลจากรุ่นเก่าที่ไม่ตรงรูปแบบ ไม่ควรถูกตีความว่าปิดเครื่อง
  return true;
}


function normalizeMachineStatus_(value) {
  const t = clean_(value).toLowerCase();
  const map = {
    'available':'available', 'ว่าง':'available', 'พร้อมใช้งาน':'available', 'พร้อม':'available',
    'reserved':'reserved', 'จองแล้ว':'reserved', 'ถูกจอง':'reserved',
    'in_use':'in_use', 'in use':'in_use', 'กำลังใช้งาน':'in_use', 'ใช้งานอยู่':'in_use',
    'maintenance':'maintenance', 'ซ่อมบำรุง':'maintenance', 'กำลังซ่อม':'maintenance', 'ชำรุด':'maintenance',
    'disabled':'disabled', 'ปิดใช้งาน':'disabled', 'ไม่ใช้งาน':'disabled'
  };
  return map[t] || 'available';
}

function machineRows_() {
  return table_(sh_(CONFIG.SHEETS.MACHINES))
    .filter(m => clean_(m.id) || clean_(m.name))
    .map(m => ({
      id: String(m.id || ''),
      name: String(m.name || ''),
      status: normalizeMachineStatus_(m.status),
      floor: String(m.floor || ''),
      zone: String(m.zone || ''),
      price: Number(m.price || 0),
      duration: Number(m.duration || 0),
      active: machineActive_(m.active),
      note: String(m.note || '')
    }));
}

function customerMachineStats_(machines) {
  const active = (machines || []).filter(m => m.active === true);
  return {
    total: active.length,
    available: active.filter(m => m.status === 'available').length,
    inUse: active.filter(m => m.status === 'in_use').length,
    reserved: active.filter(m => m.status === 'reserved').length,
    maintenance: active.filter(m => m.status === 'maintenance').length
  };
}

function seedSettings_() {
  const sh = sh_(CONFIG.SHEETS.SETTINGS);
  const existing = table_(sh);
  const defaults = [
    ['systemName','ระบบจองเครื่องซักผ้า','ชื่อระบบ'],
    ['storeLogoUrl','','ลิงก์โลโก้ร้าน'],
    ['storeLogoEnabled','false','เปิดแสดงโลโก้ร้าน'],
    ['defaultPrice','50','ราคาต่อรอบ'],
    ['defaultDuration','50','ระยะเวลาต่อรอบ (นาที)'],
    ['promptpayId','','เบอร์พร้อมเพย์'],
    ['promptpayName','','ชื่อบัญชีพร้อมเพย์'],
    ['promptpayQrUrl','','ลิงก์รูป QR พร้อมเพย์'],
    ['cashEnabled','true','เปิดรับเงินสด'],
    ['promptpayEnabled','true','เปิดรับพร้อมเพย์'],
    ['bookingAdvanceDays','30','จองล่วงหน้าได้กี่วัน'],
    ['openTime','00:00','เวลาเปิด'],
    ['closeTime','23:59','เวลาปิด']
  ];
  const keys = new Set(existing.map(r => String(r.key)));
  defaults.forEach(([key,value,label]) => {
    if (!keys.has(key)) sh.appendRow([key,value,label,new Date()]);
  });
}

function seedAdmin_() {
  const sh = sh_(CONFIG.SHEETS.USERS);
  const users = table_(sh);
  if (!users.some(u => String(u.role) === 'admin')) {
    sh.appendRow([
      uid_('USR'),
      'ผู้ดูแลระบบ',
      'admin@laundry.local',
      hash_('123456'),
      'admin',
      '',
      true,
      new Date(),
      ''
    ]);
  }
}

function seedMachines_() {
  const sh = sh_(CONFIG.SHEETS.MACHINES);
  // ห้ามใช้ getLastRow() เพียงอย่างเดียว เพราะแถวว่าง/สูตรค้างอาจทำให้คิดว่ามีข้อมูลแล้ว
  const existing = table_(sh).filter(m => clean_(m.id) || clean_(m.name));
  if (existing.length > 0) return;

  // ล้างแถวข้อมูลเดิมที่ว่างหรือผิดรูปแบบ แต่คงหัวตารางไว้
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 10)).clearContent();
  }

  const settings = settings_();
  const rows = [];
  for (let i=1;i<=10;i++) {
    rows.push([
      uid_('M'),
      'เครื่องซักผ้า 20 Kg - ' + i,
      'available',
      'ชั้น 1',
      i <= 5 ? 'โซน A' : 'โซน B',
      Number(settings.defaultPrice || 50),
      Number(settings.defaultDuration || 50),
      true,
      '',
      new Date()
    ]);
  }
  sh.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function repairMachines_() {
  const sh = sh_(CONFIG.SHEETS.MACHINES);
  const rows = table_(sh);
  const settings = settings_();
  const validStatuses = ['available','reserved','in_use','maintenance','disabled'];

  rows.forEach((m, index) => {
    // ข้ามแถวที่ว่างจริง
    if (!clean_(m.id) && !clean_(m.name)) return;

    const id = clean_(m.id) || uid_('M');
    const name = clean_(m.name) || ('เครื่องซักผ้า 20 Kg - ' + (index + 1));
    const status = normalizeMachineStatus_(m.status);
    const floor = clean_(m.floor) || 'ชั้น 1';
    const zone = clean_(m.zone) || (index < 5 ? 'โซน A' : 'โซน B');
    const price = Number(m.price) > 0 ? Number(m.price) : Number(settings.defaultPrice || 50);
    const duration = Number(m.duration) > 0 ? Number(m.duration) : Number(settings.defaultDuration || 50);
    // ข้อมูลรุ่นเก่ามักไม่มีค่า active ให้ถือว่าเปิดใช้งาน
    const active = machineActive_(m.active);
    const note = clean_(m.note || '');
    const updatedAt = m.updatedAt || new Date();

    sh.getRange(m._row, 1, 1, 10).setValues([[
      id, name, status, floor, zone, price, duration, active, note, updatedAt
    ]]);
  });
}



function migrateMachineNames_() {
  const sh = sh_(CONFIG.SHEETS.MACHINES);
  const rows = table_(sh);
  rows.forEach((m, index) => {
    const name = String(m.name || '');
    if (/^เครื่อง\s*\d+$/.test(name)) {
      sh.getRange(m._row,2).setValue('เครื่องซักผ้า 20 Kg - ' + (index + 1));
    }
  });
}

function seedTimeSlots_() {
  const sh = sh_(CONFIG.SHEETS.TIME_SLOTS);
  if (sh.getLastRow() > 1) return;
  const settings = settings_();
  const duration = Number(settings.defaultDuration || 50);
  for (let h = 0; h < 24; h++) {
    const start = String(h).padStart(2,'0') + ':00';
    const endMinutes = h * 60 + duration;
    const end = hhmm_(endMinutes % 1440);
    sh.appendRow([uid_('SLT'), start, end, true, h + 1, 'ช่วงเวลาเริ่มต้น', new Date()]);
  }
}

function table_(sh) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(r => r.some(v => v !== '')).map((r, i) => {
    const o = {_row:i+2};
    headers.forEach((h,j) => o[h]=r[j]);
    return o;
  });
}

function settings_() {
  const rows = table_(sh_(CONFIG.SHEETS.SETTINGS));
  const out = {};
  rows.forEach(r => {
    const key = settingKey_(r.key);
    if (!key) return;
    const value = r.value == null ? '' : r.value;
    // เมื่อมีคีย์ซ้ำ อย่าให้แถวว่างทับค่าที่บันทึกแล้ว
    if (!(key in out) || String(value).trim() !== '') out[key] = value;
  });
  return out;
}

function publicConfig() {
  setupSystem();
  const s = settings_();
  return {
    systemName: s.systemName || 'ระบบจองเครื่องซักผ้า',
    storeLogoUrl: publicImageUrl_(s.storeLogoUrl || ''),
    storeLogoOriginalUrl: s.storeLogoUrl || '',
    storeLogoEnabled: String(s.storeLogoEnabled) === 'true' && !!clean_(s.storeLogoUrl),
    promptpayId: s.promptpayId || '',
    promptpayName: s.promptpayName || '',
    promptpayQrUrl: publicImageUrl_(s.promptpayQrUrl || ''),
    promptpayQrOriginalUrl: s.promptpayQrUrl || '',
    cashEnabled: String(s.cashEnabled) !== 'false',
    promptpayEnabled: String(s.promptpayEnabled) !== 'false',
    bookingAdvanceDays: Number(s.bookingAdvanceDays || 30),
    openTime: s.openTime || '00:00',
    closeTime: s.closeTime || '23:59'
  };
}

function registerUser(data) {
  setupSystem();
  const name = clean_(data.name);
  const email = clean_(data.email).toLowerCase();
  const password = String(data.password || '');
  const phone = clean_(data.phone || '');
  if (!name || !email || password.length < 6) throw new Error('กรุณากรอกชื่อ อีเมล และรหัสผ่านอย่างน้อย 6 ตัว');
  const sh = sh_(CONFIG.SHEETS.USERS);
  const users = table_(sh);
  if (users.some(u => String(u.email).toLowerCase() === email)) throw new Error('อีเมลนี้ถูกใช้งานแล้ว');
  const id = uid_('USR');
  sh.appendRow([id,name,email,hash_(password),'user',phone,true,new Date(),'']);
  log_(id,'register','สมัครสมาชิก');
  return login({email,password});
}


function forgotPassword(email) {
  setupSystem();
  email = clean_(email).toLowerCase();
  if (!email) throw new Error('กรุณากรอกอีเมล');
  const sh = sh_(CONFIG.SHEETS.USERS);
  const user = table_(sh).find(u => String(u.email).toLowerCase() === email && bool_(u.active));
  // Return a neutral response to avoid exposing whether an email exists.
  if (!user) return {ok:true,message:'หากอีเมลนี้มีอยู่ในระบบ ระบบจะส่งรหัสผ่านชั่วคราวให้'};
  const temp = String(Math.floor(100000 + Math.random() * 900000));
  sh.getRange(user._row,4).setValue(hash_(temp));
  MailApp.sendEmail({
    to: email,
    subject: 'รหัสผ่านชั่วคราว - ระบบจองเครื่องซักผ้า',
    htmlBody:
      '<div style="font-family:Arial,sans-serif">' +
      '<h2>ระบบจองเครื่องซักผ้า</h2>' +
      '<p>รหัสผ่านชั่วคราวของคุณคือ</p>' +
      '<div style="font-size:28px;font-weight:bold;letter-spacing:5px">' + temp + '</div>' +
      '<p>กรุณาเข้าสู่ระบบแล้วติดต่อผู้ดูแลเพื่อเปลี่ยนรหัสผ่าน</p>' +
      '</div>'
  });
  log_(user.id,'forgot_password','ส่งรหัสผ่านชั่วคราว');
  return {ok:true,message:'ส่งรหัสผ่านชั่วคราวไปยังอีเมลแล้ว'};
}

function login(data) {
  setupSystem();
  const email = clean_(data.email).toLowerCase();
  const passwordHash = hash_(String(data.password || ''));
  const user = table_(sh_(CONFIG.SHEETS.USERS)).find(u =>
    String(u.email).toLowerCase() === email &&
    String(u.passwordHash) === passwordHash &&
    bool_(u.active)
  );
  if (!user) throw new Error('อีเมลหรือรหัสผ่านไม่ถูกต้อง');
  const token = createSession_(user.id);
  log_(user.id,'login','เข้าสู่ระบบ');
  return {token, user:safeUser_(user)};
}

function logout(token) {
  CacheService.getScriptCache().remove('SESSION_' + token);
  return {ok:true};
}

function me(token) {
  const user = auth_(token);
  return safeUser_(user);
}

function createSession_(userId) {
  const token = Utilities.getUuid().replace(/-/g,'') + new Date().getTime();
  CacheService.getScriptCache().put('SESSION_' + token, userId, CONFIG.SESSION_HOURS * 3600);
  return token;
}

function auth_(token, role) {
  if (!token) throw new Error('กรุณาเข้าสู่ระบบ');
  const userId = CacheService.getScriptCache().get('SESSION_' + token);
  if (!userId) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  const user = table_(sh_(CONFIG.SHEETS.USERS)).find(u => String(u.id) === String(userId) && bool_(u.active));
  if (!user) throw new Error('ไม่พบบัญชีผู้ใช้');
  if (role && String(user.role) !== role) throw new Error('ไม่มีสิทธิ์ใช้งานส่วนนี้');
  return user;
}

function safeUser_(u) {
  return {id:u.id,name:u.name,email:u.email,role:u.role,phone:u.phone || '',profilePhotoUrl:u.profilePhotoUrl || ''};
}

function getDashboard(token) {
  setupSystem();
  const user = auth_(token);
  const machines = machineRows_();
  const bookings = table_(sh_(CONFIG.SHEETS.BOOKINGS));
  const mine = user.role === 'admin' ? bookings : bookings.filter(b => String(b.userId) === String(user.id));
  const notifications = getNotifications(token);
  const today = fmtDate_(new Date());
  return {
    user: safeUser_(user),
    machines: machines,
    stats: customerMachineStats_(machines),
    todayBookings: mine.filter(b => String(b.date) === today).length,
    // แปลงข้อมูลการจองก่อนส่งกลับ เพื่อไม่ให้ Date object ทำให้ google.script.run คืนค่า null
    upcoming: mine.filter(b => ['pending','confirmed'].includes(String(b.bookingStatus)))
      .sort((a,b)=>String(a.date+a.startTime).localeCompare(String(b.date+b.startTime)))
      .slice(0,5)
      .map(bookingObj_),
    notifications: notifications.slice(0,5)
  };
}

function getCustomerMachineStats(token) {
  setupSystem();
  auth_(token);
  const machines = machineRows_();
  return {stats: customerMachineStats_(machines), machines: machines, serverTime: String(new Date())};
}

function getMachines(token) {
  setupSystem();
  auth_(token);
  return machineRows_();
}

function getAvailableSlots(token, data) {
  auth_(token);
  const date = String(data.date || '');
  const machineId = String(data.machineId || '');
  if (!date || !machineId) throw new Error('กรุณาเลือกวันที่และเครื่อง');
  const machine = table_(sh_(CONFIG.SHEETS.MACHINES)).find(m => String(m.id) === machineId && machineActive_(m.active));
  if (!machine) throw new Error('ไม่พบเครื่องซักผ้า');
  if (String(machine.status) === 'maintenance' || String(machine.status) === 'disabled') return [];

  const bookings = table_(sh_(CONFIG.SHEETS.BOOKINGS)).filter(b =>
    String(b.machineId) === machineId &&
    String(b.date) === date &&
    !['cancelled','expired','completed'].includes(String(b.bookingStatus))
  );

  let timeSlots = table_(sh_(CONFIG.SHEETS.TIME_SLOTS));
  if (!timeSlots.length) {
    seedTimeSlots_();
    timeSlots = table_(sh_(CONFIG.SHEETS.TIME_SLOTS));
  }

  const result = timeSlots
    .filter(slot => bool_(slot.active))
    .sort((a,b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0))
    .map(slot => {
      const startTime = normalizeTime_(slot.startTime);
      const endTime = normalizeTime_(slot.endTime);
      if (!startTime || !endTime) return null;
      const busy = bookings.some(b =>
        normalizeTime_(b.startTime) === startTime && normalizeTime_(b.endTime) === endTime
      );
      return {id:String(slot.id || ''),startTime:String(startTime),endTime:String(endTime),available:!busy,note:String(slot.note || '')};
    })
    .filter(Boolean);
  return result || [];
}
function createBooking(token, data) {
  const user = auth_(token);
  const machineId = String(data.machineId || '');
  const date = String(data.date || '');
  const startTime = normalizeTime_(data.startTime || '');
  const selectedEndTime = normalizeTime_(data.endTime || '');
  const paymentMethod = String(data.paymentMethod || '');
  if (!machineId || !date || !startTime || !['cash','promptpay'].includes(paymentMethod)) {
    throw new Error('ข้อมูลการจองไม่ครบถ้วน');
  }
  const machine = table_(sh_(CONFIG.SHEETS.MACHINES)).find(m => String(m.id)===machineId && machineActive_(m.active));
  if (!machine) throw new Error('ไม่พบเครื่องซักผ้า');
  if (['maintenance','disabled'].includes(String(machine.status))) throw new Error('เครื่องนี้ไม่พร้อมให้บริการ');
  const duration = Number(machine.duration || 50);
  const endTime = selectedEndTime || hhmm_(minutes_(startTime)+duration);
  const conflict = table_(sh_(CONFIG.SHEETS.BOOKINGS)).some(b =>
    String(b.machineId)===machineId && String(b.date)===date &&
    !['cancelled','expired','completed'].includes(String(b.bookingStatus)) &&
    overlaps_(startTime,endTime,String(b.startTime),String(b.endTime))
  );
  if (conflict) throw new Error('ช่วงเวลานี้มีผู้จองแล้ว กรุณาเลือกเวลาอื่น');

  const id = uid_('BK');
  const price = Number(machine.price || 50);
  const paymentStatus = paymentMethod === 'cash' ? 'pending_cash' : 'awaiting_slip';
  const bookingStatus = 'pending';
  const now = new Date();
  sh_(CONFIG.SHEETS.BOOKINGS).appendRow([
    id,user.id,user.name,machine.id,machine.name,date,startTime,endTime,price,
    paymentMethod,paymentStatus,bookingStatus,'',clean_(data.note || ''),now,now
  ]);
  sh_(CONFIG.SHEETS.PAYMENTS).appendRow([
    uid_('PAY'),id,user.id,paymentMethod,price,paymentStatus,'','','',now
  ]);
  notify_(user.id,'สร้างรายการจองแล้ว',
    `${machine.name} วันที่ ${date} เวลา ${startTime} - ${endTime}`,
    'booking');
  log_(user.id,'create_booking',id);
  return {ok:true,id,price,endTime};
}

function uploadSlip(token, data) {
  const user = auth_(token);
  const bookingId = String(data.bookingId || '');
  const booking = table_(sh_(CONFIG.SHEETS.BOOKINGS)).find(b => String(b.id)===bookingId);
  if (!booking) throw new Error('ไม่พบรายการจอง');
  if (user.role !== 'admin' && String(booking.userId)!==String(user.id)) throw new Error('ไม่มีสิทธิ์');
  if (!data.base64 || !data.mimeType) throw new Error('กรุณาแนบสลิป');

  const bytes = Utilities.base64Decode(String(data.base64).split(',').pop());
  const blob = Utilities.newBlob(bytes, data.mimeType, data.fileName || ('slip_'+bookingId+'.jpg'));
  const folder = getSlipFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = file.getUrl();

  const bsh = sh_(CONFIG.SHEETS.BOOKINGS);
  bsh.getRange(booking._row,13).setValue(url);
  bsh.getRange(booking._row,11).setValue('checking');
  bsh.getRange(booking._row,16).setValue(new Date());

  const payment = table_(sh_(CONFIG.SHEETS.PAYMENTS)).find(p => String(p.bookingId)===bookingId);
  if (payment) {
    const psh = sh_(CONFIG.SHEETS.PAYMENTS);
    psh.getRange(payment._row,6).setValue('checking');
    psh.getRange(payment._row,7).setValue(url);
  }
  notify_(user.id,'รับสลิปแล้ว','กำลังรอแอดมินตรวจสอบรายการ '+bookingId,'payment');
  return {ok:true,url};
}

function getSlipFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SLIP_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  const folder = DriveApp.createFolder('Laundry Booking Slips');
  props.setProperty('SLIP_FOLDER_ID',folder.getId());
  return folder;
}

function getMyBookings(token) {
  const user = auth_(token);
  const rows = table_(sh_(CONFIG.SHEETS.BOOKINGS));
  const list = user.role === 'admin' ? rows : rows.filter(b => String(b.userId)===String(user.id));
  return (list || []).sort((a,b)=>new Date(b.createdAt).getTime()-new Date(a.createdAt).getTime()).map(bookingObj_);
}

function cancelBooking(token, bookingId) {
  const user = auth_(token);
  const booking = table_(sh_(CONFIG.SHEETS.BOOKINGS)).find(b=>String(b.id)===String(bookingId));
  if (!booking) throw new Error('ไม่พบรายการจอง');
  if (user.role !== 'admin' && String(booking.userId)!==String(user.id)) throw new Error('ไม่มีสิทธิ์');
  if (!['pending','confirmed'].includes(String(booking.bookingStatus))) throw new Error('รายการนี้ไม่สามารถยกเลิกได้');
  const sh = sh_(CONFIG.SHEETS.BOOKINGS);
  sh.getRange(booking._row,12).setValue('cancelled');
  sh.getRange(booking._row,16).setValue(new Date());
  notify_(booking.userId,'ยกเลิกรายการจองแล้ว','รายการ '+booking.id+' ถูกยกเลิก','booking');
  return {ok:true};
}

function getNotifications(token) {
  const user = auth_(token);
  return table_(sh_(CONFIG.SHEETS.NOTIFICATIONS))
    .filter(n => String(n.userId)===String(user.id) || String(n.userId)==='ALL')
    .sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)))
    .map(n=>({id:n.id,title:n.title,message:n.message,type:n.type,isRead:bool_(n.isRead),createdAt:n.createdAt ? String(n.createdAt) : ''}));
}

function markNotificationRead(token, id) {
  const user = auth_(token);
  const n = table_(sh_(CONFIG.SHEETS.NOTIFICATIONS)).find(x=>String(x.id)===String(id));
  if (n && (String(n.userId)===String(user.id) || String(n.userId)==='ALL')) {
    sh_(CONFIG.SHEETS.NOTIFICATIONS).getRange(n._row,6).setValue(true);
  }
  return {ok:true};
}

function updateProfile(token, data) {
  const user = auth_(token);
  const sh = sh_(CONFIG.SHEETS.USERS);
  sh.getRange(user._row,2).setValue(clean_(data.name || user.name));
  sh.getRange(user._row,6).setValue(clean_(data.phone || ''));
  return me(token);
}


function uploadProfilePhoto(token, data) {
  const user = auth_(token);
  if (!data || !data.base64 || !data.mimeType) throw new Error('กรุณาเลือกรูปภาพ');
  const bytes = Utilities.base64Decode(String(data.base64).split(',').pop());
  if (bytes.length > 5 * 1024 * 1024) throw new Error('รูปภาพต้องไม่เกิน 5 MB');
  const blob = Utilities.newBlob(bytes, data.mimeType, data.fileName || ('profile_' + user.id + '.jpg'));
  const folder = getProfileFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const url = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  const sh = sh_(CONFIG.SHEETS.USERS);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String);
  const col = headers.indexOf('profilePhotoUrl') + 1;
  if (!col) throw new Error('ไม่พบคอลัมน์รูปโปรไฟล์ กรุณารัน setupSystem');
  sh.getRange(user._row,col).setValue(url);
  log_(user.id,'upload_profile_photo',file.getId());
  return {ok:true,url:user.profilePhotoUrl=url,user:safeUser_(Object.assign({},user,{profilePhotoUrl:url}))};
}

function getProfileFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PROFILE_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch(e) {} }
  const folder = DriveApp.createFolder('Laundry Booking Profile Photos');
  props.setProperty('PROFILE_FOLDER_ID', folder.getId());
  return folder;
}

function getMyPayments(token) {
  const user = auth_(token);
  return table_(sh_(CONFIG.SHEETS.PAYMENTS))
    .filter(p => user.role === 'admin' || String(p.userId) === String(user.id))
    .sort((a,b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())
    .map(p => ({id:p.id,bookingId:p.bookingId,method:p.method,amount:Number(p.amount||0),status:p.status,slipUrl:p.slipUrl||'',createdAt:p.createdAt ? String(p.createdAt) : ''}));
}

function changePassword(token, data) {
  const user = auth_(token);
  const current = String(data.currentPassword || '');
  const next = String(data.newPassword || '');
  if (String(user.passwordHash) !== hash_(current)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
  if (next.length < 6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัว');
  sh_(CONFIG.SHEETS.USERS).getRange(user._row,4).setValue(hash_(next));
  log_(user.id,'change_password','เปลี่ยนรหัสผ่าน');
  return {ok:true};
}

/* ===================== ADMIN ===================== */

function adminOverview(token) {
  auth_(token,'admin');
  const machines = table_(sh_(CONFIG.SHEETS.MACHINES));
  const bookings = table_(sh_(CONFIG.SHEETS.BOOKINGS));
  const users = table_(sh_(CONFIG.SHEETS.USERS));
  const today = fmtDate_(new Date());
  const todayBookings = bookings.filter(b=>String(b.date)===today);
  const revenue = todayBookings.filter(b=>String(b.paymentStatus)==='paid')
    .reduce((s,b)=>s+Number(b.price||0),0);
  const days = [];
  for(let i=6;i>=0;i--){
    const d = new Date(); d.setDate(d.getDate()-i);
    const key = fmtDate_(d);
    days.push({date:key,count:bookings.filter(b=>String(b.date)===key).length});
  }
  return {
    totalMachines:machines.filter(m=>machineActive_(m.active)).length,
    inUse:machines.filter(m=>String(m.status)==='in_use' && machineActive_(m.active)).length,
    available:machines.filter(m=>String(m.status)==='available' && machineActive_(m.active)).length,
    maintenance:machines.filter(m=>String(m.status)==='maintenance' && machineActive_(m.active)).length,
    todayBookings:todayBookings.length,
    todayRevenue:revenue,
    users:users.filter(u=>bool_(u.active)).length,
    chart:days,
    latest:bookings.slice(-8).reverse().map(bookingObj_)
  };
}

function adminGetUsers(token) {
  auth_(token,'admin');
  return table_(sh_(CONFIG.SHEETS.USERS)).map(u=>({
    id:u.id,name:u.name,email:u.email,role:u.role,phone:u.phone||'',active:bool_(u.active),createdAt:u.createdAt
  }));
}

function adminToggleUser(token, userId, active) {
  const admin = auth_(token,'admin');
  if (String(admin.id)===String(userId) && !active) throw new Error('ไม่สามารถปิดบัญชีตนเอง');
  const u = table_(sh_(CONFIG.SHEETS.USERS)).find(x=>String(x.id)===String(userId));
  if (!u) throw new Error('ไม่พบผู้ใช้');
  sh_(CONFIG.SHEETS.USERS).getRange(u._row,7).setValue(Boolean(active));
  return {ok:true};
}


function adminSaveUser(token, data) {
  const admin = auth_(token,'admin');
  data = data || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sh_(CONFIG.SHEETS.USERS);
    const users = table_(sh);
    const id = clean_(data.id || '');
    const name = clean_(data.name || '');
    const email = clean_(data.email || '').toLowerCase();
    const phone = clean_(data.phone || '');
    const role = String(data.role || 'user') === 'admin' ? 'admin' : 'user';
    const password = String(data.password || '');
    const active = data.active === false || String(data.active).toLowerCase() === 'false' ? false : true;

    if (!name) throw new Error('กรุณากรอกชื่อผู้ใช้');
    if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('กรุณากรอกอีเมลให้ถูกต้อง');
    const duplicate = users.find(u => String(u.email || '').toLowerCase() === email && String(u.id) !== String(id));
    if (duplicate) throw new Error('อีเมลนี้ถูกใช้งานแล้ว');

    if (id) {
      const u = users.find(x => String(x.id) === String(id));
      if (!u) throw new Error('ไม่พบผู้ใช้');
      if (String(admin.id) === String(id) && (!active || role !== 'admin')) {
        throw new Error('ไม่สามารถปิดบัญชีหรือลดสิทธิ์บัญชีที่กำลังใช้งาน');
      }
      if (String(u.role) === 'admin' && role !== 'admin') {
        const adminCount = users.filter(x => String(x.role) === 'admin' && bool_(x.active)).length;
        if (adminCount <= 1) throw new Error('ระบบต้องมีผู้ดูแลอย่างน้อย 1 บัญชี');
      }
      sh.getRange(u._row, 2).setValue(name);
      sh.getRange(u._row, 3).setValue(email);
      if (password) {
        if (password.length < 6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
        sh.getRange(u._row, 4).setValue(hash_(password));
      }
      sh.getRange(u._row, 5).setValue(role);
      sh.getRange(u._row, 6).setValue(phone);
      sh.getRange(u._row, 7).setValue(active);
      log_(admin.id, 'admin_update_user', id);
      return {ok:true,id:id};
    }

    if (password.length < 6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัวอักษร');
    const newId = uid_('USR');
    sh.appendRow([newId,name,email,hash_(password),role,phone,active,new Date(),'']);
    log_(admin.id, 'admin_create_user', newId);
    return {ok:true,id:newId};
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteUser(token, userId) {
  const admin = auth_(token,'admin');
  if (String(admin.id) === String(userId)) throw new Error('ไม่สามารถลบบัญชีที่กำลังใช้งาน');
  const sh = sh_(CONFIG.SHEETS.USERS);
  const users = table_(sh);
  const u = users.find(x => String(x.id) === String(userId));
  if (!u) throw new Error('ไม่พบผู้ใช้');
  if (String(u.role) === 'admin') {
    const adminCount = users.filter(x => String(x.role) === 'admin' && bool_(x.active)).length;
    if (adminCount <= 1) throw new Error('ระบบต้องมีผู้ดูแลอย่างน้อย 1 บัญชี');
  }
  sh.deleteRow(u._row);
  log_(admin.id, 'admin_delete_user', userId);
  return {ok:true};
}

function adminSaveMachine(token, data) {
  auth_(token,'admin');
  const sh = sh_(CONFIG.SHEETS.MACHINES);
  const rows = table_(sh);
  const id = String(data.id || '');
  const values = [
    clean_(data.name || 'เครื่องใหม่'),
    clean_(data.status || 'available'),
    clean_(data.floor || 'ชั้น 1'),
    clean_(data.zone || 'โซน A'),
    Number(data.price || 50),
    Number(data.duration || 50),
    data.active !== false,
    clean_(data.note || ''),
    new Date()
  ];
  if (id) {
    const m = rows.find(x=>String(x.id)===id);
    if (!m) throw new Error('ไม่พบเครื่อง');
    sh.getRange(m._row,2,1,9).setValues([values]);
    return {ok:true,id};
  }
  const newId = uid_('M');
  sh.appendRow([newId].concat(values));
  return {ok:true,id:newId};
}

function adminDeleteMachine(token, id) {
  auth_(token,'admin');
  const m = table_(sh_(CONFIG.SHEETS.MACHINES)).find(x=>String(x.id)===String(id));
  if (!m) throw new Error('ไม่พบเครื่อง');
  sh_(CONFIG.SHEETS.MACHINES).getRange(m._row,8).setValue(false);
  sh_(CONFIG.SHEETS.MACHINES).getRange(m._row,3).setValue('disabled');
  return {ok:true};
}

function adminUpdateBooking(token, bookingId, action) {
  const admin = auth_(token,'admin');
  const b = table_(sh_(CONFIG.SHEETS.BOOKINGS)).find(x=>String(x.id)===String(bookingId));
  if (!b) throw new Error('ไม่พบรายการจอง');
  const map = {
    confirm:{bookingStatus:'confirmed',paymentStatus:'paid',title:'ยืนยันการจองแล้ว'},
    reject:{bookingStatus:'cancelled',paymentStatus:'rejected',title:'การชำระเงินไม่ผ่าน'},
    start:{bookingStatus:'in_use',paymentStatus:String(b.paymentStatus),title:'เริ่มใช้งานแล้ว'},
    complete:{bookingStatus:'completed',paymentStatus:String(b.paymentStatus),title:'ใช้งานเสร็จสิ้นแล้ว'},
    cancel:{bookingStatus:'cancelled',paymentStatus:String(b.paymentStatus),title:'รายการจองถูกยกเลิก'}
  };
  const x = map[action];
  if (!x) throw new Error('คำสั่งไม่ถูกต้อง');
  const sh = sh_(CONFIG.SHEETS.BOOKINGS);
  sh.getRange(b._row,11).setValue(x.paymentStatus);
  sh.getRange(b._row,12).setValue(x.bookingStatus);
  sh.getRange(b._row,16).setValue(new Date());

  const machine = table_(sh_(CONFIG.SHEETS.MACHINES)).find(m=>String(m.id)===String(b.machineId));
  if (machine) {
    const status = action==='start' ? 'in_use' : action==='confirm' ? 'reserved' : 'available';
    sh_(CONFIG.SHEETS.MACHINES).getRange(machine._row,3).setValue(status);
  }
  const p = table_(sh_(CONFIG.SHEETS.PAYMENTS)).find(p=>String(p.bookingId)===String(bookingId));
  if (p && ['confirm','reject'].includes(action)) {
    const psh=sh_(CONFIG.SHEETS.PAYMENTS);
    psh.getRange(p._row,6).setValue(action==='confirm'?'paid':'rejected');
    psh.getRange(p._row,8).setValue(admin.name);
    psh.getRange(p._row,9).setValue(new Date());
  }
  notify_(b.userId,x.title,`${b.machineName} วันที่ ${b.date} เวลา ${b.startTime}`,'booking');
  return {ok:true};
}

function adminGetSettings(token) {
  auth_(token,'admin');
  setupSystem();
  const s = settings_() || {};
  return {
    systemName: s.systemName || 'ระบบจองเครื่องซักผ้า',
    storeLogoUrl: publicImageUrl_(s.storeLogoUrl || ''),
    storeLogoOriginalUrl: s.storeLogoUrl || '',
    storeLogoEnabled: String(s.storeLogoEnabled) === 'true',
    defaultPrice: Number(s.defaultPrice || 50),
    defaultDuration: Number(s.defaultDuration || 50),
    promptpayId: s.promptpayId || '',
    promptpayName: s.promptpayName || '',
    promptpayQrUrl: publicImageUrl_(s.promptpayQrUrl || ''),
    promptpayQrOriginalUrl: s.promptpayQrUrl || '',
    cashEnabled: String(s.cashEnabled) !== 'false',
    promptpayEnabled: String(s.promptpayEnabled) !== 'false',
    openTime: s.openTime || '00:00',
    closeTime: s.closeTime || '23:59'
  };
}

function adminSaveSettings(token, data) {
  auth_(token,'admin');
  setupSystem();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = sh_(CONFIG.SHEETS.SETTINGS);
    const allowed = [
      'systemName','storeLogoUrl','storeLogoEnabled','defaultPrice','defaultDuration','promptpayId','promptpayName',
      'promptpayQrUrl','openTime','closeTime','cashEnabled','promptpayEnabled',
      'bookingAdvanceDays'
    ];
    const rows = table_(sh);
    allowed.forEach(key => {
      if (!Object.prototype.hasOwnProperty.call(data || {}, key)) return;
      let value = data[key];
      if (['systemName','storeLogoUrl','promptpayId','promptpayName','promptpayQrUrl','openTime','closeTime'].includes(key)) {
        value = clean_(value);
      }
      if (['storeLogoEnabled','cashEnabled','promptpayEnabled'].includes(key)) {
        value = String(value) === 'true' || value === true ? 'true' : 'false';
      }
      const matches = rows.filter(x => settingKey_(x.key) === key);
      if (matches.length) {
        // อัปเดตทุกแถวซ้ำ เพื่อไม่ให้ค่าจากแถวเก่ากลับมาทับหลังรีเฟรช
        matches.forEach(r => {
          sh.getRange(r._row,1,1,4).setValues([[key,value,r.label || key,new Date()]]);
        });
      } else {
        sh.appendRow([key,value,key,new Date()]);
      }
    });
    SpreadsheetApp.flush();
    repairSettings_();
    const saved = settings_();
    return {
      ok:true,
      settings:{
        systemName:saved.systemName || 'ระบบจองเครื่องซักผ้า',
        storeLogoUrl:publicImageUrl_(saved.storeLogoUrl || ''),
        storeLogoOriginalUrl:saved.storeLogoUrl || '',
        storeLogoEnabled:String(saved.storeLogoEnabled) === 'true' && !!clean_(saved.storeLogoUrl),
        promptpayId:saved.promptpayId || '',
        promptpayName:saved.promptpayName || '',
        promptpayQrUrl:publicImageUrl_(saved.promptpayQrUrl || ''),
        promptpayQrOriginalUrl:saved.promptpayQrUrl || '',
        cashEnabled:String(saved.cashEnabled) !== 'false',
        promptpayEnabled:String(saved.promptpayEnabled) !== 'false'
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function adminReport(token, fromDate, toDate) {
  auth_(token,'admin');
  const rows = table_(sh_(CONFIG.SHEETS.BOOKINGS)).filter(b=>{
    const d=String(b.date);
    return (!fromDate || d>=fromDate) && (!toDate || d<=toDate);
  });
  const paid = rows.filter(b=>String(b.paymentStatus)==='paid');
  const byMachine = {};
  rows.forEach(b=>byMachine[b.machineName]=(byMachine[b.machineName]||0)+1);
  let topMachine = '-';
  let topCount = 0;
  Object.keys(byMachine).forEach(k=>{ if(byMachine[k]>topCount){topMachine=k;topCount=byMachine[k];} });
  return {
    rows:rows.map(bookingObj_),
    totalBookings:rows.length,
    paidBookings:paid.length,
    revenue:paid.reduce((s,b)=>s+Number(b.price||0),0),
    topMachine,topCount
  };
}



function adminResetTimeSlots(token) {
  auth_(token,'admin');
  const sh = sh_(CONFIG.SHEETS.TIME_SLOTS);
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  seedTimeSlots_();
  return adminGetTimeSlots(token);
}

function adminGetTimeSlots(token) {
  auth_(token,'admin');
  return table_(sh_(CONFIG.SHEETS.TIME_SLOTS))
    .sort((a,b)=>Number(a.sortOrder||0)-Number(b.sortOrder||0))
    .map(s=>({id:s.id,startTime:normalizeTime_(s.startTime),endTime:normalizeTime_(s.endTime),active:bool_(s.active),sortOrder:Number(s.sortOrder||0),note:s.note||''}));
}

function adminSaveTimeSlot(token, data) {
  auth_(token,'admin');
  const startTime = normalizeTime_(data.startTime);
  const endTime = normalizeTime_(data.endTime);
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(startTime) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(endTime)) {
    throw new Error('รูปแบบเวลาไม่ถูกต้อง');
  }
  const sh = sh_(CONFIG.SHEETS.TIME_SLOTS);
  const rows = table_(sh);
  const duplicate = rows.find(r => String(r.id)!==String(data.id||'') && normalizeTime_(r.startTime)===startTime && normalizeTime_(r.endTime)===endTime);
  if (duplicate) throw new Error('มีช่วงเวลานี้อยู่แล้ว');
  const values = [startTime,endTime,data.active!==false,Number(data.sortOrder||rows.length+1),clean_(data.note||''),new Date()];
  if (data.id) {
    const row = rows.find(r=>String(r.id)===String(data.id));
    if (!row) throw new Error('ไม่พบช่วงเวลา');
    sh.getRange(row._row,2,1,6).setValues([values]);
    return {ok:true,id:data.id};
  }
  const id = uid_('SLT');
  sh.appendRow([id].concat(values));
  return {ok:true,id};
}

function adminDeleteTimeSlot(token, id) {
  auth_(token,'admin');
  const row = table_(sh_(CONFIG.SHEETS.TIME_SLOTS)).find(r=>String(r.id)===String(id));
  if (!row) throw new Error('ไม่พบช่วงเวลา');
  sh_(CONFIG.SHEETS.TIME_SLOTS).deleteRow(row._row);
  return {ok:true};
}

/* ===================== HELPERS ===================== */

function bookingObj_(b) {
  return {
    id:b.id,userId:b.userId,userName:b.userName,machineId:b.machineId,machineName:b.machineName,
    date:(b.date instanceof Date ? fmtDate_(b.date) : String(b.date)),startTime:normalizeTime_(b.startTime),endTime:normalizeTime_(b.endTime),price:Number(b.price),
    paymentMethod:b.paymentMethod,paymentStatus:b.paymentStatus,bookingStatus:b.bookingStatus,
    slipUrl:b.slipUrl||'',note:b.note||'',createdAt:b.createdAt ? String(b.createdAt) : ''
  };
}
function notify_(userId,title,message,type){
  sh_(CONFIG.SHEETS.NOTIFICATIONS).appendRow([uid_('NTF'),userId,title,message,type||'info',false,new Date()]);
}
function log_(userId,action,detail){
  sh_(CONFIG.SHEETS.LOGS).appendRow([uid_('LOG'),userId,action,detail,new Date()]);
}
function uid_(prefix){ return prefix+'-'+Utilities.getUuid().split('-')[0].toUpperCase()+'-'+Date.now().toString().slice(-6); }
function hash_(text){
  const raw=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,text,Utilities.Charset.UTF_8);
  return raw.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');
}
function clean_(v){ return String(v==null?'':v).trim(); }
function bool_(v){ return v===true || String(v).toLowerCase()==='true' || String(v)==='1'; }
function fmtDate_(d){ return Utilities.formatDate(new Date(d),CONFIG.TZ,'yyyy-MM-dd'); }
function normalizeTime_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TZ, 'HH:mm');
  }

  const text = String(value == null ? '' : value).trim();
  if (!text) return '';

  const match = text.match(/^(\d{1,2}):(\d{2})/);
  if (match) {
    const hour = Math.max(0, Math.min(23, Number(match[1])));
    const minute = Math.max(0, Math.min(59, Number(match[2])));
    return String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
  }

  const date = new Date(value);
  if (!isNaN(date.getTime())) {
    return Utilities.formatDate(date, CONFIG.TZ, 'HH:mm');
  }

  throw new Error('รูปแบบเวลาไม่ถูกต้อง: ' + text);
}

function minutes_(value){
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getHours() * 60 + value.getMinutes();
  }
  const text = String(value == null ? '' : value).trim();
  const match = text.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return NaN;
  return Number(match[1]) * 60 + Number(match[2]);
}
function hhmm_(m){
  m = Number(m);
  if (!isFinite(m)) throw new Error('รูปแบบเวลาไม่ถูกต้อง');
  return String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0');
}
function overlaps_(a1,a2,b1,b2){ return minutes_(a1)<minutes_(b2) && minutes_(a2)>minutes_(b1); }
