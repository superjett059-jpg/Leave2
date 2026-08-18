/**
 * GOOGLE APPS SCRIPT — ระบบลางาน
 * อัปเดต: อนุมัติอัตโนมัติภายใน 12 ชั่วโมงเมื่อไม่มีการพิจารณา และยกเลิกการส่งเตือนความถี่
 */

// --- ตั้งค่า ---
const SPREADSHEET_ID     = '1MkIDIicBuQZl3tAinhc2grdqssTXyijGQ5aSstGnKpc';
const FOLDER_ID          = '1WlkZrUKygBzwpteKNaZpvGZ5EQrH-cwb';
const LINE_CHANNEL_TOKEN = 'jpJukrtE/SpktkYrUEKupUmD2Cl7+m+leMZdBLzyIwT+IXynA3sWTyOq55QolTEASg/nOdZ+wVbafSxkZ36qOURE70w/k+ueJf5OUdnsbpHuedNNCqwTg6+pJTj+y/wXmWjHsdqluJAljbp2ez5b8AdB04t89/1O/w1cDnyilFU=';
const MANAGER_USER_ID    = 'U4a123429599c6698b9f8df4509170d49';
const MANAGER_USER_ID_2  = 'U65f24825f41a96cd0ab8f8db17226e2f'; // ← จะมาเปลี่ยนตรงนี้ทีหลัง
const WEB_URL            = 'https://rubbertreeindustry059.github.io/Leave/';
const APPROVE_URL        = 'https://rubbertreeindustry059.github.io/Leave/approve.html';

function doGet(e) {
  const action   = e.parameter.action;
  const callback = e.parameter.callback;
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  let responseData;
  if (action === 'getData') {
    responseData = JSON.stringify({
      success: true,
      data: {
        users:  getSheetData(ss, 'Users'),
        leaves: getSheetData(ss, 'Leaves'),
        config: getSheetData(ss, 'Config')[0] || null
      }
    });
  } else {
    responseData = JSON.stringify({ success: true, message: 'ระบบลางาน Online' });
  }
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + responseData + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(responseData).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const response = createResponse({ success: true });
  if (!e || !e.postData || !e.postData.contents) return response;

  // ✅ ดัก LINE User ID อัตโนมัติ
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.events && body.events[0]) {
      const userId = body.events[0].source.userId;
      const ss2 = SpreadsheetApp.openById(SPREADSHEET_ID);
      const log = ss2.getSheetByName('LineLog') || ss2.insertSheet('LineLog');
      log.appendRow([new Date(), userId]);
      return createResponse({ success: true });
    }
  } catch(err) {}

  let params;
  try { params = JSON.parse(e.postData.contents); } catch (err) { return response; }

  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const action = params.action;

  if (action === 'saveLeaves')  { setSheetData(ss, 'Leaves', params.data); }
  if (action === 'saveUsers')   { setSheetData(ss, 'Users',  params.data); }
  if (action === 'saveConfig')  { setSheetData(ss, 'Config', [params.data]); }

  if (action === 'notifyLine') {
    const leave = params.data ? params.data.leave : params.leave;
    if (leave) notifyManager(leave);
  }

  if (action === 'approveLeave') {
    const { leaveId, note } = params.data || {};
    const leaves = getSheetData(ss, 'Leaves');
    const cleanId = function(x) { return String(x || '').replace(/,/g, '').trim(); };
    const idx = leaves.findIndex(l => cleanId(l.id) === cleanId(leaveId));
    if (idx !== -1) {
      leaves[idx].status        = 'อนุมัติแล้ว';
      leaves[idx].manager_name  = 'ผู้จัดการ';
      leaves[idx].reject_reason = note || '';
      leaves[idx].start_date    = cleanDate(leaves[idx].start_date);
      leaves[idx].end_date      = cleanDate(leaves[idx].end_date);
      setSheetData(ss, 'Leaves', leaves);
      notifyEmployee(leaves[idx], 'approve');
    }
    return createResponse({ success: true });
  }

  if (action === 'rejectLeave') {
    const { leaveId, note } = params.data || {};
    const leaves = getSheetData(ss, 'Leaves');
    const cleanId = function(x) { return String(x || '').replace(/,/g, '').trim(); };
    const idx = leaves.findIndex(l => cleanId(l.id) === cleanId(leaveId));
    if (idx !== -1) {
      leaves[idx].status        = 'ปฏิเสธ';
      leaves[idx].manager_name  = 'ผู้จัดการ';
      leaves[idx].reject_reason = note || '';
      leaves[idx].start_date    = cleanDate(leaves[idx].start_date);
      leaves[idx].end_date      = cleanDate(leaves[idx].end_date);
      setSheetData(ss, 'Leaves', leaves);
      notifyEmployee(leaves[idx], 'reject');
    }
    return createResponse({ success: true });
  }

  if (action === 'uploadFile') {
    try {
      const folder = DriveApp.getFolderById(FOLDER_ID);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(params.data.data.split(',')[1]),
        params.data.data.split(',')[0].split(':')[1].split(';')[0],
        params.data.name
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return createResponse({ success: true, url: file.getUrl() });
    } catch(err) {
      return createResponse({ success: false, error: err.toString() });
    }
  }

  return response;
}

function isLineEnabled(ss) {
  try {
    const config = getSheetData(ss, 'Config')[0];
    if (!config || config.lineNotifications === undefined || config.lineNotifications === null) return true;
    const val = String(config.lineNotifications).toLowerCase();
    return val !== 'false';
  } catch (err) {
    return true;
  }
}

function notifyManager(leave) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!isLineEnabled(ss)) return;
  const approveLink = APPROVE_URL + '?id=' + leave.id;
  const msg =
    '📋 ใบลาใหม่เข้ามาครับ!\n' +
    '👤 ชื่อ: '   + leave.employee_name + '\n' +
    '📝 ประเภท: ' + leave.leave_type    + '\n' +
    '📅 วันที่: ' + cleanDate(leave.start_date) + ' - ' + cleanDate(leave.end_date) + '\n' +
    '⏱ จำนวน: '  + formatDaysGS(leave) + '\n' +
    '💡 เหตุผล: ' + (leave.reason || '-') + '\n' +
    (leave.certificate_file_data ? '📎 ไฟล์แนบ: ' + leave.certificate_file_data + '\n' : '') + '\n' +
    '👇 กดลิงก์เพื่ออนุมัติหรือปฏิเสธได้เลยครับ\n' +
    approveLink;

  // ส่งหาทั้งสองคน
  [MANAGER_USER_ID, MANAGER_USER_ID_2].forEach(function(userId) {
    if (!userId || userId === 'WAIT_FOR_ID') return;
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ to: userId, messages: [{ type: 'text', text: msg }] }),
      muteHttpExceptions: true
    });
  });
}

function notifyEmployee(leave, action) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!isLineEnabled(ss)) return;
  const users = getSheetData(ss, 'Users');
  const user  = users.find(u => u.username === leave.employee_username);
  if (!user || !user.line_user_id) return;

  const isApproved = action === 'approve';
  const msg =
    (isApproved ? '✅ ใบลาของคุณได้รับการอนุมัติแล้ว!' : '❌ ใบลาของคุณถูกปฏิเสธ') + '\n' +
    '📝 ประเภท: ' + leave.leave_type + '\n' +
    '📅 วันที่: '  + cleanDate(leave.start_date) + ' - ' + cleanDate(leave.end_date) + '\n' +
    (leave.reject_reason ? '💬 หมายเหตุ: ' + leave.reject_reason + '\n' : '') +
    '\n🔗 ดูรายละเอียด: ' + WEB_URL;

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ to: user.line_user_id, messages: [{ type: 'text', text: msg }] }),
    muteHttpExceptions: true
  });
}

function formatDaysGS(leave) {
  if (leave.is_half_day === 'hourly') {
    const t1 = String(leave.hourly_start || '').trim();
    const t2 = String(leave.hourly_end   || '').trim();
    if (t1 && t2 && /^\d{1,2}:\d{2}/.test(t1) && /^\d{1,2}:\d{2}/.test(t2)) {
      const m1 = parseInt(t1.split(':')[0])*60 + parseInt(t1.split(':')[1]);
      const m2 = parseInt(t2.split(':')[0])*60 + parseInt(t2.split(':')[1]);
      const diff = m2 - m1;
      if (diff > 0) {
        const h = Math.floor(diff/60), m = diff%60;
        return (h > 0 ? h+' ชม.' : '') + (m > 0 ? (h>0?' ':'')+m+' นาที' : '');
      }
    }
  }
  if (leave.is_half_day === 'true' || leave.is_half_day === true) return '0.5 วัน (ครึ่งวัน)';
  return (leave.days || 0) + ' วัน';
}

function cleanDate(val) {
  if (!val) return val;
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  let s = String(val).replace(/"/g, '').trim();
  if (s.includes('T')) s = s.split('T')[0];
  return s;
}

function cleanTime(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return String(val.getHours()).padStart(2,'0') + ':' + String(val.getMinutes()).padStart(2,'0');
  }
  const s = String(val).trim();
  if (!isNaN(parseFloat(s)) && parseFloat(s) >= 0 && parseFloat(s) < 1) {
    const totalMin = Math.round(parseFloat(s) * 24 * 60);
    const h = Math.floor(totalMin / 60), m = totalMin % 60;
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }
  if (s.includes('1899')) {
    const match = s.match(/(\d{2}):(\d{2})/);
    if (match) return match[1] + ':' + match[2];
  }
  if (/^\d{1,2}:\d{2}/.test(s)) return s.substring(0, 5);
  return s;
}

function createResponse(d) {
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}

function getSheetData(ss, sn) {
  let s = ss.getSheetByName(sn);
  if (!s) return [];
  const display = s.getDataRange().getDisplayValues();
  const raw     = s.getDataRange().getValues();
  if (display.length <= 1) return [];
  const h = display[0];
  return display.slice(1).map((row, rowIdx) => {
    const rawRow = raw[rowIdx + 1];
    const o = {};
    h.forEach((k, i) => {
      let v = row[i];
      if (v === '' || v === 'null') { o[k] = null; return; }
      const rawVal = rawRow ? rawRow[i] : null;
      if (k === 'hourly_start' || k === 'hourly_end') {
        if (rawVal instanceof Date) { o[k] = cleanTime(rawVal); }
        else if (typeof rawVal === 'number' && rawVal >= 0 && rawVal < 1) { o[k] = cleanTime(rawVal); }
        else { o[k] = parseDisplayTime(v); }
        return;
      }
      if (rawVal instanceof Date) { o[k] = cleanDate(rawVal); return; }
      if (typeof rawVal === 'number' && rawVal % 1 === 0) { o[k] = String(rawVal); return; }
      o[k] = v;
    });
    return o;
  });
}

function parseDisplayTime(val) {
  if (!val) return '';
  const s = String(val).trim();
  const ampm = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?$/i);
  if (ampm) {
    let h = parseInt(ampm[1]), m = parseInt(ampm[2]);
    const period = ampm[3] ? ampm[3].toUpperCase() : null;
    if (period === 'PM' && h !== 12) h += 12;
    if (period === 'AM' && h === 12) h = 0;
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }
  return s.substring(0, 5);
}

function setSheetData(ss, sn, da) {
  let s = ss.getSheetByName(sn);
  if (!s) s = ss.insertSheet(sn);
  s.clear();
  if (!da || da.length === 0) return;
  
  // รวบรวมหัวตาราง (keys) ทั้งหมดจากทุกแถวข้อมูล
  const keysSet = {};
  da.forEach(function(o) {
    Object.keys(o).forEach(function(key) {
      keysSet[key] = true;
    });
  });
  const k = Object.keys(keysSet);
  
  s.appendRow(k);
  const r = da.map(o => k.map(x => {
    const v = o[x];
    if (v === null || v === undefined) return '';
    return typeof v === 'object' ? JSON.stringify(v) : v;
  }));
  if (r.length > 0) s.getRange(2, 1, r.length, k.length).setValues(r);
}

// ─── ระบบอนุมัติใบลาอัตโนมัติภายใน 12 ชั่วโมง ───
function autoApprovePendingLeaves() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const leaves = getSheetData(ss, 'Leaves');
  const now = Date.now();
  const autoApproveMs = 12 * 60 * 60 * 1000; // 12 ชั่วโมง
  let updated = false;

  leaves.forEach(function(leave) {
    if (leave.status === 'รอการอนุมัติ') {
      const createdAtTime = new Date(leave.created_at).getTime();
      if (isNaN(createdAtTime)) return;

      // ถ้าค้างเกิน 12 ชั่วโมง จะอนุมัติอัตโนมัติ
      if (now - createdAtTime >= autoApproveMs) {
        leave.status        = 'อนุมัติแล้ว';
        leave.manager_name  = 'ผู้จัดการ';
        leave.reject_reason = '';
        leave.start_date    = cleanDate(leave.start_date);
        leave.end_date      = cleanDate(leave.end_date);
        
        // แจ้งเตือนไปยังพนักงาน
        notifyEmployee(leave, 'approve');

        // ✅ แจ้งเตือนผู้จัดการให้ทราบด้วย
        notifyManagerOfAutoApprove(leave);

        updated = true;
      }
    }
  });

  if (updated) {
    setSheetData(ss, 'Leaves', leaves);
  }
}

// แจ้งเตือนผู้จัดการเมื่อมีใบลาอนุมัติอัตโนมัติ
function notifyManagerOfAutoApprove(leave) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (!isLineEnabled(ss)) return;
  const msg =
    '🤖 แจ้งเตือนผู้จัดการ: ระบบได้อนุมัติใบลาอัตโนมัติแล้วครับ!\n' +
    '⚠️ (เนื่องจากไม่มีการดำเนินการอนุมัติ/ปฏิเสธภายใน 12 ชั่วโมง)\n\n' +
    '👤 ชื่อพนักงาน: ' + leave.employee_name + '\n' +
    '📝 ประเภท: ' + leave.leave_type    + '\n' +
    '📅 วันที่: ' + cleanDate(leave.start_date) + ' - ' + cleanDate(leave.end_date) + '\n' +
    '⏱ จำนวน: '  + formatDaysGS(leave) + '\n' +
    '💡 เหตุผล: ' + (leave.reason || '-');

  [MANAGER_USER_ID, MANAGER_USER_ID_2].forEach(function(userId) {
    if (!userId || userId === 'WAIT_FOR_ID') return;
    UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ to: userId, messages: [{ type: 'text', text: msg }] }),
      muteHttpExceptions: true
    });
  });
}

// ─── รันระบบตรวจสอบใบลาค้างเพื่ออนุมัติอัตโนมัติ ───
// (รักษาชื่อฟังก์ชันเดิมไว้เพื่อไม่ให้ทริกเกอร์ที่ตั้งค่าใน GAS เสียหาย)
function checkAndSendPendingReminders() {
  autoApprovePendingLeaves();
}
