// Google Apps Script Backend for ผังสแตนเชียร์ (คณะสีเหลือง เทพอพอลโล)
// Admin Password: yellowisusmah

function doGet(e) {
    // Handle GET requests (e.g. browser redirects or preflight)
    const action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
    if (action === 'get') {
        return responseJSON(getBookings());
    }
    return responseJSON({ success: true, message: 'Google Apps Script is running.' });
}

function doPost(e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const action = data.action;

        if (action === 'get') {
            return responseJSON(getBookings());
        }
        else if (action === 'reserve') {
            const result = reserveSeat(
                data.seatId,
                data.studentName,
                data.className,
                data.studentNo,
                data.studentPass,
                data.cancelPass
            );
            return responseJSON(result);
        }
        else if (action === 'cancel') {
            const result = cancelSeat(data.seatId, data.cancelPass);
            return responseJSON(result);
        }
        else if (action === 'admin_clear') {
            const result = adminClearSeat(data.adminPass, data.seatId);
            return responseJSON(result);
        }
        else if (action === 'admin_clear_all') {
            const result = adminClearAll(data.adminPass);
            return responseJSON(result);
        }

        return responseJSON({ success: false, message: 'Invalid action' });
    } catch (err) {
        return responseJSON({ success: false, message: err.toString() });
    }
}

function responseJSON(data) {
    return ContentService.createTextOutput(JSON.stringify(data))
        .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Bookings');
    if (!sheet) {
        // Create the Bookings sheet if it doesn't exist
        sheet = ss.insertSheet('Bookings');
        sheet.appendRow(['Seat ID', 'ชื่อ-นามสกุล', 'ชั้นเรียน', 'เลขที่', 'รหัส7หลัก', 'รหัสยกเลิก8หลัก', 'เวลาที่จอง']);
    } else if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Seat ID', 'ชื่อ-นามสกุล', 'ชั้นเรียน', 'เลขที่', 'รหัส7หลัก', 'รหัสยกเลิก8หลัก', 'เวลาที่จอง']);
    }
    return sheet;
}

function isDateValue(val) {
    if (!val) return false;
    if (val instanceof Date) return true;
    const str = val.toString().trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) return true;
    if (/(?:Sun|Mon|Tue|Wed|Thu|Fri|Sat)\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)/i.test(str)) return true;
    if (/GMT[+-]\d{4}/i.test(str) || /เวลาอินโดจีน/i.test(str) || /Indochina Time/i.test(str)) return true;
    return false;
}

function getBookings() {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const bookings = {};

    for (let i = 1; i < data.length; i++) {
        const seatId = data[i][0] ? data[i][0].toString().trim() : '';
        let rawName = data[i][1];
        let name = rawName ? rawName.toString().trim() : '';
        const className = data[i][2] ? data[i][2].toString().trim() : '';
        const studentNo = data[i][3] ? data[i][3].toString().trim() : '';
        const studentPass = data[i][4] ? data[i][4].toString().trim() : '';
        const cancelPass = data[i][5] ? data[i][5].toString().trim() : '';
        let time = data[i][6] ? data[i][6].toString() : '';

        if (isDateValue(rawName)) {
            if (!time) time = name;
            name = '';
        }

        if (seatId && (name || className || studentNo || studentPass || cancelPass || time)) {
            bookings[seatId] = {
                name: name,
                class: className,
                num: studentNo,
                pass: studentPass,
                cancelpass: cancelPass,
                time: time
            };
        }
    }
    return bookings;
}

function reserveSeat(seatId, studentName, className, studentNo, studentPass, cancelPass) {
    var lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000); // Wait up to 10 seconds
    } catch (e) {
        return { success: false, message: 'ระบบกำลังประมวลผลจำนวนมาก กรุณาลองใหม่อีกครั้ง' };
    }

    try {
        var sheet = getSheet();
        var data = sheet.getDataRange().getValues();

        // Check if seat is already booked (inside lock for atomicity)
        for (var i = 1; i < data.length; i++) {
            if (data[i][0] && data[i][0].toString().trim() === seatId && data[i][1]) {
                return { success: false, message: 'ที่นั่ง ' + seatId + ' ถูกจองไปแล้ว!' };
            }
        }

        sheet.appendRow([
            seatId,
            studentName,
            className,
            studentNo,
            studentPass,
            cancelPass,
            new Date().toISOString()
        ]);
        SpreadsheetApp.flush(); // Force write to sheet

        return { success: true, message: 'จองที่นั่ง ' + seatId + ' สำเร็จ!' };
    } finally {
        lock.releaseLock();
    }
}

function cancelSeat(seatId, cancelPass) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: 'ระบบกำลังประมวลผล กรุณาลองใหม่อีกครั้ง' };
    }

    try {
        const sheet = getSheet();
        const data = sheet.getDataRange().getValues();

        for (let i = 1; i < data.length; i++) {
            const currentSeat = data[i][0] ? data[i][0].toString().trim() : '';
            const storedCancelPass = data[i][5] ? data[i][5].toString().trim() : '';

            if ((currentSeat === seatId || !seatId) && storedCancelPass === cancelPass) {
                sheet.deleteRow(i + 1);
                SpreadsheetApp.flush();
                return { success: true, message: 'ยกเลิกการจองที่นั่งสำเร็จ เรียบร้อยแล้ว!' };
            }
        }

        return { success: false, message: 'รหัสยกเลิก (8 หลัก) หรือรหัสที่นั่งไม่ถูกต้อง!' };
    } finally {
        lock.releaseLock();
    }
}

function adminClearSeat(adminPass, seatId) {
    if (adminPass !== 'yellowisusmah') {
        return { success: false, message: 'รหัสผ่าน Admin ไม่ถูกต้อง!' };
    }
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim() === seatId) {
            sheet.deleteRow(i + 1);
            return { success: true, message: 'Admin ปลดล็อกที่นั่ง ' + seatId + ' สำเร็จ' };
        }
    }
    return { success: false, message: 'ไม่พบที่นั่งที่ต้องการลบ' };
}

function adminClearAll(adminPass) {
    if (adminPass !== 'yellowisusmah') {
        return { success: false, message: 'รหัสผ่าน Admin ไม่ถูกต้อง!' };
    }
    const sheet = getSheet();
    sheet.clear();
    sheet.appendRow(['Seat ID', 'ชื่อ-นามสกุล', 'ชั้นเรียน', 'เลขที่', 'รหัส7หลัก', 'รหัสยกเลิก8หลัก', 'เวลาที่จอง']);
    return { success: true, message: 'Admin รีเซ็ตข้อมูลการจองทั้งหมดเรียบร้อยแล้ว' };
}
