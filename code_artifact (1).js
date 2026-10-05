function doGet(e) {
  // ดึงไฟล์ index.html มาแสดงผลเป็น Web App
  return HtmlService.createHtmlOutputFromFile('index')
    // ตั้งชื่อ Tab ของเว็บ (อัปเดตชื่อระบบ)
    .setTitle('ระบบสรุปผลการดำเนินโครงการ/กิจกรรม แบบย่อ - โรงเรียนแม่จันวิทยาคม')
    // รองรับการแสดงผลบนมือถือ (Responsive)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    // อนุญาตให้นำไปฝัง (Embed) ลงใน Google Sites หรือเว็บอื่นๆ ได้
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันนี้จะถูกเรียกจากหน้าเว็บเพื่อไปอ่านข้อมูลมาตรฐานและเป้าหมายกลาง
function getAdminDefaults() {
  try {
    // ID Sheet สำหรับตั้งค่ามาตรฐานและเป้าหมายกลาง
    const sheetId = '1r6dg-Vu1Z-LM-eczWUuRsllzuQ4_2WsMQ2-iW5ecwII';
    const sheet = SpreadsheetApp.openById(sheetId).getSheets()[0]; // อ่านจากชีตแรก
    const data = sheet.getDataRange().getDisplayValues(); // ดึงข้อมูลทุกบรรทัด

    let newDefaults = {};
    
    // วนลูปอ่านข้อมูล (เริ่มที่บรรทัด 1 เพื่อข้ามหัวข้อในบรรทัด 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = row[0] ? row[0].toString().trim() : '';
      
      if (id && id !== 'รหัสมาตรฐาน' && id !== 'StandardID') {
        newDefaults[id] = {
          quantity: row[1] || '',
          quality: row[2] || '',
          description: row[3] || '' // ดึงคำอธิบายเกณฑ์ประเมิน
        };
      }
    }
    return { success: true, data: newDefaults };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ฟังก์ชันใหม่สำหรับบันทึกและจัดการข้อมูลใน Google Sheets ฐานข้อมูล
function saveDataToSheet(records, goals, results) {
  try {
    // ID ของ Sheet ที่จะใช้เก็บข้อมูลการบันทึก (Database)
    const sheetId = '12j68-TD3VtGkHlG3BFWjP4XGLfQ4CGwfPV0zhbbecAo';
    const ss = SpreadsheetApp.openById(sheetId);
    
    // ดึงหรือสร้าง Sheet ย่อยสำหรับแยกเก็บแต่ละหมวดหมู่
    let sheetRecords = ss.getSheetByName('Activities') || ss.insertSheet('Activities');
    let sheetResults = ss.getSheetByName('Results') || ss.insertSheet('Results');
    
    // อัปเดตข้อมูลกิจกรรม (แนบ Header) - เพิ่ม project_code และ activity_code
    const activityHeaders = ['id', 'type', 'project_code', 'project_name', 'activity_code', 'activity_name', 'objectives', 'strategies', 'standards', 'created_at'];
    upsertData(sheetRecords, records, activityHeaders, 0);
    
    // อัปเดตข้อมูลผลลัพธ์
    const resultHeaders = ['result_id', 'result_goal_id', 'result_quantity_achieved', 'result_quality_achieved', 'created_at'];
    upsertData(sheetResults, results, resultHeaders, 0);
    
    // หมายเหตุ: หน้าเป้าหมาย (Goals) ถูกตัดออก จึงไม่มีการเขียนลง Sheet Goals แยกแล้ว 
    // เพราะเป้าหมายจะถูกดึงสดจาก getAdminDefaults เสมอ
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

// ฟังก์ชันช่วยเหลือสำหรับ Update (หากมี ID เดิม) หรือ Append (หากเป็นข้อมูลใหม่)
function upsertData(sheet, data, headers, idColIdx) {
  if (!data || data.length === 0) return;
  
  // สร้างหัวตารางถ้ายังไม่มี
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers); 
  }
  
  const existingData = sheet.getDataRange().getValues();
  
  // สร้างแผนที่แถวของ ID ปัจจุบัน
  let idMap = {};
  for (let i = 1; i < existingData.length; i++) {
    idMap[existingData[i][idColIdx]] = i + 1; // อ้างอิง index แถว (แถวใน Sheet เริ่มที่ 1)
  }
  
  data.forEach(obj => {
    let row = headers.map(h => {
      let val = obj[h];
      return (val !== undefined && val !== null) ? (typeof val === 'object' ? JSON.stringify(val) : val) : '';
    });
    
    let id = row[idColIdx];
    if (idMap[id]) {
      // มีข้อมูลเก่าอยู่แล้ว -> อัปเดตข้อมูล (Update)
      sheet.getRange(idMap[id], 1, 1, headers.length).setValues([row]);
    } else {
      // ไม่มีข้อมูลในระบบ -> เพิ่มแถวใหม่ (Append)
      sheet.appendRow(row);
      idMap[id] = sheet.getLastRow(); 
    }
  });
}

// ฟังก์ชันลบกิจกรรมจาก Sheet
function deleteDataFromSheet(id) {
  try {
    const sheetId = '12j68-TD3VtGkHlG3BFWjP4XGLfQ4CGwfPV0zhbbecAo';
    const ss = SpreadsheetApp.openById(sheetId);
    const sheetRecords = ss.getSheetByName('Activities');
    
    if (sheetRecords) {
       const data = sheetRecords.getDataRange().getValues();
       // ค้นหาแถวที่มี ID ตรงกันและทำการลบ
       for (let i = 1; i < data.length; i++) {
          if (data[i][0] === id) {
             sheetRecords.deleteRow(i + 1);
             break;
          }
       }
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}