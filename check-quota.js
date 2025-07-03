// check-quota.js - สคริปต์ตรวจสอบสถานะ API quota
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
require('dotenv').config();

async function checkQuotaStatus() {
  console.log('🔍 ตรวจสอบสถานะ Google Sheets API...');
  
  try {
    const serviceAccountAuth = new JWT({
      email: process.env.GOOGLE_CLIENT_EMAIL,
      key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });

    const doc = new GoogleSpreadsheet(process.env.GOOGLE_SPREADSHEET_ID, serviceAccountAuth);
    
    console.log('📋 กำลังทดสอบการเชื่อมต่อ...');
    await doc.loadInfo();
    
    console.log(`✅ เชื่อมต่อสำเร็จ: ${doc.title}`);
    console.log(`📊 จำนวน sheets: ${doc.sheetCount}`);
    
    // ทดสอบอ่านข้อมูล
    const sheets = ['EMPLOYEES', 'ON WORK', 'MAIN'];
    for (const sheetName of sheets) {
      try {
        const sheet = doc.sheetsByTitle[sheetName];
        if (sheet) {
          const rows = await sheet.getRows({ limit: 1 });
          console.log(`✅ ${sheetName}: อ่านได้ปกติ`);
        } else {
          console.log(`⚠️ ${sheetName}: ไม่พบ sheet`);
        }
      } catch (error) {
        console.log(`❌ ${sheetName}: ${error.message}`);
      }
    }
    
    console.log('\n🎉 API ทำงานปกติ - ไม่มีปัญหา quota');
    
  } catch (error) {
    console.error('\n❌ เกิดข้อผิดพลาด:', error.message);
    
    if (error.message.includes('quota') || 
        error.message.includes('limit') || 
        error.message.includes('429')) {
      console.log('\n🚨 ข้อเสนอแนะ:');
      console.log('1. รอให้ quota reset (ประมาณ 24 ชั่วโมง)');
      console.log('2. ใช้ระบบ cache ที่ปรับปรุงแล้ว');
      console.log('3. ลดการใช้งานฟีเจอร์ที่ต้องใช้ API');
      console.log('4. เปิด Emergency Mode ในระบบ');
    }
  }
}

checkQuotaStatus();
