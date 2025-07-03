// config/index.js - Application Configuration
require('dotenv').config();

const CONFIG = {
  GOOGLE_SHEETS: {
    SPREADSHEET_ID: process.env.GOOGLE_SPREADSHEET_ID,
    PRIVATE_KEY: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
    CLIENT_EMAIL: process.env.GOOGLE_CLIENT_EMAIL,
  },
  TELEGRAM: {
    BOT_TOKEN: process.env.TELEGRAM_BOT_TOKEN,
    CHAT_ID: process.env.TELEGRAM_CHAT_ID
  },
  LINE: {
    LIFF_ID: process.env.LIFF_ID
  },
  SHEETS: {
    MAIN: 'MAIN',
    EMPLOYEES: 'EMPLOYEES',
    ON_WORK: 'ON WORK'
  },
  RENDER: {
    SERVICE_URL: process.env.RENDER_SERVICE_URL || `https://${process.env.RENDER_EXTERNAL_HOSTNAME}` || 'http://localhost:3001',
    KEEP_ALIVE_ENABLED: process.env.KEEP_ALIVE_ENABLED === 'true',
    GSA_WEBHOOK_SECRET: process.env.GSA_WEBHOOK_SECRET || 'your-secret-key'
  },
  ADMIN: {
    JWT_SECRET: process.env.JWT_SECRET || 'huana-nbp-jwt-secret-2025',
    JWT_EXPIRES_IN: '24h',
    // Admin users (in production, store in database)
    USERS: [
      {
        id: 1,
        username: 'admin',
        password: '$2a$10$7ROfP4YLlJpub4cWuPkqwu2C1shrT.QbHr2zbLeDoGLE7VxSBhmCS', // khayai042315962
        name: 'ผู้ดูแลระบบ อบต.ข่าใหญ่',
        role: 'admin'
      },
      {
        id: 2,
        username: 'huana_admin',
        password: '$2a$10$AnotherHashedPasswordHere', // ต้อง hash ก่อนใช้งานจริง
        name: 'ผู้ดูแลระบบ อบต.ข่าใหญ่',
        role: 'admin'
      }
    ]
  },
  // 🆕 เพิ่มการตั้งค่าสำหรับยกเว้นการลงเวลาออกอัตโนมัติ
  AUTO_CHECKOUT: {
    // รายชื่อพนักงานที่ยกเว้นจากการลงเวลาออกอัตโนมัติ (เช่น ยามกลางคืน)
    EXEMPT_EMPLOYEES: [
      '1017-เปรมชัย ทองสงคราม' // พนักงานยามกลางคืน
    ],
    // เวลาที่ทำการลงเวลาออกอัตโนมัติ (23:59)
    CUTOFF_HOUR: 23,
    CUTOFF_MINUTE: 59
  },
  TIMEZONE: 'Asia/Bangkok'
};

// Validation function
function validateConfig() {
  const required = [
    { name: 'GOOGLE_SPREADSHEET_ID', value: CONFIG.GOOGLE_SHEETS.SPREADSHEET_ID },
    { name: 'GOOGLE_PRIVATE_KEY', value: CONFIG.GOOGLE_SHEETS.PRIVATE_KEY },
    { name: 'GOOGLE_CLIENT_EMAIL', value: CONFIG.GOOGLE_SHEETS.CLIENT_EMAIL }
  ];

  const missing = required.filter(item => !item.value);
  
  if (missing.length > 0) {
    console.error('❌ Missing required environment variables:');
    missing.forEach(item => console.error(`   - ${item.name}`));
    return false;
  }
  
  console.log('✅ Configuration validated successfully');
  return true;
}

module.exports = {
  CONFIG,
  validateConfig
};
