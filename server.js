// server.js - Time Tracker with Admin Panel and Excel Export
const express = require('express');
const cors = require('cors');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const path = require('path');
const cron = require('node-cron');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const ExcelJS = require('exceljs');
const moment = require('moment-timezone');
const fetch = require('node-fetch');
const { CONFIG, validateConfig } = require('./config');
const ExcelExportService = require('./services/excelExport');

const app = express();
const PORT = process.env.PORT || 3001;
console.log(`🔧 Using PORT: ${PORT}`);

// ========== Helper Functions ==========

/**
 * คำนวณชั่วโมงการทำงานแบบเดียวกันกับ admin stats
 * @param {string} clockInTime - เวลาเข้างาน
 * @param {string} [clockOutTime] - เวลาออกงาน (ถ้าไม่ให้จะใช้เวลาปัจจุบัน)
 * @returns {number} - ชั่วโมงการทำงาน (ทศนิยม)
 */
function calculateWorkingHours(clockInTime, clockOutTime = null) {
  if (!clockInTime) {
    console.warn('⚠️ No clock in time provided for calculation');
    return 0;
  }

  try {
    // ใช้ logic เดียวกันกับ admin stats
    const clockInMoment = moment.tz(clockInTime, 'YYYY-MM-DD H:mm:ss', CONFIG.TIMEZONE);
    const endTimeMoment = clockOutTime ? 
      moment.tz(clockOutTime, 'YYYY-MM-DD H:mm:ss', CONFIG.TIMEZONE) :
      moment().tz(CONFIG.TIMEZONE);

    if (!clockInMoment.isValid()) {
      console.error(`❌ Invalid clockInTime format: "${clockInTime}"`);
      return 0;
    }

    if (clockOutTime && !endTimeMoment.isValid()) {
      console.error(`❌ Invalid clockOutTime format: "${clockOutTime}"`);
      return 0;
    }

    // คำนวณความแตกต่างของเวลาในหน่วยชั่วโมง (เหมือน admin stats)
    const hours = endTimeMoment.diff(clockInMoment, 'hours', true);

    // Debug: แสดงการคำนวณ
    console.log(`⏰ Working hours calculation:`, {
      clockIn: clockInMoment.format('YYYY-MM-DD HH:mm:ss'),
      endTime: endTimeMoment.format('YYYY-MM-DD HH:mm:ss'),
      diffHours: hours.toFixed(2)
    });

    // ตรวจสอบให้แน่ใจว่าไม่เป็นลบ (ป้องกันปัญหา timezone)
    if (hours >= 0) {
      return hours;
    } else {
      console.warn(`⚠️ Negative working hours detected: ${hours.toFixed(2)}, setting to 0`);
      return 0;
    }
  } catch (error) {
    console.error('❌ Error calculating working hours:', error);
    return 0;
  }
}
// สร้าง hash password (ใช้ในการตั้งรหัสผ่านครั้งแรก)
async function createPassword(plainPassword) {
  return await bcrypt.hash(plainPassword, 10);
}

// ========== Middleware ==========
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Security middleware สำหรับ webhook
app.use('/api/webhook', (req, res, next) => {
  const providedSecret = req.headers['x-webhook-secret'] || req.query.secret;
  if (providedSecret !== CONFIG.RENDER.GSA_WEBHOOK_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
});

// Admin Authentication Middleware
function authenticateAdmin(req, res, next) {
  const authHeader = req.headers.authorization;
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ 
      success: false, 
      error: 'Access token required',
      errorCode: 'NO_TOKEN'
    });
  }

  try {
    const decoded = jwt.verify(token, CONFIG.ADMIN.JWT_SECRET);
    req.user = decoded;
    next();
  } catch (error) {
    console.error('JWT verification error:', error.name, error.message);
    
    // จัดการ error แต่ละประเภท
    let errorResponse = {
      success: false,
      error: 'Authentication failed'
    };

    if (error.name === 'TokenExpiredError') {
      errorResponse.error = 'Token has expired. Please login again.';
      errorResponse.errorCode = 'TOKEN_EXPIRED';
      errorResponse.expiredAt = error.expiredAt;
    } else if (error.name === 'JsonWebTokenError') {
      errorResponse.error = 'Invalid token format';
      errorResponse.errorCode = 'INVALID_TOKEN';
    } else if (error.name === 'NotBeforeError') {
      errorResponse.error = 'Token not active yet';
      errorResponse.errorCode = 'TOKEN_NOT_ACTIVE';
    } else {
      errorResponse.errorCode = 'TOKEN_ERROR';
    }

    return res.status(401).json(errorResponse);
  }
}

// Serve static files
app.use(express.static('public'));

// Admin routes - ใช้ไฟล์จาก public folder
app.get('/admin/login', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/admin/dashboard', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

app.get('/admin', (req, res) => {
  res.redirect('/admin/login');
});

// Serve ads.txt specifically
app.get('/ads.txt', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'ads.txt'));
});

// Serve robots.txt (optional)
app.get('/robots.txt', (req, res) => {
  res.type('text/plain');
  res.send('User-agent: *\nDisallow: /api/\nAllow: /ads.txt');
});

// ========== Keep-Alive Service ==========
class KeepAliveService {
  constructor() {
    this.isEnabled = CONFIG.RENDER.KEEP_ALIVE_ENABLED;
    this.serviceUrl = CONFIG.RENDER.SERVICE_URL;
    this.startTime = new Date();
    this.pingCount = 0;
    this.errorCount = 0;
  }

  init() {
    if (!this.isEnabled) {
      console.log('🔴 Keep-Alive disabled');
      return;
    }

    console.log('🟢 Keep-Alive service started');
    console.log(`📍 Service URL: ${this.serviceUrl}`);

    // เวลาทำงาน: 05:00-10:00 และ 15:00-20:00 (เวลาไทย)
    // ปิงทุก 10 นาที
    cron.schedule('*/10 * * * *', () => {
      this.checkAndPing();
    }, {
      scheduled: true,
      timezone: CONFIG.TIMEZONE
    });

    // 🆕 ตรวจสอบและจัดการคนที่ลืมลงเวลาออก - ทุกวันเวลา 23:59
    cron.schedule('59 23 * * *', async () => {
      console.log('🔍 Starting daily missed checkout check at 23:59...');
      try {
        const result = await sheetsService.checkAndHandleMissedCheckouts();
        console.log('✅ Daily missed checkout check completed:', result);
      } catch (error) {
        console.error('❌ Error in daily missed checkout check:', error);
      }
    }, {
      scheduled: true,
      timezone: CONFIG.TIMEZONE
    });

    // Ping ทันทีเมื่อเริ่มต้น
    setTimeout(() => this.ping(), 5000);
  }

  checkAndPing() {
    const now = new Date();
    const hour = now.getHours();
    
    // เช็คว่าอยู่ในเวลาทำงานไหม
    const isWorkingHour = (hour >= 5 && hour < 10) || (hour >= 15 && hour < 20);
    
    if (isWorkingHour) {
      this.ping();
    } else {
      console.log(`😴 Outside working hours (${hour}:00), skipping ping`);
    }
  }

  async ping() {
    try {
      const response = await fetch(`${this.serviceUrl}/api/ping`, {
        method: 'GET',
        headers: {
          'User-Agent': 'KeepAlive-Service/1.0'
        }
      });

      this.pingCount++;
      
      if (response.ok) {
        console.log(`✅ Keep-Alive ping #${this.pingCount} successful`);
        this.errorCount = 0; // Reset error count on success
      } else {
        throw new Error(`HTTP ${response.status}`);
      }
      
    } catch (error) {
      this.errorCount++;
      console.log(`❌ Keep-Alive ping #${this.pingCount} failed:`, error.message);
      
      // หากผิดพลาดติดต่อกัน 5 ครั้ง ให้ลอง ping ใหม่หลัง 1 นาที
      if (this.errorCount >= 5) {
        console.log('🔄 Too many errors, will retry in 1 minute');
        setTimeout(() => this.ping(), 60000);
      }
    }
  }

  getStats() {
    const uptime = Math.floor((Date.now() - this.startTime.getTime()) / 1000);
    return {
      enabled: this.isEnabled,
      uptime: uptime,
      pingCount: this.pingCount,
      errorCount: this.errorCount,
      lastPing: new Date().toISOString()
    };
  }
}

// ========== Google Sheets Service ==========
class GoogleSheetsService {
  constructor() {
    this.doc = null;
    this.isInitialized = false;
    // เพิ่มระบบ caching เพื่อลดการเรียก API
    this.cache = {
      employees: { data: null, timestamp: null, ttl: 300000 }, // 5 นาที
      onWork: { data: null, timestamp: null, ttl: 60000 }, // 1 นาที  
      main: { data: null, timestamp: null, ttl: 30000 }, // 30 วินาที
      stats: { data: null, timestamp: null, ttl: 120000 } // 2 นาที
    };
    this.emergencyMode = false; // เริ่มต้นปิดระบบ emergency mode
  }

  async initialize() {
    if (this.isInitialized) return;

    try {
      const serviceAccountAuth = new JWT({
        email: CONFIG.GOOGLE_SHEETS.CLIENT_EMAIL,
        key: CONFIG.GOOGLE_SHEETS.PRIVATE_KEY,
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });

      this.doc = new GoogleSpreadsheet(CONFIG.GOOGLE_SHEETS.SPREADSHEET_ID, serviceAccountAuth);
      await this.doc.loadInfo();
      
      console.log(`✅ Connected to Google Sheets: ${this.doc.title}`);
      this.isInitialized = true;
      
    } catch (error) {
      console.error('❌ Failed to initialize Google Sheets:', error);
      throw error;
    }
  }  // เพิ่มฟังก์ชัน cache helper
  isCacheValid(cacheKey) {
    const cache = this.cache[cacheKey];
    if (!cache || !cache.data || !cache.timestamp) return false;
    return (Date.now() - cache.timestamp) < cache.ttl;
  }

  setCache(cacheKey, data) {
    if (!this.cache[cacheKey]) {
      this.cache[cacheKey] = { data: null, timestamp: null, ttl: 300000 }; // default 5 min
    }
    this.cache[cacheKey] = {
      data: data,
      timestamp: Date.now(),
      ttl: this.cache[cacheKey].ttl
    };
  }

  getCache(cacheKey) {
    const cache = this.cache[cacheKey];
    return cache && cache.data ? cache.data : null;
  }

  clearCache(cacheKey = null) {
    if (cacheKey) {
      this.cache[cacheKey].data = null;
      this.cache[cacheKey].timestamp = null;
    } else {
      // Clear all cache
      Object.keys(this.cache).forEach(key => {
        this.cache[key].data = null;
        this.cache[key].timestamp = null;
      });
    }
  }

  async getSheet(sheetName) {
    if (!this.isInitialized) {
      await this.initialize();
    }
    
    const sheet = this.doc.sheetsByTitle[sheetName];
    if (!sheet) {
      throw new Error(`Sheet ${sheetName} not found`);
    }
    
    return sheet;
  }  // เพิ่มฟังก์ชันดึงข้อมูลพร้อม cache และ rate limiting
  async getCachedSheetData(sheetName) {
    const cacheKey = sheetName.toLowerCase().replace(/\s+/g, '');
    
    // ตรวจสอบ cache ก่อน
    if (this.isCacheValid(cacheKey)) {
      console.log(`📋 Using cached data for ${sheetName}`);
      return this.getCache(cacheKey);
    }

    // ตรวจสอบ rate limit ก่อนเรียก API
    if (!apiMonitor.canMakeAPICall()) {
      console.warn(`⚠️ API rate limit reached, using stale cache for ${sheetName}`);
      const staleData = this.getCache(cacheKey);
      if (staleData) {
        return staleData;
      }
      throw new Error('Rate limit exceeded and no cached data available');
    }

    console.log(`🔄 Fetching fresh data from ${sheetName}`);
    apiMonitor.logAPICall(`getCachedSheetData:${sheetName}`);
    
    try {
      const sheet = await this.getSheet(sheetName);
      
      let rows;
      if (sheetName === CONFIG.SHEETS.ON_WORK) {
        rows = await sheet.getRows({ offset: 1 }); // เริ่มจากแถว 3
      } else {
        rows = await sheet.getRows();
      }

      // บันทึกลง cache
      this.setCache(cacheKey, rows);
      
      // เสร็จสิ้น API call
      apiMonitor.finishCall();
      return rows;
      
    } catch (error) {
      // เสร็จสิ้น API call แม้จะ error
      apiMonitor.finishCall();
      
      console.error(`❌ API Error for ${sheetName}:`, error.message);
      
      // ถ้าเป็น quota error, ใช้ stale cache
      if (error.message.includes('quota') || error.message.includes('limit') || 
          error.message.includes('429') || error.message.includes('RATE_LIMIT')) {
        console.warn(`⚠️ Quota exceeded for ${sheetName}, using stale cache`);
        const staleData = this.getCache(cacheKey);
        if (staleData) {
          console.log(`📋 Using stale cache for ${sheetName} (${staleData.length} items)`);
          return staleData;
        }
      }
      
      throw error;
    }
  }

  // ฟังก์ชันช่วยเหลือสำหรับการเปรียบเทียบชื่อ
  normalizeEmployeeName(name) {
    if (!name) return '';
    
    return name.toString()
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  isNameMatch(inputName, compareName) {
    if (!inputName || !compareName) return false;
    
    const normalizedInput = this.normalizeEmployeeName(inputName);
    const normalizedCompare = this.normalizeEmployeeName(compareName);
    
    return normalizedInput === normalizedCompare ||
           normalizedInput.includes(normalizedCompare) ||
           normalizedCompare.includes(normalizedInput);
  }
    /**
   * ตรวจสอบว่าพนักงานคนนี้ได้รับการยกเว้นจากการลงเวลาออกอัตโนมัติหรือไม่
   * @param {string} employeeName - ชื่อพนักงาน
   * @returns {boolean} - true ถ้าได้รับการยกเว้น, false ถ้าไม่ได้
   */
  isEmployeeExempt(employeeName) {
    if (!employeeName || !CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES) {
      return false;
    }
    
    const normalizedInputName = this.normalizeEmployeeName(employeeName);
    
    // ตรวจสอบกับรายชื่อที่ยกเว้น
    return CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES.some(exemptName => {
      const normalizedExemptName = this.normalizeEmployeeName(exemptName);
      
      // ตรวจสอบแบบหลายรูปแบบ
      const isExactMatch = normalizedInputName === normalizedExemptName;
      const isPartialMatch = normalizedInputName.includes(normalizedExemptName) || 
                            normalizedExemptName.includes(normalizedInputName);
      
      if (isExactMatch || isPartialMatch) {
        console.log(`🛡️ Employee exempt match found: "${employeeName}" ↔ "${exemptName}"`);
        return true;
      }
      
      return false;
    });
  }

  async getEmployees() {
    try {
      // ใช้ cached data แทนการเรียก API ใหม่
      const rows = await this.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
      
      const employees = rows.map(row => row.get('ชื่อ-นามสกุล')).filter(name => name);
      return employees;
      
    } catch (error) {
      console.error('Error getting employees:', error);
      return [];
    }
  }  async getEmployeeStatus(employeeName) {
    try {
      // ใช้ safe method แทน
      const rows = await this.safeGetCachedSheetData(CONFIG.SHEETS.ON_WORK);
      
      console.log(`🔍 Checking status for: "${employeeName}"`);
      console.log(`📊 Total rows in ON_WORK (from row 3): ${rows.length}`);
      
      if (rows.length === 0) {
        console.log('📋 ON_WORK sheet is empty (from row 3)');
        return { isOnWork: false, workRecord: null };
      }
      
      const workRecord = rows.find(row => {
        const systemName = row.get('ชื่อในระบบ');
        const employeeName2 = row.get('ชื่อพนักงาน');
        
        const isMatch = this.isNameMatch(employeeName, systemName) || 
                       this.isNameMatch(employeeName, employeeName2);
        
        if (isMatch) {
          console.log(`✅ Found match: "${employeeName}" ↔ "${systemName || employeeName2}"`);
        }
        
        return isMatch;
      });
      
      if (workRecord) {
        let mainRowIndex = null;
        
        const rowRef1 = workRecord.get('แถวอ้างอิง');
        const rowRef2 = workRecord.get('แถวในMain');
        
        if (rowRef1 && !isNaN(parseInt(rowRef1))) {
          mainRowIndex = parseInt(rowRef1);
        } else if (rowRef2 && !isNaN(parseInt(rowRef2))) {
          mainRowIndex = parseInt(rowRef2);
        }
        
        console.log(`✅ Employee "${employeeName}" is currently working`);
        
        return {
          isOnWork: true,
          workRecord: {
            row: workRecord,
            mainRowIndex: mainRowIndex,
            clockIn: workRecord.get('เวลาเข้า'),
            systemName: workRecord.get('ชื่อในระบบ'),
            employeeName: workRecord.get('ชื่อพนักงาน')
          }
        };
      } else {
        console.log(`❌ Employee "${employeeName}" is not currently working`);
        return { isOnWork: false, workRecord: null };
      }
      
    } catch (error) {
      console.error('❌ Error checking employee status:', error);
      return { isOnWork: false, workRecord: null };
    }
  }
  // Admin functions
  async getAdminStats() {
    try {
      // ตรวจสอบ cache สำหรับ stats ก่อน
      if (this.isCacheValid('stats')) {
        console.log('📊 Using cached admin stats');
        return this.getCache('stats');
      }

      console.log('🔄 Fetching fresh admin stats data');      // ใช้ safe method แทน การเรียก API
      const [employees, onWorkRows, mainRows] = await Promise.all([
        this.safeGetCachedSheetData(CONFIG.SHEETS.EMPLOYEES),
        this.safeGetCachedSheetData(CONFIG.SHEETS.ON_WORK),
        this.safeGetCachedSheetData(CONFIG.SHEETS.MAIN)
      ]);

      const totalEmployees = employees.length;
      const workingNow = onWorkRows.length;// หาจำนวนคนที่มาทำงานวันนี้ (ใช้ข้อมูลจาก ON_WORK sheet ที่มีวันที่วันนี้)
      const today = moment().tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
      console.log(`📅 Today date for comparison: ${today}`);
      console.log(`📊 Total MAIN sheet records: ${mainRows.length}`);
      console.log(`� Total ON_WORK sheet records: ${onWorkRows.length}`);
      
      // นับจาก ON_WORK sheet ที่มีวันที่วันนี้
      const presentToday = onWorkRows.filter(row => {
        const clockInDate = row.get('เวลาเข้า');
        if (!clockInDate) return false;
        
        try {
          const employeeName = row.get('ชื่อพนักงาน') || row.get('ชื่อในระบบ');
          let dateStr = '';
          
          // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
          if (typeof clockInDate === 'string' && clockInDate.includes(' ')) {
            dateStr = clockInDate.split(' ')[0];
            const isToday = dateStr === today;
            
            if (isToday) {
              console.log(`✅ Present today (ON_WORK): ${employeeName} - ${clockInDate} (date: ${dateStr})`);
            }
            
            return isToday;
          }
          
          // ถ้าเป็น ISO format
          if (typeof clockInDate === 'string' && clockInDate.includes('T')) {
            dateStr = clockInDate.split('T')[0];
            const isToday = dateStr === today;
            
            if (isToday) {
              console.log(`✅ Present today (ON_WORK ISO): ${employeeName} - ${clockInDate} (date: ${dateStr})`);
            }
            
            return isToday;
          }
          
          return false;
        } catch (error) {
          console.warn(`⚠️ Error parsing date in ON_WORK: ${clockInDate}`, error);
          return false;
        }
      }).length;
      
      console.log(`📊 Present today count: ${presentToday} out of ${onWorkRows.length} ON_WORK records`);

      const absentToday = totalEmployees - presentToday;      // รายชื่อพนักงานที่กำลังทำงาน
      const workingEmployees = onWorkRows.map(row => {
        const clockInTime = row.get('เวลาเข้า');
        let workingHours = '0 ชม.';
        
        if (clockInTime) {
          // 🎯 ใช้ฟังก์ชันคำนวณเวลาแบบเดียวกันกับ clock out
          const hours = calculateWorkingHours(clockInTime);
          
          if (hours > 0) {
            workingHours = `${hours.toFixed(1)} ชม.`;
          } else {
            workingHours = '0 ชม.';
          }
        }

        return {
          name: row.get('ชื่อพนักงาน') || row.get('ชื่อในระบบ'),
          clockIn: clockInTime ? moment.tz(clockInTime, 'YYYY-MM-DD H:mm:ss', CONFIG.TIMEZONE).format('HH:mm') : '',
          workingHours
        };
      });      const stats = {
        totalEmployees,
        presentToday,
        workingNow,
        absentToday,
        workingEmployees
      };
      
      console.log('📊 Admin stats summary:', {
        totalEmployees,
        presentToday,
        workingNow,
        absentToday,
        workingEmployeesCount: workingEmployees.length
      });
      
      // บันทึกลง cache
      this.setCache('stats', stats);
      
      return stats;

    } catch (error) {
      console.error('Error getting admin stats:', error);
      throw error;
    }
  }
  async getReportData(type, params) {
    try {
      console.log(`📊 Getting report data for type: ${type}`, params);
      
      // ใช้ safe cached data method
      const rows = await this.safeGetCachedSheetData(CONFIG.SHEETS.MAIN);
      
      if (!rows || rows.length === 0) {
        console.log('⚠️ No data found in MAIN sheet');
        return [];
      }

      console.log(`📋 Found ${rows.length} total records in MAIN sheet`);
      
      // Debug: แสดงตัวอย่างข้อมูลไม่กี่แถวแรก
      if (rows.length > 0) {
        console.log('📋 Sample data (first 3 rows):');
        for (let i = 0; i < Math.min(3, rows.length); i++) {
          const row = rows[i];
          // ใช้ index แทนเนื่องจาก sheet ไม่มี header
          const employee = row._rawData[0]; // column 0: ชื่อพนักงาน
          const clockIn = row._rawData[3];  // column 3: เวลาเข้า
          console.log(`   Row ${i+1}: Employee="${employee}", ClockIn="${clockIn}" (type: ${typeof clockIn})`);
        }
        
        // Debug: แสดง headers ของ sheet
        console.log('📋 Sheet headers:', Object.keys(rows[0]._rawData));
        
        // Debug: แสดงค่าของแต่ละ column ในแถวแรก
        const firstRow = rows[0];
        console.log('📋 First row values by index:');
        console.log(`   Column 0: "${firstRow._rawData[0]}" (should be ชื่อพนักงาน)`);
        console.log(`   Column 1: "${firstRow._rawData[1]}" (should be Line name)`);
        console.log(`   Column 2: "${firstRow._rawData[2]}" (should be รูปภาพ)`);
        console.log(`   Column 3: "${firstRow._rawData[3]}" (should be เวลาเข้า)`);
        console.log(`   Column 4: "${firstRow._rawData[4]}" (should be userinfo/หมายเหตุ)`);
        console.log(`   Column 5: "${firstRow._rawData[5]}" (should be เวลาออก)`);
        console.log(`   Column 6: "${firstRow._rawData[6]}" (should be พิกัดเข้า)`);
        console.log(`   Column 7: "${firstRow._rawData[7]}" (should be สถานที่เข้า)`);
        console.log(`   Column 8: "${firstRow._rawData[8]}" (should be พิกัดออก)`);
        console.log(`   Column 9: "${firstRow._rawData[9]}" (should be ที่อยู่ออก)`);
        console.log(`   Column 10: "${firstRow._rawData[10]}" (should be ชั่วโมงทำงาน)`);
        console.log(`   Column 11: "${firstRow._rawData[11]}" (should be หมายเหตุเดิม - ไม่ใช้แล้ว)`);
      }
      
      let filteredRows = [];

      switch (type) {
        case 'daily':
          const targetDate = moment(params.date).tz(CONFIG.TIMEZONE).format('YYYY-MM-DD');
          console.log(`📅 Filtering for daily report: ${targetDate}`);
          
          filteredRows = rows.filter(row => {
            const clockIn = row._rawData[3]; // column 3: เวลาเข้า
            if (!clockIn) return false;
            
            try {
              let dateStr = '';
              console.log(`🔍 Checking clockIn: "${clockIn}" (type: ${typeof clockIn})`);
              
              // ถ้าเป็น string format 'DD/MM/YYYY HH:mm:ss'
              if (typeof clockIn === 'string' && clockIn.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                const datePart = clockIn.split(' ')[0]; // "26/06/2025"
                const [day, month, year] = datePart.split('/');
                dateStr = `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
              }
              // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
              else if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                dateStr = clockIn.split(' ')[0];
              } else if (typeof clockIn === 'string' && clockIn.includes('T')) {
                // ISO format
                dateStr = clockIn.split('T')[0];
              } else if (typeof clockIn === 'string' && clockIn.match(/^\d{4}-\d{2}-\d{2}$/)) {
                // Already in YYYY-MM-DD format
                dateStr = clockIn;
              } else {
                // Date object หรือ format อื่น
                const rowDate = moment(clockIn).tz(CONFIG.TIMEZONE);
                if (rowDate.isValid()) {
                  dateStr = rowDate.format('YYYY-MM-DD');
                } else {
                  console.warn(`⚠️ Invalid date format: "${clockIn}"`);
                  return false;
                }
              }
              
              console.log(`📅 Extracted date: "${dateStr}" vs target: "${targetDate}"`);
              const isMatch = dateStr === targetDate;
              if (isMatch) {
                console.log(`✅ Date match found: ${row._rawData[0]} - ${clockIn}`);
              } else if (clockIn && clockIn.includes('26')) {
                console.log(`❓ Potential match (contains '26'): ${row._rawData[0]} - ${clockIn} -> ${dateStr}`);
              }
              
              return isMatch;
            } catch (error) {
              console.warn('❌ Error parsing date for daily report:', clockIn, error);
              return false;
            }
          });
          
          console.log(`📊 Daily filter result: ${filteredRows.length} records found for ${targetDate}`);
          break;

        case 'monthly':
          const month = parseInt(params.month);
          const year = parseInt(params.year);
          console.log(`📅 Filtering for monthly report: ${month}/${year}`);
          
          filteredRows = rows.filter(row => {
            const clockIn = row._rawData[3]; // column 3: เวลาเข้า
            if (!clockIn) return false;
            
            try {
              let dateStr = '';
              
              // ใช้วิธีเดียวกับรายงานรายวัน เพื่อความสอดคล้อง
              if (typeof clockIn === 'string' && clockIn.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                const datePart = clockIn.split(' ')[0]; // "26/06/2025"
                const [day, monthPart, yearPart] = datePart.split('/');
                dateStr = `${yearPart}-${monthPart.padStart(2, '0')}-${day.padStart(2, '0')}`;
              }
              // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
              else if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                dateStr = clockIn.split(' ')[0];
              } else if (typeof clockIn === 'string' && clockIn.includes('T')) {
                // ISO format
                dateStr = clockIn.split('T')[0];
              } else if (typeof clockIn === 'string' && clockIn.match(/^\d{4}-\d{2}-\d{2}$/)) {
                // Already in YYYY-MM-DD format
                dateStr = clockIn;
              } else {
                // Date object หรือ format อื่น
                const rowDate = moment(clockIn).tz(CONFIG.TIMEZONE);
                if (rowDate.isValid()) {
                  dateStr = rowDate.format('YYYY-MM-DD');
                } else {
                  console.warn(`⚠️ Invalid date format: "${clockIn}"`);
                  return false;
                }
              }
              
              // แปลงเป็น Date object เพื่อเปรียบเทียบ
              const rowDate = moment(dateStr).tz(CONFIG.TIMEZONE);
              if (!rowDate.isValid()) return false;
              
              const isMatch = rowDate.month() + 1 === month && rowDate.year() === year;
              
              if (isMatch) {
                console.log(`✅ Monthly match found: ${row._rawData[0]} - ${clockIn} -> ${dateStr}`);
              }
              
              return isMatch;
            } catch (error) {
              console.warn('❌ Error parsing date for monthly report:', clockIn, error);
              return false;
            }
          });
          break;

        case 'range':
          const startMoment = moment(params.startDate).tz(CONFIG.TIMEZONE).startOf('day');
          const endMoment = moment(params.endDate).tz(CONFIG.TIMEZONE).endOf('day');
          console.log(`📅 Filtering for range report: ${startMoment.format('YYYY-MM-DD')} to ${endMoment.format('YYYY-MM-DD')}`);
          
          filteredRows = rows.filter(row => {
            const clockIn = row._rawData[3]; // column 3: เวลาเข้า
            if (!clockIn) return false;
            
            try {
              let rowMoment;
              
              // ถ้าเป็น string format 'DD/MM/YYYY HH:mm:ss'
              if (typeof clockIn === 'string' && clockIn.match(/^\d{2}\/\d{2}\/\d{4}/)) {
                rowMoment = moment(clockIn, 'DD/MM/YYYY HH:mm:ss').tz(CONFIG.TIMEZONE);
              }
              // ถ้าเป็น string format 'YYYY-MM-DD HH:mm:ss'
              else if (typeof clockIn === 'string' && clockIn.includes(' ')) {
                rowMoment = moment(clockIn, 'YYYY-MM-DD HH:mm:ss').tz(CONFIG.TIMEZONE);
              } else if (typeof clockIn === 'string' && clockIn.includes('T')) {
                // ISO format
                rowMoment = moment(clockIn).tz(CONFIG.TIMEZONE);
              } else {
                // Date object หรือ format อื่น
                rowMoment = moment(clockIn).tz(CONFIG.TIMEZONE);
              }
              
              if (!rowMoment.isValid()) return false;
              
              return rowMoment.isBetween(startMoment, endMoment, null, '[]');
            } catch (error) {
              console.warn('Error parsing date for range report:', clockIn, error);
              return false;
            }
          });
          break;

        default:
          throw new Error(`Unsupported report type: ${type}`);
      }

      console.log(`📊 Filtered to ${filteredRows.length} records for ${type} report`);

      // แปลงข้อมูลเป็น format ที่ใช้งานง่าย
      const reportData = filteredRows.map((row, index) => {
        // ใช้ index แทนเนื่องจาก sheet ไม่มี header
        const employee = row._rawData[0] || '';        // column 0: ชื่อพนักงาน
        const lineName = row._rawData[1] || '';        // column 1: Line name
        const clockIn = row._rawData[3] || '';         // column 3: เวลาเข้า
        const clockOut = row._rawData[5] || '';        // column 5: เวลาออก
        const userInfo = row._rawData[4] || '';        // column 4: userinfo/หมายเหตุ (ใช้แทนหมายเหตุ)
        const location = row._rawData[6] || '';        // column 6: พิกัด
        const locationName = row._rawData[7] || '';    // column 7: สถานที่เข้า
        const locationOutCoords = row._rawData[8] || ''; // column 8: พิกัดออก
        const locationOut = row._rawData[9] || '';     // column 9: ที่อยู่ออก
        const workingHours = row._rawData[10] || '';   // column 10: ชั่วโมงทำงาน
        const note = row._rawData[4] || '';            // column 4: หมายเหตุ (เปลี่ยนจาก 11 เป็น 4)
        
        // Debug: แสดงข้อมูลแต่ละ row
        if (index < 3) {
          console.log(`📋 Row ${index + 1} data:`, {
            employee: employee,
            clockIn: clockIn,
            clockOut: clockOut,
            lineName: lineName,
            userInfo: userInfo,
            location: location,
            locationName: locationName,
            locationOut: locationOut,
            workingHours: workingHours,
            note: note,
            allData: row._rawData
          });
        }
        
        return {
          no: index + 1,
          employee: employee,
          lineName: lineName,
          clockIn: clockIn,
          clockOut: clockOut,
          note: note,
          workingHours: workingHours,
          locationIn: locationName,
          locationOut: locationOut,
          userInfo: userInfo
        };
      });

      console.log(`✅ Report data prepared successfully: ${reportData.length} records`);
      return reportData;

    } catch (error) {
      console.error('❌ Error getting report data:', error);
      throw error;
    }
  }

  async clockIn(data) {
    try {
      const { employee, userinfo, lat, lon, line_name, line_picture, mock_time } = data;
      
      console.log(`⏰ Clock In request for: "${employee}"`);
      if (mock_time) {
        console.log(`🧪 Using mock time: ${mock_time}`);
      }
      
      const employeeStatus = await this.getEmployeeStatus(employee);
      
      if (employeeStatus.isOnWork) {
        console.log(`❌ Employee "${employee}" is already clocked in`);
        return {
          success: false,
          message: 'คุณลงเวลาเข้างานไปแล้ว กรุณาลงเวลาออกก่อน',
          employee,
          currentStatus: 'clocked_in',
          clockInTime: employeeStatus.workRecord?.clockIn
        };
      }

      // ใช้ mock_time หากมีการส่งมา ไม่เช่นนั้นใช้เวลาปัจจุบัน
      const timestamp = mock_time 
        ? moment(mock_time).tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss')
        : moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss');
      
      // แปลงพิกัดเป็นชื่อสถานที่
      const locationName = await this.getLocationName(lat, lon);
      console.log(`📍 Location: ${locationName}`);
      
      console.log(`✅ Proceeding with clock in for "${employee}"`);
      
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      
      const newRow = await mainSheet.addRow([
        employee,           
        line_name,          
        `=IMAGE("${line_picture}")`, 
        timestamp,          
        userinfo || '',     
        '',                 
        `${lat},${lon}`,    
        locationName,       
        '',                 
        '',                 
        ''                  
      ]);

      const mainRowIndex = newRow.rowNumber;
      console.log(`✅ Added to MAIN sheet at row: ${mainRowIndex}`);

      const onWorkSheet = await this.getSheet(CONFIG.SHEETS.ON_WORK);
      await onWorkSheet.addRow([
        timestamp,          
        employee,           
        timestamp,          
        'ทำงาน',           
        userinfo || '',     
        `${lat},${lon}`,    
        locationName,       
        mainRowIndex,       
        line_name,          
        line_picture,       
        mainRowIndex,       
        employee            
      ]);      // Clear cache เนื่องจากมีการเพิ่มข้อมูลใหม่
      this.clearCache('onwork');
      this.clearCache('main');
      this.clearCache('stats');

      console.log(`✅ Clock In successful: ${employee} at ${this.formatTime(timestamp)}, Main row: ${mainRowIndex}`);

      // ทำการ warm cache อัตโนมัติ
      setTimeout(async () => {
        try {
          await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
          await this.getAdminStats();
        } catch (error) {
          console.error('⚠️ Auto cache warming error:', error);
        }
      }, 2000);

      this.triggerMapGeneration('clockin', {
        employee, lat, lon, line_name, userinfo, timestamp
      });

      return {
        success: true,
        message: 'บันทึกเวลาเข้างานสำเร็จ',
        employee,
        time: this.formatTime(timestamp),
        currentStatus: 'clocked_in'
      };

    } catch (error) {
      console.error('❌ Clock in error:', error);
      return {
        success: false,
        message: `เกิดข้อผิดพลาด: ${error.message}`,
        employee: data.employee
      };
    }
  }

  async clockOut(data) {
    try {
      const { employee, lat, lon, line_name, mock_time } = data;
      
      console.log(`⏰ Clock Out request for: "${employee}"`);
      console.log(`📍 Location: ${lat}, ${lon}`);
      if (mock_time) {
        console.log(`🧪 Using mock time: ${mock_time}`);
      }
      
      const employeeStatus = await this.getEmployeeStatus(employee);
        if (!employeeStatus.isOnWork) {
        console.log(`❌ Employee "${employee}" is not clocked in`);
        
        // ใช้ cached data แทนการเรียก API ใหม่
        const rows = await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
        
        const suggestions = rows
          .map(row => ({
            systemName: row.get('ชื่อในระบบ'),
            employeeName: row.get('ชื่อพนักงาน')
          }))
          .filter(emp => emp.systemName || emp.employeeName)
          .filter(emp => 
            this.isNameMatch(employee, emp.systemName) ||
            this.isNameMatch(employee, emp.employeeName)
          );
        
        let message = 'คุณต้องลงเวลาเข้างานก่อน หรือตรวจสอบชื่อที่ป้อนให้ถูกต้อง';
        
        if (suggestions.length > 0) {
          const suggestedNames = suggestions.map(s => s.systemName || s.employeeName);
          message = `ไม่พบข้อมูลการลงเวลาเข้างาน ชื่อที่ใกล้เคียง: ${suggestedNames.join(', ')}`;
        }
        
        return {
          success: false,
          message: message,
          employee,
          currentStatus: 'not_clocked_in',
          suggestions: suggestions.length > 0 ? suggestions : undefined
        };
      }

      // ใช้ mock_time หากมีการส่งมา ไม่เช่นนั้นใช้เวลาปัจจุบัน
      const timestamp = mock_time 
        ? moment(mock_time).tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss')
        : moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY HH:mm:ss');
      const workRecord = employeeStatus.workRecord;
      const clockInTime = workRecord.clockIn;
      console.log(`⏰ Clock in time: ${clockInTime}`);
      
      // 🎯 ใช้ฟังก์ชันคำนวณเวลาแบบเดียวกันกับ admin stats
      const hoursWorked = calculateWorkingHours(clockInTime, timestamp);
      console.log(`✅ Working hours calculated: ${hoursWorked.toFixed(2)} hours`);
      
      // แปลงพิกัดเป็นชื่อสถานที่
      const locationName = await this.getLocationName(lat, lon);
      console.log(`📍 Clock out location: ${locationName}`);      console.log(`✅ Proceeding with clock out for "${employee}"`);
      
      // ใช้ cached data แทนการเรียก API ใหม่
      const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
      const rows = await this.getCachedSheetData(CONFIG.SHEETS.MAIN);
      
      console.log(`📊 Total rows in MAIN: ${rows.length}`);
      console.log(`🎯 Target row index: ${workRecord.mainRowIndex}`);
      
      let mainRow = null;
      
      if (workRecord.mainRowIndex && workRecord.mainRowIndex > 1) {
        const targetIndex = workRecord.mainRowIndex - 2;
        
        if (targetIndex >= 0 && targetIndex < rows.length) {
          const candidateRow = rows[targetIndex];
          const candidateEmployee = candidateRow.get('ชื่อพนักงาน');
          
          if (this.isNameMatch(employee, candidateEmployee)) {
            mainRow = candidateRow;
            console.log(`✅ Found main row by index: ${targetIndex} (row ${workRecord.mainRowIndex})`);
          } else {
            console.log(`⚠️ Row index found but employee name mismatch: "${candidateEmployee}" vs "${employee}"`);
          }
        } else {
          console.log(`⚠️ Row index out of range: ${targetIndex} (total rows: ${rows.length})`);
        }
      }
      
      if (!mainRow) {
        console.log('🔍 Searching by employee name and conditions...');
        
        const candidateRows = rows.filter(row => {
          const rowEmployee = row.get('ชื่อพนักงาน');
          const rowClockOut = row.get('เวลาออก');
          
          return this.isNameMatch(employee, rowEmployee) && !rowClockOut;
        });
        
        console.log(`Found ${candidateRows.length} candidate rows without clock out`);
        
        if (candidateRows.length === 1) {
          mainRow = candidateRows[0];
          console.log(`✅ Found unique candidate row`);
        } else if (candidateRows.length > 1) {
          let closestRow = null;
          let minTimeDiff = Infinity;
          
          candidateRows.forEach((row, index) => {
            const rowClockIn = row.get('เวลาเข้า');
            if (rowClockIn && clockInTime) {
              const timeDiff = Math.abs(new Date(rowClockIn) - new Date(clockInTime));
              console.log(`Candidate ${index}: time diff = ${timeDiff}ms`);
              if (timeDiff < minTimeDiff) {
                minTimeDiff = timeDiff;
                closestRow = row;
              }
            }
          });
          
          if (closestRow && minTimeDiff < 300000) {
            mainRow = closestRow;
            console.log(`✅ Found closest matching row (time diff: ${minTimeDiff}ms)`);
          } else {
            console.log(`❌ No close time match found (min diff: ${minTimeDiff}ms)`);
          }
        }
      }
      
      if (!mainRow) {
        console.log('🔍 Searching for latest row of this employee...');
        
        for (let i = rows.length - 1; i >= 0; i--) {
          const row = rows[i];
          const rowEmployee = row.get('ชื่อพนักงาน');
          const rowClockOut = row.get('เวลาออก');
          
          if (this.isNameMatch(employee, rowEmployee) && !rowClockOut) {
            mainRow = row;
            console.log(`✅ Found latest uncompleted row at index: ${i}`);
            break;
          }
        }
      }
      
      if (!mainRow) {
        console.log('❌ Cannot find main row to update');
        
        return {
          success: false,
          message: 'ไม่พบข้อมูลการลงเวลาเข้างานที่ตรงกัน กรุณาตรวจสอบระบบ',
          employee
        };
      }
      
      console.log('✅ Found main row, updating...');
      
      try {
        // 🔧 ใช้วิธี batch update เพื่อป้องกันการเปลี่ยนรูปแบบเวลาเข้า
        const sheet = await this.getSheet(CONFIG.SHEETS.MAIN);
        const rowNumber = mainRow.rowNumber;
        
        console.log(`📝 Updating row ${rowNumber} using batch update to preserve format`);
        
        // อัปเดตเฉพาะเซลล์ที่จำเป็น โดยไม่แตะเซลล์เวลาเข้า (column D)
        const updates = [];
        
        // Column F: เวลาออก (index 5)
        updates.push({
          range: `F${rowNumber}`,
          values: [[timestamp]]
        });
        
        // Column I: พิกัดออก (index 8) 
        updates.push({
          range: `I${rowNumber}`,
          values: [[`${lat},${lon}`]]
        });
        
        // Column J: ที่อยู่ออก (index 9)
        updates.push({
          range: `J${rowNumber}`,
          values: [[locationName]]
        });
        
        // Column K: ชั่วโมงทำงาน (index 10)
        updates.push({
          range: `K${rowNumber}`,
          values: [[hoursWorked.toFixed(2)]]
        });
        
        // ทำการอัปเดตทีละเซลล์
        for (const update of updates) {
          await sheet.loadCells(update.range);
          const cell = sheet.getCellByA1(update.range);
          
          // เซ็ตค่าเฉพาะข้อมูล ไม่ตั้งค่า format ให้ Google Sheets จัดการเอง
          cell.value = update.values[0][0];
        }
        
        await sheet.saveUpdatedCells();
        console.log('✅ Main row updated successfully using batch update (clock-in format preserved)');
        
      } catch (updateError) {
        console.error('❌ Error updating main row:', updateError);
        throw new Error('ไม่สามารถอัปเดตข้อมูลได้: ' + updateError.message);
      }      try {
        await workRecord.row.delete();
        console.log('✅ Removed from ON_WORK sheet');
          // Clear cache เนื่องจากมีการเปลี่ยนแปลงข้อมูล
        this.clearCache('onwork');
        this.clearCache('main');
        this.clearCache('stats');
        
        // ทำการ warm cache อัตโนมัติ
        setTimeout(async () => {
          try {
            await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
            await this.getAdminStats();
          } catch (error) {
            console.error('⚠️ Auto cache warming error:', error);
          }
        }, 2000);
        
      } catch (deleteError) {
        console.error('❌ Error deleting from ON_WORK:', deleteError);
      }

      console.log(`✅ Clock Out successful: ${employee} at ${this.formatTime(timestamp)} (${hoursWorked.toFixed(2)} hours)`);

      try {
        this.triggerMapGeneration('clockout', {
          employee, lat, lon, line_name, timestamp, hoursWorked
        });
      } catch (webhookError) {
        console.error('⚠️ Webhook error (non-critical):', webhookError);
      }

      return {
        success: true,
        message: 'บันทึกเวลาออกงานสำเร็จ',
        employee,
        time: this.formatTime(timestamp),
        hours: hoursWorked.toFixed(2),
        currentStatus: 'clocked_out'
      };

    } catch (error) {
      console.error('❌ Clock out error:', error);
      
      return {
        success: false,
        message: `เกิดข้อผิดพลาด: ${error.message}`,
        employee: data.employee
      };
    }
  }

  async triggerMapGeneration(action, data) {
    try {
      const gsaWebhookUrl = process.env.GSA_MAP_WEBHOOK_URL;
      if (!gsaWebhookUrl) {
        console.log('⚠️ GSA webhook URL not configured');
        return;
      }      const payload = {
        action,
        data,
        timestamp: moment().tz(CONFIG.TIMEZONE).toISOString() // ใช้เวลาไทย
      };

      await fetch(gsaWebhookUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-Webhook-Secret': CONFIG.RENDER.GSA_WEBHOOK_SECRET
        },
        body: JSON.stringify(payload)
      });

      console.log(`📍 Map generation triggered for ${action}: ${data.employee}`);
      
    } catch (error) {
      console.error('Error triggering map generation:', error);
    }
  }  formatTime(date) {
    try {
      // รองรับทั้ง Date object และ string
      if (typeof date === 'string') {
        // ถ้าเป็นรูปแบบ 'YYYY-MM-DD HH:mm:ss' จาก moment
        if (date.includes(' ') && date.length === 19) {
          return date.split(' ')[1]; // ใช้ส่วนเวลาเท่านั้น
        }
        // ลองแปลงเป็น Date object
        const parsedDate = moment(date).tz(CONFIG.TIMEZONE);
        if (parsedDate.isValid()) {
          return parsedDate.format('HH:mm:ss');
        }
        return date; // ถ้าแปลงไม่ได้ ส่งกลับเป็น string เดิม
      }
      
      // ถ้าเป็น Date object
      if (date instanceof Date && !isNaN(date.getTime())) {
        return moment(date).tz(CONFIG.TIMEZONE).format('HH:mm:ss');
      }
      
      return '';
    } catch (error) {
      console.error('Error formatting time:', error);
      return date?.toString() || '';
    }
  }

  // เพิ่มฟังก์ชันแปลงพิกัดเป็นชื่อสถานที่
  async getLocationName(lat, lon) {
    try {
      // ใช้ OpenStreetMap Nominatim API (ฟรี)
      const response = await fetch(
        `https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lon}&zoom=18&addressdetails=1&accept-language=th`
      );
      
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      
      const data = await response.json();
      
      if (data && data.display_name) {
        // ใช้ชื่อสถานที่ที่ได้จาก API
        return data.display_name;
      } else {
        // ถ้าไม่ได้ข้อมูล ใช้พิกัดแทน
        return `${lat}, ${lon}`;
      }    } catch (error) {
      console.warn(`⚠️ Location lookup failed for ${lat}, ${lon}:`, error.message);
      // ถ้าผิดพลาด ใช้พิกัดแทน
      return `${lat}, ${lon}`;
    }
  }

  // ฟังก์ชันสำหรับตรวจสอบและจัดการกรณีลืมลงเวลาออก

  async checkAndHandleMissedCheckouts() {
    try {
      console.log('🔍 Starting automatic missed checkout check...');
      
      // ดึงข้อมูลพนักงานที่ยังอยู่ในระหว่างทำงาน (ON_WORK sheet)
      const onWorkRows = await this.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
      
      if (onWorkRows.length === 0) {
        console.log('✅ No employees currently on work, no missed checkouts to handle');
        return { success: true, processedCount: 0, message: 'No employees on work' };
      }
  
      console.log(`📊 Found ${onWorkRows.length} employees currently on work`);
      
      const today = moment().tz(CONFIG.TIMEZONE);
      const cutoffTime = today.clone().set({
        hour: CONFIG.AUTO_CHECKOUT.CUTOFF_HOUR,
        minute: CONFIG.AUTO_CHECKOUT.CUTOFF_MINUTE,
        second: 59,
        millisecond: 999
      });
      
      console.log(`⏰ Processing missed checkouts for cutoff time: ${cutoffTime.format('YYYY-MM-DD HH:mm:ss')}`);
      console.log(`🛡️ Exempt employees: ${CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES.join(', ')}`);
      
      let processedCount = 0;
      let exemptedCount = 0;
      const results = [];
      
      // ประมวลผลแต่ละพนักงานที่ยังไม่ได้ลงเวลาออก
      for (const workRow of onWorkRows) {
        try {
          const employeeName = workRow.get('ชื่อพนักงาน') || workRow.get('ชื่อในระบบ');
          const clockInTime = workRow.get('เวลาเข้า');
          const mainRowIndex = workRow.get('แถวในMain') || workRow.get('แถวอ้างอิง');
          
          if (!employeeName || !clockInTime) {
            console.warn(`⚠️ Missing data for work record: ${employeeName || 'Unknown'}`);
            continue;
          }
  
          // 🛡️ ตรวจสอบว่าเป็นพนักงานที่ยกเว้นหรือไม่
          const isExempt = this.isEmployeeExempt(employeeName);
          if (isExempt) {
            console.log(`🛡️ EXEMPT: ${employeeName} - skipping auto checkout (night guard)`);
            exemptedCount++;
            results.push({
              employee: employeeName,
              action: 'exempted',
              reason: 'Night guard - exempt from auto checkout',
              clockIn: clockInTime
            });
            continue;
          }
  
          // ตรวจสอบว่าเวลาเข้าเป็นวันนี้หรือไม่
          const clockInMoment = moment.tz(clockInTime, 'YYYY-MM-DD H:mm:ss', CONFIG.TIMEZONE);
          const isToday = clockInMoment.format('YYYY-MM-DD') === today.format('YYYY-MM-DD');
          
          if (!isToday) {
            console.log(`⏭️ Skipping ${employeeName} - not clocked in today (${clockInMoment.format('YYYY-MM-DD')})`);
            continue;
          }
  
          console.log(`🔄 Processing missed checkout for: ${employeeName}`);
          console.log(`⏰ Clock in time: ${clockInTime}`);
          console.log(`📍 Main row index: ${mainRowIndex}`);
  
          // อัปเดต MAIN sheet ด้วยข้อมูลลืมลงเวลาออก
          const result = await this.processMissedCheckout({
            employeeName,
            clockInTime,
            mainRowIndex,
            cutoffTime,
            workRow
          });
  
          if (result.success) {
            processedCount++;
            results.push({
              employee: employeeName,
              action: 'missed_checkout_processed',
              clockIn: clockInTime,
              autoClockOut: result.autoClockOut
            });
            
            console.log(`✅ Processed missed checkout for ${employeeName}`);
          } else {
            console.error(`❌ Failed to process missed checkout for ${employeeName}: ${result.error}`);
            results.push({
              employee: employeeName,
              action: 'failed',
              error: result.error
            });
          }
  
        } catch (error) {
          console.error(`❌ Error processing missed checkout for employee:`, error);
          results.push({
            employee: workRow.get('ชื่อพนักงาน') || 'Unknown',
            action: 'error',
            error: error.message
          });
        }
      }
  
      console.log(`✅ Missed checkout check completed.`);
      console.log(`   📊 Total checked: ${onWorkRows.length}`);
      console.log(`   ✅ Processed: ${processedCount}`);
      console.log(`   🛡️ Exempted: ${exemptedCount}`);
      
      // ส่ง notification ถ้ามีการประมวลผลหรือการยกเว้น
      if (processedCount > 0 || exemptedCount > 0) {
        await this.sendMissedCheckoutNotification(results, processedCount, exemptedCount);
      }
  
      return {
        success: true,
        processedCount,
        exemptedCount,
        totalChecked: onWorkRows.length,
        results,
        message: `Processed ${processedCount} missed checkouts, exempted ${exemptedCount} employees`
      };
  
    } catch (error) {
      console.error('❌ Error in checkAndHandleMissedCheckouts:', error);
      return {
        success: false,
        error: error.message,
        processedCount: 0,
        exemptedCount: 0
      };
    }
  }

  // ฟังก์ชันสำหรับประมวลผลลืมลงเวลาออกของพนักงานคนหนึ่ง
  async processMissedCheckout({ employeeName, clockInTime, mainRowIndex, cutoffTime, workRow }) {
    try {
      // 🎯 ใช้ฟังก์ชันคำนวณเวลาแบบเดียวกันกับ clock out
      const autoClockOutTime = cutoffTime.format('DD/MM/YYYY HH:mm:ss');
      const hoursWorked = calculateWorkingHours(clockInTime, autoClockOutTime);
      
      // ข้อความที่จะเขียนลง sheet (คอลัมน์ E)
      const missedCheckoutNote = 'ลืมลงเวลาออก (ระบบอัตโนมัติ)';
      
      console.log(`⏰ Auto clock out for ${employeeName}: ${autoClockOutTime} (${hoursWorked.toFixed(2)} hours)`);
      console.log(`📝 Note will be written to column E: "${missedCheckoutNote}"`);

      // อัปเดต MAIN sheet
      if (mainRowIndex && !isNaN(parseInt(mainRowIndex))) {
        try {
          // 🔧 ใช้วิธี batch update เพื่อป้องกันการเปลี่ยนรูปแบบเวลาเข้า (เดียวกับ clockOut)
          const mainSheet = await this.getSheet(CONFIG.SHEETS.MAIN);
          const rowNumber = parseInt(mainRowIndex);
          
          console.log(`📝 Updating auto checkout for row ${rowNumber} using batch update to preserve format`);
          
          // อัปเดตเฉพาะเซลล์ที่จำเป็น โดยไม่แตะเซลล์เวลาเข้า (column D)
          const updates = [];
          
          // Column E: หมายเหตุ (userinfo) (index 4)
          updates.push({
            range: `E${rowNumber}`,
            values: [[missedCheckoutNote]]
          });
          
          // Column F: เวลาออก (index 5)
          updates.push({
            range: `F${rowNumber}`,
            values: [[autoClockOutTime]]
          });
          
          // Column K: ชั่วโมงทำงาน (index 10)
          updates.push({
            range: `K${rowNumber}`,
            values: [[hoursWorked.toFixed(2)]]
          });
          
          // ทำการอัปเดตทีละเซลล์
          for (const update of updates) {
            await mainSheet.loadCells(update.range);
            const cell = mainSheet.getCellByA1(update.range);
            
            // เซ็ตค่าเฉพาะข้อมูล ไม่ตั้งค่า format ให้ Google Sheets จัดการเอง
            cell.value = update.values[0][0];
          }
          
          await mainSheet.saveUpdatedCells();
          console.log(`✅ Updated MAIN sheet row ${rowNumber} for ${employeeName} using batch update (auto checkout format preserved)`);
          
        } catch (updateError) {
          console.error(`❌ Error updating auto checkout for ${employeeName}:`, updateError);
          throw new Error('ไม่สามารถอัปเดตข้อมูลลืมลงเวลาออกได้: ' + updateError.message);
        }
      }

      // ลบออกจาก ON_WORK sheet
      await workRow.delete();
      console.log(`✅ Removed ${employeeName} from ON_WORK sheet`);
      
      // ล้าง cache ที่เกี่ยวข้อง
      this.clearCache('onwork');
      this.clearCache('main');

      return {
        success: true,
        autoClockOut: autoClockOutTime,
        hoursWorked: hoursWorked.toFixed(2),
        note: missedCheckoutNote
      };

    } catch (error) {
      console.error(`❌ Error processing missed checkout for ${employeeName}:`, error);
      return {
        success: false,
        error: error.message
      };
    }
  }

  // ฟังก์ชันส่ง notification เมื่อมีการประมวลผลลืมลงเวลาออก
  async sendMissedCheckoutNotification(results, processedCount, exemptedCount = 0) {
    try {
      if (!CONFIG.TELEGRAM.BOT_TOKEN || !CONFIG.TELEGRAM.CHAT_ID) {
        console.log('⚠️ Telegram notification not configured for missed checkout alerts');
        return;
      }

      const successfulResults = results.filter(r => r.action === 'missed_checkout_processed');
      const exemptedResults = results.filter(r => r.action === 'exempted');
      const failedResults = results.filter(r => r.action === 'failed' || r.action === 'error');
      
      const today = moment().tz(CONFIG.TIMEZONE).format('DD/MM/YYYY');
      
      let message = `🤖 *รายงานลงเวลาออกอัตโนมัติ - ${today}*\n\n`;
      message += `📊 สรุปผล:\n`;
      message += `   ✅ ลงเวลาออกอัตโนมัติ: ${processedCount} คน\n`;
      message += `   🛡️ ยกเว้น (ยามกลางคืน): ${exemptedCount} คน\n`;
      message += `   ❌ ไม่สำเร็จ: ${failedResults.length} คน\n\n`;
      
      if (exemptedResults.length > 0) {
        message += `🛡️ *พนักงานที่ได้รับการยกเว้น:*\n`;
        exemptedResults.forEach(result => {
          const clockInTime = moment(result.clockIn).tz(CONFIG.TIMEZONE).format('HH:mm');
          message += `• ${result.employee} - เข้างาน ${clockInTime} (ยามกลางคืน)\n`;
        });
        message += '\n';
      }

      if (successfulResults.length > 0) {
        message += `✅ *ดำเนินการสำเร็จ:*\n`;
        successfulResults.forEach(result => {
          const clockOutTime = moment(result.autoClockOut).tz(CONFIG.TIMEZONE).format('HH:mm');
          message += `• ${result.employee} - ลงเวลาออกอัตโนมัติ ${clockOutTime}\n`;
        });
        message += '\n';
      }
      
      if (failedResults.length > 0) {
        message += `❌ *ดำเนินการไม่สำเร็จ:*\n`;
        failedResults.forEach(result => {
          message += `• ${result.employee} - ${result.error}\n`;
        });
        message += '\n';
      }
      
      message += `⏰ เวลาประมวลผล: ${moment().tz(CONFIG.TIMEZONE).format('HH:mm:ss')}\n`;
      message += `💡 พนักงานสามารถลงเวลาเข้างานวันใหม่ได้ปกติ\n`;
      message += `🛡️ พนักงานยกเว้น: ${CONFIG.AUTO_CHECKOUT.EXEMPT_EMPLOYEES.join(', ')}\n`;
      message += `📝 หมายเหตุ "ลืมลงเวลาออก (ระบบอัตโนมัติ)" ถูกเขียนลงคอลัมน์ E ใน Google Sheet`;

      // ส่งข้อความไปยัง Telegram
      const telegramUrl = `https://api.telegram.org/bot${CONFIG.TELEGRAM.BOT_TOKEN}/sendMessage`;
      
      await fetch(telegramUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          chat_id: CONFIG.TELEGRAM.CHAT_ID,
          text: message,
          parse_mode: 'Markdown'
        })
      });

      console.log('✅ Missed checkout notification sent to Telegram');

    } catch (error) {
      console.error('❌ Error sending missed checkout notification:', error);
    }
  }

  // Emergency mode functions
  setEmergencyMode(enabled) {
    this.emergencyMode = enabled;
    if (enabled) {
      console.log('🚨 Emergency mode ENABLED - Using cached data only');
      // ขยาย TTL ของ cache เป็น 1 ชั่วโมง
      Object.keys(this.cache).forEach(key => {
        this.cache[key].ttl = 3600000; // 1 hour
      });
    } else {
      console.log('✅ Emergency mode DISABLED - Normal operation resumed');
      // คืนค่า TTL เดิม
      this.cache.employees.ttl = 300000; // 5 minutes
      this.cache.onwork.ttl = 60000;     // 1 minute
      this.cache.main.ttl = 30000;       // 30 seconds
      this.cache.stats.ttl = 120000;     // 2 minutes
    }
  }

  async safeGetCachedSheetData(sheetName) {
    try {
      return await this.getCachedSheetData(sheetName);
    } catch (error) {
      console.error(`❌ Failed to get data for ${sheetName}:`, error.message);
      
      // เข้าสู่ emergency mode
      if (!this.emergencyMode) {
        this.setEmergencyMode(true);
      }
      
      // คืนค่า cache เก่า (ถ้ามี)
      const staleData = this.getCache(sheetName.toLowerCase().replace(/\s+/g, ''));
      if (staleData) {
        console.log(`📋 Using emergency cache for ${sheetName}`);
        return staleData;
      }
      
      // ถ้าไม่มี cache เลย คืนค่า array ว่าง
      console.warn(`⚠️ No cache available for ${sheetName}, returning empty data`);
      return [];
    }
  }
}

// ========== Initialize Services ==========
const sheetsService = new GoogleSheetsService();
const keepAliveService = new KeepAliveService();

// ========== Admin Authentication Routes ==========

// Admin Login
app.post('/api/admin/login', async (req, res) => {
  try {
    const { username, password } = req.body;

    if (!username || !password) {
      return res.status(400).json({
        success: false,
        message: 'กรุณากรอกชื่อผู้ใช้และรหัสผ่าน'
      });
    }

    // ค้นหาผู้ใช้
    const user = CONFIG.ADMIN.USERS.find(u => u.username === username);
    if (!user) {
      return res.status(401).json({
        success: false,
        message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'
      });
    }

    // ตรวจสอบรหัสผ่าน
    const isValidPassword = await bcrypt.compare(password, user.password);
    if (!isValidPassword) {
      return res.status(401).json({
        success: false,
        message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง'
      });
    }

    // สร้าง JWT token
    const token = jwt.sign(
      { 
        id: user.id, 
        username: user.username, 
        role: user.role 
      },
      CONFIG.ADMIN.JWT_SECRET,
      { expiresIn: CONFIG.ADMIN.JWT_EXPIRES_IN }
    );

    res.json({
      success: true,
      message: 'เข้าสู่ระบบสำเร็จ',
      token,
      user: {
        id: user.id,
        username: user.username,
        name: user.name,
        role: user.role
      }
    });

  } catch (error) {
    console.error('Admin login error:', error);
    res.status(500).json({
      success: false,
      message: 'เกิดข้อผิดพลาดภายในระบบ'
    });
  }
});

// Verify Token
app.get('/api/admin/verify-token', authenticateAdmin, (req, res) => {
  res.json({
    success: true,
    user: req.user
  });
});

// Token Refresh - สำหรับต่ออายุ token ที่ใกล้หมดอายุ
app.post('/api/admin/refresh-token', async (req, res) => {
  try {
    const authHeader = req.headers.authorization;
    const token = authHeader && authHeader.split(' ')[1];

    if (!token) {
      return res.status(401).json({
        success: false,
        error: 'Token required for refresh'
      });
    }

    // ตรวจสอบ token แม้ว่าจะหมดอายุแล้ว
    let decoded;
    try {
      decoded = jwt.verify(token, CONFIG.ADMIN.JWT_SECRET);
    } catch (error) {
      if (error.name === 'TokenExpiredError') {
        // อนุญาตให้ refresh token ที่หมดอายุไม่เกิน 7 วัน
        const expiredAt = new Date(error.expiredAt);
        const now = new Date();
        const daysSinceExpired = (now - expiredAt) / (1000 * 60 * 60 * 24);
        
        if (daysSinceExpired <= 7) {
          // ถอดรหัส token โดยไม่ตรวจสอบวันหมดอายุ
          decoded = jwt.verify(token, CONFIG.ADMIN.JWT_SECRET, { ignoreExpiration: true });
        } else {
          return res.status(401).json({
            success: false,
            error: 'Token expired too long ago. Please login again.',
            errorCode: 'TOKEN_EXPIRED_TOO_LONG'
          });
        }
      } else {
        return res.status(401).json({
          success: false,
          error: 'Invalid token',
          errorCode: 'INVALID_TOKEN'
        });
      }
    }

    // ตรวจสอบว่าผู้ใช้ยังคงอยู่ในระบบ
    const user = CONFIG.ADMIN.USERS.find(u => u.id === decoded.id);
    if (!user) {
      return res.status(401).json({
        success: false,
        error: 'User no longer exists'
      });
    }

    // สร้าง token ใหม่
    const newToken = jwt.sign(
      { 
        id: user.id, 
        username: user.username,
        name: user.name,
        role: user.role
      },
      CONFIG.ADMIN.JWT_SECRET,
      { expiresIn: CONFIG.ADMIN.JWT_EXPIRES_IN }
    );

    console.log(`🔄 Token refreshed for user: ${user.username}`);

    res.json({
      success: true,
      message: 'Token refreshed successfully',
      token: newToken,
      user: {
        id: user.id,
        username: user.username,
        name: user.name,
        role: user.role
      }
    });

  } catch (error) {
    console.error('Token refresh error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to refresh token'
    });
  }
});

// Admin Stats
app.get('/api/admin/stats', authenticateAdmin, async (req, res) => {
  try {
    const stats = await sheetsService.getAdminStats();
    res.json({
      success: true,
      data: stats
    });
  } catch (error) {
    console.error('Admin stats error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get stats'
    });
  }
});

// Export Routes
app.get('/api/admin/export/:type', authenticateAdmin, async (req, res) => {
  try {
    const { type } = req.params;
    const params = req.query;

    // ตรวจสอบประเภทรายงาน
    if (!['daily', 'monthly', 'range'].includes(type)) {
      return res.status(400).json({
        success: false,
        error: 'Invalid report type'
      });
    }

    // Log format parameter เพื่อ debug
    console.log(`📊 Export request: type=${type}, format=${params.format || 'default'}`);

    // ดึงข้อมูลจาก Google Sheets
    const reportData = await sheetsService.getReportData(type, params);

    // สร้างไฟล์ Excel
    const workbook = await ExcelExportService.createWorkbook(reportData, type, params);

    // ตั้งชื่อไฟล์ตาม format
    let filename = 'report.xlsx';
    if (type === 'monthly' && params.format === 'detailed') {
      filename = 'monthly_detailed_report.xlsx';
    } else if (type === 'monthly') {
      filename = 'monthly_summary_report.xlsx';
    } else {
      filename = `${type}_report.xlsx`;
    }

    // ตั้งค่า response headers
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${filename}`);

    // ส่งไฟล์
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to export report'
    });
  }
});

// ========== API Rate Limiting และ Monitoring ==========
class APIMonitor {
  constructor() {
    this.apiCalls = [];
    // ปรับเพิ่ม rate limit เพื่อรองรับ concurrent users มากขึ้น
    this.maxCallsPerMinute = 100; // เพิ่มจาก 30 เป็น 100 ครั้งต่อนาที
    this.maxCallsPerHour = 1000; // เพิ่มจาก 300 เป็น 1000 ครั้งต่อชั่วโมง
    
    // เพิ่ม burst allowance สำหรับ peak time
    this.burstLimit = 75; // เพิ่มจาก 50 เป็น 75 concurrent requests
    this.currentBurst = 0;
    this.lastBurstReset = Date.now();
    
    // Auto-reset burst counter every 5 seconds
    setInterval(() => {
      if (this.currentBurst > 0) {
        console.log(`🔄 Auto-resetting burst counter from ${this.currentBurst} to 0`);
        this.currentBurst = 0;
      }
    }, 5000); // 5 วินาที
  }

  logAPICall(operation) {
    const now = new Date();
    this.apiCalls.push({
      timestamp: now,
      operation: operation
    });

    // ลบ logs ที่เก่าเกิน 1 ชั่วโมง
    this.apiCalls = this.apiCalls.filter(call => 
      (now - call.timestamp) < 3600000 // 1 hour
    );

    console.log(`📊 API Call: ${operation} (Total in last hour: ${this.apiCalls.length}, Current burst: ${this.currentBurst})`);
  }

  canMakeAPICall() {
    const now = new Date();
    
    // นับจำนวน API calls ในนาทีที่แล้ว
    const callsInLastMinute = this.apiCalls.filter(call => 
      (now - call.timestamp) < 60000 // 1 minute
    ).length;

    // นับจำนวน API calls ในชั่วโมงที่แล้ว
    const callsInLastHour = this.apiCalls.length;

    // ตรวจสอบ burst limit
    if (this.currentBurst >= this.burstLimit) {
      console.warn(`⚠️ Burst limit exceeded: ${this.currentBurst}/${this.burstLimit} concurrent requests`);
      return false;
    }

    if (callsInLastMinute >= this.maxCallsPerMinute) {
      console.warn(`⚠️ Rate limit exceeded: ${callsInLastMinute} calls in last minute`);
      return false;
    }

    if (callsInLastHour >= this.maxCallsPerHour) {
      console.warn(`⚠️ Rate limit exceeded: ${callsInLastHour} calls in last hour`);
      return false;
    }

    // เพิ่ม burst counter
    this.currentBurst++;

    return true;
  }

  // เมื่อ API call เสร็จแล้ว ลด burst counter
  finishCall() {
    if (this.currentBurst > 0) {
      this.currentBurst--;
    }
  }

  getStats() {
    const now = new Date();
    const callsInLastMinute = this.apiCalls.filter(call => 
      (now - call.timestamp) < 60000
    ).length;
    const callsInLastHour = this.apiCalls.length;

    return {
      callsInLastMinute,
      callsInLastHour,
      maxCallsPerMinute: this.maxCallsPerMinute,
      maxCallsPerHour: this.maxCallsPerHour,
      percentageUsedPerMinute: (callsInLastMinute / this.maxCallsPerMinute) * 100,
      percentageUsedPerHour: (callsInLastHour / this.maxCallsPerHour) * 100
    };
  }
}

const apiMonitor = new APIMonitor();

// ========== Original Routes (unchanged) ==========

// Home page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Health check และ ping endpoint
app.get('/debug/sheet-info', async (req, res) => {
  try {
    console.log('🔍 Debug: Getting sheet info...');
    
    const mainSheet = await sheetsService.getSheet(CONFIG.SHEETS.MAIN);
    const rows = await mainSheet.getRows({ limit: 5 });
    
    if (rows.length > 0) {
      const headers = Object.keys(rows[0]._rawData);
      const firstRowData = rows[0]._rawData;
      
      console.log('📋 MAIN Sheet Headers:', headers);
      console.log('📋 First row data:', firstRowData);
      
      res.json({
        sheetTitle: mainSheet.title,
        headerCount: headers.length,
        headers: headers,
        firstRowData: firstRowData,
        sampleRows: rows.map((row, index) => ({
          rowIndex: index,
          employee: row.get('ชื่อพนักงาน'),
          clockIn: row.get('เวลาเข้า'),
          clockOut: row.get('เวลาออก'),
          rawData: row._rawData
        }))
      });
    } else {
      res.json({ error: 'No data found' });
    }
    
  } catch (error) {
    console.error('❌ Debug sheet info error:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/health', (req, res) => {  res.json({
    status: 'healthy',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString(), // ใช้เวลาไทย
    uptime: process.uptime(),
    keepAlive: keepAliveService.getStats(),
    environment: process.env.NODE_ENV || 'development',
    config: {
      hasLiffId: !!CONFIG.LINE.LIFF_ID,
      liffIdLength: CONFIG.LINE.LIFF_ID ? CONFIG.LINE.LIFF_ID.length : 0
    }
  });
});

// Ping endpoint สำหรับ keep-alive
app.get('/api/ping', (req, res) => {
  res.json({
    status: 'pong',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString(), // ใช้เวลาไทย
    uptime: process.uptime()
  });
});

// Webhook endpoint สำหรับรับ ping จาก GSA
app.post('/api/webhook/ping', (req, res) => {
  console.log('📨 Received ping from GSA');  res.json({
    status: 'received',
    timestamp: moment().tz(CONFIG.TIMEZONE).toISOString() // ใช้เวลาไทย
  });
});

// API สำหรับ Client Configuration
app.get('/api/config', (req, res) => {
  try {
    res.json({
      success: true,
      data: {
        liffId: CONFIG.LINE.LIFF_ID,
        apiUrl: CONFIG.RENDER.SERVICE_URL + '/api',
        environment: process.env.NODE_ENV || 'development',
        features: {
          keepAlive: CONFIG.RENDER.KEEP_ALIVE_ENABLED,
          liffEnabled: !!CONFIG.LINE.LIFF_ID
        }
      }
    });
  } catch (error) {
    console.error('API Error - config:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get config'
    });
  }
});

// Get employees
app.post('/api/employees', async (req, res) => {
  try {
    const employees = await sheetsService.getEmployees();
    res.json({
      success: true,
      data: employees
    });
  } catch (error) {
    console.error('API Error - employees:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get employees'
    });
  }
});

// Clock in
app.post('/api/clockin', async (req, res) => {
  try {
    const { employee, userinfo, lat, lon, line_name, line_picture, mock_time } = req.body;
    
    if (!employee || !lat || !lon) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields'
      });
    }

    // ตรวจสอบ rate limit
    if (!apiMonitor.canMakeAPICall()) {
      return res.status(429).json({
        success: false,
        error: 'Too many requests, please try again later'
      });
    }

    apiMonitor.logAPICall('clockIn');
    const result = await sheetsService.clockIn({
      employee, userinfo, lat, lon, line_name, line_picture, mock_time
    });
    
    // ลด burst counter หลังจาก API call เสร็จ
    apiMonitor.finishCall();

    res.json(result);
    
  } catch (error) {
    // ลด burst counter ถึงแม้จะ error
    apiMonitor.finishCall();
    console.error('API Error - clockin:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to clock in'
    });
  }
});

// Clock out
app.post('/api/clockout', async (req, res) => {
  try {
    const { employee, lat, lon, line_name, mock_time } = req.body;
    
    if (!employee || !lat || !lon) {
      return res.status(400).json({
        success: false,
        error: 'Missing required fields'
      });
    }

    // ตรวจสอบ rate limit
    if (!apiMonitor.canMakeAPICall()) {
      return res.status(429).json({
        success: false,
        error: 'Too many requests, please try again later'
      });
    }

    apiMonitor.logAPICall('clockOut');
    const result = await sheetsService.clockOut({
      employee, lat, lon, line_name, mock_time
    });
    
    // ลด burst counter หลังจาก API call เสร็จ
    apiMonitor.finishCall();

    res.json(result);
    
  } catch (error) {
    // ลด burst counter ถึงแม้จะ error
    apiMonitor.finishCall();
    console.error('API Error - clockout:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to clock out'
    });
  }
});

// API สำหรับตรวจสอบสถานะพนักงาน
app.post('/api/check-status', async (req, res) => {
  try {
    const { employee } = req.body;
    
    if (!employee) {
      return res.status(400).json({
        success: false,
        error: 'Missing employee name'
      });
    }    const employeeStatus = await sheetsService.getEmployeeStatus(employee);

    // ใช้ cached data แทนการเรียก API ใหม่
    const rows = await sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
    
    const currentEmployees = rows.map(row => ({
      systemName: row.get('ชื่อในระบบ'),
      employeeName: row.get('ชื่อพนักงาน'),
      clockIn: row.get('เวลาเข้า'),
      mainRowIndex: row.get('แถวในMain') || row.get('แถวอ้างอิง')
    }));

    res.json({
      success: true,
      data: {
        employee: employee,
        isOnWork: employeeStatus.isOnWork,
        hasWorkRecord: !!employeeStatus.workRecord,
        workRecord: employeeStatus.workRecord ? {
          clockIn: employeeStatus.workRecord.clockIn,
          mainRowIndex: employeeStatus.workRecord.mainRowIndex
        } : null,
        allCurrentEmployees: currentEmployees,
        suggestions: currentEmployees
          .filter(emp => 
            sheetsService.isNameMatch(employee, emp.systemName) ||
            sheetsService.isNameMatch(employee, emp.employeeName)
          )
          .map(emp => emp.systemName || emp.employeeName)
      }
    });

  } catch (error) {
    console.error('API Error - check-status:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to check status'
    });
  }
});

// API Monitoring endpoint
app.get('/api/admin/api-stats', authenticateAdmin, (req, res) => {
  const stats = apiMonitor.getStats();
  res.json({
    success: true,
    data: stats
  });
});

// API สำหรับ manual cache refresh (สำหรับ admin เท่านั้น)
app.post('/api/admin/refresh-cache', authenticateAdmin, async (req, res) => {
  try {
    console.log('🔄 Manual cache refresh initiated by admin');
    
    // Clear all cache
    sheetsService.clearCache();
    
    // Warm critical caches
    await sheetsService.getCachedSheetData(CONFIG.SHEETS.ON_WORK);
    await sheetsService.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
    
    res.json({
      success: true,
      message: 'Cache refreshed successfully'
    });
  } catch (error) {
    console.error('Cache refresh error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to refresh cache'
    });
  }
});

// API สำหรับตรวจสอบสถานะ API quota
app.get('/api/admin/quota-status', authenticateAdmin, async (req, res) => {
  try {
    const apiStats = apiMonitor.getStats();
    const isEmergencyMode = sheetsService.emergencyMode || false;
    
    // ทดสอบการเชื่อมต่อ API
    let apiHealthy = true;
    let lastError = null;
    
    try {
      await sheetsService.getCachedSheetData(CONFIG.SHEETS.EMPLOYEES);
    } catch (error) {
      apiHealthy = false;
      lastError = error.message;
    }
    
    res.json({
      success: true,
      data: {
        apiHealthy,
        emergencyMode: isEmergencyMode,
        lastError,
        apiStats,
        recommendations: apiHealthy ? 
          ['ระบบทำงานปกติ'] : 
          [
            'รอให้ quota reset (ภายใน 24 ชั่วโมง)',
            'ใช้ cached data ในระยะนี้',
            'ลดการใช้งานฟีเจอร์ที่ต้องใช้ API'
          ]
      }
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// API สำหรับเปิด/ปิด emergency mode
app.post('/api/admin/emergency-mode', authenticateAdmin, async (req, res) => {
  try {
    const { enabled } = req.body;
    sheetsService.setEmergencyMode(enabled);
    
    res.json({
      success: true,
      message: `Emergency mode ${enabled ? 'enabled' : 'disabled'}`,
      emergencyMode: enabled
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      error: error.message
    });
  }
});

// ========== Error Handling ==========
app.use((error, req, res, next) => {
  console.error('Global error handler:', error);
  res.status(500).json({
    success: false,
    error: 'Internal server error'
  });
});

app.use((req, res) => {
  res.status(404).json({
    success: false,
    error: 'Not found'
  });
});

// ========== Start Server ==========
async function startServer() {
  try {
    console.log('🚀 Starting Time Tracker Server with Admin Panel...');
    console.log(`🌍 Environment: ${process.env.NODE_ENV || 'development'}`);
    
    // ตรวจสอบ environment variables
    if (!validateConfig()) {
      console.error('❌ Server startup aborted due to missing configuration');
      process.exit(1);
    }

    // เริ่มต้น Google Sheets Service
    console.log('📊 Initializing Google Sheets Service...');
    await sheetsService.initialize();
    console.log('✅ Google Sheets Service initialized successfully');
    
    // เริ่มต้น Keep-Alive Service
    if (CONFIG.RENDER.KEEP_ALIVE_ENABLED) {
      console.log('🔄 Starting Keep-Alive Service...');
      keepAliveService.init();
    } else {
      console.log('⚠️ Keep-Alive Service is disabled');
    }

    // ตั้งค่า cron job สำหรับตรวจสอบลืมลงเวลาออก (ทุกวันเวลา 23:59:59)
    cron.schedule('59 59 23 * * *', async () => {
      console.log('🕚 Running daily missed checkout check at 23:59:59...');
      try {
        const result = await sheetsService.checkAndHandleMissedCheckouts();
        console.log(`✅ Missed checkout check completed: ${result.processedCount} employees processed`);
        
        // ส่งการแจ้งเตือนถ้ามีการประมวลผล
        if (result.processedCount > 0) {
          console.log(`📱 Auto-processed ${result.processedCount} missed checkouts`);
        }
      } catch (error) {
        console.error('❌ Error in missed checkout check:', error);
      }
    }, {
      scheduled: true,
      timezone: CONFIG.TIMEZONE
    });

    // เริ่มต้นเซิร์ฟเวอร์
    const server = app.listen(PORT, () => {
      console.log('🎉 Server Started Successfully!');
      console.log(`🌐 Server running on port ${PORT}`);
      console.log(`📱 Public URL: ${CONFIG.RENDER.SERVICE_URL}`);
      console.log(`⚙️ Admin Panel: ${CONFIG.RENDER.SERVICE_URL}/admin`);
      console.log(`🕐 Timezone: ${CONFIG.TIMEZONE}`);
      console.log(`🔧 Keep-Alive: ${CONFIG.RENDER.KEEP_ALIVE_ENABLED ? 'Enabled' : 'Disabled'}`);
      console.log('─'.repeat(50));
    });

    // Graceful shutdown
    process.on('SIGTERM', () => {
      console.log('🛑 Received SIGTERM, shutting down gracefully...');
      server.close(() => {
        console.log('✅ Server closed');
        process.exit(0);
      });
    });

    process.on('SIGINT', () => {
      console.log('🛑 Received SIGINT, shutting down gracefully...');
      server.close(() => {
        console.log('✅ Server closed');
        process.exit(0);
      });
    });

  } catch (error) {
    console.error('❌ Failed to start server:', error);
    process.exit(1);
  }
}

// เรียกใช้ฟังก์ชัน startServer
startServer();