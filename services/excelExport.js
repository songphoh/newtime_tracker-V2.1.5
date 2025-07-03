// services/excelExport.js - Excel Export Service
const ExcelJS = require('exceljs');
const moment = require('moment-timezone');
const { CONFIG } = require('../config');

class ExcelExportService {
  static async createWorkbook(data, type, params) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('รายงานการลงเวลา');

    // ตั้งค่าข้อมูลองค์กร
    const orgInfo = {
      name: 'องค์การบริหารส่วนตำบลข่าใหญ่',
      address: 'อำเภอเมือง จังหวัดหนองบัวลำภู',
      phone: '042-315962'
    };

    // สร้างหัวข้อรายงาน
    let reportTitle = '';
    let reportPeriod = '';

    switch (type) {
      case 'daily':
        reportTitle = 'รายงานการลงเวลาเข้า-ออกงาน รายวัน';
        reportPeriod = `วันที่ ${moment(params.date).tz(CONFIG.TIMEZONE).format('DD MMMM YYYY')}`;
        break;
      case 'monthly':
        const monthNames = [
          'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
          'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
        ];
        const isDetailed = params.format === 'detailed';
        reportTitle = isDetailed 
          ? 'รายงานการลงเวลาเข้า-ออกงาน รายเดือน (แบ่งตามวันชัดเจน)'
          : 'รายงานการลงเวลาเข้า-ออกงาน รายเดือน';
        reportPeriod = `เดือน ${monthNames[params.month - 1]} ${parseInt(params.year) + 543}`;
        break;
      case 'range':
        reportTitle = 'รายงานการลงเวลาเข้า-ออกงาน ช่วงวันที่';
        const startDate = moment(params.startDate).tz(CONFIG.TIMEZONE);
        const endDate = moment(params.endDate).tz(CONFIG.TIMEZONE);
        reportPeriod = `${startDate.format('DD MMMM YYYY')} - ${endDate.format('DD MMMM YYYY')}`;
        break;
    }

    // จัดรูปแบบหัวกระดาษ
    worksheet.mergeCells('A1:J3');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `${orgInfo.name}\n${reportTitle}\n${reportPeriod}`;
    titleCell.font = { name: 'Angsana New', size: 18, bold: true };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };

    // ข้อมูลองค์กร
    worksheet.getCell('A4').value = `${orgInfo.address} โทร. ${orgInfo.phone}`;
    worksheet.getCell('A4').font = { name: 'Angsana New', size: 14 };
    worksheet.getCell('A4').alignment = { horizontal: 'center' };
    worksheet.mergeCells('A4:J4');

    // สร้างหัวตาราง
    const headerRow = 6;
    const headers = [
      'ลำดับ',
      'ชื่อ-นามสกุล',
      'วันที่',
      'เวลาเข้า',
      'เวลาออก',
      'ชั่วโมงทำงาน',
      'หมายเหตุ',
      'สถานที่เข้า',
      'สถานที่ออก',
      'ชื่อไลน์'
    ];

    headers.forEach((header, index) => {
      const cell = worksheet.getCell(headerRow, index + 1);
      cell.value = header;
      cell.font = { name: 'Angsana New', size: 14, bold: true };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6FA' }
      };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    // เพิ่มข้อมูล
    if (type === 'monthly' && params.format === 'detailed') {
      // สำหรับรายงานรายเดือนแบบ detailed: จัดเรียงข้อมูลตามวันที่
      data = ExcelExportService.organizeDetailedMonthlyData(data, params);
    }
    
    data.forEach((record, index) => {
      const rowNumber = headerRow + 1 + index;
      
      // จัดการวันที่และเวลา
      let clockInDate = null;
      let clockOutDate = null;
      let dateDisplay = '';
      let clockInTime = '';
      let clockOutTime = '';

      if (record.clockIn) {
        try {
          if (typeof record.clockIn === 'string' && record.clockIn.includes(' ')) {
            // ตรวจสอบรูปแบบ DD/MM/YYYY HH:mm:ss ก่อน
            if (record.clockIn.match(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/)) {
              clockInDate = moment.tz(record.clockIn, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed DD/MM/YYYY format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            // รูปแบบ YYYY-MM-DD HH:mm:ss
            else if (record.clockIn.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
              clockInDate = moment.tz(record.clockIn, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed YYYY-MM-DD format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            else {
              // ลองให้ moment แปลงเอง
              clockInDate = moment(record.clockIn).tz(CONFIG.TIMEZONE);
              console.log(`📅 Auto-parsed format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
          } else {
            clockInDate = moment(record.clockIn).tz(CONFIG.TIMEZONE);
            console.log(`📅 Parsed non-string format: ${record.clockIn} -> ${clockInDate.format('YYYY-MM-DD HH:mm:ss')}`);
          }
          
          if (clockInDate.isValid()) {
            dateDisplay = clockInDate.format('DD/MM/YYYY');
            clockInTime = clockInDate.format('HH:mm:ss');
            console.log(`✅ Final display: Date="${dateDisplay}", Time="${clockInTime}"`);
          } else {
            console.warn(`⚠️ Invalid clockIn date: "${record.clockIn}"`);
          }
        } catch (error) {
          console.warn('Error parsing clockIn time:', record.clockIn, error);
        }
      }

      if (record.clockOut) {
        try {
          if (typeof record.clockOut === 'string' && record.clockOut.includes(' ')) {
            // ตรวจสอบรูปแบบ DD/MM/YYYY HH:mm:ss ก่อน
            if (record.clockOut.match(/^\d{2}\/\d{2}\/\d{4} \d{2}:\d{2}:\d{2}$/)) {
              clockOutDate = moment.tz(record.clockOut, 'DD/MM/YYYY HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed clockOut DD/MM/YYYY format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            // รูปแบบ YYYY-MM-DD HH:mm:ss
            else if (record.clockOut.match(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/)) {
              clockOutDate = moment.tz(record.clockOut, 'YYYY-MM-DD HH:mm:ss', CONFIG.TIMEZONE);
              console.log(`📅 Parsed clockOut YYYY-MM-DD format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
            else {
              // ลองให้ moment แปลงเอง
              clockOutDate = moment(record.clockOut).tz(CONFIG.TIMEZONE);
              console.log(`📅 Auto-parsed clockOut format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
            }
          } else {
            clockOutDate = moment(record.clockOut).tz(CONFIG.TIMEZONE);
            console.log(`📅 Parsed clockOut non-string format: ${record.clockOut} -> ${clockOutDate.format('YYYY-MM-DD HH:mm:ss')}`);
          }
          
          if (clockOutDate.isValid()) {
            clockOutTime = clockOutDate.format('HH:mm:ss');
            console.log(`✅ Final clockOut time: "${clockOutTime}"`);
          } else {
            console.warn(`⚠️ Invalid clockOut date: "${record.clockOut}"`);
          }
        } catch (error) {
          console.warn('Error parsing clockOut time:', record.clockOut, error);
        }
      }

      // จัดการชั่วโมงทำงาน
      let workingHoursDisplay = '';
      if (record.workingHours) {
        const hours = parseFloat(record.workingHours);
        if (!isNaN(hours)) {
          workingHoursDisplay = `${hours.toFixed(2)} ชม.`;
        } else {
          workingHoursDisplay = record.workingHours;
        }
      }

      const rowData = [
        record.no || (index + 1),
        record.employee || '',
        dateDisplay,
        clockInTime,
        clockOutTime,
        workingHoursDisplay,
        record.note || '',
        record.locationIn || '',
        record.locationOut || '',
        record.lineName || ''
      ];

      rowData.forEach((value, colIndex) => {
        const cell = worksheet.getCell(rowNumber, colIndex + 1);
        cell.value = value;
        cell.font = { name: 'Angsana New', size: 12 };
        cell.alignment = { 
          horizontal: colIndex === 0 ? 'center' : 'left', 
          vertical: 'middle' 
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };

        // สีพื้นหลังสำหรับแถวต่างๆ (ตรวจสอบหมายเหตุจากคอลัมน์ E)
        if (record.note && record.note.includes('ลืมลงเวลาออก')) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFCCCC' } // สีแดงอ่อน
          };
        }
      });
    });

    // ปรับขนาดคอลัมน์
    const columnWidths = [8, 25, 15, 12, 12, 15, 25, 30, 30, 20];
    columnWidths.forEach((width, index) => {
      worksheet.getColumn(index + 1).width = width;
    });

    // สรุปข้อมูล
    const summaryRow = headerRow + data.length + 2;
    
    // สถิติการทำงาน
    const totalRecords = data.length;
    const normalCheckouts = data.filter(r => !r.note || !r.note.includes('ลืมลงเวลาออก')).length;
    const missedCheckouts = data.filter(r => r.note && r.note.includes('ลืมลงเวลาออก')).length;
    
    worksheet.getCell(summaryRow, 1).value = `สรุปข้อมูล: ทั้งหมด ${totalRecords} รายการ | ลงเวลาออกปกติ ${normalCheckouts} คน | ลืมลงเวลาออก ${missedCheckouts} คน`;
    worksheet.getCell(summaryRow, 1).font = { name: 'Angsana New', size: 12, bold: true };
    worksheet.mergeCells(`A${summaryRow}:J${summaryRow}`);

    // วันที่สร้างรายงาน
    const footerRow = summaryRow + 2;
    const currentTime = moment().tz(CONFIG.TIMEZONE);
    worksheet.getCell(footerRow, 1).value = `สร้างรายงานเมื่อ: ${currentTime.format('DD/MM/YYYY HH:mm:ss')} (เวลาไทย)`;
    worksheet.getCell(footerRow, 1).font = { name: 'Angsana New', size: 10 };
    worksheet.getCell(footerRow, 1).alignment = { horizontal: 'right' };
    worksheet.mergeCells(`A${footerRow}:J${footerRow}`);

    // เพิ่มหมายเหตุเกี่ยวกับสี
    if (data.some(r => r.note && r.note.includes('ลืมลงเวลาออก'))) {
      const noteRow = footerRow + 1;
      worksheet.getCell(noteRow, 1).value = 'หมายเหตุ: แถวที่มีพื้นหลังสีแดงอ่อน = ลืมลงเวลาออก (ระบบอัตโนมัติ)';
      worksheet.getCell(noteRow, 1).font = { name: 'Angsana New', size: 10, italic: true };
      worksheet.mergeCells(`A${noteRow}:J${noteRow}`);
    }

    return workbook;
  }

  // ฟังก์ชันสำหรับจัดเรียงข้อมูลรายเดือนแบบ detailed
  static organizeDetailedMonthlyData(data, params) {
    console.log(`📊 Organizing detailed monthly data: ${data.length} records`);
    
    // จัดเรียงข้อมูลตามวันที่ และ ชื่อพนักงาน
    const sortedData = data.sort((a, b) => {
      // เรียงตามวันที่ก่อน
      const dateA = moment(a.clockIn).tz(CONFIG.TIMEZONE);
      const dateB = moment(b.clockIn).tz(CONFIG.TIMEZONE);
      
      if (dateA.format('YYYY-MM-DD') !== dateB.format('YYYY-MM-DD')) {
        return dateA.diff(dateB);
      }
      
      // ถ้าวันที่เดียวกัน เรียงตามชื่อพนักงาน
      return (a.employee || '').localeCompare(b.employee || '', 'th');
    });
    
    console.log(`✅ Sorted detailed data: ${sortedData.length} records`);
    return sortedData;
  }
}

module.exports = ExcelExportService;
