const ExcelJS = require('exceljs');

/**
 * Builds an Excel workbook for the Score List template with 100% fidelity to the original document.
 */
async function generateScoreListWorkbook(data) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ImageToExcel - AI Converter';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Score List', {
    views: [{ showGridLines: true }],
    pageSetup: {
      orientation: 'landscape',
      paperSize: 9, // A4
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0
    }
  });

  const khmerFontFamily = 'Khmer OS Siemreap';
  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };

  const headerFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE5E7EB' }
  };

  // Define Columns
  sheet.columns = [
    { key: 'no', width: 6 },
    { key: 'name_en', width: 25 },
    { key: 'name_kh', width: 22 },
    { key: 'sex', width: 7 },
    { key: 'group', width: 8 },
    { key: 'conduct', width: 9 },
    { key: 'cp', width: 7 },
    { key: 'hw1', width: 6 },
    { key: 'hw2', width: 6 },
    { key: 'hw3', width: 6 },
    { key: 'hw4', width: 6 },
    { key: 'hw5', width: 6 },
    { key: 'hw6', width: 6 },
    { key: 'qz1', width: 6 },
    { key: 'qz2', width: 6 },
    { key: 'qz3', width: 6 },
    { key: 'qz4', width: 6 },
    { key: 'qz5', width: 6 },
    { key: 'qz6', width: 6 },
    { key: 'mt1', width: 7 },
    { key: 'mt2', width: 7 },
    { key: 'mt3', width: 7 }
  ];

  // Header rows
  sheet.mergeCells('A1:D1');
  const a1 = sheet.getCell('A1');
  a1.value = data.schoolNameKh || 'សាលាអន្តរជាតិ វេស្ទើន';
  a1.font = { name: khmerFontFamily, size: 11, bold: true, color: { argb: 'FF1E3A8A' } };
  a1.alignment = { vertical: 'middle', horizontal: 'left' };

  sheet.mergeCells('A2:D2');
  const a2 = sheet.getCell('A2');
  a2.value = data.officeName || "Registrar's Office SSK";
  a2.font = { name: 'Calibri', size: 9.5, bold: true, color: { argb: 'FF374151' } };
  a2.alignment = { vertical: 'middle', horizontal: 'left' };

  // Center Banner
  sheet.mergeCells('H1:P2');
  const titleCell = sheet.getCell('H1');
  titleCell.value = data.title || 'តារាងសម្រង់ពិន្ទុ / Score List';
  titleCell.font = { name: khmerFontFamily, size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1F2937' }
  };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  // Top Right Info
  sheet.mergeCells('Q1:R1');
  sheet.getCell('Q1').value = 'Teacher Name:';
  sheet.getCell('Q1').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('Q1').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.mergeCells('S1:V1');
  const teacherCell = sheet.getCell('S1');
  teacherCell.value = data.teacherName || '';
  teacherCell.font = { name: 'Calibri', size: 10, bold: true };
  teacherCell.alignment = { vertical: 'middle', horizontal: 'left' };
  ['S1', 'T1', 'U1', 'V1'].forEach(cell => sheet.getCell(cell).border = { bottom: { style: 'thin' } });

  sheet.mergeCells('Q2:R2');
  sheet.getCell('Q2').value = 'Subject:';
  sheet.getCell('Q2').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('Q2').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.mergeCells('S2:V2');
  const subjCell = sheet.getCell('S2');
  subjCell.value = data.subject || '';
  subjCell.font = { name: 'Calibri', size: 10, bold: true };
  subjCell.alignment = { vertical: 'middle', horizontal: 'left' };
  ['S2', 'T2', 'U2', 'V2'].forEach(cell => sheet.getCell(cell).border = { bottom: { style: 'thin' } });

  sheet.mergeCells('Q3:R3');
  sheet.getCell('Q3').value = 'Grade:';
  sheet.getCell('Q3').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('Q3').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.getCell('S3').value = data.grade || '';
  sheet.getCell('S3').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('S3').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('S3').border = { bottom: { style: 'thin' } };

  sheet.getCell('T3').value = 'For :';
  sheet.getCell('T3').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('T3').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.mergeCells('U3:V3');
  sheet.getCell('U3').value = data.quarter || '';
  sheet.getCell('U3').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('U3').alignment = { vertical: 'middle', horizontal: 'left' };
  ['U3', 'V3'].forEach(cell => sheet.getCell(cell).border = { bottom: { style: 'thin' } });

  sheet.getRow(1).height = 20;
  sheet.getRow(2).height = 20;
  sheet.getRow(3).height = 20;
  sheet.getRow(4).height = 10;

  // Table Headers
  sheet.getRow(5).height = 24;
  sheet.getRow(6).height = 20;

  sheet.mergeCells('A5:A6');
  sheet.getCell('A5').value = 'N°';
  sheet.getCell('A5').alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.mergeCells('B5:B6');
  sheet.getCell('B5').value = 'Name';
  sheet.getCell('B5').alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.mergeCells('C5:C6');
  sheet.getCell('C5').value = 'ឈ្មោះ';
  sheet.getCell('C5').alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.getCell('D5').value = 'Sex';
  sheet.getCell('D5').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('D6').value = 'ភេទ';
  sheet.getCell('D6').alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.getCell('E5').value = 'Group';
  sheet.getCell('E5').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('E6').value = 'ក្រុម';
  sheet.getCell('E6').alignment = { vertical: 'middle', horizontal: 'center' };

  sheet.mergeCells('F5:F6');
  sheet.getCell('F5').value = 'Conduct';
  sheet.getCell('F5').alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90 };

  sheet.mergeCells('G5:G6');
  sheet.getCell('G5').value = 'C.P.';
  sheet.getCell('G5').alignment = { vertical: 'middle', horizontal: 'center' };

  // Homework
  sheet.mergeCells('H5:M5');
  sheet.getCell('H5').value = 'កិច្ចការផ្ទះ / Homework';
  sheet.getCell('H5').alignment = { vertical: 'middle', horizontal: 'center' };
  for (let i = 1; i <= 6; i++) {
    const c = sheet.getCell(6, 7 + i);
    c.value = i;
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  // Quizzes
  sheet.mergeCells('N5:S5');
  sheet.getCell('N5').value = 'តេស្តខ្លី / Quizzes';
  sheet.getCell('N5').alignment = { vertical: 'middle', horizontal: 'center' };
  for (let i = 1; i <= 6; i++) {
    const c = sheet.getCell(6, 13 + i);
    c.value = i;
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  // Monthly Test
  sheet.mergeCells('T5:V5');
  sheet.getCell('T5').value = 'Monthly Test\nតេស្តប្រចាំខែ';
  sheet.getCell('T5').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  for (let i = 1; i <= 3; i++) {
    const c = sheet.getCell(6, 19 + i);
    c.value = i;
    c.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  for (let r = 5; r <= 6; r++) {
    for (let c = 1; c <= 22; c++) {
      const cell = sheet.getCell(r, c);
      cell.fill = headerFill;
      cell.border = thinBorder;
      cell.font = { name: khmerFontFamily, size: 9.5, bold: true };
    }
  }

  // Data rows
  const rows = data.rows || [];
  let startRow = 7;
  rows.forEach((item, index) => {
    const rowIdx = startRow + index;
    const row = sheet.getRow(rowIdx);
    row.height = 22;

    row.getCell(1).value = item.no !== undefined ? item.no : index + 1;
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(1).font = { name: 'Calibri', size: 10 };

    row.getCell(2).value = item.name_en || '';
    row.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
    row.getCell(2).font = { name: 'Calibri', size: 9.5, bold: true };

    row.getCell(3).value = item.name_kh || '';
    row.getCell(3).alignment = { vertical: 'middle', horizontal: 'left' };
    row.getCell(3).font = { name: khmerFontFamily, size: 10 };

    row.getCell(4).value = item.sex || '';
    row.getCell(4).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(4).font = { name: 'Calibri', size: 10 };

    row.getCell(5).value = item.group || '';
    row.getCell(5).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(5).font = { name: 'Calibri', size: 10 };

    row.getCell(6).value = item.conduct !== undefined ? item.conduct : '';
    row.getCell(6).alignment = { vertical: 'middle', horizontal: 'center' };

    row.getCell(7).value = item.cp !== undefined ? item.cp : '';
    row.getCell(7).alignment = { vertical: 'middle', horizontal: 'center' };

    const hw = Array.isArray(item.hw) ? item.hw : [];
    for (let i = 0; i < 6; i++) {
      const c = row.getCell(8 + i);
      const val = hw[i];
      c.value = val !== '' && val !== null && val !== undefined && !isNaN(val) ? Number(val) : (val || '');
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.font = { name: 'Calibri', size: 10, color: { argb: 'FFB91C1C' } };
    }

    const qz = Array.isArray(item.qz) ? item.qz : [];
    for (let i = 0; i < 6; i++) {
      const c = row.getCell(14 + i);
      const val = qz[i];
      c.value = val !== '' && val !== null && val !== undefined && !isNaN(val) ? Number(val) : (val || '');
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.font = { name: 'Calibri', size: 10, color: { argb: 'FFB91C1C' } };
    }

    const mt = Array.isArray(item.mt) ? item.mt : [];
    for (let i = 0; i < 3; i++) {
      const c = row.getCell(20 + i);
      const val = mt[i];
      c.value = val !== '' && val !== null && val !== undefined && !isNaN(val) ? Number(val) : (val || '');
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.font = { name: 'Calibri', size: 10, color: { argb: 'FFB91C1C' } };
    }

    for (let c = 1; c <= 22; c++) {
      row.getCell(c).border = thinBorder;
    }
  });

  // Optional blank rows
  const lastDataRow = startRow + rows.length - 1;
  const blankCount = Math.max(0, 2);
  for (let b = 1; b <= blankCount; b++) {
    const blankRowIdx = lastDataRow + b;
    const bRow = sheet.getRow(blankRowIdx);
    bRow.height = 20;
    for (let c = 1; c <= 22; c++) {
      bRow.getCell(c).border = thinBorder;
    }
  }

  // Footer notes
  const footerStart = lastDataRow + blankCount + 2;

  sheet.mergeCells(`A${footerStart}:Q${footerStart}`);
  const fKhmer1 = sheet.getCell(`A${footerStart}`);
  fKhmer1.value = data.notesKh1 || 'ចំណាំ ៖ តារាងសម្រង់ពិន្ទុនេះគឺសម្រាប់លោកគ្រូអ្នកគ្រូកត់ពិន្ទុចូល ដើម្បីងាយស្រួលក្នុងការបញ្ចូលពិន្ទុក្នុងតារាងសម្រង់ពិន្ទុអេឡិចត្រូនិក (E-Gradebook) ។';
  fKhmer1.font = { name: khmerFontFamily, size: 8.5 };
  fKhmer1.alignment = { vertical: 'middle', horizontal: 'left' };

  sheet.mergeCells(`R${footerStart}:V${footerStart}`);
  const dateCell = sheet.getCell(`R${footerStart}`);
  dateCell.value = data.date ? `Date:  ${data.date}` : 'Date:  21-Sep-26';
  dateCell.font = { name: 'Calibri', size: 9, bold: true };
  dateCell.alignment = { vertical: 'middle', horizontal: 'right' };

  const footerLine2 = footerStart + 1;
  sheet.mergeCells(`A${footerLine2}:Q${footerLine2}`);
  const fKhmer2 = sheet.getCell(`A${footerLine2}`);
  fKhmer2.value = data.notesKh2 || '             តារាងសម្រង់ពិន្ទុនេះគឺប្រើប្រាស់សម្រាប់ធ្វើការកត់ត្រាផ្ទាល់ខ្លួនរបស់លោកគ្រូអ្នកគ្រូតែប៉ុណ្ណោះ។';
  fKhmer2.font = { name: khmerFontFamily, size: 8.5 };
  fKhmer2.alignment = { vertical: 'middle', horizontal: 'left' };

  const footerLine3 = footerLine2 + 1;
  sheet.mergeCells(`A${footerLine3}:V${footerLine3}`);
  const fEng1 = sheet.getCell(`A${footerLine3}`);
  fEng1.value = data.notesEn1 || '* NOTE: This list is for teachers to keep record of all kinds of scores which is served as hard copies for egrade-book input.';
  fEng1.font = { name: 'Calibri', size: 8.5 };
  fEng1.alignment = { vertical: 'middle', horizontal: 'left' };

  const footerLine4 = footerLine3 + 1;
  sheet.mergeCells(`A${footerLine4}:V${footerLine4}`);
  const fEng2 = sheet.getCell(`A${footerLine4}`);
  fEng2.value = data.notesEn2 || '             The list is for teacher personal use. It is not required by the office.';
  fEng2.font = { name: 'Calibri', size: 8.5 };
  fEng2.alignment = { vertical: 'middle', horizontal: 'left' };

  return workbook;
}

/**
 * Builds a generic Excel table for arbitrary extracted tables.
 */
async function generateGenericTableWorkbook(tableData) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ImageToExcel - AI Converter';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Sheet1', {
    views: [{ showGridLines: true }]
  });

  const khmerFontFamily = 'Khmer OS Siemreap';
  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };

  const headerFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE5E7EB' }
  };

  const headers = tableData.headers || [];
  const rows = tableData.rows || [];

  // Set headers
  const headerRow = sheet.getRow(1);
  headerRow.height = 24;
  headers.forEach((h, colIdx) => {
    const cell = headerRow.getCell(colIdx + 1);
    cell.value = h;
    cell.font = { name: khmerFontFamily, size: 10, bold: true };
    cell.fill = headerFill;
    cell.border = thinBorder;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  // Set rows
  rows.forEach((rowVals, rIdx) => {
    const row = sheet.getRow(rIdx + 2);
    row.height = 22;
    rowVals.forEach((val, cIdx) => {
      const cell = row.getCell(cIdx + 1);
      const isNum = val !== '' && !isNaN(val);
      cell.value = isNum ? Number(val) : val;
      cell.font = { name: khmerFontFamily, size: 10 };
      cell.border = thinBorder;
      cell.alignment = {
        vertical: 'middle',
        horizontal: isNum ? 'center' : 'left'
      };
    });
  });

  // Adjust column widths automatically
  for (let c = 1; c <= headers.length; c++) {
    let maxLen = (headers[c - 1] || '').toString().length;
    rows.forEach(r => {
      const len = (r[c - 1] || '').toString().length;
      if (len > maxLen) maxLen = len;
    });
    sheet.getColumn(c).width = Math.max(10, Math.min(maxLen + 4, 40));
  }

  return workbook;
}

module.exports = {
  generateScoreListWorkbook,
  generateGenericTableWorkbook
};

