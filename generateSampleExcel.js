const ExcelJS = require('exceljs');
const path = require('path');

async function createScoreListExcel(outputPath) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'ImageToExcel Converter';
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

  // Fonts
  const khmerFontFamily = 'Khmer OS Siemreap';
  const defaultFont = { name: 'Calibri', size: 10 };
  const khmerFont = { name: khmerFontFamily, size: 10 };
  const boldFont = { name: 'Calibri', size: 10, bold: true };
  const headerKhmerFont = { name: khmerFontFamily, size: 10, bold: true };

  // Borders
  const thinBorder = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    left: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'thin', color: { argb: 'FF000000' } },
    right: { style: 'thin', color: { argb: 'FF000000' } }
  };

  const headerFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFE5E7EB' } // Clean gray #E5E7EB
  };

  // Define Columns (21 columns)
  // A: N°
  // B: Name
  // C: ឈ្មោះ
  // D: Sex
  // E: Group
  // F: Conduct
  // G: C.P.
  // H..M: Homework 1..6
  // N..S: Quizzes 1..6
  // T..V: Monthly Test 1..3
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

  // ==================== TOP HEADER SECTION ====================
  // Row 1: School Name (left), Title Box (middle), Teacher Name (right)
  sheet.mergeCells('A1:C1');
  const a1 = sheet.getCell('A1');
  a1.value = 'សាលាអន្តរជាតិ វេស្ទើន';
  a1.font = { name: khmerFontFamily, size: 11, bold: true, color: { argb: 'FF1E3A8A' } };
  a1.alignment = { vertical: 'middle', horizontal: 'left' };

  sheet.mergeCells('A2:C2');
  const a2 = sheet.getCell('A2');
  a2.value = "Registrar's Office SSK";
  a2.font = { name: 'Calibri', size: 9, bold: true, color: { argb: 'FF374151' } };
  a2.alignment = { vertical: 'middle', horizontal: 'left' };

  // Middle Title Banner (Col H to P)
  sheet.mergeCells('H1:P2');
  const titleCell = sheet.getCell('H1');
  titleCell.value = 'តារាងសម្រង់ពិន្ទុ / Score List';
  titleCell.font = { name: khmerFontFamily, size: 13, bold: true, color: { argb: 'FFFFFFFF' } };
  titleCell.fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF1F2937' } // Dark charcoal background
  };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

  // Top Right Info
  sheet.mergeCells('Q1:R1');
  sheet.getCell('Q1').value = 'Teacher Name:';
  sheet.getCell('Q1').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('Q1').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.mergeCells('S1:V1');
  const teacherCell = sheet.getCell('S1');
  teacherCell.value = 'Bin Soratha';
  teacherCell.font = { name: 'Calibri', size: 10, bold: true };
  teacherCell.alignment = { vertical: 'middle', horizontal: 'left' };
  sheet.getCell('S1').border = { bottom: { style: 'thin' } };
  sheet.getCell('T1').border = { bottom: { style: 'thin' } };
  sheet.getCell('U1').border = { bottom: { style: 'thin' } };
  sheet.getCell('V1').border = { bottom: { style: 'thin' } };

  sheet.mergeCells('Q2:R2');
  sheet.getCell('Q2').value = 'Subject:';
  sheet.getCell('Q2').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('Q2').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.mergeCells('S2:V2');
  const subjCell = sheet.getCell('S2');
  subjCell.value = 'Math';
  subjCell.font = { name: 'Calibri', size: 10, bold: true };
  subjCell.alignment = { vertical: 'middle', horizontal: 'left' };
  sheet.getCell('S2').border = { bottom: { style: 'thin' } };
  sheet.getCell('T2').border = { bottom: { style: 'thin' } };
  sheet.getCell('U2').border = { bottom: { style: 'thin' } };
  sheet.getCell('V2').border = { bottom: { style: 'thin' } };

  // Row 3: Grade and Quarter
  sheet.mergeCells('Q3:R3');
  sheet.getCell('Q3').value = 'Grade:';
  sheet.getCell('Q3').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('Q3').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.getCell('S3').value = '1A';
  sheet.getCell('S3').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('S3').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('S3').border = { bottom: { style: 'thin' } };

  sheet.getCell('T3').value = 'For :';
  sheet.getCell('T3').font = { name: 'Calibri', size: 9, bold: true };
  sheet.getCell('T3').alignment = { vertical: 'middle', horizontal: 'right' };

  sheet.mergeCells('U3:V3');
  sheet.getCell('U3').value = 'Quarter 1';
  sheet.getCell('U3').font = { name: 'Calibri', size: 10, bold: true };
  sheet.getCell('U3').alignment = { vertical: 'middle', horizontal: 'left' };
  sheet.getCell('U3').border = { bottom: { style: 'thin' } };
  sheet.getCell('V3').border = { bottom: { style: 'thin' } };

  // Set Row heights
  sheet.getRow(1).height = 20;
  sheet.getRow(2).height = 20;
  sheet.getRow(3).height = 20;
  sheet.getRow(4).height = 10; // spacer

  // ==================== TABLE HEADERS (Rows 5 & 6) ====================
  sheet.getRow(5).height = 24;
  sheet.getRow(6).height = 20;

  // Merged Vertically: A5:A6 (N°)
  sheet.mergeCells('A5:A6');
  const hA = sheet.getCell('A5');
  hA.value = 'N°';
  hA.alignment = { vertical: 'middle', horizontal: 'center' };

  // B5:B6 (Name)
  sheet.mergeCells('B5:B6');
  const hB = sheet.getCell('B5');
  hB.value = 'Name';
  hB.alignment = { vertical: 'middle', horizontal: 'center' };

  // C5:C6 (ឈ្មោះ)
  sheet.mergeCells('C5:C6');
  const hC = sheet.getCell('C5');
  hC.value = 'ឈ្មោះ';
  hC.alignment = { vertical: 'middle', horizontal: 'center' };

  // D5: Sex / D6: ភេទ
  sheet.getCell('D5').value = 'Sex';
  sheet.getCell('D5').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('D6').value = 'ភេទ';
  sheet.getCell('D6').alignment = { vertical: 'middle', horizontal: 'center' };

  // E5: Group / E6: ក្រុម
  sheet.getCell('E5').value = 'Group';
  sheet.getCell('E5').alignment = { vertical: 'middle', horizontal: 'center' };
  sheet.getCell('E6').value = 'ក្រុម';
  sheet.getCell('E6').alignment = { vertical: 'middle', horizontal: 'center' };

  // F5:F6 (Conduct)
  sheet.mergeCells('F5:F6');
  const hF = sheet.getCell('F5');
  hF.value = 'Conduct';
  hF.alignment = { vertical: 'middle', horizontal: 'center', textRotation: 90 };

  // G5:G6 (C.P.)
  sheet.mergeCells('G5:G6');
  const hG = sheet.getCell('G5');
  hG.value = 'C.P.';
  hG.alignment = { vertical: 'middle', horizontal: 'center' };

  // Homework (H5:M5)
  sheet.mergeCells('H5:M5');
  const hH = sheet.getCell('H5');
  hH.value = 'កិច្ចការផ្ទះ / Homework';
  hH.alignment = { vertical: 'middle', horizontal: 'center' };
  for (let i = 1; i <= 6; i++) {
    const cell = sheet.getCell(6, 7 + i); // cols 8..13 -> H..M
    cell.value = i;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  // Quizzes (N5:S5)
  sheet.mergeCells('N5:S5');
  const hN = sheet.getCell('N5');
  hN.value = 'តេស្តខ្លី / Quizzes';
  hN.alignment = { vertical: 'middle', horizontal: 'center' };
  for (let i = 1; i <= 6; i++) {
    const cell = sheet.getCell(6, 13 + i); // cols 14..19 -> N..S
    cell.value = i;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  // Monthly Test (T5:V5)
  sheet.mergeCells('T5:V5');
  const hT = sheet.getCell('T5');
  hT.value = 'Monthly Test\nតេស្តប្រចាំខែ';
  hT.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  for (let i = 1; i <= 3; i++) {
    const cell = sheet.getCell(6, 19 + i); // cols 20..22 -> T..V
    cell.value = i;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  }

  // Apply header styles and borders to rows 5 & 6
  for (let r = 5; r <= 6; r++) {
    for (let c = 1; c <= 22; c++) {
      const cell = sheet.getCell(r, c);
      cell.fill = headerFill;
      cell.border = thinBorder;
      cell.font = { name: khmerFontFamily, size: 9.5, bold: true };
    }
  }

  // ==================== DATA ROWS ====================
  const data = [
    { no: 1, name_en: 'CHHEANGCHAN ROTH', name_kh: 'ឈាងច័ន្ទ រ័ត្នមុនី', sex: 'M', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 91, 95, '', '', ''], mt: [95, 94, ''] },
    { no: 2, name_en: 'HENG MANANIKA', name_kh: 'ហេង ម៉ាណានីកា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 50, 100, 100, '', ''], qz: [100, 100, 80, '', '', ''], mt: [89, 57, ''] },
    { no: 3, name_en: 'HENG ZANA', name_kh: 'ហេង ហ្សាណា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [92, 100, 100, '', '', ''], mt: [100, 98, ''] },
    { no: 4, name_en: 'HOU TIPANHALOKA', name_kh: 'ហ៊ូ តិបុញ្ញាលោកា', sex: 'M', group: 'III', conduct: '', cp: '', hw: [100, 50, 90, 98, '', ''], qz: [100, 82, 100, '', '', ''], mt: [99, 89, ''] },
    { no: 5, name_en: 'KHAI ELLIS', name_kh: 'ខៃ អេលលីស', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 100, 90, '', '', ''], mt: [100, 90, ''] },
    { no: 6, name_en: 'KRY LIXIN', name_kh: 'គ្រី លីស៊ីន', sex: 'M', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 100, 100, '', '', ''], mt: [100, 100, ''] },
    { no: 7, name_en: 'MEANG JULLY', name_kh: 'ម៉ាង ជូលី', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [92, 90, 100, '', '', ''], mt: [86, 75, ''] },
    { no: 8, name_en: 'NARONG KEMARAK', name_kh: 'ណារ៉ុង ខេមរៈ', sex: 'M', group: 'III', conduct: '', cp: '', hw: [100, 90, 100, 70, '', ''], qz: [55, 73, 90, '', '', ''], mt: [35, 71, ''] },
    { no: 9, name_en: 'NIRABEL VYRIYA', name_kh: 'នីរ៉ាបែល វីរិយា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [55, 82, 57, '', '', ''], mt: [65, 37, ''] },
    { no: 10, name_en: 'OUCH ROZANA', name_kh: 'អ៊ូច រ៉ូហ្សាណា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 60, 100, 98, '', ''], qz: [96, 87, 100, '', '', ''], mt: [100, 98, ''] },
    { no: 11, name_en: 'PEOU AHMIDHA', name_kh: 'ពៅ អាមីដា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 80, '', ''], qz: [100, 100, 65, '', '', ''], mt: [93, 83, ''] },
    { no: 12, name_en: 'PEOV KITIYA', name_kh: 'ពៅ គិតិយា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 100, 100, '', '', ''], mt: [95, 98, ''] },
    { no: 13, name_en: 'PHAL DENVICHEY', name_kh: 'ផល ដេនវិជ័យ', sex: 'M', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 90, '', ''], qz: [100, 100, 99, '', '', ''], mt: [85, 96, ''] },
    { no: 14, name_en: 'PHAT PHEAKTRA', name_kh: 'ផាត់ ភក្ត្រា', sex: 'M', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 90, 94, '', '', ''], mt: [88, 94, ''] },
    { no: 15, name_en: 'SAMSIM SOPHEAKNEI', name_kh: 'សំស៊ឹម សុភក្តិណី', sex: 'M', group: 'III', conduct: '', cp: '', hw: [100, 90, 100, 98, '', ''], qz: [100, 87, 96, '', '', ''], mt: [88, 79, ''] },
    { no: 16, name_en: 'SOVAN SEREYPHEAN', name_kh: 'សុវណ្ណ សិរីភានុមា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [96, 100, 100, '', '', ''], mt: [78, 66, ''] },
    { no: 17, name_en: 'SUN PHENG AN', name_kh: 'ស៊ុន ផេងអាន', sex: 'M', group: 'III', conduct: '', cp: '', hw: [0, 0, 95, 100, '', ''], qz: [100, 100, 100, '', '', ''], mt: [88, 50, ''] },
    { no: 18, name_en: 'SUON SETHIKA', name_kh: 'សួន សិទ្ធិកា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 80, 90, 90, '', ''], qz: [92, 96, 81, '', '', ''], mt: [87, 56, ''] },
    { no: 19, name_en: 'TAING MUYLY', name_kh: 'តាំង មុយលី', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 90, '', ''], qz: [100, 91, 88, '', '', ''], mt: [100, 96, ''] },
    { no: 20, name_en: 'TE SAMNAMMUNISAT', name_kh: 'តែ សំណាងមុនិសត្យា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 100, 100, '', '', ''], mt: [100, 100, ''] },
    { no: 21, name_en: 'TY Y XING', name_kh: 'ទី អ៊ីស៊ីង', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 100, 100, '', '', ''], mt: [100, 90, ''] },
    { no: 22, name_en: 'VANNAK SOFIYADA', name_kh: 'វណ្ណៈ សូហ្វីយ៉ាដា', sex: 'F', group: 'III', conduct: '', cp: '', hw: [100, 100, 100, 100, '', ''], qz: [100, 100, 100, '', '', ''], mt: [84, 70, ''] }
  ];

  let startRow = 7;
  data.forEach((item, index) => {
    const rowIdx = startRow + index;
    const row = sheet.getRow(rowIdx);
    row.height = 22; // Comfortable height for Khmer script

    row.getCell(1).value = item.no;
    row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(1).font = { name: 'Calibri', size: 10 };

    row.getCell(2).value = item.name_en;
    row.getCell(2).alignment = { vertical: 'middle', horizontal: 'left' };
    row.getCell(2).font = { name: 'Calibri', size: 9.5, bold: true };

    row.getCell(3).value = item.name_kh;
    row.getCell(3).alignment = { vertical: 'middle', horizontal: 'left' };
    row.getCell(3).font = { name: khmerFontFamily, size: 10 };

    row.getCell(4).value = item.sex;
    row.getCell(4).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(4).font = { name: 'Calibri', size: 10 };

    row.getCell(5).value = item.group;
    row.getCell(5).alignment = { vertical: 'middle', horizontal: 'center' };
    row.getCell(5).font = { name: 'Calibri', size: 10 };

    row.getCell(6).value = item.conduct;
    row.getCell(6).alignment = { vertical: 'middle', horizontal: 'center' };

    row.getCell(7).value = item.cp;
    row.getCell(7).alignment = { vertical: 'middle', horizontal: 'center' };

    // Scores
    for (let i = 0; i < 6; i++) {
      const c = row.getCell(8 + i);
      c.value = item.hw[i] !== '' ? Number(item.hw[i]) : '';
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.font = { name: 'Calibri', size: 10, color: { argb: 'FFB91C1C' } }; // Red score styling matching original ink
    }
    for (let i = 0; i < 6; i++) {
      const c = row.getCell(14 + i);
      c.value = item.qz[i] !== '' ? Number(item.qz[i]) : '';
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.font = { name: 'Calibri', size: 10, color: { argb: 'FFB91C1C' } };
    }
    for (let i = 0; i < 3; i++) {
      const c = row.getCell(20 + i);
      c.value = item.mt[i] !== '' ? Number(item.mt[i]) : '';
      c.alignment = { vertical: 'middle', horizontal: 'center' };
      c.font = { name: 'Calibri', size: 10, color: { argb: 'FFB91C1C' } };
    }

    // Apply borders
    for (let c = 1; c <= 22; c++) {
      row.getCell(c).border = thinBorder;
    }
  });

  // Add 3 blank rows (as seen on the template)
  const lastDataRow = startRow + data.length - 1;
  for (let b = 1; b <= 3; b++) {
    const blankRowIdx = lastDataRow + b;
    const bRow = sheet.getRow(blankRowIdx);
    bRow.height = 20;
    for (let c = 1; c <= 22; c++) {
      bRow.getCell(c).border = thinBorder;
    }
  }

  // ==================== FOOTER SECTION ====================
  const footerStart = lastDataRow + 5;

  sheet.mergeCells(`A${footerStart}:Q${footerStart}`);
  const fKhmer1 = sheet.getCell(`A${footerStart}`);
  fKhmer1.value = 'ចំណាំ ៖ តារាងសម្រង់ពិន្ទុនេះគឺសម្រាប់លោកគ្រូអ្នកគ្រូកត់ពិន្ទុចូល ដើម្បីងាយស្រួលក្នុងការបញ្ចូលពិន្ទុក្នុងតារាងសម្រង់ពិន្ទុអេឡិចត្រូនិក (E-Gradebook) ។';
  fKhmer1.font = { name: khmerFontFamily, size: 8.5 };
  fKhmer1.alignment = { vertical: 'middle', horizontal: 'left' };

  // Date on the right
  sheet.mergeCells(`R${footerStart}:V${footerStart}`);
  const dateCell = sheet.getCell(`R${footerStart}`);
  dateCell.value = 'Date:  21-Sep-26';
  dateCell.font = { name: 'Calibri', size: 9, bold: true };
  dateCell.alignment = { vertical: 'middle', horizontal: 'right' };

  const footerLine2 = footerStart + 1;
  sheet.mergeCells(`A${footerLine2}:Q${footerLine2}`);
  const fKhmer2 = sheet.getCell(`A${footerLine2}`);
  fKhmer2.value = '             តារាងសម្រង់ពិន្ទុនេះគឺប្រើប្រាស់សម្រាប់ធ្វើការកត់ត្រាផ្ទាល់ខ្លួនរបស់លោកគ្រូអ្នកគ្រូតែប៉ុណ្ណោះ។';
  fKhmer2.font = { name: khmerFontFamily, size: 8.5 };
  fKhmer2.alignment = { vertical: 'middle', horizontal: 'left' };

  const footerLine3 = footerLine2 + 1;
  sheet.mergeCells(`A${footerLine3}:V${footerLine3}`);
  const fEng1 = sheet.getCell(`A${footerLine3}`);
  fEng1.value = '* NOTE: This list is for teachers to keep record of all kinds of scores which is served as hard copies for egrade-book input.';
  fEng1.font = { name: 'Calibri', size: 8.5 };
  fEng1.alignment = { vertical: 'middle', horizontal: 'left' };

  const footerLine4 = footerLine3 + 1;
  sheet.mergeCells(`A${footerLine4}:V${footerLine4}`);
  const fEng2 = sheet.getCell(`A${footerLine4}`);
  fEng2.value = '             The list is for teacher personal use. It is not required by the office.';
  fEng2.font = { name: 'Calibri', size: 8.5 };
  fEng2.alignment = { vertical: 'middle', horizontal: 'left' };

  await workbook.xlsx.writeFile(outputPath);
  console.log(`Excel generated successfully at: ${outputPath}`);
}

createScoreListExcel(path.join(__dirname, 'Score_List_1A_Quarter1.xlsx'))
  .catch(err => {
    console.error('Error generating Excel:', err);
    process.exit(1);
  });

