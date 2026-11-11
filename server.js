const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const { extractTableFromImages } = require('./geminiExtractor');
const { generateScoreListWorkbook, generateGenericTableWorkbook } = require('./excelGenerator');

const app = express();
const PORT = process.env.PORT || 3000;
const HOST = '0.0.0.0';

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/samples', express.static(path.join(__dirname, 'samples')));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 25 * 1024 * 1024 } // 25MB per file
});

// Sample data preloaded for immediate demonstration & testing
const sampleExtractedData = {
  isScoreList: true,
  schoolNameKh: 'សាលាអន្តរជាតិ វេស្ទើន',
  officeName: "Registrar's Office SSK",
  title: 'តារាងសម្រង់ពិន្ទុ / Score List',
  teacherName: 'Bin Soratha',
  subject: 'Math',
  grade: '1A',
  quarter: 'Quarter 1',
  date: '21-Sep-26',
  notesKh1: 'ចំណាំ ៖ តារាងសម្រង់ពិន្ទុនេះគឺសម្រាប់លោកគ្រូអ្នកគ្រូកត់ពិន្ទុចូល ដើម្បីងាយស្រួលក្នុងការបញ្ចូលពិន្ទុក្នុងតារាងសម្រង់ពិន្ទុអេឡិចត្រូនិក (E-Gradebook) ។',
  notesKh2: 'តារាងសម្រង់ពិន្ទុនេះគឺប្រើប្រាស់សម្រាប់ធ្វើការកត់ត្រាផ្ទាល់ខ្លួនរបស់លោកគ្រូអ្នកគ្រូតែប៉ុណ្ណោះ។',
  notesEn1: '* NOTE: This list is for teachers to keep record of all kinds of scores which is served as hard copies for egrade-book input.',
  notesEn2: 'The list is for teacher personal use. It is not required by the office.',
  rows: [
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
  ]
};

// API: Check server configuration
app.get('/api/config', (req, res) => {
  res.json({
    hasServerApiKey: !!process.env.GEMINI_API_KEY
  });
});

// API: Load preloaded sample data
app.get('/api/sample-data', (req, res) => {
  res.json({
    success: true,
    data: sampleExtractedData,
    sampleImages: [
      '/samples/page1_sample.jpg',
      '/samples/page2_sample.jpg'
    ]
  });
});

// API: Process and extract images with Gemini
app.post('/api/extract', upload.array('images', 10), async (req, res) => {
  try {
    const files = req.files;
    if (!files || files.length === 0) {
      return res.status(400).json({ error: 'សូមជ្រើសរើស ឬទម្លាក់ File រូបភាពយ៉ាងហោចណាស់ 1 (Please upload at least 1 image file).' });
    }

    const apiKey = req.body.apiKey || req.headers['x-gemini-key'] || process.env.GEMINI_API_KEY;
    if (!apiKey) {
      return res.status(400).json({
        error: 'API_KEY_REQUIRED',
        message: 'សូមបញ្ចូល Google Gemini API Key របស់អ្នកនៅក្នុង Setting ឬផ្ដល់ជា Header។ អ្នកអាចទទួលបាន Free API Key នៅ Google AI Studio (https://aistudio.google.com/app/apikey)'
      });
    }

    const imagePayloads = files.map(file => ({
      buffer: file.buffer,
      mimeType: file.mimetype
    }));

    const result = await extractTableFromImages(imagePayloads, apiKey);
    res.json({ success: true, data: result });
  } catch (error) {
    console.error('Extraction error:', error);
    res.status(500).json({
      error: 'EXTRACTION_FAILED',
      message: error.message || 'Error processing images'
    });
  }
});

// API: Export to Excel (.xlsx)
app.post('/api/export', async (req, res) => {
  try {
    const { data, filename } = req.body;
    if (!data) {
      return res.status(400).json({ error: 'No data provided for export' });
    }

    let workbook;
    if (data.isScoreList !== false && data.rows && (data.rows.length === 0 || data.rows[0].hw !== undefined || data.rows[0].name_en !== undefined)) {
      workbook = await generateScoreListWorkbook(data);
    } else if (data.genericTable) {
      workbook = await generateGenericTableWorkbook(data.genericTable);
    } else {
      workbook = await generateScoreListWorkbook(data);
    }

    const safeFilename = (filename || 'Score_List')
      .replace(/[^a-zA-Z0-9_\-\u1780-\u17FF]/g, '_') + '.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(safeFilename)}"`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ error: error.message || 'Failed to export Excel file' });
  }
});

// Serve frontend SPA fallback
app.use((req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, HOST, () => {
  console.log(`=======================================================`);
  console.log(`🚀 Image to Excel Web App is running!`);
  console.log(`👉 Access URL: http://${HOST}:${PORT}`);
  console.log(`=======================================================`);
});
