const { GoogleGenAI } = require('@google/genai');

/**
 * Extracts structured data from image buffers using Google Gemini API.
 * @param {Array<{buffer: Buffer, mimeType: string}>} images
 * @param {string} apiKey
 */
async function extractTableFromImages(images, apiKey) {
  if (!apiKey) {
    throw new Error('API_KEY_REQUIRED: Please provide your Gemini API Key in the settings or .env file.');
  }

  const ai = new GoogleGenAI({ apiKey });

  const prompt = `
You are an expert AI document analyzer and OCR specialist specializing in recognizing both Khmer Unicode script, English text, and handwritten numbers in tables/spreadsheets.

Analyze the uploaded image(s). They represent a score sheet / grade book or a data table (e.g., from Western International School / Registrar's Office SSK: "តារាងសម្រង់ពិន្ទុ / Score List").
If there are multiple pages (e.g. Page 1 rows 1-18, Page 2 rows 19-22+), merge and sequence all the rows in order.

Please extract the data accurately and return a single valid JSON object strictly matching this JSON schema:
{
  "isScoreList": true,
  "schoolNameKh": "សាលាអន្តរជាតិ វេស្ទើន",
  "officeName": "Registrar's Office SSK",
  "title": "តារាងសម្រង់ពិន្ទុ / Score List",
  "teacherName": "Teacher's name if present",
  "subject": "Subject if present (e.g., Math)",
  "grade": "Grade if present (e.g., 1A)",
  "quarter": "Quarter if present (e.g., Quarter 1)",
  "date": "Date if present (e.g., 21-Sep-26)",
  "notesKh1": "Footer note Khmer line 1",
  "notesKh2": "Footer note Khmer line 2",
  "notesEn1": "Footer note English line 1",
  "notesEn2": "Footer note English line 2",
  "rows": [
    {
      "no": 1,
      "name_en": "STUDENT ENGLISH NAME",
      "name_kh": "ឈ្មោះសិស្សជាភាសាខ្មែរ",
      "sex": "M or F",
      "group": "III",
      "conduct": "",
      "cp": "",
      "hw": [100, 100, 100, 100, "", ""],
      "qz": [100, 91, 95, "", "", ""],
      "mt": [95, 94, ""]
    }
  ]
}

CRITICAL RULES:
1. Pay maximum attention to Khmer spelling with correct sub-consonants and vowel diacritics (e.g., ឈាងច័ន្ទ រ័ត្នមុនី, ហេង ម៉ាណានីកា, ហ៊ូ តិបុញ្ញាលោកា, ណារ៉ុង ខេមរៈ, នីរ៉ាបែល វីរិយា, អ៊ូច រ៉ូហ្សាណា, ពៅ អាមីដា, ផល ដេនវិជ័យ, ផាត់ ភក្ត្រា, សំស៊ឹម សុភក្តិណី, សុវណ្ណ សិរីភានុមា, ស៊ុន ផេងអាន, សួន សិទ្ធិកា, តាំង មុយលី, តែ សំណាងមុនិសត្យា, ទី អ៊ីស៊ីង, វណ្ណៈ សូហ្វីយ៉ាដា).
2. Transcribe handwritten score numbers precisely. Common values are 100, 95, 90, 80, 50, 0, etc. If a cell is blank/empty, represent it as "".
3. "hw" must be an array of exactly 6 items (Homework 1 to 6).
4. "qz" must be an array of exactly 6 items (Quizzes 1 to 6).
5. "mt" must be an array of exactly 3 items (Monthly Test 1 to 3).
6. If the uploaded image is NOT a Score List, but a generic table: set "isScoreList": false, and provide "genericTable": { "headers": ["Col1", "Col2", ...], "rows": [ ["val1", "val2", ...], ... ] }.
7. Output ONLY the raw JSON without markdown code fences or backticks.
`;

  const contents = [];
  contents.push(prompt);

  for (const img of images) {
    contents.push({
      inlineData: {
        data: img.buffer.toString('base64'),
        mimeType: img.mimeType || 'image/jpeg'
      }
    });
  }

  // Models to attempt in order of preference
  const candidateModels = ['gemini-3.8-flash', 'gemini-1.5-flash', 'gemini-2.0-flash'];
  let lastError = null;

  for (const modelName of candidateModels) {
    try {
      console.log(`Attempting extraction with model: ${modelName}...`);
      const response = await ai.models.generateContent({
        model: modelName,
        contents: contents,
        config: {
          temperature: 0.1,
          responseMimeType: 'application/json'
        }
      });

      const text = response.text.trim();
      const cleaned = text.replace(/^```json\s*/i, '').replace(/```$/i, '').trim();
      const parsed = JSON.parse(cleaned);
      return parsed;
    } catch (err) {
      console.warn(`Model ${modelName} failed:`, err.message || err);
      lastError = err;
      // If error indicates 404 or model unavailable, try next candidate
      if (err.status === 404 || (err.message && (err.message.includes('404') || err.message.includes('not found') || err.message.includes('no longer available')))) {
        continue;
      }
      throw err;
    }
  }

  throw lastError || new Error('All candidate Gemini models failed.');
}

module.exports = {
  extractTableFromImages
};

