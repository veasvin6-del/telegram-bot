const TelegramBot = require('node-telegram-bot-api');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const os = require('os');
const http = require('http');
const ExcelJS = require('exceljs');
const pdfjsLib = require('pdfjs-dist');

// ===== CONFIG =====
const TOKEN = process.env.BOT_TOKEN;
const GEMINI_KEY = process.env.GEMINI_API_KEY || '';
const GROQ_KEY = process.env.GROQ_API_KEY || '';
const PORT = process.env.PORT || 3000;

if (!TOKEN) { console.error('❌ BOT_TOKEN missing!'); process.exit(1); }

const bot = new TelegramBot(TOKEN, { polling: true });
const userState = {};
const userStats = {}; // Track usage per user

// ===== HELPERS =====
const sleep = ms => new Promise(r => setTimeout(r, ms));

function trackUser(chatId, action) {
  if (!userStats[chatId]) userStats[chatId] = { count: 0, actions: [] };
  userStats[chatId].count++;
  userStats[chatId].actions.push(action);
}

// ===== MENUS =====
const MAIN_MENU = {
  reply_markup: {
    keyboard: [
      [{ text: '📄 PDF → Excel' }, { text: '📝 PDF → Word' }],
      [{ text: '📋 PDF → CSV' },   { text: '📃 PDF → Text' }],
      [{ text: '🖼️ Image Tools' }, { text: '🔊 Text → Voice' }],
      [{ text: '🤖 Gemini AI' },   { text: '🦙 Groq AI' }],
      [{ text: '📊 របៀបប្រើ' },    { text: '⚙️ ការកំណត់' }],
    ],
    resize_keyboard: true,
    persistent: true
  }
};

const BACK_MENU = {
  reply_markup: {
    keyboard: [[{ text: '🔙 ត្រឡប់ Menu' }]],
    resize_keyboard: true
  }
};

const IMAGE_MENU = {
  reply_markup: {
    keyboard: [
      [{ text: '🗜️ Compress រូបភាព' }, { text: '↔️ Resize រូបភាព' }],
      [{ text: '🔄 JPG → PNG' },        { text: '🔄 PNG → JPG' }],
      [{ text: '🔙 ត្រឡប់ Menu' }],
    ],
    resize_keyboard: true
  }
};

const FORMAT_LABELS = {
  xlsx: 'Excel (.xlsx)',
  docx: 'Word (.docx)',
  csv: 'CSV (.csv)',
  txt: 'Text (.txt)'
};

// ===== /start =====
bot.onText(/\/start/, async (msg) => {
  const chatId = msg.chat.id;
  const name = msg.from.first_name || 'អ្នកប្រើ';
  const username = msg.from.username ? `@${msg.from.username}` : '';
  trackUser(chatId, 'start');

  const aiReady = [];
  if (GEMINI_KEY) aiReady.push('Gemini ✅');
  if (GROQ_KEY) aiReady.push('Groq ✅');

  const greeting = `🌟 *សួស្តី ${name}!* ${username}\n\n` +
    `ខ្ញុំជា *AI Assistant Bot* — ជំនួយការឆ្លាតវៃ\n` +
    `━━━━━━━━━━━━━━━━━━━\n\n` +
    `📋 *មុខងារសំខាន់ៗ:*\n` +
    `📄 បម្លែង PDF → Excel, Word, CSV, Text\n` +
    `🖼️ កែសម្រួលរូបភាព Compress/Resize/Convert\n` +
    `🔊 បម្លែង Text → Audio (ខ្មែរ & English)\n` +
    `🤖 AI Chat: ${aiReady.length > 0 ? aiReady.join(' | ') : 'មិនទាន់ active'}\n\n` +
    `━━━━━━━━━━━━━━━━━━━\n` +
    `👇 *ជ្រើសមុខងារពី Menu ខាងក្រោម*`;

  await bot.sendMessage(chatId, greeting, {
    parse_mode: 'Markdown',
    ...MAIN_MENU
  });
});

// ===== /help =====
bot.onText(/\/help/, async (msg) => {
  const chatId = msg.chat.id;
  const helpText =
    `📖 *របៀបប្រើប្រាស់ Bot*\n` +
    `━━━━━━━━━━━━━━━━━━━\n\n` +
    `*📄 PDF Converter:*\n` +
    `1. ចុចប៊ូតុង format ដែលអ្នកចង់បាន\n` +
    `2. ផ្ញើ PDF file\n` +
    `3. ទទួលបាន file ភ្លាមៗ ✅\n\n` +
    `*🖼️ Image Tools:*\n` +
    `1. ចុច Image Tools\n` +
    `2. ជ្រើសសកម្មភាព\n` +
    `3. ផ្ញើរូបភាព\n\n` +
    `*🔊 Text to Voice:*\n` +
    `1. ចុច Text → Voice\n` +
    `2. វាយ ឬ paste អត្ថបទ\n` +
    `3. ទទួលបាន audio file\n\n` +
    `*🤖 AI Chat:*\n` +
    `1. ចុច Gemini AI ឬ Groq AI\n` +
    `2. សួរសំណួរអ្វីក៏បាន\n` +
    `3. AI ឆ្លើយភ្លាមៗ\n\n` +
    `━━━━━━━━━━━━━━━━━━━\n` +
    `⚠️ *ដែនកំណត់:* PDF & Image ≤ 20MB\n` +
    `💡 *Tip:* Bot ដំណើរការ 24/7 ដោយស្វ័យប្រវត្តិ`;

  await bot.sendMessage(chatId, helpText, { parse_mode: 'Markdown', ...MAIN_MENU });
});

// ===== Handle Text Messages =====
bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text;
  if (!text || text.startsWith('/')) return;

  const state = userState[chatId];

  // Back button
  if (text === '🔙 ត្រឡប់ Menu') {
    delete userState[chatId];
    return bot.sendMessage(chatId,
      '🏠 *Menu ចម្បង*\nជ្រើសមុខងារដែលអ្នកចង់ប្រើ 👇',
      { parse_mode: 'Markdown', ...MAIN_MENU }
    );
  }

  // PDF formats
  const fmtMap = {
    '📄 PDF → Excel': 'xlsx',
    '📝 PDF → Word': 'docx',
    '📋 PDF → CSV': 'csv',
    '📃 PDF → Text': 'txt',
  };

  if (fmtMap[text]) {
    const fmt = fmtMap[text];
    userState[chatId] = { step: 'waiting_pdf', format: fmt };
    trackUser(chatId, `pdf_${fmt}`);
    return bot.sendMessage(chatId,
      `📤 *ផ្ញើ PDF File*\n\n` +
      `Format: *${FORMAT_LABELS[fmt]}*\n\n` +
      `✅ គ្រាន់តែ drag & drop ឬ ជ្រើស PDF file\n` +
      `⚠️ ទំហំអតិបរិមា: 20MB`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  // Image Tools
  if (text === '🖼️ Image Tools') {
    return bot.sendMessage(chatId,
      `🖼️ *Image Tools*\n\n` +
      `ជ្រើសសកម្មភាពដែលអ្នកចង់ធ្វើ:\n\n` +
      `🗜️ *Compress* — កាត់ទំហំ file\n` +
      `↔️ *Resize* — ផ្លាស់ប្ដូរ resolution\n` +
      `🔄 *Convert* — ប្ដូរ format រូបភាព`,
      { parse_mode: 'Markdown', ...IMAGE_MENU }
    );
  }

  const imgActions = {
    '🗜️ Compress រូបភាព': 'compress',
    '↔️ Resize រូបភាព': 'resize',
    '🔄 JPG → PNG': 'jpg2png',
    '🔄 PNG → JPG': 'png2jpg',
  };

  if (imgActions[text]) {
    userState[chatId] = { step: 'waiting_image', action: imgActions[text] };
    const desc = {
      compress: '🗜️ Compress — កាត់ size រូបភាព 60%',
      resize: '↔️ Resize — ផ្លាស់ ​width → 800px',
      jpg2png: '🔄 Convert JPG → PNG',
      png2jpg: '🔄 Convert PNG → JPG',
    };
    return bot.sendMessage(chatId,
      `📤 *ផ្ញើ រូបភាព*\n\n` +
      `Action: *${desc[imgActions[text]]}*\n\n` +
      `✅ ផ្ញើ photo ឬ file រូបភាពមកខ្ញុំ`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  // Text to Voice
  if (text === '🔊 Text → Voice') {
    userState[chatId] = { step: 'waiting_tts' };
    return bot.sendMessage(chatId,
      `🔊 *Text to Voice*\n\n` +
      ` វាយ ឬ paste *អត្ថបទ* ដែលអ្នកចង់ស្ដាប់\n\n` +
      `🇰🇭 គាំទ្រ: ភាសាខ្មែរ\n` +
      `🇺🇸 គាំទ្រ: English\n` +
      `🇹🇭 គាំទ្រ: ភាសាថៃ និងភាសាផ្សេងៗ`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  // Gemini AI
  if (text === '🤖 Gemini AI') {
    if (!GEMINI_KEY) {
      return bot.sendMessage(chatId,
        `⚠️ *Gemini AI មិនទាន់ Active*\n\n` +
        `ទៅ Render → Environment → បន្ថែម:\n` +
        `\`GEMINI_API_KEY\`\n\n` +
        `🔗 aistudio.google.com → Get API Key`,
        { parse_mode: 'Markdown', ...MAIN_MENU }
      );
    }
    userState[chatId] = { step: 'ai_chat', ai: 'gemini', history: [] };
    return bot.sendMessage(chatId,
      `🤖 *Gemini AI — Google*\n` +
      `━━━━━━━━━━━━━━━━━━━\n\n` +
      `✅ AI ត្រៀមខ្លួនហើយ!\n` +
      `💬 សួរសំណួរអ្វីក៏បាន ខ្ញុំនឹងឆ្លើយភ្លាម\n\n` +
      `🌐 ឆ្លើយបាន: ខ្មែរ, English, Thai...\n` +
      `🔙 ចុច ត្រឡប់ Menu ដើម្បីចេញ`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  // Groq AI
  if (text === '🦙 Groq AI') {
    if (!GROQ_KEY) {
      return bot.sendMessage(chatId,
        `⚠️ *Groq AI មិនទាន់ Active*\n\n` +
        `ទៅ Render → Environment → បន្ថែម:\n` +
        `\`GROQ_API_KEY\`\n\n` +
        `🔗 console.groq.com → API Keys`,
        { parse_mode: 'Markdown', ...MAIN_MENU }
      );
    }
    userState[chatId] = { step: 'ai_chat', ai: 'groq', history: [] };
    return bot.sendMessage(chatId,
      `🦙 *Groq AI — Llama 3.3*\n` +
      `━━━━━━━━━━━━━━━━━━━\n\n` +
      `⚡ AI លឿនបំផុត FREE!\n` +
      `💬 សួរសំណួរអ្វីក៏បាន ខ្ញុំនឹងឆ្លើយភ្លាម\n\n` +
      `🌐 ឆ្លើយបាន: ខ្មែរ, English, Thai...\n` +
      `🔙 ចុច ត្រឡប់ Menu ដើម្បីចេញ`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  // Help & Settings
  if (text === '📊 របៀបប្រើ') {
    return bot.onText(/\/help/, () => {})(msg);
  }

  if (text === '⚙️ ការកំណត់') {
    const stats = userStats[chatId] || { count: 0 };
    return bot.sendMessage(chatId,
      `⚙️ *ព័ត៌មាន Bot*\n` +
      `━━━━━━━━━━━━━━━━━━━\n\n` +
      `🤖 *Bot:* @Imvinconverter_bot\n` +
      `📊 *Version:* 2.0 Pro\n` +
      `🟢 *Status:* Online 24/7\n` +
      `📈 *ការប្រើប្រាស់របស់អ្នក:* ${stats.count} ដង\n\n` +
      `*AI Status:*\n` +
      `${GEMINI_KEY ? '🟢' : '🔴'} Gemini AI\n` +
      `${GROQ_KEY ? '🟢' : '🔴'} Groq AI\n\n` +
      `💬 *ទាក់ទង:* @your_username`,
      { parse_mode: 'Markdown', ...MAIN_MENU }
    );
  }

  // Handle ongoing states
  if (state?.step === 'waiting_tts') return handleTTS(chatId, text);
  if (state?.step === 'ai_chat') return handleAI(chatId, text, state);

  // Default response
  bot.sendMessage(chatId,
    `👇 សូមជ្រើសមុខងារពី *Menu* ខាងក្រោម`,
    { parse_mode: 'Markdown', ...MAIN_MENU }
  );
});

// ===== PDF Handler =====
bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  const state = userState[chatId];
  const doc = msg.document;

  if (!state || state.step !== 'waiting_pdf') {
    return bot.sendMessage(chatId,
      `⚠️ *សូមជ្រើស Format ជាមុន!*\n\nចុចប៊ូតុង PDF format ពី Menu`,
      { parse_mode: 'Markdown', ...MAIN_MENU }
    );
  }

  if (doc.mime_type !== 'application/pdf') {
    return bot.sendMessage(chatId,
      `❌ *File មិនត្រឹមត្រូវ!*\n\nសូមផ្ញើ *PDF file* ប៉ុណ្ណោះ\n📌 Extension: .pdf`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  const fileSizeMB = (doc.file_size / 1024 / 1024).toFixed(1);
  if (doc.file_size > 20 * 1024 * 1024) {
    return bot.sendMessage(chatId,
      `❌ *File ធំពេក!*\n\n` +
      `ទំហំ file: *${fileSizeMB}MB*\n` +
      `ទំហំអតិបរិមា: *20MB*\n\n` +
      `💡 សូម compress PDF ជាមុន`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  const fmt = state.format;
  const fmtLabel = FORMAT_LABELS[fmt];
  trackUser(chatId, `converted_${fmt}`);

  const statusMsg = await bot.sendMessage(chatId,
    `⏳ *កំពុងដំណើរការ...*\n\n` +
    `📄 File: ${doc.file_name || 'document.pdf'}\n` +
    `📦 ទំហំ: ${fileSizeMB}MB\n` +
    `🎯 Format: ${fmtLabel}\n\n` +
    `_សូមរង់ចាំបន្តិច..._`,
    { parse_mode: 'Markdown' }
  );

  try {
    const fileLink = await bot.getFileLink(doc.file_id);
    const resp = await axios.get(fileLink, { responseType: 'arraybuffer', timeout: 30000 });
    const lines = await extractPdfLines(new Uint8Array(resp.data));

    if (!lines || lines.length === 0) {
      await bot.editMessageText(
        `❌ *មិនអាចអាន PDF នេះបាន!*\n\n` +
        `*មូលហេតុ:*\n` +
        `• PDF អាចជា scanned image\n` +
        `• PDF ត្រូវបាន encrypt/protect\n\n` +
        `💡 *ដំណោះស្រាយ:* ប្រើ PDF ដែលមាន text`,
        { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
      );
      delete userState[chatId];
      return bot.sendMessage(chatId, '🏠 ត្រឡប់ Menu', MAIN_MENU);
    }

    const baseName = (doc.file_name || 'document').replace(/\.pdf$/i, '');
    const outName = `${baseName}.${fmt}`;
    const outputPath = path.join(os.tmpdir(), outName);

    if (fmt === 'xlsx') await makeXlsx(lines, outputPath);
    else if (fmt === 'docx') makeDocx(lines, outputPath);
    else if (fmt === 'csv') fs.writeFileSync(outputPath, makeCsv(lines), 'utf8');
    else fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');

    const outSizeMB = (fs.statSync(outputPath).size / 1024).toFixed(0);

    await bot.sendDocument(chatId, outputPath, {
      caption:
        `✅ *បម្លែងបានជោគជ័យ!*\n\n` +
        `📄 *${outName}*\n` +
        `📦 ទំហំ: ${outSizeMB}KB\n` +
        `📑 ${lines.filter(l => l).length} rows\n\n` +
        `_ចុច Download ដើម្បីទាញយក_`,
      parse_mode: 'Markdown'
    });

    fs.unlinkSync(outputPath);
    delete userState[chatId];

    await bot.editMessageText(
      `✅ *រួចរាល់ហើយ!*\n\nPDF → ${fmtLabel} ✨`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
    );

    bot.sendMessage(chatId,
      `🎉 *ជោគជ័យ!*\n\nចង់បម្លែង PDF ទៀតទេ? 👇`,
      { parse_mode: 'Markdown', ...MAIN_MENU }
    );

  } catch (err) {
    console.error('PDF Error:', err);
    await bot.editMessageText(
      `❌ *មានបញ្ហាកើតឡើង!*\n\n` +
      `Error: ${err.message}\n\n` +
      `💡 សូម try ម្ដងទៀត ឬ ផ្ញើ PDF ផ្សេង`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
    );
    delete userState[chatId];
    bot.sendMessage(chatId, '🏠 ត្រឡប់ Menu', MAIN_MENU);
  }
});

// ===== Image Handler =====
bot.on('photo', async (msg) => {
  const chatId = msg.chat.id;
  const state = userState[chatId];

  if (!state || state.step !== 'waiting_image') {
    return bot.sendMessage(chatId,
      `⚠️ *ចុច Image Tools ជាមុន!*`,
      { parse_mode: 'Markdown', ...MAIN_MENU }
    );
  }

  const photo = msg.photo[msg.photo.length - 1];
  const statusMsg = await bot.sendMessage(chatId,
    `⏳ *កំពុងដំណើរការរូបភាព...*\n_សូមរង់ចាំ..._`,
    { parse_mode: 'Markdown' }
  );

  try {
    const fileLink = await bot.getFileLink(photo.file_id);
    const resp = await axios.get(fileLink, { responseType: 'arraybuffer', timeout: 30000 });
    const imgBuf = Buffer.from(resp.data);
    const ext = state.action === 'jpg2png' ? 'png' : 'jpg';
    const outputPath = path.join(os.tmpdir(), `img_${Date.now()}.${ext}`);

    try {
      const sharp = require('sharp');
      let p = sharp(imgBuf);
      if (state.action === 'compress') p = p.jpeg({ quality: 60 });
      else if (state.action === 'resize') p = p.resize(800, null, { withoutEnlargement: true }).jpeg({ quality: 80 });
      else if (state.action === 'jpg2png') p = p.png();
      else if (state.action === 'png2jpg') p = p.jpeg({ quality: 90 });
      await p.toFile(outputPath);
    } catch (e) {
      fs.writeFileSync(outputPath, imgBuf);
    }

    const origKB = (imgBuf.length / 1024).toFixed(0);
    const newKB = (fs.statSync(outputPath).size / 1024).toFixed(0);
    const saved = Math.max(0, ((1 - newKB / origKB) * 100).toFixed(0));

    await bot.sendDocument(chatId, outputPath, {
      caption:
        `✅ *Image ដំណើរការបានជោគជ័យ!*\n\n` +
        `📁 ទំហំដើម: *${origKB}KB*\n` +
        `📁 ទំហំថ្មី: *${newKB}KB*\n` +
        `💾 សន្សំ: *${saved}%*`,
      parse_mode: 'Markdown'
    });

    fs.unlinkSync(outputPath);
    delete userState[chatId];
    await bot.editMessageText('✅ *Image ជោគជ័យ!*',
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
    );
    bot.sendMessage(chatId, '🖼️ ចង់ធ្វើ Image ទៀតទេ? 👇', IMAGE_MENU);

  } catch (err) {
    console.error('Image Error:', err);
    bot.editMessageText(`❌ *Error:* ${err.message}`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
    );
    delete userState[chatId];
    bot.sendMessage(chatId, '🏠 ត្រឡប់ Menu', MAIN_MENU);
  }
});

// ===== Text to Voice =====
async function handleTTS(chatId, text) {
  if (text.length > 500) {
    return bot.sendMessage(chatId,
      `⚠️ *អត្ថបទវែងពេក!*\n\nអតិបរិមា: 500 characters\nអ្នកវាយ: ${text.length} characters`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );
  }

  const statusMsg = await bot.sendMessage(chatId,
    `🔊 *កំពុងបង្កើត Audio...*\n_សូមរង់ចាំ..._`,
    { parse_mode: 'Markdown' }
  );

  try {
    const lang = /[\u1780-\u17FF]/.test(text) ? 'km' : 'en';
    const langName = lang === 'km' ? '🇰🇭 ខ្មែរ' : '🇺🇸 English';
    const url = `https://translate.google.com/translate_tts?ie=UTF-8&q=${encodeURIComponent(text)}&tl=${lang}&client=tw-ob`;

    const resp = await axios.get(url, {
      responseType: 'arraybuffer',
      headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' },
      timeout: 15000
    });

    const audioPath = path.join(os.tmpdir(), `tts_${Date.now()}.mp3`);
    fs.writeFileSync(audioPath, Buffer.from(resp.data));

    await bot.sendAudio(chatId, audioPath, {
      caption:
        `🔊 *Audio បានបង្កើតរួច!*\n\n` +
        `📝 "${text.substring(0, 80)}${text.length > 80 ? '...' : ''}"\n` +
        `🌐 ភាសា: ${langName}`,
      parse_mode: 'Markdown'
    });

    fs.unlinkSync(audioPath);
    await bot.editMessageText('✅ *Audio ជោគជ័យ!*',
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
    );
    bot.sendMessage(chatId,
      `🎧 *ចង់បង្កើត Audio ទៀតទេ?*\nផ្ញើ text ទៀត ឬ ចុច 🔙`,
      { parse_mode: 'Markdown', ...BACK_MENU }
    );

  } catch (err) {
    console.error('TTS Error:', err);
    bot.editMessageText(
      `❌ *TTS Error!*\n\n${err.message}\n\n💡 សូម try ម្ដងទៀត`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
    );
    delete userState[chatId];
    bot.sendMessage(chatId, '🏠 ត្រឡប់ Menu', MAIN_MENU);
  }
}

// ===== AI Chat =====
async function handleAI(chatId, question, state) {
  const isGemini = state.ai === 'gemini';
  const label = isGemini ? '🤖 Gemini' : '🦙 Groq';
  const statusMsg = await bot.sendMessage(chatId,
    `${label} *កំពុងគិត...*\n_⏳ សូមរង់ចាំ..._`,
    { parse_mode: 'Markdown' }
  );

  const sys = `You are a helpful, friendly AI assistant in a Telegram bot called "AI Assistant Bot". 
Always reply in the same language as the user. 
If user writes in Khmer (ខ្មែរ), reply in Khmer.
If user writes in English, reply in English.
Be concise, accurate, and helpful.
Use emoji occasionally to make responses engaging.`;

  let history = state.history || [];

  try {
    let answer = '';

    if (isGemini) {
      history.push({ role: 'user', parts: [{ text: question }] });
      if (history.length > 20) history = history.slice(-20);

      const resp = await axios.post(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`,
        {
          system_instruction: { parts: [{ text: sys }] },
          contents: history,
          generationConfig: { maxOutputTokens: 1024, temperature: 0.7 }
        },
        { timeout: 30000 }
      );
      answer = resp.data.candidates[0].content.parts[0].text;
      history.push({ role: 'model', parts: [{ text: answer }] });

    } else {
      history.push({ role: 'user', content: question });
      if (history.length > 20) history = history.slice(-20);

      const resp = await axios.post(
        'https://api.groq.com/openai/v1/chat/completions',
        {
          model: 'llama-3.3-70b-versatile',
          messages: [{ role: 'system', content: sys }, ...history],
          max_tokens: 1024,
          temperature: 0.7,
        },
        { headers: { Authorization: `Bearer ${GROQ_KEY}` }, timeout: 30000 }
      );
      answer = resp.data.choices[0].message.content;
      history.push({ role: 'assistant', content: answer });
    }

    userState[chatId].history = history;
    await bot.deleteMessage(chatId, statusMsg.message_id);

    // Split long messages
    if (answer.length > 3000) {
      const chunks = answer.match(/.{1,3000}/gs) || [answer];
      for (const chunk of chunks) {
        await bot.sendMessage(chatId, `${label}:\n\n${chunk}`, BACK_MENU);
        await sleep(300);
      }
    } else {
      bot.sendMessage(chatId, `${label}:\n\n${answer}`, BACK_MENU);
    }

  } catch (err) {
    console.error('AI Error:', err.response?.data || err.message);
    const errMsg = err.response?.data?.error?.message || err.message;
    bot.editMessageText(
      `❌ *AI Error!*\n\n${errMsg}\n\n💡 សូម try ម្ដងទៀត`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' }
    );
  }
}

// ===== PDF Helpers =====
async function extractPdfLines(data) {
  const pdf = await pdfjsLib.getDocument({ data, useWorkerFetch: false, isEvalSupported: false }).promise;
  const allLines = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const tc = await page.getTextContent();
    const yMap = {};
    tc.items.forEach(item => {
      const y = Math.round(item.transform[5]);
      if (!yMap[y]) yMap[y] = [];
      yMap[y].push(item.str);
    });
    Object.keys(yMap).map(Number).sort((a, b) => b - a).forEach(y => {
      const line = yMap[y].join(' ').trim();
      if (line) allLines.push(line);
    });
    if (i < pdf.numPages) allLines.push('');
  }
  return allLines;
}

async function makeXlsx(lines, outputPath) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Data');
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E74B5' } };
  lines.forEach((line, idx) => {
    const cells = line.split(/\t|\s{3,}/);
    const row = ws.addRow(cells.length > 1 ? cells : [line]);
    if (idx % 2 === 1) {
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F4FF' } };
    }
  });
  ws.columns.forEach(col => {
    let max = 10;
    col.eachCell(cell => { if (cell.value) max = Math.max(max, String(cell.value).length); });
    col.width = Math.min(max + 4, 50);
  });
  await wb.xlsx.writeFile(outputPath);
}

function makeDocx(lines, outputPath) {
  const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const paras = lines.map(l => l.trim()
    ? `<w:p><w:pPr><w:spacing w:after="120"/></w:pPr><w:r><w:t xml:space="preserve">${esc(l)}</w:t></w:r></w:p>`
    : '<w:p/>'
  ).join('');
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paras}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
  const wRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`;
  fs.writeFileSync(outputPath, buildZip({ '[Content_Types].xml': ct, '_rels/.rels': rels, 'word/document.xml': docXml, 'word/_rels/document.xml.rels': wRels }));
}

function buildZip(files) {
  const parts = [], cds = []; let offset = 0;
  Object.entries(files).forEach(([name, content]) => {
    const nb = Buffer.from(name), data = Buffer.from(content, 'utf8');
    const lh = Buffer.alloc(30 + nb.length);
    lh.writeUInt32LE(0x04034b50,0); lh.writeUInt16LE(20,4); lh.writeUInt32LE(0,6);
    lh.writeUInt32LE(0,10); lh.writeUInt32LE(0,14);
    lh.writeUInt32LE(data.length,18); lh.writeUInt32LE(data.length,22);
    lh.writeUInt16LE(nb.length,26); lh.writeUInt16LE(0,28); nb.copy(lh,30);
    parts.push(lh); parts.push(data);
    const cd = Buffer.alloc(46 + nb.length);
    cd.writeUInt32LE(0x02014b50,0); cd.writeUInt32LE(0x00140014,4);
    cd.writeUInt32LE(0,8); cd.writeUInt32LE(0,12); cd.writeUInt32LE(0,16);
    cd.writeUInt32LE(data.length,20); cd.writeUInt32LE(data.length,24);
    cd.writeUInt16LE(nb.length,28); cd.writeUInt32LE(0,30); cd.writeUInt32LE(0,34);
    cd.writeUInt32LE(0,38); cd.writeUInt32LE(offset,42); nb.copy(cd,46);
    offset += lh.length + data.length; cds.push(cd);
  });
  const cdBuf = Buffer.concat(cds);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50,0); eocd.writeUInt32LE(0,4);
  eocd.writeUInt16LE(Object.keys(files).length,8); eocd.writeUInt16LE(Object.keys(files).length,10);
  eocd.writeUInt32LE(cdBuf.length,12); eocd.writeUInt32LE(offset,16); eocd.writeUInt16LE(0,20);
  return Buffer.concat([...parts, cdBuf, eocd]);
}

function makeCsv(lines) {
  return lines.map(l => {
    const c = l.split(/\t|\s{3,}/);
    if (c.length > 1) return c.map(x => '"' + x.replace(/"/g,'""') + '"').join(',');
    return '"' + l.replace(/"/g,'""') + '"';
  }).join('\r\n');
}

// ===== Keep-Alive Server =====
http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({
    status: 'online',
    bot: '@Imvinconverter_bot',
    version: '2.0 Pro',
    uptime: Math.floor(process.uptime()) + 's'
  }));
}).listen(PORT, () => console.log(`✅ Keep-alive server on port ${PORT}`));

bot.on('polling_error', err => console.error('Polling error:', err.message));

console.log('🚀 AI Assistant Bot v2.0 Pro — Started!');
console.log(`🤖 Gemini: ${GEMINI_KEY ? 'Ready' : 'Not configured'}`);
console.log(`🦙 Groq: ${GROQ_KEY ? 'Ready' : 'Not configured'}`);
