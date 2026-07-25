const TelegramBot = require('node-telegram-bot-api');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const os = require('os');
const http = require('http');
const ExcelJS = require('exceljs');
const pdfjsLib = require('pdfjs-dist');

const TOKEN = process.env.BOT_TOKEN;
const GEMINI_KEY = process.env.GEMINI_API_KEY || '';
const GROQ_KEY = process.env.GROQ_API_KEY || '';

if (!TOKEN) { console.error('BOT_TOKEN missing!'); process.exit(1); }

const bot = new TelegramBot(TOKEN, { polling: true });
const userState = {};

const mainMenu = {
  reply_markup: {
    keyboard: [
      [{ text: '📄 PDF → Excel' }, { text: '📝 PDF → Word' }],
      [{ text: '📋 PDF → CSV' },   { text: '📃 PDF → Text' }],
      [{ text: '🖼️ Image Tools' }, { text: '🔊 Text → Voice' }],
      [{ text: '🤖 Gemini AI' },   { text: '🦙 Groq AI' }],
      [{ text: '❓ របៀបប្រើ' },    { text: '📞 ទាក់ទង' }],
    ],
    resize_keyboard: true,
    persistent: true
  }
};

const backMenu = {
  reply_markup: { keyboard: [[{ text: '🔙 ត្រឡប់ Menu' }]], resize_keyboard: true }
};

const imageMenu = {
  reply_markup: {
    keyboard: [
      [{ text: '🗜️ Compress រូបភាព' }, { text: '↔️ Resize រូបភាព' }],
      [{ text: '🔄 JPG → PNG' },        { text: '🔄 PNG → JPG' }],
      [{ text: '🔙 ត្រឡប់ Menu' }],
    ],
    resize_keyboard: true
  }
};

bot.onText(/\/start/, (msg) => {
  const name = msg.from.first_name || 'បង';
  const aiStatus = [];
  if (GEMINI_KEY) aiStatus.push('🤖 Gemini AI OK');
  if (GROQ_KEY) aiStatus.push('🦙 Groq AI OK');
  if (!aiStatus.length) aiStatus.push('AI មិនទាន់ active');
  bot.sendMessage(msg.chat.id,
    `👋 សួស្ដី ${name}!\n\n` +
    `🤖 AI Assistant Bot\n\n` +
    `PDF Convert / Image / TTS / AI Chat\n` +
    `${aiStatus.join(' | ')}\n\n` +
    `ជ្រើសពី Menu 👇`,
    mainMenu
  );
});

bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text;
  if (!text) return;
  const state = userState[chatId];

  if (text === '🔙 ត្រឡប់ Menu') {
    delete userState[chatId];
    return bot.sendMessage(chatId, 'Menu ចម្បង', mainMenu);
  }

  if (text === '❓ របៀបប្រើ') {
    return bot.sendMessage(chatId,
      'របៀបប្រើ:\n\n' +
      'PDF: ចុច format → ផ្ញើ PDF → ទទួល file\n' +
      'Image: ចុច Image Tools → ជ្រើស → ផ្ញើរូប\n' +
      'TTS: ចុច Text→Voice → ផ្ញើ text\n' +
      'AI: ចុច Gemini/Groq → សួរសំណួរ', mainMenu);
  }

  if (text === '📞 ទាក់ទង') {
    return bot.sendMessage(chatId, 'Telegram: @your_username', mainMenu);
  }

  const fmtMap = {
    '📄 PDF → Excel': 'xlsx',
    '📝 PDF → Word':  'docx',
    '📋 PDF → CSV':   'csv',
    '📃 PDF → Text':  'txt',
  };
  if (fmtMap[text]) {
    userState[chatId] = { step: 'waiting_pdf', format: fmtMap[text] };
    return bot.sendMessage(chatId, `ផ្ញើ PDF file មកខ្ញុំ`, backMenu);
  }

  if (text === '🖼️ Image Tools') {
    return bot.sendMessage(chatId, 'ជ្រើសសកម្មភាព:', imageMenu);
  }

  const imgActions = {
    '🗜️ Compress រូបភាព': 'compress',
    '↔️ Resize រូបភាព':   'resize',
    '🔄 JPG → PNG':       'jpg2png',
    '🔄 PNG → JPG':       'png2jpg',
  };
  if (imgActions[text]) {
    userState[chatId] = { step: 'waiting_image', action: imgActions[text] };
    return bot.sendMessage(chatId, 'ផ្ញើ រូបភាព មកខ្ញុំ', backMenu);
  }

  if (text === '🔊 Text → Voice') {
    userState[chatId] = { step: 'waiting_tts' };
    return bot.sendMessage(chatId, 'ផ្ញើ អត្ថបទ (ខ្មែរ/English)', backMenu);
  }

  if (text === '🤖 Gemini AI') {
    if (!GEMINI_KEY) return bot.sendMessage(chatId, 'បន្ថែម GEMINI_API_KEY ក្នុង Render', mainMenu);
    userState[chatId] = { step: 'ai_chat', ai: 'gemini', history: [] };
    return bot.sendMessage(chatId, 'Gemini AI - សួរសំណួរ!', backMenu);
  }

  if (text === '🦙 Groq AI') {
    if (!GROQ_KEY) return bot.sendMessage(chatId, 'បន្ថែម GROQ_API_KEY ក្នុង Render', mainMenu);
    userState[chatId] = { step: 'ai_chat', ai: 'groq', history: [] };
    return bot.sendMessage(chatId, 'Groq AI - សួរសំណួរ!', backMenu);
  }

  if (state?.step === 'waiting_tts') return handleTTS(chatId, text);
  if (state?.step === 'ai_chat') return handleAI(chatId, text, state);

  bot.sendMessage(chatId, 'ជ្រើសពី Menu', mainMenu);
});

bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  const state = userState[chatId];
  const doc = msg.document;
  if (!state || state.step !== 'waiting_pdf') return bot.sendMessage(chatId, 'ជ្រើស format ជាមុន!', mainMenu);
  if (doc.mime_type !== 'application/pdf') return bot.sendMessage(chatId, 'PDF ប៉ុណ្ណោះ!', backMenu);
  if (doc.file_size > 20 * 1024 * 1024) return bot.sendMessage(chatId, 'Max 20MB!', backMenu);

  const fmt = state.format;
  const statusMsg = await bot.sendMessage(chatId, `កំពុងបម្លែង PDF → ${fmt.toUpperCase()}...`);

  try {
    const fileLink = await bot.getFileLink(doc.file_id);
    const resp = await axios.get(fileLink, { responseType: 'arraybuffer' });
    const lines = await extractPdfLines(new Uint8Array(resp.data));
    if (!lines || lines.length === 0) throw new Error('មិនអាចអាន PDF នេះ');

    const baseName = (doc.file_name || 'document').replace(/\.pdf$/i, '');
    const outName = baseName + '.' + fmt;
    const outputPath = path.join(os.tmpdir(), outName);

    if (fmt === 'xlsx') await makeXlsx(lines, outputPath);
    else if (fmt === 'docx') makeDocx(lines, outputPath);
    else if (fmt === 'csv') fs.writeFileSync(outputPath, makeCsv(lines), 'utf8');
    else fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');

    await bot.sendDocument(chatId, outputPath, { caption: `✅ ${outName} រួចរាល់!` });
    fs.unlinkSync(outputPath);
    delete userState[chatId];
    await bot.editMessageText('✅ រួចរាល់!', { chat_id: chatId, message_id: statusMsg.message_id });
    bot.sendMessage(chatId, 'ចង់បម្លែងទៀត?', mainMenu);
  } catch (err) {
    bot.editMessageText(`❌ ${err.message}`, { chat_id: chatId, message_id: statusMsg.message_id });
    delete userState[chatId];
    bot.sendMessage(chatId, 'ត្រឡប់ Menu', mainMenu);
  }
});

bot.on('photo', async (msg) => {
  const chatId = msg.chat.id;
  const state = userState[chatId];
  if (!state || state.step !== 'waiting_image') return bot.sendMessage(chatId, 'ចុច Image Tools ជាមុន!', mainMenu);

  const photo = msg.photo[msg.photo.length - 1];
  const statusMsg = await bot.sendMessage(chatId, 'កំពុងដំណើរការ...');

  try {
    const fileLink = await bot.getFileLink(photo.file_id);
    const resp = await axios.get(fileLink, { responseType: 'arraybuffer' });
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
    await bot.sendDocument(chatId, outputPath, { caption: `✅ ${origKB}KB → ${newKB}KB` });
    fs.unlinkSync(outputPath);
    delete userState[chatId];
    await bot.editMessageText('✅ Image ជោគជ័យ!', { chat_id: chatId, message_id: statusMsg.message_id });
    bot.sendMessage(chatId, 'ចង់ធ្វើ Image ទៀត?', imageMenu);
  } catch (err) {
    bot.editMessageText(`❌ ${err.message}`, { chat_id: chatId, message_id: statusMsg.message_id });
    delete userState[chatId];
    bot.sendMessage(chatId, 'ត្រឡប់ Menu', mainMenu);
  }
});

async function handleTTS(chatId, text) {
  const statusMsg = await bot.sendMessage(chatId, 'កំពុងបង្កើត audio...');
  try {
    const lang = /[\u1780-\u17FF]/.test(text) ? 'km' : 'en';
    const url = `https://translate.google.com/translate_tts?ie=UTF-8&q=${encodeURIComponent(text)}&tl=${lang}&client=tw-ob`;
    const resp = await axios.get(url, { responseType: 'arraybuffer', headers: { 'User-Agent': 'Mozilla/5.0' } });
    const audioPath = path.join(os.tmpdir(), 'tts.mp3');
    fs.writeFileSync(audioPath, Buffer.from(resp.data));
    await bot.sendAudio(chatId, audioPath, { caption: `🔊 "${text.substring(0, 50)}"` });
    fs.unlinkSync(audioPath);
    await bot.editMessageText('✅ Audio រួចរាល់!', { chat_id: chatId, message_id: statusMsg.message_id });
    bot.sendMessage(chatId, 'ផ្ញើ text ទៀត ឬ ចុច 🔙', backMenu);
  } catch (err) {
    bot.editMessageText(`❌ TTS Error: ${err.message}`, { chat_id: chatId, message_id: statusMsg.message_id });
    delete userState[chatId];
    bot.sendMessage(chatId, 'ត្រឡប់ Menu', mainMenu);
  }
}

async function handleAI(chatId, question, state) {
  const label = state.ai === 'gemini' ? '🤖 Gemini' : '🦙 Groq';
  const statusMsg = await bot.sendMessage(chatId, `${label} កំពុងគិត...`);
  const sys = 'Reply in the same language as user. If Khmer reply Khmer.';
  let history = state.history || [];

  try {
    let answer = '';
    if (state.ai === 'gemini') {
      history.push({ role: 'user', parts: [{ text: question }] });
      if (history.length > 20) history = history.slice(-20);
      const resp = await axios.post(
        `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_KEY}`,
        { system_instruction: { parts: [{ text: sys }] }, contents: history, generationConfig: { maxOutputTokens: 1024 } }
      );
      answer = resp.data.candidates[0].content.parts[0].text;
      history.push({ role: 'model', parts: [{ text: answer }] });
    } else {
      history.push({ role: 'user', content: question });
      if (history.length > 20) history = history.slice(-20);
      const resp = await axios.post(
        'https://api.groq.com/openai/v1/chat/completions',
        { model: 'llama-3.3-70b-versatile', messages: [{ role: 'system', content: sys }, ...history], max_tokens: 1024 },
        { headers: { Authorization: `Bearer ${GROQ_KEY}` } }
      );
      answer = resp.data.choices[0].message.content;
      history.push({ role: 'assistant', content: answer });
    }
    userState[chatId].history = history;
    await bot.deleteMessage(chatId, statusMsg.message_id);
    bot.sendMessage(chatId, `${label}:\n\n${answer}`, backMenu);
  } catch (err) {
    const errMsg = err.response?.data?.error?.message || err.message;
    bot.editMessageText(`❌ AI Error: ${errMsg}`, { chat_id: chatId, message_id: statusMsg.message_id });
  }
}

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
  const ws = wb.addWorksheet('Sheet1');
  lines.forEach(line => {
    const cells = line.split(/\t|\s{3,}/);
    ws.addRow(cells.length > 1 ? cells : [line]);
  });
  await wb.xlsx.writeFile(outputPath);
}

function makeDocx(lines, outputPath) {
  const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const paras = lines.map(l => l.trim() ? `<w:p><w:r><w:t xml:space="preserve">${esc(l)}</w:t></w:r></w:p>` : '<w:p/>').join('');
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paras}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
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

const PORT = process.env.PORT || 3000;
http.createServer((req, res) => { res.writeHead(200); res.end('Bot running!'); }).listen(PORT);
bot.on('polling_error', err => console.error('Polling error:', err.message));
console.log('Bot started! PDF+Image+TTS+AI');
