const TelegramBot = require('node-telegram-bot-api');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const os = require('os');
const ExcelJS = require('exceljs');
const pdfjsLib = require('pdfjs-dist');

const TOKEN = "8834925461:AAFOs3i7ficN_v8oWGEENtjgIVDzT8dL9X0";

if (!TOKEN) { console.error('❌ BOT_TOKEN missing!'); process.exit(1); }

const bot = new TelegramBot(TOKEN, { polling: true });
const userState = {};

const mainMenu = {
  reply_markup: {
    keyboard: [
      [{ text: '📄 PDF → Excel' }, { text: '📝 PDF → Word' }],
      [{ text: '📋 PDF → CSV' }, { text: '📃 PDF → Text' }],
      [{ text: '❓ របៀបប្រើ' }, { text: '📞 ទាក់ទង' }]
    ],
    resize_keyboard: true,
    persistent: true
  }
};

bot.onText(/\/start/, (msg) => {
  const name = msg.from.first_name || 'បង';
  bot.sendMessage(msg.chat.id,
    `👋 សួស្ដី *${name}*!\n\n🤖 ខ្ញុំជា *PDF Converter Bot*\nជ្រើសប្រភេទការបម្លែងពី Menu ខាងក្រោម 👇`,
    { parse_mode: 'Markdown', ...mainMenu }
  );
});

bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text;
  if (!text) return;

  const formatMap = {
    '📄 PDF → Excel': 'xlsx',
    '📝 PDF → Word': 'docx',
    '📋 PDF → CSV': 'csv',
    '📃 PDF → Text': 'txt',
  };

  if (formatMap[text]) {
    const fmt = formatMap[text];
    userState[chatId] = { step: 'waiting_pdf', format: fmt };
    const labels = { xlsx: 'Excel (.xlsx)', docx: 'Word (.docx)', csv: 'CSV (.csv)', txt: 'Text (.txt)' };
    return bot.sendMessage(chatId,
      `✅ ជ្រើស *${labels[fmt]}* ហើយ!\n\n📤 ឥឡូវ *ផ្ញើ PDF file* មកខ្ញុំ`,
      { parse_mode: 'Markdown', reply_markup: { keyboard: [[{ text: '🔙 ត្រឡប់ Menu' }]], resize_keyboard: true } }
    );
  }

  if (text === '🔙 ត្រឡប់ Menu') {
    delete userState[chatId];
    return bot.sendMessage(chatId, '🏠 *Menu ចម្បង*', { parse_mode: 'Markdown', ...mainMenu });
  }

  if (text === '❓ របៀបប្រើ') {
    return bot.sendMessage(chatId,
      `📖 *របៀបប្រើ PDF Converter Bot*\n\n1️⃣ ចុច format ដែលចង់បាន\n2️⃣ ផ្ញើ PDF file\n3️⃣ ទទួល file ✅\n\n⚠️ *ដែនកំណត់:*\n• PDF ≤ 20MB\n• PDF ត្រូវមាន text`,
      { parse_mode: 'Markdown', ...mainMenu }
    );
  }

  if (text === '📞 ទាក់ទង') {
    return bot.sendMessage(chatId,
      `📞 *ទាក់ទងអ្នកអភិវឌ្ឍន៍*\n\n💬 Telegram: @your_username`,
      { parse_mode: 'Markdown', ...mainMenu }
    );
  }
});

bot.on('document', async (msg) => {
  const chatId = msg.chat.id;
  const doc = msg.document;
  const state = userState[chatId];

  if (!state || state.step !== 'waiting_pdf') {
    return bot.sendMessage(chatId, `⚠️ សូមជ្រើស *format* ពី Menu ជាមុនសិន!`, { parse_mode: 'Markdown', ...mainMenu });
  }
  if (doc.mime_type !== 'application/pdf') {
    return bot.sendMessage(chatId, `⚠️ ផ្ញើ *PDF file* ប៉ុណ្ណោះ!`, { parse_mode: 'Markdown' });
  }
  if (doc.file_size > 20 * 1024 * 1024) {
    return bot.sendMessage(chatId, `⚠️ File ធំពេក! Max *20MB*`, { parse_mode: 'Markdown' });
  }

  const fmt = state.format;
  const fmtLabel = { xlsx: 'Excel', docx: 'Word', csv: 'CSV', txt: 'Text' }[fmt];

  const statusMsg = await bot.sendMessage(chatId,
    `⏳ *កំពុងបម្លែង PDF → ${fmtLabel}...*\nសូមរង់ចាំ...`,
    { parse_mode: 'Markdown' }
  );

  try {
    const fileLink = await bot.getFileLink(doc.file_id);
    const resp = await axios.get(fileLink, { responseType: 'arraybuffer' });
    const pdfBuffer = new Uint8Array(resp.data);

    // Extract text using pdfjs-dist
    const lines = await extractPdfLines(pdfBuffer);

    if (!lines || lines.length === 0) {
      await bot.editMessageText(`❌ មិនអាចអានអត្ថបទពី PDF នេះ!\n_PDF នេះអាចជា scanned image_`,
        { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' });
      delete userState[chatId];
      return bot.sendMessage(chatId, '🏠 ត្រឡប់ Menu', mainMenu);
    }

    const baseName = (doc.file_name || 'document').replace(/\.pdf$/i, '');
    let outputPath, outName;

    if (fmt === 'xlsx') {
      outName = baseName + '.xlsx';
      outputPath = path.join(os.tmpdir(), outName);
      await makeXlsx(lines, outputPath);
    } else if (fmt === 'docx') {
      outName = baseName + '.docx';
      outputPath = path.join(os.tmpdir(), outName);
      makeDocx(lines, outputPath);
    } else if (fmt === 'csv') {
      outName = baseName + '.csv';
      outputPath = path.join(os.tmpdir(), outName);
      fs.writeFileSync(outputPath, makeCsv(lines), 'utf8');
    } else {
      outName = baseName + '.txt';
      outputPath = path.join(os.tmpdir(), outName);
      fs.writeFileSync(outputPath, lines.join('\n'), 'utf8');
    }

    await bot.sendDocument(chatId, outputPath, {
      caption: `✅ *បម្លែង PDF → ${fmtLabel} ជោគជ័យ!*\n📄 ${outName}`,
      parse_mode: 'Markdown'
    });

    fs.unlinkSync(outputPath);
    delete userState[chatId];

    await bot.editMessageText(`✅ *រួចរាល់!* PDF → ${fmtLabel}`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' });

    bot.sendMessage(chatId, '🏠 ចង់បម្លែង PDF ទៀត? ជ្រើសពី Menu 👇', mainMenu);

  } catch (err) {
    console.error('Error:', err);
    bot.editMessageText(`❌ *មានបញ្ហា:* ${err.message}`,
      { chat_id: chatId, message_id: statusMsg.message_id, parse_mode: 'Markdown' });
    delete userState[chatId];
    bot.sendMessage(chatId, '🏠 ត្រឡប់ Menu', mainMenu);
  }
});

// ===== Extract lines from PDF using pdfjs-dist =====
async function extractPdfLines(data) {
  const pdf = await pdfjsLib.getDocument({ data, useWorkerFetch: false, isEvalSupported: false }).promise;
  const allLines = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const tc = await page.getTextContent();
    // Group items by Y position into lines
    const yMap = {};
    tc.items.forEach(item => {
      const y = Math.round(item.transform[5]);
      if (!yMap[y]) yMap[y] = [];
      yMap[y].push(item.str);
    });
    const sortedYs = Object.keys(yMap).map(Number).sort((a, b) => b - a);
    sortedYs.forEach(y => {
      const line = yMap[y].join(' ').trim();
      if (line) allLines.push(line);
    });
    if (i < pdf.numPages) allLines.push('');
  }
  return allLines;
}

// ===== Excel using ExcelJS =====
async function makeXlsx(lines, outputPath) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  lines.forEach(line => {
    const cells = line.split(/\t|\s{3,}/);
    sheet.addRow(cells.length > 1 ? cells : [line]);
  });
  await workbook.xlsx.writeFile(outputPath);
}

// ===== Word =====
function makeDocx(lines, outputPath) {
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const paras = lines.map(l => {
    if (!l.trim()) return '<w:p/>';
    return `<w:p><w:r><w:t xml:space="preserve">${esc(l)}</w:t></w:r></w:p>`;
  }).join('');
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
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt32LE(0, 6);
    lh.writeUInt32LE(0, 10); lh.writeUInt32LE(0, 14);
    lh.writeUInt32LE(data.length, 18); lh.writeUInt32LE(data.length, 22);
    lh.writeUInt16LE(nb.length, 26); lh.writeUInt16LE(0, 28); nb.copy(lh, 30);
    parts.push(lh); parts.push(data);
    const cd = Buffer.alloc(46 + nb.length);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt32LE(0x00140014, 4);
    cd.writeUInt32LE(0, 8); cd.writeUInt32LE(0, 12); cd.writeUInt32LE(0, 16);
    cd.writeUInt32LE(data.length, 20); cd.writeUInt32LE(data.length, 24);
    cd.writeUInt16LE(nb.length, 28); cd.writeUInt32LE(0, 30); cd.writeUInt32LE(0, 34);
    cd.writeUInt32LE(0, 38); cd.writeUInt32LE(offset, 42); nb.copy(cd, 46);
    offset += lh.length + data.length; cds.push(cd);
  });
  const cdBuf = Buffer.concat(cds);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt32LE(0, 4);
  eocd.writeUInt16LE(Object.keys(files).length, 8); eocd.writeUInt16LE(Object.keys(files).length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12); eocd.writeUInt32LE(offset, 16); eocd.writeUInt16LE(0, 20);
  return Buffer.concat([...parts, cdBuf, eocd]);
}

function makeCsv(lines) {
  return lines.map(l => {
    const c = l.split(/\t|\s{3,}/);
    if (c.length > 1) return c.map(x => '"' + x.replace(/"/g, '""') + '"').join(',');
    return '"' + l.replace(/"/g, '""') + '"';
  }).join('\r\n');
}

bot.on('polling_error', err => console.error('Polling error:', err.message));

// ===== Keep-Alive Web Server (for UptimeRobot) =====
const http = require('http');
const PORT = process.env.PORT || 3000;
http.createServer((req, res) => {
  res.writeHead(200, { 'Content-Type': 'text/plain' });
  res.end('🤖 PDF Converter Bot is running!');
}).listen(PORT, () => {
  console.log(`✅ Keep-alive server running on port ${PORT}`);
});

console.log('🤖 PDF Converter Bot running...');
