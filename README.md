# 🤖 PDF Converter Telegram Bot

Bot Telegram សម្រាប់បម្លែង PDF → Excel / Word / CSV / Text

---

## 📋 Menu របស់ Bot

```
[ 📄 PDF → Excel ]  [ 📝 PDF → Word ]
[ 📋 PDF → CSV   ]  [ 📃 PDF → Text ]
[ ❓ របៀបប្រើ   ]  [ 📞 ទាក់ទង    ]
```

---

## 🚀 ជំហានដំឡើង

### ជំហាន 1: បង្កើត Bot Token

1. បើក Telegram → ស្វែងរក **@BotFather**
2. ផ្ញើ `/newbot`
3. ដាក់ឈ្មោះ Bot (ឧ: `PDF Converter Bot`)
4. ដាក់ username (ឧ: `my_pdf_converter_bot`)
5. BotFather នឹងផ្ញើ **Token** មកអ្នក → **Copy token នោះ**

### ជំហាន 2: Deploy លើ Railway (ឥតគិតថ្លៃ)

1. ទៅ **https://railway.app** → Sign up ជាមួយ GitHub
2. ចុច **"New Project"** → **"Deploy from GitHub repo"**
3. Upload folder `pdf-converter-bot` នេះទៅ GitHub ជាមុន:
   ```bash
   git init
   git add .
   git commit -m "PDF Converter Bot"
   git push origin main
   ```
4. ក្នុង Railway → ជ្រើស repo → Deploy

### ជំហាន 3: បន្ថែម BOT_TOKEN

1. ក្នុង Railway project → ចុច **"Variables"**
2. ចុច **"New Variable"**
3. Name: `BOT_TOKEN`
4. Value: (paste token ពី BotFather)
5. ចុច **Add** → Bot ចាប់ផ្ដើមដំណើរការ!

---

## 💻 Run លើ Computer ផ្ទាល់ (optional)

```bash
# ដំឡើង dependencies
npm install

# Set token
export BOT_TOKEN="your_token_here"   # Mac/Linux
set BOT_TOKEN=your_token_here        # Windows

# Start bot
npm start
```

---

## 📁 Project Structure

```
pdf-converter-bot/
├── src/
│   └── bot.js          # Bot code ចម្បង
├── package.json
├── railway.toml        # Railway config
└── README.md
```

---

## ⚠️ ដែនកំណត់

- PDF ទំហំអតិបរិមា **20MB**
- PDF ត្រូវតែមាន **text** (មិនដំណើរការ scanned image)
- Free tier Railway: **500 hours/month**

---

## 🛠 Tech Stack

- **Node.js** + node-telegram-bot-api
- **pdf-parse** — អានអត្ថបទពី PDF
- **xlsx** — បង្កើត Excel file
- **Railway** — Free hosting
