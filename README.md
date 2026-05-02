Snapdeal P&L Analyzer — Flawsome
A deployable web tool for SKU-level P&L analysis on Snapdeal.
What it does
Upload your monthly Snapdeal P&L sheet + Consolidated Order Report
Automatically detects SKUs, pack sizes, and product types
Calculates true net-to-seller (verified ₹110.50/order for ₹199 SP)
Shows per-SKU P&L with every Snapdeal charge explained
Insights: which SKUs to kill, review, or scale
Tech stack
Backend: Node.js + Express
File parsing: xlsx (SheetJS)
Frontend: Vanilla HTML/CSS/JS (no framework needed)
---
Run locally
```bash
npm install
npm start
# Open http://localhost:3000
```
---
Deploy to Railway (recommended — free tier)
Push this folder to a GitHub repo
Go to https://railway.app → New Project → Deploy from GitHub
Select your repo → Railway auto-detects Node.js
Set start command: `node server.js`
Deploy — you get a public URL instantly
---
Deploy to Render (free tier)
Push to GitHub
Go to https://render.com → New Web Service
Connect repo, set:
Build command: `npm install`
Start command: `node server.js`
Deploy
---
Deploy to Vercel
Not recommended for this app (uses file uploads via multer which needs a persistent server).
Use Railway or Render instead.
---
Deploy to a VPS (DigitalOcean/AWS/GCP)
```bash
# On your server
git clone <your-repo>
cd snapdeal-pl-tool
npm install
npm install -g pm2
pm2 start server.js --name snapdeal-pl
pm2 startup
pm2 save

# With nginx proxy (optional)
# Point nginx to localhost:3000
```
---
Environment variables
Variable	Default	Description
`PORT`	`3000`	Server port
---
File structure
```
snapdeal-pl-tool/
├── server.js          # Express API + file parsing
├── package.json
└── public/
    └── index.html     # Full frontend (single file)
```
---
How the P&L math works
Snapdeal's fee structure is opaque. Their system:
Credits you a variable invoice amount (₹168–₹199 for ₹199 SP) depending on promotions
Then separately charges courier, commission, payment fee, IGST in the charges sheet
Net result is always ₹110.50 per delivered order at ₹199 SP — verified across every suborder in March 2026 P&L
The tool uses this verified net figure instead of double-counting invoice + charges.
Gross profit per SKU = (Delivered × ₹110.50) − (Returned × ₹110.50) − Ad spend − COGS
