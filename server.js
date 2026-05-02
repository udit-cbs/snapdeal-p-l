const express = require('express');
const multer  = require('multer');
const XLSX    = require('xlsx');
const cors    = require('cors');
const path    = require('path');

const app    = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ── parse helpers ──────────────────────────────────────────────────────────
function sheetToRows(wb, hints) {
  for (const name of wb.SheetNames) {
    if (hints.some(h => name.toLowerCase().includes(h)))
      return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
  }
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
}

function inferType(sku, name) {
  const s = (sku + ' ' + name).toLowerCase();
  if (s.includes('liner')) return 'liner';
  if (s.includes('panty')) return 'panty';
  return 'pad';
}

function inferQty(sku) {
  const m = sku.match(/[_\-](\d+)$/);
  if (m) return parseInt(m[1]);
  const m2 = sku.match(/(\d+)$/);
  if (m2) return parseInt(m2[1]);
  return 1;
}

// ── /api/analyze ───────────────────────────────────────────────────────────
app.post('/api/analyze',
  upload.fields([{ name: 'pl', maxCount: 1 }, { name: 'orders', maxCount: 1 }]),
  (req, res) => {
    try {
      if (!req.files?.orders) return res.status(400).json({ error: 'Order report file is required.' });

      const costs = {
        pad:   parseFloat(req.body.c_pad)   || 6,
        liner: parseFloat(req.body.c_liner) || 3,
        panty: parseFloat(req.body.c_panty) || 12,
        pkg:   parseFloat(req.body.c_pkg)   || 10,
        misc:  parseFloat(req.body.c_misc)  || 2,
      };

      const skuConfigs = JSON.parse(req.body.skuConfigs || '[]');

      const wb2 = XLSX.read(req.files.orders[0].buffer, { type: 'buffer' });
      const wb1 = req.files.pl ? XLSX.read(req.files.pl[0].buffer, { type: 'buffer' }) : null;

      const orderRows   = sheetToRows(wb2, ['order', 'consolidate', 'comp']);
      const chargesRows = wb1 ? sheetToRows(wb1, ['commission', 'charges', 'marketing']) : [];

      // Build SKU config map
      const cfgMap = {};
      skuConfigs.forEach(c => { cfgMap[c.sku] = c; });

      function cogsForSku(sku, productName) {
        const cfg = cfgMap[sku];
        const type = cfg ? cfg.type : inferType(sku, productName);
        const qty  = cfg ? cfg.qty  : inferQty(sku);
        const unit = type === 'pad' ? costs.pad : type === 'liner' ? costs.liner : costs.panty;
        return { cogs: (unit * qty) + costs.pkg + costs.misc, type, qty, unit };
      }

      const FALLBACK_NET = 110.50;
      const skuMap = {};
      const invToSku = {}, subToSku = {}, subToInvAmt = {};

      // Pass 1: build SKU map from order report
      for (const o of orderRows) {
        const sku   = String(o['SKU CODE'] || '').trim();
        const inv   = String(o['INVOICE NUMBER'] || '').trim();
        const sub   = String(o['SUBORDER CODE'] || '').trim();
        const state = String(o['CURRENT ORDER STATE'] || '').toLowerCase();
        const sp    = parseFloat(o['SELLING PRICE']) || 0;
        const invAmt = parseFloat(o['SELLER INVOICE AMOUNT']) || 0;
        const name  = String(o['PRODUCT NAME'] || '');
        const attr  = String(o['ATTRIBUTES'] || '');
        if (!sku) continue;
        if (inv) invToSku[inv] = sku;
        if (sub) { subToSku[sub] = sku; subToInvAmt[sub] = invAmt; }
        if (!skuMap[sku]) {
          const { cogs, type, qty } = cogsForSku(sku, name);
          skuMap[sku] = { sku, productName: name, attr, sp, type, qty,
            orders: 0, returned: 0, totalCOGS: 0, unitCOGS: cogs,
            totalNetDelivered: 0, totalNetReturned: 0, totalSnapCut: 0,
            remoteOrders: 0, standardOrders: 0 };
        }
        const s = skuMap[sku];
        const isRet = state.includes('return');
        const isCan = state.includes('cancel');
        if (isRet)       { s.returned++; }
        else if (!isCan) { s.orders++; s.totalCOGS += s.unitCOGS; }
      }

      // Pass 2: calculate exact net per order from charges sheet
      let totalAdSpend = 0;
      for (const c of chargesRows) {
        const tx    = String(c['Transaction Type'] || '').toLowerCase();
        const sub   = String(c['Sub Order No'] || '').trim();
        const inv   = String(c['Invoice Number'] || '').trim();
        const total = parseFloat(c['Total Commission Amount'] || 0);
        const courier = Math.abs(parseFloat(c['Courier Fee'] || 0));

        if (tx.includes('advertise') || tx.includes('ads income')) {
          totalAdSpend += Math.abs(total); continue;
        }

        // Delivered invoices: net = invoice_amount + total_commission
        if (tx.includes('vendor invoice')) {
          const invAmt = subToInvAmt[sub] || 0;
          const net    = invAmt + total;
          const sku    = subToSku[sub] || invToSku[inv];
          if (sku && skuMap[sku]) {
            skuMap[sku].totalNetDelivered += net;
            skuMap[sku].totalSnapCut      += (skuMap[sku].sp - net);
            if (courier > 80) skuMap[sku].remoteOrders++;
            else              skuMap[sku].standardOrders++;
          }
        }

        // Returns: Snapdeal reverses the net amount
        if (tx.includes('return to vendor')) {
          const sku = subToSku[sub] || invToSku[inv];
          if (sku && skuMap[sku]) skuMap[sku].totalNetReturned += Math.abs(total);
        }
      }

      const skuList = Object.keys(skuMap);
      const adShare = skuList.length > 0 ? totalAdSpend / skuList.length : 0;

      const detectedSKUs = skuList.map(k => ({
        sku: skuMap[k].sku, productName: skuMap[k].productName,
        type: skuMap[k].type, qty: skuMap[k].qty,
      }));

      const skus = Object.values(skuMap).map(s => {
        const hasActual       = s.totalNetDelivered > 0;
        const avgNetPerOrder  = hasActual && s.orders > 0 ? s.totalNetDelivered / s.orders : FALLBACK_NET;
        const avgCutPerOrder  = s.sp - avgNetPerOrder;
        const netDelivered    = hasActual ? s.totalNetDelivered : s.orders * FALLBACK_NET;
        const netReturned     = s.totalNetReturned > 0 ? s.totalNetReturned : s.returned * avgNetPerOrder;
        const totalSnapNet    = netDelivered - netReturned;
        const totalSnapCut    = hasActual ? s.totalSnapCut : s.orders * avgCutPerOrder;
        const grossProfit     = totalSnapNet - adShare - s.totalCOGS;
        const totalAttempts   = s.orders + s.returned;
        const returnRate      = totalAttempts > 0 ? s.returned / totalAttempts * 100 : 0;
        const potentialRev    = s.sp * s.orders;
        const grossMarginPct  = potentialRev > 0 ? grossProfit / potentialRev * 100 : 0;
        const profitPerUnit   = s.orders > 0 ? grossProfit / s.orders : 0;
        const returnLoss      = avgNetPerOrder + s.unitCOGS;
        const totalReturnLoss = s.returned * returnLoss;
        return {
          ...s, avgNetPerOrder, avgCutPerOrder, netDelivered, netReturned,
          totalSnapNet, totalSnapCut, adShare, grossProfit, totalAttempts,
          returnRate, potentialRev, grossMarginPct, profitPerUnit,
          returnLoss, totalReturnLoss, NET_PER_ORDER: avgNetPerOrder, hasActual,
        };
      });

      res.json({ skus, totalAdSpend, detectedSKUs, NET_PER_ORDER: FALLBACK_NET });

    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message });
    }
  }
);

// ── /api/detect-skus ───────────────────────────────────────────────────────
app.post('/api/detect-skus', upload.single('orders'), (req, res) => {
  try {
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const rows = sheetToRows(wb, ['order', 'consolidate', 'comp']);
    const seen = {};
    const skus = [];
    for (const o of rows) {
      const sku  = String(o['SKU CODE'] || '').trim();
      const name = String(o['PRODUCT NAME'] || '').trim();
      if (!sku || seen[sku]) continue;
      seen[sku] = true;
      skus.push({ sku, productName: name, type: inferType(sku, name), qty: inferQty(sku) });
    }
    res.json({ skus });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Snapdeal P&L Analyzer running on http://localhost:${PORT}`));
