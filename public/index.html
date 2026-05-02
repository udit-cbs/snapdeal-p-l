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

function sheetToRows(wb, hints) {
  for (const name of wb.SheetNames) {
    if (hints.some(h => name.toLowerCase().includes(h)))
      return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
  }
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
}

function n(v) { return parseFloat(v) || 0; }

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

function normSub(v) {
  if (!v || String(v).trim() === '' || String(v) === 'nan') return '';
  const s = String(v).trim();
  if (s.includes('e') || s.includes('.')) {
    try { return String(Math.round(parseFloat(s))); } catch(e) { return s; }
  }
  return s;
}

app.post('/api/analyze',
  upload.fields([{ name: 'pl', maxCount: 1 }, { name: 'orders', maxCount: 1 }]),
  (req, res) => {
    try {
      if (!req.files?.orders) return res.status(400).json({ error: 'Order report file is required.' });

      const costs = {
        pad:   n(req.body.c_pad)   || 6,
        liner: n(req.body.c_liner) || 3,
        panty: n(req.body.c_panty) || 12,
        pkg:   n(req.body.c_pkg)   || 10,
        misc:  n(req.body.c_misc)  || 2,
      };

      const skuConfigs = JSON.parse(req.body.skuConfigs || '[]');
      const cfgMap = {};
      skuConfigs.forEach(c => { cfgMap[c.sku] = c; });

      function cogsForSku(sku, productName) {
        const cfg  = cfgMap[sku];
        const type = cfg ? cfg.type : inferType(sku, productName);
        const qty  = cfg ? cfg.qty  : inferQty(sku);
        const unit = type === 'pad' ? costs.pad : type === 'liner' ? costs.liner : costs.panty;
        return { cogs: (unit * qty) + costs.pkg + costs.misc, type, qty };
      }

      const wb2 = XLSX.read(req.files.orders[0].buffer, { type: 'buffer' });
      const wb1 = req.files.pl ? XLSX.read(req.files.pl[0].buffer, { type: 'buffer' }) : null;

      const orderRows    = sheetToRows(wb2, ['order', 'consolidate', 'comp']);
      const chargesRows  = wb1 ? sheetToRows(wb1, ['commission', 'charges', 'marketing']) : [];
      const nonOrderRows = wb1 ? sheetToRows(wb1, ['non order', 'non-order', 'nonorder', 'non_order']) : [];

      // ── Pass 1: build SKU map ──────────────────────────────────────────────
      const skuMap = {}, invToSku = {}, subToSku = {}, subToInvAmt = {};

      for (const o of orderRows) {
        const sku    = String(o['SKU CODE'] || '').trim();
        const inv    = String(o['INVOICE NUMBER'] || '').trim();
        const sub    = normSub(o['SUBORDER CODE']);
        const state  = String(o['CURRENT ORDER STATE'] || '').toLowerCase();
        const sp     = n(o['SELLING PRICE']);
        const invAmt = n(o['SELLER INVOICE AMOUNT']);
        const name   = String(o['PRODUCT NAME'] || '');
        const attr   = String(o['ATTRIBUTES'] || '');
        if (!sku) continue;
        if (inv) invToSku[inv] = sku;
        if (sub) { subToSku[sub] = sku; subToInvAmt[sub] = invAmt; }
        if (!skuMap[sku]) {
          const { cogs, type, qty } = cogsForSku(sku, name);
          skuMap[sku] = {
            sku, productName: name, attr, sp, type, qty,
            orders: 0, returned: 0,
            unitCOGS: cogs, totalCOGS: 0,
            // Stage 1 components
            totalInvoiceAmt: 0,
            totalMarketingFee: 0,
            totalCourierFee: 0,
            totalPaymentFee: 0,
            totalIGST: 0,
            totalWebAds: 0,
            grossPayableDelivered: 0,
            totalReturnReversal: 0,
            // Stage 2
            totalTDS: 0,
            totalTCS: 0,
            // zone
            remoteOrders: 0, standardOrders: 0,
          };
        }
        const s = skuMap[sku];
        const isRet = state.includes('return');
        const isCan = state.includes('cancel');
        if (isRet)       { s.returned++; }
        else if (!isCan) { s.orders++; s.totalCOGS += s.unitCOGS; }
      }

      // ── Pass 2: charges sheet ──────────────────────────────────────────────
      let totalAdSpend = 0;

      for (const c of chargesRows) {
        const tx    = String(c['Transaction Type'] || '').toLowerCase();
        const sub   = normSub(c['Sub Order No']);
        const inv   = String(c['Invoice Number'] || '').trim();
        const total = n(c['Total Commission Amount']);
        const mkt   = n(c['Marketing Fee']);
        const cour  = n(c['Courier Fee']);
        const pmt   = n(c['Payment Collection Fee']);
        const igst  = n(c['Igst'] || c['IGST'] || 0);

        // Bulk ad spend — account level, no SKU
        if (tx.includes('advertise') || tx.includes('ads income')) {
          totalAdSpend += Math.abs(total); continue;
        }

        const sku = subToSku[sub] || invToSku[inv];

        // Delivered order invoice
        if (tx.includes('vendor invoice')) {
          if (!sku || !skuMap[sku]) continue;
          const s = skuMap[sku];
          const invAmt = subToInvAmt[sub] || 0;
          s.totalInvoiceAmt   += invAmt;
          s.totalMarketingFee += mkt;
          s.totalCourierFee   += cour;
          s.totalPaymentFee   += pmt;
          s.totalIGST         += igst;
          s.grossPayableDelivered += (invAmt + total);
          if (Math.abs(cour) > 80) s.remoteOrders++;
          else                     s.standardOrders++;
        }

        // Per-order web ads & stock-out charges
        if (tx.includes('web ads') || tx.includes('stock out')) {
          if (!sku || !skuMap[sku]) continue;
          skuMap[sku].totalWebAds         += total;
          skuMap[sku].grossPayableDelivered += total;
        }

        // Return reversal
        if (tx.includes('return to vendor')) {
          if (!sku || !skuMap[sku]) continue;
          skuMap[sku].totalReturnReversal += Math.abs(total);
        }
      }

      // ── Pass 3: Non-Order Transactions → TDS/TCS ──────────────────────────
      for (const r of nonOrderRows) {
        const tx  = String(r['Transaction Type'] || '').toLowerCase();
        const sub = normSub(r['Sub Order No']);
        const amt = n(r['Gross Amount']);
        if (!tx.includes('tds') && !tx.includes('tcs')) continue;
        const sku = subToSku[sub];
        if (!sku || !skuMap[sku]) continue;
        if (tx.includes('tcs')) skuMap[sku].totalTCS += amt;
        else                    skuMap[sku].totalTDS += amt;
      }

      // ── Allocate ad spend equally across SKUs ──────────────────────────────
      const skuList = Object.keys(skuMap);
      const adShare = skuList.length > 0 ? totalAdSpend / skuList.length : 0;

      // ── Final P&L per SKU ──────────────────────────────────────────────────
      const FALLBACK_GROSS = 91.71;

      const skus = Object.values(skuMap).map(s => {
        const hasActual = s.grossPayableDelivered > 0;

        // STAGE 1 — Gross Seller Payable
        const grossPayableDelivered = hasActual ? s.grossPayableDelivered : s.orders * FALLBACK_GROSS;
        const grossPayableNet       = grossPayableDelivered - s.totalReturnReversal;

        // STAGE 2 — Net Seller Payable (after TDS/TCS)
        const totalTaxDeductions = s.totalTDS + s.totalTCS;
        const netSellerPayable   = grossPayableNet + totalTaxDeductions;

        // STAGE 3 — Gross Profit (after advertising + COGS)
        const grossProfit = netSellerPayable - adShare - s.totalCOGS;

        // Derived metrics
        const totalAttempts  = s.orders + s.returned;
        const returnRate     = totalAttempts > 0 ? s.returned / totalAttempts * 100 : 0;
        const potentialRev   = s.sp * s.orders;
        const grossMarginPct = potentialRev > 0 ? grossProfit / potentialRev * 100 : 0;
        const profitPerUnit  = s.orders > 0 ? grossProfit / s.orders : 0;

        // Per-order averages
        const avgGrossPayable  = s.orders > 0 ? grossPayableDelivered / s.orders : FALLBACK_GROSS;
        const avgCutPerOrder   = s.sp - avgGrossPayable;
        const avgTDSPerOrder   = s.orders > 0 ? totalTaxDeductions / s.orders : 0;
        const avgNetPerOrder   = avgGrossPayable + avgTDSPerOrder;
        const avgMarketingFee  = s.orders > 0 ? s.totalMarketingFee / s.orders : 0;
        const avgCourierFee    = s.orders > 0 ? s.totalCourierFee   / s.orders : 0;
        const avgPaymentFee    = s.orders > 0 ? s.totalPaymentFee   / s.orders : 0;
        const avgIGST          = s.orders > 0 ? s.totalIGST         / s.orders : 0;
        const avgWebAds        = s.orders > 0 ? s.totalWebAds       / s.orders : 0;

        // Return loss per unit = reversal + COGS already spent dispatching
        const avgReturnReversal = s.returned > 0 ? s.totalReturnReversal / s.returned : avgGrossPayable;
        const returnLossPerUnit = avgReturnReversal + s.unitCOGS;
        const totalReturnLoss   = s.returned * returnLossPerUnit;

        return {
          sku: s.sku, productName: s.productName, attr: s.attr, sp: s.sp,
          type: s.type, qty: s.qty, unitCOGS: s.unitCOGS,
          orders: s.orders, returned: s.returned, totalAttempts,
          returnRate, potentialRev, hasActual,
          remoteOrders: s.remoteOrders, standardOrders: s.standardOrders,
          // Stage 1
          totalInvoiceAmt:      s.totalInvoiceAmt,
          totalMarketingFee:    s.totalMarketingFee,
          totalCourierFee:      s.totalCourierFee,
          totalPaymentFee:      s.totalPaymentFee,
          totalIGST:            s.totalIGST,
          totalWebAds:          s.totalWebAds,
          grossPayableDelivered, totalReturnReversal: s.totalReturnReversal, grossPayableNet,
          // Stage 2
          totalTDS:             s.totalTDS,
          totalTCS:             s.totalTCS,
          totalTaxDeductions, netSellerPayable,
          // Stage 3
          totalCOGS:            s.totalCOGS,
          adShare, grossProfit, grossMarginPct, profitPerUnit,
          // Return
          returnLossPerUnit, totalReturnLoss,
          // Per-order averages
          avgGrossPayable, avgCutPerOrder, avgTDSPerOrder, avgNetPerOrder,
          avgMarketingFee, avgCourierFee, avgPaymentFee, avgIGST, avgWebAds,
          // Aliases
          NET_PER_ORDER: avgNetPerOrder,
          totalSnapNet: netSellerPayable,
          totalSnapCut: s.orders * avgCutPerOrder,
        };
      });

      res.json({ skus, totalAdSpend, detectedSKUs: skuList.map(k => ({
        sku: skuMap[k].sku, productName: skuMap[k].productName,
        type: skuMap[k].type, qty: skuMap[k].qty,
      }))});

    } catch (err) {
      console.error(err);
      res.status(500).json({ error: err.message });
    }
  }
);

app.post('/api/detect-skus', upload.single('orders'), (req, res) => {
  try {
    const wb   = XLSX.read(req.file.buffer, { type: 'buffer' });
    const rows = sheetToRows(wb, ['order', 'consolidate', 'comp']);
    const seen = {}, skus = [];
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
