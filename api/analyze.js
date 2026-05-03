const XLSX   = require('xlsx');
const busboy = require('busboy');

// ── CRITICAL: tell Vercel NOT to parse the body — busboy does it ──────────
module.exports.config = { api: { bodyParser: false } };

// ── helpers ────────────────────────────────────────────────────────────────
function n(v) { return parseFloat(v) || 0; }

function normSub(v) {
  if (!v || String(v).trim() === '' || String(v) === 'nan') return '';
  const s = String(v).trim();
  if (s.includes('e') || s.includes('.')) {
    try { return String(Math.round(parseFloat(s))); } catch(e) { return s; }
  }
  return s;
}

function sheetToRows(wb, hints) {
  for (const name of wb.SheetNames) {
    if (hints.some(h => name.toLowerCase().includes(h)))
      return XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' });
  }
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
}

function inferType(sku, name) {
  const s = (sku+' '+name).toLowerCase();
  if (s.includes('liner')) return 'liner';
  if (s.includes('panty')) return 'panty';
  return 'pad';
}

function inferQty(sku) {
  const m = sku.match(/[_\-](\d+)$/);  if (m) return parseInt(m[1]);
  const m2 = sku.match(/(\d+)$/);       if (m2) return parseInt(m2[1]);
  return 1;
}

// ── parse multipart (returns { fields, files }) ────────────────────────────
function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const bb = busboy({ headers: req.headers, limits: { fileSize: 20*1024*1024 } });
    const fields = {}, files = {};

    bb.on('field', (name, val) => { fields[name] = val; });

    bb.on('file', (name, stream) => {
      const chunks = [];
      stream.on('data', d => chunks.push(d));
      stream.on('end', () => { files[name] = Buffer.concat(chunks); });
    });

    bb.on('finish', () => resolve({ fields, files }));
    bb.on('error',  reject);
    req.pipe(bb);
  });
}

// ── main handler ───────────────────────────────────────────────────────────
async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'POST')    { res.status(405).json({ error: 'POST only' }); return; }

  try {
    const { fields, files } = await parseMultipart(req);

    if (!files.orders) return res.status(400).json({ error: 'Order report file is required.' });

    const costs = {
      pad:   n(fields.c_pad)   || 6,
      liner: n(fields.c_liner) || 3,
      panty: n(fields.c_panty) || 12,
      pkg:   n(fields.c_pkg)   || 10,
      misc:  n(fields.c_misc)  || 2,
    };

    const skuConfigs = JSON.parse(fields.skuConfigs || '[]');
    const cfgMap = {};
    skuConfigs.forEach(c => { cfgMap[c.sku] = c; });

    function cogsForSku(sku, productName) {
      const cfg  = cfgMap[sku];
      const type = cfg ? cfg.type : inferType(sku, productName);
      const qty  = cfg ? cfg.qty  : inferQty(sku);
      const unit = type === 'pad' ? costs.pad : type === 'liner' ? costs.liner : costs.panty;
      return { cogs: (unit * qty) + costs.pkg + costs.misc, type, qty };
    }

    const wb2 = XLSX.read(files.orders, { type: 'buffer' });
    const wb1 = files.pl ? XLSX.read(files.pl, { type: 'buffer' }) : null;

    const orderRows    = sheetToRows(wb2, ['order','consolidate','comp']);
    const chargesRows  = wb1 ? sheetToRows(wb1, ['commission','charges','marketing']) : [];
    const nonOrderRows = wb1 ? sheetToRows(wb1, ['non order','non-order','nonorder','non_order']) : [];

    // ── Pass 1: SKU map from order report ──────────────────────────────────
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
          totalInvoiceAmt: 0,
          totalMarketingFee: 0, totalCourierFee: 0,
          totalPaymentFee: 0,   totalIGST: 0,
          totalWebAds: 0,
          grossPayableDelivered: 0,
          totalReturnReversal: 0,
          totalTDS: 0, totalTCS: 0,
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

      if (tx.includes('advertise') || tx.includes('ads income')) {
        totalAdSpend += Math.abs(total); continue;
      }

      const sku = subToSku[sub] || invToSku[inv];

      if (tx.includes('vendor invoice')) {
        if (!sku || !skuMap[sku]) continue;
        const s      = skuMap[sku];
        const invAmt = subToInvAmt[sub] || 0;
        s.totalInvoiceAmt       += invAmt;
        s.totalMarketingFee     += mkt;
        s.totalCourierFee       += cour;
        s.totalPaymentFee       += pmt;
        s.totalIGST             += igst;
        s.grossPayableDelivered += (invAmt + total);
        if (Math.abs(cour) > 80) s.remoteOrders++;
        else                     s.standardOrders++;
      }

      if (tx.includes('web ads') || tx.includes('stock out')) {
        if (!sku || !skuMap[sku]) continue;
        skuMap[sku].totalWebAds           += total;
        skuMap[sku].grossPayableDelivered += total;
      }

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

    // ── Allocate ad spend equally ──────────────────────────────────────────
    const skuList = Object.keys(skuMap);
    const adShare = skuList.length > 0 ? totalAdSpend / skuList.length : 0;

    // ── Final P&L per SKU ──────────────────────────────────────────────────
    const FALLBACK_GROSS = 91.71;

    const skus = Object.values(skuMap).map(s => {
      const hasActual = s.grossPayableDelivered > 0;

      const grossPayableDelivered = hasActual ? s.grossPayableDelivered : s.orders * FALLBACK_GROSS;
      const grossPayableNet       = grossPayableDelivered - s.totalReturnReversal;

      const totalTaxDeductions = s.totalTDS + s.totalTCS;
      const netSellerPayable   = grossPayableNet + totalTaxDeductions;

      const grossProfit = netSellerPayable - adShare - s.totalCOGS;

      const totalAttempts  = s.orders + s.returned;
      const returnRate     = totalAttempts > 0 ? s.returned / totalAttempts * 100 : 0;
      const potentialRev   = s.sp * s.orders;
      const grossMarginPct = potentialRev > 0 ? grossProfit / potentialRev * 100 : 0;
      const profitPerUnit  = s.orders > 0 ? grossProfit / s.orders : 0;

      const avgGrossPayable = s.orders > 0 ? grossPayableDelivered / s.orders : FALLBACK_GROSS;
      const avgCutPerOrder  = s.sp - avgGrossPayable;
      const avgTDSPerOrder  = s.orders > 0 ? totalTaxDeductions / s.orders : 0;
      const avgNetPerOrder  = avgGrossPayable + avgTDSPerOrder;
      const avgMarketingFee = s.orders > 0 ? s.totalMarketingFee / s.orders : 0;
      const avgCourierFee   = s.orders > 0 ? s.totalCourierFee   / s.orders : 0;
      const avgPaymentFee   = s.orders > 0 ? s.totalPaymentFee   / s.orders : 0;
      const avgIGST         = s.orders > 0 ? s.totalIGST         / s.orders : 0;
      const avgWebAds       = s.orders > 0 ? s.totalWebAds       / s.orders : 0;

      const avgReturnReversal = s.returned > 0 ? s.totalReturnReversal / s.returned : avgGrossPayable;
      const returnLossPerUnit = avgReturnReversal + s.unitCOGS;
      const totalReturnLoss   = s.returned * returnLossPerUnit;

      return {
        sku: s.sku, productName: s.productName, attr: s.attr, sp: s.sp,
        type: s.type, qty: s.qty, unitCOGS: s.unitCOGS,
        orders: s.orders, returned: s.returned, totalAttempts,
        returnRate, potentialRev, hasActual,
        remoteOrders: s.remoteOrders, standardOrders: s.standardOrders,
        totalInvoiceAmt:    s.totalInvoiceAmt,
        totalMarketingFee:  s.totalMarketingFee,
        totalCourierFee:    s.totalCourierFee,
        totalPaymentFee:    s.totalPaymentFee,
        totalIGST:          s.totalIGST,
        totalWebAds:        s.totalWebAds,
        grossPayableDelivered, totalReturnReversal: s.totalReturnReversal, grossPayableNet,
        totalTDS: s.totalTDS, totalTCS: s.totalTCS, totalTaxDeductions, netSellerPayable,
        totalCOGS: s.totalCOGS, adShare, grossProfit, grossMarginPct, profitPerUnit,
        returnLossPerUnit, totalReturnLoss,
        avgGrossPayable, avgCutPerOrder, avgTDSPerOrder, avgNetPerOrder,
        avgMarketingFee, avgCourierFee, avgPaymentFee, avgIGST, avgWebAds,
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

module.exports = handler;
module.exports.config = { api: { bodyParser: false } };
