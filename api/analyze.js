const XLSX   = require('xlsx');
const busboy = require('busboy');

module.exports.config = { api: { bodyParser: false } };

function f(v) { return parseFloat(v) || 0; }
function str(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }

function sheetByHint(wb, hints) {
  for (const name of wb.SheetNames) {
    if (hints.some(h => name.toLowerCase().includes(h.toLowerCase())))
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
  const m  = sku.match(/[_\-](\d+)$/); if (m)  return parseInt(m[1]);
  const m2 = sku.match(/(\d+)$/);       if (m2) return parseInt(m2[1]);
  return 1;
}

function parseMultipart(req) {
  return new Promise((resolve, reject) => {
    const bb = busboy({ headers: req.headers, limits: { fileSize: 20 * 1024 * 1024 } });
    const fields = {}, files = {};
    bb.on('field', (name, val) => { fields[name] = val; });
    bb.on('file', (name, stream) => {
      const chunks = [];
      stream.on('data', d => chunks.push(d));
      stream.on('end', () => { files[name] = Buffer.concat(chunks); });
    });
    bb.on('finish', () => resolve({ fields, files }));
    bb.on('error', reject);
    req.pipe(bb);
  });
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'POST')   { res.status(405).json({ error: 'POST only' }); return; }

  try {
    const { fields, files } = await parseMultipart(req);
    if (!files.orders) return res.status(400).json({ error: 'Order Report file is required.' });

    const costs = {
      pad:   f(fields.c_pad)   || 6,
      liner: f(fields.c_liner) || 3,
      panty: f(fields.c_panty) || 12,
      pkg:   f(fields.c_pkg)   || 10,
      misc:  f(fields.c_misc)  || 2,
    };
    const skuConfigs = JSON.parse(fields.skuConfigs || '[]');
    const cfgMap = {};
    skuConfigs.forEach(c => { cfgMap[c.sku] = c; });

    function calcUnitCOGS(sku, productName) {
      const cfg  = cfgMap[sku];
      const type = cfg ? cfg.type : inferType(sku, productName);
      const qty  = cfg ? cfg.qty  : inferQty(sku);
      const unit = type === 'pad' ? costs.pad : type === 'liner' ? costs.liner : costs.panty;
      return { cogs: (unit * qty) + costs.pkg + costs.misc, type, qty };
    }

    // ── Read files ─────────────────────────────────────────────────────────
    const wbOrders = XLSX.read(files.orders, { type: 'buffer' });
    const wbPL     = files.pl ? XLSX.read(files.pl, { type: 'buffer' }) : null;

    const orderRows     = sheetByHint(wbOrders, ['consolidate', 'order', 'comp']);
    const totalSubRows  = wbPL ? sheetByHint(wbPL, ['total_sub', 'total sub'])            : [];
    const commRows      = wbPL ? sheetByHint(wbPL, ['commission and other', 'commission']) : [];
    const nonOrderRows  = wbPL ? sheetByHint(wbPL, ['non order', 'non_order'])             : [];
    const returnsRows   = wbPL ? sheetByHint(wbPL, ['returns'])                            : [];

    // ── Pass 1: sub→SKU from order report ──────────────────────────────────
    const subToSKU = {}, subToSP = {}, skuMeta = {};

    for (const o of orderRows) {
      const sub  = str(o['SUBORDER CODE']);
      const sku  = str(o['SKU CODE']);
      const name = str(o['PRODUCT NAME']);
      const sp   = f(o['SELLING PRICE']);
      if (!sku || !sub) continue;
      subToSKU[sub] = sku;
      subToSP[sub]  = sp;
      if (!skuMeta[sku]) {
        const { cogs, type, qty } = calcUnitCOGS(sku, name);
        skuMeta[sku] = { productName: name, attr: str(o['ATTRIBUTES']), type, qty, cogs, sp };
      }
    }

    // ── Pass 2: financial data from P&L sheets ─────────────────────────────
    // subData[sub] = per-suborder financial breakdown
    const subData = {};

    // A) Total_Suboders → invoice amount
    for (const r of totalSubRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;
      if (tx.includes('Vendor Invoice')) {
        subData[sub] = subData[sub] || {};
        subData[sub].invoiceAmt = f(r['Invoice Amount']);
      }
    }

    // B) Commission sheet
    let totalAdSpend = 0;
    for (const r of commRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;

      if (tx.includes('Advertise Ads Income') || tx.includes('Web Ads Invoice')) {
        totalAdSpend += f(r['Total Commission Amount']);
        continue;
      }
      if (tx.includes('Stock Out') || tx.includes('RTO Charges')) continue;

      subData[sub] = subData[sub] || {};
      if (tx.includes('COD Vendor Invoice')) {
        subData[sub].commTotal   = f(r['Total Commission Amount']);
        subData[sub].commMkt     = f(r['Marketing Fee']);
        subData[sub].commCourier = f(r['Courier Fee']);
        subData[sub].commPmt     = f(r['Payment Collection Fee']);
        subData[sub].commIGST    = f(r['Igst'] || r['IGST'] || 0);
      }
      if (tx.includes('COD Return to Vendor')) {
        subData[sub].returnComm    = f(r['Total Commission Amount']); // positive reversal
        subData[sub].returnCourier = f(r['Courier Fee']);
      }
    }

    // C) Non-Order Transactions → TDS
    for (const r of nonOrderRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;
      subData[sub] = subData[sub] || {};
      if (tx.includes('TDS INV')) subData[sub].tdsInv = f(r['Gross Amount']); // negative
      if (tx.includes('TDS DM'))  subData[sub].tdsDM  = f(r['Gross Amount']); // positive
    }

    // D) Returns sheet → debit invoice for returned orders
    for (const r of returnsRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;
      if (tx.includes('Return to Vendor')) {
        subData[sub] = subData[sub] || {};
        subData[sub].returnInv = f(r['Invoice Amount']); // negative
      }
    }

    // ── Pass 3: aggregate into per-SKU stats ──────────────────────────────
    // We build stats with the EXACT field names the frontend expects
    const skuStats = {};

    function initSKU(sku) {
      const m = skuMeta[sku] || { productName: 'Unknown', attr: '', type: 'pad', qty: 1, cogs: costs.pkg + costs.misc, sp: 0 };
      skuStats[sku] = {
        sku,
        productName: m.productName,
        attr: m.attr,
        sp: m.sp,
        type: m.type,
        qty: m.qty,
        unitCOGS: m.cogs,
        // Order counts
        orders: 0,          // delivered (has invoice)
        returned: 0,        // returned this month
        totalAttempts: 0,
        // Stage 1 — Gross Seller Payable
        totalInvoiceAmt: 0,
        totalMarketingFee: 0,
        totalCourierFee: 0,
        totalPaymentFee: 0,
        totalIGST: 0,
        totalWebAds: 0,     // web ads per-order charges (not bulk ad spend)
        grossPayableDelivered: 0,
        totalReturnReversal: 0,  // absolute value of what was clawed back
        grossPayableNet: 0,
        // Stage 2
        totalTDS: 0,
        totalTCS: 0,
        netSellerPayable: 0,
        // Stage 3
        totalCOGS: 0,
        adShare: 0,
        grossProfit: 0,
        grossMarginPct: 0,
        profitPerUnit: 0,
        // Return economics
        returnLossPerUnit: 0,
        totalReturnLoss: 0,
        // Per-order averages
        avgGrossPayable: 0,
        avgCutPerOrder: 0,
        avgTDSPerOrder: 0,
        avgNetPerOrder: 0,
        avgMarketingFee: 0,
        avgCourierFee: 0,
        avgPaymentFee: 0,
        avgIGST: 0,
        avgWebAds: 0,
        // Zone
        remoteOrders: 0,
        standardOrders: 0,
        // Aliases
        potentialRev: 0,
        totalSnapCut: 0,
        NET_PER_ORDER: 0,
        totalSnapNet: 0,
      };
    }

    for (const [sub, d] of Object.entries(subData)) {
      const sku = subToSKU[sub] || 'PRIOR_MONTH_RETURNS';
      if (!skuStats[sku]) initSKU(sku);
      const s = skuStats[sku];

      const invAmt     = d.invoiceAmt   || 0;
      const commTotal  = d.commTotal    || 0;
      const commMkt    = d.commMkt      || 0;
      const commCourier= d.commCourier  || 0;
      const commPmt    = d.commPmt      || 0;
      const commIGST   = d.commIGST     || 0;
      const tdsInv     = d.tdsInv       || 0;
      const returnInv  = d.returnInv    || 0;
      const returnComm = d.returnComm   || 0;
      const tdsDM      = d.tdsDM        || 0;

      const isDelivered = invAmt > 0;
      const isReturn    = returnInv < 0;

      if (isDelivered) {
        // Net = invoice + commission (neg) + TDS INV (neg)
        const grossPayable = invAmt + commTotal + tdsInv;
        s.orders++;
        s.totalInvoiceAmt    += invAmt;
        s.totalMarketingFee  += commMkt;
        s.totalCourierFee    += commCourier;
        s.totalPaymentFee    += commPmt;
        s.totalIGST          += commIGST;
        s.grossPayableDelivered += grossPayable;
        s.totalTDS           += tdsInv;
        // Zone classification: courier > -80 means remote
        if (Math.abs(commCourier) > 80) s.remoteOrders++;
        else s.standardOrders++;
      }

      if (isReturn) {
        // Return net = returnInv (neg) + returnComm (pos reversal) + tdsDM (pos)
        const returnNet = returnInv + returnComm + tdsDM;
        s.returned++;
        s.totalReturnReversal += Math.abs(returnNet); // store as positive loss magnitude
        s.totalTDS += tdsDM; // TDS DM is a credit (positive), add to TDS bucket
      }
    }

    // ── Pass 4: compute derived fields per SKU ────────────────────────────
    const totalDelivered = Object.values(skuStats)
      .filter(s => s.sku !== 'PRIOR_MONTH_RETURNS')
      .reduce((sum, s) => sum + s.orders, 0);

    for (const s of Object.values(skuStats)) {
      // Gross payable net = delivered gross payable − return reversals
      s.grossPayableNet = s.grossPayableDelivered - s.totalReturnReversal;

      // Stage 2: TCS (not in this data but keep field)
      s.netSellerPayable = s.grossPayableNet + s.totalTCS;

      // Stage 3
      s.totalCOGS  = s.orders * s.unitCOGS;
      s.adShare    = totalDelivered > 0
        ? totalAdSpend * (s.orders / totalDelivered)
        : 0;
      s.grossProfit = s.netSellerPayable - s.totalCOGS + s.adShare; // adShare is negative

      // Potential revenue (SP × delivered)
      s.potentialRev = s.sp * s.orders;
      s.grossMarginPct = s.potentialRev > 0
        ? (s.grossProfit / s.potentialRev) * 100
        : 0;
      s.profitPerUnit = s.orders > 0 ? s.grossProfit / s.orders : 0;

      // Counts
      s.totalAttempts = s.orders + s.returned;
      s.returnRate    = s.totalAttempts > 0
        ? (s.returned / s.totalAttempts) * 100
        : 0;

      // Per-order averages
      s.avgGrossPayable = s.orders > 0 ? s.grossPayableDelivered / s.orders : 0;
      s.avgTDSPerOrder  = s.orders > 0 ? s.totalTDS / s.orders : 0;
      s.avgNetPerOrder  = s.avgGrossPayable + s.avgTDSPerOrder;
      s.avgMarketingFee = s.orders > 0 ? s.totalMarketingFee / s.orders : 0;
      s.avgCourierFee   = s.orders > 0 ? s.totalCourierFee   / s.orders : 0;
      s.avgPaymentFee   = s.orders > 0 ? s.totalPaymentFee   / s.orders : 0;
      s.avgIGST         = s.orders > 0 ? s.totalIGST         / s.orders : 0;
      s.avgWebAds       = s.orders > 0 ? s.totalWebAds       / s.orders : 0;
      s.avgCutPerOrder  = s.sp - s.avgGrossPayable;

      // Return economics
      const avgReturnReversal = s.returned > 0 ? s.totalReturnReversal / s.returned : s.avgGrossPayable;
      s.returnLossPerUnit = avgReturnReversal + s.unitCOGS;
      s.totalReturnLoss   = s.returned * s.returnLossPerUnit;

      // Aliases
      s.NET_PER_ORDER = s.avgNetPerOrder;
      s.totalSnapNet  = s.netSellerPayable;
      s.totalSnapCut  = s.orders * s.avgCutPerOrder;
    }

    res.json({
      skus: Object.values(skuStats),
      totalAdSpend,
      totalDelivered,
      detectedSKUs: Object.keys(skuMeta).map(k => ({
        sku: k,
        productName: skuMeta[k].productName,
        type: skuMeta[k].type,
        qty: skuMeta[k].qty,
      })),
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
};
