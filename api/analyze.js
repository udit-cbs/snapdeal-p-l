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

    // ── User cost inputs ───────────────────────────────────────────────────
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

    function unitCOGS(sku, productName) {
      const cfg  = cfgMap[sku];
      const type = cfg ? cfg.type : inferType(sku, productName);
      const qty  = cfg ? cfg.qty  : inferQty(sku);
      const unit = type === 'pad' ? costs.pad : type === 'liner' ? costs.liner : costs.panty;
      return { cogs: (unit * qty) + costs.pkg + costs.misc, type, qty };
    }

    // ── Read files ─────────────────────────────────────────────────────────
    const wbOrders = XLSX.read(files.orders, { type: 'buffer' });
    const wbPL     = files.pl ? XLSX.read(files.pl, { type: 'buffer' }) : null;

    // Order report: single sheet
    const orderRows = sheetByHint(wbOrders, ['consolidate', 'order', 'comp']);

    // P&L sheet rows (all optional — only available if P&L file uploaded)
    const totalSubRows  = wbPL ? sheetByHint(wbPL, ['total_sub', 'total sub'])   : [];
    const commRows      = wbPL ? sheetByHint(wbPL, ['commission and other', 'commission']) : [];
    const nonOrderRows  = wbPL ? sheetByHint(wbPL, ['non order', 'non_order'])   : [];
    const returnsRows   = wbPL ? sheetByHint(wbPL, ['returns'])                  : [];

    // ── Pass 1: build sub→SKU map from order report ────────────────────────
    const subToSKU  = {};   // suborder code → sku
    const subToSP   = {};   // suborder code → selling price
    const skuMeta   = {};   // sku → { productName, type, qty, cogs, sp }

    for (const o of orderRows) {
      const sub  = str(o['SUBORDER CODE']);
      const sku  = str(o['SKU CODE']);
      const name = str(o['PRODUCT NAME']);
      const sp   = f(o['SELLING PRICE']);
      if (!sku || !sub) continue;

      subToSKU[sub] = sku;
      subToSP[sub]  = sp;

      if (!skuMeta[sku]) {
        const { cogs, type, qty } = unitCOGS(sku, name);
        skuMeta[sku] = { productName: name, type, qty, cogs, sp };
      }
    }

    // ── Pass 2: per-suborder financial data from P&L sheets ───────────────
    // sub_data[sub] = { invoiceAmt, commTotal, commMkt, commCourier, commPmt,
    //                   commIGST, tdsInv, tdsDM, returnInv, returnComm }
    const subData = {};

    // A) Total_Suboders → invoice amount for each COD Vendor Invoice
    for (const r of totalSubRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;
      if (tx.includes('Vendor Invoice')) {
        subData[sub] = subData[sub] || {};
        subData[sub].invoiceAmt = f(r['Invoice Amount']);
      }
    }

    // B) Commission sheet → per-order charges + returns + ad spend
    let totalAdSpend = 0;

    for (const r of commRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;

      // Ad spend: Advertise Ads Income + Web Ads Invoice (monthly totals, not per-sub)
      if (tx.includes('Advertise Ads Income') || tx.includes('Web Ads Invoice')) {
        totalAdSpend += f(r['Total Commission Amount']);
        continue;
      }

      // Skip zero-value entries
      if (tx.includes('Stock Out') || tx.includes('RTO Charges')) continue;

      subData[sub] = subData[sub] || {};

      if (tx.includes('COD Vendor Invoice')) {
        // Snapdeal's cut on a delivered order (all negative)
        subData[sub].commTotal   = f(r['Total Commission Amount']);
        subData[sub].commMkt     = f(r['Marketing Fee']);
        subData[sub].commCourier = f(r['Courier Fee']);
        subData[sub].commPmt     = f(r['Payment Collection Fee']);
        subData[sub].commIGST    = f(r['Igst'] || r['IGST'] || 0);
      }

      if (tx.includes('COD Return to Vendor')) {
        // Commission reversal on a returned order (positive — Snapdeal gives back its cut)
        subData[sub].returnComm    = f(r['Total Commission Amount']);
        subData[sub].returnCourier = f(r['Courier Fee']);
      }
    }

    // C) Non-Order Transactions → TDS INV (deducted on delivery) & TDS DM (credited on return)
    for (const r of nonOrderRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;
      subData[sub] = subData[sub] || {};
      if (tx.includes('TDS INV')) subData[sub].tdsInv = f(r['Gross Amount']); // negative
      if (tx.includes('TDS DM'))  subData[sub].tdsDM  = f(r['Gross Amount']); // positive
    }

    // D) Returns sheet → invoice debit for returned orders (negative)
    for (const r of returnsRows) {
      const sub = str(r['Sub Order No']);
      const tx  = str(r['Transaction Type']);
      if (!sub || sub === 'nan') continue;
      if (tx.includes('Return to Vendor')) {
        subData[sub] = subData[sub] || {};
        subData[sub].returnInv = f(r['Invoice Amount']); // negative
      }
    }

    // ── Pass 3: aggregate per SKU ──────────────────────────────────────────
    const skuStats = {}; // sku → aggregated financials

    function getSKUStat(sku, meta) {
      if (!skuStats[sku]) {
        const m = meta || skuMeta[sku] || { productName: 'Unknown', type: 'pad', qty: 1, cogs: costs.pkg + costs.misc, sp: 0 };
        skuStats[sku] = {
          sku,
          productName: m.productName,
          type: m.type,
          qty: m.qty,
          unitCOGS: m.cogs,
          sp: m.sp,
          // Delivered
          deliveredCount: 0,
          totalInvoiceAmt: 0,    // from Total_Suboders
          totalCommission: 0,    // from Commission (negative)
          totalTDSInv: 0,        // from Non-Order TDS INV (negative)
          netDelivered: 0,       // invoiceAmt + commission + tdsInv per delivered sub
          // Returns this month (including prior-month orders returned this month)
          returnedCount: 0,
          totalReturnInv: 0,     // from Returns sheet (negative)
          totalReturnComm: 0,    // commission reversal (positive)
          totalTDSDM: 0,         // TDS debit memo (positive)
          netReturns: 0,         // returnInv + returnComm + tdsDM per return sub
          // COGS & ads allocated after
          totalCOGS: 0,
          adShare: 0,
          grossProfit: 0,
          // Per-order breakdown
          orders: [],            // per-order detail rows
        };
      }
      return skuStats[sku];
    }

    // Process each suborder
    for (const [sub, d] of Object.entries(subData)) {
      const sku  = subToSKU[sub];
      const isDelivered = (d.invoiceAmt || 0) > 0;
      const isReturn    = (d.returnInv  || 0) < 0;

      if (!isDelivered && !isReturn) continue; // skip zero-impact rows

      // For returned subs not in order report, we still record financial impact
      // under 'PRIOR_RETURNS' bucket if SKU unknown
      const effectiveSKU = sku || 'PRIOR_MONTH_RETURNS';

      const stat = getSKUStat(effectiveSKU, skuMeta[sku]);
      const sp   = subToSP[sub] || stat.sp || 0;

      if (isDelivered && !isReturn) {
        // Pure delivered order
        const invAmt  = d.invoiceAmt  || 0;
        const comm    = d.commTotal   || 0;
        const tdsInv  = d.tdsInv      || 0;
        const net     = invAmt + comm + tdsInv;

        stat.deliveredCount++;
        stat.totalInvoiceAmt += invAmt;
        stat.totalCommission += comm;
        stat.totalTDSInv     += tdsInv;
        stat.netDelivered    += net;
        stat.orders.push({
          sub, sku: effectiveSKU, sp,
          invoiceAmt: invAmt, commission: comm, tdsInv, net,
          type: 'delivered',
          courier: d.commCourier || 0,
          mkt: d.commMkt || 0,
          pmt: d.commPmt || 0,
          igst: d.commIGST || 0,
        });
      }

      if (isDelivered && isReturn) {
        // Order that was delivered in this month AND returned in this month
        // Net on invoice side = 0 (charged then reversed), but COGS + return loss apply
        const invAmt     = d.invoiceAmt  || 0;
        const comm       = d.commTotal   || 0;
        const tdsInv     = d.tdsInv      || 0;
        const retInv     = d.returnInv   || 0;
        const retComm    = d.returnComm  || 0;
        const tdsDM      = d.tdsDM       || 0;
        const netDel     = invAmt + comm + tdsInv;         // ~0 or small
        const netRet     = retInv + retComm + tdsDM;       // negative
        const net        = netDel + netRet;

        // Count as delivered (COGS applies) AND returned
        stat.deliveredCount++;
        stat.totalInvoiceAmt += invAmt;
        stat.totalCommission += comm;
        stat.totalTDSInv     += tdsInv;
        stat.netDelivered    += netDel;
        stat.returnedCount++;
        stat.totalReturnInv  += retInv;
        stat.totalReturnComm += retComm;
        stat.totalTDSDM      += tdsDM;
        stat.netReturns      += netRet;
        stat.orders.push({
          sub, sku: effectiveSKU, sp,
          invoiceAmt: invAmt, commission: comm, tdsInv,
          returnInv: retInv, returnComm: retComm, tdsDM, net,
          type: 'delivered_and_returned',
          courier: d.commCourier || 0,
        });
      }

      if (!isDelivered && isReturn) {
        // Return for a prior-month order — no delivery invoice this month
        const retInv  = d.returnInv  || 0;
        const retComm = d.returnComm || 0;
        const tdsDM   = d.tdsDM      || 0;
        const net     = retInv + retComm + tdsDM;

        stat.returnedCount++;
        stat.totalReturnInv  += retInv;
        stat.totalReturnComm += retComm;
        stat.totalTDSDM      += tdsDM;
        stat.netReturns      += net;
        stat.orders.push({
          sub, sku: effectiveSKU, sp: 0,
          returnInv: retInv, returnComm: retComm, tdsDM, net,
          type: 'return_only',
        });
      }
    }

    // ── Pass 4: allocate COGS & ad spend, compute gross profit ────────────
    const totalDelivered = Object.values(skuStats)
      .filter(s => s.sku !== 'PRIOR_MONTH_RETURNS')
      .reduce((sum, s) => sum + s.deliveredCount, 0);

    for (const stat of Object.values(skuStats)) {
      stat.totalCOGS = stat.deliveredCount * stat.unitCOGS;
      stat.adShare   = totalDelivered > 0
        ? totalAdSpend * (stat.deliveredCount / totalDelivered)
        : 0;
      // gross profit = net received from Snapdeal (delivered + returns) - COGS + ad allocation
      stat.grossProfit = stat.netDelivered + stat.netReturns - stat.totalCOGS + stat.adShare;
      stat.grossMarginPct = stat.netDelivered > 0
        ? (stat.grossProfit / (stat.sp * stat.deliveredCount || stat.netDelivered)) * 100
        : 0;
      stat.profitPerUnit = stat.deliveredCount > 0
        ? stat.grossProfit / stat.deliveredCount
        : 0;
      stat.avgNetPerOrder = stat.deliveredCount > 0
        ? stat.netDelivered / stat.deliveredCount
        : 0;
      stat.returnRate = (stat.deliveredCount + stat.returnedCount) > 0
        ? stat.returnedCount / (stat.deliveredCount + stat.returnedCount) * 100
        : 0;
    }

    const skus = Object.values(skuStats);

    res.json({
      skus,
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
