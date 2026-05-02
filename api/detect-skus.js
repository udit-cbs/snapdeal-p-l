const XLSX   = require('xlsx');
const busboy = require('busboy');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') { res.status(200).end(); return; }
  if (req.method !== 'POST')    { res.status(405).json({ error: 'POST only' }); return; }

  try {
    const buf = await parseFile(req, 'orders');
    const wb  = XLSX.read(buf, { type: 'buffer' });
    const rows = sheetToRows(wb, ['order','consolidate','comp']);
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
};

function parseFile(req, fieldName) {
  return new Promise((resolve, reject) => {
    const bb = busboy({ headers: req.headers, limits: { fileSize: 20*1024*1024 } });
    let found = false;
    bb.on('file', (name, stream) => {
      if (name !== fieldName) { stream.resume(); return; }
      found = true;
      const chunks = [];
      stream.on('data', d => chunks.push(d));
      stream.on('end', () => resolve(Buffer.concat(chunks)));
    });
    bb.on('finish', () => { if (!found) reject(new Error(`Field '${fieldName}' not found`)); });
    bb.on('error', reject);
    req.pipe(bb);
  });
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
