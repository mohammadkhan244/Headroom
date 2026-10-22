const { SHEETS_URL, SHEET_SECRET } = process.env;
const ALLOWED_SHEETS = new Set(['Expenses', 'Balances', 'Config', 'all']);

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  // Whitelist sheet param
  const sheetParam = req.query.sheet || (req.body && req.body.sheet) || '';
  const sheet = ALLOWED_SHEETS.has(sheetParam) ? sheetParam : 'Expenses';

  try {
    if (req.method === 'GET') {
      const qs = sheet ? `&sheet=${encodeURIComponent(sheet)}` : '';
      const url = `${SHEETS_URL}?secret=${encodeURIComponent(SHEET_SECRET)}&t=${Date.now()}${qs}`;
      const upstream = await fetch(url, { redirect: 'follow' });
      const data = await upstream.json();
      return res.json(data);
    }

    if (req.method === 'POST') {
      const body = { ...req.body, secret: SHEET_SECRET };
      if (sheetParam && ALLOWED_SHEETS.has(sheetParam)) body.sheet = sheetParam;
      const upstream = await fetch(SHEETS_URL, {
        method: 'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      });
      const data = await upstream.json();
      return res.json(data);
    }

    return res.status(405).end();
  } catch {
    // Never expose SHEETS_URL or SHEET_SECRET in error messages
    return res.status(502).json({ error: 'proxy error' });
  }
};
