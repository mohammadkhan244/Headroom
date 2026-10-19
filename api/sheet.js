const { timingSafeEqual } = require('crypto');

const { SHEETS_URL, SHEET_SECRET, APP_PASSCODE } = process.env;

function validPasscode(req) {
  const h = String(req.headers['x-passcode'] || '');
  if (!APP_PASSCODE || h.length !== APP_PASSCODE.length) return false;
  try {
    return timingSafeEqual(Buffer.from(h), Buffer.from(APP_PASSCODE));
  } catch { return false; }
}

module.exports = async function handler(req, res) {
  if (!validPasscode(req)) {
    await new Promise(r => setTimeout(r, 500));
    return res.status(401).json({ error: 'unauthorized' });
  }

  res.setHeader('Cache-Control', 'no-store');

  try {
    if (req.method === 'GET') {
      const url = `${SHEETS_URL}?secret=${encodeURIComponent(SHEET_SECRET)}&t=${Date.now()}`;
      const upstream = await fetch(url, { redirect: 'follow' });
      const data = await upstream.json();
      return res.json(data);
    }

    if (req.method === 'POST') {
      const body = { ...req.body, secret: SHEET_SECRET };
      const upstream = await fetch(SHEETS_URL, {
        method: 'POST',
        redirect: 'follow',
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
