const { SHEETS_URL, SHEET_SECRET } = process.env;
const ALLOWED_SHEETS = new Set(['Expenses', 'Balances', 'Config', 'all']);

// Strip secrets before logging or returning to client
function scrub(s) {
  s = String(s || '');
  if (SHEETS_URL)   s = s.split(SHEETS_URL).join('[URL]');
  if (SHEET_SECRET) s = s.split(SHEET_SECRET).join('[SECRET]');
  return s;
}

function timedFetch(url, opts, ms) {
  const ctrl = new AbortController();
  const t = setTimeout(() => ctrl.abort(), ms);
  return fetch(url, { ...opts, signal: ctrl.signal }).finally(() => clearTimeout(t));
}

module.exports = async function handler(req, res) {
  res.setHeader('Cache-Control', 'no-store');

  const sheetParam = req.query.sheet || (req.body && req.body.sheet) || '';
  const sheet = ALLOWED_SHEETS.has(sheetParam) ? sheetParam : 'Expenses';

  try {
    let upstream;

    if (req.method === 'GET') {
      const qs  = `&sheet=${encodeURIComponent(sheet)}`;
      const url = `${SHEETS_URL}?secret=${encodeURIComponent(SHEET_SECRET)}&t=${Date.now()}${qs}`;
      try {
        upstream = await timedFetch(url, { redirect: 'follow' }, 25000);
      } catch {
        // one retry on network failure (GET is idempotent)
        upstream = await timedFetch(url, { redirect: 'follow' }, 25000);
      }
    } else if (req.method === 'POST') {
      const body = { ...req.body, secret: SHEET_SECRET };
      if (sheetParam && ALLOWED_SHEETS.has(sheetParam)) body.sheet = sheetParam;
      upstream = await timedFetch(SHEETS_URL, {
        method:  'POST',
        redirect: 'follow',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify(body),
      }, 25000);
    } else {
      return res.status(405).end();
    }

    const ct   = upstream.headers.get('content-type') || '';
    const text = await upstream.text();

    if (ct.includes('text/html')) {
      const err = { error: 'upstream returned html', upstreamStatus: upstream.status };
      console.error('[sheet.js] upstream html:', upstream.status, scrub(text).slice(0, 80));
      return res.status(502).json(err);
    }

    let data;
    try {
      data = JSON.parse(text);
    } catch {
      const snippet = scrub(text).slice(0, 80);
      const err = { error: 'upstream json parse failed', upstreamStatus: upstream.status, snippet };
      console.error('[sheet.js] json parse failed:', err);
      return res.status(502).json(err);
    }

    return res.json(data);

  } catch (e) {
    const isTimeout = e && (e.name === 'AbortError' || e.name === 'TimeoutError');
    const err = { error: isTimeout ? 'upstream timeout' : 'proxy error', isTimeout: !!isTimeout };
    console.error('[sheet.js] catch:', err, scrub(String(e)));
    return res.status(502).json(err);
  }
};
