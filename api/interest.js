/**
 * Teaspoon catering — short interest form relay (Vercel serverless function)
 *
 * The /catering page posts to /api/interest. This function adds a secret and forwards
 * the request to its own Google Apps Script (google-apps-script/interest-form.gs),
 * which saves it to the "interest" spreadsheet. The detailed form (/catering-details)
 * keeps using /api/submit and its own spreadsheet.
 *
 * Set these in Vercel → Project → Settings → Environment Variables:
 *   INTEREST_SCRIPT_URL   the interest Apps Script Web app URL (ends in /exec)
 *   INTEREST_SECRET       the secret printed by running setup() in that Apps Script
 * Then redeploy.
 */

const FIELDS = ['Package', 'Event date', 'Time of day', 'Email', 'Contact via', 'Phone',
  'Heard via', 'Notes', 'Email specials', 'Text specials', 'Agreed to terms'];
const REQUIRED = ['Package', 'Event date', 'Time of day', 'Email'];
const MAX_LEN = 2000;

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  }

  const url = process.env.INTEREST_SCRIPT_URL;
  const secret = process.env.INTEREST_SECRET;
  if (!url || !secret) {
    console.error('Missing INTEREST_SCRIPT_URL or INTEREST_SECRET environment variable');
    return res.status(500).json({ ok: false, error: 'Server is not set up yet' });
  }

  let data = req.body;
  if (typeof data === 'string') {
    try { data = JSON.parse(data); } catch (e) { data = null; }
  }
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    return res.status(400).json({ ok: false, error: 'Invalid request' });
  }

  // Spam trap: real visitors never see or fill the hidden "website" field.
  if (data.website) return res.status(200).json({ ok: true, id: 'TS-RECEIVED' });

  // Only known fields are passed on, so the sheet columns stay fixed.
  const clean = {};
  for (const k of FIELDS) clean[k] = String(data[k] == null ? '' : data[k]).slice(0, MAX_LEN).trim();
  if (REQUIRED.some(k => !clean[k])) {
    return res.status(400).json({ ok: false, error: 'Missing required fields' });
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean.Email)) {
    return res.status(400).json({ ok: false, error: 'Invalid email' });
  }
  if (clean['Contact via'] !== 'Email' && clean.Phone.replace(/\D/g, '').length < 10) {
    return res.status(400).json({ ok: false, error: 'Invalid phone' });
  }

  try {
    // Apps Script answers POSTs with a redirect to the result; fetch follows it.
    const r = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ secret, data: clean }),
      redirect: 'follow',
    });
    const out = await r.json().catch(() => null);
    if (!out || !out.ok) {
      console.error('Apps Script error:', out && out.error, 'HTTP', r.status);
      return res.status(502).json({ ok: false, error: 'Could not save the request' });
    }
    return res.status(200).json({ ok: true, id: out.id });
  } catch (err) {
    console.error('Relay failed:', err);
    return res.status(502).json({ ok: false, error: 'Could not save the request' });
  }
};
