/**
 * Teaspoon catering — order relay (Vercel serverless function)
 *
 * The order page posts to /api/submit on its own Vercel domain. This function adds a
 * secret and forwards the order to Google Apps Script, so the Apps Script URL and the
 * secret never appear in the page or in the GitHub repo.
 *
 * Set these in Vercel → Project → Settings → Environment Variables:
 *   APPS_SCRIPT_URL   your Apps Script Web app URL (ends in /exec)
 *   FORM_SECRET       the secret printed by running setup() in Apps Script
 * Then redeploy.
 */

const REQUIRED = ['Name', 'Email', 'Phone', 'Package'];
const MAX_FIELDS = 40;
const MAX_LEN = 2000;

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST');
    return res.status(405).json({ ok: false, error: 'Method not allowed' });
  }

  const url = process.env.APPS_SCRIPT_URL;
  const secret = process.env.FORM_SECRET;
  if (!url || !secret) {
    console.error('Missing APPS_SCRIPT_URL or FORM_SECRET environment variable');
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
  delete data.website;

  const keys = Object.keys(data);
  if (keys.length > MAX_FIELDS) return res.status(400).json({ ok: false, error: 'Invalid request' });
  const clean = {};
  for (const k of keys) clean[String(k).slice(0, 60)] = String(data[k] == null ? '' : data[k]).slice(0, MAX_LEN);
  if (REQUIRED.some(k => !clean[k] || !clean[k].trim())) {
    return res.status(400).json({ ok: false, error: 'Missing required information' });
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean.Email.trim())) {
    return res.status(400).json({ ok: false, error: 'Invalid email' });
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
    return res.status(200).json({ ok: true, id: out.id, phone: out.phone || '' });
  } catch (err) {
    console.error('Relay failed:', err);
    return res.status(502).json({ ok: false, error: 'Could not save the request' });
  }
};
