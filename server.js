const express = require('express');
const { Pool } = require('pg');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const JWT_SECRET = process.env.JWT_SECRET || 'eicr-field-pro-secret-key-change-me';
const PORT = process.env.PORT || 3000;

// Server-side API keys from Railway env vars (fallback if no user key configured)
const ENV_KEYS = {
  gemini: process.env.GEMINI_API_KEY || '',
  anthropic: process.env.ANTHROPIC_API_KEY || ''
};
const DEFAULT_MODELS = {
  gemini: process.env.GEMINI_MODEL || 'gemini-2.5-flash',
  anthropic: process.env.ANTHROPIC_MODEL || 'claude-sonnet-4-20250514'
};

// ---- DATABASE ----
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

async function initDB() {
  const client = await pool.connect();
  try {
    await client.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        name TEXT NOT NULL,
        role TEXT DEFAULT 'inspector',
        created_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS ai_keys (
        id SERIAL PRIMARY KEY,
        user_id INT REFERENCES users(id),
        label TEXT NOT NULL,
        provider TEXT NOT NULL,
        model TEXT NOT NULL,
        api_key TEXT NOT NULL
      );

      CREATE TABLE IF NOT EXISTS sites (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL,
        address TEXT,
        postcode TEXT,
        client_name TEXT,
        client_address TEXT,
        client_tel TEXT,
        description TEXT DEFAULT 'Commercial',
        purpose TEXT DEFAULT 'Periodic inspection',
        wiring_age INT,
        additions TEXT DEFAULT 'No',
        additions_age INT,
        last_inspection_date TEXT,
        records_available TEXT DEFAULT 'No',
        extent TEXT,
        agreed_limitations TEXT,
        agreed_with TEXT,
        operational_limitations TEXT,
        overall_assessment TEXT,
        next_inspection_date TEXT,
        general_condition TEXT,
        earthing_type TEXT,
        num_phases TEXT,
        supply_type TEXT DEFAULT 'AC',
        nominal_voltage TEXT DEFAULT '230/400',
        nominal_frequency TEXT DEFAULT '50',
        ipf_at_origin TEXT,
        ze_at_origin TEXT,
        num_supplies TEXT DEFAULT '1',
        supply_device_bsen TEXT,
        supply_device_type TEXT,
        supply_device_rating TEXT,
        means_of_earthing TEXT,
        earth_electrode_type TEXT,
        earth_electrode_resistance TEXT,
        earth_electrode_location TEXT,
        main_switch_location TEXT,
        main_switch_bsen TEXT,
        main_switch_poles TEXT,
        main_switch_current_rating TEXT,
        main_switch_fuse_rating TEXT,
        main_switch_voltage_rating TEXT,
        main_switch_rcd_type TEXT,
        main_switch_rcd_idn TEXT,
        main_switch_rcd_delay TEXT,
        main_switch_rcd_time TEXT,
        earthing_conductor_material TEXT DEFAULT 'Copper',
        earthing_conductor_csa TEXT,
        earthing_conductor_verified TEXT,
        bonding_conductor_material TEXT DEFAULT 'Copper',
        bonding_conductor_csa TEXT,
        bonding_conductor_verified TEXT,
        bonding_water TEXT,
        bonding_gas TEXT,
        bonding_oil TEXT,
        bonding_lightning TEXT,
        bonding_steel TEXT,
        bonding_other TEXT,
        inspector_name TEXT,
        inspector_position TEXT,
        inspector_date TEXT,
        company_name TEXT,
        company_address TEXT,
        company_reg TEXT,
        company_tel TEXT,
        report_ref TEXT,
        created_by INT REFERENCES users(id),
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS company_profile (
        id SERIAL PRIMARY KEY,
        name TEXT,
        address TEXT,
        postcode TEXT,
        tel TEXT,
        registration_body TEXT,
        registration_number TEXT,
        logo TEXT
      );

      CREATE TABLE IF NOT EXISTS signatures (
        id SERIAL PRIMARY KEY,
        user_id INT UNIQUE REFERENCES users(id),
        signature_data TEXT,
        updated_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS photos (
        id SERIAL PRIMARY KEY,
        site_id INT,
        board_id INT,
        circuit_id INT,
        observation_id INT,
        data TEXT NOT NULL,
        caption TEXT,
        created_by INT REFERENCES users(id),
        created_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS inspections (
        id SERIAL PRIMARY KEY,
        site_id INT REFERENCES sites(id) ON DELETE CASCADE,
        item_ref TEXT NOT NULL,
        outcome TEXT DEFAULT 'N/V',
        UNIQUE(site_id, item_ref)
      );

      CREATE TABLE IF NOT EXISTS boards (
        id SERIAL PRIMARY KEY,
        site_id INT REFERENCES sites(id) ON DELETE CASCADE,
        ref TEXT NOT NULL,
        location TEXT,
        supplied_from TEXT,
        dist_bsen TEXT,
        dist_type TEXT,
        dist_rating TEXT,
        num_phases TEXT DEFAULT '1',
        spd_types TEXT,
        spd_status TEXT,
        supply_polarity TEXT,
        phase_sequence TEXT,
        ze TEXT,
        ipf TEXT,
        instruments_multi TEXT,
        instruments_ir TEXT,
        instruments_cont TEXT,
        instruments_earth TEXT,
        instruments_loop TEXT,
        instruments_rcd TEXT,
        photo TEXT,
        sort_order INT DEFAULT 0,
        created_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS circuits (
        id SERIAL PRIMARY KEY,
        board_id INT REFERENCES boards(id) ON DELETE CASCADE,
        number TEXT,
        description TEXT,
        wiring_type TEXT,
        ref_method TEXT,
        num_points TEXT,
        live_mm TEXT,
        cpc_mm TEXT,
        max_disconnect TEXT DEFAULT '0.4',
        ocpd_bsen TEXT,
        ocpd_type TEXT,
        ocpd_rating TEXT,
        ocpd_breaking_cap TEXT,
        max_zs TEXT,
        rcd_bsen TEXT,
        rcd_type TEXT,
        rcd_idn TEXT,
        rcd_rating_a TEXT,
        r1 TEXT,
        rn TEXT,
        r2 TEXT,
        r1r2 TEXT,
        r1r2_or_r2 TEXT,
        r2_ring TEXT,
        test_voltage TEXT DEFAULT '500',
        ir_ll TEXT,
        ir_le TEXT,
        polarity TEXT,
        zs_measured TEXT,
        rcd_time_x1 TEXT,
        rcd_time_x5 TEXT,
        rcd_test_button TEXT,
        afdd_test TEXT,
        remarks TEXT,
        sort_order INT DEFAULT 0,
        created_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS observations (
        id SERIAL PRIMARY KEY,
        site_id INT REFERENCES sites(id) ON DELETE CASCADE,
        board_id INT,
        item_no INT,
        description TEXT NOT NULL,
        code TEXT NOT NULL DEFAULT 'C2',
        location TEXT,
        photo TEXT,
        created_by INT REFERENCES users(id),
        created_at TIMESTAMPTZ DEFAULT NOW()
      );
    `);

    // Seed default users if none exist
    const { rows } = await client.query('SELECT COUNT(*) FROM users');
    if (parseInt(rows[0].count) === 0) {
      const hash1 = await bcrypt.hash('wayne1', 10);
      const hash2 = await bcrypt.hash('john1', 10);
      await client.query(`INSERT INTO users (username, password_hash, name, role) VALUES ('wayne', $1, 'Wayne Harrow', 'admin')`, [hash1]);
      await client.query(`INSERT INTO users (username, password_hash, name, role) VALUES ('john', $1, 'John Harrow', 'inspector')`, [hash2]);
      console.log('Default users created: wayne/wayne1, john/john1');
    }
  } finally {
    client.release();
  }
}

// ---- AUTH MIDDLEWARE ----
function authMiddleware(req, res, next) {
  const token = req.headers.authorization?.replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'No token' });
  try {
    req.user = jwt.verify(token, JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: 'Invalid token' });
  }
}

// ---- AUTH ROUTES ----
app.post('/api/auth/login', async (req, res) => {
  const { username, password } = req.body;
  const { rows } = await pool.query('SELECT * FROM users WHERE username = $1', [username]);
  if (!rows[0]) return res.status(401).json({ error: 'Invalid credentials' });
  const valid = await bcrypt.compare(password, rows[0].password_hash);
  if (!valid) return res.status(401).json({ error: 'Invalid credentials' });
  const token = jwt.sign({ id: rows[0].id, username: rows[0].username, name: rows[0].name, role: rows[0].role }, JWT_SECRET, { expiresIn: '30d' });
  res.json({ token, user: { id: rows[0].id, name: rows[0].name, username: rows[0].username, role: rows[0].role } });
});

app.get('/api/auth/me', authMiddleware, (req, res) => res.json(req.user));

// ---- AI KEYS ----
app.get('/api/ai-keys', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT id, label, provider, model FROM ai_keys WHERE user_id = $1', [req.user.id]);
  res.json(rows);
});

app.post('/api/ai-keys', authMiddleware, async (req, res) => {
  const { label, provider, model, api_key } = req.body;
  const { rows } = await pool.query('INSERT INTO ai_keys (user_id, label, provider, model, api_key) VALUES ($1,$2,$3,$4,$5) RETURNING id, label, provider, model', [req.user.id, label, provider, model, api_key]);
  res.json(rows[0]);
});

app.delete('/api/ai-keys/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM ai_keys WHERE id = $1 AND user_id = $2', [req.params.id, req.user.id]);
  res.json({ ok: true });
});

// ---- SITES CRUD ----
app.get('/api/sites', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM sites ORDER BY created_at DESC');
  res.json(rows);
});

app.post('/api/sites', authMiddleware, async (req, res) => {
  const d = req.body;
  const { rows } = await pool.query(
    `INSERT INTO sites (name, address, postcode, client_name, client_address, client_tel, description, report_ref, created_by)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9) RETURNING *`,
    [d.name, d.address, d.postcode, d.client_name, d.client_address, d.client_tel, d.description || 'Commercial', d.report_ref, req.user.id]
  );
  res.json(rows[0]);
});

app.get('/api/sites/:id', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM sites WHERE id = $1', [req.params.id]);
  if (!rows[0]) return res.status(404).json({ error: 'Not found' });
  res.json(rows[0]);
});

app.put('/api/sites/:id', authMiddleware, async (req, res) => {
  const d = req.body;
  const fields = Object.keys(d).filter(k => k !== 'id' && k !== 'created_at' && k !== 'created_by');
  if (!fields.length) return res.json({ ok: true });
  const sets = fields.map((f, i) => `${f} = $${i + 2}`);
  const vals = fields.map(f => d[f]);
  await pool.query(`UPDATE sites SET ${sets.join(', ')}, updated_at = NOW() WHERE id = $1`, [req.params.id, ...vals]);
  const { rows } = await pool.query('SELECT * FROM sites WHERE id = $1', [req.params.id]);
  res.json(rows[0]);
});

app.delete('/api/sites/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM sites WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// ---- INSPECTIONS ----
app.get('/api/sites/:id/inspections', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM inspections WHERE site_id = $1 ORDER BY item_ref', [req.params.id]);
  res.json(rows);
});

app.put('/api/sites/:id/inspections', authMiddleware, async (req, res) => {
  const items = req.body; // [{item_ref, outcome}, ...]
  for (const item of items) {
    await pool.query(
      `INSERT INTO inspections (site_id, item_ref, outcome) VALUES ($1, $2, $3)
       ON CONFLICT (site_id, item_ref) DO UPDATE SET outcome = $3`,
      [req.params.id, item.item_ref, item.outcome]
    );
  }
  res.json({ ok: true });
});

// ---- BOARDS CRUD ----
app.get('/api/sites/:id/boards', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM boards WHERE site_id = $1 ORDER BY sort_order, id', [req.params.id]);
  res.json(rows);
});

app.post('/api/sites/:id/boards', authMiddleware, async (req, res) => {
  const d = req.body;
  const { rows } = await pool.query(
    `INSERT INTO boards (site_id, ref, location, supplied_from, num_phases) VALUES ($1,$2,$3,$4,$5) RETURNING *`,
    [req.params.id, d.ref, d.location, d.supplied_from, d.num_phases || '1']
  );
  res.json(rows[0]);
});

app.put('/api/boards/:id', authMiddleware, async (req, res) => {
  const d = req.body;
  const fields = Object.keys(d).filter(k => k !== 'id' && k !== 'site_id' && k !== 'created_at');
  if (!fields.length) return res.json({ ok: true });
  const sets = fields.map((f, i) => `${f} = $${i + 2}`);
  const vals = fields.map(f => d[f]);
  await pool.query(`UPDATE boards SET ${sets.join(', ')} WHERE id = $1`, [req.params.id, ...vals]);
  const { rows } = await pool.query('SELECT * FROM boards WHERE id = $1', [req.params.id]);
  res.json(rows[0]);
});

app.delete('/api/boards/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM boards WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// ---- CIRCUITS CRUD ----
app.get('/api/boards/:id/circuits', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM circuits WHERE board_id = $1 ORDER BY sort_order, id', [req.params.id]);
  res.json(rows);
});

app.post('/api/boards/:id/circuits', authMiddleware, async (req, res) => {
  const d = req.body;
  const { rows } = await pool.query(
    `INSERT INTO circuits (board_id, number, description, sort_order) VALUES ($1,$2,$3,$4) RETURNING *`,
    [req.params.id, d.number, d.description, d.sort_order || 0]
  );
  res.json(rows[0]);
});

app.put('/api/circuits/:id', authMiddleware, async (req, res) => {
  const d = req.body;
  const fields = Object.keys(d).filter(k => k !== 'id' && k !== 'board_id' && k !== 'created_at');
  if (!fields.length) return res.json({ ok: true });
  const sets = fields.map((f, i) => `${f} = $${i + 2}`);
  const vals = fields.map(f => d[f]);
  await pool.query(`UPDATE circuits SET ${sets.join(', ')} WHERE id = $1`, [req.params.id, ...vals]);
  const { rows } = await pool.query('SELECT * FROM circuits WHERE id = $1', [req.params.id]);
  res.json(rows[0]);
});

app.put('/api/boards/:id/circuits/bulk', authMiddleware, async (req, res) => {
  const updates = req.body; // [{id or number, ...fields}]
  const boardId = req.params.id;
  for (const u of updates) {
    if (u.id) {
      const fields = Object.keys(u).filter(k => k !== 'id' && k !== 'board_id');
      const sets = fields.map((f, i) => `${f} = $${i + 2}`);
      const vals = fields.map(f => u[f]);
      await pool.query(`UPDATE circuits SET ${sets.join(', ')} WHERE id = $1`, [u.id, ...vals]);
    } else {
      // Match by number or create new
      const { rows: existing } = await pool.query('SELECT id FROM circuits WHERE board_id = $1 AND number = $2', [boardId, u.number]);
      if (existing[0]) {
        const fields = Object.keys(u).filter(k => k !== 'id' && k !== 'board_id' && k !== 'number');
        const sets = fields.map((f, i) => `${f} = $${i + 2}`);
        const vals = fields.map(f => u[f]);
        if (sets.length) await pool.query(`UPDATE circuits SET ${sets.join(', ')} WHERE id = $1`, [existing[0].id, ...vals]);
      } else {
        const fields = Object.keys(u).filter(k => k !== 'id');
        const cols = ['board_id', ...fields];
        const placeholders = cols.map((_, i) => `$${i + 1}`);
        const vals = [boardId, ...fields.map(f => u[f])];
        await pool.query(`INSERT INTO circuits (${cols.join(',')}) VALUES (${placeholders.join(',')})`, vals);
      }
    }
  }
  const { rows } = await pool.query('SELECT * FROM circuits WHERE board_id = $1 ORDER BY sort_order, id', [boardId]);
  res.json(rows);
});

app.delete('/api/circuits/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM circuits WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// ---- OBSERVATIONS CRUD ----
app.get('/api/sites/:id/observations', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT o.*, u.name as created_by_name FROM observations o LEFT JOIN users u ON o.created_by = u.id WHERE o.site_id = $1 ORDER BY o.item_no', [req.params.id]);
  res.json(rows);
});

app.post('/api/sites/:id/observations', authMiddleware, async (req, res) => {
  const d = req.body;
  const { rows: maxRow } = await pool.query('SELECT COALESCE(MAX(item_no), 0) + 1 as next FROM observations WHERE site_id = $1', [req.params.id]);
  const { rows } = await pool.query(
    `INSERT INTO observations (site_id, board_id, item_no, description, code, location, photo, created_by) VALUES ($1,$2,$3,$4,$5,$6,$7,$8) RETURNING *`,
    [req.params.id, d.board_id, d.item_no || maxRow[0].next, d.description, d.code || 'C2', d.location, d.photo, req.user.id]
  );
  res.json(rows[0]);
});

app.put('/api/observations/:id', authMiddleware, async (req, res) => {
  const d = req.body;
  const fields = Object.keys(d).filter(k => !['id', 'site_id', 'created_at', 'created_by'].includes(k));
  if (!fields.length) return res.json({ ok: true });
  const sets = fields.map((f, i) => `${f} = $${i + 2}`);
  const vals = fields.map(f => d[f]);
  await pool.query(`UPDATE observations SET ${sets.join(', ')} WHERE id = $1`, [req.params.id, ...vals]);
  res.json({ ok: true });
});

app.delete('/api/observations/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM observations WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// ---- AI PROXY ----
// Endpoint to list available AI providers (including env var ones)
app.get('/api/ai/providers', authMiddleware, async (req, res) => {
  const providers = [];
  if (ENV_KEYS.gemini) providers.push({ id: 'env_gemini', label: 'Gemini (Server)', provider: 'gemini', model: DEFAULT_MODELS.gemini });
  if (ENV_KEYS.anthropic) providers.push({ id: 'env_anthropic', label: 'Claude (Server)', provider: 'anthropic', model: DEFAULT_MODELS.anthropic });
  // Also include user's own keys
  const { rows } = await pool.query('SELECT id, label, provider, model FROM ai_keys WHERE user_id = $1', [req.user.id]);
  rows.forEach(r => providers.push({ id: String(r.id), label: r.label, provider: r.provider, model: r.model }));
  res.json(providers);
});

app.post('/api/ai/process', authMiddleware, async (req, res) => {
  const { key_id, prompt, image_base64, system_prompt } = req.body;

  // Resolve the API key — either from env vars or user's saved keys
  var keyRow;
  if (key_id === 'env_gemini' && ENV_KEYS.gemini) {
    keyRow = { provider: 'gemini', model: DEFAULT_MODELS.gemini, api_key: ENV_KEYS.gemini };
  } else if (key_id === 'env_anthropic' && ENV_KEYS.anthropic) {
    keyRow = { provider: 'anthropic', model: DEFAULT_MODELS.anthropic, api_key: ENV_KEYS.anthropic };
  } else {
    const { rows: keys } = await pool.query('SELECT * FROM ai_keys WHERE id = $1 AND user_id = $2', [key_id, req.user.id]);
    if (!keys[0]) return res.status(400).json({ error: 'No API key selected. Add one in Settings or ask admin to set env vars.' });
    keyRow = keys[0];
  }

  try {
    if (keyRow.provider === 'gemini') {
      const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/${keyRow.model}:generateContent?key=${keyRow.api_key}`;
      const parts = [{ text: prompt }];
      if (image_base64) {
        parts.push({ inlineData: { mimeType: 'image/jpeg', data: image_base64 } });
      }
      const payload = {
        contents: [{ parts }],
        generationConfig: { responseMimeType: 'application/json' }
      };
      if (system_prompt) {
        payload.systemInstruction = { parts: [{ text: system_prompt }] };
      }
      const r = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
      const data = await r.json();
      if (data.error) return res.status(400).json({ error: data.error.message });
      const text = data.candidates?.[0]?.content?.parts?.[0]?.text || '';
      res.json({ result: text });

    } else if (keyRow.provider === 'anthropic') {
      const messages = [{ role: 'user', content: [] }];
      if (image_base64) {
        messages[0].content.push({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: image_base64 } });
      }
      messages[0].content.push({ type: 'text', text: prompt });

      const payload = { model: keyRow.model, max_tokens: 4096, messages };
      if (system_prompt) payload.system = system_prompt;

      const r = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': keyRow.api_key,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify(payload)
      });
      const data = await r.json();
      if (data.error) return res.status(400).json({ error: data.error.message });
      const text = data.content?.[0]?.text || '';
      res.json({ result: text });
    } else {
      res.status(400).json({ error: 'Unknown provider' });
    }
  } catch (e) {
    console.error('AI proxy error:', e);
    res.status(500).json({ error: e.message });
  }
});

// ---- COMPANY PROFILE ----
app.get('/api/company', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM company_profile LIMIT 1');
  res.json(rows[0] || {});
});

app.put('/api/company', authMiddleware, async (req, res) => {
  const d = req.body;
  const { rows: existing } = await pool.query('SELECT id FROM company_profile LIMIT 1');
  if (existing[0]) {
    const fields = Object.keys(d).filter(k => k !== 'id');
    const sets = fields.map((f, i) => `${f} = $${i + 2}`);
    const vals = fields.map(f => d[f]);
    await pool.query(`UPDATE company_profile SET ${sets.join(', ')} WHERE id = $1`, [existing[0].id, ...vals]);
  } else {
    const fields = Object.keys(d).filter(k => k !== 'id');
    const cols = fields.join(', ');
    const placeholders = fields.map((_, i) => `$${i + 1}`).join(', ');
    const vals = fields.map(f => d[f]);
    await pool.query(`INSERT INTO company_profile (${cols}) VALUES (${placeholders})`, vals);
  }
  const { rows } = await pool.query('SELECT * FROM company_profile LIMIT 1');
  res.json(rows[0]);
});

// ---- SIGNATURES ----
app.get('/api/signature', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM signatures WHERE user_id = $1', [req.user.id]);
  res.json(rows[0] || {});
});

app.put('/api/signature', authMiddleware, async (req, res) => {
  const { signature_data } = req.body;
  await pool.query(
    `INSERT INTO signatures (user_id, signature_data, updated_at) VALUES ($1, $2, NOW())
     ON CONFLICT (user_id) DO UPDATE SET signature_data = $2, updated_at = NOW()`,
    [req.user.id, signature_data]
  );
  res.json({ ok: true });
});

// ---- PHOTOS ----
app.post('/api/photos', authMiddleware, async (req, res) => {
  const { site_id, board_id, circuit_id, observation_id, data, caption } = req.body;
  const { rows } = await pool.query(
    `INSERT INTO photos (site_id, board_id, circuit_id, observation_id, data, caption, created_by) VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING id, caption, created_at`,
    [site_id, board_id, circuit_id, observation_id, data, caption, req.user.id]
  );
  res.json(rows[0]);
});

app.get('/api/photos', authMiddleware, async (req, res) => {
  const { site_id, board_id } = req.query;
  let q = 'SELECT id, site_id, board_id, circuit_id, observation_id, caption, created_at FROM photos WHERE 1=1';
  const vals = [];
  if (site_id) { vals.push(site_id); q += ` AND site_id = $${vals.length}`; }
  if (board_id) { vals.push(board_id); q += ` AND board_id = $${vals.length}`; }
  q += ' ORDER BY created_at DESC';
  const { rows } = await pool.query(q, vals);
  res.json(rows);
});

app.get('/api/photos/:id', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM photos WHERE id = $1', [req.params.id]);
  if (!rows[0]) return res.status(404).json({ error: 'Not found' });
  res.json(rows[0]);
});

app.delete('/api/photos/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM photos WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// ---- FILL N/A (EasyCert-style) ----
// Fill all empty circuit fields with N/A for a board
app.post('/api/boards/:id/fill-na', authMiddleware, async (req, res) => {
  const boardId = req.params.id;
  const { rows: circuits } = await pool.query('SELECT * FROM circuits WHERE board_id = $1', [boardId]);
  const fillableFields = [
    'wiring_type', 'ref_method', 'num_points', 'live_mm', 'cpc_mm',
    'ocpd_bsen', 'ocpd_type', 'ocpd_rating', 'ocpd_breaking_cap', 'max_zs',
    'rcd_bsen', 'rcd_type', 'rcd_idn', 'rcd_rating_a',
    'r1', 'rn', 'r2', 'r1r2', 'r1r2_or_r2', 'r2_ring',
    'ir_ll', 'ir_le', 'polarity', 'zs_measured',
    'rcd_time_x1', 'rcd_time_x5', 'rcd_test_button', 'afdd_test', 'remarks'
  ];
  let filled = 0;
  for (const c of circuits) {
    const updates = {};
    fillableFields.forEach(f => {
      if (!c[f] || c[f].trim() === '') updates[f] = 'N/A';
    });
    if (Object.keys(updates).length) {
      const sets = Object.keys(updates).map((f, i) => `${f} = $${i + 2}`);
      const vals = Object.values(updates);
      await pool.query(`UPDATE circuits SET ${sets.join(', ')} WHERE id = $1`, [c.id, ...vals]);
      filled += Object.keys(updates).length;
    }
  }
  res.json({ ok: true, filled });
});

// ---- VERIFY ZS (EasyCert-style) ----
// Check all circuits in a board against max Zs tables
app.get('/api/boards/:id/verify-zs', authMiddleware, async (req, res) => {
  const { rows: circuits } = await pool.query('SELECT * FROM circuits WHERE board_id = $1 ORDER BY sort_order, id', [req.params.id]);
  const results = [];
  for (const c of circuits) {
    if (!c.zs_measured || c.zs_measured === 'N/A' || !c.ocpd_bsen || !c.ocpd_type || !c.ocpd_rating) {
      results.push({ circuit_id: c.id, number: c.number, status: 'skip', reason: 'Missing data' });
      continue;
    }
    const bsen = c.ocpd_bsen.replace(/\s/g, '');
    const table = MAX_ZS_TABLE[bsen] || MAX_ZS_TABLE['60898'];
    const curve = table ? table[c.ocpd_type] : null;
    const maxZs = curve ? curve[parseInt(c.ocpd_rating)] : null;
    if (!maxZs) {
      results.push({ circuit_id: c.id, number: c.number, status: 'unknown', reason: 'No Zs table entry for ' + bsen + ' ' + c.ocpd_type + ' ' + c.ocpd_rating });
      continue;
    }
    const measured = parseFloat(c.zs_measured);
    const limit80 = Math.round(maxZs * 0.8 * 100) / 100;
    const pass = measured <= limit80;
    results.push({
      circuit_id: c.id, number: c.number,
      status: pass ? 'pass' : 'fail',
      measured, max_zs: maxZs, max_zs_80: limit80
    });
    // Auto-populate max_zs field on circuit
    await pool.query('UPDATE circuits SET max_zs = $1 WHERE id = $2', [String(maxZs), c.id]);
  }
  res.json(results);
});

// ---- MAX ZS LOOKUP ----
const MAX_ZS_TABLE = {
  // BS(EN) 60898 — 0.4s disconnection
  '60898': {
    'B': { 6: 7.28, 10: 4.37, 16: 2.73, 20: 2.19, 25: 1.75, 32: 1.37, 40: 1.09, 50: 0.87, 63: 0.69, 80: 0.55, 100: 0.44, 125: 0.35 },
    'C': { 6: 3.64, 10: 2.19, 16: 1.37, 20: 1.09, 25: 0.87, 32: 0.68, 40: 0.55, 50: 0.44, 63: 0.35, 80: 0.27, 100: 0.22, 125: 0.17 },
    'D': { 6: 1.82, 10: 1.09, 16: 0.68, 20: 0.55, 25: 0.44, 32: 0.34, 40: 0.27, 50: 0.22, 63: 0.17, 80: 0.14, 100: 0.11, 125: 0.09 }
  },
  // 61009 RCBOs same as 60898
  '61009': null, // alias
  // BS 3036 fuses
  '3036': {
    'gG': { 5: 9.58, 15: 3.20, 20: 2.30, 30: 1.44, 45: 0.96 }
  },
  // BS 88-2 / BS(EN) 60269 fuses
  '88-2': {
    'gG': { 6: 7.67, 10: 4.60, 16: 2.87, 20: 2.30, 25: 1.77, 32: 1.37, 40: 1.04, 50: 0.82, 63: 0.57, 80: 0.43, 100: 0.32, 125: 0.25 }
  },
  // BS(EN) 60947-2 MCCBs (instantaneous trip = 10x for Type C-equivalent)
  '60947-2': {
    'C': { 16: 1.37, 20: 1.09, 25: 0.87, 32: 0.68, 40: 0.55, 50: 0.44, 63: 0.35, 80: 0.27, 100: 0.22, 125: 0.17, 160: 0.14, 200: 0.11, 250: 0.09 },
    'D': { 16: 0.68, 20: 0.55, 25: 0.44, 32: 0.34, 40: 0.27, 50: 0.22, 63: 0.17, 80: 0.14, 100: 0.11, 125: 0.09, 160: 0.07, 200: 0.05, 250: 0.04 }
  },
  // BS(EN) 60947-3 Isolators / switch-disconnectors (no Zs — they don't provide fault protection)
  '60947-3': {}
};
MAX_ZS_TABLE['61009'] = MAX_ZS_TABLE['60898'];
MAX_ZS_TABLE['60269'] = MAX_ZS_TABLE['88-2'];

app.get('/api/max-zs', (req, res) => {
  const { bsen, type, rating } = req.query;
  const table = MAX_ZS_TABLE[bsen] || MAX_ZS_TABLE['60898'];
  if (!table) return res.json({ max_zs: null });
  const curve = table[type];
  if (!curve) return res.json({ max_zs: null });
  const zs = curve[parseInt(rating)];
  res.json({ max_zs: zs || null, max_zs_80: zs ? Math.round(zs * 0.8 * 100) / 100 : null });
});

// ---- REPORT GENERATION ----
app.get('/api/sites/:id/report', authMiddleware, async (req, res) => {
  const siteId = req.params.id;
  const { rows: [site] } = await pool.query('SELECT * FROM sites WHERE id = $1', [siteId]);
  if (!site) return res.status(404).json({ error: 'Site not found' });
  const { rows: siteBoards } = await pool.query('SELECT * FROM boards WHERE site_id = $1 ORDER BY sort_order, id', [siteId]);
  const { rows: obs } = await pool.query('SELECT * FROM observations WHERE site_id = $1 ORDER BY item_no', [siteId]);
  const { rows: insp } = await pool.query('SELECT * FROM inspections WHERE site_id = $1 ORDER BY item_ref', [siteId]);

  // Build circuits per board
  const boardCircuits = {};
  for (const b of siteBoards) {
    const { rows } = await pool.query('SELECT * FROM circuits WHERE board_id = $1 ORDER BY sort_order, id', [b.id]);
    boardCircuits[b.id] = rows;
  }

  const e = (s) => (s || '').toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const hasC1C2FI = obs.some(o => ['C1', 'C2', 'FI'].includes(o.code));

  let html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>EICR - ${e(site.name)}</title>
<style>
  @page { size: A4; margin: 15mm; }
  body { font-family: Arial, sans-serif; font-size: 10px; line-height: 1.4; color: #000; }
  h1 { font-size: 16px; text-align: center; margin: 0 0 10px; }
  h2 { font-size: 12px; background: #e8e8e8; padding: 4px 8px; margin: 15px 0 5px; page-break-after: avoid; }
  table { border-collapse: collapse; width: 100%; margin: 5px 0; }
  th, td { border: 1px solid #999; padding: 3px 5px; text-align: left; font-size: 9px; }
  th { background: #f0f0f0; font-weight: bold; }
  .field-label { font-weight: bold; background: #f8f8f8; width: 30%; }
  .obs-c1 { background: #fdd; } .obs-c2 { background: #fed; } .obs-fi { background: #def; }
  .pass { color: green; } .fail { color: red; font-weight: bold; }
  .page-break { page-break-before: always; }
  @media print { .no-print { display: none; } }
</style></head><body>
<div class="no-print" style="text-align:center;padding:10px;background:#3b82f6;color:white;margin-bottom:20px">
  <button onclick="window.print()" style="padding:10px 30px;font-size:14px;font-weight:bold;cursor:pointer">Print / Save as PDF</button>
</div>`;

  // Page 1: Report Details
  html += `<h1>ELECTRICAL INSTALLATION CONDITION REPORT</h1>
<p style="text-align:center;font-size:9px">Requirements for Electrical Installations - BS 7671:2018+A2:2022</p>
<p style="text-align:right;font-size:9px">Report Ref: ${e(site.report_ref)}</p>
<table>
  <tr><td class="field-label" colspan="2">DETAILS OF THE PERSON ORDERING THE REPORT</td></tr>
  <tr><td class="field-label">Client:</td><td>${e(site.client_name)}</td></tr>
  <tr><td class="field-label">Address:</td><td>${e(site.client_address)} ${e(site.postcode)}</td></tr>
  <tr><td class="field-label">Telephone:</td><td>${e(site.client_tel)}</td></tr>
</table>
<table>
  <tr><td class="field-label" colspan="2">REASON FOR PRODUCING THIS REPORT</td></tr>
  <tr><td class="field-label">Purpose:</td><td>${e(site.purpose)}</td></tr>
</table>
<table>
  <tr><td class="field-label" colspan="2">DETAILS OF THE INSTALLATION</td></tr>
  <tr><td class="field-label">Installation Address:</td><td>${e(site.address)} ${e(site.postcode)}</td></tr>
  <tr><td class="field-label">Description:</td><td>${e(site.description)}</td></tr>
  <tr><td class="field-label">Estimated age of wiring:</td><td>${e(site.wiring_age)} years</td></tr>
  <tr><td class="field-label">Additions/Alterations:</td><td>${e(site.additions)}${site.additions_age ? ' (' + site.additions_age + ' years)' : ''}</td></tr>
  <tr><td class="field-label">Date of last inspection:</td><td>${e(site.last_inspection_date)}</td></tr>
  <tr><td class="field-label">Records available:</td><td>${e(site.records_available)}</td></tr>
</table>
<table>
  <tr><td class="field-label" colspan="2">EXTENT AND LIMITATIONS</td></tr>
  <tr><td class="field-label">Extent:</td><td>${e(site.extent)}</td></tr>
  <tr><td class="field-label">Agreed limitations:</td><td>${e(site.agreed_limitations)}</td></tr>
  <tr><td class="field-label">Operational limitations:</td><td>${e(site.operational_limitations)}</td></tr>
</table>
<table>
  <tr><td class="field-label" colspan="2">SUMMARY</td></tr>
  <tr><td class="field-label">Overall assessment:</td><td style="font-size:12px;font-weight:bold;color:${hasC1C2FI ? 'red' : 'green'}">${hasC1C2FI ? 'UNSATISFACTORY' : 'SATISFACTORY'}</td></tr>
  <tr><td class="field-label">Next inspection by:</td><td>${e(site.next_inspection_date)}</td></tr>
</table>`;

  // Observations page
  if (obs.length) {
    html += `<div class="page-break"></div><h2>OBSERVATIONS AND RECOMMENDATIONS</h2>
    <table><thead><tr><th>#</th><th>Observation</th><th>Location</th><th>Code</th></tr></thead><tbody>`;
    obs.forEach(o => {
      const cls = o.code === 'C1' ? 'obs-c1' : o.code === 'C2' ? 'obs-c2' : o.code === 'FI' ? 'obs-fi' : '';
      html += `<tr class="${cls}"><td>${o.item_no}</td><td>${e(o.description)}</td><td>${e(o.location)}</td><td><strong>${e(o.code)}</strong></td></tr>`;
    });
    html += `</tbody></table>`;
  }

  // Supply Characteristics
  html += `<div class="page-break"></div><h2>SUPPLY CHARACTERISTICS AND EARTHING ARRANGEMENTS</h2>
  <table>
    <tr><td class="field-label">Earthing:</td><td>${e(site.earthing_type)}</td><td class="field-label">Phases:</td><td>${e(site.num_phases)}</td></tr>
    <tr><td class="field-label">Nominal Voltage:</td><td>${e(site.nominal_voltage)}V</td><td class="field-label">Frequency:</td><td>${e(site.nominal_frequency)}Hz</td></tr>
    <tr><td class="field-label">Ipf at origin:</td><td>${e(site.ipf_at_origin)} kA</td><td class="field-label">Ze:</td><td>${e(site.ze_at_origin)} &Omega;</td></tr>
    <tr><td class="field-label">Supply device BS(EN):</td><td>${e(site.supply_device_bsen)}</td><td class="field-label">Rating:</td><td>${e(site.supply_device_rating)}A</td></tr>
  </table>`;

  // Inspection Schedule
  if (insp.length) {
    html += `<div class="page-break"></div><h2>SCHEDULE OF INSPECTIONS</h2>
    <table><thead><tr><th>Item</th><th>Description</th><th>Outcome</th></tr></thead><tbody>`;
    // We'll add a helper to get inspection outcome
    const inspMap = {};
    insp.forEach(i => { inspMap[i.item_ref] = i.outcome; });
    // The full inspection items would be rendered from the INSPECTION_ITEMS constant
    insp.forEach(i => {
      html += `<tr><td>${e(i.item_ref)}</td><td></td><td>${e(i.outcome)}</td></tr>`;
    });
    html += `</tbody></table>`;
  }

  // Per-board test results
  for (const b of siteBoards) {
    const ccts = boardCircuits[b.id] || [];
    html += `<div class="page-break"></div>
    <h2>DB: ${e(b.ref)} — ${e(b.location)}</h2>
    <table>
      <tr><td class="field-label">Supplied from:</td><td>${e(b.supplied_from)}</td><td class="field-label">Ze at DB:</td><td>${e(b.ze)}&Omega;</td></tr>
      <tr><td class="field-label">OCPD BS(EN):</td><td>${e(b.dist_bsen)}</td><td class="field-label">Ipf at DB:</td><td>${e(b.ipf)} kA</td></tr>
    </table>`;

    if (ccts.length) {
      html += `<table><thead><tr>
        <th>#</th><th>Description</th><th>Type</th><th>BS(EN)</th><th>Curve</th><th>In(A)</th>
        <th>Live</th><th>CPC</th><th>R1+R2</th><th>R2</th><th>IR L-L</th><th>IR L-E</th>
        <th>Pol</th><th>Zs</th><th>Max Zs</th><th>RCD x1</th><th>RCD x5</th><th>Remarks</th>
      </tr></thead><tbody>`;
      ccts.forEach(c => {
        const zsNum = parseFloat(c.zs_measured);
        const maxNum = parseFloat(c.max_zs);
        const zsFail = zsNum && maxNum && zsNum > maxNum * 0.8;
        html += `<tr>
          <td>${e(c.number)}</td><td>${e(c.description)}</td><td>${e(c.wiring_type)}</td>
          <td>${e(c.ocpd_bsen)}</td><td>${e(c.ocpd_type)}</td><td>${e(c.ocpd_rating)}</td>
          <td>${e(c.live_mm)}</td><td>${e(c.cpc_mm)}</td>
          <td>${e(c.r1r2)}</td><td>${e(c.r2_ring)}</td>
          <td>${e(c.ir_ll)}</td><td>${e(c.ir_le)}</td>
          <td>${e(c.polarity)}</td>
          <td class="${zsFail ? 'fail' : 'pass'}">${e(c.zs_measured)}</td>
          <td>${e(c.max_zs)}</td>
          <td>${e(c.rcd_time_x1)}</td><td>${e(c.rcd_time_x5)}</td>
          <td>${e(c.remarks)}</td>
        </tr>`;
      });
      html += `</tbody></table>`;
    }
  }

  html += `</body></html>`;
  res.setHeader('Content-Type', 'text/html');
  res.send(html);
});

// ---- BOOT ----
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

initDB().then(() => {
  app.listen(PORT, () => console.log(`EICR Field Pro running on port ${PORT}`));
}).catch(err => {
  console.error('DB init failed:', err);
  process.exit(1);
});
