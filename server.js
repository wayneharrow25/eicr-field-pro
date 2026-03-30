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

      CREATE TABLE IF NOT EXISTS board_photos (
        id SERIAL PRIMARY KEY,
        board_id INT REFERENCES boards(id) ON DELETE CASCADE,
        photo TEXT NOT NULL,
        caption TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS testers (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL,
        position TEXT DEFAULT 'Electrician',
        ecs_number TEXT,
        napit_number TEXT,
        niceic_number TEXT,
        jib_number TEXT,
        other_qual TEXT,
        signature_data TEXT,
        is_qs BOOLEAN DEFAULT false,
        created_at TIMESTAMPTZ DEFAULT NOW()
      );

      CREATE TABLE IF NOT EXISTS site_testers (
        id SERIAL PRIMARY KEY,
        site_id INT REFERENCES sites(id) ON DELETE CASCADE,
        tester_id INT REFERENCES testers(id) ON DELETE CASCADE,
        role TEXT DEFAULT 'tester',
        date_signed TEXT,
        UNIQUE(site_id, tester_id)
      );
    `);

    // Add missing columns to sites table (safe — IF NOT EXISTS)
    const siteColumns = [
      'client_phone TEXT', 'client_email TEXT', 'occupier TEXT', 'desc_premises TEXT',
      'estimated_age TEXT', 'evidence_alteration TEXT', 'date_last_inspection TEXT',
      'date_inspection TEXT', 'next_inspection TEXT', 'next_interval TEXT', 'sampling TEXT',
      'limitations TEXT', 'supply_system TEXT', 'supply_protective TEXT', 'supply_protective_rating TEXT',
      'nominal_voltage_ll TEXT', 'nominal_voltage_ln TEXT', 'nominal_freq TEXT',
      'pscc TEXT', 'pfc_confirmed TEXT', 'ze TEXT',
      'main_switch_type TEXT', 'main_switch_rating TEXT', 'main_switch_voltage TEXT',
      'bonding_structural TEXT', 'bonding_condition TEXT'
    ];
    for (const col of siteColumns) {
      const colName = col.split(' ')[0];
      await client.query(`ALTER TABLE sites ADD COLUMN IF NOT EXISTS ${col}`).catch(() => {});
    }

    // Add missing columns to observations table
    const obsColumns = ['materials TEXT'];
    for (const col of obsColumns) {
      await client.query(`ALTER TABLE observations ADD COLUMN IF NOT EXISTS ${col}`).catch(() => {});
    }

    // Add missing columns to boards table
    const boardColumns = [
      'ocpd_bsen TEXT', 'ocpd_type TEXT', 'ocpd_rating TEXT', 'poles TEXT'
    ];
    for (const col of boardColumns) {
      await client.query(`ALTER TABLE boards ADD COLUMN IF NOT EXISTS ${col}`).catch(() => {});
    }

    // Seed default users if none exist
    const { rows } = await client.query('SELECT COUNT(*) FROM users');
    if (parseInt(rows[0].count) === 0) {
      const hash1 = await bcrypt.hash('wayne1', 10);
      const hash2 = await bcrypt.hash('john1', 10);
      await client.query(`INSERT INTO users (username, password_hash, name, role) VALUES ('wayne', $1, 'Wayne Harrow', 'admin')`, [hash1]);
      await client.query(`INSERT INTO users (username, password_hash, name, role) VALUES ('john', $1, 'John Harrow', 'inspector')`, [hash2]);
      console.log('Default users created: wayne/wayne1, john/john1');
    }

    // Seed company profile if none exists
    const { rows: compRows } = await client.query('SELECT COUNT(*) FROM company_profile');
    if (parseInt(compRows[0].count) === 0) {
      await client.query(`INSERT INTO company_profile (name, address, postcode, tel, registration_body, registration_number) VALUES ($1, $2, $3, $4, $5, $6)`, [
        'Expert Energy Group',
        'Unit 21, Industrial Estate, Old Church Road, East Hanningfield, Essex',
        'CM3 8AB',
        '08000016724',
        'NAPIT',
        '64559'
      ]);
      console.log('Company profile seeded');
    }

    // Seed testers if none exist
    const { rows: testerRows } = await client.query('SELECT COUNT(*) FROM testers');
    if (parseInt(testerRows[0].count) === 0) {
      await client.query(`INSERT INTO testers (name, position, napit_number, is_qs) VALUES ($1, $2, $3, $4)`, [
        'Wayne Harrow', 'Qualified Supervisor', '64559', true
      ]);
      await client.query(`INSERT INTO testers (name, position) VALUES ($1, $2)`, [
        'John Harrow', 'Electrician'
      ]);
      console.log('Default testers seeded: Wayne (QS), John (Electrician)');
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
  const { rows } = await pool.query(
    `SELECT b.*, COALESCE(c.cnt, 0)::int AS circuit_count
     FROM boards b
     LEFT JOIN (SELECT board_id, COUNT(*) AS cnt FROM circuits GROUP BY board_id) c
       ON b.id = c.board_id
     WHERE b.site_id = $1 ORDER BY b.sort_order, b.id`,
    [req.params.id]
  );
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

// Apply instruments to all boards in a site
app.put('/api/sites/:id/boards/instruments', authMiddleware, async (req, res) => {
  const d = req.body;
  const instrumentCols = ['instruments_multi', 'instruments_ir', 'instruments_cont', 'instruments_earth', 'instruments_loop', 'instruments_rcd'];
  const fields = instrumentCols.filter(k => d[k] !== undefined);
  if (!fields.length) return res.json({ count: 0 });
  const sets = fields.map((f, i) => `${f} = $${i + 2}`);
  const vals = fields.map(f => d[f]);
  const result = await pool.query(`UPDATE boards SET ${sets.join(', ')} WHERE site_id = $1`, [req.params.id, ...vals]);
  res.json({ count: result.rowCount });
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

// ---- BOARD PHOTOS ----
app.get('/api/boards/:id/photos', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT id, caption, created_at FROM board_photos WHERE board_id = $1 ORDER BY id', [req.params.id]);
  res.json(rows);
});

app.get('/api/board-photos/:id', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM board_photos WHERE id = $1', [req.params.id]);
  if (!rows.length) return res.status(404).json({ error: 'Not found' });
  res.json(rows[0]);
});

app.post('/api/boards/:id/photos', authMiddleware, async (req, res) => {
  const { photo, caption } = req.body;
  if (!photo) return res.status(400).json({ error: 'photo required' });
  const { rows } = await pool.query(
    'INSERT INTO board_photos (board_id, photo, caption) VALUES ($1, $2, $3) RETURNING *',
    [req.params.id, photo, caption || '']
  );
  res.json(rows[0]);
});

app.delete('/api/board-photos/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM board_photos WHERE id = $1', [req.params.id]);
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

// Map frontend field names to actual DB column names
const CIRCUIT_FIELD_MAP = {
  points: 'num_points',
  breaking_cap: 'ocpd_breaking_cap',
  rcd_rating: 'rcd_rating_a',
  r2_measured: 'r1r2_or_r2',
  ir_test_v: 'test_voltage',
  rcd_test_btn: 'rcd_test_button'
};
function mapCircuitFields(obj) {
  const mapped = {};
  for (const [k, v] of Object.entries(obj)) {
    const dbCol = CIRCUIT_FIELD_MAP[k] || k;
    mapped[dbCol] = v;
  }
  return mapped;
}

app.put('/api/boards/:id/circuits/bulk', authMiddleware, async (req, res) => {
  const updates = req.body; // [{id or number, ...fields}] or with ?replace=1 to delete-all-then-insert
  const boardId = req.params.id;
  const replaceMode = req.query.replace === '1';

  if (replaceMode) {
    // Delete all existing circuits and insert fresh — used by renumber to avoid duplication
    await pool.query('DELETE FROM circuits WHERE board_id = $1', [boardId]);
    for (const raw of updates) {
      const u = mapCircuitFields(raw);
      const fields = Object.keys(u).filter(k => k !== 'id' && k !== 'board_id');
      const cols = ['board_id', ...fields];
      const placeholders = cols.map((_, i) => `$${i + 1}`);
      const vals = [boardId, ...fields.map(f => u[f])];
      await pool.query(`INSERT INTO circuits (${cols.join(',')}) VALUES (${placeholders.join(',')})`, vals);
    }
  } else {
    for (const raw of updates) {
      const u = mapCircuitFields(raw);
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
  }
  const { rows } = await pool.query('SELECT * FROM circuits WHERE board_id = $1 ORDER BY sort_order, id', [boardId]);
  res.json(rows);
});

// Bulk delete circuits by ID list
app.post('/api/boards/:id/circuits/bulk-delete', authMiddleware, async (req, res) => {
  const ids = req.body.ids || [];
  if (!ids.length) return res.json({ deleted: 0 });
  const placeholders = ids.map((_, i) => `$${i + 2}`).join(',');
  await pool.query(`DELETE FROM circuits WHERE board_id = $1 AND id IN (${placeholders})`, [req.params.id, ...ids]);
  const { rows } = await pool.query('SELECT * FROM circuits WHERE board_id = $1 ORDER BY sort_order, id', [req.params.id]);
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

app.put('/api/observations/:id', authMiddleware, async (req, res) => {
  const { code, description, location, photo } = req.body;
  const fields = [];
  const vals = [];
  let idx = 1;
  if (code !== undefined) { fields.push('code = $' + idx++); vals.push(code); }
  if (description !== undefined) { fields.push('description = $' + idx++); vals.push(description); }
  if (location !== undefined) { fields.push('location = $' + idx++); vals.push(location); }
  if (photo !== undefined) { fields.push('photo = $' + idx++); vals.push(photo); }
  if (!fields.length) return res.json(req.body);
  vals.push(req.params.id);
  const { rows } = await pool.query(
    'UPDATE observations SET ' + fields.join(', ') + ' WHERE id = $' + idx + ' RETURNING *',
    vals
  );
  res.json(rows[0] || {});
});

app.delete('/api/observations/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM observations WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// ---- AI PROXY ----
// Endpoint to list available AI providers (including env var ones)
app.get('/api/ai/providers', authMiddleware, async (req, res) => {
  const providers = [];
  // Gemini key works for ALL models — offer Flash, Pro and legacy
  if (ENV_KEYS.gemini) {
    providers.push({ id: 'env_gemini_flash', label: 'Gemini 2.5 Flash (fast)', provider: 'gemini', model: 'gemini-2.5-flash' });
    providers.push({ id: 'env_gemini_pro', label: 'Gemini 2.5 Pro (smart)', provider: 'gemini', model: 'gemini-2.5-pro' });
    providers.push({ id: 'env_gemini_31_pro', label: 'Gemini 3.1 Pro (latest)', provider: 'gemini', model: 'gemini-3.1-pro' });
    providers.push({ id: 'env_gemini_3_pro', label: 'Gemini 3.0 Pro', provider: 'gemini', model: 'gemini-3.0-pro' });
    providers.push({ id: 'env_gemini_flash_2', label: 'Gemini 2.0 Flash', provider: 'gemini', model: 'gemini-2.0-flash' });
    providers.push({ id: 'env_gemini_pro_15', label: 'Gemini 1.5 Pro', provider: 'gemini', model: 'gemini-1.5-pro' });
  }
  if (ENV_KEYS.anthropic) providers.push({ id: 'env_anthropic', label: 'Claude Sonnet', provider: 'anthropic', model: DEFAULT_MODELS.anthropic });
  // Also include user's own keys
  const { rows } = await pool.query('SELECT id, label, provider, model FROM ai_keys WHERE user_id = $1', [req.user.id]);
  rows.forEach(r => providers.push({ id: String(r.id), label: r.label, provider: r.provider, model: r.model }));
  res.json(providers);
});

// Base EICR context injected into ALL AI requests automatically
const EICR_BASE_CONTEXT = `You are an AI assistant for Expert Energy Group, helping qualified electricians complete BS 7671:2018+A2:2022 Electrical Installation Condition Reports (EICRs) for commercial properties (schools, offices, industrial).

KEY KNOWLEDGE:
- Protective devices: MCBs (BS EN 60898), RCBOs (BS EN 61009), RCDs (BS EN 61008), MCCBs (BS EN 60947-2), HRC fuses (BS EN 88-2)
- Curve types: B (general/domestic), C (motors/fluorescent), D (high inrush transformers)
- RCD types: AC (standard), A (pulsating DC), B (smooth DC), S (selective/time-delayed)
- Common IΔn ratings: 30mA (personal protection), 100mA, 300mA (fire protection)
- Earthing systems: TN-S (separate earth), TN-C-S (PME/combined), TT (earth electrode), IT
- Supply types: Single phase (230V), Three phase (400V)
- Classification codes: C1 (danger present), C2 (potentially dangerous), C3 (improvement recommended), FI (further investigation required)
- Test readings: Ze (external earth fault loop impedance), Zs (earth fault loop impedance), R1+R2 (continuity), Insulation resistance (IR in MΩ), RCD trip times (ms)
- Max Zs values use 80% rule for on-site measurements
- Wiring types: Flat T+E (twin & earth), SWA (steel wire armoured), MICC (mineral insulated), Singles in conduit/trunking
- Common CSA sizes: 1.0, 1.5, 2.5, 4.0, 6.0, 10.0, 16.0, 25.0mm²

Always use correct UK electrical terminology and BS 7671 conventions.`;

app.post('/api/ai/process', authMiddleware, async (req, res) => {
  const { key_id, prompt, image_base64, system_prompt } = req.body;

  // Combine base EICR context with any specific system prompt
  const fullSystemPrompt = system_prompt
    ? EICR_BASE_CONTEXT + '\n\n' + system_prompt
    : EICR_BASE_CONTEXT;

  // Resolve the API key — either from env vars or user's saved keys
  var keyRow;
  if (key_id === 'env_gemini_flash' && ENV_KEYS.gemini) {
    keyRow = { provider: 'gemini', model: 'gemini-2.5-flash', api_key: ENV_KEYS.gemini };
  } else if (key_id === 'env_gemini_pro' && ENV_KEYS.gemini) {
    keyRow = { provider: 'gemini', model: 'gemini-2.5-pro', api_key: ENV_KEYS.gemini };
  } else if (key_id === 'env_gemini_31_pro' && ENV_KEYS.gemini) {
    keyRow = { provider: 'gemini', model: 'gemini-3.1-pro', api_key: ENV_KEYS.gemini };
  } else if (key_id === 'env_gemini_3_pro' && ENV_KEYS.gemini) {
    keyRow = { provider: 'gemini', model: 'gemini-3.0-pro', api_key: ENV_KEYS.gemini };
  } else if (key_id === 'env_gemini_flash_2' && ENV_KEYS.gemini) {
    keyRow = { provider: 'gemini', model: 'gemini-2.0-flash', api_key: ENV_KEYS.gemini };
  } else if (key_id === 'env_gemini_pro_15' && ENV_KEYS.gemini) {
    keyRow = { provider: 'gemini', model: 'gemini-1.5-pro', api_key: ENV_KEYS.gemini };
  } else if (key_id === 'env_gemini' && ENV_KEYS.gemini) {
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
      const payload = { contents: [{ parts }] };
      payload.systemInstruction = { parts: [{ text: fullSystemPrompt }] };
      // Support multi-turn conversation
      if (req.body.history && req.body.history.length > 0) {
        payload.contents = req.body.history.concat(payload.contents);
      }
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 120000); // 2 min timeout
      const r = await fetch(apiUrl, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload), signal: controller.signal });
      clearTimeout(timeout);
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
      payload.system = fullSystemPrompt;

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

// ---- TESTERS ----
app.get('/api/testers', authMiddleware, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM testers ORDER BY is_qs DESC, name');
  res.json(rows);
});

app.post('/api/testers', authMiddleware, async (req, res) => {
  const { name, position, ecs_number, napit_number, niceic_number, jib_number, other_qual, signature_data, is_qs } = req.body;
  const { rows } = await pool.query(
    `INSERT INTO testers (name, position, ecs_number, napit_number, niceic_number, jib_number, other_qual, signature_data, is_qs)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9) RETURNING *`,
    [name, position || 'Electrician', ecs_number, napit_number, niceic_number, jib_number, other_qual, signature_data, is_qs || false]
  );
  res.json(rows[0]);
});

app.put('/api/testers/:id', authMiddleware, async (req, res) => {
  const d = req.body;
  const fields = Object.keys(d).filter(k => k !== 'id' && k !== 'created_at');
  if (!fields.length) return res.status(400).json({ error: 'No fields' });
  const sets = fields.map((f, i) => `${f} = $${i + 2}`);
  const vals = fields.map(f => d[f]);
  await pool.query(`UPDATE testers SET ${sets.join(', ')} WHERE id = $1`, [req.params.id, ...vals]);
  const { rows } = await pool.query('SELECT * FROM testers WHERE id = $1', [req.params.id]);
  res.json(rows[0]);
});

app.delete('/api/testers/:id', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM testers WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// Site testers — assign testers to a site
app.get('/api/sites/:id/testers', authMiddleware, async (req, res) => {
  const { rows } = await pool.query(
    `SELECT st.*, t.name, t.position, t.ecs_number, t.napit_number, t.niceic_number, t.jib_number, t.other_qual, t.is_qs, t.signature_data
     FROM site_testers st JOIN testers t ON st.tester_id = t.id WHERE st.site_id = $1 ORDER BY t.is_qs DESC, t.name`,
    [req.params.id]
  );
  res.json(rows);
});

app.post('/api/sites/:id/testers', authMiddleware, async (req, res) => {
  const { tester_id, role, date_signed } = req.body;
  const { rows } = await pool.query(
    `INSERT INTO site_testers (site_id, tester_id, role, date_signed) VALUES ($1,$2,$3,$4)
     ON CONFLICT (site_id, tester_id) DO UPDATE SET role = $3, date_signed = $4 RETURNING *`,
    [req.params.id, tester_id, role || 'tester', date_signed]
  );
  res.json(rows[0]);
});

app.delete('/api/sites/:id/testers/:testerId', authMiddleware, async (req, res) => {
  await pool.query('DELETE FROM site_testers WHERE site_id = $1 AND tester_id = $2', [req.params.id, req.params.testerId]);
  res.json({ ok: true });
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
// Report accepts token via query param (for new tab/PDF download)
app.get('/api/sites/:id/report', (req, res, next) => {
  if (req.query.token) req.headers.authorization = 'Bearer ' + req.query.token;
  authMiddleware(req, res, next);
}, async (req, res) => {
  const siteId = req.params.id;
  const { rows: [site] } = await pool.query('SELECT * FROM sites WHERE id = $1', [siteId]);
  if (!site) return res.status(404).json({ error: 'Site not found' });
  const { rows: siteBoards } = await pool.query('SELECT * FROM boards WHERE site_id = $1 ORDER BY sort_order, id', [siteId]);
  const { rows: obs } = await pool.query('SELECT * FROM observations WHERE site_id = $1 ORDER BY item_no', [siteId]);
  const { rows: insp } = await pool.query('SELECT * FROM inspections WHERE site_id = $1 ORDER BY item_ref', [siteId]);
  const { rows: [company] } = await pool.query('SELECT * FROM company_profile LIMIT 1');
  const { rows: siteTestersRaw } = await pool.query(
    `SELECT st.*, t.name, t.position, t.ecs_number, t.napit_number, t.niceic_number, t.jib_number, t.other_qual, t.is_qs, t.signature_data
     FROM site_testers st JOIN testers t ON st.tester_id = t.id WHERE st.site_id = $1 ORDER BY t.is_qs DESC, t.name`, [siteId]
  );

  const boardCircuits = {};
  for (const b of siteBoards) {
    const { rows } = await pool.query('SELECT * FROM circuits WHERE board_id = $1 ORDER BY sort_order, id', [b.id]);
    boardCircuits[b.id] = rows;
  }

  const e = (s) => (s || '').toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const tick = (v) => (v === 'true' || v === true || v === 'Yes' || v === 'yes') ? '&#10003;' : '';
  const earthTick = (type, val) => (site.earthing_type || '').toLowerCase().replace(/[^a-z0-9]/g, '').includes(type.toLowerCase().replace(/[^a-z0-9]/g, '')) ? '&#10003;' : '';
  const hasC1C2FI = obs.some(o => ['C1', 'C2', 'FI'].includes(o.code));
  const comp = company || {};
  const qs = siteTestersRaw.find(t => t.is_qs) || {};
  const testers = siteTestersRaw.filter(t => !t.is_qs);
  const certNum = site.report_ref || '';
  const totalBoards = siteBoards.length;

  // Count pages for footer
  // Page 1: Main report, Page 2: Obs, Per-DB obs pages, General+Declaration, Supply, Particulars, Inspection schedule pages, Per-board Form 4 pages, Guidance
  // We'll use JS-based page counting in print
  let pageNum = 0;

  // Inspection schedule master list with BS 7671 regulation references
  const inspectionSchedule = [
    { ref: '1.0', desc: 'CONSUMER UNIT / DISTRIBUTION BOARD', items: [] },
    { ref: '1.1', desc: 'Adequacy of access and working space', reg: 'Reg 132.12' },
    { ref: '1.2', desc: 'Security of fixing', reg: '' },
    { ref: '1.3', desc: 'Condition of enclosure(s) in terms of damage and deterioration', reg: '' },
    { ref: '1.4', desc: 'Suitability of enclosure(s) for IP and fire ratings', reg: 'Reg 416.2, 421.1.6, 526.5' },
    { ref: '1.5', desc: 'Enclosure not damaged/deteriorated so as to impair safety', reg: 'Reg 514.4.2' },
    { ref: '1.6', desc: 'Presence of main linked switch (functional)', reg: 'Reg 537.1.4' },
    { ref: '1.7', desc: 'Operation of main switch (functional check)', reg: 'Reg 612.13.2' },
    { ref: '1.8', desc: 'Manual operation of circuit-breakers and RCDs to prove disconnection', reg: 'Reg 612.13.2' },
    { ref: '1.9', desc: 'Correct identification of each circuit (labelling)', reg: 'Reg 514.8.1, 514.9.1' },
    { ref: '1.10', desc: 'Presence of RCD quarterly test notice at or near origin', reg: 'Reg 514.12.2' },
    { ref: '1.11', desc: 'Presence of diagrams, charts or schedules at or near distribution board', reg: 'Reg 514.9.1' },
    { ref: '1.12', desc: 'Presence of non-standard (mixed) cable colour warning notice', reg: 'Reg 514.14' },
    { ref: '1.13', desc: 'Presence of alternative supply warning notice at or near origin', reg: 'Reg 514.15' },
    { ref: '1.14', desc: 'Adequacy of arrangements for isolation/switching of each circuit', reg: 'Reg 132.15, 537' },
    { ref: '1.15', desc: 'Correct connection of conductors (no single-pole devices in N)', reg: 'Reg 132.14.1, 530.3.2' },
    { ref: '1.16', desc: 'Adequacy of connections, including CPCs, within accessories', reg: 'Reg 526' },
    { ref: '1.17', desc: 'Adequacy of conductor sizing (current-carrying capacity)', reg: 'Reg 523, 433' },
    { ref: '1.18', desc: 'Presence of linked circuit-breaker or linked switch to each unmetered supply', reg: '' },
    { ref: '1.19', desc: 'Confirmation that ALL conductors are correctly connected', reg: 'Reg 526.3' },
    { ref: '1.20', desc: 'Confirmation that indicators and other devices are correctly connected', reg: '' },
    { ref: '1.21', desc: 'No basic insulation of a conductor visible outside enclosure', reg: 'Reg 526.8' },
    { ref: '1.22', desc: 'Suitability of surge protection device(s) (SPD) if fitted', reg: 'Reg 534' },
    { ref: '2.0', desc: 'PARALLEL OR SWITCHED ALTERNATIVE SOURCES OF SUPPLY', items: [] },
    { ref: '2.1', desc: 'Correct connection of alternative supply', reg: '' },
    { ref: '2.2', desc: 'Means of isolation of alternative supply', reg: '' },
    { ref: '2.3', desc: 'Adequate warning notices', reg: 'Reg 514.15' },
    { ref: '3.0', desc: 'DISTRIBUTION CIRCUITS', items: [] },
    { ref: '3.1', desc: 'Cables correctly supported throughout', reg: 'Table 4A, 4C' },
    { ref: '3.2', desc: 'Condition of insulation of live parts', reg: '' },
    { ref: '3.3', desc: 'Non-sheathed cables protected by enclosure in accordance with BS 7671', reg: 'Reg 521.10.1' },
    { ref: '3.4', desc: 'Cables concealed under floors, above ceilings and in walls adequately protected against damage', reg: 'Reg 522.6' },
    { ref: '3.5', desc: 'Provision of additional protection by RCD not exceeding 30 mA where required', reg: 'Reg 411.3.3, 411.3.4' },
    { ref: '3.6', desc: 'Adequacy of cables for current-carrying capacity with respect to the type and nature of installation', reg: 'Reg 523' },
    { ref: '3.7', desc: 'Cables adequately protected against mechanical damage and/or electromagnetic effects', reg: 'Reg 522.5, 522.6' },
    { ref: '3.8', desc: 'No basic insulation of a conductor visible outside enclosure', reg: 'Reg 526.8' },
    { ref: '3.9', desc: 'Connections soundly made and under no undue strain', reg: 'Reg 526.6' },
    { ref: '3.10', desc: 'No signs of overheating at connections', reg: '' },
    { ref: '4.0', desc: 'FINAL CIRCUITS', items: [] },
    { ref: '4.1', desc: 'Identification of conductors', reg: 'Table 51' },
    { ref: '4.2', desc: 'Cables correctly supported throughout', reg: 'Table 4A, 4C' },
    { ref: '4.3', desc: 'Condition of insulation of live parts', reg: '' },
    { ref: '4.4', desc: 'Non-sheathed cables protected by enclosure in accordance with BS 7671', reg: 'Reg 521.10.1' },
    { ref: '4.5', desc: 'Cables concealed under floors, above ceilings and in walls adequately protected against damage', reg: 'Reg 522.6' },
    { ref: '4.6', desc: 'Provision of additional protection by RCD not exceeding 30 mA', reg: 'Reg 411.3.3, 411.3.4' },
    { ref: '4.7', desc: 'Cables adequately protected against mechanical damage', reg: 'Reg 522.5, 522.6' },
    { ref: '4.8', desc: 'No basic insulation of a conductor visible outside enclosure', reg: 'Reg 526.8' },
    { ref: '4.9', desc: 'Connections soundly made and under no undue strain', reg: 'Reg 526.6' },
    { ref: '4.10', desc: 'No signs of overheating at connections', reg: '' },
    { ref: '4.11', desc: 'Adequacy of cables for current-carrying capacity', reg: 'Reg 523' },
    { ref: '5.0', desc: 'ISOLATION AND SWITCHING', items: [] },
    { ref: '5.1', desc: 'Presence and correct operation of appropriate devices for isolation and switching', reg: 'Reg 537' },
    { ref: '5.2', desc: 'Correct functioning of all isolators and switches', reg: 'Reg 537' },
    { ref: '5.3', desc: 'Isolator/switch accessible and clearly identified', reg: 'Reg 537.3' },
    { ref: '5.4', desc: 'Correct operation of all circuit-breakers', reg: 'Reg 612.13.2' },
    { ref: '5.5', desc: 'Condition of all enclosures and accessories', reg: '' },
    { ref: '6.0', desc: 'CURRENT-USING EQUIPMENT (Permanently Connected)', items: [] },
    { ref: '6.1', desc: 'Suitability of equipment in terms of IP rating and fire rating', reg: '' },
    { ref: '6.2', desc: 'Enclosure not damaged/deteriorated so as to impair safety', reg: '' },
    { ref: '6.3', desc: 'Suitability for the environment and external influences', reg: '' },
    { ref: '6.4', desc: 'Security of fixing', reg: '' },
    { ref: '6.5', desc: 'Cable entry holes adequately sealed', reg: '' },
    { ref: '7.0', desc: 'EARTHING AND BONDING', items: [] },
    { ref: '7.1', desc: 'Presence and adequacy of earthing conductor', reg: 'Table 54.7' },
    { ref: '7.2', desc: 'Presence and adequacy of circuit protective conductors', reg: 'Reg 543' },
    { ref: '7.3', desc: 'Presence and adequacy of main protective bonding conductors', reg: 'Reg 544.1' },
    { ref: '7.4', desc: 'Presence and adequacy of supplementary bonding conductors (where required)', reg: 'Reg 544.2' },
    { ref: '7.5', desc: 'Adequacy of earthing/bonding labels at all appropriate locations', reg: 'Reg 514.13' },
    { ref: '7.6', desc: 'Accessibility and condition of earthing and bonding connections', reg: '' },
    { ref: '7.7', desc: 'Accessibility and condition of earth electrode connection (where applicable)', reg: '' },
    { ref: '8.0', desc: 'GENERAL', items: [] },
    { ref: '8.1', desc: 'Adequacy of access to switchgear, equipment etc.', reg: 'Reg 132.12' },
    { ref: '8.2', desc: 'Presence of danger or warning notices and other required notices', reg: 'Reg 514' },
    { ref: '8.3', desc: 'Condition of accessories including socket-outlets, switches, fused connection units, etc.', reg: '' },
    { ref: '8.4', desc: 'Single-pole switches or devices in line conductors only', reg: 'Reg 530.3.2' },
    { ref: '8.5', desc: 'Protection against mechanical damage where cables pass through walls, floors and ceilings', reg: 'Reg 522.6' },
    { ref: '8.6', desc: 'Additional protection for cables concealed in walls at a depth less than 50 mm', reg: 'Reg 522.6.101, 522.6.102, 522.6.103' },
    { ref: '8.7', desc: 'Provision of fire barriers, sealing and protection against thermal effects', reg: 'Reg 527' },
    { ref: '8.8', desc: 'External condition of wiring system components', reg: '' },
    { ref: '8.9', desc: 'Security of fixing of wiring systems', reg: '' },
    { ref: '8.10', desc: 'Condition of enclosures', reg: '' },
    { ref: '9.0', desc: 'PROSUMER\'S LOW VOLTAGE INSTALLATION', items: [] },
    { ref: '9.1', desc: 'Confirmation of the type of PEI and its external influences', reg: '' },
    { ref: '9.2', desc: 'Confirmation of correct type and operation of interface protection', reg: '' },
    { ref: '9.3', desc: 'Presence of all appropriate notices and labels', reg: '' },
    { ref: '10.0', desc: 'ELECTRIC VEHICLE CHARGING INSTALLATION', items: [] },
    { ref: '10.1', desc: 'Correct type and rating of EVCP', reg: '' },
    { ref: '10.2', desc: 'Supply cable adequately sized and protected', reg: '' },
    { ref: '10.3', desc: 'Earthing arrangement adequate', reg: '' },
    { ref: '10.4', desc: 'Protective devices correctly rated and suitable', reg: '' },
    { ref: '11.0', desc: 'LOCATIONS CONTAINING A BATH OR SHOWER', items: [] },
    { ref: '11.1', desc: 'Suitability of equipment for the zone in which it is installed', reg: 'Section 701' },
    { ref: '11.2', desc: 'Suitability of equipment for external influences (IP rating)', reg: 'Reg 701.512.2' },
    { ref: '11.3', desc: 'Supplementary bonding conductors (where required)', reg: 'Reg 701.415.2' },
    { ref: '11.4', desc: 'Additional protection by 30 mA RCD for circuits serving the location', reg: 'Reg 701.411.3.3' },
  ];

  // Build inspection outcome lookup
  const inspMap = {};
  insp.forEach(i => { inspMap[i.item_ref] = i.outcome; });

  // Per-board observations
  const boardObsMap = {};
  obs.forEach(o => {
    if (o.board_id) {
      if (!boardObsMap[o.board_id]) boardObsMap[o.board_id] = [];
      boardObsMap[o.board_id].push(o);
    }
  });

  // Classify observations
  const c1Items = obs.filter(o => o.code === 'C1').map((o,i) => (i+1));
  const c2Items = obs.filter(o => o.code === 'C2').map((o,i) => (i+1));
  const c3Items = obs.filter(o => o.code === 'C3').map((o,i) => (i+1));
  const fiItems = obs.filter(o => o.code === 'FI').map((o,i) => (i+1));

  // Re-map item numbers from full obs list
  const c1Nos = []; const c2Nos = []; const c3Nos = []; const fiNos = [];
  obs.forEach((o, i) => {
    const num = i + 1;
    if (o.code === 'C1') c1Nos.push(num);
    if (o.code === 'C2') c2Nos.push(num);
    if (o.code === 'C3') c3Nos.push(num);
    if (o.code === 'FI') fiNos.push(num);
  });

  // Description checkbox helper
  const descType = (site.description || '').toLowerCase();
  const chk = (type) => descType.includes(type) ? '&#10003;' : '';

  let html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>EICR - ${e(site.name)}</title>
<style>
  @page { size: A4 portrait; margin: 10mm 10mm 18mm 10mm; }
  @page landscape { size: A4 landscape; margin: 10mm 10mm 18mm 10mm; }
  * { box-sizing: border-box; }
  body { font-family: Arial, Helvetica, sans-serif; font-size: 9px; line-height: 1.3; color: #000; margin: 0; padding: 0; }
  .page { width: 100%; position: relative; page-break-after: always; }
  .page:last-child { page-break-after: auto; }
  .page-landscape { page-break-before: always; }

  /* Header bar */
  .hdr { display: flex; align-items: center; justify-content: space-between; border-bottom: 2px solid #000; padding: 4px 0; margin-bottom: 6px; }
  .hdr img { max-height: 50px; max-width: 150px; }
  .hdr-right { text-align: right; font-size: 8px; line-height: 1.2; }
  .hdr-title { font-size: 11px; font-weight: bold; }

  /* Section headers */
  .sec-hdr { background: #000; color: #fff; font-weight: bold; font-size: 9px; padding: 3px 6px; margin: 0; border: 1px solid #000; }
  .sec-hdr-gray { background: #666; color: #fff; font-weight: bold; font-size: 8px; padding: 2px 6px; margin: 0; border: 1px solid #000; }

  /* Tables */
  table { border-collapse: collapse; width: 100%; margin: 0; }
  th, td { border: 1px solid #000; padding: 1.5px 3px; font-size: 8px; text-align: left; vertical-align: top; }
  th { background: #e0e0e0; font-weight: bold; font-size: 7.5px; text-align: center; }
  .fl { font-weight: bold; background: #f0f0f0; width: 25%; font-size: 8px; }
  .fl2 { font-weight: bold; background: #f0f0f0; font-size: 8px; }
  .fv { font-size: 8px; min-height: 14px; }
  .fv-large { font-size: 14px; font-weight: bold; text-align: center; padding: 6px; }
  .center { text-align: center; }
  .tick { text-align: center; font-size: 12px; }

  /* Assessment box */
  .assess-sat { background: #fff; font-size: 16px; font-weight: bold; text-align: center; padding: 8px; border: 3px solid #000; }
  .assess-unsat { background: #fff; font-size: 16px; font-weight: bold; text-align: center; padding: 8px; border: 3px solid #000; }

  /* Observation colours */
  .obs-c1 { background: #ffcccc; } .obs-c2 { background: #ffe0b2; } .obs-c3 { background: #cce5ff; } .obs-fi { background: #e1d5f0; }

  /* Footer */
  .page-footer { position: fixed; bottom: 0; left: 0; right: 0; text-align: center; font-size: 7px; color: #333; border-top: 1px solid #000; padding: 3px 10mm; background: #fff; }

  /* Signature */
  .sig-img { max-height: 30px; }

  /* Circuit table */
  .cct-table th { font-size: 6.5px; padding: 1px 2px; white-space: nowrap; }
  .cct-table td { font-size: 7px; padding: 1px 2px; text-align: center; }
  .cct-table td:nth-child(2) { text-align: left; }

  /* Print overrides */
  @media print {
    .no-print { display: none !important; }
    body { margin: 0; }
    .page-footer { position: fixed; bottom: 0; }
  }

  /* Wiring codes legend */
  .legend-table td { font-size: 7px; padding: 1px 3px; border: 1px solid #000; }
  .legend-table { margin-top: 4px; }

  /* Inspection schedule */
  .insp-table th { font-size: 7.5px; }
  .insp-table td { font-size: 7.5px; }
  .insp-section td { background: #e0e0e0; font-weight: bold; font-size: 8px; }
  .insp-pass { color: #000; }
  .insp-fail { color: #000; font-weight: bold; }
</style></head><body>

<div class="no-print" style="text-align:center;padding:10px;background:#333;color:white;margin-bottom:0">
  <button onclick="window.print()" style="padding:10px 30px;font-size:14px;font-weight:bold;cursor:pointer;border:none;background:#0066cc;color:#fff;border-radius:4px;margin-right:10px">Print / Save as PDF</button>
  <span style="font-size:12px">Optimised for A4 printing &mdash; Use Chrome Print &gt; Save as PDF</span>
</div>`;

  // ========== PAGE 1: MAIN REPORT ==========
  html += `<div class="page">`;
  // Header
  html += `<div class="hdr">
    <div><img src="/logo.jpeg" alt="Logo" onerror="this.style.display='none'"></div>
    <div style="text-align:center;flex:1">
      <div style="font-size:13px;font-weight:bold">ELECTRICAL INSTALLATION CONDITION REPORT</div>
      <div style="font-size:8px">(Requirements for Electrical Installations &mdash; BS 7671: 2018+A2:2022)</div>
    </div>
    <div class="hdr-right">
      <div>Ref: ${e(certNum)}</div>
    </div>
  </div>`;

  // Section 1
  html += `<table>
    <tr><td class="sec-hdr" colspan="4">SECTION 1: DETAILS OF THE PERSON ORDERING THE REPORT</td></tr>
    <tr><td class="fl" style="width:20%">Client:</td><td class="fv" colspan="3">${e(site.client_name)}</td></tr>
    <tr><td class="fl">Address:</td><td class="fv" colspan="3">${e(site.client_address || site.address)} ${e(site.postcode)}</td></tr>
  </table>`;

  // Section 2
  html += `<table>
    <tr><td class="sec-hdr" colspan="4">SECTION 2: REASON FOR PRODUCING THIS REPORT</td></tr>
    <tr><td class="fl" style="width:20%">Reason:</td><td class="fv" colspan="3">${e(site.purpose)}</td></tr>
    <tr><td class="fl">Date(s) of inspection:</td><td class="fv">${e(site.inspector_date || '')}</td><td class="fl" style="width:20%">Report Ref:</td><td class="fv">${e(certNum)}</td></tr>
  </table>`;

  // Section 3
  html += `<table>
    <tr><td class="sec-hdr" colspan="6">SECTION 3: DETAILS OF THE INSTALLATION</td></tr>
    <tr><td class="fl" style="width:20%">Installation Address:</td><td class="fv" colspan="5">${e(site.address)} ${e(site.postcode)}</td></tr>
    <tr>
      <td class="fl">Description of Premises:</td>
      <td class="fv" style="width:15%">Domestic <span class="tick">${chk('domestic')}</span></td>
      <td class="fv" style="width:15%">Commercial <span class="tick">${chk('commercial')}</span></td>
      <td class="fv" style="width:15%">Industrial <span class="tick">${chk('industrial')}</span></td>
      <td class="fv" style="width:15%">Other <span class="tick">${chk('other')}</span></td>
      <td class="fv">${e(site.description)}</td>
    </tr>
    <tr>
      <td class="fl">Estimated age of the<br>electrical installation (years):</td><td class="fv">${e(site.wiring_age)}${site.wiring_age ? '' : ''}</td>
      <td class="fl2">Evidence of additions<br>or alterations:</td><td class="fv">${e(site.additions)}${site.additions_age ? ' (' + e(site.additions_age) + ' years est.)' : ''}</td>
      <td class="fl2">Installation records<br>available:</td><td class="fv">${e(site.records_available)}</td>
    </tr>
    <tr><td class="fl">Date of last inspection:</td><td class="fv" colspan="5">${e(site.last_inspection_date)}</td></tr>
  </table>`;

  // Section 4
  html += `<table>
    <tr><td class="sec-hdr" colspan="2">SECTION 4: EXTENT AND LIMITATIONS OF THE INSPECTION AND TESTING</td></tr>
    <tr><td class="fl" style="width:20%">Extent of the installation covered<br>by this report:</td><td class="fv">${e(site.extent)}</td></tr>
    <tr><td class="fl">Agreed limitations including<br>reasons (see also Section 12):</td><td class="fv">${e(site.agreed_limitations)}</td></tr>
    <tr><td class="fl">Agreed with:</td><td class="fv">${e(site.agreed_with)}</td></tr>
    <tr><td class="fl">Operational limitations including<br>reasons (see also Section 12):</td><td class="fv">${e(site.operational_limitations)}</td></tr>
    <tr><td colspan="2" style="font-size:7px;padding:3px">The inspection and testing has been carried out in accordance with BS 7671: 2018+A2:2022 as amended. The inspection and testing detailed within this report, and described within the limitations given in Section 12, has been carried out in a manner intended to identify, so far as is reasonably practicable, any damage, deterioration, defects, dangerous conditions and non-compliances with the requirements of BS 7671 that may give rise to danger.</td></tr>
  </table>`;

  // Section 5
  html += `<table>
    <tr><td class="sec-hdr" colspan="4">SECTION 5: SUMMARY OF THE CONDITION OF THE INSTALLATION</td></tr>
    <tr>
      <td class="fl" style="width:30%">General condition of the installation<br>(in terms of electrical safety):</td>
      <td style="width:30%">
        <div class="assess-${hasC1C2FI ? 'unsat' : 'sat'}">${hasC1C2FI ? 'UNSATISFACTORY' : 'SATISFACTORY'}</div>
      </td>
      <td class="fl" style="width:20%">Date of next inspection<br>(recommended):</td>
      <td class="fv" style="font-size:11px;font-weight:bold">${e(site.next_inspection || site.next_inspection_date)}</td>
    </tr>
  </table>`;

  html += `</div>`;

  // ========== PAGE 2: OBSERVATIONS & RECOMMENDATIONS (Section 7) ==========
  html += `<div class="page">`;
  html += `<div class="hdr">
    <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
    <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">OBSERVATIONS AND RECOMMENDATIONS</div></div>
    <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
  </div>`;

  html += `<table>
    <tr><td class="sec-hdr" colspan="3">SECTION 7: OBSERVATIONS AND RECOMMENDATIONS FOR ACTIONS TO BE TAKEN</td></tr>
  </table>`;

  html += `<table>
    <thead><tr>
      <th style="width:40px">Item No</th>
      <th>Observations<br>(See regulation numbers, where given, for further guidance)</th>
      <th style="width:60px">Classification<br>Code</th>
    </tr></thead><tbody>`;
  if (obs.length) {
    obs.forEach((o, i) => {
      const cls = o.code === 'C1' ? 'obs-c1' : o.code === 'C2' ? 'obs-c2' : o.code === 'C3' ? 'obs-c3' : o.code === 'FI' ? 'obs-fi' : '';
      html += `<tr class="${cls}"><td class="center">${i + 1}</td><td>${e(o.description)}${o.location ? ' [' + e(o.location) + ']' : ''}</td><td class="center"><strong>${e(o.code)}</strong></td></tr>`;
    });
  } else {
    html += `<tr><td class="center">-</td><td style="color:#666;font-style:italic">No observations recorded</td><td class="center">-</td></tr>`;
  }
  html += `</tbody></table>`;

  // Classification legend
  html += `<table style="margin-top:6px">
    <tr><td class="sec-hdr-gray" colspan="2">CLASSIFICATION CODE</td></tr>
    <tr class="obs-c1"><td style="width:30px;text-align:center;font-weight:bold">C1</td><td>Danger present. Risk of injury. Immediate remedial action required.</td></tr>
    <tr class="obs-c2"><td style="text-align:center;font-weight:bold">C2</td><td>Potentially dangerous. Urgent remedial action required.</td></tr>
    <tr class="obs-c3"><td style="text-align:center;font-weight:bold">C3</td><td>Improvement recommended.</td></tr>
    <tr class="obs-fi"><td style="text-align:center;font-weight:bold">FI</td><td>Further investigation required without delay.</td></tr>
  </table>`;

  // Summary lines
  html += `<table style="margin-top:6px">
    <tr><td class="fl2" style="width:60%">Immediate remedial action required for items classified C1:</td><td class="fv">${c1Nos.length ? c1Nos.join(', ') : 'None'}</td></tr>
    <tr><td class="fl2">Urgent remedial action required for items classified C2:</td><td class="fv">${c2Nos.length ? c2Nos.join(', ') : 'None'}</td></tr>
    <tr><td class="fl2">Improvement recommended for items classified C3:</td><td class="fv">${c3Nos.length ? c3Nos.join(', ') : 'None'}</td></tr>
    <tr><td class="fl2">Further investigation required for items classified FI:</td><td class="fv">${fiNos.length ? fiNos.join(', ') : 'None'}</td></tr>
  </table>`;

  html += `</div>`;

  // ========== PER-DB OBSERVATIONS PAGES ==========
  for (const b of siteBoards) {
    const bObs = boardObsMap[b.id];
    if (!bObs || !bObs.length) continue;
    html += `<div class="page">`;
    html += `<div class="hdr">
      <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
      <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">OBSERVATIONS &mdash; ${e(b.ref)} (${e(b.location || '')})</div></div>
      <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
    </div>`;
    html += `<table>
      <thead><tr>
        <th style="width:40px">Item No</th>
        <th>Observations</th>
        <th style="width:60px">Code</th>
      </tr></thead><tbody>`;
    bObs.forEach((o, i) => {
      const cls = o.code === 'C1' ? 'obs-c1' : o.code === 'C2' ? 'obs-c2' : o.code === 'C3' ? 'obs-c3' : o.code === 'FI' ? 'obs-fi' : '';
      html += `<tr class="${cls}"><td class="center">${o.item_no || (i + 1)}</td><td>${e(o.description)}${o.location ? ' [' + e(o.location) + ']' : ''}</td><td class="center"><strong>${e(o.code)}</strong></td></tr>`;
    });
    html += `</tbody></table>`;
    html += `</div>`;
  }

  // ========== GENERAL CONDITION & DECLARATION (Sections 8-9) ==========
  html += `<div class="page">`;
  html += `<div class="hdr">
    <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
    <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">GENERAL CONDITION &amp; DECLARATION</div></div>
    <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
  </div>`;

  // Section 8
  html += `<table>
    <tr><td class="sec-hdr" colspan="2">SECTION 8: GENERAL CONDITION OF THE INSTALLATION</td></tr>
    <tr><td class="fl" style="width:30%">General condition of the installation<br>(in terms of electrical safety):</td><td class="fv">${e(site.general_condition)}</td></tr>
    <tr><td class="fl">Overall assessment:</td><td class="fv" style="font-size:12px;font-weight:bold">${hasC1C2FI ? 'UNSATISFACTORY' : 'SATISFACTORY'}</td></tr>
  </table>`;

  // Section 9
  html += `<table style="margin-top:6px">
    <tr><td class="sec-hdr" colspan="4">SECTION 9: DECLARATION</td></tr>
    <tr><td colspan="4" style="font-size:7.5px;padding:3px">I/We, being the person(s) responsible for the inspection and testing of the electrical installation (as indicated by my/our signatures below), particulars of which are described in this report, having exercised reasonable skill and care when carrying out the inspection and testing, hereby declare that the information in this report, including the observations and the attached schedules, provides an accurate assessment of the condition of the electrical installation taking into account the stated extent and limitations.</td></tr>
  </table>`;

  // Contractor details
  html += `<table>
    <tr><td class="fl" style="width:25%">Trading title:</td><td class="fv" style="width:25%">${e(comp.name || '')}</td><td class="fl" style="width:25%">Registration number:</td><td class="fv" style="width:25%">${e(comp.registration_number || '')}</td></tr>
    <tr><td class="fl">Address:</td><td class="fv" colspan="3">${e(comp.address || '')} ${e(comp.postcode || '')}</td></tr>
    <tr><td class="fl">Telephone:</td><td class="fv">${e(comp.tel || '')}</td><td class="fl">Registration body:</td><td class="fv">${e(comp.registration_body || '')}</td></tr>
  </table>`;

  // QS signature
  if (qs.name) {
    html += `<table style="margin-top:6px">
      <tr><td class="sec-hdr-gray" colspan="4">QUALIFIED SUPERVISOR</td></tr>
      <tr><td class="fl" style="width:25%">Name:</td><td class="fv" style="width:25%">${e(qs.name)}</td><td class="fl" style="width:25%">Position:</td><td class="fv" style="width:25%">${e(qs.position || 'Qualified Supervisor')}</td></tr>`;
    const qNum = qs.napit_number || qs.niceic_number || qs.ecs_number || qs.jib_number || qs.other_qual || '';
    html += `<tr><td class="fl">Qualification/Registration No:</td><td class="fv">${e(qNum)}</td><td class="fl">Date:</td><td class="fv">${e(qs.date_signed || '')}</td></tr>`;
    html += `<tr><td class="fl">Signature:</td><td colspan="3">${qs.signature_data ? '<img class="sig-img" src="' + qs.signature_data + '">' : ''}</td></tr>`;
    html += `</table>`;
  }

  // Inspector/Tester signatures
  testers.forEach(t => {
    html += `<table style="margin-top:4px">
      <tr><td class="sec-hdr-gray" colspan="4">INSPECTOR / TESTER</td></tr>
      <tr><td class="fl" style="width:25%">Name:</td><td class="fv" style="width:25%">${e(t.name)}</td><td class="fl" style="width:25%">Position:</td><td class="fv" style="width:25%">${e(t.position || 'Electrician')}</td></tr>`;
    const tNum = t.napit_number || t.niceic_number || t.ecs_number || t.jib_number || t.other_qual || '';
    html += `<tr><td class="fl">Qualification/Registration No:</td><td class="fv">${e(tNum)}</td><td class="fl">Date:</td><td class="fv">${e(t.date_signed || '')}</td></tr>`;
    html += `<tr><td class="fl">Signature:</td><td colspan="3">${t.signature_data ? '<img class="sig-img" src="' + t.signature_data + '">' : ''}</td></tr>`;
    html += `</table>`;
  });

  html += `</div>`;

  // ========== SUPPLY CHARACTERISTICS (Section 10) ==========
  html += `<div class="page">`;
  html += `<div class="hdr">
    <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
    <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">SUPPLY CHARACTERISTICS AND EARTHING ARRANGEMENTS</div></div>
    <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
  </div>`;

  html += `<table>
    <tr><td class="sec-hdr" colspan="6">SECTION 10: SUPPLY CHARACTERISTICS AND EARTHING ARRANGEMENTS</td></tr>
  </table>`;

  // Nature of supply
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="6">Nature of Supply Parameters</td></tr>
    <tr>
      <td class="fl2">AC <span class="tick">${(site.supply_type || 'AC').toUpperCase() === 'AC' ? '&#10003;' : ''}</span></td>
      <td class="fl2">DC <span class="tick">${(site.supply_type || '').toUpperCase() === 'DC' ? '&#10003;' : ''}</span></td>
      <td class="fl2">Single phase <span class="tick">${(site.num_phases || '').includes('1') || (site.num_phases || '').toLowerCase().includes('single') ? '&#10003;' : ''}</span></td>
      <td class="fl2">Two phase <span class="tick">${(site.num_phases || '').includes('2') ? '&#10003;' : ''}</span></td>
      <td class="fl2">Three phase <span class="tick">${(site.num_phases || '').includes('3') ? '&#10003;' : ''}</span></td>
      <td class="fv"></td>
    </tr>
  </table>`;

  // Earthing arrangements
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="6">Earthing Arrangements</td></tr>
    <tr>
      <td class="fl2">TN-S <span class="tick">${earthTick('TNS', site.earthing_type)}</span></td>
      <td class="fl2">TN-C-S <span class="tick">${earthTick('TNCS', site.earthing_type)}</span></td>
      <td class="fl2">TT <span class="tick">${earthTick('TT', site.earthing_type)}</span></td>
      <td class="fl2">TN-C <span class="tick">${earthTick('TNC', site.earthing_type)}</span></td>
      <td class="fl2">IT <span class="tick">${earthTick('IT', site.earthing_type)}</span></td>
      <td class="fv">${e(site.earthing_type)}</td>
    </tr>
  </table>`;

  // Supply details
  html += `<table>
    <tr><td class="fl" style="width:25%">Nominal voltage, U / U<sub>0</sub> (V):</td><td class="fv" style="width:25%">${e(site.nominal_voltage)}</td><td class="fl" style="width:25%">Nominal frequency, f (Hz):</td><td class="fv" style="width:25%">${e(site.nominal_frequency)}</td></tr>
    <tr><td class="fl">Prospective fault current,<br>Ipf (kA):</td><td class="fv">${e(site.ipf_at_origin)}</td><td class="fl">External loop impedance,<br>Ze (&Omega;):</td><td class="fv">${e(site.ze_at_origin)}</td></tr>
  </table>`;

  // Supply protective device
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="6">Supply Protective Device</td></tr>
    <tr><td class="fl" style="width:20%">BS (EN):</td><td class="fv" style="width:15%">${e(site.supply_device_bsen)}</td><td class="fl" style="width:15%">Type:</td><td class="fv" style="width:15%">${e(site.supply_device_type)}</td><td class="fl" style="width:15%">Rating (A):</td><td class="fv" style="width:20%">${e(site.supply_device_rating)}</td></tr>
    <tr><td class="fl">Number of supplies:</td><td class="fv" colspan="5">${e(site.num_supplies || '1')}</td></tr>
  </table>`;

  html += `</div>`;

  // ========== PARTICULARS OF INSTALLATION (Section 11) ==========
  html += `<div class="page">`;
  html += `<div class="hdr">
    <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
    <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">PARTICULARS OF INSTALLATION REFERRED TO IN THE REPORT</div></div>
    <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
  </div>`;

  html += `<table>
    <tr><td class="sec-hdr" colspan="4">SECTION 11: PARTICULARS OF INSTALLATION AT THE ORIGIN</td></tr>
  </table>`;

  // Means of earthing
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="4">Means of Earthing</td></tr>
    <tr>
      <td class="fl" style="width:30%">Distributor's facility:</td>
      <td class="fv" style="width:20%"><span class="tick">${(site.means_of_earthing || '').toLowerCase().includes('distributor') ? '&#10003;' : ''}</span></td>
      <td class="fl" style="width:30%">Installation earth electrode:</td>
      <td class="fv" style="width:20%"><span class="tick">${(site.means_of_earthing || '').toLowerCase().includes('electrode') ? '&#10003;' : ''}</span></td>
    </tr>
  </table>`;

  // Earth electrode details
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="6">Earth Electrode (where applicable)</td></tr>
    <tr>
      <td class="fl" style="width:20%">Type:</td><td class="fv" style="width:15%">${e(site.earth_electrode_type)}</td>
      <td class="fl" style="width:15%">Location:</td><td class="fv" style="width:15%">${e(site.earth_electrode_location)}</td>
      <td class="fl" style="width:15%">Resistance (&Omega;):</td><td class="fv" style="width:20%">${e(site.earth_electrode_resistance)}</td>
    </tr>
  </table>`;

  // Main switch details
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="6">Main Switch / Switch-fuse / RCD / RCBO</td></tr>
    <tr>
      <td class="fl" style="width:20%">Location:</td><td class="fv" style="width:15%">${e(site.main_switch_location)}</td>
      <td class="fl" style="width:15%">BS (EN):</td><td class="fv" style="width:15%">${e(site.main_switch_bsen)}</td>
      <td class="fl" style="width:15%">No. of poles:</td><td class="fv" style="width:20%">${e(site.main_switch_poles)}</td>
    </tr>
    <tr>
      <td class="fl">Current rating (A):</td><td class="fv">${e(site.main_switch_current_rating)}</td>
      <td class="fl">Fuse/device rating (A):</td><td class="fv">${e(site.main_switch_fuse_rating)}</td>
      <td class="fl">Voltage rating (V):</td><td class="fv">${e(site.main_switch_voltage_rating)}</td>
    </tr>
    <tr>
      <td class="fl">RCD type (if applicable):</td><td class="fv">${e(site.main_switch_rcd_type)}</td>
      <td class="fl">RCD I&Delta;n (mA):</td><td class="fv">${e(site.main_switch_rcd_idn)}</td>
      <td class="fl">RCD operating time (ms):</td><td class="fv">${e(site.main_switch_rcd_time)}</td>
    </tr>
  </table>`;

  // Earthing and protective bonding conductors
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="6">Earthing and Protective Bonding Conductors</td></tr>
    <tr>
      <th></th><th>Material</th><th>CSA (mm&sup2;)</th><th colspan="3">Connection verified &#10003;</th>
    </tr>
    <tr>
      <td class="fl2">Earthing conductor:</td>
      <td class="fv center">${e(site.earthing_conductor_material)}</td>
      <td class="fv center">${e(site.earthing_conductor_csa)}</td>
      <td class="fv center tick" colspan="3">${tick(site.earthing_conductor_verified)}</td>
    </tr>
    <tr>
      <td class="fl2">Main protective bonding<br>conductor:</td>
      <td class="fv center">${e(site.bonding_conductor_material)}</td>
      <td class="fv center">${e(site.bonding_conductor_csa)}</td>
      <td class="fv center tick" colspan="3">${tick(site.bonding_conductor_verified)}</td>
    </tr>
  </table>`;

  // Bonding table
  html += `<table>
    <tr><td class="sec-hdr-gray" colspan="7">Main Protective Bonding Connections to:</td></tr>
    <tr>
      <th>Water</th><th>Gas</th><th>Oil</th><th>Lightning</th><th>Structural Steel</th><th>Other</th><th>Details</th>
    </tr>
    <tr>
      <td class="fv center tick">${tick(site.bonding_water) || e(site.bonding_water)}</td>
      <td class="fv center tick">${tick(site.bonding_gas) || e(site.bonding_gas)}</td>
      <td class="fv center tick">${tick(site.bonding_oil) || e(site.bonding_oil)}</td>
      <td class="fv center tick">${tick(site.bonding_lightning) || e(site.bonding_lightning)}</td>
      <td class="fv center tick">${tick(site.bonding_steel) || e(site.bonding_steel)}</td>
      <td class="fv center tick">${tick(site.bonding_other) || e(site.bonding_other)}</td>
      <td class="fv"></td>
    </tr>
  </table>`;

  html += `</div>`;

  // ========== INSPECTION SCHEDULE (Section 12) ==========
  // Split schedule into chunks to fit pages
  const schedChunks = [];
  let currentChunk = [];
  inspectionSchedule.forEach((item, idx) => {
    currentChunk.push(item);
    if (currentChunk.length >= 40 || idx === inspectionSchedule.length - 1) {
      schedChunks.push(currentChunk);
      currentChunk = [];
    }
  });

  schedChunks.forEach((chunk, ci) => {
    html += `<div class="page">`;
    html += `<div class="hdr">
      <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
      <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">SCHEDULE OF INSPECTIONS</div></div>
      <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
    </div>`;

    if (ci === 0) {
      html += `<table><tr><td class="sec-hdr" colspan="10">SECTION 12: SCHEDULE OF INSPECTIONS (in accordance with BS 7671)</td></tr></table>`;
      html += `<table style="margin-bottom:4px"><tr>
        <td style="font-size:7px;border:1px solid #000;padding:2px 4px"><strong>Outcomes:</strong> Pass (P) = acceptable condition &bull; C1 = Danger present &bull; C2 = Potentially dangerous &bull; C3 = Improvement recommended &bull; FI = Further investigation &bull; N/A = Not applicable &bull; LIM = Not inspected (limitation) &bull; N/V = Not verified</td>
      </tr></table>`;
    }

    html += `<table class="insp-table">
      <thead><tr>
        <th style="width:35px">Ref</th>
        <th>Inspection Item</th>
        <th style="width:40px">Outcome</th>
      </tr></thead><tbody>`;

    chunk.forEach(item => {
      if (item.items) {
        // Section header
        html += `<tr class="insp-section"><td>${e(item.ref)}</td><td colspan="2">${e(item.desc)}</td></tr>`;
      } else {
        const outcome = inspMap[item.ref] || 'N/V';
        const oClass = outcome === 'Pass' ? 'insp-pass' : ['C1','C2','FI'].includes(outcome) ? 'insp-fail' : '';
        html += `<tr>
          <td class="center">${e(item.ref)}</td>
          <td>${e(item.desc)}${item.reg ? ' <span style="color:#666;font-size:6.5px">(' + e(item.reg) + ')</span>' : ''}</td>
          <td class="center ${oClass}"><strong>${e(outcome)}</strong></td>
        </tr>`;
      }
    });

    html += `</tbody></table>`;
    html += `</div>`;
  });

  // ========== PER-BOARD TEST RESULTS (Form 4) ==========
  for (const b of siteBoards) {
    const ccts = boardCircuits[b.id] || [];

    html += `<div class="page">`;
    html += `<div class="hdr">
      <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
      <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">SCHEDULE OF CIRCUIT DETAILS AND TEST RESULTS</div><div style="font-size:8px">(FORM 4)</div></div>
      <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
    </div>`;

    // DB Header box
    html += `<table>
      <tr><td class="sec-hdr" colspan="6">DISTRIBUTION BOARD: ${e(b.ref)}</td></tr>
      <tr>
        <td class="fl" style="width:18%">DB reference:</td><td class="fv" style="width:15%">${e(b.ref)}</td>
        <td class="fl" style="width:18%">Location:</td><td class="fv" style="width:15%">${e(b.location)}</td>
        <td class="fl" style="width:18%">Supplied from:</td><td class="fv" style="width:16%">${e(b.supplied_from)}</td>
      </tr>
    </table>`;

    html += `<table>
      <tr><td class="sec-hdr-gray" colspan="8">Distribution circuit OCPD &amp; DB details</td></tr>
      <tr>
        <td class="fl2" style="width:12%">BS (EN):</td><td class="fv" style="width:13%">${e(b.dist_bsen)}</td>
        <td class="fl2" style="width:12%">Type:</td><td class="fv" style="width:13%">${e(b.dist_type)}</td>
        <td class="fl2" style="width:12%">Rating (A):</td><td class="fv" style="width:13%">${e(b.dist_rating)}</td>
        <td class="fl2" style="width:12%">No. of phases:</td><td class="fv" style="width:13%">${e(b.num_phases || '1')}</td>
      </tr>
      <tr>
        <td class="fl2">SPD type(s):</td><td class="fv">${e(b.spd_types)}</td>
        <td class="fl2">SPD status:</td><td class="fv">${e(b.spd_status)}</td>
        <td class="fl2">Supply polarity confirmed:</td><td class="fv tick">${tick(b.supply_polarity)}</td>
        <td class="fl2">Phase sequence:</td><td class="fv">${e(b.phase_sequence)}</td>
      </tr>
      <tr>
        <td class="fl2">Zs at DB (&Omega;):</td><td class="fv">${e(b.ze)}</td>
        <td class="fl2">Ipf at DB (kA):</td><td class="fv" colspan="5">${e(b.ipf)}</td>
      </tr>
    </table>`;

    // Circuit table
    if (ccts.length) {
      html += `<table class="cct-table" style="margin-top:4px">
        <thead>
          <tr>
            <th colspan="8" style="background:#000;color:#fff;font-size:7px">CIRCUIT DETAILS</th>
            <th colspan="5" style="background:#000;color:#fff;font-size:7px">OVERCURRENT PROTECTIVE DEVICE</th>
            <th colspan="4" style="background:#000;color:#fff;font-size:7px">RCD</th>
            <th colspan="11" style="background:#000;color:#fff;font-size:7px">TEST RESULT DETAILS</th>
          </tr>
          <tr>
            <th rowspan="2" style="width:22px">Cct<br>No</th>
            <th rowspan="2" style="min-width:60px">Circuit<br>description</th>
            <th rowspan="2" style="width:22px">Type<br>wiring</th>
            <th rowspan="2" style="width:20px">Ref<br>method</th>
            <th rowspan="2" style="width:18px">No.<br>pts</th>
            <th rowspan="2" style="width:20px">Live<br>mm&sup2;</th>
            <th rowspan="2" style="width:20px">cpc<br>mm&sup2;</th>
            <th rowspan="2" style="width:22px">Max<br>disc<br>time</th>
            <th rowspan="2" style="width:24px">BS<br>(EN)</th>
            <th rowspan="2" style="width:18px">Type</th>
            <th rowspan="2" style="width:22px">Rating<br>(A)</th>
            <th rowspan="2" style="width:22px">Brk<br>cap<br>(kA)</th>
            <th rowspan="2" style="width:25px">Max<br>Zs<br>(&Omega;)</th>
            <th rowspan="2" style="width:24px">BS<br>(EN)</th>
            <th rowspan="2" style="width:18px">Type</th>
            <th rowspan="2" style="width:22px">I&Delta;n<br>(mA)</th>
            <th rowspan="2" style="width:22px">Rating<br>(A)</th>
            <th colspan="3" style="background:#e0e0e0;font-size:6px">Ring final cct</th>
            <th colspan="2" style="background:#e0e0e0;font-size:6px">Continuity</th>
            <th colspan="2" style="background:#e0e0e0;font-size:6px">Insulation resistance</th>
            <th rowspan="2" style="width:18px">Pol<br>&#10003;</th>
            <th rowspan="2" style="width:25px">Zs<br>(&Omega;)</th>
            <th colspan="2" style="background:#e0e0e0;font-size:6px">RCD</th>
            <th rowspan="2" style="width:18px">AFDD<br>&#10003;</th>
          </tr>
          <tr>
            <th style="width:20px;font-size:6px">r1<br>(&Omega;)</th>
            <th style="width:20px;font-size:6px">rn<br>(&Omega;)</th>
            <th style="width:20px;font-size:6px">r2<br>(&Omega;)</th>
            <th style="width:25px;font-size:6px">R1+R2<br>(&Omega;)</th>
            <th style="width:20px;font-size:6px">R2<br>(&Omega;)</th>
            <th style="width:25px;font-size:6px">L-L<br>(M&Omega;)</th>
            <th style="width:25px;font-size:6px">L-E<br>(M&Omega;)</th>
            <th style="width:22px;font-size:6px">x1<br>(ms)</th>
            <th style="width:18px;font-size:6px">Btn<br>&#10003;</th>
          </tr>
        </thead><tbody>`;

      ccts.forEach(c => {
        const zsNum = parseFloat(c.zs_measured);
        const maxNum = parseFloat(c.max_zs);
        const zsFail = zsNum && maxNum && zsNum > maxNum * 0.8;
        const polTick = (c.polarity === 'true' || c.polarity === true) ? '&#10003;' : e(c.polarity);
        const rcdBtnTick = (c.rcd_test_button === 'true' || c.rcd_test_button === true) ? '&#10003;' : e(c.rcd_test_button);
        const afddTick = (c.afdd_test === 'true' || c.afdd_test === true) ? '&#10003;' : e(c.afdd_test);

        html += `<tr>
          <td>${e(c.number)}</td>
          <td style="text-align:left">${e(c.description)}</td>
          <td>${e(c.wiring_type)}</td>
          <td>${e(c.ref_method)}</td>
          <td>${e(c.num_points)}</td>
          <td>${e(c.live_mm)}</td>
          <td>${e(c.cpc_mm)}</td>
          <td>${e(c.max_disconnect)}</td>
          <td>${e(c.ocpd_bsen)}</td>
          <td>${e(c.ocpd_type)}</td>
          <td>${e(c.ocpd_rating)}</td>
          <td>${e(c.ocpd_breaking_cap)}</td>
          <td>${e(c.max_zs)}</td>
          <td>${e(c.rcd_bsen)}</td>
          <td>${e(c.rcd_type)}</td>
          <td>${e(c.rcd_idn)}</td>
          <td>${e(c.rcd_rating_a)}</td>
          <td>${e(c.r1)}</td>
          <td>${e(c.rn)}</td>
          <td>${e(c.r2)}</td>
          <td>${e(c.r1r2)}</td>
          <td>${e(c.r2_ring)}</td>
          <td>${e(c.ir_ll)}</td>
          <td>${e(c.ir_le)}</td>
          <td>${polTick}</td>
          <td style="${zsFail ? 'font-weight:bold' : ''}">${e(c.zs_measured)}</td>
          <td>${e(c.rcd_time_x1)}</td>
          <td>${rcdBtnTick}</td>
          <td>${afddTick}</td>
        </tr>`;
      });
      html += `</tbody></table>`;
    } else {
      html += `<p style="color:#666;font-style:italic;font-size:8px;margin:8px 0">No circuits recorded for this distribution board.</p>`;
    }

    // Wiring codes legend
    html += `<table class="legend-table" style="margin-top:6px">
      <tr><td class="sec-hdr-gray" colspan="10">WIRING TYPE CODES</td></tr>
      <tr>
        <td><strong>A</strong></td><td>Thermoplastic insulated and sheathed (flat/round)</td>
        <td><strong>B</strong></td><td>Thermoplastic singles in metallic conduit</td>
        <td><strong>C</strong></td><td>Thermoplastic singles in non-metallic conduit</td>
        <td><strong>D</strong></td><td>Thermoplastic singles in metallic trunking</td>
        <td><strong>E</strong></td><td>Thermoplastic singles in non-metallic trunking</td>
      </tr>
      <tr>
        <td><strong>F</strong></td><td>Thermoplastic SWA cable</td>
        <td><strong>G</strong></td><td>Thermosetting SWA cable</td>
        <td><strong>H</strong></td><td>Mineral insulated</td>
        <td><strong>O</strong></td><td>Other (state in remarks)</td>
        <td colspan="2"></td>
      </tr>
    </table>`;

    // Test instruments block
    html += `<table style="margin-top:4px">
      <tr><td class="sec-hdr-gray" colspan="4">TEST INSTRUMENTS USED (serial numbers)</td></tr>
      <tr><td class="fl2" style="width:25%">Multi-functional:</td><td class="fv" style="width:25%">${e(b.instruments_multi)}</td><td class="fl2" style="width:25%">Insulation resistance:</td><td class="fv" style="width:25%">${e(b.instruments_ir)}</td></tr>
      <tr><td class="fl2">Continuity:</td><td class="fv">${e(b.instruments_cont)}</td><td class="fl2">Earth electrode resistance:</td><td class="fv">${e(b.instruments_earth)}</td></tr>
      <tr><td class="fl2">Earth fault loop impedance:</td><td class="fv">${e(b.instruments_loop)}</td><td class="fl2">RCD:</td><td class="fv">${e(b.instruments_rcd)}</td></tr>
    </table>`;

    // Tested by
    const boardTester = testers[0] || qs;
    if (boardTester.name) {
      html += `<table style="margin-top:4px">
        <tr><td class="sec-hdr-gray" colspan="4">TESTED BY</td></tr>
        <tr><td class="fl2" style="width:15%">Name:</td><td class="fv" style="width:35%">${e(boardTester.name)}</td><td class="fl2" style="width:15%">Position:</td><td class="fv" style="width:35%">${e(boardTester.position || '')}</td></tr>
        <tr><td class="fl2">Signature:</td><td>${boardTester.signature_data ? '<img class="sig-img" src="' + boardTester.signature_data + '">' : ''}</td><td class="fl2">Date:</td><td class="fv">${e(boardTester.date_signed || '')}</td></tr>
      </table>`;
    }

    html += `</div>`;
  }

  // ========== GUIDANCE FOR RECIPIENTS ==========
  html += `<div class="page">`;
  html += `<div class="hdr">
    <div><img src="/logo.jpeg" alt="" onerror="this.style.display='none'"></div>
    <div style="text-align:center;flex:1"><div style="font-size:11px;font-weight:bold">GUIDANCE FOR RECIPIENTS</div></div>
    <div class="hdr-right"><div>Ref: ${e(certNum)}</div></div>
  </div>`;

  html += `<table>
    <tr><td class="sec-hdr">GUIDANCE FOR RECIPIENTS OF AN ELECTRICAL INSTALLATION CONDITION REPORT</td></tr>
    <tr><td style="font-size:8px;line-height:1.5;padding:6px">
      <p style="margin:0 0 6px"><strong>This report is an important and valuable document which should be retained for future reference.</strong></p>

      <p style="margin:0 0 4px"><strong>1. Purpose of the report</strong></p>
      <p style="margin:0 0 6px">The purpose of an Electrical Installation Condition Report is to confirm, so far as reasonably practicable, whether or not the electrical installation is in a satisfactory condition for continued service. The report should identify any damage, deterioration, defects and/or conditions which may give rise to danger together with observations for which improvement is recommended.</p>

      <p style="margin:0 0 4px"><strong>2. The report</strong></p>
      <p style="margin:0 0 6px">The report is based on the condition of the electrical installation at the time of the inspection and, where appropriate, the results of tests on the installation. It is not a guarantee that the installation will remain in a satisfactory condition. The condition of any electrical installation will change over time due to age, wear, damage, alteration and additions.</p>

      <p style="margin:0 0 4px"><strong>3. The person ordering the report</strong></p>
      <p style="margin:0 0 6px">The person ordering the report should have received the original report and any attachments. If you are a person other than the person who ordered the report, you should contact the person ordering the report for the original. You should also ensure that the report is complete and relates to the electrical installation concerned.</p>

      <p style="margin:0 0 4px"><strong>4. The electrical installation</strong></p>
      <p style="margin:0 0 6px">The report relates to the fixed electrical installation only. It does not cover portable or transportable electrical equipment connected to, or intended to be connected to, the installation by means of a plug and socket-outlet, or similar means, unless specifically agreed.</p>

      <p style="margin:0 0 4px"><strong>5. Classification codes</strong></p>
      <p style="margin:0 0 3px">Where the inspection and testing has revealed that the condition of an item is such that a classification code has been allocated, the following remedial action is required:</p>
      <p style="margin:0 0 2px"><strong>C1 (Danger present):</strong> Risk of injury. Immediate remedial action required.</p>
      <p style="margin:0 0 2px"><strong>C2 (Potentially dangerous):</strong> Urgent remedial action required.</p>
      <p style="margin:0 0 2px"><strong>C3 (Improvement recommended):</strong> Improvement is recommended.</p>
      <p style="margin:0 0 6px"><strong>FI (Further investigation required):</strong> Further investigation is required, without delay, to determine the extent and nature of the identified deficiency.</p>

      <p style="margin:0 0 4px"><strong>6. Overall assessment</strong></p>
      <p style="margin:0 0 6px">The overall assessment of SATISFACTORY indicates that the person or persons who carried out the inspection and testing, at the time of the inspection, were of the opinion that the condition of the electrical installation was satisfactory. An overall assessment of UNSATISFACTORY indicates that one or more observations classified C1, C2 or FI have been recorded and that the installation requires remedial action.</p>

      <p style="margin:0 0 4px"><strong>7. Recommended date of next inspection</strong></p>
      <p style="margin:0 0 6px">The recommended interval between inspections of an electrical installation is determined in accordance with the requirements of BS 7671. The maximum period recommended between inspections is typically: Domestic 10 years, Commercial 5 years, Industrial 3 years. However, these periods can vary depending on the use, external influences and condition of the installation.</p>

      <p style="margin:0 0 4px"><strong>8. Remedial action</strong></p>
      <p style="margin:0">Where the report indicates that remedial work is necessary, the remedial work should be carried out without delay. Remedial work should be carried out only by an electrician competent in such work. A minor electrical installation works certificate or an electrical installation certificate should be issued to confirm that the remedial work has been completed and the installation retested as necessary.</p>
    </td></tr>
  </table>`;

  html += `</div>`;

  // Footer script for page numbers
  html += `
  <div class="page-footer">
    This form is based on the model shown in Appendix 6 of BS 7671:2018+A2:2022. Ref: ${e(certNum)}
  </div>`;

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
