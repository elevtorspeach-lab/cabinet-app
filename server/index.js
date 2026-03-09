const express = require('express');
const fs = require('fs');
const fsp = require('fs/promises');
const path = require('path');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || '0.0.0.0';
const WEB_DIR = path.join(__dirname, '..');

const DATA_DIR = path.join(__dirname, 'data');
const STATE_FILE = path.join(DATA_DIR, 'state.json');
const BACKUP_DIR = path.join(DATA_DIR, 'backups');
const SERVER_BACKUP_RETENTION_COUNT = 20;
const SERVER_BACKUP_MIN_INTERVAL_MS = 3 * 60 * 1000;

const DEFAULT_STATE = {
  clients: [],
  salleAssignments: [],
  users: [],
  audienceDraft: {},
  recycleBin: [],
  recycleArchive: [],
  updatedAt: new Date().toISOString()
};

let cachedState = null;
let lastBackupSignature = '';
let lastBackupAt = 0;
const sseClients = new Set();

app.use(express.json({ limit: '120mb' }));
app.use(express.static(WEB_DIR, {
  index: false
}));

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(204);
  next();
});

async function ensureDataFile() {
  await fsp.mkdir(DATA_DIR, { recursive: true });
  await fsp.mkdir(BACKUP_DIR, { recursive: true });
  try {
    await fsp.access(STATE_FILE, fs.constants.F_OK);
  } catch {
    await writeState(DEFAULT_STATE);
  }
}

async function readState() {
  if (cachedState) return cachedState;
  try {
    const raw = await fsp.readFile(STATE_FILE, 'utf8');
    const parsed = JSON.parse(raw);
    cachedState = {
      ...DEFAULT_STATE,
      ...(parsed && typeof parsed === 'object' ? parsed : {}),
      updatedAt: String(parsed?.updatedAt || new Date().toISOString())
    };
    return cachedState;
  } catch {
    cachedState = { ...DEFAULT_STATE };
    return cachedState;
  }
}

function buildBackupSignature(state) {
  try {
    return JSON.stringify({
      clients: Array.isArray(state?.clients) ? state.clients : [],
      salleAssignments: Array.isArray(state?.salleAssignments) ? state.salleAssignments : [],
      users: Array.isArray(state?.users) ? state.users : [],
      audienceDraft: state?.audienceDraft && typeof state.audienceDraft === 'object' ? state.audienceDraft : {},
      recycleBin: Array.isArray(state?.recycleBin) ? state.recycleBin : [],
      recycleArchive: Array.isArray(state?.recycleArchive) ? state.recycleArchive : []
    });
  } catch {
    return '';
  }
}

function buildBackupFileName(ts = new Date()) {
  const pad = (value) => String(value).padStart(2, '0');
  return `state_${ts.getFullYear()}-${pad(ts.getMonth() + 1)}-${pad(ts.getDate())}_${pad(ts.getHours())}-${pad(ts.getMinutes())}-${pad(ts.getSeconds())}.json`;
}

async function pruneBackupFiles() {
  try {
    const entries = await fsp.readdir(BACKUP_DIR, { withFileTypes: true });
    const files = entries
      .filter((entry) => entry.isFile() && entry.name.endsWith('.json'))
      .map((entry) => entry.name)
      .sort()
      .reverse();
    const extras = files.slice(SERVER_BACKUP_RETENTION_COUNT);
    await Promise.all(extras.map((name) => fsp.unlink(path.join(BACKUP_DIR, name)).catch(() => {})));
  } catch (err) {
    console.warn('Failed to prune state backups:', err);
  }
}

async function maybeWriteBackupSnapshot(state) {
  const now = Date.now();
  const signature = buildBackupSignature(state);
  if (signature && signature === lastBackupSignature) return;
  if (lastBackupAt && (now - lastBackupAt) < SERVER_BACKUP_MIN_INTERVAL_MS) return;

  const snapshot = {
    savedAt: new Date(now).toISOString(),
    ...state
  };
  const backupPath = path.join(BACKUP_DIR, buildBackupFileName(new Date(now)));
  await fsp.writeFile(backupPath, JSON.stringify(snapshot, null, 2), 'utf8');
  lastBackupAt = now;
  lastBackupSignature = signature;
  await pruneBackupFiles();
}

async function writeState(nextState) {
  const safe = {
    ...DEFAULT_STATE,
    ...(nextState && typeof nextState === 'object' ? nextState : {}),
    updatedAt: new Date().toISOString()
  };
  const tmpFile = `${STATE_FILE}.tmp`;
  await fsp.writeFile(tmpFile, JSON.stringify(safe), 'utf8');
  await fsp.rename(tmpFile, STATE_FILE);
  cachedState = safe;
  await maybeWriteBackupSnapshot(safe);
  return safe;
}

function broadcastStateUpdated(payload) {
  const data = `event: state-updated\ndata: ${JSON.stringify({
    updatedAt: payload?.updatedAt || new Date().toISOString()
  })}\n\n`;
  sseClients.forEach((res) => {
    try {
      res.write(data);
    } catch {
      sseClients.delete(res);
    }
  });
}

app.get('/api/health', async (req, res) => {
  await ensureDataFile();
  res.json({ ok: true, service: 'cabinet-api', ts: new Date().toISOString() });
});

app.get('/api/state', async (req, res) => {
  await ensureDataFile();
  const state = await readState();
  res.json(state);
});

app.post('/api/state', async (req, res) => {
  await ensureDataFile();
  const body = req.body && typeof req.body === 'object' ? req.body : {};
  const saved = await writeState(body);
  broadcastStateUpdated(saved);
  res.json({ ok: true, updatedAt: saved.updatedAt });
});

app.get('/api/state/stream', async (req, res) => {
  await ensureDataFile();
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders?.();

  sseClients.add(res);
  const current = await readState();
  res.write(`event: state-updated\ndata: ${JSON.stringify({ updatedAt: current.updatedAt })}\n\n`);

  const keepAlive = setInterval(() => {
    try {
      res.write(': keep-alive\n\n');
    } catch {}
  }, 20000);

  req.on('close', () => {
    clearInterval(keepAlive);
    sseClients.delete(res);
    res.end();
  });
});

app.get('/', (req, res) => {
  res.sendFile(path.join(WEB_DIR, 'index.html'));
});

ensureDataFile()
  .then(() => {
    app.listen(PORT, HOST, () => {
      console.log(`Cabinet API running on http://${HOST}:${PORT}`);
    });
  })
  .catch((err) => {
    console.error('Failed to start API:', err);
    process.exit(1);
  });
