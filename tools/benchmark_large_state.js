const fs = require('fs/promises');
const path = require('path');
const os = require('os');
const { performance } = require('perf_hooks');

const dossierArg = process.argv.find((arg) => arg.startsWith('--dossiers='));
const audienceArg = process.argv.find((arg) => arg.startsWith('--audiences='));
const clientArg = process.argv.find((arg) => arg.startsWith('--clients='));
const diligenceArg = process.argv.find((arg) => arg.startsWith('--diligences='));
const DOSSIER_COUNT = Math.max(1, Number(process.env.BENCH_DOSSIERS || (dossierArg ? dossierArg.slice('--dossiers='.length) : 60000)) || 60000);
const AUDIENCE_COUNT = Math.max(0, Number(process.env.BENCH_AUDIENCES || (audienceArg ? audienceArg.slice('--audiences='.length) : 80000)) || 80000);
const CLIENT_COUNT = Math.max(1, Number(process.env.BENCH_CLIENTS || (clientArg ? clientArg.slice('--clients='.length) : 300)) || 300);
const DILIGENCE_COUNT = Math.max(0, Number(process.env.BENCH_DILIGENCES || (diligenceArg ? diligenceArg.slice('--diligences='.length) : 0)) || 0);
const OUTPUT_DIR = path.join(os.tmpdir(), 'applicationversion1-benchmark');
const postUrlArg = process.argv.find((arg) => arg.startsWith('--post-url='));
const POST_URL = postUrlArg ? postUrlArg.slice('--post-url='.length) : '';

function pad(value) {
  return String(value).padStart(2, '0');
}

function formatMs(value) {
  return `${value.toFixed(1)} ms`;
}

function formatBytes(bytes) {
  const units = ['B', 'KB', 'MB', 'GB'];
  let size = bytes;
  let index = 0;
  while (size >= 1024 && index < units.length - 1) {
    size /= 1024;
    index += 1;
  }
  return `${size.toFixed(size >= 10 || index === 0 ? 1 : 2)} ${units[index]}`;
}

function timed(label, fn) {
  const start = performance.now();
  const result = fn();
  const end = performance.now();
  return { label, ms: end - start, result };
}

async function timedAsync(label, fn) {
  const start = performance.now();
  const result = await fn();
  const end = performance.now();
  return { label, ms: end - start, result };
}

function makeAudienceDetail(index) {
  const day = pad((index % 28) + 1);
  const month = pad((index % 12) + 1);
  return {
    referenceClient: `AUD-${index + 1}`,
    audience: `${day}/${month}/2026`,
    juge: `Juge ${index % 35}`,
    sort: ['Renvoi', 'Délibéré', 'Jugement', 'Plaidoirie'][index % 4],
    tribunal: ['Casablanca', 'Rabat', 'Marrakech', 'Tanger'][index % 4],
    depotLe: `${day}/${month}/2025`,
    color: ['blue', 'green', 'red', 'yellow'][index % 4]
  };
}

function makeDossier(index, audienceSlots) {
  const procOptions = ['SFDC', 'S/bien', 'Injonction'];
  const procedure = procOptions[index % procOptions.length];
  const base = {
    dateAffectation: `${pad((index % 28) + 1)}/${pad(((index + 3) % 12) + 1)}/2025`,
    type: ['Auto', 'Crédit', 'Recouvrement'][index % 3],
    procedure,
    referenceClient: `REF-${index + 1}`,
    debiteur: `Debiteur ${index + 1}`,
    montant: String(1000 + (index % 90000)),
    ww: `WW-${100000 + index}`,
    marque: ['Dacia', 'Renault', 'Peugeot', 'Hyundai'][index % 4],
    adresse: `Adresse ${index + 1}`,
    ville: ['Casablanca', 'Rabat', 'Fes', 'Agadir'][index % 4],
    statut: index % 7 === 0 ? 'Actif' : 'En cours',
    avancement: `Etape ${index % 9}`,
    note: '',
    procedureDetails: {},
    history: [],
    montantByProcedure: {}
  };

  base.procedureDetails[procedure] = {
    referenceClient: `PROC-${index + 1}`
  };

  for (let i = 0; i < audienceSlots.length; i += 1) {
    const key = audienceSlots[i].proc;
    base.procedureDetails[key] = {
      ...(base.procedureDetails[key] || {}),
      ...makeAudienceDetail(audienceSlots[i].audienceIndex)
    };
  }

  return base;
}

function getProcedureBaseName(procName) {
  const raw = String(procName || '').trim();
  if (!raw) return '';
  return raw.replace(/\d+$/, '').trim() || raw;
}

function isExecutionProcedure(procName) {
  const base = getProcedureBaseName(procName);
  return base === 'SFDC' || base === 'S/bien' || base === 'Injonction';
}

function parseDdMmYyyy(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;
  const out = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  return Number.isNaN(out.getTime()) ? null : out;
}

function isAssAudienceDue(details, referenceDate = new Date()) {
  const parsed = parseDdMmYyyy(details?.audience);
  if (!(parsed instanceof Date)) return false;
  const audienceDay = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()).getTime();
  const todayStart = new Date(referenceDate);
  todayStart.setHours(0, 0, 0, 0);
  return audienceDay < todayStart.getTime();
}

function countDiligenceRows(payload, referenceDate = new Date()) {
  let count = 0;
  const clients = Array.isArray(payload?.clients) ? payload.clients : [];
  clients.forEach((client) => {
    (Array.isArray(client?.dossiers) ? client.dossiers : []).forEach((dossier) => {
      const details = dossier?.procedureDetails && typeof dossier.procedureDetails === 'object'
        ? dossier.procedureDetails
        : {};
      Object.entries(details).forEach(([procName, procDetails]) => {
        const base = getProcedureBaseName(procName);
        if (isExecutionProcedure(base)) {
          count += 1;
          return;
        }
        if (base === 'ASS' && isAssAudienceDue(procDetails, referenceDate)) {
          count += 1;
        }
      });
    });
  });
  return count;
}

function getNextProcedureVariantName(existingNames, sourceProc) {
  const base = getProcedureBaseName(sourceProc);
  let maxIndex = 1;
  (existingNames || []).forEach((name) => {
    const raw = String(name || '').trim();
    if (!raw || getProcedureBaseName(raw) !== base) return;
    const match = raw.match(/(\d+)$/);
    if (!match) return;
    const value = Number(match[1]);
    if (Number.isFinite(value) && value > maxIndex) maxIndex = value;
  });
  return `${base}${maxIndex + 1}`;
}

function augmentPayloadToTargetDiligenceCount(payload, diligenceTarget) {
  const target = Math.max(0, Number(diligenceTarget) || 0);
  if (!target) return payload;

  let currentCount = countDiligenceRows(payload);
  if (currentCount >= target) return payload;

  const baseProcedures = ['SFDC', 'S/bien', 'Injonction'];
  let sequence = 0;

  outer:
  while (currentCount < target) {
    let changed = false;
    for (const client of payload.clients || []) {
      for (const dossier of client.dossiers || []) {
        if (currentCount >= target) break outer;
        if (!dossier.procedureDetails || typeof dossier.procedureDetails !== 'object') {
          dossier.procedureDetails = {};
        }
        const existingNames = Object.keys(dossier.procedureDetails);
        const baseProc = baseProcedures[sequence % baseProcedures.length];
        const procName = getNextProcedureVariantName(existingNames, baseProc);
        dossier.procedureDetails[procName] = {
          referenceClient: `DIL-${currentCount + 1}`,
          tribunal: ['Casablanca', 'Rabat', 'Marrakech', 'Tanger'][sequence % 4],
          executionNo: `EX-${100000 + currentCount}`,
          sort: ['att', 'ok', 'ko'][sequence % 3],
          dateDepot: `${pad((sequence % 28) + 1)}/${pad(((sequence + 4) % 12) + 1)}/2025`
        };
        currentCount += 1;
        sequence += 1;
        changed = true;
      }
    }
    if (!changed) break;
  }

  return payload;
}

function buildPayload(options = {}) {
  const dossierCount = Math.max(1, Number(options.dossiers || DOSSIER_COUNT) || DOSSIER_COUNT);
  const audienceCount = Math.max(0, Number(options.audiences || AUDIENCE_COUNT) || AUDIENCE_COUNT);
  const clientCount = Math.max(1, Number(options.clients || CLIENT_COUNT) || CLIENT_COUNT);
  const diligenceTarget = Math.max(0, Number(options.diligences || DILIGENCE_COUNT) || 0);
  const clients = [];
  const audienceAssignments = Array.from({ length: dossierCount }, () => []);
  for (let i = 0; i < audienceCount; i += 1) {
    const dossierIndex = i % dossierCount;
    const proc = ['ASS', 'Restitution', 'ASS2', 'Audience'][i % 4];
    audienceAssignments[dossierIndex].push({ proc, audienceIndex: i });
  }

  let globalDossierIndex = 0;
  for (let clientIndex = 0; clientIndex < clientCount; clientIndex += 1) {
    const dossiers = [];
    const countForClient = Math.floor(dossierCount / clientCount) + (clientIndex < (dossierCount % clientCount) ? 1 : 0);
    for (let i = 0; i < countForClient; i += 1) {
      dossiers.push(makeDossier(globalDossierIndex, audienceAssignments[globalDossierIndex]));
      globalDossierIndex += 1;
    }
    clients.push({
      id: clientIndex + 1,
      name: `Client ${clientIndex + 1}`,
      dossiers
    });
  }

  const payload = {
    clients,
    salleAssignments: [],
    users: [
      { id: 1, username: 'manager', password: '1234', role: 'manager', clientIds: [] }
    ],
    audienceDraft: {},
    recycleBin: [],
    recycleArchive: []
  };

  if (diligenceTarget > 0) {
    augmentPayloadToTargetDiligenceCount(payload, diligenceTarget);
  }

  return payload;
}

async function writeStateAndBackup(jsonText) {
  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  const statePath = path.join(OUTPUT_DIR, 'state.json');
  const backupName = `state_${new Date().getFullYear()}-${pad(new Date().getMonth() + 1)}-${pad(new Date().getDate())}_${pad(new Date().getHours())}-${pad(new Date().getMinutes())}-${pad(new Date().getSeconds())}.json`;
  const backupDir = path.join(OUTPUT_DIR, 'backups');
  await fs.mkdir(backupDir, { recursive: true });
  await fs.writeFile(statePath, jsonText, 'utf8');
  await fs.writeFile(path.join(backupDir, backupName), jsonText, 'utf8');
  const stats = await fs.stat(statePath);
  return { statePath, backupDir, bytes: stats.size };
}

async function postState(jsonText) {
  if (!POST_URL) return null;
  const response = await fetch(POST_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: jsonText
  });
  const body = await response.text();
  return {
    ok: response.ok,
    status: response.status,
    body
  };
}

async function main() {
  console.log(`Benchmarking ${DOSSIER_COUNT} dossiers / ${AUDIENCE_COUNT} audience entries`);

  const generate = timed('generate payload', () => buildPayload({
    dossiers: DOSSIER_COUNT,
    audiences: AUDIENCE_COUNT,
    clients: CLIENT_COUNT,
    diligences: DILIGENCE_COUNT
  }));
  const payload = generate.result;
  console.log(`${generate.label}: ${formatMs(generate.ms)}`);
  console.log(`estimated diligence rows: ${countDiligenceRows(payload)}`);

  const stringify = timed('JSON stringify', () => JSON.stringify(payload));
  const jsonText = stringify.result;
  console.log(`${stringify.label}: ${formatMs(stringify.ms)}`);
  console.log(`payload size: ${formatBytes(Buffer.byteLength(jsonText, 'utf8'))}`);

  const parse = timed('JSON parse', () => JSON.parse(jsonText));
  console.log(`${parse.label}: ${formatMs(parse.ms)}`);

  const write = await timedAsync('write state + backup files', () => writeStateAndBackup(jsonText));
  console.log(`${write.label}: ${formatMs(write.ms)}`);
  console.log(`state path: ${write.result.statePath}`);
  console.log(`backup dir: ${write.result.backupDir}`);
  console.log(`written bytes: ${formatBytes(write.result.bytes)}`);

  if (POST_URL) {
    const post = await timedAsync('HTTP POST /api/state', () => postState(jsonText));
    console.log(`${post.label}: ${formatMs(post.ms)}`);
    console.log(`http status: ${post.result.status} | ok: ${post.result.ok}`);
  }

  const mem = process.memoryUsage();
  console.log(`rss: ${formatBytes(mem.rss)} | heapUsed: ${formatBytes(mem.heapUsed)}`);
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = {
  buildPayload,
  countDiligenceRows,
  DOSSIER_COUNT,
  AUDIENCE_COUNT,
  CLIENT_COUNT,
  DILIGENCE_COUNT
};
